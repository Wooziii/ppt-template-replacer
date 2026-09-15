from __future__ import annotations

import os
import queue
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # pragma: no cover - optional dependency in dev/runtime environments.
    DND_FILES = None
    TkinterDnD = None

from .config import PrototypeOptions
from .pipeline import process_source
from .template_probe import probe_template
from .template_profile import extract_template_profile
from .template_theme import extract_template_theme
from .workspace import (
    default_input_scan_root,
    default_workspace_root,
    derive_probe_output_root,
    derive_replace_output_root,
    discover_initial_pptx_inputs,
    is_frozen_app,
)


DND_AVAILABLE = TkinterDnD is not None and DND_FILES is not None


@dataclass
class ResultRecord:
    source: str
    case_dir: str
    prototype_pptx: str
    report_path: str
    results_json: str
    risk_slides: list[int]
    engine_used: str
    template_palette_json: str


@dataclass
class ProbeRecord:
    template: str
    output_dir: str
    report_path: str
    profile_json: str
    palette_json: str
    export_error: str | None
    theme: dict[str, Any]


def _split_drop_paths(root: tk.Misc, raw_data: str) -> list[str]:
    try:
        parts = list(root.tk.splitlist(raw_data))
    except tk.TclError:
        parts = [raw_data]
    cleaned: list[str] = []
    for item in parts:
        path = item.strip().strip("{}").strip()
        if path:
            cleaned.append(str(Path(path).expanduser()))
    return cleaned


class PrototypeGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PPT 模板替换工具")
        self.root.geometry("1320x900")
        self.root.minsize(1180, 800)

        self.template_path_var = tk.StringVar()
        self.quick_source_path_var = tk.StringVar()
        self.workspace_root_var = tk.StringVar(value=str(default_workspace_root()))
        self.max_report_slides_var = tk.IntVar(value=6)
        self.edit_engine_var = tk.StringVar(value="com")
        self.skip_render_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="待开始")
        self.template_probe_status_var = tk.StringVar(value="待开始")
        self.quick_status_var = tk.StringVar(value="先选择模板和待替换 PPT")
        self.profile_summary_var = tk.StringVar(value="尚未选择模板")
        self.template_probe_summary_var = tk.StringVar(value="尚未选择模板")
        self.runtime_summary_var = tk.StringVar(value=self._build_runtime_summary())
        self.quick_template_summary_var = tk.StringVar(value="")
        self.quick_source_summary_var = tk.StringVar(value="")
        self.quick_output_summary_var = tk.StringVar(value="")
        self.quick_result_summary_var = tk.StringVar(value="生成完成后，可在这里一键打开结果。")
        self.replace_output_summary_var = tk.StringVar(value="正式替换输出：尚未设置")
        self.probe_output_summary_var = tk.StringVar(value="模板探针输出：尚未设置")

        self.source_files: list[str] = []
        self.results_by_item: dict[str, ResultRecord] = {}
        self.probe_results_by_item: dict[str, ProbeRecord] = {}
        self.last_replace_result: ResultRecord | None = None
        self.event_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self._build_layout()
        self._auto_fill_from_workspace()
        self._refresh_sources()
        self._refresh_output_summary()
        self.root.after(120, self._poll_events)

    def _build_layout(self) -> None:
        self.root.configure(bg="#f0f3f9")
        self._build_styles()

        root_frame = ttk.Frame(self.root, padding=16)
        root_frame.pack(fill=tk.BOTH, expand=True)
        root_frame.columnconfigure(0, weight=1)
        root_frame.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(root_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        quick_tab = ttk.Frame(self.notebook, padding=16)
        replace_tab = ttk.Frame(self.notebook, padding=16)
        probe_tab = ttk.Frame(self.notebook, padding=16)
        self.notebook.add(quick_tab, text="一键替换")
        self.notebook.add(replace_tab, text="批量替换")
        self.notebook.add(probe_tab, text="模板管理")

        self._build_quick_tab(quick_tab)
        self._build_replace_tab(replace_tab)
        self._build_probe_tab(probe_tab)

    def _build_styles(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Header.TLabel", background="#ffffff", foreground="#0b0a18", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Body.TLabel", background="#ffffff", foreground="#334155", font=("Microsoft YaHei UI", 9))
        style.configure("Muted.TLabel", background="#ffffff", foreground="#64748b", font=("Microsoft YaHei UI", 9))
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Accent.Horizontal.TProgressbar", troughcolor="#e5e7eb", background="#42D7FF", bordercolor="#e5e7eb")
        style.configure("TNotebook.Tab", padding=(14, 8))

    def _build_quick_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        hero_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        hero_card.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        hero_card.columnconfigure(0, weight=1)
        ttk.Label(hero_card, text="一键替换", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            hero_card,
            text="只需要选 1 个模板和 1 个待替换 PPT。默认输出到文档目录；需要批量处理时再进“批量替换”。",
            style="Body.TLabel",
            wraplength=1080,
            justify=tk.LEFT,
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        content = ttk.Frame(parent)
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=2)
        content.rowconfigure(0, weight=1)

        chooser_card = ttk.Frame(content, style="Card.TFrame", padding=20)
        chooser_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        chooser_card.columnconfigure(0, weight=1)
        chooser_card.columnconfigure(1, weight=1)
        chooser_card.rowconfigure(1, weight=1)

        ttk.Label(chooser_card, text="第 1 步：选择文件", style="Header.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")

        template_zone = self._create_drop_zone(
            chooser_card,
            title="模板 PPTX",
            textvariable=self.quick_template_summary_var,
            button_text="选择模板",
            button_command=self._choose_template,
        )
        template_zone.grid(row=1, column=0, sticky="nsew", padx=(0, 10), pady=(14, 0))
        self._register_drop_target(template_zone, self._on_template_drop)
        self._register_drop_target(template_zone.drop_label, self._on_template_drop)

        source_zone = self._create_drop_zone(
            chooser_card,
            title="待替换 PPTX",
            textvariable=self.quick_source_summary_var,
            button_text="选择待替换 PPT",
            button_command=self._choose_quick_source,
        )
        source_zone.grid(row=1, column=1, sticky="nsew", padx=(10, 0), pady=(14, 0))
        self._register_drop_target(source_zone, self._on_quick_source_drop)
        self._register_drop_target(source_zone.drop_label, self._on_quick_source_drop)

        action_card = ttk.Frame(content, style="Card.TFrame", padding=20)
        action_card.grid(row=0, column=1, sticky="nsew")
        action_card.columnconfigure(0, weight=1)

        ttk.Label(action_card, text="第 2 步：开始替换", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(action_card, textvariable=self.quick_output_summary_var, style="Body.TLabel", wraplength=360, justify=tk.LEFT).grid(
            row=1, column=0, sticky="w", pady=(10, 0)
        )
        ttk.Button(action_card, text="开始替换", style="Primary.TButton", command=self._start_quick_run).grid(
            row=2, column=0, sticky="ew", pady=(16, 0)
        )
        ttk.Label(action_card, textvariable=self.quick_status_var, style="Body.TLabel", wraplength=360, justify=tk.LEFT).grid(
            row=3, column=0, sticky="w", pady=(10, 0)
        )
        self.quick_progressbar = ttk.Progressbar(action_card, style="Accent.Horizontal.TProgressbar", mode="indeterminate")
        self.quick_progressbar.grid(row=4, column=0, sticky="ew", pady=(12, 0))

        ttk.Label(action_card, text="最近结果", style="Header.TLabel").grid(row=5, column=0, sticky="w", pady=(18, 0))
        ttk.Label(action_card, textvariable=self.quick_result_summary_var, style="Body.TLabel", wraplength=360, justify=tk.LEFT).grid(
            row=6, column=0, sticky="w", pady=(8, 0)
        )

        quick_actions = ttk.Frame(action_card, style="Card.TFrame")
        quick_actions.grid(row=7, column=0, sticky="ew", pady=(14, 0))
        ttk.Button(quick_actions, text="打开生成 PPT", command=self._open_last_generated_ppt).pack(side=tk.LEFT)
        ttk.Button(quick_actions, text="打开输出目录", command=self._open_last_output_folder).pack(side=tk.LEFT, padx=6)
        ttk.Button(quick_actions, text="高级设置", command=lambda: self.notebook.select(1)).pack(side=tk.LEFT)

        ttk.Label(
            action_card,
            text="支持把 `.pptx` 直接拖进左侧两个框里。模板探针、批量处理和高级参数都保留在后面的页签。",
            style="Muted.TLabel",
            wraplength=360,
            justify=tk.LEFT,
        ).grid(row=8, column=0, sticky="w", pady=(18, 0))

    def _create_drop_zone(
        self,
        parent: ttk.Frame,
        title: str,
        textvariable: tk.StringVar,
        button_text: str,
        button_command,
    ) -> tk.Frame:
        frame = tk.Frame(parent, bg="#ffffff", highlightbackground="#d7e3f4", highlightthickness=1, bd=0)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        title_label = tk.Label(
            frame,
            text=title,
            bg="#ffffff",
            fg="#0b0a18",
            font=("Microsoft YaHei UI", 11, "bold"),
            anchor="w",
        )
        title_label.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))

        drop_label = tk.Label(
            frame,
            textvariable=textvariable,
            bg="#f8fbff",
            fg="#334155",
            relief="groove",
            bd=2,
            padx=24,
            pady=32,
            justify="center",
            font=("Microsoft YaHei UI", 10),
            cursor="hand2",
        )
        drop_label.grid(row=1, column=0, sticky="nsew", padx=16)
        drop_label.bind("<Button-1>", lambda _event: button_command())

        button = ttk.Button(frame, text=button_text, command=button_command)
        button.grid(row=2, column=0, sticky="w", padx=16, pady=(10, 16))

        for widget in (frame, title_label, drop_label):
            widget.bind("<Button-1>", lambda _event: button_command())
        frame.drop_label = drop_label
        return frame

    def _build_replace_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=3)
        parent.columnconfigure(1, weight=2)
        parent.rowconfigure(1, weight=1)
        parent.rowconfigure(2, weight=1)

        top_card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        top_card.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 12))
        top_card.columnconfigure(1, weight=1)
        top_card.columnconfigure(4, weight=1)

        ttk.Label(top_card, text="模板文件", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(top_card, textvariable=self.template_path_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(top_card, text="选择模板", command=self._choose_template).grid(row=0, column=2, sticky="ew")

        ttk.Label(top_card, text="工作区目录", style="Header.TLabel").grid(row=0, column=3, sticky="w", padx=(18, 0))
        ttk.Entry(top_card, textvariable=self.workspace_root_var).grid(row=0, column=4, sticky="ew", padx=(8, 8))
        ttk.Button(top_card, text="选择目录", command=self._choose_workspace_root).grid(row=0, column=5, sticky="ew")

        ttk.Label(top_card, text="模板摘要", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Label(top_card, textvariable=self.profile_summary_var, style="Body.TLabel", wraplength=980).grid(
            row=1, column=1, columnspan=5, sticky="w", pady=(12, 0)
        )
        ttk.Label(top_card, text="输出位置", style="Body.TLabel").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Label(top_card, textvariable=self.replace_output_summary_var, style="Body.TLabel", wraplength=980).grid(
            row=2, column=1, columnspan=5, sticky="w", pady=(8, 0)
        )
        ttk.Label(top_card, text="运行环境", style="Body.TLabel").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Label(top_card, textvariable=self.runtime_summary_var, style="Body.TLabel", wraplength=980).grid(
            row=3, column=1, columnspan=5, sticky="w", pady=(8, 0)
        )

        left_card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        left_card.grid(row=1, column=0, rowspan=2, sticky="nsew", padx=(0, 12))
        left_card.columnconfigure(0, weight=1)
        left_card.rowconfigure(1, weight=1)
        left_card.rowconfigure(3, weight=1)

        ttk.Label(left_card, text="待处理文件", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        source_actions = ttk.Frame(left_card, style="Card.TFrame")
        source_actions.grid(row=0, column=1, sticky="e")
        ttk.Button(source_actions, text="添加文件", command=self._add_sources).pack(side=tk.LEFT)
        ttk.Button(source_actions, text="移除选中", command=self._remove_selected_source).pack(side=tk.LEFT, padx=6)
        ttk.Button(source_actions, text="清空", command=self._clear_sources).pack(side=tk.LEFT)

        self.source_listbox = tk.Listbox(left_card, activestyle="none", font=("Microsoft YaHei UI", 10), selectmode=tk.EXTENDED)
        self.source_listbox.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(10, 14))

        options_frame = ttk.LabelFrame(left_card, text="运行选项", padding=12)
        options_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
        options_frame.columnconfigure(1, weight=1)
        ttk.Label(options_frame, text="抽样页数").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(options_frame, from_=3, to=20, textvariable=self.max_report_slides_var, width=6).grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(options_frame, text="编辑引擎").grid(row=0, column=2, sticky="w", padx=(18, 0))
        engine_combo = ttk.Combobox(options_frame, textvariable=self.edit_engine_var, values=("pptx", "com"), state="readonly", width=14)
        engine_combo.grid(row=0, column=3, sticky="w", padx=(8, 0))
        ttk.Checkbutton(
            options_frame,
            text="跳过截图导出（更稳）",
            variable=self.skip_render_var,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Label(options_frame, text="建议流程：先去“模板管理”探针，再回来正式替换。", style="Body.TLabel").grid(
            row=2, column=0, columnspan=4, sticky="w", pady=(8, 0)
        )

        run_frame = ttk.Frame(left_card, style="Card.TFrame")
        run_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(16, 0))
        run_frame.columnconfigure(0, weight=1)
        run_frame.rowconfigure(2, weight=1)
        ttk.Button(run_frame, text="开始处理", style="Primary.TButton", command=self._start_run).grid(row=0, column=0, sticky="ew")
        ttk.Label(run_frame, textvariable=self.status_var, style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 6))
        self.progressbar = ttk.Progressbar(run_frame, style="Accent.Horizontal.TProgressbar", mode="indeterminate")
        self.progressbar.grid(row=1, column=0, sticky="e", padx=(260, 0))
        self.log_text = tk.Text(run_frame, height=12, wrap="word", font=("Consolas", 9), background="#0f172a", foreground="#e2e8f0")
        self.log_text.grid(row=2, column=0, sticky="nsew")

        right_top = ttk.Frame(parent, style="Card.TFrame", padding=16)
        right_top.grid(row=1, column=1, sticky="nsew")
        right_top.columnconfigure(0, weight=1)
        right_top.rowconfigure(1, weight=1)
        ttk.Label(right_top, text="运行结果", style="Header.TLabel").grid(row=0, column=0, sticky="w")

        columns = ("source", "engine", "risks")
        self.results_tree = ttk.Treeview(right_top, columns=columns, show="headings", height=10)
        self.results_tree.heading("source", text="源文件")
        self.results_tree.heading("engine", text="引擎")
        self.results_tree.heading("risks", text="风险页")
        self.results_tree.column("source", width=280, anchor="w")
        self.results_tree.column("engine", width=90, anchor="center")
        self.results_tree.column("risks", width=120, anchor="w")
        self.results_tree.grid(row=1, column=0, sticky="nsew", pady=(10, 10))

        result_actions = ttk.Frame(right_top, style="Card.TFrame")
        result_actions.grid(row=2, column=0, sticky="ew")
        ttk.Button(result_actions, text="打开报告", command=self._open_selected_report).pack(side=tk.LEFT)
        ttk.Button(result_actions, text="打开输出目录", command=self._open_selected_folder).pack(side=tk.LEFT, padx=6)
        ttk.Button(result_actions, text="打开生成 PPT", command=self._open_selected_ppt).pack(side=tk.LEFT)
        ttk.Button(result_actions, text="打开 palette", command=self._open_selected_palette).pack(side=tk.LEFT, padx=6)

        right_bottom = ttk.Frame(parent, style="Card.TFrame", padding=16)
        right_bottom.grid(row=2, column=1, sticky="nsew", pady=(12, 0))
        right_bottom.columnconfigure(0, weight=1)
        ttk.Label(right_bottom, text="使用提示", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        hints = (
            "1. 先去“模板管理”探针模板，确认背景色、标题色、抬头位置是否抽对。\n"
            "2. 现在只需要维护一个“工作区目录”，正式替换和模板探针会自动分目录落盘；默认位置在“文档”目录。\n"
            "3. 默认使用 `com` 原生应用模板母版/版式，`pptx` 只保留为无母版兜底路径。\n"
            "4. 默认勾选“跳过截图导出”，优先保证成品生成稳定；需要截图报告时再取消。\n"
            "5. 打包版默认只扫描 exe 同目录下的 PPTX，不再依赖系统当前工作目录。\n"
            "6. 处理完成后，优先看报告、palette 和生成后的 PPT。"
        )
        ttk.Label(right_bottom, text=hints, style="Body.TLabel", wraplength=380, justify=tk.LEFT).grid(row=1, column=0, sticky="nw", pady=(10, 0))

    def _build_probe_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=3)
        parent.columnconfigure(1, weight=2)
        parent.rowconfigure(1, weight=1)
        parent.rowconfigure(2, weight=1)

        top_card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        top_card.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 12))
        top_card.columnconfigure(1, weight=1)
        top_card.columnconfigure(4, weight=1)

        ttk.Label(top_card, text="模板文件", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(top_card, textvariable=self.template_path_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(top_card, text="选择模板", command=self._choose_template).grid(row=0, column=2, sticky="ew")

        ttk.Label(top_card, text="工作区目录", style="Header.TLabel").grid(row=0, column=3, sticky="w", padx=(18, 0))
        ttk.Entry(top_card, textvariable=self.workspace_root_var).grid(row=0, column=4, sticky="ew", padx=(8, 8))
        ttk.Button(top_card, text="选择目录", command=self._choose_workspace_root).grid(row=0, column=5, sticky="ew")

        ttk.Label(top_card, text="当前模板画像", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Label(top_card, textvariable=self.template_probe_summary_var, style="Body.TLabel", wraplength=980, justify=tk.LEFT).grid(
            row=1, column=1, columnspan=5, sticky="w", pady=(12, 0)
        )
        ttk.Label(top_card, text="输出位置", style="Body.TLabel").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Label(top_card, textvariable=self.probe_output_summary_var, style="Body.TLabel", wraplength=980, justify=tk.LEFT).grid(
            row=2, column=1, columnspan=5, sticky="w", pady=(8, 0)
        )
        ttk.Label(top_card, text="运行环境", style="Body.TLabel").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Label(top_card, textvariable=self.runtime_summary_var, style="Body.TLabel", wraplength=980, justify=tk.LEFT).grid(
            row=3, column=1, columnspan=5, sticky="w", pady=(8, 0)
        )

        left_card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        left_card.grid(row=1, column=0, rowspan=2, sticky="nsew", padx=(0, 12))
        left_card.columnconfigure(0, weight=1)
        left_card.rowconfigure(1, weight=1)

        header_row = ttk.Frame(left_card, style="Card.TFrame")
        header_row.grid(row=0, column=0, sticky="ew")
        ttk.Label(header_row, text="模板探针结果", style="Header.TLabel").pack(side=tk.LEFT)
        ttk.Button(header_row, text="开始探针", style="Primary.TButton", command=self._start_template_probe).pack(side=tk.RIGHT)
        ttk.Button(header_row, text="切到替换页", command=lambda: self.notebook.select(0)).pack(side=tk.RIGHT, padx=(0, 8))

        probe_columns = ("template", "background", "title", "status")
        self.template_probe_tree = ttk.Treeview(left_card, columns=probe_columns, show="headings", height=12)
        self.template_probe_tree.heading("template", text="模板")
        self.template_probe_tree.heading("background", text="背景")
        self.template_probe_tree.heading("title", text="标题色")
        self.template_probe_tree.heading("status", text="预览")
        self.template_probe_tree.column("template", width=250, anchor="w")
        self.template_probe_tree.column("background", width=100, anchor="center")
        self.template_probe_tree.column("title", width=100, anchor="center")
        self.template_probe_tree.column("status", width=120, anchor="center")
        self.template_probe_tree.grid(row=1, column=0, sticky="nsew", pady=(12, 10))

        probe_actions = ttk.Frame(left_card, style="Card.TFrame")
        probe_actions.grid(row=2, column=0, sticky="ew")
        ttk.Button(probe_actions, text="打开报告", command=self._open_selected_probe_report).pack(side=tk.LEFT)
        ttk.Button(probe_actions, text="打开输出目录", command=self._open_selected_probe_folder).pack(side=tk.LEFT, padx=6)
        ttk.Button(probe_actions, text="打开 profile", command=self._open_selected_probe_profile).pack(side=tk.LEFT)
        ttk.Button(probe_actions, text="打开 palette", command=self._open_selected_probe_palette).pack(side=tk.LEFT, padx=6)

        right_card = ttk.Frame(parent, style="Card.TFrame", padding=16)
        right_card.grid(row=1, column=1, rowspan=2, sticky="nsew")
        right_card.columnconfigure(0, weight=1)
        right_card.rowconfigure(2, weight=1)

        ttk.Label(right_card, text="模板管理", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            right_card,
            text="建议流程：先探针模板，看 profile / palette 和预览，再回“批量替换”正式处理源文件。",
            style="Body.TLabel",
            wraplength=380,
            justify=tk.LEFT,
        ).grid(row=1, column=0, sticky="w", pady=(10, 8))
        ttk.Label(right_card, textvariable=self.template_probe_status_var, style="Body.TLabel").grid(row=2, column=0, sticky="nw")
        self.template_probe_progressbar = ttk.Progressbar(right_card, style="Accent.Horizontal.TProgressbar", mode="indeterminate")
        self.template_probe_progressbar.grid(row=2, column=0, sticky="ne")
        self.template_probe_log_text = tk.Text(
            right_card,
            height=18,
            wrap="word",
            font=("Consolas", 9),
            background="#0f172a",
            foreground="#e2e8f0",
        )
        self.template_probe_log_text.grid(row=3, column=0, sticky="nsew", pady=(10, 0))

    def _auto_fill_from_workspace(self) -> None:
        template, sources = discover_initial_pptx_inputs(default_input_scan_root())
        if template:
            self.template_path_var.set(template)
            self._refresh_template_summary()
        if sources and not self.quick_source_path_var.get().strip():
            self.quick_source_path_var.set(sources[0])
        self._refresh_output_summary()
        self.source_files = sources
        self._refresh_quick_selection_summaries()

    def _build_runtime_summary(self) -> str:
        mode = "打包版" if is_frozen_app() else "源码版"
        return (
            f"模式：{mode}；默认工作区写入 {default_workspace_root()}；"
            f"自动识别同级 PPTX 目录：{default_input_scan_root()}；"
            "安装 PowerPoint 后可用 COM 高保真编辑和预览导出。"
        )

    def _register_drop_target(self, widget: tk.BaseWidget, handler) -> None:
        if not DND_AVAILABLE:
            return
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>", lambda event: handler(event.data))

    def _refresh_quick_selection_summaries(self) -> None:
        template = self.template_path_var.get().strip()
        source = self.quick_source_path_var.get().strip()
        drag_hint = "支持拖拽到这里" if DND_AVAILABLE else "点击按钮选择文件"

        if template and Path(template).exists():
            self.quick_template_summary_var.set(
                f"已选择模板\n{Path(template).name}\n\n点击更换，或继续拖入新模板"
            )
        else:
            self.quick_template_summary_var.set(
                f"把模板 PPTX 拖到这里\n\n{drag_hint}\n只支持 .pptx"
            )

        if source and Path(source).exists():
            self.quick_source_summary_var.set(
                f"已选择待替换 PPT\n{Path(source).name}\n\n点击更换，或继续拖入新文件"
            )
        else:
            self.quick_source_summary_var.set(
                f"把待替换 PPT 拖到这里\n\n{drag_hint}\n只支持 .pptx"
            )

    def _refresh_quick_output_summary(self) -> None:
        workspace_raw = self.workspace_root_var.get().strip()
        workspace_root = Path(workspace_raw).expanduser() if workspace_raw else default_workspace_root()
        template_raw = self.template_path_var.get().strip()
        replace_root = derive_replace_output_root(workspace_root, template_raw or None)
        self.quick_output_summary_var.set(
            f"默认保存位置：{replace_root}\n如需改位置，可在“批量替换”页修改工作区目录。"
        )

    def _on_template_drop(self, raw_data: str) -> None:
        paths = _split_drop_paths(self.root, raw_data)
        if not paths:
            return
        template_path = paths[0]
        if Path(template_path).suffix.lower() != ".pptx":
            messagebox.showerror("文件类型错误", "模板只支持 .pptx 文件。")
            return
        self.template_path_var.set(template_path)
        self._refresh_template_summary()
        self._refresh_output_summary()
        self._refresh_quick_selection_summaries()

    def _on_quick_source_drop(self, raw_data: str) -> None:
        paths = _split_drop_paths(self.root, raw_data)
        if not paths:
            return
        source_path = paths[0]
        if Path(source_path).suffix.lower() != ".pptx":
            messagebox.showerror("文件类型错误", "待替换文件只支持 .pptx。")
            return
        self.quick_source_path_var.set(source_path)
        self._refresh_quick_selection_summaries()

    def _choose_quick_source(self) -> None:
        file_path = filedialog.askopenfilename(title="选择待替换 PPTX", filetypes=[("PowerPoint", "*.pptx")])
        if file_path:
            self.quick_source_path_var.set(file_path)
            self._refresh_quick_selection_summaries()

    def _validate_replace_inputs(self, source_files: list[str]) -> bool:
        template = Path(self.template_path_var.get().strip())
        if not template.exists():
            messagebox.showerror("模板缺失", "请先选择有效的模板 PPTX。")
            return False
        if not source_files:
            messagebox.showerror("源文件缺失", "请至少选择一个待处理 PPTX。")
            return False
        missing = [item for item in source_files if not Path(item).exists()]
        if missing:
            messagebox.showerror("源文件不存在", "\n".join(missing[:5]))
            return False
        return True

    def _refresh_template_summary(self) -> None:
        template = self.template_path_var.get().strip()
        if not template:
            self.profile_summary_var.set("尚未选择模板")
            self.template_probe_summary_var.set("尚未选择模板")
            return
        path = Path(template)
        if not path.exists():
            self.profile_summary_var.set("模板路径不存在")
            self.template_probe_summary_var.set("模板路径不存在")
            return
        try:
            profile = extract_template_profile(path)
            theme = extract_template_theme(path)
            self.profile_summary_var.set(
                "header_left={:.3f}, header_top={:.3f}, header_width={:.3f}, safe_zone_bottom={:.3f}, bg={}, title={}, accent={}".format(
                    profile.header_left_ratio,
                    profile.header_top_ratio,
                    profile.header_width_ratio,
                    profile.safe_zone_bottom_ratio,
                    theme.palette.background_dark,
                    theme.palette.title_cyan,
                    theme.palette.accent_orange,
                )
            )
            self.template_probe_summary_var.set(
                "主背景色：{}  标题色：{}  强调色：{}  面板色：{} / {}  宽版抬头宽度：{:.3f}  安全区：{:.3f}".format(
                    theme.palette.background_dark,
                    theme.palette.title_cyan,
                    theme.palette.accent_orange,
                    theme.palette.panel_mid,
                    theme.palette.panel_alt,
                    max(profile.header_width_ratio, 0.58),
                    profile.safe_zone_bottom_ratio,
                )
            )
        except Exception as exc:
            self.profile_summary_var.set(f"模板解析失败: {exc}")
            self.template_probe_summary_var.set(f"模板解析失败: {exc}")

    def _refresh_output_summary(self) -> None:
        workspace_raw = self.workspace_root_var.get().strip()
        workspace_root = Path(workspace_raw).expanduser() if workspace_raw else default_workspace_root()
        template_raw = self.template_path_var.get().strip()
        replace_root = derive_replace_output_root(workspace_root, template_raw or None)
        probe_root = derive_probe_output_root(workspace_root)
        self.replace_output_summary_var.set(f"正式替换输出：{replace_root}")
        self.probe_output_summary_var.set(f"模板探针输出：{probe_root}")
        self._refresh_quick_output_summary()

    def _refresh_sources(self) -> None:
        self.source_listbox.delete(0, tk.END)
        for item in self.source_files:
            self.source_listbox.insert(tk.END, item)

    def _choose_template(self) -> None:
        file_path = filedialog.askopenfilename(title="选择模板 PPTX", filetypes=[("PowerPoint", "*.pptx")])
        if file_path:
            self.template_path_var.set(file_path)
            self._refresh_template_summary()
            self._refresh_output_summary()
            self._refresh_quick_selection_summaries()

    def _choose_workspace_root(self) -> None:
        directory = filedialog.askdirectory(title="选择工作区目录")
        if directory:
            self.workspace_root_var.set(directory)
            self._refresh_output_summary()

    def _add_sources(self) -> None:
        files = filedialog.askopenfilenames(title="选择源 PPTX", filetypes=[("PowerPoint", "*.pptx")])
        if not files:
            return
        existing = set(self.source_files)
        for file in files:
            if file not in existing:
                self.source_files.append(file)
        self._refresh_sources()

    def _remove_selected_source(self) -> None:
        selected = list(self.source_listbox.curselection())
        if not selected:
            return
        for index in reversed(selected):
            del self.source_files[index]
        self._refresh_sources()

    def _clear_sources(self) -> None:
        self.source_files.clear()
        self._refresh_sources()

    def _append_log(self, message: str) -> None:
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def _append_probe_log(self, message: str) -> None:
        self.template_probe_log_text.insert(tk.END, message + "\n")
        self.template_probe_log_text.see(tk.END)

    def _set_replace_running(self, running: bool) -> None:
        if running:
            self.progressbar.start(12)
            self.quick_progressbar.start(12)
        else:
            self.progressbar.stop()
            self.quick_progressbar.stop()

    def _set_probe_running(self, running: bool) -> None:
        if running:
            self.template_probe_progressbar.start(12)
        else:
            self.template_probe_progressbar.stop()

    def _validate_template_for_probe(self) -> bool:
        template = Path(self.template_path_var.get().strip())
        if not template.exists():
            messagebox.showerror("模板缺失", "请先选择有效的模板 PPTX。")
            return False
        return True

    def _start_quick_run(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("正在运行", "当前还有任务在执行。")
            return
        quick_source = self.quick_source_path_var.get().strip()
        if not self._validate_replace_inputs([quick_source] if quick_source else []):
            return
        self._begin_replace_run([quick_source])

    def _start_run(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("正在运行", "当前还有任务在执行。")
            return
        if not self._validate_replace_inputs(self.source_files):
            return
        self._begin_replace_run(self.source_files)

    def _begin_replace_run(self, source_files: list[str]) -> None:
        self.results_tree.delete(*self.results_tree.get_children())
        self.results_by_item.clear()
        self.last_replace_result = None
        self.log_text.delete("1.0", tk.END)
        self.status_var.set("正在处理")
        self.quick_status_var.set("正在处理")
        self.quick_result_summary_var.set("正在生成，请稍候。")
        self._set_replace_running(True)

        template = Path(self.template_path_var.get().strip())
        workspace_root = Path(self.workspace_root_var.get().strip()).expanduser()
        output_root = derive_replace_output_root(workspace_root, template)
        max_slides = self.max_report_slides_var.get()
        edit_engine = self.edit_engine_var.get()
        skip_render = self.skip_render_var.get()
        profile = extract_template_profile(template)
        theme = extract_template_theme(template)
        sources = [Path(item) for item in source_files]

        def worker() -> None:
            try:
                for index, source in enumerate(sources, start=1):
                    self.event_queue.put(("replace_status", f"处理中 {index}/{len(sources)}: {source.name}"))
                    options = PrototypeOptions(
                        template_path=template,
                        source_path=source,
                        output_root=output_root,
                        max_report_slides=max_slides,
                        palette=theme.palette,
                        palette_metadata=theme.metadata,
                        profile=profile,
                        edit_engine=edit_engine,
                        skip_render=skip_render,
                    )
                    result = process_source(options, progress=lambda msg: self.event_queue.put(("replace_log", f"[{source.name}] {msg}")))
                    self.event_queue.put(("replace_result", result))
                self.event_queue.put(("replace_done", None))
            except Exception as exc:
                self.event_queue.put(("replace_error", str(exc)))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def _start_template_probe(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("正在运行", "当前还有任务在执行。")
            return
        if not self._validate_template_for_probe():
            return

        self.template_probe_log_text.delete("1.0", tk.END)
        self.template_probe_status_var.set("正在探针")
        self._set_probe_running(True)

        template = Path(self.template_path_var.get().strip())
        workspace_root = Path(self.workspace_root_var.get().strip()).expanduser()
        output_root = derive_probe_output_root(workspace_root)

        def worker() -> None:
            try:
                self.event_queue.put(("probe_log", f"[{template.name}] 开始模板探针"))
                result = probe_template(
                    template,
                    output_root,
                    progress=lambda msg: self.event_queue.put(("probe_log", f"[{template.name}] {msg}")),
                )
                self.event_queue.put(("probe_result", result))
                self.event_queue.put(("probe_done", None))
            except Exception as exc:
                self.event_queue.put(("probe_error", str(exc)))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def _poll_events(self) -> None:
        try:
            while True:
                event, payload = self.event_queue.get_nowait()
                if event == "replace_log":
                    self._append_log(str(payload))
                elif event == "probe_log":
                    self._append_probe_log(str(payload))
                elif event == "replace_status":
                    self.status_var.set(str(payload))
                    self.quick_status_var.set(str(payload))
                elif event == "probe_status":
                    self.template_probe_status_var.set(str(payload))
                elif event == "replace_result":
                    self._handle_replace_result(payload)
                elif event == "probe_result":
                    self._handle_probe_result(payload)
                elif event == "replace_done":
                    self.status_var.set("全部处理完成")
                    self.quick_status_var.set("全部处理完成")
                    self._set_replace_running(False)
                    self._append_log("全部任务处理完成。")
                elif event == "probe_done":
                    self.template_probe_status_var.set("模板探针完成")
                    self._set_probe_running(False)
                    self._append_probe_log("模板探针完成。")
                elif event == "replace_error":
                    self.status_var.set("运行失败")
                    self.quick_status_var.set("运行失败")
                    self._set_replace_running(False)
                    self._append_log(f"运行失败: {payload}")
                    self.quick_result_summary_var.set(f"生成失败：{payload}")
                    messagebox.showerror("运行失败", str(payload))
                elif event == "probe_error":
                    self.template_probe_status_var.set("探针失败")
                    self._set_probe_running(False)
                    self._append_probe_log(f"模板探针失败: {payload}")
                    messagebox.showerror("模板探针失败", str(payload))
        except queue.Empty:
            pass
        self.root.after(120, self._poll_events)

    def _handle_replace_result(self, payload: dict[str, Any]) -> None:
        source_name = Path(payload["source"]).name
        risks = ", ".join(str(item) for item in payload["risk_slides"]) if payload["risk_slides"] else "无"
        item_id = self.results_tree.insert("", tk.END, values=(source_name, payload["engine_used"], risks))
        record = ResultRecord(**payload)
        self.results_by_item[item_id] = record
        self.last_replace_result = record
        self.quick_result_summary_var.set(
            f"已生成：{Path(record.prototype_pptx).name}\n输出目录：{record.case_dir}"
        )
        self._append_log(f"完成: {source_name} -> {payload['prototype_pptx']}")

    def _handle_probe_result(self, payload: dict[str, Any]) -> None:
        palette = payload["theme"]["palette"]
        status = "预览失败" if payload.get("export_error") else "完成"
        template_name = Path(payload["template"]).name
        item_id = self.template_probe_tree.insert(
            "",
            tk.END,
            values=(template_name, palette["background_dark"], palette["title_cyan"], status),
        )
        self.probe_results_by_item[item_id] = ProbeRecord(**payload)
        self.template_probe_status_var.set(f"模板探针完成: {template_name}")
        self._append_probe_log(
            "完成: {} -> report={} palette={} export={}".format(
                template_name,
                payload["report_path"],
                payload["palette_json"],
                "failed" if payload.get("export_error") else "ok",
            )
        )

    def _selected_result(self) -> ResultRecord | None:
        selected = self.results_tree.selection()
        if not selected:
            messagebox.showinfo("未选择结果", "请先在结果列表中选择一项。")
            return None
        return self.results_by_item.get(selected[0])

    def _selected_probe_result(self) -> ProbeRecord | None:
        selected = self.template_probe_tree.selection()
        if not selected:
            messagebox.showinfo("未选择模板探针结果", "请先在模板管理页选择一项。")
            return None
        return self.probe_results_by_item.get(selected[0])

    def _open_path(self, path_str: str) -> None:
        path = Path(path_str)
        if not path.exists():
            messagebox.showerror("文件不存在", str(path))
            return
        os.startfile(str(path))

    def _open_selected_report(self) -> None:
        record = self._selected_result()
        if record:
            self._open_path(record.report_path)

    def _open_selected_folder(self) -> None:
        record = self._selected_result()
        if record:
            self._open_path(record.case_dir)

    def _open_selected_ppt(self) -> None:
        record = self._selected_result()
        if record:
            self._open_path(record.prototype_pptx)

    def _open_selected_palette(self) -> None:
        record = self._selected_result()
        if record:
            self._open_path(record.template_palette_json)

    def _open_selected_probe_report(self) -> None:
        record = self._selected_probe_result()
        if record:
            self._open_path(record.report_path)

    def _open_selected_probe_folder(self) -> None:
        record = self._selected_probe_result()
        if record:
            self._open_path(record.output_dir)

    def _open_selected_probe_profile(self) -> None:
        record = self._selected_probe_result()
        if record:
            self._open_path(record.profile_json)

    def _open_selected_probe_palette(self) -> None:
        record = self._selected_probe_result()
        if record:
            self._open_path(record.palette_json)

    def _open_last_generated_ppt(self) -> None:
        if not self.last_replace_result:
            messagebox.showinfo("暂无结果", "请先完成一次替换。")
            return
        self._open_path(self.last_replace_result.prototype_pptx)

    def _open_last_output_folder(self) -> None:
        if not self.last_replace_result:
            messagebox.showinfo("暂无结果", "请先完成一次替换。")
            return
        self._open_path(self.last_replace_result.case_dir)


def create_app() -> tk.Tk:
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    PrototypeGUI(root)
    return root


def main(smoke_test: bool = False) -> None:
    root = create_app()
    if smoke_test:
        root.update_idletasks()
        root.destroy()
        return
    root.mainloop()
