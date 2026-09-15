from __future__ import annotations

import os
import sys
from pathlib import Path


APP_WORKSPACE_NAME = "PPT模板替换输出"
REPLACE_DIR_NAME = "正式替换"
PROBE_DIR_NAME = "模板探针"


def is_frozen_app() -> bool:
    return bool(getattr(sys, "frozen", False))


def app_base_dir() -> Path:
    if is_frozen_app():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def default_documents_dir() -> Path:
    candidates: list[Path] = []
    userprofile = os.environ.get("USERPROFILE")
    if userprofile:
        user_root = Path(userprofile)
        candidates.extend([user_root / "Documents", user_root / "文档"])

    home = Path.home()
    candidates.extend([home / "Documents", home / "文档", home])

    for candidate in candidates:
        if candidate.exists():
            return candidate.resolve()
    return home.resolve()


def default_workspace_root() -> Path:
    return (default_documents_dir() / APP_WORKSPACE_NAME).resolve()


def default_input_scan_root() -> Path:
    return app_base_dir()


def sanitize_dir_name(value: str) -> str:
    cleaned = "".join("_" if char in '<>:"/\\|?*' else char for char in value.strip())
    cleaned = " ".join(cleaned.split())
    return cleaned.rstrip(". ") or "未命名模板"


def derive_replace_output_root(workspace_root: Path, template_path: str | Path | None) -> Path:
    if template_path:
        template_name = sanitize_dir_name(Path(template_path).stem)
    else:
        template_name = "未选择模板"
    return workspace_root / REPLACE_DIR_NAME / template_name


def derive_probe_output_root(workspace_root: Path) -> Path:
    return workspace_root / PROBE_DIR_NAME


def discover_initial_pptx_inputs(scan_root: Path) -> tuple[str | None, list[str]]:
    pptx_files = sorted(scan_root.glob("*.pptx"))
    template: str | None = None
    sources: list[str] = []

    for file in pptx_files:
        lower_name = file.name.lower()
        if template is None and ("模板" in file.name or "template" in lower_name):
            template = str(file.resolve())
            continue
        sources.append(str(file.resolve()))

    return template, sources
