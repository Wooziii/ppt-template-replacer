from __future__ import annotations

import argparse
from pathlib import Path

from .config import PrototypeOptions
from .pipeline import process_source
from .template_profile import extract_template_profile, write_template_profile_json
from .template_theme import extract_template_theme
from .workspace import default_workspace_root, derive_replace_output_root


def ensure_path(path_str: str) -> Path:
    path = Path(path_str).expanduser()
    return path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PPT 模板替换第一版原型")
    parser.add_argument("--template", required=True, help="模板 PPTX 路径")
    parser.add_argument(
        "--source", required=True, nargs="+", help="一个或多个源 PPTX 路径"
    )
    parser.add_argument("--output-root", help="输出目录；默认写入文档\\PPT模板替换输出")
    parser.add_argument(
        "--max-report-slides", type=int, default=10, help="报告中展示的最大页数"
    )
    parser.add_argument("--render-width", type=int, default=1600, help="渲染宽度")
    parser.add_argument("--render-height", type=int, default=900, help="渲染高度")
    parser.add_argument(
        "--edit-engine",
        choices=["pptx", "com"],
        default="com",
        help="编辑引擎：默认 com，用于原生 PowerPoint 模板/母版应用；pptx 为无母版兜底路径",
    )
    parser.add_argument(
        "--skip-render",
        action="store_true",
        help="跳过 PowerPoint COM 截图导出步骤，只生成 PPTX；推荐与 com 一起使用以提升稳定性",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    template = ensure_path(args.template)
    sources = [ensure_path(item) for item in args.source]
    if args.output_root:
        output_root = ensure_path(args.output_root)
    else:
        output_root = derive_replace_output_root(default_workspace_root(), template)
    output_root.mkdir(parents=True, exist_ok=True)
    theme = extract_template_theme(template)
    profile = extract_template_profile(template)
    summaries = []
    for source in sources:
        options = PrototypeOptions(
            template_path=template,
            source_path=source,
            output_root=output_root,
            max_report_slides=args.max_report_slides,
            palette=theme.palette,
            palette_metadata=theme.metadata,
            profile=profile,
            edit_engine=args.edit_engine,
            skip_render=args.skip_render,
        )
        result = process_source(options)
        summaries.append(result)
        print(f"[done] {source.name} -> {result['prototype_pptx']}")
        print(f"       report: {result['report_path']}")
        print(f"       risks: {result['risk_slides']}")
        print(f"       engine: {result['engine_used']}")
    if len(summaries) > 1:
        print(f"[summary] processed {len(summaries)} decks into {output_root}")
