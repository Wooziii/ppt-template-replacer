from __future__ import annotations

import argparse
from pathlib import Path

from prototype_tool.config import RenderConfig
from prototype_tool.template_probe import probe_template


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="模板接入探针")
    parser.add_argument("--template", required=True, help="模板 PPTX 路径")
    parser.add_argument("--output-root", default="template_probe_runs", help="输出目录")
    parser.add_argument("--render-width", type=int, default=1600, help="预览图宽度")
    parser.add_argument("--render-height", type=int, default=900, help="预览图高度")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    template = Path(args.template).expanduser()
    output_root = Path(args.output_root).expanduser()
    result = probe_template(
        template,
        output_root,
        render_config=RenderConfig(width=args.render_width, height=args.render_height),
    )
    print(f"[done] {template.name} -> {result['output_dir']}")
    print(f"       report: {result['report_path']}")
    print(f"       profile: {result['profile_json']}")
    print(f"       palette: {result['palette_json']}")


if __name__ == "__main__":
    main()
