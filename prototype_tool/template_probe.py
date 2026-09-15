from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Callable

from .config import RenderConfig
from .render import export_presentation_to_png
from .reporting import make_montage
from .template_assets import extract_template_assets
from .template_profile import extract_template_profile, write_template_profile_json
from .template_theme import extract_template_theme, write_template_palette_json


ProgressCallback = Callable[[str], None] | None


def _emit(callback: ProgressCallback, message: str) -> None:
    if callback:
        callback(message)


def write_template_probe_report(
    output_dir: Path,
    template_path: Path,
    profile_json: Path,
    palette_json: Path,
    assets,
    theme,
    montage_path: Path | None,
    export_error: str | None,
) -> Path:
    report_path = output_dir / "模板接入报告.md"
    lines = [
        "# 模板接入报告",
        "",
        f"- 模板：[{template_path.name}]({template_path.as_posix()})",
        f"- profile：[{profile_json.name}]({profile_json.as_posix()})",
        f"- palette：[{palette_json.name}]({palette_json.as_posix()})",
        f"- 背景图：cover=`{bool(assets.cover_background_blob)}` content=`{bool(assets.content_background_blob)}` section=`{bool(assets.section_background_blob)}`",
        f"- 紧凑抬头：left=`{assets.header_style_compact.left_ratio}` top=`{assets.header_style_compact.top_ratio}` width=`{assets.header_style_compact.width_ratio}` color=`{assets.header_style_compact.color_hex}`",
        f"- 宽版抬头：left=`{assets.header_style_wide.left_ratio}` top=`{assets.header_style_wide.top_ratio}` width=`{assets.header_style_wide.width_ratio}` color=`{assets.header_style_wide.color_hex}`",
        f"- 背景主色：`{theme.palette.background_dark}`",
        f"- 标题主色：`{theme.palette.title_cyan}`",
        f"- 强调色：`{theme.palette.accent_orange}`",
        "",
        "## 模板预览",
        "",
    ]
    if montage_path and montage_path.exists():
        lines.extend([f"![template montage]({montage_path.as_posix()})", ""])
    elif export_error:
        lines.extend([f"- 预览图导出失败：`{export_error}`", ""])
    lines.extend([
        "## 主题抽取元数据",
        "",
        "```json",
        json.dumps(theme.metadata, ensure_ascii=False, indent=2),
        "```",
    ])
    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path


def probe_template(template_path: Path, output_root: Path, render_config: RenderConfig | None = None, progress: ProgressCallback = None) -> dict[str, Any]:
    render_config = render_config or RenderConfig()
    output_dir = output_root / template_path.stem
    export_dir = output_dir / "template_exports"
    output_dir.mkdir(parents=True, exist_ok=True)

    _emit(progress, f"抽取模板资产: {template_path.name}")
    assets = extract_template_assets(template_path)
    profile = extract_template_profile(template_path)
    theme = extract_template_theme(template_path, assets)
    profile_json = write_template_profile_json(output_dir, profile)
    palette_json = write_template_palette_json(output_dir, theme)

    montage_path = None
    export_error = None
    try:
        _emit(progress, "导出模板预览图")
        export_presentation_to_png(
            template_path,
            export_dir,
            width=render_config.width,
            height=render_config.height,
            retries=max(6, render_config.retries),
        )
        preview_images = sorted(export_dir.glob("*.PNG"))[:6]
        if preview_images:
            montage_path = output_dir / "template_montage.jpg"
            make_montage(preview_images, montage_path, template_path.stem)
    except Exception as exc:
        export_error = str(exc)
        _emit(progress, f"模板预览图导出失败，保留接入结果: {exc}")

    _emit(progress, "生成模板接入报告")
    report_path = write_template_probe_report(output_dir, template_path, profile_json, palette_json, assets, theme, montage_path, export_error)
    return {
        "template": str(template_path),
        "output_dir": str(output_dir),
        "report_path": str(report_path),
        "profile_json": str(profile_json),
        "palette_json": str(palette_json),
        "export_error": export_error,
        "theme": {
            "palette": theme.palette.__dict__,
            "metadata": theme.metadata,
        },
    }
