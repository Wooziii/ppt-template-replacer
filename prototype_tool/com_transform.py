from __future__ import annotations

import csv
import math
import subprocess
import shutil
import time
from pathlib import Path
from typing import Any

from .template_assets import (
    TemplateAssets,
    choose_template_layout_name,
    extract_template_assets,
)


EMU_PER_PT = 12700
POWERPOINT_TYPELIB = ("{91493440-5A91-11CF-8700-00AA0060263B}", 0, 2, 12)


def rgb_int(hex_color: str) -> int:
    hex_color = hex_color.lstrip("#")
    red = int(hex_color[0:2], 16)
    green = int(hex_color[2:4], 16)
    blue = int(hex_color[4:6], 16)
    return red + (green << 8) + (blue << 16)


def rgb_tuple(rgb_value: int) -> tuple[int, int, int]:
    red = rgb_value & 0xFF
    green = (rgb_value >> 8) & 0xFF
    blue = (rgb_value >> 16) & 0xFF
    return red, green, blue


def luminance(color: tuple[int, int, int]) -> float:
    red, green, blue = [channel / 255.0 for channel in color]
    return 0.2126 * red + 0.7152 * green + 0.0722 * blue


def retry_com(action, retries: int = 15, delay_s: float = 0.15):
    import pythoncom

    last_error = None
    for attempt in range(1, retries + 1):
        try:
            return action()
        except Exception as exc:  # pragma: no cover - PowerPoint COM timing is environment-specific.
            last_error = exc
            pythoncom.PumpWaitingMessages()
            time.sleep(delay_s * attempt)
    raise last_error


def create_powerpoint_app():
    import win32com.client
    from win32com.client import gencache

    gencache.EnsureModule(*POWERPOINT_TYPELIB)
    app = retry_com(lambda: win32com.client.DispatchEx("PowerPoint.Application"), retries=20, delay_s=0.2)
    try:
        retry_com(lambda: setattr(app, "DisplayAlerts", 0), retries=10, delay_s=0.2)
    except Exception:
        pass
    retry_com(lambda: app.Presentations.Count, retries=20, delay_s=0.2)
    return app


def list_powerpoint_pids() -> set[int]:
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError:
        return set()

    pids: set[int] = set()
    for raw_line in result.stdout.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("INFO:"):
            continue
        try:
            row = next(csv.reader([line]))
        except Exception:
            continue
        if len(row) < 2 or not row[0].upper().startswith("POWERPNT"):
            continue
        try:
            pids.add(int(row[1]))
        except ValueError:
            continue
    return pids


def detect_new_powerpoint_pid(before_pids: set[int], timeout_s: float = 5.0) -> int | None:
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        new_pids = list_powerpoint_pids() - before_pids
        if new_pids:
            return sorted(new_pids)[-1]
        time.sleep(0.2)
    return None


def force_kill_powerpoint_pid(app_pid: int | None) -> None:
    if app_pid is None:
        return
    try:
        subprocess.run(
            ["taskkill", "/PID", str(app_pid), "/T", "/F"],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError:
        return


def close_powerpoint_app(app, app_pid: int | None = None) -> None:
    if app is not None:
        try:
            retry_com(lambda: app.Quit(), retries=20, delay_s=0.2)
        except Exception:
            force_kill_powerpoint_pid(app_pid)
            return
    if app_pid is not None:
        deadline = time.time() + 3.0
        while time.time() < deadline:
            if app_pid not in list_powerpoint_pids():
                return
            time.sleep(0.2)
        force_kill_powerpoint_pid(app_pid)


def has_text(shape) -> bool:
    try:
        return bool(shape.HasTextFrame) and bool(shape.TextFrame.HasText)
    except Exception:
        return False


def get_text(shape) -> str:
    try:
        return shape.TextFrame.TextRange.Text.strip()
    except Exception:
        return ""


def get_shape_fill_rgb(shape) -> int | None:
    try:
        if shape.Fill.Visible:
            return shape.Fill.ForeColor.RGB
    except Exception:
        pass
    return None


def intersection_ratio(shape_a, shape_b) -> float:
    left = max(shape_a.Left, shape_b.Left)
    top = max(shape_a.Top, shape_b.Top)
    right = min(shape_a.Left + shape_a.Width, shape_b.Left + shape_b.Width)
    bottom = min(shape_a.Top + shape_a.Height, shape_b.Top + shape_b.Height)
    if right <= left or bottom <= top:
        return 0.0
    inter = (right - left) * (bottom - top)
    area = max(1.0, shape_a.Width * shape_a.Height)
    return inter / area


def shape_local_bg_is_light(shape, underlying_shapes: list, fallback_rgb: int) -> bool:
    color = get_shape_fill_rgb(shape)
    if color is not None:
        return luminance(rgb_tuple(color)) > 0.75
    for under in reversed(underlying_shapes):
        if getattr(under, "Id", None) == getattr(shape, "Id", None):
            continue
        try:
            if intersection_ratio(shape, under) < 0.45:
                continue
        except Exception:
            continue
        color = get_shape_fill_rgb(under)
        if color is not None:
            return luminance(rgb_tuple(color)) > 0.75
    return luminance(rgb_tuple(fallback_rgb)) > 0.55


def set_text_color(shape, hex_color: str) -> None:
    if not has_text(shape):
        return
    retry_com(lambda: setattr(shape.TextFrame.TextRange.Font.Color, "RGB", rgb_int(hex_color)))


def set_text_font(shape, font_name: str) -> None:
    if not has_text(shape):
        return
    for attribute in ("Name", "NameAscii", "NameFarEast", "NameOther"):
        try:
            retry_com(lambda attr=attribute: setattr(shape.TextFrame.TextRange.Font, attr, font_name))
        except Exception:
            continue


def normalize_header_text(text: str) -> str:
    return " ".join(text.split())


def estimate_text_units(text: str) -> float:
    units = 0.0
    for character in text:
        if ord(character) > 127:
            units += 1.0
        elif character.isupper():
            units += 0.72
        elif character.islower():
            units += 0.58
        elif character.isdigit():
            units += 0.6
        elif character in {" ", "/", "-", "_", ".", ",", ":", ";", "|"}:
            units += 0.32
        else:
            units += 0.5
    return units


def choose_header_style(title_text: str, assets: TemplateAssets):
    normalized = normalize_header_text(title_text)
    if len(normalized) > 18 or any(character.isascii() for character in normalized):
        return assets.header_style_wide
    return assets.header_style_compact


def choose_header_width_ratio(style, title_text: str) -> float:
    width_ratio = style.width_ratio
    if width_ratio < 0.35:
        return width_ratio
    normalized_text = normalize_header_text(title_text)
    text_units = estimate_text_units(normalized_text)
    ascii_chars = sum(1 for character in normalized_text if character.isascii() and not character.isspace())
    visible_chars = sum(1 for character in normalized_text if not character.isspace())
    ascii_ratio = ascii_chars / max(1, visible_chars)
    if ascii_ratio > 0.35 and text_units > 14:
        width_ratio += 0.05 + min(0.04, max(0.0, text_units - 14) * 0.003)
    elif text_units > 18:
        width_ratio += 0.04
    return round(min(0.7, width_ratio), 4)


def choose_fitted_header_font_size(style, slide_width: float, title_text: str, width_ratio: float | None = None) -> float:
    font_size = style.font_size_pt
    normalized_text = normalize_header_text(title_text)
    resolved_width_ratio = width_ratio if width_ratio is not None else choose_header_width_ratio(style, title_text)
    box_width_pt = ((slide_width * resolved_width_ratio) - style.margin_left - style.margin_right) / EMU_PER_PT
    text_units = estimate_text_units(normalized_text)
    ascii_chars = sum(1 for character in normalized_text if character.isascii() and not character.isspace())
    visible_chars = sum(1 for character in normalized_text if not character.isspace())
    ascii_ratio = ascii_chars / max(1, visible_chars)
    if ascii_ratio > 0.75:
        if visible_chars > 28:
            font_size = min(font_size, 24.0)
        elif visible_chars > 22:
            font_size = min(font_size, 27.0)
    if ascii_ratio > 0.75:
        width_factor = 0.78
    elif ascii_ratio > 0.35:
        width_factor = 0.72
    else:
        width_factor = 0.62
    single_line_min_size = 18.0 if ascii_ratio > 0.35 else 22.0
    while font_size > single_line_min_size:
        capacity = max(1.0, box_width_pt / (font_size * width_factor))
        if math.ceil(text_units / capacity) <= 1:
            return round(font_size, 2)
        font_size -= 1.0

    font_size = min(single_line_min_size, style.font_size_pt)
    two_line_min_size = 16.0 if ascii_ratio > 0.35 else 19.0
    while font_size > two_line_min_size:
        capacity = max(1.0, box_width_pt / (font_size * width_factor))
        if math.ceil(text_units / capacity) <= 2:
            return round(font_size, 2)
        font_size -= 1.0
    return round(two_line_min_size, 2)


def open_presentation(app, path: Path):
    retry_com(lambda: app.Presentations.Open(str(path), False, False, False), retries=20, delay_s=0.2)
    return retry_com(lambda: app.Presentations.Item(app.Presentations.Count), retries=20, delay_s=0.2)


def close_presentation(presentation) -> None:
    retry_com(lambda: presentation.Close(), retries=12, delay_s=0.2)


def layout_name(slide) -> str:
    try:
        return slide.CustomLayout.Name or ""
    except Exception:
        return ""


def apply_template_file(presentation, template_path: Path) -> None:
    retry_com(lambda: presentation.ApplyTemplate(str(template_path)), retries=20, delay_s=0.2)
    retry_com(presentation.Save, retries=20, delay_s=0.2)


def apply_template_layouts(presentation, analysis: dict[str, Any], assets: TemplateAssets) -> dict[int, dict[str, str]]:
    layout_map: dict[str, Any] = {}
    for idx in range(1, presentation.Slides.Count + 1):
        slide = retry_com(lambda slide_idx=idx: presentation.Slides(slide_idx))
        name = layout_name(slide)
        if name and name not in layout_map:
            layout_map[name] = slide.CustomLayout

    applied_layouts: dict[int, dict[str, str]] = {}
    for slide_analysis in analysis["slides"]:
        target_name = choose_template_layout_name(slide_analysis, assets)
        if not target_name or target_name not in layout_map:
            continue
        slide_idx = slide_analysis["slide_index"]
        slide = retry_com(lambda idx=slide_idx: presentation.Slides(idx))
        before_name = layout_name(slide)
        if before_name != target_name:
            retry_com(lambda s=slide, custom_layout=layout_map[target_name]: setattr(s, "CustomLayout", custom_layout))
        try:
            retry_com(lambda s=slide: setattr(s, "FollowMasterBackground", True))
        except Exception:
            pass
        try:
            retry_com(lambda s=slide: setattr(s, "DisplayMasterShapes", True))
        except Exception:
            pass
        applied_layouts[slide_idx] = {"before": before_name, "after": target_name}
    return applied_layouts


def add_header(slide, slide_width: float, slide_height: float, title_text: str, assets: TemplateAssets) -> dict[str, Any]:
    normalized_title = normalize_header_text(title_text)
    if not normalized_title:
        return {}
    style = choose_header_style(title_text, assets)
    header_width_ratio = choose_header_width_ratio(style, normalized_title)
    fitted_font_size = choose_fitted_header_font_size(style, slide_width, normalized_title, width_ratio=header_width_ratio)
    header = retry_com(
        lambda: slide.Shapes.AddTextbox(
            1,
            slide_width * style.left_ratio,
            slide_height * style.top_ratio,
            slide_width * header_width_ratio,
            slide_height * style.height_ratio,
        )
    )
    retry_com(lambda: setattr(header.Fill, "Visible", False))
    retry_com(lambda: setattr(header.Line, "Visible", False))
    retry_com(lambda: setattr(header.TextFrame, "MarginLeft", style.margin_left / EMU_PER_PT))
    retry_com(lambda: setattr(header.TextFrame, "MarginRight", style.margin_right / EMU_PER_PT))
    retry_com(lambda: setattr(header.TextFrame, "MarginTop", style.margin_top / EMU_PER_PT))
    retry_com(lambda: setattr(header.TextFrame, "MarginBottom", style.margin_bottom / EMU_PER_PT))
    retry_com(lambda: setattr(header.TextFrame.TextRange, "Text", normalized_title))
    retry_com(lambda: setattr(header.TextFrame.TextRange.ParagraphFormat, "Alignment", 1))
    retry_com(lambda: setattr(header.TextFrame.TextRange.Font, "Bold", True))
    retry_com(lambda: setattr(header.TextFrame.TextRange.Font, "Size", fitted_font_size))
    retry_com(lambda: setattr(header.TextFrame.TextRange.Font, "Name", style.font_name))
    for attribute in ("NameAscii", "NameFarEast", "NameOther"):
        try:
            retry_com(lambda attr=attribute: setattr(header.TextFrame.TextRange.Font, attr, style.font_name))
        except Exception:
            continue
    retry_com(lambda: setattr(header.TextFrame.TextRange.Font.Color, "RGB", rgb_int(style.color_hex)))
    return {
        "font_name": style.font_name,
        "font_size_pt": fitted_font_size,
        "color_hex": style.color_hex,
        "style": "wide" if style == assets.header_style_wide else "compact",
        "width_ratio": header_width_ratio,
    }


def iter_shapes(slide) -> list:
    return [slide.Shapes(index) for index in range(1, slide.Shapes.Count + 1)]


def style_cover_slide(slide, slide_width: float, slide_height: float, title_text: str, palette) -> dict[str, Any]:
    for shape in iter_shapes(slide):
        if has_text(shape) and get_text(shape):
            set_text_color(shape, palette.text_light)
            set_text_font(shape, palette.body_font)
    return {"mode": "cover_preview", "engine": "com", "native_master_applied": True, "title_kept": bool(title_text)}


def style_section_slide(slide, slide_width: float, slide_height: float, palette) -> dict[str, Any]:
    for shape in iter_shapes(slide):
        if has_text(shape) and get_text(shape):
            set_text_color(shape, palette.text_light)
            set_text_font(shape, palette.body_font)
    return {"mode": "section_preview", "engine": "com", "native_master_applied": True}


def style_content_slide(
    slide,
    slide_width: float,
    slide_height: float,
    slide_analysis: dict[str, Any],
    palette,
    profile,
    assets: TemplateAssets,
) -> dict[str, Any]:
    bg_rgb = rgb_int(palette.background_dark)
    safe_zone_bottom = slide_height * profile.safe_zone_bottom_ratio
    title_shape_index = slide_analysis["title_detection"]["shape_index"]
    title_text = slide_analysis["title_detection"]["title"]
    moved_shapes = []

    # Delete title first while indices still align with analysis.
    if title_shape_index and title_shape_index <= slide.Shapes.Count:
        title_shape = slide.Shapes(title_shape_index)
        if has_text(title_shape) and get_text(title_shape):
            retry_com(title_shape.Delete)

    shapes = iter_shapes(slide)
    for idx, shape in enumerate(shapes):
        text = get_text(shape)
        if not text:
            continue
        if shape.Top < safe_zone_bottom and shape.Left < slide_width * 0.62:
            delta = safe_zone_bottom - shape.Top + slide_height * 0.015
            new_top = min(shape.Top + delta, max(0, slide_height - shape.Height))
            retry_com(lambda target=new_top: setattr(shape, "Top", target))
            moved_shapes.append({"text": text[:60], "delta": round(delta, 2)})
        is_light = shape_local_bg_is_light(shape, shapes[:idx], bg_rgb)
        set_text_color(shape, palette.text_dark if is_light else palette.text_light)
        set_text_font(shape, palette.body_font)

    header_meta = add_header(slide, slide_width, slide_height, title_text, assets)
    return {
        "mode": "content_preview",
        "engine": "com",
        "native_master_applied": True,
        "header_added": bool(header_meta),
        "header_meta": header_meta,
        "safe_zone_bottom_ratio": profile.safe_zone_bottom_ratio,
        "moved_shapes": moved_shapes,
    }


def transform_and_export_with_com(
    source_path: Path,
    output_pptx: Path,
    export_dir: Path,
    analysis: dict[str, Any],
    palette,
    profile,
    template_path: Path,
    width: int,
    height: int,
    *,
    skip_render: bool = False,
) -> dict[str, Any]:
    import pythoncom

    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    export_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_path, output_pptx)
    time.sleep(0.3)
    assets = extract_template_assets(template_path)

    pythoncom.CoInitialize()
    app = None
    app_pid = None
    presentation = None
    try:
        before_pids = list_powerpoint_pids()
        app = create_powerpoint_app()
        app_pid = detect_new_powerpoint_pid(before_pids)
        presentation = open_presentation(app, output_pptx)
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight
        apply_template_file(presentation, template_path)
        close_presentation(presentation)
        presentation = None
        time.sleep(0.8)

        presentation = open_presentation(app, output_pptx)
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight
        applied_layouts = apply_template_layouts(presentation, analysis, assets)

        transforms = []
        for slide_analysis in analysis["slides"]:
            slide = presentation.Slides(slide_analysis["slide_index"])
            if slide_analysis["role"] == "cover":
                meta = style_cover_slide(
                    slide,
                    slide_width,
                    slide_height,
                    slide_analysis["title_detection"]["title"],
                    palette,
                )
            elif slide_analysis["role"] == "section":
                meta = style_section_slide(slide, slide_width, slide_height, palette)
            else:
                meta = style_content_slide(
                    slide,
                    slide_width,
                    slide_height,
                    slide_analysis,
                    palette,
                    profile,
                    assets,
                )
            if slide_analysis["slide_index"] in applied_layouts:
                meta["layout_applied"] = applied_layouts[slide_analysis["slide_index"]]
            transforms.append({"slide_index": slide_analysis["slide_index"], "transform": meta})

        retry_com(presentation.Save)
        if not skip_render:
            retry_com(lambda: presentation.Export(str(export_dir), "PNG", width, height), retries=20, delay_s=0.2)
        return {
            "output_pptx": str(output_pptx),
            "transforms": transforms,
            "engine": "com",
            "render_skipped": skip_render,
            "template_profile": profile.__dict__,
        }
    finally:
        if presentation is not None:
            try:
                close_presentation(presentation)
            except Exception:
                pass
        if app is not None:
            try:
                close_powerpoint_app(app, app_pid)
            except Exception:
                pass
