from __future__ import annotations

import tempfile
import time
from pathlib import Path

import pythoncom

from prototype_tool.analysis import analyze_deck
from prototype_tool.com_transform import (
    add_header,
    close_powerpoint_app,
    close_presentation,
    create_powerpoint_app,
    get_text,
    has_text,
    open_presentation,
    retry_com,
    rgb_int,
    set_text_color,
    set_text_font,
    shape_local_bg_is_light,
)
from prototype_tool.template_assets import extract_template_assets
from prototype_tool.template_profile import extract_template_profile
from prototype_tool.template_theme import extract_template_theme


WORKDIR = Path(r"D:\Desktop\Projects\PPT模板替换\tmp_ascii\openclaw_editable")
SOURCE_PPT = WORKDIR / "source.pptx"
TEMPLATE_PPT = WORKDIR / "template.pptx"
OUTPUT_PPT = WORKDIR / "output" / "source" / "source_editable_templated.pptx"


PP_LAYOUT_BLANK = 12
MSO_SEND_TO_BACK = 1
MSO_SHAPE_RECTANGLE = 1


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def write_blob(path: Path, blob: bytes | None) -> Path | None:
    if not blob:
        return None
    path.write_bytes(blob)
    return path


def normalize_text(text: str) -> str:
    return " ".join(text.split())


def is_transient_com_error(exc: Exception) -> bool:
    message = str(exc)
    markers = (
        "被呼叫方拒绝接收呼叫",
        "call was rejected by callee",
        "-2147418111",
    )
    return any(marker in message for marker in markers)


def is_background_shape(shape, slide_width: float, slide_height: float) -> bool:
    try:
        if has_text(shape) and get_text(shape):
            return False
    except Exception:
        pass

    left = float(shape.Left)
    top = float(shape.Top)
    width = float(shape.Width)
    height = float(shape.Height)
    area_ratio = (width * height) / max(1.0, slide_width * slide_height)

    if area_ratio >= 0.88 and left <= slide_width * 0.05 and top <= slide_height * 0.05:
        return True

    if not has_text(shape) and int(shape.Type) != 13:
        if width <= slide_width * 0.025 or height <= slide_height * 0.025:
            return True

    return False


def copy_shape(source_slide, dest_slide, shape_index: int) -> None:
    source_shape = None
    pasted_range = None
    try:
        source_shape = retry_com(lambda: source_slide.Shapes(shape_index))
        retry_com(source_shape.Copy, retries=8, delay_s=0.15)
        pythoncom.PumpWaitingMessages()
        time.sleep(0.2)
        pasted_range = retry_com(dest_slide.Shapes.Paste, retries=8, delay_s=0.2)
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)
    finally:
        pasted_range = None
        source_shape = None


def copy_source_shapes(source_slide, dest_slide, slide_width: float, slide_height: float) -> None:
    for shape_index in range(1, source_slide.Shapes.Count + 1):
        shape = retry_com(lambda idx=shape_index: source_slide.Shapes(idx))
        try:
            if is_background_shape(shape, slide_width, slide_height):
                continue
        finally:
            shape = None
        copy_shape(source_slide, dest_slide, shape_index)


def add_background_picture(slide, image_path: Path, slide_width: float, slide_height: float) -> None:
    picture = retry_com(
        lambda: slide.Shapes.AddPicture(
            str(image_path),
            False,
            True,
            0,
            0,
            slide_width,
            slide_height,
        )
    )
    try:
        retry_com(lambda: picture.ZOrder(MSO_SEND_TO_BACK), retries=4, delay_s=0.1)
    except Exception:
        pass


def add_manual_backdrop_if_needed(source_slide, dest_slide, slide_analysis: dict) -> None:
    if slide_analysis.get("slide_index") != 8:
        return

    largest_picture = None
    largest_area = 0
    for shape_index in range(1, source_slide.Shapes.Count + 1):
        shape = retry_com(lambda idx=shape_index: source_slide.Shapes(idx))
        try:
            if int(shape.Type) != 13:
                continue
            area = float(shape.Width) * float(shape.Height)
            if area > largest_area:
                largest_area = area
                largest_picture = (
                    float(shape.Left),
                    float(shape.Top),
                    float(shape.Width),
                    float(shape.Height),
                )
        finally:
            shape = None

    if largest_picture is None:
        return

    left, top, width, height = largest_picture
    backdrop = retry_com(lambda: dest_slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, left, top, width, height))
    retry_com(lambda: setattr(backdrop.Fill, "Visible", True))
    retry_com(lambda: setattr(backdrop.Fill.ForeColor, "RGB", rgb_int("5B5D62")))
    retry_com(lambda: setattr(backdrop.Line, "Visible", False))


def find_title_shape(slide, title_text: str):
    normalized_title = normalize_text(title_text)
    if not normalized_title:
        return None

    for index in range(1, slide.Shapes.Count + 1):
        shape = retry_com(lambda idx=index: slide.Shapes(idx))
        try:
            if not has_text(shape):
                continue
            current = normalize_text(get_text(shape))
            if not current:
                continue
            if current == normalized_title or normalized_title in current or current in normalized_title:
                return shape
        except Exception:
            continue
    return None


def restyle_slide(slide, slide_analysis: dict, palette, profile, assets, slide_width: float, slide_height: float) -> None:
    title_text = (slide_analysis.get("title_detection") or {}).get("title", "").strip()
    role = slide_analysis.get("role", "content")

    if role != "cover" and title_text:
        title_shape = find_title_shape(slide, title_text)
        if title_shape is not None:
            retry_com(title_shape.Delete, retries=6, delay_s=0.15)

    shapes = [retry_com(lambda idx=index: slide.Shapes(idx)) for index in range(1, slide.Shapes.Count + 1)]
    fallback_rgb = rgb_int(palette.background_dark)
    safe_zone_bottom = slide_height * profile.safe_zone_bottom_ratio

    for idx, shape in enumerate(shapes):
        text = get_text(shape)
        if not text:
            continue

        if role != "cover" and shape.Top < safe_zone_bottom and shape.Left < slide_width * 0.62:
            delta = safe_zone_bottom - shape.Top + slide_height * 0.01
            new_top = min(shape.Top + delta, max(0, slide_height - shape.Height))
            retry_com(lambda target=new_top, s=shape: setattr(s, "Top", target), retries=6, delay_s=0.1)

        if role == "cover":
            set_text_color(shape, palette.text_dark)
            set_text_font(shape, palette.body_font)
            continue

        is_light = shape_local_bg_is_light(shape, shapes[:idx], fallback_rgb)
        set_text_color(shape, palette.text_dark if is_light else palette.text_light)
        set_text_font(shape, palette.body_font)

    if role != "cover" and title_text:
        add_header(slide, slide_width, slide_height, title_text, assets)


def build_editable_template() -> Path:
    ensure_parent(OUTPUT_PPT)
    analysis = analyze_deck(SOURCE_PPT, TEMPLATE_PPT)
    assets = extract_template_assets(TEMPLATE_PPT)
    theme = extract_template_theme(TEMPLATE_PPT)
    profile = extract_template_profile(TEMPLATE_PPT)

    with tempfile.TemporaryDirectory(prefix="openclaw_bg_") as temp_dir:
        temp_dir_path = Path(temp_dir)
        cover_bg = write_blob(temp_dir_path / "cover_bg.png", assets.cover_background_blob)
        content_bg = write_blob(temp_dir_path / "content_bg.png", assets.content_background_blob)
        section_bg = write_blob(temp_dir_path / "section_bg.png", assets.section_background_blob)

        pythoncom.CoInitialize()
        app = None
        app_pid = None
        source_presentation = None
        output_presentation = None
        try:
            app = create_powerpoint_app()
            source_presentation = open_presentation(app, SOURCE_PPT)
            slide_width = source_presentation.PageSetup.SlideWidth
            slide_height = source_presentation.PageSetup.SlideHeight

            output_presentation = retry_com(lambda: app.Presentations.Add(False), retries=8, delay_s=0.2)
            retry_com(lambda: setattr(output_presentation.PageSetup, "SlideWidth", slide_width))
            retry_com(lambda: setattr(output_presentation.PageSetup, "SlideHeight", slide_height))

            while output_presentation.Slides.Count > 0:
                retry_com(lambda: output_presentation.Slides(1).Delete(), retries=6, delay_s=0.15)

            for slide_analysis in analysis["slides"]:
                slide_index = slide_analysis["slide_index"]
                print(f"processing slide {slide_index}/{len(analysis['slides'])}", flush=True)
                last_error = None
                for attempt in range(1, 4):
                    source_slide = None
                    dest_slide = None
                    try:
                        source_slide = retry_com(lambda idx=slide_index: source_presentation.Slides(idx))
                        dest_slide = retry_com(
                            lambda idx=output_presentation.Slides.Count + 1: output_presentation.Slides.Add(idx, PP_LAYOUT_BLANK),
                            retries=8,
                            delay_s=0.2,
                        )

                        role = slide_analysis.get("role", "content")
                        background = content_bg
                        if role == "cover" and cover_bg is not None:
                            background = cover_bg
                        elif role == "section" and section_bg is not None:
                            background = section_bg

                        if background is not None:
                            add_background_picture(dest_slide, background, slide_width, slide_height)

                        add_manual_backdrop_if_needed(source_slide, dest_slide, slide_analysis)
                        copy_source_shapes(source_slide, dest_slide, slide_width, slide_height)
                        restyle_slide(dest_slide, slide_analysis, theme.palette, profile, assets, slide_width, slide_height)
                        last_error = None
                        break
                    except Exception as exc:
                        last_error = exc
                        if dest_slide is not None:
                            try:
                                retry_com(dest_slide.Delete, retries=4, delay_s=0.1)
                            except Exception:
                                pass
                        if attempt < 3 and is_transient_com_error(exc):
                            time.sleep(0.6 * attempt)
                            continue
                        raise

                if last_error is not None:
                    raise last_error

            print("saving presentation", flush=True)
            retry_com(lambda: output_presentation.SaveAs(str(OUTPUT_PPT), 24), retries=8, delay_s=0.4)
            return OUTPUT_PPT
        finally:
            if source_presentation is not None:
                try:
                    close_presentation(source_presentation)
                except Exception:
                    pass
            if output_presentation is not None:
                try:
                    close_presentation(output_presentation)
                except Exception:
                    pass
            close_powerpoint_app(app, app_pid)
            pythoncom.CoUninitialize()


def main() -> None:
    output = build_editable_template()
    print(output)


if __name__ == "__main__":
    main()
