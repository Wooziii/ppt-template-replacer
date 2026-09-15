from __future__ import annotations

import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client


RPC_E_CALL_REJECTED = -2147418111
MsoTrue = -1
MsoFalse = 0


def retry_com(func, *args, retries: int = 12, delay: float = 1.0, **kwargs):
    last_error = None
    for _ in range(retries):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as exc:
            last_error = exc
            hresult = exc.hresult if hasattr(exc, "hresult") else exc.args[0]
            if hresult != RPC_E_CALL_REJECTED:
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(delay)
    raise last_error


def build_output_path(source_path: Path, suffix: str) -> Path:
    return source_path.with_name(f"{source_path.stem}{suffix}{source_path.suffix}")


def slide_text(slide) -> str:
    texts: list[str] = []
    count = retry_com(lambda: slide.Shapes.Count)
    for idx in range(1, count + 1):
        shape = retry_com(lambda: slide.Shapes.Item(idx))
        try:
            if retry_com(lambda: shape.HasTextFrame) == MsoTrue and retry_com(lambda: shape.TextFrame.HasText) == MsoTrue:
                text = retry_com(lambda: shape.TextFrame.TextRange.Text).strip()
                if text:
                    texts.append(text)
        finally:
            pass
    return " ".join(texts)


def choose_base_slide(slide_index: int, text: str) -> int:
    lowered = text.lower()
    if slide_index == 1:
        return 1
    if "q&a" in lowered or "问答" in text or "答疑" in text:
        return 11
    if "thank you" in lowered or "thanks" in lowered or "感谢" in text or "谢谢" in text or "结束页" in text:
        return 12
    return 5


def remove_text_shapes(slide) -> None:
    count = retry_com(lambda: slide.Shapes.Count)
    for idx in range(count, 0, -1):
        shape = retry_com(lambda: slide.Shapes.Item(idx))
        try:
            has_text_frame = retry_com(lambda: shape.HasTextFrame) == MsoTrue
            if has_text_frame and retry_com(lambda: shape.TextFrame.HasText) == MsoTrue:
                retry_com(shape.Delete)
        finally:
            pass


def should_skip_shape(shape, slide_width: float, slide_height: float) -> bool:
    try:
        if retry_com(lambda: shape.HasTextFrame) == MsoTrue and retry_com(lambda: shape.TextFrame.HasText) == MsoTrue:
            return False
    except pywintypes.com_error:
        return False

    left = float(retry_com(lambda: shape.Left))
    top = float(retry_com(lambda: shape.Top))
    width = float(retry_com(lambda: shape.Width))
    height = float(retry_com(lambda: shape.Height))
    area_ratio = (width * height) / (slide_width * slide_height)
    return area_ratio >= 0.90 and left <= 5 and top <= 5


def copy_shapes(source_slide, dest_slide, slide_width: float, slide_height: float) -> None:
    count = retry_com(lambda: source_slide.Shapes.Count)
    for idx in range(1, count + 1):
        shape = retry_com(lambda: source_slide.Shapes.Item(idx))
        try:
            if should_skip_shape(shape, slide_width, slide_height):
                continue
            retry_com(shape.Copy)
            time.sleep(0.05)
            retry_com(dest_slide.Shapes.Paste)
        finally:
            pass


def rebuild_deck(app, template_path: Path, source_path: Path, suffix: str) -> tuple[Path, int]:
    output_path = build_output_path(source_path, suffix)
    if output_path.exists():
        output_path.unlink()

    source = None
    output = None
    try:
        source = retry_com(
            lambda: app.Presentations.Open(
                str(source_path),
                MsoTrue,
                MsoFalse,
                MsoFalse,
            )
        )
        output = retry_com(
            lambda: app.Presentations.Open(
                str(template_path),
                MsoTrue,
                MsoTrue,
                MsoFalse,
            )
        )

        slide_width = float(retry_com(lambda: source.PageSetup.SlideWidth))
        slide_height = float(retry_com(lambda: source.PageSetup.SlideHeight))

        template_slide_count = retry_com(lambda: output.Slides.Count)
        source_slide_count = retry_com(lambda: source.Slides.Count)

        for slide_index in range(1, source_slide_count + 1):
            source_slide = retry_com(lambda: source.Slides.Item(slide_index))
            text = slide_text(source_slide)
            base_index = choose_base_slide(slide_index, text)

            duplicate_range = retry_com(lambda: output.Slides.Item(base_index).Duplicate())
            dest_slide = retry_com(lambda: duplicate_range.Item(1))
            retry_com(dest_slide.MoveTo, retry_com(lambda: output.Slides.Count))
            remove_text_shapes(dest_slide)
            copy_shapes(source_slide, dest_slide, slide_width, slide_height)

        for idx in range(template_slide_count, 0, -1):
            retry_com(lambda: output.Slides.Item(idx).Delete())

        retry_com(output.SaveAs, str(output_path))
        return output_path, source_slide_count
    finally:
        if source is not None:
            try:
                retry_com(source.Close)
            except Exception:
                pass
        if output is not None:
            try:
                retry_com(output.Close)
            except Exception:
                pass


def main(argv: list[str]) -> int:
    if len(argv) < 3:
        print("Usage: python rebuild_ppt_on_template.py TEMPLATE SOURCE [SOURCE ...] [--suffix SUFFIX]")
        return 1

    args = argv[1:]
    suffix = "-模板重建版"
    if "--suffix" in args:
        idx = args.index("--suffix")
        suffix = args[idx + 1]
        del args[idx : idx + 2]

    template = Path(args[0]).resolve()
    sources = [Path(item).resolve() for item in args[1:]]

    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.Visible = MsoTrue
        app.DisplayAlerts = 0
        time.sleep(5)

        for source in sources:
            output_path, slides = rebuild_deck(app, template, source, suffix)
            print(f"{source} -> {output_path} ({slides} slides)")
        return 0
    finally:
        if app is not None:
            try:
                retry_com(app.Quit)
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
