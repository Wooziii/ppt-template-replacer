from __future__ import annotations

import time
from pathlib import Path

from .com_transform import (
    close_powerpoint_app,
    close_presentation,
    create_powerpoint_app,
    detect_new_powerpoint_pid,
    list_powerpoint_pids,
    open_presentation,
    retry_com,
)


def export_presentation_to_png(pptx_path: Path, output_dir: Path, width: int = 1600, height: int = 900, retries: int = 6) -> None:
    import pythoncom

    output_dir.mkdir(parents=True, exist_ok=True)
    last_error = None
    for attempt in range(1, max(6, retries) + 1):
        pythoncom.CoInitialize()
        app = None
        app_pid = None
        presentation = None
        try:
            before_pids = list_powerpoint_pids()
            app = create_powerpoint_app()
            app_pid = detect_new_powerpoint_pid(before_pids)
            presentation = open_presentation(app, pptx_path)
            pythoncom.PumpWaitingMessages()
            time.sleep(0.4 * attempt)
            retry_com(lambda: presentation.Export(str(output_dir), "PNG", width, height), retries=20, delay_s=0.2)
            close_presentation(presentation)
            presentation = None
            close_powerpoint_app(app, app_pid)
            app = None
            return
        except Exception as exc:  # pragma: no cover - COM errors are environment-specific.
            last_error = exc
            pythoncom.PumpWaitingMessages()
            time.sleep(1.2 * attempt)
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
    raise RuntimeError(f"PowerPoint export failed after {retries} attempts: {last_error}")
