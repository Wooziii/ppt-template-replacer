from __future__ import annotations

import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from prototype_tool.cli import parse_args
from prototype_tool.pipeline import _is_transient_com_error
from prototype_tool.render import export_presentation_to_png


class CliAndRenderTests(unittest.TestCase):
    def test_cli_defaults_to_com_engine(self) -> None:
        with patch.object(
            sys,
            "argv",
            ["run_prototype.py", "--template", "template.pptx", "--source", "source.pptx"],
        ):
            args = parse_args()

        self.assertEqual(args.edit_engine, "com")
        self.assertFalse(args.skip_render)

    def test_render_uses_stable_com_helpers(self) -> None:
        fake_pythoncom = types.SimpleNamespace(
            CoInitialize=MagicMock(),
            PumpWaitingMessages=MagicMock(),
        )
        fake_presentation = MagicMock()
        fake_app = MagicMock()

        with tempfile.TemporaryDirectory() as temp_dir:
            pptx_path = Path(temp_dir) / "deck.pptx"
            output_dir = Path(temp_dir) / "exports"
            pptx_path.write_bytes(b"pptx")

            with patch.dict(sys.modules, {"pythoncom": fake_pythoncom}):
                with patch("prototype_tool.render.create_powerpoint_app", return_value=fake_app) as create_app:
                    with patch("prototype_tool.render.open_presentation", return_value=fake_presentation) as open_prs:
                        with patch("prototype_tool.render.close_presentation") as close_prs:
                            with patch("prototype_tool.render.list_powerpoint_pids", return_value=set()):
                                with patch("prototype_tool.render.detect_new_powerpoint_pid", return_value=1234):
                                    with patch("prototype_tool.render.close_powerpoint_app") as close_app:
                                        with patch(
                                            "prototype_tool.render.retry_com",
                                            side_effect=lambda action, **_: action(),
                                        ) as retry:
                                            with patch("prototype_tool.render.time.sleep"):
                                                export_presentation_to_png(pptx_path, output_dir)

        create_app.assert_called_once()
        open_prs.assert_called_once_with(fake_app, pptx_path)
        fake_presentation.Export.assert_called_once_with(str(output_dir), "PNG", 1600, 900)
        close_prs.assert_called_once_with(fake_presentation)
        close_app.assert_called_once_with(fake_app, 1234)
        self.assertGreaterEqual(retry.call_count, 1)

    def test_transient_com_error_detection_matches_busy_rejection(self) -> None:
        self.assertTrue(_is_transient_com_error(RuntimeError("(-2147418111, '被呼叫方拒绝接收呼叫。', None, None)")))
        self.assertFalse(_is_transient_com_error(RuntimeError("模板版式不存在")))


if __name__ == "__main__":
    unittest.main()
