from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from prototype_tool.workspace import (
    APP_WORKSPACE_NAME,
    app_base_dir,
    default_workspace_root,
    derive_probe_output_root,
    derive_replace_output_root,
    discover_initial_pptx_inputs,
)


class WorkspaceTests(unittest.TestCase):
    def test_default_workspace_prefers_documents_root(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            documents_dir = Path(temp_dir) / "Documents"
            documents_dir.mkdir()
            with patch("prototype_tool.workspace.default_documents_dir", return_value=documents_dir):
                self.assertEqual(default_workspace_root(), documents_dir / APP_WORKSPACE_NAME)

    def test_discover_initial_inputs_picks_first_template_and_other_sources(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "0329-长沙站-PPT模板.pptx"
            source_a = root / "A.pptx"
            source_b = root / "B.pptx"
            for file in (template, source_a, source_b):
                file.write_bytes(b"pptx")

            detected_template, sources = discover_initial_pptx_inputs(root)

        self.assertEqual(detected_template, str(template.resolve()))
        self.assertEqual(sources, [str(source_a.resolve()), str(source_b.resolve())])

    def test_output_roots_are_grouped_by_template_name(self) -> None:
        workspace_root = Path(r"D:\OutputRoot")
        replace_root = derive_replace_output_root(workspace_root, r"D:\Decks\TVP: 模板?.pptx")
        probe_root = derive_probe_output_root(workspace_root)

        self.assertEqual(replace_root, workspace_root / "正式替换" / "TVP_ 模板_")
        self.assertEqual(probe_root, workspace_root / "模板探针")

    def test_app_base_dir_uses_executable_parent_when_frozen(self) -> None:
        fake_executable = Path(r"D:\Apps\PPTTemplateReplacer\PPTTemplateReplacer.exe")
        with patch("prototype_tool.workspace.is_frozen_app", return_value=True):
            with patch.object(sys, "executable", str(fake_executable)):
                self.assertEqual(app_base_dir(), fake_executable.parent)


if __name__ == "__main__":
    unittest.main()
