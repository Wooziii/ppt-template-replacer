from pathlib import Path
import unittest

from prototype_tool.template_assets import extract_template_assets
from prototype_tool.template_theme import extract_template_theme, hex_to_rgb_tuple, luminance


BASE_DIR = Path(__file__).resolve().parents[1]


class TemplateThemeTests(unittest.TestCase):
    def test_extract_template_theme_uses_header_color_and_background(self) -> None:
        template_path = BASE_DIR / "OpenClaw TVP研讨会-PPT模板.pptx"
        assets = extract_template_assets(template_path)
        theme = extract_template_theme(template_path, assets)

        self.assertEqual(theme.palette.title_cyan, assets.header_style_compact.color_hex)
        self.assertTrue(theme.palette.background_dark)
        self.assertEqual(theme.palette.heading_font, "腾讯体 W7")
        self.assertGreaterEqual(len(theme.palette.chart_series), 5)
        self.assertIn("background_candidates", theme.metadata)

    def test_light_template_gets_soft_surfaces_and_non_accent_body_text(self) -> None:
        template_path = BASE_DIR / "上海-架构师城市沙龙-PPT模板.pptx"
        assets = extract_template_assets(template_path)
        theme = extract_template_theme(template_path, assets)

        self.assertFalse(theme.metadata["is_dark_theme"])
        self.assertNotEqual(theme.palette.text_light, theme.palette.title_cyan)
        self.assertLess(luminance(hex_to_rgb_tuple(theme.palette.text_dark)), 0.4)
        self.assertGreater(luminance(hex_to_rgb_tuple(theme.palette.panel_mid)), 0.82)
        self.assertNotEqual(theme.palette.border_muted, theme.palette.title_cyan)


if __name__ == "__main__":
    unittest.main()
