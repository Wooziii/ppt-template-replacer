from pathlib import Path
import unittest

from prototype_tool.analysis import analyze_deck, detect_title_candidate
from prototype_tool.template_assets import choose_template_layout_name, extract_template_assets
from prototype_tool.template_profile import extract_template_profile
from pptx import Presentation


BASE_DIR = Path(__file__).resolve().parents[1]


class AnalysisTests(unittest.TestCase):
    def test_title_detection_known_slides(self) -> None:
        prs = Presentation(BASE_DIR / "EverMemOS_Introduction.pptx")
        self.assertEqual(detect_title_candidate(prs, 2)["title"], "Founder's Vision")
        self.assertEqual(detect_title_candidate(prs, 5)["title"], "Industry Background")

        prs2 = Presentation(BASE_DIR / "手把手带你玩转腾讯版自研小龙虾 WorkBuddy.pptx")
        self.assertEqual(
            detect_title_candidate(prs2, 4)["title"],
            "开场用 AI  Coding 一个抽奖应用，从这个小应用开始",
        )

    def test_analyze_deck_flags_complex_slides(self) -> None:
        analysis = analyze_deck(
            BASE_DIR / "手把手带你玩转腾讯版自研小龙虾 WorkBuddy.pptx",
            BASE_DIR / "OpenClaw TVP研讨会-PPT模板.pptx",
        )
        slide4 = next(item for item in analysis["slides"] if item["slide_index"] == 4)
        self.assertIn("image_heavy", slide4["risks"])
        self.assertGreaterEqual(analysis["slide_count"], 40)

    def test_template_profile_extraction(self) -> None:
        profile = extract_template_profile(BASE_DIR / "OpenClaw TVP研讨会-PPT模板.pptx")
        self.assertGreater(profile.header_left_ratio, 0.04)
        self.assertLess(profile.header_left_ratio, 0.12)
        self.assertGreater(profile.header_width_ratio, 0.18)
        self.assertLess(profile.header_width_ratio, 0.35)
        self.assertGreater(profile.safe_zone_bottom_ratio, 0.10)
        self.assertLess(profile.safe_zone_bottom_ratio, 0.22)

    def test_template_assets_capture_background_and_header_style(self) -> None:
        assets = extract_template_assets(BASE_DIR / "OpenClaw TVP研讨会-PPT模板.pptx")
        self.assertIsNotNone(assets.cover_background_blob)
        self.assertIsNotNone(assets.content_background_blob)
        self.assertIsNotNone(assets.section_background_blob)
        self.assertEqual(assets.header_style_compact.font_name, "腾讯体 W7")
        self.assertAlmostEqual(assets.header_style_compact.font_size_pt, 32.0, delta=0.2)
        self.assertGreater(assets.header_style_wide.width_ratio, assets.header_style_compact.width_ratio)

    def test_template_assets_capture_native_layout_hints(self) -> None:
        assets = extract_template_assets(BASE_DIR / "上海-架构师城市沙龙-PPT模板.pptx")
        self.assertEqual(assets.cover_layout_name, "封底")
        self.assertEqual(assets.toc_layout_name, "1_自定义版式")
        self.assertEqual(assets.chapter_layout_name, "2_自定义版式")
        self.assertEqual(assets.content_layout_name, "自定义版式")

        self.assertEqual(
            choose_template_layout_name(
                {"slide_index": 2, "role": "content", "title_detection": {"title": "目录"}},
                assets,
            ),
            "1_自定义版式",
        )
        self.assertEqual(
            choose_template_layout_name(
                {"slide_index": 21, "role": "section", "title_detection": {"title": "THANKS"}},
                assets,
            ),
            "封底",
        )


if __name__ == "__main__":
    unittest.main()
