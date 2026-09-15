from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class TemplatePalette:
    background_dark: str = "0B0A18"
    title_cyan: str = "42D7FF"
    accent_orange: str = "FF4A2D"
    text_light: str = "F3F6FB"
    text_dark: str = "1E293B"
    text_muted: str = "64748B"
    panel_dark: str = "111827"
    panel_mid: str = "16243A"
    panel_alt: str = "1F3554"
    panel_teal: str = "153634"
    panel_warm: str = "35241D"
    border_muted: str = "2E597C"
    chart_teal: str = "55D6C2"
    chart_gold: str = "D9B763"
    chart_blue: str = "6FA7FF"
    body_font: str = "腾讯体 W7"
    heading_font: str = "腾讯体 W7"
    chart_series: tuple[str, ...] = ("42D7FF", "FF6A3D", "D9B763", "55D6C2", "6FA7FF")


@dataclass(frozen=True)
class RenderConfig:
    width: int = 1600
    height: int = 900
    retries: int = 6


@dataclass(frozen=True)
class TemplateProfile:
    header_left_ratio: float = 0.07
    header_top_ratio: float = 0.033
    header_width_ratio: float = 0.52
    header_height_ratio: float = 0.07
    accent_left_ratio: float = 0.022
    accent_top_ratio: float = 0.041
    accent_width_ratio: float = 0.018
    accent_height_ratio: float = 0.014
    safe_zone_bottom_ratio: float = 0.16


@dataclass(frozen=True)
class PrototypeOptions:
    template_path: Path
    source_path: Path
    output_root: Path
    max_report_slides: int = 10
    render_config: RenderConfig = RenderConfig()
    palette: TemplatePalette = TemplatePalette()
    palette_metadata: dict[str, object] = field(default_factory=dict)
    profile: TemplateProfile = TemplateProfile()
    edit_engine: str = "com"
    skip_render: bool = False


DEFAULT_SAFE_ZONE_BOTTOM_RATIO = 0.16
