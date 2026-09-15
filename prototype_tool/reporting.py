from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Iterable

from PIL import Image, ImageDraw
from pptx import Presentation


def annotate_title_image(src_image: Path, out_image: Path, prs: Presentation, detection: dict[str, Any]) -> None:
    if not detection["candidates"]:
        image = Image.open(src_image).convert("RGB")
        image.save(out_image, quality=92)
        return
    image = Image.open(src_image).convert("RGB")
    draw = ImageDraw.Draw(image)
    slide_w = prs.slide_width
    slide_h = prs.slide_height
    bbox = detection["candidates"][0]["bbox"]
    left = int(bbox["left"] / slide_w * image.width)
    top = int(bbox["top"] / slide_h * image.height)
    width = int(bbox["width"] / slide_w * image.width)
    height = int(bbox["height"] / slide_h * image.height)
    rect_left = max(0, min(image.width - 1, left))
    rect_top = max(0, min(image.height - 1, top))
    rect_right = max(rect_left + 1, min(image.width, left + width))
    rect_bottom = max(rect_top + 1, min(image.height, top + height))
    draw.rectangle([rect_left, rect_top, rect_right, rect_bottom], outline="#FF4A2D", width=6)
    label = f"title: {detection['title'][:48]}  conf={detection['confidence']}"
    label_left = rect_left
    label_right = min(image.width, rect_left + 700)
    if rect_top >= 36:
        label_top = rect_top - 36
        label_bottom = rect_top
    else:
        label_top = rect_bottom
        label_bottom = min(image.height, rect_bottom + 36)
    if label_bottom > label_top:
        draw.rectangle([label_left, label_top, label_right, label_bottom], fill="#0B0A18")
        draw.text((label_left + 8, label_top + 6), label, fill="#42D7FF")
    image.save(out_image, quality=92)


def build_before_after_board(before_path: Path, after_path: Path, board_path: Path, title: str) -> None:
    before = Image.open(before_path).convert("RGB")
    after = Image.open(after_path).convert("RGB")
    card_w = 820
    card_h = 520
    canvas = Image.new("RGB", (card_w * 2 + 60, card_h + 120), "#10131a")
    draw = ImageDraw.Draw(canvas)

    for idx, img in enumerate([before, after]):
        img.thumbnail((card_w, card_h - 40))
        x = 20 + idx * (card_w + 20)
        y = 60
        panel = Image.new("RGB", (card_w, card_h), "#171c26")
        panel.paste(img, ((card_w - img.width) // 2, 20))
        canvas.paste(panel, (x, y))
        draw.text((x + 16, y + card_h - 18), "Before" if idx == 0 else "After", fill="#f3f6fb")

    draw.text((20, 18), title, fill="#42D7FF")
    canvas.save(board_path, quality=92)


def make_montage(images: Iterable[Path], out_path: Path, label_prefix: str) -> None:
    images = list(images)
    if not images:
        return
    thumbs = []
    for idx, image_path in enumerate(images, start=1):
        image = Image.open(image_path).convert("RGB")
        image.thumbnail((360, 202))
        card = Image.new("RGB", (380, 250), "white")
        card.paste(image, ((380 - image.width) // 2, 10))
        draw = ImageDraw.Draw(card)
        draw.text((10, 220), f"{label_prefix} {idx}", fill="black")
        thumbs.append(card)
    cols = 2
    rows = (len(thumbs) + cols - 1) // cols
    canvas = Image.new("RGB", (cols * 380, rows * 250), "#d9d9d9")
    for idx, thumb in enumerate(thumbs):
        canvas.paste(thumb, ((idx % cols) * 380, (idx // cols) * 250))
    canvas.save(out_path, quality=90)


def select_report_slides(analysis: dict[str, Any], max_report_slides: int) -> list[int]:
    slides = analysis["slides"]
    prioritized = sorted(
        slides,
        key=lambda item: (
            -len(item["risks"]),
            item["title_detection"]["confidence"],
            -item["slide_index"],
        ),
    )
    selected = []
    if slides:
        selected.append(1)
    for slide in prioritized:
        if slide["slide_index"] not in selected:
            selected.append(slide["slide_index"])
        if len(selected) >= max_report_slides:
            break
    return sorted(selected)


def write_results_json(output_dir: Path, payload: dict[str, Any]) -> Path:
    path = output_dir / "results.json"
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def write_report(
    output_dir: Path,
    source_path: Path,
    template_path: Path,
    analysis: dict[str, Any],
    selected_slides: list[int],
    compare_montage: Path | None,
    title_montage: Path | None,
    template_palette,
    export_notes: list[str] | None = None,
) -> Path:
    report_path = output_dir / "报告.md"
    lines = [
        "# PPT 模板替换原型报告",
        "",
        f"- 模板：[{template_path.name}]({template_path.as_posix()})",
        f"- 源文件：[{source_path.name}]({source_path.as_posix()})",
        f"- 幻灯片总数：`{analysis['slide_count']}`",
        f"- 风险页：`{', '.join(str(i) for i in analysis['summary']['high_risk_slides']) or '无'}`",
        f"- 缺失字体：`{', '.join(analysis['missing_fonts']) or '无'}`",
        f"- 模板背景色：`{template_palette.background_dark}`",
        f"- 模板标题色：`{template_palette.title_cyan}`",
        f"- 模板强调色：`{template_palette.accent_orange}`",
        "",
        "## 识别与对比",
        "",
    ]
    if title_montage and title_montage.exists():
        lines.extend([f"![title montage]({title_montage.as_posix()})", ""])
    if compare_montage and compare_montage.exists():
        lines.extend([f"![compare montage]({compare_montage.as_posix()})", ""])
    if export_notes:
        lines.append("## 导出说明")
        lines.append("")
        for note in export_notes:
            lines.append(f"- {note}")
        lines.append("")
    lines.append("## 风险统计")
    lines.append("")
    for risk, count in sorted(analysis["summary"]["risk_counts"].items(), key=lambda item: (-item[1], item[0])):
        lines.append(f"- `{risk}`: {count}")
    lines.append("")
    lines.append("## 抽样页结论")
    lines.append("")
    slide_map = {slide["slide_index"]: slide for slide in analysis["slides"]}
    for slide_index in selected_slides:
        slide = slide_map[slide_index]
        lines.append(
            f"- 第 {slide_index} 页：标题 `{slide['title_detection']['title']}`，置信度 `{slide['title_detection']['confidence']}`，角色 `{slide['role']}`，风险 `{', '.join(slide['risks']) or '无'}`。"
        )
    lines.append("")
    lines.append("## 说明")
    lines.append("")
    lines.append("- 第一版原型仍是单模板风格配置，不是任意模板通用引擎。")
    lines.append("- 当前编辑阶段使用对象级规则替换，截图渲染使用 PowerPoint 原生导出。")
    lines.append("- 图表、表格、复杂组合对象目前只做风险标记，还没有做深度重排。")
    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path
