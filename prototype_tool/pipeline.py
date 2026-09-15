from __future__ import annotations

import time
from pathlib import Path
from typing import Any, Callable

from pptx import Presentation

from .analysis import analyze_deck
from .config import PrototypeOptions
from .com_transform import transform_and_export_with_com
from .render import export_presentation_to_png
from .reporting import (
    annotate_title_image,
    build_before_after_board,
    make_montage,
    select_report_slides,
    write_report,
    write_results_json,
)
from .template_profile import write_template_profile_json
from .template_theme import TemplateTheme, write_template_palette_json
from .transform import transform_deck


ProgressCallback = Callable[[str], None] | None


def build_case_outputs(
    options: PrototypeOptions, analysis: dict[str, Any]
) -> dict[str, Path]:
    case_dir = options.output_root / options.source_path.stem
    return {
        "case_dir": case_dir,
        "source_export_dir": case_dir / "source_exports",
        "prototype_pptx": case_dir / f"{options.source_path.stem}_prototype.pptx",
        "prototype_export_dir": case_dir / "prototype_exports",
        "title_detection_dir": case_dir / "title_detection",
        "compare_dir": case_dir / "compare",
        "compare_montage": case_dir / "compare_montage.jpg",
        "title_montage": case_dir / "title_detection_montage.jpg",
        "com_error_log": case_dir / "com_edit_error.txt",
    }


def _emit(callback: ProgressCallback, message: str) -> None:
    if callback:
        callback(message)


def _wait_until_unlocked(path: Path, timeout_s: float = 8.0) -> None:
    if not path.exists():
        return
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            with path.open("ab"):
                return
        except PermissionError:
            time.sleep(0.4)


def _is_transient_com_error(exc: Exception) -> bool:
    message = str(exc)
    markers = (
        "被呼叫方拒绝接收呼叫",
        "call was rejected by callee",
        "-2147418111",
    )
    return any(marker in message for marker in markers)


def process_source(
    options: PrototypeOptions, progress: ProgressCallback = None
) -> dict[str, Any]:
    _emit(progress, f"分析源文件: {options.source_path.name}")
    analysis = analyze_deck(options.source_path, options.template_path)
    outputs = build_case_outputs(options, analysis)
    outputs["case_dir"].mkdir(parents=True, exist_ok=True)
    profile_json = write_template_profile_json(outputs["case_dir"], options.profile)
    theme = TemplateTheme(palette=options.palette, metadata=options.palette_metadata)
    palette_json = write_template_palette_json(outputs["case_dir"], theme)
    export_notes: list[str] = []

    if options.skip_render:
        _emit(progress, "跳过源文件截图（--skip-render）")
    else:
        _emit(progress, "导出源文件截图")
        try:
            export_presentation_to_png(
                options.source_path,
                outputs["source_export_dir"],
                width=options.render_config.width,
                height=options.render_config.height,
                retries=options.render_config.retries,
            )
        except Exception as exc:
            export_notes.append(f"源文件截图导出失败：{exc}")
            _emit(progress, f"源文件截图导出失败，继续处理: {exc}")

    engine_used = options.edit_engine
    if options.edit_engine == "com":
        _emit(progress, "尝试 COM 编辑路径")
        try:
            transform_result = None
            last_com_error = None
            for com_attempt in range(1, 3):
                try:
                    transform_result = transform_and_export_with_com(
                        options.source_path,
                        outputs["prototype_pptx"],
                        outputs["prototype_export_dir"],
                        analysis,
                        options.palette,
                        options.profile,
                        options.template_path,
                        options.render_config.width,
                        options.render_config.height,
                        skip_render=options.skip_render,
                    )
                    break
                except Exception as exc:
                    last_com_error = exc
                    outputs["com_error_log"].write_text(str(exc), encoding="utf-8")
                    if com_attempt < 2 and _is_transient_com_error(exc):
                        _emit(progress, f"COM 编辑第 {com_attempt} 次失败，准备自动重试: {exc}")
                        _wait_until_unlocked(outputs["prototype_pptx"])
                        time.sleep(1.0 * com_attempt)
                        continue
                    raise
            if transform_result is None:
                raise last_com_error or RuntimeError("COM 编辑未生成结果")
        except Exception as exc:
            outputs["com_error_log"].write_text(str(exc), encoding="utf-8")
            engine_used = "pptx-fallback"
            _emit(progress, f"COM 编辑失败，自动回退: {exc}")
            _wait_until_unlocked(outputs["prototype_pptx"])
            transform_result = transform_deck(
                options.source_path,
                outputs["prototype_pptx"],
                analysis,
                options.palette,
                options.profile,
                options.template_path,
            )
            if not options.skip_render:
                try:
                    export_presentation_to_png(
                        outputs["prototype_pptx"],
                        outputs["prototype_export_dir"],
                        width=options.render_config.width,
                        height=options.render_config.height,
                        retries=options.render_config.retries,
                    )
                except Exception as exc:
                    export_notes.append(f"成品截图导出失败：{exc}")
                    _emit(progress, f"成品截图导出失败，继续输出报告: {exc}")
    else:
        _emit(progress, "使用 python-pptx 编辑路径")
        transform_result = transform_deck(
            options.source_path,
            outputs["prototype_pptx"],
            analysis,
            options.palette,
            options.profile,
            options.template_path,
        )
        if not options.skip_render:
            try:
                export_presentation_to_png(
                    outputs["prototype_pptx"],
                    outputs["prototype_export_dir"],
                    width=options.render_config.width,
                    height=options.render_config.height,
                    retries=options.render_config.retries,
                )
            except Exception as exc:
                export_notes.append(f"成品截图导出失败：{exc}")
                _emit(progress, f"成品截图导出失败，继续输出报告: {exc}")

    _emit(progress, "生成标注图与对比图")
    prs = Presentation(options.source_path)
    report_slides = select_report_slides(analysis, options.max_report_slides)
    outputs["title_detection_dir"].mkdir(exist_ok=True)
    outputs["compare_dir"].mkdir(exist_ok=True)
    title_images = []
    compare_images = []
    slide_map = {slide["slide_index"]: slide for slide in analysis["slides"]}
    for slide_index in report_slides:
        slide_analysis = slide_map[slide_index]
        source_image = outputs["source_export_dir"] / f"幻灯片{slide_index}.PNG"
        prototype_image = outputs["prototype_export_dir"] / f"幻灯片{slide_index}.PNG"
        annotated_path = (
            outputs["title_detection_dir"] / f"slide_{slide_index:02d}_title.png"
        )
        compare_path = outputs["compare_dir"] / f"slide_{slide_index:02d}_compare.jpg"
        if source_image.exists():
            annotate_title_image(
                source_image, annotated_path, prs, slide_analysis["title_detection"]
            )
            title_images.append(annotated_path)
        if source_image.exists() and prototype_image.exists():
            build_before_after_board(
                source_image,
                prototype_image,
                compare_path,
                f"{options.source_path.stem} slide {slide_index}",
            )
            compare_images.append(compare_path)
    compare_montage = outputs["compare_montage"] if compare_images else None
    title_montage = outputs["title_montage"] if title_images else None
    if compare_images:
        make_montage(
            compare_images,
            outputs["compare_montage"],
            f"{options.source_path.stem} compare",
        )
    if title_images:
        make_montage(
            title_images, outputs["title_montage"], f"{options.source_path.stem} title"
        )

    _emit(progress, "输出报告")
    report_path = write_report(
        outputs["case_dir"],
        options.source_path,
        options.template_path,
        analysis,
        report_slides,
        compare_montage,
        title_montage,
        options.palette,
        export_notes=export_notes,
    )

    result_payload = {
        "analysis": analysis,
        "template_profile": options.profile.__dict__,
        "template_palette": options.palette.__dict__,
        "transform": transform_result,
        "engine_requested": options.edit_engine,
        "engine_used": engine_used,
        "export_notes": export_notes,
        "report_slides": report_slides,
        "report_path": str(report_path),
        "template_profile_json": str(profile_json),
        "template_palette_json": str(palette_json),
    }
    results_json = write_results_json(outputs["case_dir"], result_payload)
    _emit(progress, f"完成: {options.source_path.name}")
    return {
        "source": str(options.source_path),
        "case_dir": str(outputs["case_dir"]),
        "prototype_pptx": str(outputs["prototype_pptx"]),
        "report_path": str(report_path),
        "results_json": str(results_json),
        "risk_slides": analysis["summary"]["high_risk_slides"],
        "engine_used": engine_used,
        "template_palette_json": str(palette_json),
    }
