"""月报生成业务流水线。"""

from __future__ import annotations

import re
import shutil
from datetime import date, datetime, time
from pathlib import Path
from typing import Iterable

from core.data_cleaner import attach_images_to_records, clean_dataframe
from core.excel_reader import read_excel_table
from core.image_extractor import extract_images_by_row
from core.kitchen_waste_report_generator import generate_kitchen_waste_report, summarize_kitchen_waste_records
from core.models import ProcessingResult, ReportParams
from core.parking_station_report_generator import generate_parking_station_report, summarize_parking_records
from core.report_generator import generate_report
from core.sorting_station_report_generator import generate_sorting_station_report, summarize_sorting_station_records
from core.statistics import summarize_records
from utils.file_utils import cleanup_old_run_dirs, cleanup_run_dir, create_run_dir, safe_filename, safe_upload_filename
from utils.logger import AppLogger


def _copy_input_file(source: Path, target_dir: Path, fallback_name: str) -> Path:
    """把用户选择的文件复制到本次运行目录。"""

    filename = safe_upload_filename(source.name, fallback_name)
    target = target_dir / filename
    shutil.copy2(source, target)
    return target


def _validate_params(params: ReportParams) -> None:
    """校验用户输入。"""

    if not params.excel_path.exists():
        raise ValueError("请选择有效的 Excel 检查数据文件。")
    if params.excel_path.suffix.lower() not in {".xls", ".xlsx"}:
        raise ValueError("Excel 文件必须是 .xls 或 .xlsx 格式。")
    if not params.template_path.exists():
        raise ValueError("请选择有效的 Word 模板文件。")
    if params.template_path.suffix.lower() != ".docx":
        raise ValueError("Word 模板文件必须是 .docx 格式。")


def _ensure_output_does_not_overwrite_inputs(output_path: Path, *input_paths: Path) -> None:
    """避免输出文件覆盖用户选择的原始台账或模板。"""

    resolved_output = output_path.resolve()
    for input_path in input_paths:
        try:
            resolved_input = input_path.resolve()
        except OSError:
            continue
        if resolved_output == resolved_input:
            raise ValueError("输出文件不能与上传的台账或模板文件相同，请更换输出文件名或输出目录。")


def build_preview(records: list[dict], limit: int = 20) -> list[dict]:
    """生成界面预览数据。"""

    return [
        {
            "行号": record.get("row_number"),
            "街道": record.get("street"),
            "点位": record.get("location"),
            "问题描述": record.get("specific_problem") or record.get("problem"),
            "识别指标": "、".join(record.get("indicators", [])) or "无",
            "图片数": len(record.get("images", [])),
        }
        for record in records[:limit]
    ]


def _parse_date_bound(value: str, *, is_end: bool = False) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%Y年%m月%d日", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            parsed = datetime.strptime(text, fmt).date()
            return datetime.combine(parsed, time.max if is_end else time.min)
        except ValueError:
            continue
    digits = [int(part) for part in re.findall(r"\d+", text)]
    if len(digits) >= 3:
        try:
            parsed = date(digits[0], digits[1], digits[2])
            return datetime.combine(parsed, time.max if is_end else time.min)
        except ValueError:
            return None
    return None


def _parse_record_datetime(record: dict) -> datetime | None:
    raw_values: Iterable[object] = (
        record.get("check_date"),
        record.get("raw", {}).get("创建时间"),
        record.get("raw", {}).get("案件上报时间"),
        record.get("raw", {}).get("上报时间"),
        record.get("raw", {}).get("检查时间"),
        record.get("raw", {}).get("检查日期"),
    )
    for value in raw_values:
        if isinstance(value, datetime):
            return value
        if isinstance(value, date):
            return datetime.combine(value, time.min)
        text = str(value or "").strip()
        if not text or text.lower() == "nan":
            continue
        text = text.split(".")[0]
        for fmt in (
            "%Y-%m-%d %H:%M:%S",
            "%Y/%m/%d %H:%M:%S",
            "%Y年%m月%d日 %H:%M:%S",
            "%Y-%m-%d",
            "%Y/%m/%d",
            "%Y年%m月%d日",
        ):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                continue
        digits = [int(part) for part in re.findall(r"\d+", text)]
        if len(digits) >= 3:
            try:
                parts = [digits[0], digits[1], digits[2], 0, 0, 0]
                for index, digit in enumerate(digits[3:6], start=3):
                    parts[index] = digit
                return datetime(*parts)
            except ValueError:
                continue
    return None


def _filter_records_by_date_range(records: list[dict], start_date: str, end_date: str) -> tuple[list[dict], list[int]]:
    start = _parse_date_bound(start_date)
    end = _parse_date_bound(end_date, is_end=True)
    if start is None or end is None:
        return records, []

    filtered: list[dict] = []
    ignored_rows: list[int] = []
    for record in records:
        record_dt = _parse_record_datetime(record)
        if record_dt is not None and not (start <= record_dt <= end):
            ignored_rows.append(record.get("row_number"))
            continue
        filtered.append(record)
    return filtered, [row for row in ignored_rows if row is not None]


def process_report(params: ReportParams, keep_temp: bool = False) -> ProcessingResult:
    """执行完整月报生成流程。"""

    _validate_params(params)
    logger = AppLogger()
    logger.extend(cleanup_old_run_dirs())

    run_dir = create_run_dir()
    try:
        excel_path = _copy_input_file(params.excel_path, run_dir, "data.xlsx")
        template_path = _copy_input_file(params.template_path, run_dir, "template.docx")

        df, actual_excel_path = read_excel_table(excel_path, run_dir)
        logger.extend(df.attrs.get("warnings", []))

        if params.report_type == "kitchen_waste":
            image_limit = 50
        elif params.report_type == "parking_station":
            image_limit = 15
        elif params.report_type == "sorting_station":
            image_limit = 36
        else:
            image_limit = 3
        clean = clean_dataframe(df, params.report_type)
        logger.extend(clean.warnings)
        records_for_report = clean.records
        date_filtered_rows: list[int] = []
        if params.report_type == "transfer_station":
            records_for_report, date_filtered_rows = _filter_records_by_date_range(
                clean.records,
                params.start_date,
                params.end_date,
            )
            if date_filtered_rows:
                logger.info(f"已按检查起止日期过滤掉 {len(date_filtered_rows)} 条交投点记录。")

        images_by_row, image_warnings = extract_images_by_row(actual_excel_path, run_dir / "images", max_images_per_row=image_limit)
        logger.extend(image_warnings)

        variable_image_report = params.report_type in {"parking_station", "kitchen_waste", "sorting_station"}
        records, attach_warnings = attach_images_to_records(
            records_for_report,
            images_by_row,
            max_images_per_record=image_limit,
            warn_on_missing=not variable_image_report,
            warn_on_partial=not variable_image_report,
            ignored_row_numbers=[*clean.filtered_row_numbers, *date_filtered_rows],
        )
        logger.extend(attach_warnings)

        if params.report_type == "parking_station":
            stats = summarize_parking_records(records, params.start_date, params.end_date, params.report_month)
        elif params.report_type == "kitchen_waste":
            stats = summarize_kitchen_waste_records(records, params.report_month)
        elif params.report_type == "sorting_station":
            stats = summarize_sorting_station_records(records, params.start_date, params.end_date)
        else:
            stats = summarize_records(records, params.start_date, params.end_date, params.report_type)
        output_path = params.output_dir / safe_filename(params.output_name)
        _ensure_output_does_not_overwrite_inputs(output_path, params.excel_path, params.template_path, excel_path, template_path)
        if params.report_type == "sorting_station":
            generate_sorting_station_report(
                records=records,
                stats=stats,
                output_path=output_path,
                title=params.title,
                start_date=params.start_date,
                end_date=params.end_date,
                report_month=params.report_month,
                template_path=template_path,
                temp_dir=run_dir,
            )
        elif params.report_type == "parking_station":
            generate_parking_station_report(
                records=records,
                output_path=output_path,
                title=params.title,
                start_date=params.start_date,
                end_date=params.end_date,
                report_month=params.report_month,
                template_path=template_path,
            )
        elif params.report_type == "kitchen_waste":
            generate_kitchen_waste_report(
                records=records,
                output_path=output_path,
                title=params.title,
                report_month=params.report_month,
                template_path=template_path,
            )
        else:
            generate_report(
                records=records,
                stats=stats,
                output_path=output_path,
                title=params.title,
                start_date=params.start_date,
                end_date=params.end_date,
                template_path=template_path,
            )
        for record in records:
            problem_text = str(record.get("specific_problem") or record.get("problem") or "").strip()
            if record.get("has_problem") and not problem_text:
                logger.warning(
                    f"第 {record.get('row_number')} 行点位“{record.get('report_point') or record.get('location', '')}”缺少问题内容，已在 Word 中标红。"
                )

        logger.info(f"Word 月报已生成：{output_path}")
        cleanup_warning = None if keep_temp else cleanup_run_dir(run_dir)
        if keep_temp:
            logger.info(f"本次临时目录：{run_dir}")
        elif cleanup_warning:
            logger.warning(cleanup_warning)
        else:
            logger.info("本次临时文件已自动清理。")

        return ProcessingResult(
            record_count=stats.get("record_count", 0),
            street_count=stats.get("street_count", 0),
            image_count=sum(len(record.get("images", [])) for record in records),
            problem_record_count=stats.get("problem_record_count", 0),
            field_mapping=clean.field_mapping,
            missing_fields=clean.missing_fields,
            summary_text=stats.get("summary_text", ""),
            preview=build_preview(records),
            logs=logger.messages,
            output_path=output_path,
            report_month=params.report_month,
            temp_dir=run_dir if keep_temp else None,
        )
    except Exception:
        if not keep_temp:
            cleanup_run_dir(run_dir)
        raise
