"""Excel 数据清洗和字段标准化。"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, Dict, Iterable, List, Tuple

import pandas as pd

from config.indicators import (
    FIELD_ALIASES,
    NON_PROBLEM_DISPLAY_MAP,
    NO_PROBLEM_WORDS,
    PROBLEM_INDICATOR_MAP,
    indicator_keywords_for,
    secondary_locations_for,
)


REQUIRED_FIELDS = ["street", "location", "problem"]
TRANSFER_STATION_GOOD_INDICATOR = "良好，未发现问题"
TRANSFER_STATION_DISPLAY_STATUSES = {
    "未见点位或点位已撤": ["未见点位或点位已撤"],
    "停止营业": ["停止营业"],
    "未营业": ["未营业", "无人营业"],
    "未发现点位": ["未发现点位"],
}


@dataclass
class CleanResult:
    records: List[Dict[str, Any]]
    field_mapping: Dict[str, str]
    missing_fields: List[str]
    warnings: List[str] = field(default_factory=list)
    filtered_row_numbers: List[int] = field(default_factory=list)


def normalize_header(value: Any) -> str:
    """标准化表头用于匹配。"""

    return str(value or "").strip().replace(" ", "").replace("\n", "").lower()


def guess_field_mapping(columns: Iterable[str]) -> Dict[str, str]:
    """根据字段别名推断 Excel 列到标准字段的映射。"""

    normalized = {normalize_header(col): col for col in columns}
    mapping: Dict[str, str] = {}
    for standard, aliases in FIELD_ALIASES.items():
        candidates = [normalize_header(a) for a in aliases]
        for key, original in normalized.items():
            if key in candidates or any(alias in key for alias in candidates):
                mapping[standard] = original
                break
    return mapping


def cell_to_text(value: Any) -> str:
    """把单元格值转为前端和 Word 友好的文本。"""

    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (datetime, date)):
        return f"{value.year}年{value.month}月{value.day}日"
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def normalize_problem_text(value: Any) -> str:
    """标准化具体问题原文，用于命中固化的问题-指标映射。"""

    text = cell_to_text(value)
    replacements = {
        " ": "",
        "\n": "",
        "\r": "",
        "\t": "",
        "，": ",",
        "；": ";",
        "：": ":",
        "（": "(",
        "）": ")",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    return text


def problem_display_text(value: Any) -> str:
    """把原始具体问题转为 Word 中应展示在点位后的文字。"""

    text = cell_to_text(value)
    normalized = normalize_problem_text(text)
    if normalized in NON_PROBLEM_DISPLAY_MAP:
        return NON_PROBLEM_DISPLAY_MAP[normalized]
    return text.strip()


def transfer_station_status_display_text(value: Any, secondary_indicator: Any = "") -> str:
    """交投点良好记录中仍需展示的特殊状态文本。"""

    if cell_to_text(secondary_indicator) != TRANSFER_STATION_GOOD_INDICATOR:
        return ""
    text = cell_to_text(value)
    if not text:
        return ""
    compact = normalize_problem_text(text)
    for display_text, keywords in TRANSFER_STATION_DISPLAY_STATUSES.items():
        if any(normalize_problem_text(keyword) in compact for keyword in keywords):
            return display_text
    return ""


def _transfer_station_indicator_source(specific_problem: str, problem: str, secondary_indicator: str) -> str:
    """交投点问题类型以 2级指标为准，具体问题只处理占位值。"""

    if secondary_indicator:
        return secondary_indicator
    return specific_problem or problem


def is_fixed_non_problem_status(value: Any) -> bool:
    """判断是否为需要特殊展示但不计入问题的状态。"""

    return normalize_problem_text(value) in NON_PROBLEM_DISPLAY_MAP


def is_indicator_placeholder(value: Any) -> bool:
    """判断具体问题是否只是指向指标列的占位值。"""

    return normalize_problem_text(value) == "1"


def detect_indicators(problem_text: str, row: Dict[str, Any] | None = None, report_type: str = "transfer_station") -> List[str]:
    """从问题描述和可能存在的指标列中识别问题类型。"""

    row = row or {}
    normalized = normalize_problem_text(problem_text)
    if normalized in NON_PROBLEM_DISPLAY_MAP:
        return []

    compact = cell_to_text(problem_text).replace(" ", "")
    found: List[str] = []
    mapped = PROBLEM_INDICATOR_MAP.get(normalized, [])
    found.extend(mapped)

    for indicator, groups in indicator_keywords_for(report_type).items():
        explicit_value = cell_to_text(row.get(indicator))
        if explicit_value and explicit_value not in {"0", "否", "无", "正常"} and indicator not in found:
            found.append(indicator)
            continue
        for group in groups:
            if indicator not in found and all(keyword in compact for keyword in group):
                found.append(indicator)
                break
    return found


def is_no_problem(problem_text: str, result_text: str = "") -> bool:
    """判断一条记录是否无问题。"""

    problem = cell_to_text(problem_text).replace(" ", "")
    result = cell_to_text(result_text).replace(" ", "")
    if not problem and not result:
        return True
    return normalize_problem_text(problem_text) in NON_PROBLEM_DISPLAY_MAP or problem in NO_PROBLEM_WORDS or result in NO_PROBLEM_WORDS


def _should_preserve_problem_text(report_type: str, value: Any) -> bool:
    """停车场报告正文按“具体问题”原文输出，不再依赖旧指标词库判定。"""

    if report_type != "parking_station":
        return False
    return not is_no_problem(cell_to_text(value))


def _strip_point_suffix(value: str) -> str:
    """4级点位只取点位本名，后缀在报告中固定补充。"""

    text = value.strip()
    suffixes = ["可回收物交投点", "可回收物路侧交投点", "交投点"]
    for suffix in suffixes:
        if text.endswith(suffix):
            return text[: -len(suffix)].strip("：: ")
    return text


def _normalize_filter_text(value: Any) -> str:
    """标准化用于二级点位匹配的文本。"""

    return cell_to_text(value).replace(" ", "").replace("\n", "").replace("\r", "")


def _matches_secondary_location(value: Any, allowed_locations: List[str]) -> bool:
    """判断当前 2级点位是否属于当前报告类型。"""

    text = _normalize_filter_text(value)
    if not text:
        return True
    for allowed in allowed_locations:
        allowed_text = _normalize_filter_text(allowed)
        if text == allowed_text or allowed_text in text:
            return True
    return False


def clean_dataframe(df: pd.DataFrame, report_type: str = "transfer_station") -> CleanResult:
    """清洗原始 DataFrame 并返回标准记录。"""

    df = df.dropna(how="all").copy()
    mapping = guess_field_mapping(df.columns)
    missing = []
    for field in REQUIRED_FIELDS:
        if field in mapping:
            continue
        if field == "street" and "report_group" in mapping:
            continue
        if field == "location" and "report_point" in mapping:
            continue
        missing.append(field)
    warnings: List[str] = []
    if missing:
        warnings.append(f"缺失必要字段：{', '.join(missing)}")

    records: List[Dict[str, Any]] = []
    filtered_row_numbers: List[int] = []
    allowed_secondary_locations = secondary_locations_for(report_type)
    excel_start_row = int(df.attrs.get("excel_start_row", 2))
    for position, (_, row) in enumerate(df.iterrows()):
        row_number = position + excel_start_row
        raw = row.to_dict()
        secondary_location = cell_to_text(raw.get(mapping.get("secondary_location", ""), ""))
        if allowed_secondary_locations and not _matches_secondary_location(secondary_location, allowed_secondary_locations):
            filtered_row_numbers.append(row_number)
            continue

        problem = cell_to_text(raw.get(mapping.get("problem", ""), ""))
        specific_problem = cell_to_text(raw.get(mapping.get("specific_problem", ""), "")) or problem
        secondary_indicator = cell_to_text(raw.get(mapping.get("secondary_indicator", ""), ""))
        if report_type == "transfer_station":
            indicator_source = _transfer_station_indicator_source(specific_problem, problem, secondary_indicator)
        else:
            indicator_source = secondary_indicator if is_indicator_placeholder(specific_problem or problem) else specific_problem or problem
        status_display_problem = (
            transfer_station_status_display_text(specific_problem or problem, secondary_indicator)
            if report_type == "transfer_station"
            else ""
        )
        display_problem = status_display_problem or problem_display_text(indicator_source)
        raw_street = cell_to_text(raw.get(mapping.get("street", ""), ""))
        raw_location = cell_to_text(raw.get(mapping.get("location", ""), ""))
        report_group = cell_to_text(raw.get(mapping.get("report_group", ""), "")) or raw_street
        report_point = cell_to_text(raw.get(mapping.get("report_point", ""), "")) or raw_location
        street = report_group or raw_street
        location = report_point or raw_location

        record = {
            "row_number": row_number,
            "street": street,
            "location": location,
            "secondary_location": secondary_location,
            "location_type": cell_to_text(raw.get(mapping.get("location_type", ""), "")),
            "result": cell_to_text(raw.get(mapping.get("result", ""), "")),
            "problem": indicator_source,
            "specific_problem": display_problem,
            "secondary_indicator": secondary_indicator,
            "report_group": report_group,
            "report_point": _strip_point_suffix(report_point),
            "check_date": cell_to_text(raw.get(mapping.get("check_date", ""), "")),
            "images": [],
            "raw": raw,
        }
        if not record["street"] and not record["location"] and not record["problem"] and not report_group and not report_point:
            continue
        indicators = detect_indicators(indicator_source, raw, report_type)
        record["indicators"] = indicators
        record["has_problem"] = bool(indicators)
        if not indicators and status_display_problem:
            record["specific_problem"] = status_display_problem
            record["problem"] = status_display_problem
        elif not indicators and not is_fixed_non_problem_status(indicator_source) and not _should_preserve_problem_text(report_type, indicator_source):
            record["specific_problem"] = "无问题"
            record["problem"] = "无问题"
        if not record["specific_problem"] and not record["has_problem"]:
            record["specific_problem"] = "无问题"
        if not record["problem"] and not record["has_problem"]:
            record["problem"] = "无问题"
        records.append(record)

    if filtered_row_numbers:
        warnings.append(f"已按{report_type}月报过滤掉 {len(filtered_row_numbers)} 条非对应2级点位记录。")

    return CleanResult(
        records=records,
        field_mapping=mapping,
        missing_fields=missing,
        warnings=warnings,
        filtered_row_numbers=filtered_row_numbers,
    )


def attach_images_to_records(
    records: List[Dict[str, Any]],
    images_by_row: Dict[int, List[str]],
    max_images_per_record: int = 3,
    warn_on_missing: bool = True,
    warn_on_partial: bool = True,
    ignored_row_numbers: Iterable[int] | None = None,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """把图片按 Excel 行号关联到记录。"""

    warnings: List[str] = []
    ignored_rows = set(ignored_row_numbers or [])
    row_to_record = {record["row_number"]: record for record in records}
    for row_number, paths in images_by_row.items():
        if row_number in ignored_rows:
            continue
        record = row_to_record.get(row_number)
        if record is None:
            warnings.append(f"图片所在行 {row_number} 未匹配到数据记录。")
            continue
        record.setdefault("images", []).extend(paths[:max_images_per_record])

    for record in records:
        record["images"] = record.get("images", [])[:max_images_per_record]
        image_count = len(record["images"])
        if image_count == 0 and warn_on_missing:
            warnings.append(f"第 {record['row_number']} 行未匹配到图片：{record.get('report_point') or record.get('location', '')}")
        elif image_count != max_images_per_record and warn_on_partial:
            warnings.append(f"第 {record['row_number']} 行匹配到 {image_count} 张图片，期望 {max_images_per_record} 张：{record.get('report_point') or record.get('location', '')}")
    return records, warnings
