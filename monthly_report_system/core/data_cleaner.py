"""Excel 数据清洗和字段标准化。"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, Dict, Iterable, List, Tuple

import pandas as pd

from config.indicators import FIELD_ALIASES, INDICATOR_KEYWORDS, NO_PROBLEM_WORDS


REQUIRED_FIELDS = ["street", "location", "problem"]


@dataclass
class CleanResult:
    records: List[Dict[str, Any]]
    field_mapping: Dict[str, str]
    missing_fields: List[str]
    warnings: List[str] = field(default_factory=list)


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


def detect_indicators(problem_text: str, row: Dict[str, Any] | None = None) -> List[str]:
    """从问题描述和可能存在的指标列中识别问题类型。"""

    row = row or {}
    compact = cell_to_text(problem_text).replace(" ", "")
    found: List[str] = []

    for indicator, groups in INDICATOR_KEYWORDS.items():
        explicit_value = cell_to_text(row.get(indicator))
        if explicit_value and explicit_value not in {"0", "否", "无", "正常"}:
            found.append(indicator)
            continue
        for group in groups:
            if all(keyword in compact for keyword in group):
                found.append(indicator)
                break
    return found


def is_no_problem(problem_text: str, result_text: str = "") -> bool:
    """判断一条记录是否无问题。"""

    problem = cell_to_text(problem_text).replace(" ", "")
    result = cell_to_text(result_text).replace(" ", "")
    if not problem and not result:
        return True
    return problem in NO_PROBLEM_WORDS or result in NO_PROBLEM_WORDS


def _strip_point_suffix(value: str) -> str:
    """4级点位只取点位本名，后缀在报告中固定补充。"""

    text = value.strip()
    suffixes = ["可回收物交投点", "可回收物路侧交投点", "交投点"]
    for suffix in suffixes:
        if text.endswith(suffix):
            return text[: -len(suffix)].strip("：: ")
    return text


def clean_dataframe(df: pd.DataFrame) -> CleanResult:
    """清洗原始 DataFrame 并返回标准记录。"""

    df = df.dropna(how="all").copy()
    mapping = guess_field_mapping(df.columns)
    missing = [field for field in REQUIRED_FIELDS if field not in mapping]
    warnings: List[str] = []
    if missing:
        warnings.append(f"缺失必要字段：{', '.join(missing)}")

    records: List[Dict[str, Any]] = []
    excel_start_row = int(df.attrs.get("excel_start_row", 2))
    for position, (_, row) in enumerate(df.iterrows()):
        raw = row.to_dict()
        problem = cell_to_text(raw.get(mapping.get("problem", ""), ""))
        specific_problem = cell_to_text(raw.get(mapping.get("specific_problem", ""), "")) or problem
        raw_street = cell_to_text(raw.get(mapping.get("street", ""), ""))
        raw_location = cell_to_text(raw.get(mapping.get("location", ""), ""))
        report_group = cell_to_text(raw.get(mapping.get("report_group", ""), "")) or raw_street
        report_point = cell_to_text(raw.get(mapping.get("report_point", ""), "")) or raw_location
        street = report_group or raw_street
        location = report_point or raw_location

        record = {
            "row_number": position + excel_start_row,
            "street": street,
            "location": location,
            "location_type": cell_to_text(raw.get(mapping.get("location_type", ""), "")),
            "result": cell_to_text(raw.get(mapping.get("result", ""), "")),
            "problem": problem,
            "specific_problem": specific_problem,
            "report_group": report_group,
            "report_point": _strip_point_suffix(report_point),
            "check_date": cell_to_text(raw.get(mapping.get("check_date", ""), "")),
            "images": [],
            "raw": raw,
        }
        if not record["street"] and not record["location"] and not record["problem"] and not report_group and not report_point:
            continue
        indicators = detect_indicators(specific_problem or problem, raw)
        record["indicators"] = indicators
        record["has_problem"] = bool(indicators) or not is_no_problem(specific_problem or problem, record["result"])
        if not record["specific_problem"] and not record["has_problem"]:
            record["specific_problem"] = "无问题"
        if not record["problem"] and not record["has_problem"]:
            record["problem"] = "无问题"
        records.append(record)

    return CleanResult(records=records, field_mapping=mapping, missing_fields=missing, warnings=warnings)


def attach_images_to_records(records: List[Dict[str, Any]], images_by_row: Dict[int, List[str]]) -> Tuple[List[Dict[str, Any]], List[str]]:
    """把图片按 Excel 行号关联到记录。"""

    warnings: List[str] = []
    row_to_record = {record["row_number"]: record for record in records}
    for row_number, paths in images_by_row.items():
        record = row_to_record.get(row_number)
        if record is None:
            warnings.append(f"图片所在行 {row_number} 未匹配到数据记录。")
            continue
        record.setdefault("images", []).extend(paths[:3])

    for record in records:
        record["images"] = record.get("images", [])[:3]
        image_count = len(record["images"])
        if image_count == 0:
            warnings.append(f"第 {record['row_number']} 行未匹配到图片：{record.get('report_point') or record.get('location', '')}")
        elif image_count != 3:
            warnings.append(f"第 {record['row_number']} 行匹配到 {image_count} 张图片，期望 3 张：{record.get('report_point') or record.get('location', '')}")
    return records, warnings
