"""月报统计逻辑。"""

from __future__ import annotations

from collections import Counter, defaultdict
from typing import Any, Dict, List

from config.indicators import indicators_for


def summarize_records(records: List[Dict[str, Any]], start_date: str, end_date: str, report_type: str = "transfer_station") -> Dict[str, Any]:
    """计算总体统计、街道统计和附件表数据。"""

    streets = sorted({record.get("street", "") for record in records if record.get("street")})
    indicators = indicators_for(report_type)
    problem_records = [record for record in records if _record_indicators(record, report_type, indicators)]
    indicator_by_street: Dict[str, Counter] = defaultdict(Counter)

    for record in records:
        street = record.get("street", "")
        record_indicators = _record_indicators(record, report_type, indicators)
        if record_indicators:
            for indicator in record_indicators:
                indicator_by_street[street][indicator] += 1

    total_by_indicator = Counter()
    for counter in indicator_by_street.values():
        total_by_indicator.update(counter)

    total_problem_count = sum(total_by_indicator.values())
    table_rows = [_build_table_row("总计", total_by_indicator, total_problem_count, indicators)]
    for street in streets:
        row_total = sum(indicator_by_street[street].values())
        table_rows.append(_build_table_row(street, indicator_by_street[street], row_total, indicators))

    return {
        "street_count": len(streets),
        "record_count": len(records),
        "problem_record_count": len(problem_records),
        "problem_count": total_problem_count,
        "streets": streets,
        "total_by_indicator": dict(total_by_indicator),
        "indicator_by_street": {street: dict(counter) for street, counter in indicator_by_street.items()},
        "uncategorized_by_street": {},
        "table_rows": table_rows,
        "summary_text": build_summary_text(
            start_date=start_date,
            end_date=end_date,
            street_count=len(streets),
            record_count=len(records),
            problem_count=total_problem_count,
            indicator_by_street=indicator_by_street,
        ),
    }


def _record_indicators(record: Dict[str, Any], report_type: str, allowed_indicators: List[str]) -> List[str]:
    """生成统计和汇总应使用的问题指标。"""

    if report_type == "transfer_station":
        secondary_indicator = str(record.get("secondary_indicator") or "").strip()
        if secondary_indicator in allowed_indicators:
            return [secondary_indicator]

    return [indicator for indicator in record.get("indicators", []) if indicator in allowed_indicators]


def _build_table_row(street_name: str, counter: Counter, total: int, indicators: List[str]) -> Dict[str, Any]:
    """生成附件统计表单行。"""

    row = {"街道名称": street_name, "问题总数": total}
    for indicator in indicators:
        row[indicator] = counter.get(indicator, 0)
    return row


def _format_street_counts(counter: Counter) -> str:
    return "、".join(f"{street}{count}个" for street, count in counter.items() if count)


def build_summary_text(
    start_date: str,
    end_date: str,
    street_count: int,
    record_count: int,
    problem_count: int,
    indicator_by_street: Dict[str, Counter],
) -> str:
    """生成总体情况段落。"""

    base = f"{start_date}至{end_date}对本区{street_count}个街道可回收物交投点进行{record_count}个次检查。检查共计发现问题{problem_count}个"
    indicator_totals: Dict[str, Counter] = defaultdict(Counter)
    for street, counter in indicator_by_street.items():
        for indicator, count in counter.items():
            indicator_totals[indicator][street] += count

    details = []
    for indicator, street_counter in indicator_totals.items():
        detail = _format_street_counts(street_counter)
        if detail:
            details.append(f"{indicator}（{detail}）")

    return base + ("，" + "；".join(details) if details else "") + "。"


def problem_display(record: Dict[str, Any]) -> str:
    """生成案例中展示的问题文本。"""

    transfer_indicators = indicators_for("transfer_station")
    secondary_indicator = str(record.get("secondary_indicator") or "").strip()
    if secondary_indicator in transfer_indicators:
        return secondary_indicator

    indicators = [indicator for indicator in record.get("indicators", []) if indicator in transfer_indicators]
    if indicators:
        return "、".join(indicators)
    problem = str(record.get("specific_problem") or record.get("problem") or "").strip()
    return problem or "无问题"
