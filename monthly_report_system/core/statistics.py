"""月报统计逻辑。"""

from __future__ import annotations

from collections import Counter, defaultdict
from typing import Any, Dict, List

from config.indicators import INDICATORS


def summarize_records(records: List[Dict[str, Any]], start_date: str, end_date: str) -> Dict[str, Any]:
    """计算总体统计、街道统计和附件表数据。"""

    streets = sorted({record.get("street", "") for record in records if record.get("street")})
    problem_records = [record for record in records if record.get("has_problem")]
    indicator_by_street: Dict[str, Counter] = defaultdict(Counter)
    uncategorized_by_street: Counter = Counter()

    for record in records:
        street = record.get("street", "")
        indicators = record.get("indicators", [])
        if indicators:
            for indicator in indicators:
                indicator_by_street[street][indicator] += 1
        elif record.get("has_problem"):
            uncategorized_by_street[street] += 1

    total_by_indicator = Counter()
    for counter in indicator_by_street.values():
        total_by_indicator.update(counter)

    table_rows = []
    total_problem_count = sum(total_by_indicator.values()) + sum(uncategorized_by_street.values())
    total_row = {"街道名称": "总计", "问题总数": total_problem_count}
    for indicator in INDICATORS:
        total_row[indicator] = total_by_indicator.get(indicator, 0)
    table_rows.append(total_row)

    for street in streets:
        row_total = sum(indicator_by_street[street].values()) + uncategorized_by_street[street]
        row = {"街道名称": street, "问题总数": row_total}
        for indicator in INDICATORS:
            row[indicator] = indicator_by_street[street].get(indicator, 0)
        table_rows.append(row)

    return {
        "street_count": len(streets),
        "record_count": len(records),
        "problem_record_count": len(problem_records),
        "problem_count": total_problem_count,
        "streets": streets,
        "total_by_indicator": dict(total_by_indicator),
        "indicator_by_street": {street: dict(counter) for street, counter in indicator_by_street.items()},
        "uncategorized_by_street": dict(uncategorized_by_street),
        "table_rows": table_rows,
        "summary_text": build_summary_text(
            start_date=start_date,
            end_date=end_date,
            street_count=len(streets),
            record_count=len(records),
            problem_count=total_problem_count,
            indicator_by_street=indicator_by_street,
            uncategorized_by_street=uncategorized_by_street,
        ),
    }


def _format_street_counts(counter: Counter) -> str:
    parts = [f"{street}{count}个" for street, count in counter.items() if count]
    return "、".join(parts)


def build_summary_text(
    start_date: str,
    end_date: str,
    street_count: int,
    record_count: int,
    problem_count: int,
    indicator_by_street: Dict[str, Counter],
    uncategorized_by_street: Counter,
) -> str:
    """生成总体情况段落。"""

    segments = [
        f"{start_date}至{end_date}对本区{street_count}个街道可回收物交投点进行{record_count}个次检查。"
        f"检查共计发现问题{problem_count}个"
    ]
    indicator_totals: Dict[str, Counter] = defaultdict(Counter)
    for street, counter in indicator_by_street.items():
        for indicator, count in counter.items():
            indicator_totals[indicator][street] += count

    detail_parts = []
    for indicator, street_counter in indicator_totals.items():
        detail = _format_street_counts(street_counter)
        if detail:
            detail_parts.append(f"{indicator}（{detail}）")

    if uncategorized_by_street:
        detail = _format_street_counts(uncategorized_by_street)
        if detail:
            detail_parts.append(f"其他问题（{detail}）")

    if detail_parts:
        return segments[0] + "，" + "；".join(detail_parts) + "。"
    return segments[0] + "。"


def problem_display(record: Dict[str, Any]) -> str:
    """生成案例中展示的问题文本。"""

    indicators = record.get("indicators") or []
    if indicators:
        return "、".join(indicators)
    problem = str(record.get("problem") or "").strip()
    return problem or "无问题"
