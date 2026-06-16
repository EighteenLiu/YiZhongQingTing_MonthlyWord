"""项目内共享的数据模型。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List


@dataclass(frozen=True)
class ReportParams:
    """一次月报生成所需的用户参数。"""

    excel_path: Path
    template_path: Path
    title: str
    start_date: str
    end_date: str
    report_month: str
    output_dir: Path
    output_name: str
    report_type: str = "transfer_station"


@dataclass
class CheckRecord:
    """清洗后的单条检查记录。"""

    row_number: int
    street: str
    location: str
    report_group: str
    report_point: str
    problem: str
    specific_problem: str
    images: List[str] = field(default_factory=list)
    secondary_location: str = ""
    location_type: str = ""
    result: str = ""
    check_date: str = ""
    indicators: List[str] = field(default_factory=list)
    has_problem: bool = False
    raw: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        """兼容现有统计和 Word 生成函数的 dict 结构。"""

        return {
            "row_number": self.row_number,
            "street": self.street,
            "location": self.location,
            "secondary_location": self.secondary_location,
            "report_group": self.report_group,
            "report_point": self.report_point,
            "problem": self.problem,
            "specific_problem": self.specific_problem,
            "images": self.images,
            "location_type": self.location_type,
            "result": self.result,
            "check_date": self.check_date,
            "indicators": self.indicators,
            "has_problem": self.has_problem,
            "raw": self.raw,
        }


@dataclass
class ProcessingResult:
    """生成完成后返回给界面的结果。"""

    record_count: int
    street_count: int
    image_count: int
    problem_record_count: int
    field_mapping: Dict[str, str]
    missing_fields: List[str]
    summary_text: str
    preview: List[Dict[str, Any]]
    logs: List[str]
    output_path: Path
    report_month: str
    temp_dir: Path | None = None

    def to_dict(self) -> Dict[str, Any]:
        """兼容 Tkinter 展示层目前使用的 dict 访问。"""

        return {
            "record_count": self.record_count,
            "street_count": self.street_count,
            "image_count": self.image_count,
            "problem_record_count": self.problem_record_count,
            "field_mapping": self.field_mapping,
            "missing_fields": self.missing_fields,
            "summary_text": self.summary_text,
            "preview": self.preview,
            "logs": self.logs,
            "output_path": str(self.output_path),
            "report_month": self.report_month,
            "temp_dir": str(self.temp_dir) if self.temp_dir else "",
        }
