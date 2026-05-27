"""月报生成系统的集中配置。"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List


FIELD_ALIASES: Dict[str, List[str]] = {
    "street": ["街道", "街道名称", "所属街道", "所属街乡", "所在街道"],
    "location": ["点位", "点位名称", "地址", "检查点位", "场所名称", "交投点名称"],
    "location_type": ["点位类型", "类型", "设施类型", "交投点类型"],
    "result": ["检查结果", "结果", "是否合格", "检查情况"],
    "problem": ["问题", "问题描述", "检查问题", "存在问题", "问题情况", "情况描述"],
    "specific_problem": ["具体问题", "问题明细", "具体情况", "检查具体问题"],
    "report_group": ["3级点位", "三级点位", "三层点位", "街道案例标题"],
    "report_point": ["4级点位", "四级点位", "四层点位", "案例点位标题"],
    "check_date": ["检查日期", "日期", "巡查日期", "检查时间"],
    "images": ["图片", "现场照片", "整改前照片", "检查照片", "照片"],
}


INDICATORS: List[str] = [
    "无防火标识",
    "无称重系统",
    "无回收价格表",
    "无灭火器",
    "灭火器未成组",
    "灭火器不合格",
    "占道经营",
    "无消防安全水源",
    "周边环境脏乱",
    "无七禁收八不准承诺书",
]


INDICATOR_KEYWORDS: Dict[str, List[List[str]]] = {
    "无防火标识": [["防火标识", "无"], ["防火标识", "缺失"], ["未见", "防火标识"]],
    "无称重系统": [["称重系统", "无"], ["称重系统", "缺失"], ["未见", "称重系统"]],
    "无回收价格表": [["价格表", "无"], ["价格表", "缺失"], ["未见", "价格表"]],
    "无灭火器": [["无灭火器"], ["灭火器", "无"], ["未见", "灭火器"]],
    "灭火器未成组": [["灭火器未成组"], ["灭火器不成组"], ["灭火器", "未成组"], ["灭火器", "不成组"]],
    "灭火器不合格": [["灭火器不合格"], ["灭火器", "不合格"]],
    "占道经营": [["占道经营"], ["占道"]],
    "无消防安全水源": [["消防安全水源"], ["消防水源"], ["未见", "水源"]],
    "周边环境脏乱": [["环境脏乱"], ["周边环境", "脏乱"], ["环境", "脏乱"]],
    "无七禁收八不准承诺书": [["七禁收八不准"], ["承诺书"]],
}


NO_PROBLEM_WORDS = ["无问题", "未发现问题", "正常", "合格", "无"]


@dataclass(frozen=True)
class ReportStyle:
    """Word 输出样式参数。"""

    page_margin_cm: float = 2.5
    title_font: str = "方正小标宋简体"
    body_font: str = "楷体_GB2312"
    fallback_font: str = "楷体"
    title_size_pt: int = 22
    heading_size_pt: int = 16
    body_size_pt: int = 15
    table_size_pt: int = 10
    first_line_indent_chars: int = 2
    image_width_cm: float = 5.1
    image_height_cm: float = 3
    pictures_per_row: int = 3
    table_zero_text: str = "0"


REPORT_STYLE = ReportStyle()
