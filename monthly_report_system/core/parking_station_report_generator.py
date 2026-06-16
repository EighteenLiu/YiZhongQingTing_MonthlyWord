"""环卫停车场 Jinja 模板报告生成。"""

from __future__ import annotations

from collections import OrderedDict
from datetime import date, datetime
import re
from types import SimpleNamespace
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List

from docx.shared import Cm


PARKING_ITEMS = ["作业安全", "设施设备", "消防安全", "场内环境", "场内容器"]
DEFAULT_DEPARTMENT = "北京市西城区城市管理委员会"
DEFAULT_CLOSING_TEXT = (
    "所查点位发现问题，所属单位立即整改，排除安全隐患/加强日常管理/做好培训工作。"
    "所查点位未发现问题，望各单位继续保持，共建美丽西城。"
)
NO_PROBLEM_TEXTS = {
    "",
    "1",
    "无",
    "无。",
    "无问题",
    "无问题。",
    "未发现问题",
    "未发现问题。",
    "正常",
    "合格",
    "良好，未发现问题",
}
PHOTO_ITEM_NAMES = {"照片", "图片", "现场照片", "问题照片", "大门图片"}
CN_NUMBERS = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]


ImageFactory = Callable[[str], Any]


def _cn_index(index: int) -> str:
    if 1 <= index <= len(CN_NUMBERS):
        return CN_NUMBERS[index - 1]
    if index < 20:
        return "十" + CN_NUMBERS[index - 11]
    return str(index)


def _text(value: Any) -> str:
    return str(value or "").strip()


def _normalize_problem_text(value: Any) -> str:
    text = _text(value)
    return re.sub(r"\s+", "", text).strip("。；;，,")


def _is_no_problem(value: Any) -> bool:
    return _normalize_problem_text(value) in {_normalize_problem_text(item) for item in NO_PROBLEM_TEXTS}


def _default_item_text(item_name: str) -> str:
    return "无问题。"


def _site_name(record: Dict[str, Any]) -> str:
    return _text(record.get("report_point")) or _text(record.get("location")) or "未识别停车场"


def _item_name(record: Dict[str, Any]) -> str | None:
    raw_item = _text(record.get("secondary_indicator")) or _text(record.get("problem"))
    if raw_item in PARKING_ITEMS:
        return raw_item
    if raw_item in PHOTO_ITEM_NAMES:
        return None

    source = f"{raw_item} {_text(record.get('specific_problem'))}"
    for item in PARKING_ITEMS:
        if item in source:
            return item
    if any(keyword in source for keyword in ("作业", "安全生产", "培训", "制度")):
        return "作业安全"
    if any(keyword in source for keyword in ("设施", "设备", "称重", "价格", "公示")):
        return "设施设备"
    if any(keyword in source for keyword in ("消防", "灭火", "火源", "水源")):
        return "消防安全"
    if any(keyword in source for keyword in ("环境", "脏乱", "杂乱", "占道")):
        return "场内环境"
    if any(keyword in source for keyword in ("容器", "垃圾桶", "桶")):
        return "场内容器"
    return None


def _issue_text(record: Dict[str, Any]) -> str:
    problem = _text(record.get("specific_problem") or record.get("problem"))
    if _is_no_problem(problem):
        return ""
    return problem


def _with_sentence_period(value: str) -> str:
    text = value.strip()
    if not text:
        return ""
    return text if text.endswith(("。", "！", "？", "；", ";")) else f"{text}。"


def _append_unique(parts: List[str], value: str) -> None:
    text = value.strip()
    if text and text not in parts:
        parts.append(text)


def _photo_rows(image_paths: Iterable[str], image_factory: ImageFactory | None = None) -> List[Dict[str, Any]]:
    images = [image_factory(path) if image_factory else path for path in image_paths]
    rows = []
    for index in range(0, len(images), 2):
        rows.append(
            {
                "left": images[index],
                "right": images[index + 1] if index + 1 < len(images) else "",
            }
        )
    return rows


def _to_attr_object(value: Any) -> Any:
    if isinstance(value, dict):
        return SimpleNamespace(**{key: _to_attr_object(item) for key, item in value.items()})
    if isinstance(value, list):
        return [_to_attr_object(item) for item in value]
    return value


def _issue_date_from_report_month(report_month: str) -> str:
    text = _text(report_month)
    match = re.match(r"(\d{4})年(\d{1,2})月", text)
    if match:
        return f"{match.group(1)}年{int(match.group(2))}月26日"
    try:
        parsed = datetime.strptime(text, "%Y-%m").date()
        return f"{parsed.year}年{parsed.month}月26日"
    except ValueError:
        today = date.today()
        return f"{today.year}年{today.month}月{today.day}日"


def build_parking_station_context(
    records: List[Dict[str, Any]],
    start_date: str,
    end_date: str,
    report_month: str,
    title: str = "西城区环卫作业停车场检查报告",
    image_factory: ImageFactory | None = None,
) -> Dict[str, Any]:
    """把清洗后的停车场记录转换为模板所需上下文。"""

    grouped: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
    for record in records:
        site_name = _site_name(record)
        if site_name not in grouped:
            grouped[site_name] = {
                "name": site_name,
                "items": OrderedDict(),
                "issues": [],
                "images": [],
            }

        item_name = _item_name(record)
        if item_name is None:
            grouped[site_name]["images"].extend(record.get("images", []))
            continue

        if item_name not in grouped[site_name]["items"]:
            grouped[site_name]["items"][item_name] = {"name": item_name, "issues": [], "images": []}
        item = grouped[site_name]["items"][item_name]
        item["images"].extend(record.get("images", []))

        issue_text = _issue_text(record)
        if issue_text:
            _append_unique(item["issues"], issue_text)
            _append_unique(grouped[site_name]["issues"], issue_text)

    sites = []
    for index, site in enumerate(grouped.values(), start=1):
        items = []
        for item_name in PARKING_ITEMS:
            item = site["items"].get(item_name)
            if item is None:
                continue
            result_text = _with_sentence_period("；".join(item["issues"])) if item["issues"] else _default_item_text(item_name)
            items.append(
                {
                    "name": item_name,
                    "result_text": result_text,
                    "photo_rows": _photo_rows(item["images"], image_factory=image_factory),
                }
            )
        site_photo_rows = _photo_rows(site["images"], image_factory=image_factory) if site["images"] else _photo_rows(
            [image for item in items for row in item["photo_rows"] for image in (row["left"], row["right"]) if image],
            image_factory=image_factory,
        )
        sites.append(
            {
                "index_cn": _cn_index(index),
                "name": site["name"],
                "items": items,
                "photo_rows": site_photo_rows,
                "issue_summary": _with_sentence_period("；".join(site["issues"])) if site["issues"] else "无问题。",
            }
        )

    return {
        "title": title or "西城区环卫作业停车场检查报告",
        "report_range_text": f"{start_date}-{end_date}",
        "site_names_text": "；".join(site["name"] for site in sites),
        "department": DEFAULT_DEPARTMENT,
        "sites": sites,
        "closing_text": DEFAULT_CLOSING_TEXT,
        "issue_date_text": _issue_date_from_report_month(report_month),
    }


def summarize_parking_records(records: List[Dict[str, Any]], start_date: str, end_date: str, report_month: str = "") -> Dict[str, Any]:
    """停车场报告专用统计口径。"""

    context = build_parking_station_context(records, start_date, end_date, report_month)
    problem_items = [
        item
        for site in context["sites"]
        for item in site["items"]
        if _normalize_problem_text(item["result_text"]) not in {_normalize_problem_text(_default_item_text(item["name"]))}
    ]
    return {
        "street_count": len(context["sites"]),
        "site_count": len(context["sites"]),
        "record_count": len(records),
        "problem_record_count": len(problem_items),
        "problem_count": len(problem_items),
        "summary_text": (
            f"{context['report_range_text']}西城区城管委固废科对"
            f"{context['site_names_text']}环卫作业停车场进行现场检查。"
        ),
        "table_rows": [],
    }


def generate_parking_station_report(
    records: List[Dict[str, Any]],
    output_path: Path,
    title: str,
    start_date: str,
    end_date: str,
    report_month: str,
    template_path: Path,
) -> Path:
    """使用 docxtpl 渲染环卫停车场检查报告。"""

    if output_path.resolve() == template_path.resolve():
        raise ValueError("输出文件不能覆盖模板文件，请更换输出文件名或输出目录。")

    try:
        from docxtpl import DocxTemplate, InlineImage
    except ImportError as exc:
        raise RuntimeError("缺少 docxtpl/jinja2 依赖，请运行 run_app.bat 重新安装 requirements.txt。") from exc

    template = DocxTemplate(str(template_path))

    def image_factory(path: str) -> Any:
        return InlineImage(template, path, width=Cm(7.21), height=Cm(4.06))

    context = build_parking_station_context(
        records=records,
        start_date=start_date,
        end_date=end_date,
        report_month=report_month,
        title=title,
        image_factory=image_factory,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    template.render({key: _to_attr_object(value) for key, value in context.items()})
    template.save(output_path)
    return output_path
