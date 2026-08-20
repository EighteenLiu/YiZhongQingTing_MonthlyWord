"""生活垃圾分类驿站 Jinja 模板检查通报生成。"""

from __future__ import annotations

from collections import Counter, OrderedDict
import math
from pathlib import Path
import re
import tempfile
from typing import Any, Callable, Dict, Iterable, List
from zipfile import ZipFile

from docx.shared import Cm

from core.data_cleaner import _is_sorting_station_good_result, detect_indicators, is_non_problem_explicit_value
from core.report_generator import _format_problem_text, _split_problem_parts
from utils.file_utils import TEMP_DIR


def _compact_metric(value: Any) -> str:
    text = str(value or "").strip()
    replacements = {
        " ": "",
        "\n": "",
        "\r": "",
        "\t": "",
        "（": "(",
        "）": ")",
        "，": ",",
        "；": ";",
        "：": ":",
        "台帐": "台账",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    return re.sub(r"\s+", "", text).strip("。；;，,")


SORTING_METRICS = [
    "四分类桶不成组",
    "无驿站公示牌或驿站公示牌不合格",
    "未按时开门运行",
    "无备案公示",
    "无称重系统或称重系统损坏",
    "驿站周边环境脏乱",
    "驿站内环境脏乱",
    "使用大功率电器",
    "备案公示过期",
    "无人值守",
    "无回收服务五公开或回收服务五公开不齐全",
    "无回收人员信息",
    "无安全风险公告",
    "容器颜色不符",
    "无便利性措施",
    "居民居住",
    "无消防设备器材或不合格",
    "灭火器缺少",
    "无消防水源",
    "无企安安",
    "无消防安全管理制度或不齐全",
    "无消防安全应急预案",
    "容器无标识、标识不符、脱落、破损",
    "无防火标识",
    "无安全生产责任制度",
    "无价格公式",
    "无隐患排查台账或隐患排查台账不合格",
    "无消防培训演练资料",
    "无生活垃圾分类管理台账",
    "无巡查记录表",
    "无精细化管理台账",
]
SORTING_METRIC_ALIASES = {
    "未开门运行": "未按时开门运行",
    "无消防设备器材或消防设备器不合格": "无消防设备器材或不合格",
    "无消防培训演练资料或消防培训演练资料不合格": "无消防培训演练资料",
    "无价格公示": "无价格公式",
    "备案公示": "无备案公示",
    "备案公示缺少": "无备案公示",
    "无备案信息": "无备案公示",
    "企安安": "无企安安",
    "无生活垃圾分类管理台帐": "无生活垃圾分类管理台账",
    "细化台帐": "无精细化管理台账",
    "细化台账": "无精细化管理台账",
    "称重": "无称重系统或称重系统损坏",
    "称重计量": "无称重系统或称重系统损坏",
    "周边不洁": "驿站周边环境脏乱",
    "堆积杂物": "驿站内环境脏乱",
    "杂物堆积": "驿站内环境脏乱",
    "备案信息已过期": "备案公示过期",
    "备案信息过期": "备案公示过期",
    "无精细化管理台帐": "无精细化管理台账",
    "无隐患排查台帐或隐患排查台帐不合格": "无隐患排查台账或隐患排查台账不合格",
}
SORTING_METRIC_LOOKUP = {
    _compact_metric(metric): metric for metric in SORTING_METRICS
}
SORTING_METRIC_LOOKUP.update(
    {_compact_metric(alias): canonical for alias, canonical in SORTING_METRIC_ALIASES.items()}
)
NO_ISSUE_TEXTS = {"", "1", "无", "无问题", "无问题。", "正常", "合格", "良好，未发现问题"}
EXPLICIT_NO_ISSUE_TEXTS = {
    "无问题",
    "无问题。",
    "未发现问题",
    "未发现问题。",
    "暂无问题",
    "没有问题",
    "无异常",
    "未见异常",
    "正常",
    "合格",
    "良好，未发现问题",
}
CN_NUMBERS = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]


ImageFactory = Callable[[str], Any]


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value or "").strip()


def _compact(value: Any) -> str:
    return re.sub(r"\s+", "", _text(value)).strip("。；;，,")


def _canonical_metric(value: Any) -> str:
    return SORTING_METRIC_LOOKUP.get(_compact_metric(value), "")


def _contained_metrics(value: Any) -> List[str]:
    if _is_sorting_station_good_result(value):
        return []

    compact_value = _compact_metric(value)
    if not compact_value:
        return []

    metrics: List[str] = []
    for phrase, metric in SORTING_METRIC_LOOKUP.items():
        if phrase and phrase in compact_value and metric not in metrics:
            metrics.append(metric)
    return metrics


def _is_generic_ledger_phrase(value: Any) -> bool:
    compact_value = _compact_metric(value)
    if "台账" not in compact_value:
        return False
    specific_ledger_words = ["隐患排查", "生活垃圾分类", "巡查", "精细化"]
    return not any(word in compact_value for word in specific_ledger_words)


def _split_metric_parts(value: Any) -> List[str]:
    text = _text(value)
    if not text:
        return []
    parts: List[str] = []
    for part in re.split(r"[;；、,，\s]+", text):
        part = part.strip()
        if part:
            parts.append(part)
    return parts


def _no_issue_compacts() -> set[str]:
    return {_compact(text) for text in NO_ISSUE_TEXTS}


def _explicit_no_issue_compacts() -> set[str]:
    return {_compact(text) for text in EXPLICIT_NO_ISSUE_TEXTS}


def _cn_index(index: int) -> str:
    if 1 <= index <= len(CN_NUMBERS):
        return CN_NUMBERS[index - 1]
    if index < 20:
        return "十" + CN_NUMBERS[index - 11]
    return str(index)


def _raw_value(record: Dict[str, Any], key: str) -> str:
    return _text((record.get("raw") or {}).get(key))


def _street_name(record: Dict[str, Any]) -> str:
    return _text(record.get("report_group")) or _text(record.get("street")) or _raw_value(record, "3级点位") or "未识别街道"


def _point_name(record: Dict[str, Any]) -> str:
    return _text(record.get("report_point")) or _text(record.get("location")) or _raw_value(record, "4级点位") or "未识别点位"


def _source_problem(record: Dict[str, Any]) -> str:
    raw_problem = _raw_value(record, "具体问题")
    if raw_problem:
        return raw_problem
    return _text(record.get("specific_problem") or record.get("problem"))


def _secondary_indicator(record: Dict[str, Any]) -> str:
    return _text(record.get("secondary_indicator")) or _raw_value(record, "2级指标")


def _tertiary_indicator(record: Dict[str, Any]) -> str:
    return _text(record.get("tertiary_indicator")) or _raw_value(record, "3级指标")


def _append_metric(metrics: List[str], metric: str) -> None:
    if metric and metric not in metrics:
        metrics.append(metric)


def _normalize_metric_priority(metrics: List[str]) -> List[str]:
    if "备案公示过期" in metrics and "无备案公示" in metrics:
        return [metric for metric in metrics if metric != "无备案公示"]
    return metrics


def _metrics_from_text_parts(value: Any, record: Dict[str, Any]) -> List[str]:
    metrics: List[str] = []
    for part in _split_metric_parts(value) or _split_problem_parts(_text(value)) or [_text(value)]:
        metric = _canonical_metric(part)
        _append_metric(metrics, metric)
        for contained_metric in _contained_metrics(part):
            _append_metric(metrics, contained_metric)
        for indicator in detect_indicators(part, {}, report_type="sorting_station"):
            _append_metric(metrics, _canonical_metric(indicator))
        if _is_generic_ledger_phrase(part):
            _append_metric(metrics, "无精细化管理台账")

    for indicator in detect_indicators(_text(value), record.get("raw") or {}, report_type="sorting_station"):
        _append_metric(metrics, _canonical_metric(indicator))
    return _normalize_metric_priority(metrics)


def _has_explicit_metric_value(record: Dict[str, Any]) -> bool:
    for key, value in (record.get("raw") or {}).items():
        if not _canonical_metric(key):
            continue
        explicit_value = _compact(value)
        if explicit_value and not is_non_problem_explicit_value(explicit_value, "sorting_station"):
            return True
    return False


def _is_no_issue(record: Dict[str, Any]) -> bool:
    if _compact(_raw_value(record, "具体问题")) in _explicit_no_issue_compacts():
        return True

    if _has_explicit_metric_value(record):
        return False

    tertiary = _tertiary_indicator(record)
    if tertiary:
        return _compact(tertiary) in _no_issue_compacts() or not _metrics_from_text_parts(tertiary, record)

    secondary = _secondary_indicator(record)
    if secondary:
        if _compact(secondary) not in _no_issue_compacts():
            return not _metrics_from_text_parts(secondary, record)
        source = _source_problem(record)
        return not source or _compact(source) in _no_issue_compacts()

    source = _source_problem(record)
    return _compact(source) in _no_issue_compacts()


def _record_metrics(record: Dict[str, Any]) -> List[str]:
    if _is_no_issue(record):
        return []

    tertiary = _tertiary_indicator(record)
    secondary = _secondary_indicator(record)
    indicator_source = tertiary or secondary or _source_problem(record)
    metrics: List[str] = []
    for metric in _metrics_from_text_parts(indicator_source, record):
        _append_metric(metrics, metric)
    source_problem = _source_problem(record)
    if not tertiary and not secondary and source_problem and _compact(source_problem) not in _no_issue_compacts():
        for metric in _metrics_from_text_parts(source_problem, record):
            _append_metric(metrics, metric)
    for key, value in (record.get("raw") or {}).items():
        metric = _canonical_metric(key)
        if not metric or metric in metrics:
            continue
        explicit_value = _compact(value)
        if explicit_value and not is_non_problem_explicit_value(explicit_value, "sorting_station"):
            metrics.append(metric)
    return metrics


def _issue_parts(record: Dict[str, Any], metrics: List[str]) -> List[str]:
    if _is_no_issue(record):
        return []

    tertiary = _tertiary_indicator(record)
    if tertiary:
        tertiary_metrics = _metrics_from_text_parts(tertiary, record)
        if tertiary_metrics:
            return tertiary_metrics
        return []

    secondary = _secondary_indicator(record)
    if secondary:
        secondary_metrics = _metrics_from_text_parts(secondary, record)
        if secondary_metrics:
            return secondary_metrics
        return []

    source_problem = _source_problem(record)
    if source_problem and _compact(source_problem) not in _no_issue_compacts():
        problem_metrics = _metrics_from_text_parts(source_problem, record)
        if problem_metrics:
            return problem_metrics

    source = source_problem
    if _compact(source) in _no_issue_compacts():
        return metrics
    return _split_problem_parts(source) or metrics


def _append_unique(parts: List[str], value: str) -> bool:
    text = value.strip()
    if text and text not in parts:
        parts.append(text)
        return True
    return False


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


def _complete_metric_counts(metrics: Counter) -> Dict[str, int]:
    return {metric: metrics.get(metric, 0) for metric in SORTING_METRICS}


def _group_points(records: List[Dict[str, Any]]) -> "OrderedDict[str, OrderedDict[str, Dict[str, Any]]]":
    grouped: "OrderedDict[str, OrderedDict[str, Dict[str, Any]]]" = OrderedDict()
    for record in records:
        street = _street_name(record)
        point = _point_name(record)
        grouped.setdefault(street, OrderedDict())
        points = grouped[street]
        if point not in points:
            points[point] = {"display_name": point, "issue_parts": [], "metrics": Counter(), "images": []}

        point_data = points[point]
        metrics = _record_metrics(record)
        for issue in _issue_parts(record, metrics):
            if _append_unique(point_data["issue_parts"], issue):
                for metric in _metrics_from_text_parts(issue, {"raw": {}}):
                    point_data["metrics"][metric] += 1
        point_data["images"].extend(record.get("images", []))
    return grouped


def build_sorting_station_context(
    records: List[Dict[str, Any]],
    start_date: str,
    end_date: str,
    title: str = "西城区生活垃圾分类驿站检查通报",
    image_factory: ImageFactory | None = None,
) -> Dict[str, Any]:
    """把清洗后的驿站记录转换为 Jinja 模板上下文。"""

    grouped = _group_points(records)
    streets = []
    metric_rows = []
    total_metrics = Counter()
    notified_problem_site_count = 0

    for street_index, (street_name, points) in enumerate(grouped.items(), start=1):
        point_context = []
        street_metrics = Counter()
        for point in points.values():
            issue_text = _format_problem_text("；".join(point["issue_parts"])) if point["issue_parts"] else ""
            if issue_text:
                notified_problem_site_count += 1
            street_metrics.update(point["metrics"])
            if not issue_text and not point["images"]:
                continue
            point_context.append(
                {
                    "display_name": point["display_name"],
                    "issue_text": issue_text,
                    "photo_rows": _photo_rows(point["images"], image_factory=image_factory),
                    # 兼容当前模板未修正前的 if point.photos 判断。
                    "photos": point["images"],
                }
            )

        total_metrics.update(street_metrics)
        streets.append(
            {
                "index_cn": _cn_index(street_index),
                "name": street_name,
                "points": point_context,
            }
        )
        metric_rows.append({"street_name": street_name, "metrics": _complete_metric_counts(street_metrics)})

    metric_rows.insert(0, {"street_name": "总计", "metrics": _complete_metric_counts(total_metrics)})

    return {
        "title": title or "西城区生活垃圾分类驿站检查通报",
        "report_range_text": f"{start_date}至{end_date}",
        "street_count": len(streets),
        "check_count": len(records),
        "notified_problem_site_count": notified_problem_site_count,
        "streets": streets,
        "metric_rows": metric_rows,
    }


def summarize_sorting_station_records(records: List[Dict[str, Any]], start_date: str, end_date: str) -> Dict[str, Any]:
    """生活垃圾分类驿站 Jinja 报告专用统计口径。"""

    context = build_sorting_station_context(records, start_date, end_date)
    return {
        "street_count": context["street_count"],
        "record_count": context["check_count"],
        "problem_record_count": context["notified_problem_site_count"],
        "problem_count": context["notified_problem_site_count"],
        "summary_text": (
            f"{context['report_range_text']}对本区{context['street_count']}个街道生活垃圾分类驿站"
            f"进行了{context['check_count']}个次检查。检查发现需通报问题站次{context['notified_problem_site_count']}个。"
        ),
        "table_rows": [],
    }


def _patch_sorting_template(template_path: Path, temp_dir: Path | None = None) -> Path:
    """修正当前驿站 Jinja 模板中的 point.photo_rows 变量不一致。"""

    with ZipFile(template_path, "r") as source:
        entries = {name: source.read(name) for name in source.namelist()}

    changed = False
    for name, data in list(entries.items()):
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        text = data.decode("utf-8", errors="ignore")
        patched = text.replace("point.photos", "point.photo_rows").replace("item.photo_rows", "point.photo_rows")
        if patched != text:
            entries[name] = patched.encode("utf-8")
            changed = True

    if not changed:
        return template_path

    target_temp_dir = temp_dir or TEMP_DIR
    target_temp_dir.mkdir(parents=True, exist_ok=True)
    temp_file = tempfile.NamedTemporaryFile(
        prefix=f"{template_path.stem}_patched_",
        suffix=".docx",
        dir=target_temp_dir,
        delete=False,
    )
    temp_name = temp_file.name
    temp_file.close()
    with ZipFile(temp_name, "w") as target:
        for name, data in entries.items():
            target.writestr(name, data)
    return Path(temp_name)


def generate_sorting_station_report(
    records: List[Dict[str, Any]],
    stats: Dict[str, Any],
    output_path: Path,
    title: str,
    start_date: str,
    end_date: str,
    report_month: str,
    template_path: Path | None = None,
    temp_dir: Path | None = None,
) -> Path:
    """使用 docxtpl 渲染生活垃圾分类驿站检查通报。"""

    if template_path is None:
        raise ValueError("生活垃圾分类驿站检查通报必须提供 Jinja 模板。")
    if output_path.resolve() == template_path.resolve():
        raise ValueError("输出文件不能覆盖模板文件，请更换输出文件名或输出目录。")

    try:
        from docxtpl import DocxTemplate, InlineImage
    except ImportError as exc:
        raise RuntimeError("缺少 docxtpl/jinja2 依赖，请运行 run_app.bat 重新安装 requirements.txt。") from exc

    patched_template = _patch_sorting_template(template_path, temp_dir=temp_dir)
    template = DocxTemplate(str(patched_template))

    def image_factory(path: str) -> Any:
        return InlineImage(template, path, width=Cm(7.21), height=Cm(4.06))

    context = build_sorting_station_context(
        records=records,
        start_date=start_date,
        end_date=end_date,
        title=title,
        image_factory=image_factory,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    template.render(context)
    template.save(output_path)
    return output_path
