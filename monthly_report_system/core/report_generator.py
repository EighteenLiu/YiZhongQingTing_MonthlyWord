"""Word 月报生成，支持基于上传模板写入内容。"""

from __future__ import annotations

from collections import OrderedDict
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, Iterable, List

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from config.indicators import INDICATORS, REPORT_STYLE, ReportStyle


CHINESE_NUMBERS = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
A4_WIDTH_CM = 21
A4_HEIGHT_CM = 29.7


def _clear_body_keep_section(document: Document) -> None:
    """清空模板正文内容，同时保留 sectPr、样式、页眉页脚和页面设置。"""

    body = document._body._element  # noqa: SLF001
    section_properties = body.sectPr
    saved_section_properties = deepcopy(section_properties) if section_properties is not None else None
    for child in list(body):
        body.remove(child)
    if saved_section_properties is not None:
        body.append(saved_section_properties)


def _set_run_font(run, font_name: str, size_pt: int, bold: bool = False) -> None:
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)  # noqa: SLF001
    run.font.size = Pt(size_pt)
    run.bold = bold


def _set_paragraph_font(paragraph, font_name: str, size_pt: int, bold: bool = False) -> None:
    for run in paragraph.runs:
        _set_run_font(run, font_name, size_pt, bold)


def _set_cell_text(cell, text: str, font_name: str, size_pt: int, bold: bool = False) -> None:
    cell.text = str(text)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_paragraph_font(paragraph, font_name, size_pt, bold)


def _set_table_borders(table) -> None:
    tbl = table._tbl  # noqa: SLF001
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "6")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "000000")
        borders.append(element)


def _set_a4_section(section, orientation: WD_ORIENT, style: ReportStyle) -> None:
    """设置 A4 页面方向和边距。"""

    section.orientation = orientation
    if orientation == WD_ORIENT.LANDSCAPE:
        section.page_width = Cm(A4_HEIGHT_CM)
        section.page_height = Cm(A4_WIDTH_CM)
    else:
        section.page_width = Cm(A4_WIDTH_CM)
        section.page_height = Cm(A4_HEIGHT_CM)
    section.top_margin = Cm(style.page_margin_cm)
    section.bottom_margin = Cm(style.page_margin_cm)
    section.left_margin = Cm(style.page_margin_cm)
    section.right_margin = Cm(style.page_margin_cm)


def _add_landscape_section(document: Document, style: ReportStyle) -> None:
    """为附件统计表创建独立横向 A4 分节。"""

    section = document.add_section(WD_SECTION.NEW_PAGE)
    _set_a4_section(section, WD_ORIENT.LANDSCAPE, style)


def _add_title(document: Document, title: str, style: ReportStyle) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(title)
    _set_run_font(run, style.title_font, style.title_size_pt, True)


def _add_heading(document: Document, text: str, level: int) -> None:
    """按 Word 标题级别添加标题，优先使用模板中的标题样式。"""

    paragraph = document.add_heading(text, level=level)
    paragraph.paragraph_format.first_line_indent = Pt(0)


def _add_body_paragraph(document: Document, text: str, style: ReportStyle, indent: bool = True) -> None:
    paragraph = document.add_paragraph()
    if indent:
        paragraph.paragraph_format.first_line_indent = Pt(style.body_size_pt * style.first_line_indent_chars)
    run = paragraph.add_run(text)
    _set_run_font(run, style.body_font, style.body_size_pt)


def _add_point_problem_paragraph(document: Document, heading_text: str, problem_text: str, style: ReportStyle) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.first_line_indent = Pt(0)
    problem_text = " ".join(str(problem_text).splitlines()).strip()
    run = paragraph.add_run(f"{heading_text}{problem_text}")
    _set_run_font(run, style.body_font, style.body_size_pt)


def _add_images(document: Document, image_paths: Iterable[str], style: ReportStyle) -> None:
    """在同一个段落中连续插入图片，不使用表格，也不插入空格。"""

    paths = list(image_paths)[: style.pictures_per_row]
    if not paths:
        return

    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    for image_path in paths:
        run = paragraph.add_run()
        try:
            run.add_picture(image_path, width=Cm(style.image_width_cm), height=Cm(style.image_height_cm))
        except Exception:
            run.add_text("[图片插入失败]")


def _add_statistics_table(document: Document, rows: List[Dict[str, Any]], style: ReportStyle) -> None:
    columns = ["街道名称", *INDICATORS, "问题总数"]
    table = document.add_table(rows=1, cols=len(columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for i, column in enumerate(columns):
        _set_cell_text(table.rows[0].cells[i], column, style.body_font, style.table_size_pt, True)
    for row_data in rows:
        cells = table.add_row().cells
        for i, column in enumerate(columns):
            value = row_data.get(column, 0)
            if isinstance(value, int):
                value = style.table_zero_text if value == 0 else str(value)
            _set_cell_text(cells[i], value, style.body_font, style.table_size_pt)
    _set_table_borders(table)


def _new_document(template_path: Path | None, style: ReportStyle) -> Document:
    """创建报告文档；有模板时继承模板格式，无模板时使用默认页面。"""

    if template_path:
        document = Document(str(template_path))
        _clear_body_keep_section(document)
        for section in document.sections:
            _set_a4_section(section, WD_ORIENT.PORTRAIT, style)
        return document

    document = Document()
    _set_a4_section(document.sections[0], WD_ORIENT.PORTRAIT, style)
    return document


def _cn_index(index: int) -> str:
    """把 1-based 序号转为简单中文序号。"""

    if 1 <= index <= len(CHINESE_NUMBERS):
        return CHINESE_NUMBERS[index - 1]
    if index < 20:
        return "十" + CHINESE_NUMBERS[index - 11]
    return str(index)


def _group_records(records: List[Dict[str, Any]]) -> "OrderedDict[str, List[Dict[str, Any]]]":
    grouped: "OrderedDict[str, List[Dict[str, Any]]]" = OrderedDict()
    for record in records:
        group = record.get("report_group") or record.get("street") or "未识别街道"
        grouped.setdefault(group, []).append(record)
    return grouped


def _problem_text(record: Dict[str, Any]) -> str:
    text = str(record.get("specific_problem") or record.get("problem") or "").strip()
    return text or "无问题"


def generate_report(
    records: List[Dict[str, Any]],
    stats: Dict[str, Any],
    output_path: Path,
    title: str,
    start_date: str,
    end_date: str,
    style: ReportStyle = REPORT_STYLE,
    template_path: Path | None = None,
) -> Path:
    """根据记录和统计结果生成 Word 月报。"""

    document = _new_document(template_path, style)

    _add_title(document, title, style)
    _add_heading(document, "一、总体情况", level=1)
    _add_body_paragraph(document, stats["summary_text"], style)
    _add_heading(document, "二、各街道案例", level=1)

    for group_index, (group_name, group_records) in enumerate(_group_records(records).items(), start=1):
        _add_heading(document, f"（{_cn_index(group_index)}）{group_name}", level=2)
        for point_index, record in enumerate(group_records, start=1):
            point_name = record.get("report_point") or record.get("location") or "未识别点位"
            _add_point_problem_paragraph(document, f"{point_index}.{point_name}可回收物交投点：", _problem_text(record), style)
            _add_images(document, record.get("images", []), style)

    _add_landscape_section(document, style)
    _add_heading(document, f"附件：{start_date}至{end_date}可回收物交投点各项指标问题数", level=1)
    _add_statistics_table(document, stats["table_rows"], style)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path
