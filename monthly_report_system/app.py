"""Flask 前端入口。"""

from __future__ import annotations

import contextlib
import io
from datetime import date, datetime
from pathlib import Path

from flask import Flask, flash, redirect, render_template, request, send_file, url_for

from utils.file_utils import OUTPUT_DIR, cleanup_old_run_dirs, cleanup_run_dir, create_run_dir, safe_filename, safe_upload_filename, save_upload
from utils.logger import AppLogger


app = Flask(__name__)
app.secret_key = "monthly-report-system"


def _load_processing_modules():
    """延迟加载重依赖，避免环境问题导致首页无法启动。"""

    try:
        with contextlib.redirect_stderr(io.StringIO()):
            from core.data_cleaner import attach_images_to_records, clean_dataframe
            from core.excel_reader import read_excel_table
            from core.image_extractor import extract_images_by_row
            from core.report_generator import generate_report
            from core.statistics import summarize_records
    except ImportError as exc:
        raise RuntimeError("依赖环境不可用，请先执行 pip install -r requirements.txt，确保 numpy<2 与 pandas 版本兼容。") from exc
    return attach_images_to_records, clean_dataframe, read_excel_table, extract_images_by_row, generate_report, summarize_records


def format_cn_date(value: str) -> str:
    """把 HTML 日期值格式化为中文日期。"""

    parsed = datetime.strptime(value, "%Y-%m-%d").date()
    return f"{parsed.year}年{parsed.month}月{parsed.day}日"


def format_cn_month(value: str) -> str:
    """把 HTML 月份值格式化为中文月份。"""

    parsed = datetime.strptime(value, "%Y-%m").date()
    return f"{parsed.year}年{parsed.month}月"


def default_context() -> dict:
    """页面默认参数。"""

    today = date.today()
    return {
        "title": "西城区可回收物交投点检查情况通报",
        "start_date": "2026-04-20",
        "end_date": "2026-05-19",
        "report_month": "2026-05",
        "output_dir": str(OUTPUT_DIR),
        "output_name": f"西城区{today.year}年{today.month}月可回收物交投点检查通报.docx",
        "result": None,
    }


def _save_uploaded_file(file_storage, target_dir: Path, fallback_name: str) -> Path:
    """保存 Flask 上传文件，保留中文文件名。"""

    filename = safe_upload_filename(file_storage.filename or fallback_name, fallback_name)
    return save_upload(file_storage.stream, target_dir, filename)


def _process_request(form, files) -> dict:
    """处理上传文件并生成月报。"""

    logger = AppLogger()
    excel_file = files.get("excel_file")
    template_file = files.get("template_file")
    if not excel_file or not excel_file.filename:
        raise ValueError("请上传 Excel 检查数据文件。")
    if not template_file or not template_file.filename:
        raise ValueError("请上传 Word 模板文件。")
    if Path(template_file.filename).suffix.lower() != ".docx":
        raise ValueError("模板文件必须是 .docx 格式。")

    cleanup_warnings = cleanup_old_run_dirs()
    (
        attach_images_to_records,
        clean_dataframe,
        read_excel_table,
        extract_images_by_row,
        generate_report,
        summarize_records,
    ) = _load_processing_modules()
    title = form.get("title", "").strip() or "西城区可回收物交投点检查情况通报"
    start_date = format_cn_date(form.get("start_date", "2026-04-20"))
    end_date = format_cn_date(form.get("end_date", "2026-05-19"))
    report_month = format_cn_month(form.get("report_month", "2026-05"))
    output_name = safe_filename(form.get("output_name", "月报.docx"))
    output_dir = Path(form.get("output_dir") or OUTPUT_DIR).expanduser()

    run_dir = create_run_dir()
    logger.extend(cleanup_warnings)
    try:
        excel_path = _save_uploaded_file(excel_file, run_dir, "data.xlsx")
        template_path = _save_uploaded_file(template_file, run_dir, "template.docx")

        df, actual_excel_path = read_excel_table(excel_path, run_dir)
        logger.extend(df.attrs.get("warnings", []))
        clean = clean_dataframe(df)
        logger.extend(clean.warnings)

        images_by_row, image_warnings = extract_images_by_row(actual_excel_path, run_dir / "images")
        logger.extend(image_warnings)
        records, attach_warnings = attach_images_to_records(clean.records, images_by_row)
        logger.extend(attach_warnings)
        stats = summarize_records(records, start_date, end_date)

        output_path = output_dir / output_name
        generate_report(
            records=records,
            stats=stats,
            output_path=output_path,
            title=title,
            start_date=start_date,
            end_date=end_date,
            template_path=template_path,
        )
        logger.info("Word 版式：点位与具体问题写入同一段同一文本；图片按 7cm × 4.5cm 插入，超出页面宽度由 Word 自动换行。")

        preview = [
            {
                "行号": r.get("row_number"),
                "街道": r.get("street"),
                "点位": r.get("location"),
                "问题描述": r.get("problem"),
                "识别指标": "、".join(r.get("indicators", [])) or "无",
                "图片数": len(r.get("images", [])),
            }
            for r in records[:20]
        ]
    finally:
        cleanup_message = cleanup_run_dir(run_dir)
        if cleanup_message:
            logger.warning(cleanup_message)
        else:
            logger.info(f"已自动清理本次临时目录：{run_dir.name}")

    return {
        "record_count": stats.get("record_count", 0),
        "street_count": stats.get("street_count", 0),
        "image_count": sum(len(r.get("images", [])) for r in records),
        "problem_record_count": stats.get("problem_record_count", 0),
        "field_mapping": clean.field_mapping,
        "missing_fields": clean.missing_fields,
        "summary_text": stats.get("summary_text", ""),
        "preview": preview,
        "logs": logger.messages,
        "output_path": str(output_path),
        "download_url": url_for("download_file", path=str(output_path)),
        "report_month": report_month,
    }


@app.route("/", methods=["GET", "POST"])
def index():
    """首页：上传 Excel 和模板并生成月报。"""

    context = default_context()
    if request.method == "POST":
        context.update(request.form.to_dict())
        try:
            context["result"] = _process_request(request.form, request.files)
            flash("月报生成成功。", "success")
        except (ValueError, PermissionError, RuntimeError) as exc:
            flash(f"处理失败：{exc}", "error")
        except Exception as exc:  # noqa: BLE001 - 前端需要展示所有运行错误
            flash(f"处理失败：{exc}", "error")
    return render_template("index.html", **context)


@app.route("/download")
def download_file():
    """下载生成的 Word 文件。"""

    path = Path(request.args.get("path", ""))
    if not path.exists():
        flash("下载失败：文件不存在。", "error")
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True, download_name=path.name)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8502, debug=False)
