"""桌面窗口入口：生成可回收物交投点 Word 月报。"""

from __future__ import annotations

import contextlib
import io
import os
import shutil
import threading
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from utils.file_utils import OUTPUT_DIR, cleanup_old_run_dirs, cleanup_run_dir, create_run_dir, safe_filename, safe_upload_filename
from utils.logger import AppLogger


def _load_processing_modules():
    """延迟加载重依赖，避免窗口启动阶段被 pandas/numpy 环境问题阻断。"""

    try:
        with contextlib.redirect_stderr(io.StringIO()):
            from core.data_cleaner import attach_images_to_records, clean_dataframe
            from core.excel_reader import read_excel_table
            from core.image_extractor import extract_images_by_row
            from core.report_generator import generate_report
            from core.statistics import summarize_records
    except ImportError as exc:
        raise RuntimeError("依赖环境不可用，请先运行 run_app.bat 自动安装 requirements.txt。") from exc
    return attach_images_to_records, clean_dataframe, read_excel_table, extract_images_by_row, generate_report, summarize_records


def format_cn_date(value: str) -> str:
    parsed = datetime.strptime(value, "%Y-%m-%d").date()
    return f"{parsed.year}年{parsed.month}月{parsed.day}日"


def format_cn_month(value: str) -> str:
    parsed = datetime.strptime(value, "%Y-%m").date()
    return f"{parsed.year}年{parsed.month}月"


def default_values() -> dict[str, str]:
    today = date.today()
    return {
        "title": "西城区可回收物交投点检查情况通报",
        "start_date": "2026-04-20",
        "end_date": "2026-05-19",
        "report_month": "2026-05",
        "output_dir": str(OUTPUT_DIR),
        "output_name": f"西城区{today.year}年{today.month}月可回收物交投点检查通报.docx",
    }


def _copy_input_file(source: Path, target_dir: Path, fallback_name: str) -> Path:
    filename = safe_upload_filename(source.name, fallback_name)
    target = target_dir / filename
    shutil.copy2(source, target)
    return target


def process_report(values: dict[str, str]) -> dict:
    """根据窗口表单值生成 Word 月报，返回前端展示结果。"""

    logger = AppLogger()
    excel_source = Path(values.get("excel_path", "")).expanduser()
    template_source = Path(values.get("template_path", "")).expanduser()
    if not excel_source.exists():
        raise ValueError("请选择有效的 Excel 检查数据文件。")
    if excel_source.suffix.lower() not in {".xls", ".xlsx"}:
        raise ValueError("Excel 文件必须是 .xls 或 .xlsx 格式。")
    if not template_source.exists():
        raise ValueError("请选择有效的 Word 模板文件。")
    if template_source.suffix.lower() != ".docx":
        raise ValueError("Word 模板文件必须是 .docx 格式。")

    cleanup_warnings = cleanup_old_run_dirs()
    logger.extend(cleanup_warnings)
    (
        attach_images_to_records,
        clean_dataframe,
        read_excel_table,
        extract_images_by_row,
        generate_report,
        summarize_records,
    ) = _load_processing_modules()

    title = values.get("title", "").strip() or default_values()["title"]
    start_date = format_cn_date(values.get("start_date", default_values()["start_date"]))
    end_date = format_cn_date(values.get("end_date", default_values()["end_date"]))
    report_month = format_cn_month(values.get("report_month", default_values()["report_month"]))
    output_name = safe_filename(values.get("output_name", "月报.docx"))
    output_dir = Path(values.get("output_dir") or OUTPUT_DIR).expanduser()

    run_dir = create_run_dir()
    try:
        excel_path = _copy_input_file(excel_source, run_dir, "data.xlsx")
        template_path = _copy_input_file(template_source, run_dir, "template.docx")

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
        logger.info("Word 版式：正文为竖向 A4，附件统计表为横向 A4；图片按 5.1cm x 3cm 插入且不额外留空格。")

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
        "report_month": report_month,
    }


class MonthlyReportWindow(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("可回收物交投点月报生成系统")
        self.geometry("1160x760")
        self.minsize(980, 640)
        self.result: dict | None = None
        self.vars = {key: tk.StringVar(value=value) for key, value in default_values().items()}
        self.vars["excel_path"] = tk.StringVar()
        self.vars["template_path"] = tk.StringVar()
        self._build_style()
        self._build_ui()

    def _build_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f4f6f8")
        style.configure("Card.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("TLabel", background="#f4f6f8", foreground="#172033", font=("Microsoft YaHei UI", 10))
        style.configure("Card.TLabel", background="#ffffff", foreground="#172033", font=("Microsoft YaHei UI", 10))
        style.configure("Title.TLabel", background="#f4f6f8", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Section.TLabel", background="#ffffff", foreground="#172033", font=("Microsoft YaHei UI", 13, "bold"))
        style.configure("Subsection.TLabel", background="#ffffff", foreground="#172033", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Metric.TLabel", background="#ffffff", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("TButton", font=("Microsoft YaHei UI", 10), padding=(12, 6))
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(16, 7))

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=18)
        root.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(root)
        header.pack(fill=tk.X, pady=(0, 14))
        ttk.Label(header, text="可回收物交投点月报生成系统", style="Title.TLabel").pack(side=tk.LEFT)
        self.status_var = tk.StringVar(value="请选择 Excel 数据和 Word 模板后生成。")
        ttk.Label(header, textvariable=self.status_var).pack(side=tk.RIGHT)

        content = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        content.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(content, style="Card.TFrame", padding=16)
        right = ttk.Frame(content, style="Card.TFrame", padding=16)
        content.add(left, weight=1)
        content.add(right, weight=2)

        self._build_form(left)
        self._build_result(right)

    def _build_form(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="文件与参数", style="Section.TLabel").pack(anchor=tk.W, pady=(0, 12))
        self._file_row(parent, "Excel 检查数据", "excel_path", [("Excel 文件", "*.xls *.xlsx")])
        self._file_row(parent, "Word 月报模板", "template_path", [("Word 文档", "*.docx")])
        self._entry_row(parent, "报告标题", "title")
        self._entry_row(parent, "报告月份（YYYY-MM）", "report_month")
        self._entry_row(parent, "检查开始日期（YYYY-MM-DD）", "start_date")
        self._entry_row(parent, "检查结束日期（YYYY-MM-DD）", "end_date")
        self._dir_row(parent, "输出目录", "output_dir")
        self._entry_row(parent, "输出文件名", "output_name")

        buttons = ttk.Frame(parent, style="Card.TFrame")
        buttons.pack(fill=tk.X, pady=(18, 0))
        self.generate_button = ttk.Button(buttons, text="生成 Word 月报", style="Accent.TButton", command=self._generate)
        self.generate_button.pack(side=tk.LEFT)
        ttk.Button(buttons, text="重置", command=self._reset).pack(side=tk.LEFT, padx=(8, 0))
        self.open_file_button = ttk.Button(buttons, text="打开文件", command=self._open_output, state=tk.DISABLED)
        self.open_file_button.pack(side=tk.RIGHT)

    def _build_result(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="生成结果", style="Section.TLabel").pack(anchor=tk.W, pady=(0, 12))
        metrics = ttk.Frame(parent, style="Card.TFrame")
        metrics.pack(fill=tk.X)
        self.metric_vars = {
            "record_count": tk.StringVar(value="0"),
            "street_count": tk.StringVar(value="0"),
            "image_count": tk.StringVar(value="0"),
            "problem_record_count": tk.StringVar(value="0"),
        }
        for label, key in [("记录数", "record_count"), ("街道数", "street_count"), ("图片数", "image_count"), ("问题记录", "problem_record_count")]:
            item = ttk.Frame(metrics, style="Card.TFrame", padding=10)
            item.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
            ttk.Label(item, text=label, style="Card.TLabel").pack(anchor=tk.W)
            ttk.Label(item, textvariable=self.metric_vars[key], style="Metric.TLabel").pack(anchor=tk.W)

        self.summary_var = tk.StringVar(value="尚未生成月报。")
        ttk.Label(parent, textvariable=self.summary_var, style="Card.TLabel", wraplength=640).pack(fill=tk.X, pady=12)

        columns = ("行号", "街道", "点位", "问题描述", "识别指标", "图片数")
        self.preview = ttk.Treeview(parent, columns=columns, show="headings", height=10)
        for column in columns:
            self.preview.heading(column, text=column)
            self.preview.column(column, width=90 if column in {"行号", "图片数"} else 140, stretch=True)
        self.preview.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        ttk.Label(parent, text="运行日志", style="Subsection.TLabel").pack(anchor=tk.W)
        self.log_text = tk.Text(parent, height=8, wrap=tk.WORD, relief=tk.SOLID, borderwidth=1)
        self.log_text.pack(fill=tk.BOTH, expand=False, pady=(6, 0))

    def _file_row(self, parent: ttk.Frame, label: str, key: str, filetypes) -> None:
        row = self._row(parent, label)
        ttk.Entry(row, textvariable=self.vars[key]).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="选择", command=lambda: self._choose_file(key, filetypes)).pack(side=tk.LEFT, padx=(8, 0))

    def _dir_row(self, parent: ttk.Frame, label: str, key: str) -> None:
        row = self._row(parent, label)
        ttk.Entry(row, textvariable=self.vars[key]).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="选择", command=lambda: self._choose_dir(key)).pack(side=tk.LEFT, padx=(8, 0))

    def _entry_row(self, parent: ttk.Frame, label: str, key: str) -> None:
        row = self._row(parent, label)
        ttk.Entry(row, textvariable=self.vars[key]).pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _row(self, parent: ttk.Frame, label: str) -> ttk.Frame:
        ttk.Label(parent, text=label, style="Card.TLabel").pack(anchor=tk.W, pady=(8, 4))
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill=tk.X)
        return row

    def _choose_file(self, key: str, filetypes) -> None:
        path = filedialog.askopenfilename(title="选择文件", filetypes=filetypes)
        if path:
            self.vars[key].set(path)

    def _choose_dir(self, key: str) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.vars[key].set(path)

    def _reset(self) -> None:
        for key, value in default_values().items():
            self.vars[key].set(value)
        self.vars["excel_path"].set("")
        self.vars["template_path"].set("")

    def _generate(self) -> None:
        values = {key: var.get() for key, var in self.vars.items()}
        self.generate_button.configure(state=tk.DISABLED)
        self.open_file_button.configure(state=tk.DISABLED)
        self.status_var.set("正在生成，请稍候...")
        self._set_logs(["[INFO] 开始生成 Word 月报..."])
        threading.Thread(target=self._run_generate, args=(values,), daemon=True).start()

    def _run_generate(self, values: dict[str, str]) -> None:
        try:
            result = process_report(values)
        except Exception as exc:  # noqa: BLE001 - 窗口需要展示所有处理错误
            self.after(0, self._generation_failed, exc)
        else:
            self.after(0, self._generation_done, result)

    def _generation_done(self, result: dict) -> None:
        self.result = result
        for key, var in self.metric_vars.items():
            var.set(str(result.get(key, 0)))
        self.summary_var.set(result.get("summary_text") or "生成完成。")
        for item in self.preview.get_children():
            self.preview.delete(item)
        for row in result.get("preview", []):
            self.preview.insert("", tk.END, values=tuple(row.get(column, "") for column in self.preview["columns"]))
        self._set_logs(result.get("logs", []))
        self.generate_button.configure(state=tk.NORMAL)
        self.open_file_button.configure(state=tk.NORMAL)
        self.status_var.set(f"生成完成：{result.get('output_path')}")
        messagebox.showinfo("生成完成", f"Word 月报已生成：\n{result.get('output_path')}")

    def _generation_failed(self, exc: Exception) -> None:
        self.generate_button.configure(state=tk.NORMAL)
        self.status_var.set("生成失败。")
        self._set_logs([f"[ERROR] {exc}"])
        messagebox.showerror("生成失败", str(exc))

    def _set_logs(self, logs: list[str]) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.insert(tk.END, "\n".join(logs) if logs else "暂无日志。")
        self.log_text.configure(state=tk.DISABLED)

    def _open_output(self) -> None:
        if not self.result:
            return
        path = self.result.get("output_path")
        if path and Path(path).exists():
            os.startfile(path)  # noqa: S606 - 桌面程序按用户点击打开生成文件
        else:
            messagebox.showwarning("文件不存在", "输出文件不存在或已被移动。")


def main() -> None:
    app = MonthlyReportWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
