"""桌面窗口入口：生成可回收物交投点 Word 月报。"""

from __future__ import annotations

import calendar
import contextlib
import io
import os
import sys
import threading
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core.models import ReportParams
from utils.file_utils import OUTPUT_DIR


if getattr(sys, "frozen", False):
    APP_DIR = Path(getattr(sys, "_MEIPASS")).resolve()
else:
    APP_DIR = Path(__file__).resolve().parents[1]


REPORT_TYPES = {
    "transfer_station": {
        "nav_text": "生成交投点检查通报",
        "window_title": "交投点检查通报生成系统",
        "template_label": "交投点检查通报模板",
        "title": "西城区可回收物交投点检查情况通报",
        "output_suffix": "可回收物交投点检查通报.docx",
    },
    "parking_station": {
        "nav_text": "生成环卫停车场检查报告",
        "window_title": "环卫停车场检查报告生成系统",
        "template_label": "环卫停车场检查报告模板",
        "title": "西城区环卫作业停车场检查报告",
        "output_suffix": "环卫停车场检查报告.docx",
        "default_template": str(APP_DIR / "input" / "module" / "环卫停车场检查报告_Jinja模板.docx"),
    },
    "kitchen_waste": {
        "nav_text": "生成厨余垃圾就地处理检查报告",
        "window_title": "厨余垃圾就地处理检查报告生成系统",
        "template_label": "厨余垃圾就地处理检查报告模板",
        "title": "西城区厨余垃圾就地处理检查报告",
        "output_suffix": "厨余垃圾就地处理检查报告.docx",
        "default_template": str(APP_DIR / "input" / "module" / "厨余垃圾就地处理检查报告_Jinja模板.docx"),
    },
    "sorting_station": {
        "nav_text": "生成垃圾分类驿站检查通报",
        "window_title": "垃圾分类驿站检查通报生成系统",
        "template_label": "垃圾分类驿站检查通报模板",
        "title": "西城区生活垃圾分类驿站检查通报",
        "output_suffix": "生活垃圾分类驿站检查通报.docx",
        "default_template": str(APP_DIR / "input" / "module" / "生活垃圾分类驿站检查通报_Jinja模板.docx"),
    },
}
DEFAULT_REPORT_TYPE = "transfer_station"


def _load_pipeline():
    """延迟加载重依赖，避免窗口启动阶段被 pandas/numpy 问题阻断。"""

    try:
        with contextlib.redirect_stderr(io.StringIO()):
            from core.pipeline import process_report
    except ImportError as exc:
        raise RuntimeError("依赖环境不可用，请先运行 run_app.bat 自动安装 requirements.txt。") from exc
    return process_report


def format_cn_date(value: str) -> str:
    """把 YYYY-MM-DD 转为中文日期。"""

    parsed = datetime.strptime(value, "%Y-%m-%d").date()
    return f"{parsed.year}年{parsed.month}月{parsed.day}日"


def format_cn_month(value: str) -> str:
    """把 YYYY-MM 转为中文月份。"""

    parsed = datetime.strptime(value, "%Y-%m").date()
    return f"{parsed.year}年{parsed.month}月"


def output_name_for(report_type_key: str, end_date: str, fallback_month: str = "") -> str:
    """根据检查结束日期生成默认文件名。"""

    report_type = REPORT_TYPES.get(report_type_key, REPORT_TYPES[DEFAULT_REPORT_TYPE])
    try:
        parsed = datetime.strptime(end_date, "%Y-%m-%d").date()
    except ValueError:
        try:
            parsed = datetime.strptime(fallback_month, "%Y-%m").date()
        except ValueError:
            parsed = date.today()
    return f"西城区{parsed.year}年{parsed.month}月{report_type['output_suffix']}"


def default_values() -> dict[str, str]:
    """窗口默认值。"""

    report_type = REPORT_TYPES[DEFAULT_REPORT_TYPE]
    report_month = "2026-05"
    return {
        "report_type": DEFAULT_REPORT_TYPE,
        "title": report_type["title"],
        "start_date": "2026-04-20",
        "end_date": "2026-05-19",
        "report_month": report_month,
        "output_dir": str(OUTPUT_DIR),
        "output_name": output_name_for(DEFAULT_REPORT_TYPE, "2026-05-19", report_month),
    }


def build_params(values: dict[str, str]) -> ReportParams:
    """把窗口表单值转为业务参数。"""

    defaults = default_values()
    report_type = values.get("report_type", defaults["report_type"])
    report_month = values.get("report_month", defaults["report_month"])
    end_date = values.get("end_date", defaults["end_date"])
    return ReportParams(
        excel_path=Path(values.get("excel_path", "")).expanduser(),
        template_path=Path(values.get("template_path", "")).expanduser(),
        title=values.get("title", "").strip() or defaults["title"],
        start_date=format_cn_date(values.get("start_date", defaults["start_date"])),
        end_date=format_cn_date(end_date),
        report_month=format_cn_month(report_month),
        output_dir=Path(values.get("output_dir") or OUTPUT_DIR).expanduser(),
        output_name=output_name_for(report_type, end_date, report_month),
        report_type=report_type,
    )


class DatePickerPopup(tk.Toplevel):
    """轻量日期选择弹窗，避免为桌面端额外引入依赖。"""

    def __init__(self, master: tk.Misc, variable: tk.StringVar, value_format: str, title: str) -> None:
        super().__init__(master)
        self.variable = variable
        self.value_format = value_format
        self.selected_date = self._parse_current_value()
        self.display_year = self.selected_date.year
        self.display_month = self.selected_date.month
        self.title(title)
        self.resizable(False, False)
        self.transient(master)
        self._build_ui()
        self.grab_set()
        self.focus_set()

    def _parse_current_value(self) -> date:
        value = self.variable.get().strip()
        for fmt in ("%Y-%m-%d", "%Y-%m"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
        return date.today()

    def _build_ui(self) -> None:
        body = ttk.Frame(self, padding=10)
        body.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(body)
        header.pack(fill=tk.X)
        ttk.Button(header, text="<", width=3, command=lambda: self._change_month(-1)).pack(side=tk.LEFT)
        self.month_label = ttk.Label(header, anchor=tk.CENTER, style="Card.TLabel")
        self.month_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Button(header, text=">", width=3, command=lambda: self._change_month(1)).pack(side=tk.RIGHT)

        weekdays = ("一", "二", "三", "四", "五", "六", "日")
        weekday_row = ttk.Frame(body)
        weekday_row.pack(fill=tk.X, pady=(10, 4))
        for weekday in weekdays:
            ttk.Label(weekday_row, text=weekday, anchor=tk.CENTER, width=4).pack(side=tk.LEFT)

        self.days_frame = ttk.Frame(body)
        self.days_frame.pack()
        self._render_days()

    def _change_month(self, delta: int) -> None:
        month_index = self.display_month - 1 + delta
        self.display_year += month_index // 12
        self.display_month = month_index % 12 + 1
        self._render_days()

    def _render_days(self) -> None:
        for child in self.days_frame.winfo_children():
            child.destroy()

        self.month_label.configure(text=f"{self.display_year}年{self.display_month}月")
        month_days = calendar.Calendar(firstweekday=0).monthdayscalendar(self.display_year, self.display_month)
        for week in month_days:
            row = ttk.Frame(self.days_frame)
            row.pack(fill=tk.X, pady=1)
            for day in week:
                if day == 0:
                    ttk.Label(row, text="", width=4).pack(side=tk.LEFT, padx=1)
                    continue
                button_text = str(day)
                command = lambda selected_day=day: self._select_day(selected_day)
                ttk.Button(row, text=button_text, width=4, command=command).pack(side=tk.LEFT, padx=1)

    def _select_day(self, day: int) -> None:
        selected = date(self.display_year, self.display_month, day)
        self.variable.set(selected.strftime(self.value_format))
        self.destroy()


class MonthlyReportWindow(tk.Tk):
    """可回收物交投点月报生成桌面窗口。"""

    def __init__(self) -> None:
        super().__init__()
        self.title(REPORT_TYPES[DEFAULT_REPORT_TYPE]["window_title"])
        self.geometry("1160x760")
        self.minsize(980, 640)
        self.result: dict | None = None
        self.vars = {key: tk.StringVar(value=value) for key, value in default_values().items()}
        self.vars["excel_path"] = tk.StringVar()
        self.vars["template_path"] = tk.StringVar()
        self.template_label_var = tk.StringVar(value=REPORT_TYPES[DEFAULT_REPORT_TYPE]["template_label"])
        self.nav_buttons: dict[str, ttk.Button] = {}
        self.vars["report_month"].trace_add("write", lambda *_: self._sync_output_name())
        self.vars["end_date"].trace_add("write", lambda *_: self._sync_output_name())
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
        style.configure("Nav.TButton", font=("Microsoft YaHei UI", 10), padding=(10, 7))
        style.configure("Active.Nav.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(10, 7))

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=18)
        root.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(root)
        header.pack(fill=tk.X, pady=(0, 14))
        self.title_label = ttk.Label(header, text=REPORT_TYPES[DEFAULT_REPORT_TYPE]["window_title"], style="Title.TLabel")
        self.title_label.pack(side=tk.LEFT)
        self.status_var = tk.StringVar(value="请选择台账和通报模板后生成。")
        ttk.Label(header, textvariable=self.status_var).pack(side=tk.RIGHT)

        self._build_nav(root)

        content = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        content.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(content, style="Card.TFrame", padding=16)
        right = ttk.Frame(content, style="Card.TFrame", padding=16)
        content.add(left, weight=1)
        content.add(right, weight=2)

        self._build_form(left)
        self._build_result(right)

    def _build_nav(self, parent: ttk.Frame) -> None:
        nav = ttk.Frame(parent)
        nav.pack(fill=tk.X, pady=(0, 14))
        for report_key, config in REPORT_TYPES.items():
            button = ttk.Button(
                nav,
                text=config["nav_text"],
                style="Nav.TButton",
                command=lambda selected=report_key: self._select_report_type(selected),
            )
            button.pack(side=tk.LEFT, padx=(0, 8))
            self.nav_buttons[report_key] = button
        self._refresh_nav()

    def _build_form(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="文件与参数", style="Section.TLabel").pack(anchor=tk.W, pady=(0, 12))
        self._file_row(parent, "台账上传", "excel_path", [("Excel 文件", "*.xls *.xlsx")])
        self._file_row(parent, self.template_label_var, "template_path", [("Word 文档", "*.docx")])
        self._date_row(parent, "报告月份（YYYY-MM）", "report_month", "%Y-%m")
        self._date_row(parent, "检查开始日期（YYYY-MM-DD）", "start_date", "%Y-%m-%d")
        self._date_row(parent, "检查结束日期（YYYY-MM-DD）", "end_date", "%Y-%m-%d")
        self._dir_row(parent, "输出目录", "output_dir")
        self._entry_row(parent, "输出文件名", "output_name")

        buttons = ttk.Frame(parent, style="Card.TFrame")
        buttons.pack(fill=tk.X, pady=(18, 0))
        self.generate_button = ttk.Button(
            buttons,
            text=REPORT_TYPES[DEFAULT_REPORT_TYPE]["nav_text"],
            style="Accent.TButton",
            command=self._generate,
        )
        self.generate_button.pack(side=tk.LEFT)
        ttk.Button(buttons, text="重置", command=self._reset).pack(side=tk.LEFT, padx=(8, 0))
        self.open_file_button = ttk.Button(buttons, text="打开文件", command=self._open_output, state=tk.DISABLED)
        self.open_file_button.pack(side=tk.RIGHT)

    def _select_report_type(self, report_type: str) -> None:
        if report_type not in REPORT_TYPES:
            return
        self.vars["report_type"].set(report_type)
        config = REPORT_TYPES[report_type]
        self.vars["title"].set(config["title"])
        self._sync_output_name()
        default_template = config.get("default_template")
        if default_template and Path(default_template).exists():
            self.vars["template_path"].set(default_template)
        self.template_label_var.set(config["template_label"])
        self.title(config["window_title"])
        self.title_label.configure(text=config["window_title"])
        self.status_var.set(f"已切换到：{config['nav_text']}。")
        self.generate_button.configure(text=config["nav_text"])
        self._refresh_nav()

    def _refresh_nav(self) -> None:
        active = self.vars["report_type"].get()
        for report_key, button in self.nav_buttons.items():
            style = "Active.Nav.TButton" if report_key == active else "Nav.TButton"
            button.configure(style=style)

    def _current_generate_text(self) -> str:
        report_type = self.vars["report_type"].get()
        return REPORT_TYPES.get(report_type, REPORT_TYPES[DEFAULT_REPORT_TYPE])["nav_text"]

    def _sync_output_name(self) -> None:
        report_type = self.vars["report_type"].get()
        report_month = self.vars["report_month"].get()
        end_date = self.vars["end_date"].get()
        self.vars["output_name"].set(output_name_for(report_type, end_date, report_month))

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
            width = 90 if column in {"行号", "图片数"} else 140
            self.preview.column(column, width=width, stretch=True)
        self.preview.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        ttk.Label(parent, text="运行日志", style="Subsection.TLabel").pack(anchor=tk.W)
        self.log_text = tk.Text(parent, height=8, wrap=tk.WORD, relief=tk.SOLID, borderwidth=1)
        self.log_text.pack(fill=tk.BOTH, expand=False, pady=(6, 0))

    def _file_row(self, parent: ttk.Frame, label: str | tk.StringVar, key: str, filetypes) -> None:
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

    def _date_row(self, parent: ttk.Frame, label: str, key: str, value_format: str) -> None:
        row = self._row(parent, label)
        ttk.Entry(row, textvariable=self.vars[key]).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="选择", command=lambda: self._choose_date(key, value_format)).pack(side=tk.LEFT, padx=(8, 0))

    def _row(self, parent: ttk.Frame, label: str | tk.StringVar) -> ttk.Frame:
        if isinstance(label, tk.StringVar):
            ttk.Label(parent, textvariable=label, style="Card.TLabel").pack(anchor=tk.W, pady=(8, 4))
        else:
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

    def _choose_date(self, key: str, value_format: str) -> None:
        title = "选择报告月份" if value_format == "%Y-%m" else "选择日期"
        popup = DatePickerPopup(self, self.vars[key], value_format, title)
        self.update_idletasks()
        x = self.winfo_rootx() + 120
        y = self.winfo_rooty() + 120
        popup.geometry(f"+{x}+{y}")

    def _reset(self) -> None:
        for key, value in default_values().items():
            self.vars[key].set(value)
        self.vars["excel_path"].set("")
        self.vars["template_path"].set("")
        self._select_report_type(DEFAULT_REPORT_TYPE)

    def _generate(self) -> None:
        self._sync_output_name()
        values = {key: var.get() for key, var in self.vars.items()}
        self.generate_button.configure(state=tk.DISABLED, text="正在生成")
        self.open_file_button.configure(state=tk.DISABLED)
        self.status_var.set("正在生成，请稍候...")
        self._set_logs([f"[INFO] 开始{self._current_generate_text()}..."])
        threading.Thread(target=self._run_generate, args=(values,), daemon=True).start()

    def _run_generate(self, values: dict[str, str]) -> None:
        try:
            params = build_params(values)
            process_report = _load_pipeline()
            result = process_report(params).to_dict()
        except Exception as exc:  # noqa: BLE001 - 桌面程序需要展示所有处理错误
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
        self.generate_button.configure(state=tk.NORMAL, text=self._current_generate_text())
        self.open_file_button.configure(state=tk.NORMAL)
        self.status_var.set(f"生成完成：{result.get('output_path')}")
        messagebox.showinfo("生成完成", f"检查通报已生成：\n{result.get('output_path')}")

    def _generation_failed(self, exc: Exception) -> None:
        self.generate_button.configure(state=tk.NORMAL, text=self._current_generate_text())
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
            os.startfile(path)  # noqa: S606 - 用户点击打开生成文件
        else:
            messagebox.showwarning("文件不存在", "输出文件不存在或已被移动。")


def main() -> None:
    app = MonthlyReportWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
