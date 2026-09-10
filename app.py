from __future__ import annotations

import json
import calendar
import os
import sys
import threading
import traceback
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core import ExportCancelled, ImportStore, available_dates, export_task, load_customer_names, load_region_rules, scan_headers


BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
# Keep the working database in the user's local profile. This avoids SQLite WAL
# and file-locking problems when the executable is launched from a shared folder
# such as a Parallels `\\Mac` drive or a read-only installation directory.
if getattr(sys, "frozen", False):
    DATA_DIR = Path(os.environ.get("LOCALAPPDATA", BASE_DIR)) / "ExcelCustomerSplitter"
else:
    DATA_DIR = BASE_DIR
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "splitter.sqlite3"
ERROR_LOG = DATA_DIR / "error.log"
REGION_CONFIG = BASE_DIR / "regions.json"
CUSTOMER_CONFIG = BASE_DIR / "customer_config.json"


class DateCalendar(ttk.LabelFrame):
    """Small dependency-free range calendar that enables only imported dates."""

    def __init__(self, master, on_select):
        super().__init__(master, text="导出日期范围（可选）", padding=6)
        self.on_select = on_select
        self.available: set[date] = set()
        self.min_date: date | None = None
        self.max_date: date | None = None
        self.start: date | None = None
        self.end: date | None = None
        self.month = date.today().replace(day=1)
        header = ttk.Frame(self)
        header.pack(fill="x")
        ttk.Button(header, text="‹", width=3, command=lambda: self.change_month(-1)).pack(side="left")
        self.month_label = ttk.Label(header, anchor="center")
        self.month_label.pack(side="left", fill="x", expand=True)
        ttk.Button(header, text="›", width=3, command=lambda: self.change_month(1)).pack(side="right")
        self.grid_frame = ttk.Frame(self)
        self.grid_frame.pack(fill="x", pady=(4, 2))
        ttk.Button(self, text="清除日期范围（导出全部日期）", command=self.clear).pack(fill="x", pady=(2, 0))
        self.buttons: list[ttk.Button] = []
        self.render()

    def set_dates(self, values: list[str]):
        parsed: set[date] = set()
        for value in values:
            try:
                parsed.add(date.fromisoformat(value))
            except ValueError:
                continue
        self.available = parsed
        self.min_date = min(parsed) if parsed else None
        self.max_date = max(parsed) if parsed else None
        if (not parsed or
                (self.start and self.min_date and self.start < self.min_date) or
                (self.end and self.max_date and self.end > self.max_date)):
            self.start = None
            self.end = None
            self.on_select(None, None)
        if parsed:
            anchor = self.start or min(parsed)
            self.month = anchor.replace(day=1)
        self.render()

    def change_month(self, offset: int):
        index = self.month.year * 12 + self.month.month - 1 + offset
        self.month = date(index // 12, index % 12 + 1, 1)
        self.render()

    def clear(self):
        self.start = None
        self.end = None
        self.on_select(None, None)
        self.render()

    def choose(self, value: date):
        if self.start is None or self.end is not None:
            self.start = value
            self.end = None
        elif value < self.start:
            self.start, self.end = value, self.start
        else:
            self.end = value
        self.on_select(self.start.isoformat() if self.start else None, self.end.isoformat() if self.end else None)
        self.render()

    def render(self):
        for child in self.grid_frame.winfo_children():
            child.destroy()
        self.buttons = []
        self.month_label.configure(text=f"{self.month.year} 年 {self.month.month} 月")
        for column, label in enumerate(("一", "二", "三", "四", "五", "六", "日")):
            ttk.Label(self.grid_frame, text=label, anchor="center", width=4).grid(row=0, column=column, padx=1, pady=1)
        calendar_rows = calendar.monthcalendar(self.month.year, self.month.month)
        for row_index, week in enumerate(calendar_rows, start=1):
            for column, day_number in enumerate(week):
                if not day_number:
                    ttk.Label(self.grid_frame, text="", width=4).grid(row=row_index, column=column, padx=1, pady=1)
                    continue
                value = date(self.month.year, self.month.month, day_number)
                button = ttk.Button(self.grid_frame, text=str(day_number), width=4, command=lambda item=value: self.choose(item))
                if not self.min_date or value < self.min_date or value > self.max_date:
                    button.state(["disabled"])
                if value == self.start or value == self.end:
                    button.configure(text=f"[{day_number}]")
                elif self.start and self.end and self.start < value < self.end:
                    button.configure(text=f"·{day_number}·")
                button.grid(row=row_index, column=column, padx=1, pady=1)
                self.buttons.append(button)


class ScrollableFrame(ttk.Frame):
    """Scrollable settings panel so the date calendar is never clipped."""

    def __init__(self, master, width=360):
        super().__init__(master, width=width)
        self.canvas = tk.Canvas(self, highlightthickness=0, width=width - 24)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.inner.bind("<Configure>", self._update_scrollregion)
        self.canvas.bind("<Configure>", self._fit_inner_width)
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        # Tk may calculate the scrollregion after the first geometry pass and
        # leave a newly created canvas at a non-zero y offset. Always present
        # the top of the settings panel on launch so the date picker is seen.
        self.after_idle(lambda: self.canvas.yview_moveto(0))

    def _update_scrollregion(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _fit_inner_width(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _bind_mousewheel(self, _event=None):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event=None):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")


class SplitterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("客户 Excel 拆分工具")
        # Keep enough vertical room for the date picker on first launch. The
        # settings panel remains scrollable for smaller screens and when the
        # customer/configuration sections are expanded.
        self.geometry("1080x820")
        self.minsize(960, 680)
        self.files: list[str] = []
        self.task_id: str | None = None
        self.headers: list[str] = []
        self.selected_start_date: str | None = None
        self.selected_end_date: str | None = None
        self.customer_names: list[str] = []
        self._cancel_event = threading.Event()
        self._worker_thread: threading.Thread | None = None
        self._closing = False
        self.store = ImportStore(DB_PATH)
        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=14)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text="客户 Excel 拆分工具", font=("Arial", 18, "bold")).pack(anchor="w")
        ttk.Label(root, text="导入多个 Excel，选择字段后按日期 / 分区 / 客户生成独立文件。", foreground="#555").pack(anchor="w", pady=(2, 12))

        file_frame = ttk.LabelFrame(root, text="1. 输入文件")
        file_frame.pack(fill="x", pady=6)
        self.file_label = ttk.Label(file_frame, text="尚未选择文件")
        self.file_label.pack(side="left", padx=8, pady=10)
        ttk.Button(file_frame, text="选择 Excel 文件", command=self.choose_files).pack(side="right", padx=8, pady=7)

        mapping_frame = ttk.LabelFrame(root, text="2. 字段设置")
        mapping_frame.pack(fill="both", expand=True, pady=6)
        left = ttk.Frame(mapping_frame)
        left.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        ttk.Label(left, text="导出字段（Ctrl/Command 可多选）").pack(anchor="w")
        list_frame = ttk.Frame(left)
        list_frame.pack(fill="both", expand=True, pady=5)
        self.field_list = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, exportselection=False)
        self.field_list.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.field_list.yview)
        scrollbar.pack(side="right", fill="y")
        self.field_list.configure(yscrollcommand=scrollbar.set)
        ttk.Button(left, text="扫描表头并导入 SQLite", command=self.scan_and_import).pack(anchor="w", pady=5)

        right_panel = ScrollableFrame(mapping_frame, width=390)
        right_panel.pack(side="right", fill="both", padx=8, pady=8)
        right = right_panel.inner
        right_panel.pack_propagate(False)
        self.customer_var = tk.StringVar()
        self.region_var = tk.StringVar()
        self.year_var = tk.StringVar()
        self.month_var = tk.StringVar()
        self.day_var = tk.StringVar()
        for label, variable, key in (("客户字段（必选）", self.customer_var, "customer"), ("分区字段（可选）", self.region_var, "region")):
            ttk.Label(right, text=label).pack(anchor="w", pady=(10, 2))
            combo = ttk.Combobox(right, textvariable=variable, state="readonly")
            combo.pack(fill="x")
            setattr(self, key + "_combo", combo)

        # Date columns are detected from the standard 年 / 月 / 日 headers.
        # Keep only one date interaction in the UI: selecting the export range
        # on the calendar. The internal variables remain so older workbooks
        # and the export engine continue to use the same date filtering path.
        self.calendar = DateCalendar(right, self.set_selected_date)
        self.calendar.pack(fill="x", pady=(10, 0))
        self.selected_date_label = ttk.Label(right, text="未选择日期范围，将导出全部日期", foreground="#555", wraplength=350)
        self.selected_date_label.pack(anchor="w", pady=(4, 0))

        customer_config_frame = ttk.LabelFrame(right, text="客户名单配置（可选）", padding=5)
        customer_config_frame.pack(fill="x", pady=(12, 0))
        self.customer_config_var = tk.StringVar(value=str(CUSTOMER_CONFIG))
        customer_config_row = ttk.Frame(customer_config_frame)
        customer_config_row.pack(fill="x")
        ttk.Entry(customer_config_row, textvariable=self.customer_config_var).pack(side="left", fill="x", expand=True)
        ttk.Button(customer_config_row, text="...", width=3, command=self.choose_customer_config).pack(side="right", padx=(4, 0))
        ttk.Button(customer_config_frame, text="加载预设客户", command=self.load_customer_config).pack(anchor="w", pady=(4, 2))
        customer_list_frame = ttk.Frame(customer_config_frame)
        customer_list_frame.pack(fill="x")
        self.customer_list = tk.Listbox(customer_list_frame, selectmode=tk.MULTIPLE, height=4, exportselection=False)
        self.customer_list.pack(side="left", fill="x", expand=True)
        customer_scrollbar = ttk.Scrollbar(customer_list_frame, orient="vertical", command=self.customer_list.yview)
        customer_scrollbar.pack(side="right", fill="y")
        self.customer_list.configure(yscrollcommand=customer_scrollbar.set)

        ttk.Label(right, text="提示：程序按表头自动识别 年、月、日；日历支持点击开始和结束日期。滚动右侧面板可查看完整设置。", foreground="#666", wraplength=350).pack(anchor="w", pady=(8, 0))
        ttk.Label(right, text="分区配置").pack(anchor="w", pady=(14, 2))
        self.region_path_var = tk.StringVar(value=str(REGION_CONFIG))
        ttk.Entry(right, textvariable=self.region_path_var).pack(fill="x")
        ttk.Button(right, text="选择配置文件", command=self.choose_region_config).pack(anchor="w", pady=4)
        ttk.Label(right, text="输出目录").pack(anchor="w", pady=(14, 2))
        self.output_var = tk.StringVar(value=str(BASE_DIR / "output"))
        output_row = ttk.Frame(right)
        output_row.pack(fill="x")
        ttk.Entry(output_row, textvariable=self.output_var).pack(side="left", fill="x", expand=True)
        ttk.Button(output_row, text="...", width=4, command=self.choose_output).pack(side="right", padx=(4, 0))

        action_frame = ttk.Frame(root)
        action_frame.pack(fill="x", pady=8)
        self.run_button = ttk.Button(action_frame, text="开始拆分导出", command=self.start_export)
        self.run_button.pack(side="right")
        self.stop_button = ttk.Button(action_frame, text="结束任务并释放资源", command=self.stop_and_cleanup, state="disabled")
        self.stop_button.pack(side="right", padx=(8, 0))
        self.progress = ttk.Progressbar(action_frame, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 10))
        cache_frame = ttk.Frame(root)
        cache_frame.pack(fill="x", pady=(0, 6))
        self.cache_info_var = tk.StringVar()
        ttk.Label(cache_frame, textvariable=self.cache_info_var, foreground="#666").pack(side="left", fill="x", expand=True)
        ttk.Button(cache_frame, text="清除本地数据库缓存", command=self.clear_database_cache).pack(side="right")
        self.log = tk.Text(root, height=8, state="disabled", background="#f7f7f7")
        self.log.pack(fill="both", expand=False)
        self.refresh_cache_info()

    def log_line(self, message: str):
        self.after(0, lambda: (self.log.configure(state="normal"), self.log.insert("end", message + "\n"), self.log.see("end"), self.log.configure(state="disabled")))

    def choose_files(self):
        files = filedialog.askopenfilenames(title="选择 Excel 文件", filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")])
        if files:
            self.files = list(files)
            self.file_label.configure(text=f"已选择 {len(self.files)} 个文件")
            self.log_line("已选择: " + ", ".join(Path(item).name for item in self.files))

    def choose_region_config(self):
        path = filedialog.askopenfilename(title="选择分区配置", filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")])
        if path:
            self.region_path_var.set(path)

    def choose_customer_config(self):
        path = filedialog.askopenfilename(title="选择客户名单配置", filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")])
        if path:
            self.customer_config_var.set(path)
            self.load_customer_config()

    def load_customer_config(self):
        try:
            self.customer_names = load_customer_names(self.customer_config_var.get())
        except Exception as exc:
            messagebox.showerror("客户配置错误", f"无法读取客户名单：{exc}")
            return
        self.customer_list.delete(0, tk.END)
        for name in self.customer_names:
            self.customer_list.insert(tk.END, name)
        if self.customer_names:
            self.customer_list.selection_set(0, tk.END)
            self.log_line(f"已加载 {len(self.customer_names)} 个预设客户。")
        else:
            self.log_line("客户名单配置为空或文件不存在，将导出全部客户。")

    def choose_output(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_var.set(path)

    def scan_and_import(self):
        if not self.files:
            messagebox.showwarning("缺少文件", "请先选择 Excel 文件。")
            return
        self.run_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.progress.start(10)
        self._cancel_event.clear()
        self._worker_thread = threading.Thread(target=self._import_worker, daemon=True)
        self._worker_thread.start()

    def _import_worker(self):
        try:
            self.log_line("开始扫描并导入 SQLite...")
            task_id, headers = self.store.import_files(self.files, self.log_line, self._cancel_event)
            self.task_id, self.headers = task_id, headers
            self.after(0, self.populate_fields)
            self.log_line(f"导入完成，共识别 {len(headers)} 个字段。")
        except ExportCancelled as exc:
            self.log_line(str(exc) + "，临时导入数据已清理。")
        except Exception as exc:
            ERROR_LOG.write_text(traceback.format_exc(), encoding="utf-8")
            error_message = str(exc)
            self.log_line(f"导入失败: {error_message}")
            self.after(0, lambda message=error_message: messagebox.showerror("导入失败", f"{message}\n\n详细日志：{ERROR_LOG}"))
        finally:
            self.after(0, lambda: (self.progress.stop(), self.run_button.configure(state="normal"), self.stop_button.configure(state="disabled")))

    def set_selected_date(self, start: str | None, end: str | None):
        self.selected_start_date = start
        self.selected_end_date = end
        if start and end:
            self.selected_date_label.configure(text=f"已选择导出范围：{start} 至 {end}")
        elif start:
            self.selected_date_label.configure(text=f"已选择开始日期：{start}，请再选择结束日期")
        else:
            self.selected_date_label.configure(text="未选择日期范围，将导出全部日期")

    def date_mapping_changed(self, _event=None):
        self.selected_start_date = None
        self.selected_end_date = None
        self.calendar.clear()
        self.refresh_available_dates()

    def refresh_available_dates(self):
        if not self.task_id:
            return
        fields = (self.year_var.get().strip(), self.month_var.get().strip(), self.day_var.get().strip())
        if not all(fields):
            self.calendar.set_dates([])
            return
        task_id = self.task_id

        def worker():
            try:
                values = available_dates(self.store, task_id, *fields)
                self.after(0, lambda dates=values: self.calendar.set_dates(dates))
                self.log_line(f"可选日期：{len(values)} 个。")
            except Exception as exc:
                error_message = str(exc)
                self.log_line(f"读取可选日期失败: {error_message}")

        threading.Thread(target=worker, daemon=True).start()

    def populate_fields(self):
        self.field_list.delete(0, tk.END)
        for header in self.headers:
            self.field_list.insert(tk.END, header)
        values = [""] + self.headers
        self.customer_combo["values"] = values
        self.region_combo["values"] = values
        for index in range(len(self.headers)):
            self.field_list.selection_set(index)
        self.set_default_field(self.customer_var, self.customer_combo, ("客户名称", "客户"))
        self.set_default_field(self.region_var, self.region_combo, ("代表处", "区域", "分区"))
        self.set_default_field(self.year_var, None, ("年",))
        self.set_default_field(self.month_var, None, ("月",))
        self.set_default_field(self.day_var, None, ("日",))
        if Path(self.customer_config_var.get()).exists():
            self.load_customer_config()
        self.refresh_available_dates()

    def set_default_field(self, variable, combo, candidates):
        normalized = {item.replace(" ", ""): item for item in self.headers}
        for candidate in candidates:
            value = normalized.get(candidate.replace(" ", ""))
            if value:
                variable.set(value)
                if combo is not None:
                    combo.set(value)
                return
        variable.set("")
        if combo is not None:
            combo.set("")

    def start_export(self):
        if not self.task_id:
            messagebox.showwarning("尚未导入", "请先扫描表头并导入 SQLite。")
            return
        selected = [self.headers[index] for index in self.field_list.curselection()]
        customer = self.customer_var.get().strip()
        if not selected or not customer:
            messagebox.showwarning("字段不完整", "至少选择一个导出字段，并选择客户字段。")
            return
        if self.selected_start_date and not self.selected_end_date:
            messagebox.showwarning("日期范围不完整", "请选择结束日期；或者点击清除日期筛选后导出全部日期。")
            return
        date_fields = (self.year_var.get().strip(), self.month_var.get().strip(), self.day_var.get().strip())
        if any(date_fields) and not all(date_fields):
            messagebox.showwarning("日期字段不完整", "如果使用日期分文件，请同时选择年、月、日三个字段。")
            return
        self.run_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.progress.start(10)
        selected_customers = [self.customer_names[index] for index in self.customer_list.curselection()]
        if self.customer_names and not selected_customers:
            messagebox.showwarning("客户未选择", "请至少选择一个预设客户，或清空客户配置后导出全部客户。")
            self.progress.stop()
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            return
        # Read all Tk variables on the main thread before starting the worker.
        # Tkinter StringVar.get() is not safe to call from a background thread.
        task_id = self.task_id
        output_path = self.output_var.get().strip()
        region_path = self.region_path_var.get().strip()
        region_value = self.region_var.get().strip() or None
        self._worker_thread = threading.Thread(
            target=self._export_worker,
            args=(task_id, output_path, region_path, selected, customer, region_value, *date_fields, self.selected_start_date, self.selected_end_date, selected_customers, self._cancel_event),
            daemon=True,
        )
        self._worker_thread.start()

    def _export_worker(self, task_id, output_path, region_path, selected, customer, region, year_field, month_field, day_field, start_date, end_date, selected_customers, cancel_event):
        try:
            rules = load_region_rules(region_path)
            result = export_task(
                self.store,
                task_id,
                output_path,
                selected,
                customer,
                region,
                year_field or None,
                month_field or None,
                day_field or None,
                rules,
                start_date,
                self.log_line,
                end_date,
                selected_customers,
                cancel_event,
            )
            self.log_line(f"导出完成：{result['rows']} 行，{result['files']} 个文件。")
            self.after(0, lambda: messagebox.showinfo("完成", f"已生成 {result['files']} 个客户文件。"))
        except ExportCancelled as exc:
            self.log_line(str(exc) + "，临时表和工作簿已清理。")
        except Exception as exc:
            ERROR_LOG.write_text(traceback.format_exc(), encoding="utf-8")
            error_message = str(exc)
            self.log_line(f"导出失败: {error_message}")
            self.after(0, lambda message=error_message: messagebox.showerror("导出失败", f"{message}\n\n详细日志：{ERROR_LOG}"))
        finally:
            self.after(0, lambda: (self.progress.stop(), self.run_button.configure(state="normal"), self.stop_button.configure(state="disabled")))

    def stop_and_cleanup(self):
        """Request cooperative cancellation, then reopen SQLite to release handles."""
        worker = self._worker_thread
        if worker and worker.is_alive():
            self._cancel_event.set()
            self.stop_button.configure(state="disabled")
            self.log_line("已请求结束任务，正在等待当前批次收尾并释放数据库...")
            self.after(100, self._finish_cleanup_when_idle)
            return
        self._reopen_store()

    def _finish_cleanup_when_idle(self):
        worker = self._worker_thread
        if worker and worker.is_alive():
            self.after(100, self._finish_cleanup_when_idle)
            return
        self._reopen_store()

    def _reopen_store(self):
        try:
            self.store.close()
        except Exception:
            pass
        self.store = ImportStore(DB_PATH)
        self.log_line("后台任务已结束，SQLite 连接已关闭并重新建立，文件占用已释放。")
        self.refresh_cache_info()

    def _cache_paths(self) -> tuple[Path, ...]:
        return (DB_PATH, Path(f"{DB_PATH}-wal"), Path(f"{DB_PATH}-shm"))

    def refresh_cache_info(self):
        total_size = sum(path.stat().st_size for path in self._cache_paths() if path.exists())
        size_mb = total_size / (1024 * 1024)
        self.cache_info_var.set(f"本地数据库缓存：{size_mb:.1f} MB（{DATA_DIR}）")

    def clear_database_cache(self):
        worker = self._worker_thread
        if worker and worker.is_alive():
            messagebox.showwarning("任务正在运行", "请先点击“结束任务并释放资源”，等待后台任务结束后再清除数据库缓存。")
            return
        if not messagebox.askyesno(
            "确认清除缓存",
            f"将删除本地数据库缓存及 SQLite 临时文件：\n{DB_PATH}\n\n这不会删除 Excel、导出结果或配置文件。是否继续？",
        ):
            return
        try:
            self.store.close()
            deleted = 0
            for path in self._cache_paths():
                try:
                    path.unlink()
                    deleted += 1
                except FileNotFoundError:
                    pass
            self.store = ImportStore(DB_PATH)
            self.task_id = None
            self.headers = []
            self.field_list.delete(0, tk.END)
            self.customer_combo["values"] = []
            self.region_combo["values"] = []
            for variable in (self.customer_var, self.region_var, self.year_var, self.month_var, self.day_var):
                variable.set("")
            self.calendar.clear()
            self.log_line(f"已清除本地数据库缓存，删除 {deleted} 个文件；请重新选择 Excel 并导入。")
            self.refresh_cache_info()
        except Exception as exc:
            ERROR_LOG.write_text(traceback.format_exc(), encoding="utf-8")
            try:
                self.store = ImportStore(DB_PATH)
            except Exception:
                pass
            messagebox.showerror("清除缓存失败", f"{exc}\n\n详细日志：{ERROR_LOG}")

    def destroy(self):
        if self._closing:
            return
        worker = self._worker_thread
        if worker and worker.is_alive():
            self._closing = True
            self._cancel_event.set()
            self.after(100, self._destroy_when_idle)
            return
        self.store.close()
        super().destroy()

    def _destroy_when_idle(self):
        worker = self._worker_thread
        if worker and worker.is_alive():
            self.after(100, self._destroy_when_idle)
            return
        self.store.close()
        super().destroy()


if __name__ == "__main__":
    SplitterApp().mainloop()
