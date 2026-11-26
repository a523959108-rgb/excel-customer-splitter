from __future__ import annotations

import json
import calendar
import os
import sys
import tempfile
from queue import Empty, SimpleQueue
import threading
import traceback
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core import ExportCancelled, load_customer_names, load_region_rules
from warehouse import Warehouse


BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("LOCALAPPDATA") or Path.home() / ".local" / "share") / "ExcelCustomerSplitter"
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
    def __init__(self, data_dir=DATA_DIR):
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
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.settings = self.data_dir / "settings.json"
        self.store = None
        self._callbacks = SimpleQueue()
        self._build_ui()
        self.after(30, self._drain_callbacks)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        try:
            if self.settings.exists():
                path = Path(json.loads(self.settings.read_text(encoding="utf-8"))["database"])
                self.open_store(path, existing=True)
            else:
                self.open_store(self.data_dir / "splitter.duckdb")
        except Exception as exc:
            self.log_line(f"无法打开数据库：{exc}。请点击“打开数据库”或“新建数据库”。")
        if any((folder / "splitter.sqlite3").is_file() for folder in (self.data_dir, BASE_DIR)):
            self.log_line("检测到旧 SQLite 数据，可在空数据库中点击“迁移旧 SQLite”。")

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
        self.import_button = ttk.Button(left, text="扫描表头并导入数据库", command=self.scan_and_import)
        self.import_button.pack(anchor="w", pady=5)
        ttk.Label(left, text="同一天的完整明细请一次选齐。再次导入该日会覆盖整天。", wraplength=410).pack(anchor="w")

        right_panel = ScrollableFrame(mapping_frame, width=390)
        right_panel.pack(side="right", fill="both", padx=8, pady=8)
        right = right_panel.inner
        right_panel.pack_propagate(False)
        self.customer_var = tk.StringVar()
        self.region_var = tk.StringVar()
        for label, variable, key in (("客户字段（必选）", self.customer_var, "customer"), ("分区字段（可选）", self.region_var, "region")):
            ttk.Label(right, text=label).pack(anchor="w", pady=(10, 2))
            combo = ttk.Combobox(right, textvariable=variable, state="readonly")
            combo.pack(fill="x")
            setattr(self, key + "_combo", combo)

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
        self.database_buttons = []
        for label, command in (("新建数据库", self.create_database), ("打开数据库", self.choose_database),
                               ("迁移旧 SQLite", self.migrate_database), ("整理事务日志", self.clear_database_cache)):
            button = ttk.Button(cache_frame, text=label, command=command)
            button.pack(side="left", padx=2)
            self.database_buttons.append(button)
        self.log = tk.Text(root, height=8, state="disabled", background="#f7f7f7")
        self.log.pack(fill="both", expand=False)
        self.refresh_cache_info()

    def log_line(self, message: str):
        def append():
            self.log.configure(state="normal")
            self.log.insert("end", message + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self._callbacks.put(append)

    def _drain_callbacks(self):
        try:
            for _ in range(100):
                self._callbacks.get_nowait()()
        except Empty:
            pass
        self.after(30, self._drain_callbacks)

    def open_store(self, path, existing=False):
        if self._worker_thread is not None:
            return
        path = Path(path).resolve()
        if existing and not path.is_file():
            raise FileNotFoundError(path)
        store = Warehouse(path)
        try:
            task, headers = store.restore()
            with tempfile.NamedTemporaryFile(dir=self.data_dir, delete=False) as fp:
                tmp = Path(fp.name)
            try:
                tmp.write_text(json.dumps({"database": str(path)}, ensure_ascii=False), encoding="utf-8")
                tmp.replace(self.settings)
            finally:
                tmp.unlink(missing_ok=True)
        except Exception:
            store.close()
            raise
        if self.store:
            self.store.close()
        self.store = store
        self.task_id, self.headers = task, headers
        self.calendar.clear()
        self.populate_fields()
        self.refresh_cache_info()
        self.log_line(f"已打开数据库：{path}")

    def choose_database(self):
        path = filedialog.askopenfilename(title="打开客户数据库", filetypes=[("DuckDB", "*.duckdb"), ("所有文件", "*.*")])
        if path:
            self._select_database(path, existing=True)

    def create_database(self):
        path = filedialog.asksaveasfilename(title="新建客户数据库", defaultextension=".duckdb", filetypes=[("DuckDB", "*.duckdb")])
        if path:
            if Path(path).exists():
                messagebox.showwarning("文件已存在", "请选择新文件名，或用“打开数据库”加载已有数据库。")
                return
            self._select_database(path, existing=False)

    def _select_database(self, path, existing):
        try:
            self.open_store(path, existing)
        except Exception as exc:
            messagebox.showerror("打开数据库失败", str(exc))

    def migrate_database(self):
        if not self.store or self._worker_thread is not None:
            return
        path = filedialog.askopenfilename(title="迁移旧 SQLite 到当前空库", filetypes=[("SQLite", "*.sqlite3 *.sqlite *.db"), ("所有文件", "*.*")])
        if path:
            self._start_job("迁移", lambda: self.store.migrate_sqlite(path, self.log_line, self._cancel_event), self._show_import)

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
        if not self.store:
            messagebox.showwarning("请选择数据库", "请先打开或新建数据库。")
            return
        paths = tuple(self.files)
        self._start_job("导入", lambda: self.store.import_files(paths, self.log_line, self._cancel_event), self._show_import)

    def _show_import(self, result):
        self.task_id, self.headers = result
        self.populate_fields()

    def _start_job(self, name, operation, on_success):
        if self._worker_thread is not None or self._closing:
            return
        self._cancel_event.clear()
        self.stop_button.configure(state="normal")
        for button in [self.run_button, self.import_button, *self.database_buttons]:
            button.configure(state="disabled")
        self.progress.start(10)
        self.log_line(f"开始{name}…")

        def worker():
            try:
                result = operation()
                self._callbacks.put(lambda: None if self._closing else on_success(result))
            except ExportCancelled:
                self.log_line(f"{name}已停止。")
            except Exception as exc:
                error = str(exc)
                self.log_line(f"{name}失败：{error}")
                self._callbacks.put(lambda: None if self._closing else messagebox.showerror(f"{name}失败", error))
                (self.data_dir / "error.log").write_text(traceback.format_exc(), encoding="utf-8")
            finally:
                self._callbacks.put(self._finish_job)
        self._worker_thread = threading.Thread(target=worker, daemon=True)
        self._worker_thread.start()

    def _finish_job(self):
        self._worker_thread = None
        self.progress.stop()
        self.stop_button.configure(state="disabled")
        for button in [self.run_button, self.import_button, *self.database_buttons]:
            button.configure(state="normal")
        self.refresh_cache_info()

    def set_selected_date(self, start: str | None, end: str | None):
        self.selected_start_date = start
        self.selected_end_date = end
        if start and end:
            self.selected_date_label.configure(text=f"已选择导出范围：{start} 至 {end}")
        elif start:
            self.selected_date_label.configure(text=f"已选择开始日期：{start}，请再选择结束日期")
        else:
            self.selected_date_label.configure(text="未选择日期范围，将导出全部日期")

    def refresh_available_dates(self):
        self.calendar.set_dates(self.store.dates() if self.task_id else [])

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
            messagebox.showwarning("尚未导入", "请先导入 Excel 或打开已有数据库。")
            return
        selected = [self.headers[index] for index in self.field_list.curselection()]
        customer = self.customer_var.get().strip()
        if not selected or not customer:
            messagebox.showwarning("字段不完整", "至少选择一个导出字段，并选择客户字段。")
            return
        if self.selected_start_date and not self.selected_end_date:
            messagebox.showwarning("日期范围不完整", "请选择结束日期；或者点击清除日期筛选后导出全部日期。")
            return
        selected_customers = [self.customer_names[index] for index in self.customer_list.curselection()]
        if self.customer_names and not selected_customers:
            messagebox.showwarning("客户未选择", "请至少选择一个预设客户，或清空客户配置后导出全部客户。")
            return
        output = self.output_var.get().strip()
        region_path = self.region_path_var.get().strip()
        region = self.region_var.get().strip() or None
        start, end = self.selected_start_date, self.selected_end_date

        def export():
            return self.store.export(output, selected, customer, region, load_region_rules(region_path),
                                     start, end, selected_customers, self.log_line, self._cancel_event)

        def complete(result):
            message = f"共导出 {result['rows']:,} 行，{result['files']:,} 个文件。"
            self.log_line(message)
            self.log_line(f"本次文件清单：{result['manifest']}")
            messagebox.showinfo("导出完成", f"{message}\n本次文件清单：{result['manifest']}")
        self._start_job("导出", export, complete)

    def stop_and_cleanup(self):
        if self._worker_thread is not None:
            self._cancel_event.set()
            self.store.interrupt()
            self.stop_button.configure(state="disabled")
            self.log_line("已请求停止，正在释放任务资源…")

    def refresh_cache_info(self):
        if self.store:
            path = self.store.path
            size = sum(p.stat().st_size for p in (path, Path(f"{path}.wal")) if p.is_file())
            self.cache_info_var.set(f"数据库 {size / (1024**2):,.1f} MiB  {path}")
        else:
            self.cache_info_var.set("请选择数据库")

    def clear_database_cache(self):
        if self.store and self._worker_thread is None:
            self._start_job("整理事务日志", self.store.checkpoint,
                            lambda _: self.log_line("已将事务日志写入持久库，临时文件随任务结束清理。"))

    def destroy(self):
        self._closing = True
        if self._worker_thread is not None and self._worker_thread.is_alive():
            self.stop_and_cleanup()
            self.after(100, self.destroy)
            return
        if self.store:
            self.store.close()
        super().destroy()


if __name__ == "__main__":
    if sys.argv[1:] == ["--self-test"]:
        from smoke_test import run
        run(SplitterApp)
    else:
        SplitterApp().mainloop()
