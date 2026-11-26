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
                button = ttk.Button(self.grid_frame, text=str(day_number), style="Calendar.TButton", width=4, command=lambda item=value: self.choose(item))
                if not self.min_date or value < self.min_date or value > self.max_date:
                    button.state(["disabled"])
                if value == self.start or value == self.end:
                    button.configure(style="Chosen.Calendar.TButton")
                elif self.start and self.end and self.start < value < self.end:
                    button.configure(style="Range.Calendar.TButton")
                button.grid(row=row_index, column=column, padx=1, pady=1)
                self.buttons.append(button)


class SplitterApp(tk.Tk):
    def __init__(self, data_dir=DATA_DIR):
        super().__init__()
        self.title("客户 Excel 拆分工具")
        width = min(1120, self.winfo_screenwidth() - 32)
        height = min(720, self.winfo_screenheight() - 80)
        self.geometry(f"{width}x{height}")
        self.minsize(min(width, 850), min(height, 520))
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
            self.toggle_log()
            self.log_line(f"无法打开数据库：{exc}。请点击“打开库”或“新建库”。")
        if any((folder / "splitter.sqlite3").is_file() for folder in (self.data_dir, BASE_DIR)):
            self.log_line("检测到旧 SQLite 数据，可在空数据库的“配置与维护”中点击“迁移旧 SQLite”。")

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        ink, muted, edge = "#0F172A", "#475569", "#CBD5E1"
        accent, base = "#2563EB", "#F4F6F9"
        self.configure(background=base)
        style.configure(".", background="white", foreground=ink, bordercolor=edge, focuscolor=accent)
        style.configure("TFrame", background="white")
        style.configure("TLabel", background="white")
        style.configure("Title.TLabel", font=("TkDefaultFont", 18, "bold"))
        style.configure("Section.TLabel", font=("TkDefaultFont", 12, "bold"))
        style.configure("Muted.TLabel", foreground=muted)
        style.configure("Hint.TLabel", foreground="#92400E", background="#FFFBEB", padding=(9, 4))
        style.configure("TButton", background="#F8FAFC", borderwidth=1, padding=(9, 5))
        style.map("TButton", background=[("active", "#E0ECFF")])
        style.map("TCombobox", fieldbackground=[("readonly", "white")])
        style.configure("Accent.TButton", background=accent, foreground="white", padding=(18, 8),
                        font=("TkDefaultFont", 10, "bold"))
        style.map("Accent.TButton", background=[("disabled", "#DCE3EC"), ("active", "#1D4ED8")],
                  foreground=[("disabled", muted), ("!disabled", "white")])
        style.configure("TNotebook", borderwidth=0, tabmargins=0)
        style.configure("TNotebook.Tab", padding=(12, 5))
        style.map("TNotebook.Tab", background=[("selected", "#E0ECFF")],
                  foreground=[("selected", "#1747A6")])
        style.configure("Calendar.TButton", padding=(4, 3))
        style.configure("Chosen.Calendar.TButton", foreground="white", background=accent)
        style.configure("Range.Calendar.TButton", foreground="#1747A6", background="#E0ECFF")
        style.configure("Horizontal.TProgressbar", background=accent, troughcolor="#E2E8F0")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        self.database_buttons = []
        self.cache_info_var = tk.StringVar()
        self.status_var = tk.StringVar(value="就绪")
        self.output_var = tk.StringVar(value=str(BASE_DIR / "output"))
        self.customer_config_var = tk.StringVar(value=str(CUSTOMER_CONFIG))
        self.region_path_var = tk.StringVar(value=str(REGION_CONFIG))
        self.customer_var, self.region_var = tk.StringVar(), tk.StringVar()

        header = ttk.Frame(self, padding=(18, 8))
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="客户拆分", style="Title.TLabel").pack(side="left", padx=(0, 18))
        ttk.Label(header, text="当前库", style="Muted.TLabel").pack(side="left", padx=(0, 6))
        ttk.Entry(header, textvariable=self.cache_info_var, state="readonly", width=16).pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        for label, command in (("打开库", self.choose_database), ("新建库", self.create_database)):
            button = ttk.Button(header, text=label, command=command)
            button.pack(side="left", padx=2)
            self.database_buttons.append(button)

        intake = ttk.Frame(self, padding=(14, 5))
        intake.grid(row=1, column=0, sticky="ew", padx=12, pady=(8, 0))
        self.import_toggle = ttk.Button(intake, command=self.toggle_import)
        self.import_toggle.pack(anchor="w")
        self.import_panel = ttk.Frame(intake)
        ttk.Button(self.import_panel, text="选择 Excel 文件", command=self.choose_files).pack(side="left")
        self.file_label = ttk.Label(self.import_panel, text="尚未选择文件", style="Muted.TLabel")
        self.file_label.pack(side="left", padx=8)
        ttk.Label(self.import_panel, text="同日文件须一次选齐，再次导入会覆盖整日明细。",
                  style="Hint.TLabel").pack(side="left")
        self.import_button = ttk.Button(self.import_panel, text="开始导入", command=self.scan_and_import)
        self.import_button.pack(side="right")
        self.set_import_open(True)

        body = ttk.Frame(self)
        body.grid(row=2, column=0, sticky="nsew", padx=12, pady=8)
        body.columnconfigure(0, weight=1, uniform="panel")
        body.columnconfigure(2, weight=2, uniform="panel")
        body.rowconfigure(0, weight=1)
        customers = ttk.Frame(body, padding=(14, 10))
        customers.grid(row=0, column=0, sticky="nsew")
        ttk.Separator(body, orient="vertical").grid(row=0, column=1, sticky="ns", pady=10)
        ttk.Label(customers, text="客户名单", style="Section.TLabel").pack(anchor="w")
        self.customer_count = tk.StringVar(value="全部客户")
        ttk.Label(customers, textvariable=self.customer_count, style="Muted.TLabel").pack(anchor="w", pady=(2, 6))
        self.customer_list = self.make_list(customers)
        self.customer_list.bind("<<ListboxSelect>>", self.update_counts)
        self.empty_customers = ttk.Label(self.customer_list, text="未加载名单，将导出全部客户", style="Muted.TLabel", wraplength=180)
        ttk.Button(customers, text="选择客户名单…", command=self.choose_customer_config).pack(anchor="w", pady=(7, 0))

        right = ttk.Frame(body, padding=(14, 8))
        right.grid(row=0, column=2, sticky="nsew")
        tabs = ttk.Notebook(right)
        tabs.pack(fill="both", expand=True)
        rules = ttk.Frame(tabs, padding=(10, 7))
        settings = ttk.Frame(tabs, padding=(10, 7))
        tabs.add(rules, text="拆分设置")
        tabs.add(settings, text="配置与维护")
        rules.columnconfigure((0, 1), weight=1, uniform="fields")
        rules.rowconfigure(4, weight=1)
        for col, (label, var, key) in enumerate((("客户字段", self.customer_var, "customer"),
                                                ("分区字段（可选）", self.region_var, "region"))):
            ttk.Label(rules, text=label).grid(row=0, column=col, sticky="w")
            combo = ttk.Combobox(rules, textvariable=var, state="readonly", width=12)
            combo.grid(row=1, column=col, sticky="ew", padx=(0, 7), pady=(2, 5))
            setattr(self, key + "_combo", combo)

        dates = ttk.Frame(rules)
        dates.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        ttk.Button(dates, text="选择日期…", command=self.show_calendar).pack(side="left")
        self.selected_date_label = ttk.Label(dates, text="全部日期", style="Muted.TLabel")
        self.selected_date_label.pack(side="left", padx=8)
        self.calendar_window = tk.Toplevel(self)
        self.calendar_window.withdraw()
        self.calendar_window.title("选择导出日期")
        self.calendar_window.transient(self)
        self.calendar_window.protocol("WM_DELETE_WINDOW", self.calendar_window.withdraw)
        self.calendar_window.bind("<Escape>", lambda _: self.calendar_window.withdraw())
        self.calendar = DateCalendar(self.calendar_window, self.set_selected_date)
        self.calendar.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Button(self.calendar_window, text="完成", command=self.calendar_window.withdraw).pack(
            anchor="e", padx=8, pady=(0, 8))

        self.field_count = tk.StringVar(value="导出字段")
        ttk.Label(rules, textvariable=self.field_count).grid(row=3, column=0, columnspan=2, sticky="w")
        block = ttk.Frame(rules)
        block.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=(3, 0))
        self.field_list = self.make_list(block)
        self.field_list.bind("<<ListboxSelect>>", self.update_counts)
        for label, var, browse in (("客户名单配置", self.customer_config_var, self.choose_customer_config),
                                   ("分区配置", self.region_path_var, self.choose_region_config)):
            ttk.Label(settings, text=label).pack(anchor="w")
            self.path_row(settings, var, browse)
        controls = ttk.Frame(settings)
        controls.pack(anchor="w", pady=(5, 0))
        ttk.Button(controls, text="加载预设客户", command=self.load_customer_config).pack(side="left")
        for label, command in (("迁移旧 SQLite", self.migrate_database), ("整理事务日志", self.clear_database_cache)):
            button = ttk.Button(controls, text=label, command=command)
            button.pack(side="left", padx=(6, 0))
            self.database_buttons.append(button)

        bottom = ttk.Frame(self, padding=(16, 8))
        bottom.grid(row=4, column=0, sticky="ew")
        destination = ttk.Frame(bottom)
        destination.pack(fill="x")
        ttk.Label(destination, text="输出目录").pack(side="left", padx=(0, 8))
        self.path_row(destination, self.output_var, self.choose_output)
        actions = ttk.Frame(bottom)
        actions.pack(fill="x", pady=(4, 0))
        self.log_toggle = ttk.Button(actions, text="显示日志", command=self.toggle_log)
        self.log_toggle.pack(side="left")
        ttk.Label(actions, textvariable=self.status_var, style="Muted.TLabel").pack(side="left", padx=8)
        self.progress = ttk.Progressbar(actions, length=82)
        self.progress.pack(side="left")
        self.run_button = ttk.Button(actions, text="开始拆分导出", style="Accent.TButton", command=self.start_export)
        self.run_button.pack(side="right")
        self.stop_button = ttk.Button(actions, text="停止任务", command=self.stop_and_cleanup, state="disabled")
        self.stop_button.pack(side="right", padx=8)
        self.log_panel = ttk.Frame(self, padding=(16, 4))
        self.log = tk.Text(self.log_panel, height=3, state="disabled", background="#F8FAFC",
                           foreground=ink, borderwidth=0, wrap="word", highlightthickness=0)
        scroll = ttk.Scrollbar(self.log_panel, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)
        self.refresh_cache_info()

    def make_list(self, parent):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        box = tk.Listbox(frame, selectmode=tk.MULTIPLE, exportselection=False, height=2, width=12,
                         relief="flat", borderwidth=0, highlightthickness=1, highlightbackground="#DCE3EC",
                         highlightcolor="#2563EB", background="white", foreground="#0F172A",
                         selectbackground="#E0ECFF", selectforeground="#173D7F", activestyle="dotbox")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=box.yview)
        scrollbar.pack(side="right", fill="y")
        box.configure(yscrollcommand=scrollbar.set)
        box.pack(side="left", fill="both", expand=True)
        return box

    def path_row(self, parent, variable, browse):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(2, 4))
        ttk.Entry(row, textvariable=variable, width=12).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="浏览…", command=browse).pack(side="left", padx=(6, 0))

    def set_import_open(self, visible):
        if visible:
            self.import_panel.pack(fill="x", pady=(6, 0))
        else:
            self.import_panel.pack_forget()
        self.import_toggle.configure(text="收起 Excel 导入" if visible else "展开 Excel 导入")

    def toggle_import(self):
        self.set_import_open(not self.import_panel.winfo_manager())

    def toggle_log(self):
        if self.log_panel.winfo_manager():
            self.log_panel.grid_remove()
            self.log_toggle.configure(text="显示日志")
        else:
            self.log_panel.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 6))
            self.log_toggle.configure(text="收起日志")

    def show_calendar(self):
        self.calendar_window.deiconify()
        self.calendar_window.lift()
        self.calendar_window.focus_set()

    def update_counts(self, _event=None):
        self.field_count.set(f"导出字段  {len(self.field_list.curselection())} / {len(self.headers)}")
        if self.customer_names:
            self.empty_customers.place_forget()
            self.customer_count.set(f"已选 {len(self.customer_list.curselection())} / {len(self.customer_names)}")
        else:
            self.empty_customers.place(relx=0.5, rely=0.5, anchor="center")
            self.customer_count.set("全部客户")

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
        self.set_import_open(not bool(headers))
        self.log_line(f"已打开数据库：{path}")

    def choose_database(self):
        path = filedialog.askopenfilename(title="打开客户数据库", filetypes=[("DuckDB", "*.duckdb"), ("所有文件", "*.*")])
        if path:
            self._select_database(path, existing=True)

    def create_database(self):
        path = filedialog.asksaveasfilename(title="新建客户数据库", defaultextension=".duckdb", filetypes=[("DuckDB", "*.duckdb")])
        if path:
            if Path(path).exists():
                messagebox.showwarning("文件已存在", "请选择新文件名，或用“打开库”加载已有数据库。")
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

        self.update_counts()

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
        self.set_import_open(not bool(self.headers))

    def _start_job(self, name, operation, on_success):
        if self._worker_thread is not None or self._closing:
            return
        self._cancel_event.clear()
        self.stop_button.configure(state="normal")
        for button in [self.run_button, self.import_button, *self.database_buttons]:
            button.configure(state="disabled")
        self.status_var.set(f"正在{name}")
        self.progress.configure(mode="indeterminate")
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
        self.progress.configure(mode="determinate", value=0)
        self.status_var.set("就绪")
        self.stop_button.configure(state="disabled")
        for button in [self.run_button, self.import_button, *self.database_buttons]:
            button.configure(state="normal")
        self.refresh_cache_info()

    def set_selected_date(self, start: str | None, end: str | None):
        self.selected_start_date = start
        self.selected_end_date = end
        if start and end:
            self.selected_date_label.configure(text=f"{start} 至 {end}")
        elif start:
            self.selected_date_label.configure(text=f"从 {start} 开始，待选结束日期")
        else:
            self.selected_date_label.configure(text="全部日期")

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
        self.update_counts()

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
            self.status_var.set("正在停止")
            self.stop_button.configure(state="disabled")
            self.log_line("已请求停止，正在释放任务资源…")

    def refresh_cache_info(self):
        if self.store:
            path = self.store.path
            size = sum(p.stat().st_size for p in (path, Path(f"{path}.wal")) if p.is_file())
            self.cache_info_var.set(f"{path.name}   ·   {size / (1024**2):,.1f} MiB   ·   {path}")
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
