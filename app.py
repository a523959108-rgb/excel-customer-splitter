from __future__ import annotations

import calendar
import json
import os
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import traceback
from datetime import date
from pathlib import Path
from queue import Empty, SimpleQueue
from tkinter import filedialog, font as tkfont, messagebox, ttk

from core import ExportCancelled, load_customer_names, load_region_rules
from warehouse import Warehouse


BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("LOCALAPPDATA") or Path.home() / ".local" / "share") / "ExcelCustomerSplitter"
REGION_CONFIG = BASE_DIR / "regions.json"
CUSTOMER_CONFIG = BASE_DIR / "customer_config.json"


PALETTE = {
    "canvas": "#EEF2F6",
    "header_bg": "#132238",
    "header_card": "#1D314F",
    "header_hover": "#284166",
    "header_border": "#2E4A73",
    "header_ink": "#F8FAFC",
    "header_muted": "#94A3B8",
    "card": "#FFFFFF",
    "subtle": "#F7F9FC",
    "hover": "#F1F5F9",
    "border": "#DCE2EA",
    "border_strong": "#CBD5E1",
    "divider": "#E6EBF1",
    "ink": "#0F172A",
    "ink_sec": "#475569",
    "ink_muted": "#64748B",
    "accent": "#1D4ED8",
    "accent_hover": "#1E40AF",
    "accent_soft": "#EEF4FF",
    "accent_border": "#C7D7FE",
    "accent_ink": "#1E40AF",
    "teal": "#0F766E",
    "teal_hover": "#115E59",
    "teal_soft": "#E6F6F1",
    "teal_border": "#A7F3D0",
    "amber": "#B45309",
    "amber_soft": "#FEF7E6",
    "amber_border": "#F5D08A",
    "danger": "#DC2626",
    "danger_hover": "#B91C1C",
    "danger_soft": "#FEF2F2",
    "danger_border": "#FECACA",
    "disabled_bg": "#E2E8F0",
    "disabled_ink": "#94A3B8",
}


def enable_windows_dpi_awareness():
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            import ctypes

            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def pick_ui_font(root: tk.Misc) -> str:
    try:
        available = set(tkfont.families(root))
    except Exception:
        return "TkDefaultFont"
    for candidate in ("Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", "SimHei"):
        if candidate in available:
            return candidate
    return "TkDefaultFont"


def pick_mono_font(root: tk.Misc) -> str:
    try:
        available = set(tkfont.families(root))
    except Exception:
        return "TkFixedFont"
    for candidate in ("Cascadia Mono", "Consolas", "SF Mono", "Menlo", "Courier New"):
        if candidate in available:
            return candidate
    return "TkFixedFont"


def make_checkbox_icon(checked: bool, size: int = 16) -> tk.PhotoImage:
    """Render a crisp rounded-square checkbox scaled for standard or high-DPI screens."""
    size = max(14, size)
    img = tk.PhotoImage(width=size, height=size)
    bg_fill = PALETTE["accent"] if checked else "#FFFFFF"
    border = PALETTE["accent_hover"] if checked else "#B8C4D4"
    corner = PALETTE["accent_soft"] if checked else "#DCE4F0"

    # Build a thick proportional checkmark stroke inside [1..size-2]
    check_pixels: set[tuple[int, int]] = set()
    if checked:
        p1 = (round(size * 0.25), round(size * 0.52))
        p2 = (round(size * 0.42), round(size * 0.69))
        p3 = (round(size * 0.74), round(size * 0.33))
        for (x0, y0), (x1, y1) in ((p1, p2), (p2, p3)):
            steps = max(abs(x1 - x0), abs(y1 - y0), 1)
            for s in range(steps + 1):
                cx = round(x0 + (x1 - x0) * s / steps)
                cy = round(y0 + (y1 - y0) * s / steps)
                check_pixels.add((cx, cy))
                check_pixels.add((cx, cy + 1))

    rows = []
    transparent: list[tuple[int, int]] = []
    for y in range(size):
        row = []
        for x in range(size):
            if x == 0 or x == size - 1 or y == 0 or y == size - 1:
                row.append("#000000")
                transparent.append((x, y))
                continue
            in_corner = (x in (1, size - 2)) and (y in (1, size - 2))
            if in_corner:
                row.append("#000000")
                transparent.append((x, y))
                continue
            is_soft_corner = (
                (x in (1, size - 2) and y in (2, size - 3))
                or (y in (1, size - 2) and x in (2, size - 3))
            )
            on_edge = x == 1 or x == size - 2 or y == 1 or y == size - 2
            if on_edge:
                row.append(corner if is_soft_corner else border)
            elif checked and (x, y) in check_pixels:
                row.append("#FFFFFF")
            else:
                row.append(bg_fill)
        rows.append("{" + " ".join(row) + "}")

    img.put(" ".join(rows))
    for x, y in transparent:
        img.transparency_set(x, y, True)
    return img


def open_in_file_manager(path: Path):
    target = path if path.is_dir() else path.parent
    if not target.is_dir():
        raise FileNotFoundError(f"路径不存在：{path}")
    if sys.platform == "win32":
        os.startfile(str(target))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(target)])
    else:
        subprocess.Popen(["xdg-open", str(target)])


class Tooltip:
    def __init__(self, widget: tk.Widget, text_getter, ui_font: str = "TkDefaultFont"):
        self.widget = widget
        self.text_getter = text_getter
        self.ui_font = ui_font
        self.tip: tk.Toplevel | None = None
        self._after_id: str | None = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        self._after_id = self.widget.after(450, self._show)

    def _cancel(self):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self):
        text = self.text_getter() if callable(self.text_getter) else self.text_getter
        if not text or self.tip is not None:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        except Exception:
            return
        self.tip = tk.Toplevel(self.widget)
        self.tip.overrideredirect(True)
        self.tip.attributes("-topmost", True)
        frame = tk.Frame(self.tip, bg="#0F172A", padx=10, pady=6)
        frame.pack()
        tk.Label(
            frame,
            text=text,
            bg="#0F172A",
            fg="#F8FAFC",
            font=(self.ui_font, 9),
            justify="left",
            wraplength=420,
        ).pack()
        self.tip.geometry(f"+{x}+{y}")

    def _hide(self, _event=None):
        self._cancel()
        if self.tip is not None:
            try:
                self.tip.destroy()
            except Exception:
                pass
            self.tip = None


class DateCalendar(tk.Frame):
    """Small dependency-free range calendar that enables only imported dates."""

    def __init__(self, master, on_select, ui_font="TkDefaultFont", scale: float = 1.0):
        super().__init__(master, bg=PALETTE["card"], padx=round(12 * scale), pady=round(10 * scale))
        self.on_select = on_select
        self.ui_font = ui_font
        self.available: set[date] = set()
        self.min_date: date | None = None
        self.max_date: date | None = None
        self.start: date | None = None
        self.end: date | None = None
        self.month = date.today().replace(day=1)

        header = tk.Frame(self, bg=PALETTE["card"])
        header.pack(fill="x", pady=(0, round(6 * scale)))
        ttk.Button(header, text="‹", width=3, style="Soft.TButton", command=lambda: self.change_month(-1)).pack(side="left")
        self.month_label = tk.Label(
            header,
            bg=PALETTE["card"],
            fg=PALETTE["ink"],
            font=(ui_font, 10, "bold"),
            anchor="center",
        )
        self.month_label.pack(side="left", fill="x", expand=True)
        ttk.Button(header, text="›", width=3, style="Soft.TButton", command=lambda: self.change_month(1)).pack(side="right")

        self.summary_label = tk.Label(
            self,
            text="未选择日期（导出全部日期）",
            bg=PALETTE["subtle"],
            fg=PALETTE["ink_sec"],
            font=(ui_font, 8),
            pady=round(4 * scale),
            padx=round(8 * scale),
            anchor="w",
        )
        self.summary_label.pack(fill="x", pady=(0, round(6 * scale)))

        self.grid_frame = tk.Frame(self, bg=PALETTE["card"])
        self.grid_frame.pack(fill="x")

        footer = tk.Frame(self, bg=PALETTE["card"])
        footer.pack(fill="x", pady=(round(8 * scale), 0))
        ttk.Button(footer, text="清除日期（导出全部）", style="Soft.TButton", command=self.clear).pack(side="left")
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
        if (
            not parsed
            or (self.start and self.min_date and self.start < self.min_date)
            or (self.end and self.max_date and self.end > self.max_date)
        ):
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

    def _update_summary(self):
        if self.start and self.end:
            days = (self.end - self.start).days + 1
            self.summary_label.configure(
                text=f"已选：{self.start.isoformat()} 至 {self.end.isoformat()}（{days} 天）",
                bg=PALETTE["accent_soft"],
                fg=PALETTE["accent_ink"],
            )
        elif self.start:
            self.summary_label.configure(
                text=f"起始日：{self.start.isoformat()}，请点选结束日",
                bg=PALETTE["amber_soft"],
                fg=PALETTE["amber"],
            )
        elif self.min_date and self.max_date:
            self.summary_label.configure(
                text=f"全部日期（库内 {self.min_date.isoformat()} ~ {self.max_date.isoformat()}）",
                bg=PALETTE["subtle"],
                fg=PALETTE["ink_sec"],
            )
        else:
            self.summary_label.configure(
                text="暂无按日快照，导入 Excel 后可筛选",
                bg=PALETTE["subtle"],
                fg=PALETTE["ink_muted"],
            )

    def render(self):
        for child in self.grid_frame.winfo_children():
            child.destroy()
        self.buttons = []
        self.month_label.configure(text=f"{self.month.year} 年 {self.month.month} 月")
        self._update_summary()

        for column, label in enumerate(("一", "二", "三", "四", "五", "六", "日")):
            tk.Label(
                self.grid_frame,
                text=label,
                bg=PALETTE["subtle"],
                fg=PALETTE["ink_muted"],
                font=(self.ui_font, 8, "bold"),
                width=4,
                pady=3,
            ).grid(row=0, column=column, padx=1, pady=(0, 3), sticky="ew")

        calendar_rows = calendar.monthcalendar(self.month.year, self.month.month)
        for row_index, week in enumerate(calendar_rows, start=1):
            for column, day_number in enumerate(week):
                if not day_number:
                    tk.Label(self.grid_frame, text="", bg=PALETTE["card"], width=4).grid(
                        row=row_index, column=column, padx=1, pady=1
                    )
                    continue
                value = date(self.month.year, self.month.month, day_number)
                in_range_bounds = bool(self.min_date and self.max_date and self.min_date <= value <= self.max_date)
                has_data = value in self.available
                label_text = f"{day_number}·" if (has_data and value not in (self.start, self.end)) else str(day_number)

                button = ttk.Button(
                    self.grid_frame,
                    text=label_text,
                    style="Calendar.TButton",
                    width=4,
                    command=lambda item=value: self.choose(item),
                )
                if not in_range_bounds:
                    button.state(["disabled"])
                if value == self.start or value == self.end:
                    button.configure(style="Chosen.Calendar.TButton")
                elif self.start and self.end and self.start < value < self.end:
                    button.configure(style="Range.Calendar.TButton")
                elif has_data:
                    button.configure(style="Available.Calendar.TButton")
                button.grid(row=row_index, column=column, padx=1, pady=1)
                self.buttons.append(button)


class FieldChipGrid(tk.Frame):
    """Scrollable multi-column checkbox chip matrix without solid-blue selection blocks."""

    def __init__(self, master, ui_font, icon_on, icon_off, on_change, scale: float = 1.0):
        super().__init__(master, bg=PALETTE["card"])
        self.ui_font = ui_font
        self.icon_on = icon_on
        self.icon_off = icon_off
        self.on_change = on_change
        self.scale = scale
        self.items: list[str] = []
        self.visible_indices: list[int] = []
        self.selected: set[int] = set()
        self.chips: dict[int, tk.Label] = {}
        self._cols = 3
        self._filter = ""

        self.canvas = tk.Canvas(
            self,
            bg=PALETTE["card"],
            width=round(200 * scale),
            height=round(90 * scale),
            highlightthickness=0,
            borderwidth=0,
        )
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = tk.Frame(self.canvas, bg=PALETTE["card"])
        self._window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.empty_box = tk.Frame(
            self.canvas,
            bg=PALETTE["subtle"],
            highlightbackground=PALETTE["border"],
            highlightthickness=1,
        )
        self.empty_title = tk.Label(
            self.empty_box,
            text="暂无导出字段",
            bg=PALETTE["subtle"],
            fg=PALETTE["ink"],
            font=(ui_font, 9, "bold"),
        )
        self.empty_desc = tk.Label(
            self.empty_box,
            text="导入 Excel 或打开已有库后将在此展示字段。",
            bg=PALETTE["subtle"],
            fg=PALETTE["ink_muted"],
            font=(ui_font, 8),
            wraplength=round(260 * scale),
            justify="center",
        )
        self.empty_title.pack(padx=round(18 * scale), pady=(round(10 * scale), round(3 * scale)))
        self.empty_desc.pack(padx=round(18 * scale), pady=(0, round(10 * scale)))

        self.inner.bind("<Configure>", self._sync_scroll_region)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        for widget in (self.canvas, self.inner):
            self._bind_wheel(widget)

    def _bind_wheel(self, widget):
        widget.bind("<MouseWheel>", self._on_mousewheel)
        widget.bind("<Button-4>", self._on_linux_wheel)
        widget.bind("<Button-5>", self._on_linux_wheel)

    def _on_mousewheel(self, event):
        if self.inner.winfo_height() <= self.canvas.winfo_height():
            return
        delta = -1 * (event.delta // 120) if abs(event.delta) >= 120 else (-1 if event.delta > 0 else 1)
        self.canvas.yview_scroll(delta, "units")

    def _on_linux_wheel(self, event):
        if self.inner.winfo_height() <= self.canvas.winfo_height():
            return
        self.canvas.yview_scroll(-1 if event.num == 4 else 1, "units")

    def _sync_scroll_region(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self._window_id, width=event.width)
        logical_w = event.width / self.scale
        new_cols = 3 if logical_w >= 480 else (2 if logical_w >= 280 else 1)
        if new_cols != self._cols:
            self._cols = new_cols
            self._relayout()

    def set_items(self, headers: list[str], select_all: bool = True):
        self.items = list(headers)
        self.selected = set(range(len(self.items))) if select_all else set()
        self.rebuild()

    def set_filter(self, keyword: str):
        self._filter = keyword.strip().lower()
        self.rebuild()

    def select_all_visible(self):
        self.selected.update(self.visible_indices)
        self._refresh_styles()
        self.on_change()

    def clear_visible(self):
        self.selected.difference_update(self.visible_indices)
        self._refresh_styles()
        self.on_change()

    def invert_visible(self):
        for idx in self.visible_indices:
            if idx in self.selected:
                self.selected.remove(idx)
            else:
                self.selected.add(idx)
        self._refresh_styles()
        self.on_change()

    def toggle_index(self, idx: int):
        if idx in self.selected:
            self.selected.remove(idx)
        else:
            self.selected.add(idx)
        self._style_chip(idx)
        self.on_change()

    def selected_items(self) -> list[str]:
        return [self.items[i] for i in range(len(self.items)) if i in self.selected]

    def selected_count(self) -> int:
        return len(self.selected)

    def rebuild(self):
        for child in self.inner.winfo_children():
            child.destroy()
        self.chips.clear()

        self.visible_indices = [
            i for i, name in enumerate(self.items) if not self._filter or self._filter in name.lower()
        ]
        if not self.items:
            self.empty_title.configure(text="暂无导出字段")
            self.empty_desc.configure(text="导入 Excel 或打开已有库后将在此展示字段。")
            self.empty_box.place(relx=0.5, rely=0.5, anchor="center")
            self.canvas.yview_moveto(0)
            self.on_change()
            return

        if not self.visible_indices:
            self.empty_title.configure(text="没有匹配字段")
            self.empty_desc.configure(text=f"未找到包含“{self._filter}”的字段。")
            self.empty_box.place(relx=0.5, rely=0.5, anchor="center")
            self.on_change()
            return

        self.empty_box.place_forget()
        px_x = round(8 * self.scale)
        px_y = round(5 * self.scale)
        for idx in self.visible_indices:
            chip = tk.Label(
                self.inner,
                text=f"  {self.items[idx]}",
                compound="left",
                anchor="w",
                font=(self.ui_font, 9),
                padx=px_x,
                pady=px_y,
                cursor="hand2",
                takefocus=True,
                highlightcolor=PALETTE["accent"],
                highlightthickness=1,
            )
            self._bind_wheel(chip)
            chip.bind("<space>", lambda _e, item=idx: self.toggle_index(item))
            chip.bind("<FocusIn>", self._reveal_chip)
            chip.bind("<Button-1>", lambda _e, item=idx: self.toggle_index(item))
            chip.bind("<Enter>", lambda _e, item=idx: self._on_hover(item, True))
            chip.bind("<Leave>", lambda _e, item=idx: self._on_hover(item, False))
            self.chips[idx] = chip
            self._style_chip(idx)

        self._relayout()
        self.canvas.yview_moveto(0)
        self.on_change()

    def _reveal_chip(self, event):
        top = self.canvas.canvasy(0)
        bottom = top + self.canvas.winfo_height()
        y = event.widget.winfo_y()
        end = y + event.widget.winfo_height()
        if y < top or end > bottom:
            target = y if y < top else end - self.canvas.winfo_height()
            self.canvas.yview_moveto(target / max(1, self.inner.winfo_height()))

    def _relayout(self):
        if not self.visible_indices:
            return
        for col in range(4):
            self.inner.columnconfigure(col, weight=0, uniform="")
        for col in range(self._cols):
            self.inner.columnconfigure(col, weight=1, uniform="fieldchip")
        pad = round(3 * self.scale)
        for pos, idx in enumerate(self.visible_indices):
            chip = self.chips[idx]
            row, col = divmod(pos, self._cols)
            chip.grid(row=row, column=col, sticky="ew", padx=pad, pady=pad)

    def _on_hover(self, idx: int, entering: bool):
        chip = self.chips.get(idx)
        if not chip:
            return
        if entering:
            chip.configure(highlightbackground=PALETTE["accent"])
        else:
            self._style_chip(idx)

    def _style_chip(self, idx: int):
        chip = self.chips.get(idx)
        if not chip:
            return
        checked = idx in self.selected
        chip.configure(
            image=self.icon_on if checked else self.icon_off,
            bg=PALETTE["card"] if checked else PALETTE["subtle"],
            fg=PALETTE["ink"] if checked else PALETTE["ink_muted"],
            highlightbackground=PALETTE["accent_border"] if checked else PALETTE["border"],
        )

    def _refresh_styles(self):
        for idx in self.chips:
            self._style_chip(idx)


class SplitterApp(tk.Tk):
    def __init__(self, data_dir=DATA_DIR, *, install_dir=BASE_DIR):
        enable_windows_dpi_awareness()
        super().__init__()
        self.ui_font = pick_ui_font(self)
        self.mono_font = pick_mono_font(self)
        self.dpi_scale = max(1.0, min(2.5, self.winfo_fpixels("1i") / 96.0)) if sys.platform == "win32" else 1.0

        self.title("客户 Excel 拆分工具")
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(round(1080 * self.dpi_scale), screen_w - 32)
        height = min(round(680 * self.dpi_scale), screen_h - 64)
        self.geometry(f"{width}x{height}")
        self.minsize(min(width, round(800 * self.dpi_scale)), min(height, round(500 * self.dpi_scale)))

        self.files: list[str] = []
        self.task_id: str | None = None
        self.headers: list[str] = []
        self.selected_start_date: str | None = None
        self.selected_end_date: str | None = None
        self.customer_names: list[str] = []
        self.selected_customers: set[int] = set()
        self._filtered_customer_indices: list[int] = []
        self._cancel_event = threading.Event()
        self._worker_thread: threading.Thread | None = None
        self._closing = False
        self._active_view = "export"

        self.data_dir = Path(data_dir)
        self.default_database_dir = self.data_dir
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.settings = self.data_dir / "settings.json"
        self.store = None
        self._callbacks = SimpleQueue()

        icon_size = 20 if self.dpi_scale >= 1.35 else 16
        self._cb_on = make_checkbox_icon(True, size=icon_size)
        self._cb_off = make_checkbox_icon(False, size=icon_size)

        self._build_ui()
        self._after_id = self.after(30, self._drain_callbacks)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        try:
            default_location = Path(install_dir) / "database-directory.txt"
            if default_location.is_file():
                location = Path(default_location.read_text(encoding="utf-8-sig").rstrip("\r\n"))
                if not location.is_absolute():
                    raise ValueError("安装配置中的数据库目录必须为完整路径")
                self.default_database_dir = location
            if self.settings.exists():
                path = Path(json.loads(self.settings.read_text(encoding="utf-8"))["database"])
                self.open_store(path, existing=True)
            else:
                self.open_store(self.default_database_dir / "splitter.duckdb")
        except Exception as exc:
            self.toggle_log()
            self.log_line(f"无法打开数据库：{exc}。请点击“打开库”或“新建库”。")
        if any((folder / "splitter.sqlite3").is_file() for folder in (self.data_dir, BASE_DIR)):
            self.log_line("检测到旧 SQLite 数据，可在顶部“配置维护”中迁入新建空库。")

    def _s(self, value: int) -> int:
        return round(value * self.dpi_scale)

    def _configure_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        f_ui = self.ui_font
        s = self._s

        style.configure(".", background=PALETTE["card"], foreground=PALETTE["ink"], bordercolor=PALETTE["border"], focuscolor=PALETTE["accent"], font=(f_ui, 9))
        style.configure("TFrame", background=PALETTE["card"])
        style.configure("TLabel", background=PALETTE["card"], foreground=PALETTE["ink"], font=(f_ui, 9))
        style.configure("Muted.TLabel", background=PALETTE["card"], foreground=PALETTE["ink_muted"], font=(f_ui, 9))

        style.configure(
            "Header.TButton",
            background=PALETTE["header_card"],
            foreground=PALETTE["header_ink"],
            bordercolor=PALETTE["header_border"],
            lightcolor=PALETTE["header_card"],
            darkcolor=PALETTE["header_card"],
            borderwidth=1,
            padding=(s(8), s(4)),
            font=(f_ui, 9),
        )
        style.map(
            "Header.TButton",
            background=[("disabled", "#172740"), ("active", PALETTE["header_hover"])],
            foreground=[("disabled", "#64748B"), ("!disabled", PALETTE["header_ink"])],
        )
        style.configure(
            "HeaderActive.TButton",
            background=PALETTE["accent"],
            foreground="#FFFFFF",
            bordercolor=PALETTE["accent"],
            lightcolor=PALETTE["accent"],
            darkcolor=PALETTE["accent"],
            borderwidth=1,
            padding=(s(8), s(4)),
            font=(f_ui, 9, "bold"),
        )
        style.map("HeaderActive.TButton", background=[("active", PALETTE["accent_hover"])])

        style.configure(
            "TButton",
            background=PALETTE["subtle"],
            foreground=PALETTE["ink"],
            bordercolor=PALETTE["border_strong"],
            lightcolor=PALETTE["subtle"],
            darkcolor=PALETTE["subtle"],
            borderwidth=1,
            padding=(s(8), s(4)),
            font=(f_ui, 9),
        )
        style.map(
            "TButton",
            background=[("disabled", PALETTE["subtle"]), ("active", PALETTE["hover"])],
            foreground=[("disabled", PALETTE["disabled_ink"]), ("!disabled", PALETTE["ink"])],
        )

        style.configure(
            "Soft.TButton",
            background=PALETTE["card"],
            foreground=PALETTE["ink_sec"],
            bordercolor=PALETTE["border"],
            lightcolor=PALETTE["card"],
            darkcolor=PALETTE["card"],
            borderwidth=1,
            padding=(s(8), s(4)),
            font=(f_ui, 9),
        )
        style.map(
            "Soft.TButton",
            background=[("disabled", PALETTE["subtle"]), ("active", PALETTE["accent_soft"])],
            foreground=[("disabled", PALETTE["disabled_ink"]), ("active", PALETTE["accent_ink"])],
            bordercolor=[("active", PALETTE["accent_border"])],
        )

        style.configure(
            "Mini.TButton",
            background=PALETTE["card"],
            foreground=PALETTE["ink_sec"],
            bordercolor=PALETTE["border"],
            lightcolor=PALETTE["card"],
            darkcolor=PALETTE["card"],
            borderwidth=1,
            padding=(s(6), s(2)),
            font=(f_ui, 8),
        )
        style.map(
            "Mini.TButton",
            background=[("disabled", PALETTE["subtle"]), ("active", PALETTE["accent_soft"])],
            foreground=[("disabled", PALETTE["disabled_ink"]), ("active", PALETTE["accent_ink"])],
            bordercolor=[("active", PALETTE["accent_border"])],
        )

        style.configure(
            "SoftAccent.TButton",
            background=PALETTE["accent_soft"],
            foreground=PALETTE["accent_ink"],
            bordercolor=PALETTE["accent_border"],
            lightcolor=PALETTE["accent_soft"],
            darkcolor=PALETTE["accent_soft"],
            borderwidth=1,
            padding=(s(9), s(4)),
            font=(f_ui, 9, "bold"),
        )
        style.map(
            "SoftAccent.TButton",
            background=[("disabled", PALETTE["subtle"]), ("active", "#DCE8FF")],
            foreground=[("disabled", PALETTE["disabled_ink"]), ("!disabled", PALETTE["accent_ink"])],
        )

        style.configure(
            "Teal.TButton",
            background=PALETTE["teal"],
            foreground="#FFFFFF",
            bordercolor=PALETTE["teal"],
            lightcolor=PALETTE["teal"],
            darkcolor=PALETTE["teal"],
            borderwidth=1,
            padding=(s(12), s(4)),
            font=(f_ui, 9, "bold"),
        )
        style.map(
            "Teal.TButton",
            background=[("disabled", PALETTE["disabled_bg"]), ("active", PALETTE["teal_hover"])],
            foreground=[("disabled", PALETTE["disabled_ink"]), ("!disabled", "#FFFFFF")],
        )

        style.configure(
            "Accent.TButton",
            background=PALETTE["accent"],
            foreground="#FFFFFF",
            bordercolor=PALETTE["accent"],
            lightcolor=PALETTE["accent"],
            darkcolor=PALETTE["accent"],
            borderwidth=1,
            padding=(s(14), s(6)),
            font=(f_ui, 9, "bold"),
        )
        style.map(
            "Accent.TButton",
            background=[("disabled", PALETTE["disabled_bg"]), ("active", PALETTE["accent_hover"])],
            foreground=[("disabled", PALETTE["disabled_ink"]), ("!disabled", "#FFFFFF")],
        )

        style.configure(
            "Danger.TButton",
            background=PALETTE["danger"],
            foreground="#FFFFFF",
            bordercolor=PALETTE["danger"],
            lightcolor=PALETTE["danger"],
            darkcolor=PALETTE["danger"],
            borderwidth=1,
            padding=(s(10), s(6)),
            font=(f_ui, 9, "bold"),
        )
        style.map(
            "Danger.TButton",
            background=[("disabled", PALETTE["subtle"]), ("active", PALETTE["danger_hover"])],
            foreground=[("disabled", PALETTE["disabled_ink"]), ("!disabled", "#FFFFFF")],
            bordercolor=[("disabled", PALETTE["border"]), ("!disabled", PALETTE["danger"])],
        )

        style.configure(
            "TEntry",
            fieldbackground="#FFFFFF",
            foreground=PALETTE["ink"],
            bordercolor=PALETTE["border_strong"],
            lightcolor="#FFFFFF",
            darkcolor="#FFFFFF",
            padding=(s(6), s(4)),
            font=(f_ui, 9),
        )
        style.map("TEntry", bordercolor=[("focus", PALETTE["accent"])])

        style.configure(
            "TCombobox",
            fieldbackground="#FFFFFF",
            background=PALETTE["subtle"],
            foreground=PALETTE["ink"],
            bordercolor=PALETTE["border_strong"],
            arrowcolor=PALETTE["ink_sec"],
            padding=(s(6), s(3)),
            font=(f_ui, 9),
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#FFFFFF"), ("disabled", PALETTE["subtle"])],
            bordercolor=[("focus", PALETTE["accent"])],
        )

        style.configure(
            "Customer.Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=PALETTE["ink"],
            borderwidth=0,
            rowheight=s(26),
            font=(f_ui, 9),
        )
        style.map(
            "Customer.Treeview",
            background=[("selected", PALETTE["hover"])],
            foreground=[("selected", PALETTE["ink"])],
        )

        style.configure("Calendar.TButton", background="#FFFFFF", foreground=PALETTE["ink"], bordercolor=PALETTE["border"], padding=(s(3), s(3)), font=(f_ui, 9))
        style.map("Calendar.TButton", background=[("active", PALETTE["hover"]), ("disabled", "#FAFAFA")], foreground=[("disabled", "#CBD5E1")])
        style.configure("Available.Calendar.TButton", background=PALETTE["subtle"], foreground=PALETTE["ink"], bordercolor=PALETTE["border_strong"], padding=(s(3), s(3)), font=(f_ui, 9, "bold"))
        style.map("Available.Calendar.TButton", background=[("active", PALETTE["accent_soft"])])
        style.configure("Chosen.Calendar.TButton", background=PALETTE["accent"], foreground="#FFFFFF", bordercolor=PALETTE["accent"], padding=(s(3), s(3)), font=(f_ui, 9, "bold"))
        style.map("Chosen.Calendar.TButton", background=[("active", PALETTE["accent_hover"])])
        style.configure("Range.Calendar.TButton", background=PALETTE["accent_soft"], foreground=PALETTE["accent_ink"], bordercolor=PALETTE["accent_border"], padding=(s(3), s(3)), font=(f_ui, 9))

        style.configure(
            "Horizontal.TProgressbar",
            background=PALETTE["accent"],
            troughcolor="#E2E8F0",
            bordercolor="#E2E8F0",
            lightcolor=PALETTE["accent"],
            darkcolor=PALETTE["accent"],
            thickness=s(6),
        )

    def _build_ui(self):
        self._configure_styles()
        self.configure(background=PALETTE["canvas"])
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        s = self._s

        self.database_buttons = []
        self.cache_info_var = tk.StringVar()
        self.db_title_var = tk.StringVar(value="未加载数据库")
        self.db_meta_var = tk.StringVar(value="请打开或新建库")
        self.db_path_full = ""
        self.status_var = tk.StringVar(value="就绪")
        self.output_var = tk.StringVar(value=str(BASE_DIR / "output"))
        self.customer_config_var = tk.StringVar(value=str(CUSTOMER_CONFIG))
        self.region_path_var = tk.StringVar(value=str(REGION_CONFIG))
        self.customer_var = tk.StringVar()
        self.region_var = tk.StringVar()
        self.customer_search_var = tk.StringVar()
        self.field_search_var = tk.StringVar()
        self.customer_count = tk.StringVar(value="全部客户")
        self.field_count = tk.StringVar(value="0 / 0")
        self.preview_structure_var = tk.StringVar(value="输出结构：日期 / 未分区 / 客户.xlsx")

        # ------------------------------------------------------------------
        # Row 0: Deep Ink-Navy Command Header (grid keeps right buttons locked visible)
        # ------------------------------------------------------------------
        header = tk.Frame(self, bg=PALETTE["header_bg"], padx=s(14), pady=s(8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)

        brand = tk.Frame(header, bg=PALETTE["header_bg"])
        brand.grid(row=0, column=0, sticky="w")
        tk.Frame(brand, bg="#3B82F6", width=s(4), height=s(20)).pack(side="left", padx=(0, s(8)))
        tk.Label(
            brand,
            text="客户拆分",
            bg=PALETTE["header_bg"],
            fg=PALETTE["header_ink"],
            font=(self.ui_font, 12, "bold"),
        ).pack(side="left", padx=(0, s(10)))

        # Database Context Capsule (shrinks cleanly instead of pushing right controls out)
        self.db_capsule = tk.Frame(
            header,
            bg=PALETTE["header_card"],
            highlightbackground=PALETTE["header_border"],
            highlightthickness=1,
            height=s(28),
            padx=s(9),
        )
        self.db_capsule.pack_propagate(False)
        self.db_capsule.grid(row=0, column=1, sticky="ew", padx=(0, s(10)))

        self.db_dot = tk.Canvas(self.db_capsule, width=s(8), height=s(8), bg=PALETTE["header_card"], highlightthickness=0)
        self.db_dot.pack(side="left", padx=(0, s(6)))
        self._db_dot_id = self.db_dot.create_oval(1, 1, s(7), s(7), fill="#FBBF24", outline="")
        tk.Label(
            self.db_capsule,
            textvariable=self.db_title_var,
            bg=PALETTE["header_card"],
            fg=PALETTE["header_ink"],
            font=(self.ui_font, 9, "bold"),
        ).pack(side="left")
        tk.Label(
            self.db_capsule,
            textvariable=self.db_meta_var,
            bg=PALETTE["header_card"],
            fg=PALETTE["header_muted"],
            font=(self.ui_font, 8),
        ).pack(side="left", padx=(s(8), 0))
        Tooltip(self.db_capsule, lambda: self.db_path_full or "当前未打开数据库", ui_font=self.ui_font)

        # Right Header Controls (always reachable)
        nav = tk.Frame(header, bg=PALETTE["header_bg"])
        nav.grid(row=0, column=2, sticky="e")
        for label, command in (("打开库", self.choose_database), ("新建库", self.create_database)):
            btn = ttk.Button(nav, text=label, style="Header.TButton", command=command)
            btn.pack(side="left", padx=(0, s(4)))
            self.database_buttons.append(btn)

        tk.Frame(nav, bg=PALETTE["header_border"], width=1, height=s(18)).pack(side="left", padx=(s(4), s(8)))
        self.nav_export_btn = ttk.Button(nav, text="拆分导出", style="HeaderActive.TButton", command=lambda: self.switch_view("export"))
        self.nav_export_btn.pack(side="left", padx=(0, s(4)))
        self.nav_settings_btn = ttk.Button(nav, text="配置维护", style="Header.TButton", command=lambda: self.switch_view("settings"))
        self.nav_settings_btn.pack(side="left")

        # ------------------------------------------------------------------
        # Row 1: Streamlined Excel Intake Strip (right buttons packed first)
        # ------------------------------------------------------------------
        self.intake_outer = tk.Frame(
            self,
            bg=PALETTE["card"],
            highlightbackground=PALETTE["border"],
            highlightthickness=1,
            padx=s(12),
            pady=s(6),
        )
        self.intake_outer.grid(row=1, column=0, sticky="ew", padx=s(12), pady=(s(8), 0))

        intake_bar = tk.Frame(self.intake_outer, bg=PALETTE["card"])
        intake_bar.pack(fill="x")
        self.import_toggle = ttk.Button(intake_bar, text="展开 Excel 导入", style="Soft.TButton", command=self.toggle_import)
        self.import_toggle.pack(side="right")

        tk.Label(
            intake_bar,
            text="Excel 导入",
            bg=PALETTE["card"],
            fg=PALETTE["ink"],
            font=(self.ui_font, 9, "bold"),
        ).pack(side="left")
        self.intake_summary = tk.Label(
            intake_bar,
            text="同日文件须一次选齐，重导整日覆盖",
            bg=PALETTE["card"],
            fg=PALETTE["ink_muted"],
            font=(self.ui_font, 8),
        )
        self.intake_summary.pack(side="left", padx=(s(8), s(8)))

        self.import_panel = tk.Frame(
            self.intake_outer,
            bg=PALETTE["amber_soft"],
            highlightbackground=PALETTE["amber_border"],
            highlightthickness=1,
            padx=s(10),
            pady=s(6),
        )
        self.import_button = ttk.Button(self.import_panel, text="开始导入", style="Teal.TButton", command=self.scan_and_import)
        self.import_button.pack(side="right")
        ttk.Button(self.import_panel, text="选择 Excel…", style="SoftAccent.TButton", command=self.choose_files).pack(side="left")
        self.file_label = tk.Label(
            self.import_panel,
            text="未选文件（可多选 .xlsx）",
            bg=PALETTE["amber_soft"],
            fg=PALETTE["ink"],
            font=(self.ui_font, 9),
        )
        self.file_label.pack(side="left", padx=(s(8), 0))
        self.set_import_open(True)

        # ------------------------------------------------------------------
        # Row 2: Main Workspace (Export Workbench / Settings)
        # ------------------------------------------------------------------
        self.workspace_host = tk.Frame(self, bg=PALETTE["canvas"])
        self.workspace_host.grid(row=2, column=0, sticky="nsew", padx=s(12), pady=s(8))
        self.workspace_host.columnconfigure(0, weight=1)
        self.workspace_host.rowconfigure(0, weight=1)

        self._build_export_view()
        self._build_settings_view()
        self.switch_view("export")

        # ------------------------------------------------------------------
        # Row 3: Collapsible Log Drawer
        # ------------------------------------------------------------------
        self.log_panel = tk.Frame(
            self,
            bg=PALETTE["card"],
            highlightbackground=PALETTE["border"],
            highlightthickness=1,
        )
        log_head = tk.Frame(self.log_panel, bg=PALETTE["subtle"], padx=s(10), pady=s(4))
        log_head.pack(fill="x")
        self.open_manifest_btn = ttk.Button(log_head, text="打开目录", style="Mini.TButton", command=self.open_output_dir)
        self.open_manifest_btn.pack(side="right", padx=(s(4), 0))
        ttk.Button(log_head, text="清空", style="Mini.TButton", command=self.clear_log).pack(side="right")
        tk.Label(log_head, text="运行日志", bg=PALETTE["subtle"], fg=PALETTE["ink"], font=(self.ui_font, 9, "bold")).pack(side="left")

        log_body = tk.Frame(self.log_panel, bg=PALETTE["card"], padx=s(10), pady=s(5))
        log_body.pack(fill="both", expand=True)
        self.log = tk.Text(
            log_body,
            height=4,
            state="disabled",
            background="#FFFFFF",
            foreground=PALETTE["ink"],
            borderwidth=0,
            wrap="word",
            highlightthickness=0,
            font=(self.mono_font, 9),
        )
        self.log.tag_configure("ok", foreground=PALETTE["teal"])
        self.log.tag_configure("warn", foreground=PALETTE["amber"])
        self.log.tag_configure("err", foreground=PALETTE["danger"])
        scroll = ttk.Scrollbar(log_body, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)

        # ------------------------------------------------------------------
        # Row 4: Docked Output & Action Bar (right buttons packed first)
        # ------------------------------------------------------------------
        dock = tk.Frame(
            self,
            bg=PALETTE["card"],
            highlightbackground=PALETTE["border"],
            highlightthickness=1,
            padx=s(12),
            pady=s(8),
        )
        dock.grid(row=4, column=0, sticky="ew", padx=s(12), pady=(0, s(8)))

        dest_row = tk.Frame(dock, bg=PALETTE["card"])
        dest_row.pack(fill="x", pady=(0, s(6)))
        ttk.Button(dest_row, text="打开", style="Soft.TButton", command=self.open_output_dir).pack(side="right")
        ttk.Button(dest_row, text="浏览…", style="Soft.TButton", command=self.choose_output).pack(side="right", padx=(s(6), s(4)))
        tk.Label(dest_row, text="输出目录", bg=PALETTE["card"], fg=PALETTE["ink_sec"], font=(self.ui_font, 9, "bold")).pack(side="left", padx=(0, s(8)))
        ttk.Entry(dest_row, textvariable=self.output_var).pack(side="left", fill="x", expand=True)

        tk.Frame(dock, bg=PALETTE["divider"], height=1).pack(fill="x", pady=(0, s(6)))

        actions = tk.Frame(dock, bg=PALETTE["card"])
        actions.pack(fill="x")
        self.run_button = ttk.Button(actions, text="开始拆分导出", style="Accent.TButton", command=self.start_export)
        self.run_button.pack(side="right")
        self.stop_button = ttk.Button(actions, text="停止任务", style="Danger.TButton", command=self.stop_and_cleanup, state="disabled")
        self.stop_button.pack(side="right", padx=(0, s(6)))

        self.log_toggle = ttk.Button(actions, text="显示日志", style="Soft.TButton", command=self.toggle_log)
        self.log_toggle.pack(side="left")

        status_box = tk.Frame(actions, bg=PALETTE["subtle"], padx=s(8), pady=s(3))
        status_box.pack(side="left", padx=(s(8), s(8)))
        self.status_dot = tk.Canvas(status_box, width=s(7), height=s(7), bg=PALETTE["subtle"], highlightthickness=0)
        self.status_dot.pack(side="left", padx=(0, s(5)))
        self._status_dot_id = self.status_dot.create_oval(1, 1, s(6), s(6), fill=PALETTE["teal"], outline="")
        tk.Label(status_box, textvariable=self.status_var, bg=PALETTE["subtle"], fg=PALETTE["ink"], font=(self.ui_font, 8, "bold")).pack(side="left")

        self.progress = ttk.Progressbar(actions, length=s(56))
        self.progress.pack(side="left")

        self.refresh_cache_info()

    def _build_export_view(self):
        s = self._s
        self.export_view = tk.Frame(self.workspace_host, bg=PALETTE["canvas"])
        self.export_view.columnconfigure(0, weight=32, minsize=s(195))
        self.export_view.columnconfigure(1, weight=68, minsize=s(360))
        self.export_view.rowconfigure(0, weight=1)

        # ---- Left Card: Customer List ----
        cust_card = tk.Frame(
            self.export_view,
            bg=PALETTE["card"],
            highlightbackground=PALETTE["border"],
            highlightthickness=1,
            padx=s(10),
            pady=s(9),
        )
        cust_card.grid(row=0, column=0, sticky="nsew", padx=(0, s(8)))
        cust_card.columnconfigure(0, weight=1)
        cust_card.rowconfigure(2, weight=1)

        cust_head = tk.Frame(cust_card, bg=PALETTE["card"])
        cust_head.grid(row=0, column=0, sticky="ew")
        self.cust_quick_btns = tk.Frame(cust_head, bg=PALETTE["card"])
        self.cust_quick_btns.pack(side="right")
        ttk.Button(self.cust_quick_btns, text="清空", style="Mini.TButton", command=self.clear_all_customers).pack(side="right")
        ttk.Button(self.cust_quick_btns, text="全选", style="Mini.TButton", command=self.select_all_customers).pack(side="right", padx=(0, s(3)))

        tk.Label(cust_head, text="客户名单", bg=PALETTE["card"], fg=PALETTE["ink"], font=(self.ui_font, 10, "bold")).pack(side="left")
        self.customer_badge = tk.Label(
            cust_head,
            textvariable=self.customer_count,
            bg=PALETTE["accent_soft"],
            fg=PALETTE["accent_ink"],
            font=(self.ui_font, 8, "bold"),
            padx=s(6),
            pady=s(1),
        )
        self.customer_badge.pack(side="left", padx=(s(6), 0))

        # Search row
        cust_search_row = tk.Frame(cust_card, bg=PALETTE["card"])
        cust_search_row.grid(row=1, column=0, sticky="ew", pady=(s(6), s(6)))
        ttk.Button(cust_search_row, text="×", width=2, style="Mini.TButton", command=lambda: self.customer_search_var.set("")).pack(side="right", padx=(s(4), 0))
        self.cust_search_entry = ttk.Entry(cust_search_row, textvariable=self.customer_search_var, width=8)
        self.cust_search_entry.pack(side="left", fill="x", expand=True)
        self.customer_search_var.trace_add("write", lambda *_: self._refresh_customer_tree())

        # Tree container + Compact Empty state
        cust_list_wrap = tk.Frame(
            cust_card,
            bg=PALETTE["card"],
            highlightbackground=PALETTE["border"],
            highlightthickness=1,
        )
        cust_list_wrap.grid(row=2, column=0, sticky="nsew")
        cust_list_wrap.columnconfigure(0, weight=1)
        cust_list_wrap.rowconfigure(0, weight=1)

        self.customer_tree = ttk.Treeview(
            cust_list_wrap,
            show="tree",
            height=4,
            selectmode="browse",
            style="Customer.Treeview",
        )
        self.customer_tree.column("#0", width=s(140), stretch=True, minwidth=s(80))
        self.customer_tree.tag_configure("even", background="#FFFFFF")
        self.customer_tree.tag_configure("odd", background="#FAFBFD")
        self.customer_tree.tag_configure("unchecked", foreground=PALETTE["ink_muted"])
        self.customer_tree.bind("<Button-1>", self._on_customer_click)
        self.customer_tree.bind("<space>", self._on_customer_space)

        cust_scroll = ttk.Scrollbar(cust_list_wrap, orient="vertical", command=self.customer_tree.yview)
        self.customer_tree.configure(yscrollcommand=cust_scroll.set)
        self.customer_tree.grid(row=0, column=0, sticky="nsew")
        cust_scroll.grid(row=0, column=1, sticky="ns")

        # Concise empty-state overlay (never clips on compact high-DPI heights)
        self.empty_customers = tk.Frame(cust_list_wrap, bg=PALETTE["subtle"], padx=s(10), pady=s(10))
        tk.Label(
            self.empty_customers,
            text="未加载名单 · 导出全部客户",
            bg=PALETTE["subtle"],
            fg=PALETTE["ink"],
            font=(self.ui_font, 9, "bold"),
            wraplength=s(180),
            justify="center",
        ).pack(pady=(s(2), s(3)))
        tk.Label(
            self.empty_customers,
            text="可导入 JSON 筛选指定客户。",
            bg=PALETTE["subtle"],
            fg=PALETTE["ink_muted"],
            font=(self.ui_font, 8),
            wraplength=s(180),
            justify="center",
        ).pack(pady=(0, s(8)))
        ttk.Button(self.empty_customers, text="选择名单…", style="SoftAccent.TButton", command=self.choose_customer_config).pack()

        cust_foot = tk.Frame(cust_card, bg=PALETTE["card"])
        cust_foot.grid(row=3, column=0, sticky="ew", pady=(s(7), 0))
        self.clear_cust_cfg_btn = ttk.Button(cust_foot, text="全部客户", style="Mini.TButton", command=self.reset_to_all_customers)
        self.clear_cust_cfg_btn.pack(side="right")
        ttk.Button(cust_foot, text="选择名单…", style="Soft.TButton", command=self.choose_customer_config).pack(side="left")

        # ---- Right Card: Split Rules & Field Matrix ----
        rules_card = tk.Frame(
            self.export_view,
            bg=PALETTE["card"],
            highlightbackground=PALETTE["border"],
            highlightthickness=1,
            padx=s(10),
            pady=s(9),
        )
        rules_card.grid(row=0, column=1, sticky="nsew")
        rules_card.columnconfigure(0, weight=1)
        rules_card.rowconfigure(2, weight=1)

        # Top Mapping + Date Strip
        rule_box = tk.Frame(
            rules_card,
            bg=PALETTE["subtle"],
            highlightbackground=PALETTE["border"],
            highlightthickness=1,
            padx=s(10),
            pady=s(7),
        )
        rule_box.grid(row=0, column=0, sticky="ew")
        rule_box.columnconfigure((0, 1, 2), weight=1, uniform="rulecol")

        tk.Label(rule_box, text="客户字段", bg=PALETTE["subtle"], fg=PALETTE["ink_sec"], font=(self.ui_font, 8, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(rule_box, text="分区字段（可选）", bg=PALETTE["subtle"], fg=PALETTE["ink_sec"], font=(self.ui_font, 8, "bold")).grid(row=0, column=1, sticky="w", padx=(s(8), 0))
        tk.Label(rule_box, text="日期范围", bg=PALETTE["subtle"], fg=PALETTE["ink_sec"], font=(self.ui_font, 8, "bold")).grid(row=0, column=2, sticky="w", padx=(s(8), 0))

        self.customer_combo = ttk.Combobox(rule_box, textvariable=self.customer_var, state="readonly", width=8)
        self.customer_combo.grid(row=1, column=0, sticky="ew", pady=(s(3), 0))
        self.region_combo = ttk.Combobox(rule_box, textvariable=self.region_var, state="readonly", width=8)
        self.region_combo.grid(row=1, column=1, sticky="ew", padx=(s(8), 0), pady=(s(3), 0))
        self.region_var.trace_add("write", lambda *_: self._update_preview_structure())

        date_wrap = tk.Frame(rule_box, bg=PALETTE["subtle"])
        date_wrap.grid(row=1, column=2, sticky="ew", padx=(s(8), 0), pady=(s(3), 0))
        date_wrap.columnconfigure(0, weight=1)
        self.date_button = ttk.Button(date_wrap, text="📅 全部日期", style="SoftAccent.TButton", command=self.show_calendar)
        self.date_button.grid(row=0, column=0, sticky="ew")
        self.date_clear_btn = ttk.Button(date_wrap, text="×", width=2, style="Mini.TButton", command=lambda: self.calendar.clear())

        tk.Label(
            rule_box,
            textvariable=self.preview_structure_var,
            bg=PALETTE["subtle"],
            fg=PALETTE["ink_muted"],
            font=(self.ui_font, 8),
            anchor="w",
        ).grid(row=2, column=0, columnspan=3, sticky="ew", pady=(s(5), 0))

        # Calendar Toplevel
        self.calendar_window = tk.Toplevel(self)
        self.calendar_window.withdraw()
        self.calendar_window.title("选择导出日期")
        self.calendar_window.configure(bg=PALETTE["card"])
        self.calendar_window.transient(self)
        self.calendar_window.resizable(False, False)
        self.calendar_window.protocol("WM_DELETE_WINDOW", self.calendar_window.withdraw)
        self.calendar_window.bind("<Escape>", lambda _: self.calendar_window.withdraw())
        self.calendar = DateCalendar(self.calendar_window, self.set_selected_date, ui_font=self.ui_font, scale=self.dpi_scale)
        self.calendar.pack(fill="both", expand=True)
        cal_bar = tk.Frame(self.calendar_window, bg=PALETTE["subtle"], padx=s(12), pady=s(8))
        cal_bar.pack(fill="x")
        ttk.Button(cal_bar, text="完成", style="Accent.TButton", command=self.calendar_window.withdraw).pack(side="right")

        # Export Fields Header (right action buttons packed first so never pushed out)
        field_head = tk.Frame(rules_card, bg=PALETTE["card"])
        field_head.grid(row=1, column=0, sticky="ew", pady=(s(8), s(6)))
        ttk.Button(field_head, text="清空", style="Mini.TButton", command=lambda: self.field_grid.clear_visible()).pack(side="right")
        ttk.Button(field_head, text="反选", style="Mini.TButton", command=lambda: self.field_grid.invert_visible()).pack(side="right", padx=(0, s(3)))
        ttk.Button(field_head, text="全选", style="Mini.TButton", command=lambda: self.field_grid.select_all_visible()).pack(side="right", padx=(0, s(3)))

        tk.Label(field_head, text="导出字段", bg=PALETTE["card"], fg=PALETTE["ink"], font=(self.ui_font, 10, "bold")).pack(side="left")
        self.field_badge = tk.Label(
            field_head,
            textvariable=self.field_count,
            bg=PALETTE["accent_soft"],
            fg=PALETTE["accent_ink"],
            font=(self.ui_font, 8, "bold"),
            padx=s(6),
            pady=s(1),
        )
        self.field_badge.pack(side="left", padx=(s(6), s(8)))

        self.field_search_entry = ttk.Entry(field_head, textvariable=self.field_search_var, width=5)
        self.field_search_entry.pack(side="left", fill="x", expand=True, padx=(0, s(4)))
        self.field_search_var.trace_add("write", lambda *_: self.field_grid.set_filter(self.field_search_var.get()))

        # Field Chip Matrix
        self.field_grid = FieldChipGrid(
            rules_card,
            self.ui_font,
            self._cb_on,
            self._cb_off,
            on_change=self.update_counts,
            scale=self.dpi_scale,
        )
        self.field_grid.grid(row=2, column=0, sticky="nsew")

    def _build_settings_view(self):
        s = self._s
        self.settings_view = tk.Frame(self.workspace_host, bg=PALETTE["canvas"])
        self.settings_view.columnconfigure((0, 1), weight=1, uniform="setcol")
        self.settings_view.rowconfigure(0, weight=1)

        cfg_card = tk.Frame(
            self.settings_view,
            bg=PALETTE["card"],
            highlightbackground=PALETTE["border"],
            highlightthickness=1,
            padx=s(14),
            pady=s(12),
        )
        cfg_card.grid(row=0, column=0, sticky="nsew", padx=(0, s(8)))
        tk.Label(cfg_card, text="配置文件与规则", bg=PALETTE["card"], fg=PALETTE["ink"], font=(self.ui_font, 10, "bold")).pack(anchor="w", pady=(0, s(6)))

        for label, var, browse in (
            ("客户名单配置 (.json)", self.customer_config_var, self.choose_customer_config),
            ("分区映射配置 (.json)", self.region_path_var, self.choose_region_config),
        ):
            tk.Label(cfg_card, text=label, bg=PALETTE["card"], fg=PALETTE["ink_sec"], font=(self.ui_font, 8, "bold")).pack(anchor="w", pady=(s(4), s(2)))
            self.path_row(cfg_card, var, browse)

        ttk.Button(cfg_card, text="重新加载预设客户名单", style="SoftAccent.TButton", command=self.load_customer_config).pack(anchor="w", pady=(s(8), 0))

        db_card = tk.Frame(
            self.settings_view,
            bg=PALETTE["card"],
            highlightbackground=PALETTE["border"],
            highlightthickness=1,
            padx=s(14),
            pady=s(12),
        )
        db_card.grid(row=0, column=1, sticky="nsew")
        tk.Label(db_card, text="数据库维护与历史迁移", bg=PALETTE["card"], fg=PALETTE["ink"], font=(self.ui_font, 10, "bold")).pack(anchor="w", pady=(0, s(8)))

        for title, desc, btn_text, cmd, btn_style in (
            (
                "迁移旧 SQLite 数据库",
                "只读迁入旧 splitter.sqlite3 历史批次并核对分月金额（需在新建空库中执行）。",
                "迁移旧 SQLite…",
                self.migrate_database,
                "SoftAccent.TButton",
            ),
            (
                "整理事务日志（Checkpoint）",
                "将挂起的 DuckDB 事务日志写入持久文件并清理临时缓存。",
                "整理日志",
                self.clear_database_cache,
                "Soft.TButton",
            ),
        ):
            row = tk.Frame(db_card, bg=PALETTE["subtle"], highlightbackground=PALETTE["border"], highlightthickness=1, padx=s(10), pady=s(8))
            row.pack(fill="x", pady=(0, s(8)))
            btn = ttk.Button(row, text=btn_text, style=btn_style, command=cmd)
            btn.pack(side="right", padx=(s(8), 0))
            self.database_buttons.append(btn)
            left = tk.Frame(row, bg=PALETTE["subtle"])
            left.pack(side="left", fill="x", expand=True)
            tk.Label(left, text=title, bg=PALETTE["subtle"], fg=PALETTE["ink"], font=(self.ui_font, 9, "bold")).pack(anchor="w")
            tk.Label(left, text=desc, bg=PALETTE["subtle"], fg=PALETTE["ink_muted"], font=(self.ui_font, 8), wraplength=s(230), justify="left").pack(anchor="w", pady=(s(2), 0))

    def switch_view(self, view: str):
        self._active_view = view
        if view == "export":
            self.settings_view.grid_forget()
            self.export_view.grid(row=0, column=0, sticky="nsew")
            self.nav_export_btn.configure(style="HeaderActive.TButton")
            self.nav_settings_btn.configure(style="Header.TButton")
        else:
            self.export_view.grid_forget()
            self.settings_view.grid(row=0, column=0, sticky="nsew")
            self.nav_export_btn.configure(style="Header.TButton")
            self.nav_settings_btn.configure(style="HeaderActive.TButton")

    def path_row(self, parent, variable, browse):
        row = tk.Frame(parent, bg=parent.cget("bg"))
        row.pack(fill="x", pady=(self._s(2), self._s(4)))
        ttk.Button(row, text="浏览…", style="Soft.TButton", command=browse).pack(side="right", padx=(self._s(6), 0))
        ttk.Entry(row, textvariable=variable, width=12).pack(side="left", fill="x", expand=True)

    def set_import_open(self, visible: bool):
        if visible:
            self.import_panel.pack(fill="x", pady=(self._s(6), 0))
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
            self.log_panel.grid(row=3, column=0, sticky="ew", padx=self._s(12), pady=(0, self._s(6)))
            self.log_toggle.configure(text="收起日志")

    def clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.configure(state="disabled")

    def open_output_dir(self):
        target = Path(self.output_var.get().strip() or (BASE_DIR / "output"))
        try:
            open_in_file_manager(target)
        except Exception as exc:
            messagebox.showwarning("无法打开目录", str(exc))

    def show_calendar(self):
        self.update_idletasks()
        try:
            bx = self.date_button.winfo_rootx()
            by = self.date_button.winfo_rooty() + self.date_button.winfo_height() + 4
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            self.calendar_window.update_idletasks()
            cw = max(self._s(260), self.calendar_window.winfo_reqwidth())
            ch = max(self._s(260), self.calendar_window.winfo_reqheight())
            x = max(8, min(bx, sw - cw - 12))
            y = max(8, min(by, sh - ch - 48))
            self.calendar_window.geometry(f"+{x}+{y}")
        except Exception:
            pass
        self.calendar_window.deiconify()
        self.calendar_window.lift()
        self.calendar_window.focus_set()

    def _on_customer_click(self, event):
        iid = self.customer_tree.identify_row(event.y)
        if not iid:
            return
        idx = int(iid)
        self._toggle_customer_index(idx)
        self.customer_tree.selection_set(iid)
        return "break"

    def _on_customer_space(self, _event=None):
        sel = self.customer_tree.selection()
        if sel:
            self._toggle_customer_index(int(sel[0]))
        return "break"

    def _toggle_customer_index(self, idx: int):
        if idx in self.selected_customers:
            self.selected_customers.remove(idx)
        else:
            self.selected_customers.add(idx)
        self._refresh_customer_row(idx)
        self.update_counts()

    def select_all_customers(self):
        if not self.customer_names:
            return
        self.selected_customers.update(self._filtered_customer_indices)
        self._refresh_customer_tree()
        self.update_counts()

    def clear_all_customers(self):
        if not self.customer_names:
            return
        self.selected_customers.difference_update(self._filtered_customer_indices)
        self._refresh_customer_tree()
        self.update_counts()

    def reset_to_all_customers(self):
        self.customer_names = []
        self.selected_customers.clear()
        self.customer_search_var.set("")
        self._refresh_customer_tree()
        self.update_counts()
        self.log_line("已清空预设客户过滤，将导出全部客户。")

    def _refresh_customer_row(self, idx: int):
        iid = str(idx)
        if not self.customer_tree.exists(iid):
            return
        checked = idx in self.selected_customers
        pos = self._filtered_customer_indices.index(idx) if idx in self._filtered_customer_indices else idx
        zebra = "even" if pos % 2 == 0 else "odd"
        tags = (zebra,) if checked else (zebra, "unchecked")
        self.customer_tree.item(
            iid,
            image=self._cb_on if checked else self._cb_off,
            text=f" {self.customer_names[idx]}",
            tags=tags,
        )

    def _refresh_customer_tree(self):
        for iid in self.customer_tree.get_children():
            self.customer_tree.delete(iid)
        keyword = self.customer_search_var.get().strip().lower()
        self._filtered_customer_indices = [
            i for i, name in enumerate(self.customer_names) if not keyword or keyword in name.lower()
        ]
        for pos, idx in enumerate(self._filtered_customer_indices):
            checked = idx in self.selected_customers
            zebra = "even" if pos % 2 == 0 else "odd"
            tags = (zebra,) if checked else (zebra, "unchecked")
            self.customer_tree.insert(
                "",
                "end",
                iid=str(idx),
                image=self._cb_on if checked else self._cb_off,
                text=f" {self.customer_names[idx]}",
                tags=tags,
            )

    def _update_preview_structure(self):
        region_col = self.region_var.get().strip()
        if self.selected_start_date and self.selected_end_date:
            date_part = (
                self.selected_start_date
                if self.selected_start_date == self.selected_end_date
                else f"{self.selected_start_date}至{self.selected_end_date}"
            )
        else:
            date_part = "按天日期目录"
        area_part = f"分区({region_col})" if region_col else "未分区"
        self.preview_structure_var.set(f"输出结构：{date_part} / {area_part} / 客户.xlsx")

    def update_counts(self, _event=None):
        sel_fields = self.field_grid.selected_count()
        total_fields = len(self.headers)
        self.field_count.set(f"{sel_fields} / {total_fields}")

        if self.customer_names:
            self.empty_customers.place_forget()
            self.clear_cust_cfg_btn.configure(state="normal")
            self.customer_count.set(f"已选 {len(self.selected_customers)}/{len(self.customer_names)}")
            self.customer_badge.configure(bg=PALETTE["accent_soft"], fg=PALETTE["accent_ink"])
            cust_summary = f"{len(self.selected_customers)} 客户"
        else:
            self.empty_customers.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.9)
            self.clear_cust_cfg_btn.configure(state="disabled")
            self.customer_count.set("全部客户")
            self.customer_badge.configure(bg=PALETTE["teal_soft"], fg=PALETTE["teal"])
            cust_summary = "全部客户"

        if self._worker_thread is None:
            if self.headers:
                self.status_var.set(f"就绪 · {cust_summary} · {sel_fields} 字段")
                self.status_dot.itemconfigure(self._status_dot_id, fill=PALETTE["teal"])
            else:
                self.status_var.set("空库 · 待导入")
                self.status_dot.itemconfigure(self._status_dot_id, fill=PALETTE["amber"])

    def log_line(self, message: str):
        def append():
            tag = None
            if any(k in message for k in ("完成", "已打开", "已生成", "通过", "导入 ")):
                tag = "ok"
            elif any(k in message for k in ("停止", "警告", "检测", "为空")):
                tag = "warn"
            elif "失败" in message or "错误" in message or "无法" in message:
                tag = "err"
            self.log.configure(state="normal")
            self.log.insert("end", message + "\n", tag)
            self.log.see("end")
            self.log.configure(state="disabled")

        self._callbacks.put(append)

    def _drain_callbacks(self):
        try:
            for _ in range(100):
                self._callbacks.get_nowait()()
        except Empty:
            pass
        if not self._closing:
            self._after_id = self.after(30, self._drain_callbacks)

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
        self.switch_view("export")
        self.log_line(f"已打开数据库：{path}")

    def choose_database(self):
        path = filedialog.askopenfilename(title="打开客户数据库", filetypes=[("DuckDB", "*.duckdb"), ("所有文件", "*.*")])
        if path:
            self._select_database(path, existing=True)

    def create_database(self):
        path = filedialog.asksaveasfilename(
            title="新建客户数据库", initialdir=str(self.default_database_dir),
            defaultextension=".duckdb", filetypes=[("DuckDB", "*.duckdb")],
        )
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
            names = [Path(item).name for item in self.files]
            preview = names[0] if len(names) == 1 else f"{names[0]} 等 {len(names)} 个文件"
            self.file_label.configure(text=f"已选 {len(self.files)} 个：{preview}")
            self.log_line("已选择: " + ", ".join(names))

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
        if self.customer_names:
            self.selected_customers = set(range(len(self.customer_names)))
            self.log_line(f"已加载 {len(self.customer_names)} 个预设客户。")
        else:
            self.selected_customers.clear()
            self.log_line("客户名单配置为空或文件不存在，将导出全部客户。")
        self._refresh_customer_tree()
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
        self.refresh_cache_info()

    def _start_job(self, name, operation, on_success):
        if self._worker_thread is not None or self._closing:
            return
        self._cancel_event.clear()
        self.stop_button.configure(state="normal")
        for button in [self.run_button, self.import_button, *self.database_buttons]:
            button.configure(state="disabled")
        self.status_var.set(f"正在{name}…")
        self.status_dot.itemconfigure(self._status_dot_id, fill=PALETTE["accent"])
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
        self.stop_button.configure(state="disabled")
        for button in [self.run_button, self.import_button, *self.database_buttons]:
            button.configure(state="normal")
        self.refresh_cache_info()
        self.update_counts()

    def set_selected_date(self, start: str | None, end: str | None):
        self.selected_start_date = start
        self.selected_end_date = end
        if start and end:
            text = start if start == end else f"{start} 至 {end}"
            self.date_button.configure(text=f"📅 {text}")
            self.date_clear_btn.grid(row=0, column=1, padx=(self._s(4), 0))
        elif start:
            self.date_button.configure(text=f"📅 {start} ~ 待选")
            self.date_clear_btn.grid(row=0, column=1, padx=(self._s(4), 0))
        else:
            self.date_button.configure(text="📅 全部日期")
            self.date_clear_btn.grid_forget()
        self._update_preview_structure()
        self.update_counts()

    def refresh_available_dates(self):
        self.calendar.set_dates(self.store.dates() if self.task_id else [])

    def populate_fields(self):
        self.field_search_var.set("")
        self.field_grid.set_items(self.headers, select_all=True)
        values = [""] + self.headers
        self.customer_combo["values"] = values
        self.region_combo["values"] = values
        self.set_default_field(self.customer_var, self.customer_combo, ("客户名称", "客户"))
        self.set_default_field(self.region_var, self.region_combo, ("代表处", "区域", "分区"))
        if Path(self.customer_config_var.get()).exists():
            self.load_customer_config()
        else:
            self._refresh_customer_tree()
        self.refresh_available_dates()
        self._update_preview_structure()
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
        selected = self.field_grid.selected_items()
        customer = self.customer_var.get().strip()
        if not selected or not customer:
            messagebox.showwarning("字段不完整", "至少选择一个导出字段，并选择客户字段。")
            return
        if self.selected_start_date and not self.selected_end_date:
            messagebox.showwarning("日期范围不完整", "请选择结束日期；或者点击清除日期筛选后导出全部日期。")
            return
        selected_customers = [self.customer_names[i] for i in range(len(self.customer_names)) if i in self.selected_customers]
        if self.customer_names and not selected_customers:
            messagebox.showwarning("客户未选择", "请至少选择一个预设客户，或清空客户配置后导出全部客户。")
            return
        output = self.output_var.get().strip()
        region_path = self.region_path_var.get().strip()
        region = self.region_var.get().strip() or None
        start, end = self.selected_start_date, self.selected_end_date

        def export():
            return self.store.export(
                output,
                selected,
                customer,
                region,
                load_region_rules(region_path),
                start,
                end,
                selected_customers,
                self.log_line,
                self._cancel_event,
            )

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
            self.status_var.set("正在停止…")
            self.status_dot.itemconfigure(self._status_dot_id, fill=PALETTE["amber"])
            self.stop_button.configure(state="disabled")
            self.log_line("已请求停止，正在释放任务资源…")

    def refresh_cache_info(self):
        if self.store:
            path = self.store.path
            size = sum(p.stat().st_size for p in (path, Path(f"{path}.wal")) if p.is_file())
            self.cache_info_var.set(f"{path.name}   ·   {size / (1024**2):,.1f} MiB   ·   {path}")
            self.db_title_var.set(path.name)
            self.db_path_full = str(path)
            days = self.store.dates() if self.task_id else []
            if self.headers and days:
                span = days[0] if len(days) == 1 else f"{days[0]} ~ {days[-1]}"
                self.db_meta_var.set(f"{size / (1024**2):,.1f} MiB · {span}（{len(days)}天）· {len(self.headers)}列")
                self.db_dot.itemconfigure(self._db_dot_id, fill="#34D399")
            elif self.headers:
                self.db_meta_var.set(f"{size / (1024**2):,.1f} MiB · {len(self.headers)}列")
                self.db_dot.itemconfigure(self._db_dot_id, fill="#34D399")
            else:
                self.db_meta_var.set(f"{size / (1024**2):,.1f} MiB · 空白库")
                self.db_dot.itemconfigure(self._db_dot_id, fill="#FBBF24")
        else:
            self.cache_info_var.set("请选择数据库")
            self.db_title_var.set("未加载数据库")
            self.db_meta_var.set("请打开或新建库")
            self.db_path_full = ""
            self.db_dot.itemconfigure(self._db_dot_id, fill="#F87171")

    def clear_database_cache(self):
        if self.store and self._worker_thread is None:
            self._start_job(
                "整理事务日志",
                self.store.checkpoint,
                lambda _: self.log_line("已将事务日志写入持久库，临时文件随任务结束清理。"),
            )

    def destroy(self):
        self._closing = True
        if self._worker_thread is not None and self._worker_thread.is_alive():
            self.stop_and_cleanup()
            self.after(100, self.destroy)
            return
        if getattr(self, "_after_id", None) is not None:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        if self.store:
            self.store.close()
        super().destroy()


if __name__ == "__main__":
    if sys.argv[1:] == ["--self-test"]:
        from smoke_test import run

        run(SplitterApp)
    else:
        SplitterApp().mainloop()
