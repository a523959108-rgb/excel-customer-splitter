"""Exercise the desktop entry points, including in a frozen Windows executable."""

import time
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook, load_workbook


def run(app_type):
    def settle(app):
        deadline = time.monotonic() + 180
        while app._worker_thread is not None:
            if time.monotonic() > deadline:
                raise TimeoutError("Desktop background task timed out")
            app.update()
            time.sleep(0.02)
        if errors:
            raise RuntimeError(errors)

    with TemporaryDirectory() as tmp:
        root = Path(tmp)
        profile = root / "profile"
        default_db_dir = root / "客户 数据目录"
        default_db_dir.mkdir()
        (root / "database-directory.txt").write_text(str(default_db_dir) + "\r\n", encoding="utf-8-sig")
        path = root / "测试库.duckdb"
        errors = []
        with (
            patch("tkinter.messagebox.showerror", side_effect=lambda *args: errors.append(args)),
            patch("tkinter.messagebox.showwarning", side_effect=lambda *args: errors.append(args)),
            patch("tkinter.messagebox.showinfo"),
        ):
            app = app_type(profile, install_dir=root)
            try:
                app.withdraw()
                if app.store.path != (default_db_dir / "splitter.duckdb").resolve():
                    raise ValueError("First launch did not use the installer's database directory")
                with patch("tkinter.filedialog.asksaveasfilename", return_value=str(path)) as save:
                    app.create_database()
                    if Path(save.call_args.kwargs["initialdir"]) != default_db_dir:
                        raise ValueError("New database dialog did not use the installed default")
                if not app.import_panel.winfo_manager():
                    raise ValueError("New database should show Excel import")
                for i, records in enumerate(
                    ([(1, "甲", 10), (2, "旧", 20)], [(2, "乙", 21), (2, "乙", 21), (3, "甲", 30)])
                ):
                    book = Workbook()
                    ws = book.active
                    ws.append(["年", "月", "日", "客户", "金额"])
                    for d, c, amount in records:
                        ws.append([2025, 1, d, c, amount])
                    source = root / f"source{i}.xlsx"
                    book.save(source)
                    book.close()
                    if i:
                        app.import_toggle.invoke()
                    if not app.import_panel.winfo_manager():
                        raise ValueError("Excel file selector is hidden")
                    with patch("tkinter.filedialog.askopenfilenames", return_value=[str(source)]):
                        app.choose_files()
                    app.import_button.invoke()
                    settle(app)
            finally:
                app.destroy()
            app = app_type(profile, install_dir=root)
            try:
                app.withdraw()
                if app.store.path != path.resolve() or len(app.calendar.available) != 3:
                    raise ValueError("Saved database or calendar did not reopen")
                if app.import_panel.winfo_manager():
                    raise ValueError("Existing database should open in the export view")
                app.output_var.set(str(root / "output"))
                config = root / "customers.json"
                config.write_text('{"customers": ["甲", "乙"]}', encoding="utf-8")
                with patch("tkinter.filedialog.askopenfilename", return_value=str(config)):
                    app.choose_customer_config()
                app.run_button.invoke()
                settle(app)
                contents = []
                for file in (root / "output").rglob("*.xlsx"):
                    book = load_workbook(file, read_only=True, data_only=True)
                    try:
                        contents.extend(list(book.active.values)[1:])
                    finally:
                        book.close()
                if sorted(contents) != [
                    (2025, 1, 1, "甲", 10),
                    (2025, 1, 2, "乙", 21),
                    (2025, 1, 2, "乙", 21),
                    (2025, 1, 3, "甲", 30),
                ]:
                    raise ValueError("Desktop export round-trip failed")
                # Searching limits the controls on screen. Hidden selections still export.
                app.customer_search_var.set("甲")
                app.clear_all_customers()
                app.field_search_var.set("年")
                app.field_grid.clear_visible()
                app.field_search_var.set("客户")
                filtered = root / "filtered"
                app.output_var.set(str(filtered))
                app.run_button.invoke()
                settle(app)
                paths = list(filtered.rglob("*.xlsx"))
                if len(paths) != 1:
                    raise ValueError("Filtered export did not preserve customer selection")
                book = load_workbook(paths[0], read_only=True, data_only=True)
                try:
                    if list(book.active.values) != [
                        ("月", "日", "客户", "金额"), (1, 2, "乙", 21), (1, 2, "乙", 21)
                    ]:
                        raise ValueError("Search changed hidden selections or lost duplicate rows")
                finally:
                    book.close()
                app.field_search_var.set("")
                with patch("tkinter.filedialog.asksaveasfilename", return_value=str(root / "empty.duckdb")):
                    app.create_database()
                if app.field_grid.visible_indices:
                    raise ValueError("Empty database retained fields from the previous database")
                app.field_grid._relayout()
            finally:
                app.destroy()
