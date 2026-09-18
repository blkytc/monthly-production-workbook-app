"""Tkinter desktop interface for the monthly workbook generator."""

from __future__ import annotations

import datetime as dt
import re
import sys
import tempfile
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .workbook import DimensionRule, generate_workbook, read_workbook_title
from .legacy_reports import (
    ReportKind,
    SupplierChange,
    convert_xls,
    detect_report,
    report_title,
    update_legacy_report,
)


APP_NAME = "月度生产表生成器"
UNITS = ("字符", "磅", "厘米", "毫米", "英寸")
FILENAME_DATE_RANGE = re.compile(
    r"-?\d{4}-\d{2}-\d{2}-\d{4}-\d{2}-\d{2}$"
)
THEME = {
    "background": "#17191c",
    "panel": "#202327",
    "field": "#2b2f34",
    "text": "#f2f3f5",
    "muted": "#a7adb5",
    "border": "#3a3f46",
    "accent": "#2f80ed",
    "selected": "#245fba",
}


def apply_app_theme(style: ttk.Style, root: tk.Tk) -> None:
    if "clam" in style.theme_names():
        style.theme_use("clam")
    root.configure(background=THEME["background"])
    style.configure(".", background=THEME["background"], foreground=THEME["text"])
    style.configure("TFrame", background=THEME["background"])
    style.configure("TLabel", background=THEME["background"], foreground=THEME["text"])
    style.configure("TLabelframe", background=THEME["background"], foreground=THEME["text"])
    style.configure("TLabelframe.Label", background=THEME["background"], foreground=THEME["text"])
    style.configure(
        "TEntry", fieldbackground=THEME["field"], foreground=THEME["text"],
        insertcolor=THEME["text"], bordercolor=THEME["border"], padding=5,
    )
    style.configure(
        "TCombobox", fieldbackground=THEME["field"], foreground=THEME["text"],
        background=THEME["field"], arrowcolor=THEME["text"], padding=4,
    )
    style.configure(
        "TButton", background=THEME["field"], foreground=THEME["text"],
        bordercolor=THEME["border"], padding=(10, 6),
    )
    style.map(
        "TButton",
        background=[("pressed", THEME["selected"]), ("active", THEME["accent"])],
        foreground=[("disabled", THEME["muted"]), ("!disabled", THEME["text"])],
    )
    style.configure(
        "Treeview", background=THEME["panel"], fieldbackground=THEME["panel"],
        foreground=THEME["text"], bordercolor=THEME["border"], rowheight=26,
    )
    style.configure(
        "Treeview.Heading", background=THEME["field"], foreground=THEME["text"],
        relief="flat", padding=(6, 5),
    )
    style.map("Treeview", background=[("selected", THEME["selected"])], foreground=[("selected", "#ffffff")])
    style.map("TCombobox", fieldbackground=[("readonly", THEME["field"])], foreground=[("readonly", THEME["text"])])


def parse_user_date(text: str) -> dt.date:
    try:
        parts = text.strip().replace("/", "-").split("-")
        if len(parts) != 3 or len(parts[0]) != 4:
            raise ValueError
        return dt.date(*(int(part) for part in parts))
    except (TypeError, ValueError):
        raise ValueError("日期格式应为年/月/日，例如 2026/10/25") from None


def default_output_path(source: Path, start: dt.date, end: dt.date) -> Path:
    filename = f"生产、入库、发运表{start.isoformat()}-{end.isoformat()}.xlsx"
    return source.parent / filename


def normalize_filename(value: str) -> str:
    name = value.strip()
    if not name or re.search(r'[<>:"/\\|?*]', name) or name.endswith((" ", ".")):
        raise ValueError("文件名称不能为空，也不能包含 < > : \" / \\ | ? *")
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


def update_filename_dates(filename: str, start: dt.date, end: dt.date) -> str:
    suffix = filename[-5:] if filename.lower().endswith(".xlsx") else ""
    stem = filename[:-5] if suffix else filename
    date_range = f"{start.isoformat()}-{end.isoformat()}"
    if FILENAME_DATE_RANGE.search(stem):
        stem = FILENAME_DATE_RANGE.sub(f"-{date_range}", stem)
    else:
        stem = f"{stem}-{date_range}"
    return f"{stem}{suffix}"


def resource_path(name: str) -> Path:
    root = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    candidate = root / name
    return candidate if candidate.exists() else Path(__file__).parent / name


class RuleDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, kind: str):
        super().__init__(parent)
        self.result: DimensionRule | None = None
        self.kind = kind
        self.title("添加行高规则" if kind == "row" else "添加列宽规则")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        prompt = "行号或范围" if kind == "row" else "列名或范围"
        example = "例如：4-8,12,15-18" if kind == "row" else "例如：A:C,F,H:J"
        ttk.Label(self, text=prompt).grid(row=0, column=0, padx=12, pady=(14, 5), sticky="w")
        self.selection = ttk.Entry(self, width=28)
        self.selection.grid(row=1, column=0, columnspan=2, padx=12, sticky="ew")
        ttk.Label(self, text=example, foreground=THEME["muted"]).grid(row=2, column=0, columnspan=2, padx=12, pady=(3, 10), sticky="w")

        ttk.Label(self, text="尺寸").grid(row=3, column=0, padx=12, sticky="w")
        ttk.Label(self, text="单位").grid(row=3, column=1, padx=12, sticky="w")
        self.value = ttk.Entry(self, width=14)
        self.value.grid(row=4, column=0, padx=12, pady=(5, 14), sticky="ew")
        self.unit = ttk.Combobox(self, values=UNITS, state="readonly", width=10)
        self.unit.set("磅" if kind == "row" else "字符")
        self.unit.grid(row=4, column=1, padx=12, pady=(5, 14), sticky="ew")

        buttons = ttk.Frame(self)
        buttons.grid(row=5, column=0, columnspan=2, padx=12, pady=(0, 14), sticky="e")
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="确定", command=self._accept).pack(side="right")
        self.bind("<Return>", lambda _event: self._accept())
        self.bind("<Escape>", lambda _event: self.destroy())
        self.selection.focus_set()

    def _accept(self) -> None:
        try:
            value = float(self.value.get())
            if value <= 0:
                raise ValueError
            selection = self.selection.get().strip()
            if not selection:
                raise ValueError
        except ValueError:
            messagebox.showerror("输入错误", "请填写有效范围和大于 0 的尺寸。", parent=self)
            return
        self.result = DimensionRule(selection, value, self.unit.get())
        self.destroy()


class SupplierDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, kind: ReportKind):
        super().__init__(parent)
        self.result: SupplierChange | None = None
        self.title("添加供应商操作")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        areas = {
            ReportKind.CHEMICAL_USAGE: ("供应商",),
            ReportKind.DISPATCH_INFO: ("外购供应商",),
            ReportKind.INVENTORY_PLAN: ("生产日累", "当日配送量"),
        }[kind]
        ttk.Label(self, text="操作").grid(row=0, column=0, padx=12, pady=(14, 5), sticky="w")
        ttk.Label(self, text="区域").grid(row=0, column=1, padx=12, pady=(14, 5), sticky="w")
        self.action = ttk.Combobox(self, values=("新增供应商", "修改名称", "删除供应商"), state="readonly", width=14)
        self.action.set("新增供应商")
        self.action.grid(row=1, column=0, padx=12, sticky="ew")
        self.area = ttk.Combobox(self, values=areas, state="readonly", width=16)
        self.area.set(areas[0])
        self.area.grid(row=1, column=1, padx=12, sticky="ew")
        ttk.Label(self, text="原供应商名称（新增时留空）").grid(row=2, column=0, columnspan=2, padx=12, pady=(12, 5), sticky="w")
        self.old_name = ttk.Entry(self, width=42)
        self.old_name.grid(row=3, column=0, columnspan=2, padx=12, sticky="ew")
        ttk.Label(self, text="新供应商名称").grid(row=4, column=0, columnspan=2, padx=12, pady=(12, 5), sticky="w")
        self.new_name = ttk.Entry(self, width=42)
        self.new_name.grid(row=5, column=0, columnspan=2, padx=12, sticky="ew")
        buttons = ttk.Frame(self)
        buttons.grid(row=6, column=0, columnspan=2, padx=12, pady=14, sticky="e")
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="确定", command=self._accept).pack(side="right")
        self.new_name.focus_set()

    def _accept(self) -> None:
        action = {"修改名称": "rename", "新增供应商": "add", "删除供应商": "delete"}[self.action.get()]
        old_name = self.old_name.get().strip()
        new_name = self.new_name.get().strip()
        if action == "add" and not new_name or action in ("rename", "delete") and not old_name or action == "rename" and not new_name:
            messagebox.showerror("输入错误", "请填写该操作需要的供应商名称。", parent=self)
            return
        self.result = SupplierChange(action, self.area.get(), old_name, new_name)
        self.destroy()


class WorkbookApp(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=18)
        master.title(APP_NAME)
        master.minsize(820, 760)
        self.pack(fill="both", expand=True)
        self.source = tk.StringVar()
        self.title_text = tk.StringVar()
        today = dt.date.today()
        self.start = tk.StringVar(value=today.strftime("%Y/%m/%d"))
        self.end = tk.StringVar(value=(today + dt.timedelta(days=30)).strftime("%Y/%m/%d"))
        self.output_dir = tk.StringVar()
        self.filename = tk.StringVar()
        self.row_rules: list[DimensionRule] = []
        self.column_rules: list[DimensionRule] = []
        self.supplier_changes: list[SupplierChange] = []
        self.report_kind: ReportKind | None = None
        self.converted_source: Path | None = None
        self.temp_directory = tempfile.TemporaryDirectory(prefix="monthly-workbook-")
        self.status = tk.StringVar(value="请选择 Excel 模板。")
        self.supplier_status = tk.StringVar(value="待执行操作：0 项")
        self._build()
        self.start.trace_add("write", self._dates_changed)
        self.end.trace_add("write", self._dates_changed)

    def _build(self) -> None:
        self.columnconfigure(1, weight=1)
        ttk.Label(self, text=APP_NAME, font=("TkDefaultFont", 18, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))
        self._file_row(1, "Excel 模板", self.source, self._choose_source)
        self._entry_row(2, "表格标题", self.title_text)

        date_frame = ttk.Frame(self)
        date_frame.grid(row=3, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Label(self, text="起始日期").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(date_frame, textvariable=self.start, width=18).pack(side="left")
        ttk.Label(date_frame, text="结束日期", padding=(18, 0, 8, 0)).pack(side="left")
        ttk.Entry(date_frame, textvariable=self.end, width=18).pack(side="left")
        ttk.Label(date_frame, text="年/月/日", foreground=THEME["muted"], padding=(8, 0)).pack(side="left")

        rules = ttk.Frame(self)
        rules.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=(12, 8))
        rules.columnconfigure((0, 1), weight=1)
        self.row_tree = self._rule_panel(rules, 0, "行高设置", "row")
        self.column_tree = self._rule_panel(rules, 1, "列宽设置", "column")

        supplier = ttk.LabelFrame(self, text="供应商操作（可选）", padding=8)
        supplier.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(4, 8))
        self.supplier_tree = ttk.Treeview(supplier, columns=("action", "area", "old", "new"), show="headings", height=4)
        for key, text_value, width in (("action", "操作", 90), ("area", "区域", 120), ("old", "原名称", 190), ("new", "新名称", 190)):
            self.supplier_tree.heading(key, text=text_value)
            self.supplier_tree.column(key, width=width)
        self.supplier_tree.pack(side="left", fill="both", expand=True)
        supplier_buttons = ttk.Frame(supplier)
        supplier_buttons.pack(side="right", fill="y", padx=(8, 0))
        ttk.Button(supplier_buttons, text="添加", command=self._add_supplier_change).pack(fill="x")
        ttk.Button(supplier_buttons, text="删除", command=self._delete_supplier_change).pack(fill="x", pady=(6, 0))
        ttk.Label(supplier_buttons, textvariable=self.supplier_status).pack(pady=(12, 0))

        self._file_row(6, "保存文件夹", self.output_dir, self._choose_output_dir)
        self._entry_row(7, "文件名称", self.filename)
        footer = ttk.Frame(self)
        footer.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(14, 0))
        ttk.Label(footer, textvariable=self.status).pack(side="left")
        ttk.Button(footer, text="帮助", command=self._open_help).pack(side="right")
        ttk.Button(footer, text="生成工作表", command=self._generate).pack(side="right", padx=(0, 8))

    def _entry_row(self, row: int, label: str, variable: tk.StringVar) -> None:
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(self, textvariable=variable).grid(row=row, column=1, columnspan=2, sticky="ew", pady=5)

    def _file_row(self, row: int, label: str, variable: tk.StringVar, command) -> None:
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(self, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=(0, 8), pady=5)
        ttk.Button(self, text="浏览…", command=command).grid(row=row, column=2, sticky="e", pady=5)

    def _rule_panel(self, parent: ttk.Frame, column: int, title: str, kind: str) -> ttk.Treeview:
        frame = ttk.LabelFrame(parent, text=title, padding=8)
        frame.grid(row=0, column=column, sticky="nsew", padx=(0, 6) if column == 0 else (6, 0))
        tree = ttk.Treeview(frame, columns=("range", "size"), show="headings", height=8)
        tree.heading("range", text="范围")
        tree.heading("size", text="尺寸")
        tree.column("range", width=150)
        tree.column("size", width=110)
        tree.pack(fill="both", expand=True)
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x", pady=(8, 0))
        ttk.Button(toolbar, text="添加", command=lambda: self._add_rule(kind, tree)).pack(side="left")
        ttk.Button(toolbar, text="删除", command=lambda: self._delete_rule(kind, tree)).pack(side="left", padx=(6, 0))
        return tree

    def _add_rule(self, kind: str, tree: ttk.Treeview) -> None:
        dialog = RuleDialog(self, kind)
        self.wait_window(dialog)
        if dialog.result is None:
            return
        target = self.row_rules if kind == "row" else self.column_rules
        target.append(dialog.result)
        tree.insert("", "end", values=(dialog.result.selection, f"{dialog.result.value:g} {dialog.result.unit}"))

    def _delete_rule(self, kind: str, tree: ttk.Treeview) -> None:
        selected = tree.selection()
        if not selected:
            return
        indexes = sorted((tree.index(item) for item in selected), reverse=True)
        target = self.row_rules if kind == "row" else self.column_rules
        for item, index in zip(reversed(selected), indexes):
            tree.delete(item)
            target.pop(index)

    def _add_supplier_change(self) -> None:
        if self.report_kind is None:
            messagebox.showinfo("供应商操作", "当前生产、入库、发运表没有供应商列表设置。", parent=self)
            return
        dialog = SupplierDialog(self, self.report_kind)
        self.wait_window(dialog)
        if dialog.result is None:
            return
        change = dialog.result
        self.supplier_changes.append(change)
        action = {"rename": "修改名称", "add": "新增供应商", "delete": "删除供应商"}[change.action]
        self.supplier_tree.insert("", "end", values=(action, change.area, change.old_name, change.new_name))
        self.supplier_status.set(f"待执行操作：{len(self.supplier_changes)} 项")

    def _delete_supplier_change(self) -> None:
        selected = list(self.supplier_tree.selection())
        for item in sorted(selected, key=self.supplier_tree.index, reverse=True):
            index = self.supplier_tree.index(item)
            self.supplier_tree.delete(item)
            self.supplier_changes.pop(index)
        self.supplier_status.set(f"待执行操作：{len(self.supplier_changes)} 项")

    def _choose_source(self) -> None:
        value = filedialog.askopenfilename(title="选择 Excel 模板", filetypes=(("Excel 工作簿", "*.xlsx *.xls"),))
        if not value:
            return
        try:
            selected = Path(value)
            working = convert_xls(selected, Path(self.temp_directory.name)) if selected.suffix.lower() == ".xls" else selected
            try:
                kind = detect_report(working)
            except ValueError:
                kind = None
            title = report_title(working, kind) if kind else read_workbook_title(working)
        except Exception as error:
            messagebox.showerror("无法读取模板", str(error), parent=self)
            return
        self.source.set(value)
        self.converted_source = working
        self.report_kind = kind
        self.supplier_changes.clear()
        self.supplier_status.set("待执行操作：0 项")
        for item in self.supplier_tree.get_children():
            self.supplier_tree.delete(item)
        self.title_text.set(title)
        self._refresh_output()
        report_name = kind.value if kind else "生产、入库、发运表"
        self.status.set(f"已识别：{report_name}")

    def _refresh_output(self) -> None:
        try:
            source = Path(self.source.get())
            start, end = parse_user_date(self.start.get()), parse_user_date(self.end.get())
            if self.report_kind:
                output = source.with_name(
                    update_filename_dates(source.with_suffix(".xlsx").name, start, end)
                )
            else:
                output = default_output_path(source, start, end)
            self.output_dir.set(str(output.parent))
            self.filename.set(output.name)
        except ValueError:
            pass

    def _dates_changed(self, *_args) -> None:
        if not self.source.get():
            return
        try:
            start = parse_user_date(self.start.get())
            end = parse_user_date(self.end.get())
            if end < start:
                return
        except ValueError:
            return
        if self.filename.get().strip():
            self.filename.set(update_filename_dates(self.filename.get(), start, end))
        else:
            self._refresh_output()

    def _choose_output_dir(self) -> None:
        value = filedialog.askdirectory(title="选择保存文件夹", initialdir=self.output_dir.get() or None)
        if value:
            self.output_dir.set(value)

    def _generate(self) -> None:
        try:
            start, end = parse_user_date(self.start.get()), parse_user_date(self.end.get())
            if not self.output_dir.get().strip() or not self.filename.get().strip():
                self._refresh_output()
            output = Path(self.output_dir.get()) / normalize_filename(self.filename.get())
            if self.report_kind:
                update_legacy_report(
                    self.converted_source or self.source.get(), output, start, end,
                    self.supplier_changes, self.title_text.get(), self.row_rules, self.column_rules,
                )
                sheet_count = (end - start).days + 1
            else:
                result = generate_workbook(
                    self.source.get(), output, start, end, self.title_text.get(),
                    self.row_rules, self.column_rules,
                )
                sheet_count = result.sheet_count
        except Exception as error:
            messagebox.showerror("生成失败", str(error), parent=self)
            self.status.set("生成失败，请检查输入。")
            return
        self.status.set(f"已生成 {sheet_count} 天的工作表。")
        supplier_message = f"\n供应商操作：已执行 {len(self.supplier_changes)} 项" if self.report_kind else ""
        messagebox.showinfo("生成完成", f"文件已保存到：\n{output}{supplier_message}", parent=self)

    def _open_help(self) -> None:
        webbrowser.open(resource_path("help.html").as_uri())


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    apply_app_theme(style, root)
    WorkbookApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
