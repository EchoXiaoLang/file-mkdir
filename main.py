# -*- coding: utf-8 -*-
"""
表格图片解析 -> 产品编号/客户型号 -> 可纠错 -> 按产品编号建目录，按客户编号复制文件到 YG。
支持 GUI（main.py）和命令行测试（main.py --cli）。
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import List, Optional


def _run_cli() -> None:
    """命令行模式：不依赖 tkinter，从 Excel 读取表格后建目录、复制文件。"""
    from excel_parser import parse_excel
    from file_ops import run_pipeline

    print("=== 命令行测试模式（--cli）===\n")
    excel_path = input("Excel 文件路径（含「产品编号」「客户型号」两列）: ").strip()
    if not excel_path or not Path(excel_path).exists():
        print("文件不存在，退出。")
        return
    print("读取中...")
    try:
        data = parse_excel(excel_path)
    except Exception as e:
        print(f"读取失败: {e}")
        return
    print(f"解析到 {len(data)} 行:")
    for i, row in enumerate(data, 1):
        print(f"  {i}. 产品编号={row.get('产品编号')!r}  客户型号={row.get('客户型号')!r}")

    out_root = input("\n生成目录位置（如 /tmp/测试）: ").strip()
    source_dir = input("客户文件所在目录（可回车跳过复制）: ").strip()
    if not out_root:
        print("未输入生成目录，退出。")
        return

    logs: List[str] = []

    def log(msg: str) -> None:
        logs.append(msg)
        print(msg)

    run_pipeline(Path(out_root), Path(source_dir) if source_dir else Path(out_root), data, log)
    print("\n完成。")


if __name__ == "__main__" and ("--cli" in sys.argv or "-c" in sys.argv):
    _run_cli()
    sys.exit(0)

# GUI 依赖
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from excel_parser import parse_excel
from file_ops import run_pipeline


class EditableTable(ttk.Frame):
    """可编辑的表格：双击单元格弹出编辑框。"""

    def __init__(self, parent, columns: List[str], **kwargs):
        super().__init__(parent, **kwargs)
        self.columns = columns
        self._data: List[dict] = []
        self._build_ui()

    def _build_ui(self):
        self.tree = ttk.Treeview(self, columns=self.columns, show="headings", height=12)
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=180)
        vsb = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self._on_double_click)

    def _on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id or not col:
            return
        col_idx = int(col.replace("#", "")) - 1
        if col_idx < 0 or col_idx >= len(self.columns):
            return
        col_name = self.columns[col_idx]
        item = self.tree.item(row_id)
        values = list(item["values"])
        if col_idx >= len(values):
            values.extend([""] * (col_idx - len(values) + 1))
        current = values[col_idx] if col_idx < len(values) else ""

        # 弹窗编辑
        win = tk.Toplevel(self)
        win.title(f"编辑 - {col_name}")
        win.geometry("400x80")
        win.transient(self.winfo_toplevel())
        entry = ttk.Entry(win, width=50)
        entry.pack(fill=tk.X, padx=10, pady=10)
        entry.insert(0, str(current))
        entry.focus_set()

        def ok():
            new_val = entry.get().strip()
            values[col_idx] = new_val
            self.tree.item(row_id, values=values)
            idx = self.tree.index(row_id)
            if 0 <= idx < len(self._data):
                self._data[idx][col_name] = new_val
            win.destroy()

        def cancel():
            win.destroy()

        entry.bind("<Return>", lambda e: ok())
        ttk.Button(win, text="确定", command=ok).pack(side=tk.RIGHT, padx=5, pady=5)
        ttk.Button(win, text="取消", command=cancel).pack(side=tk.RIGHT, padx=5, pady=5)

    def set_data(self, data: List[dict]) -> None:
        self._data = [dict(d) for d in data]
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self._data:
            values = [row.get(c, "") for c in self.columns]
            self.tree.insert("", tk.END, values=values)

    def get_data(self) -> List[dict]:
        out = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            out.append({c: (values[i] if i < len(values) else "") for i, c in enumerate(self.columns)})
        self._data = out
        return out


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel 表格 - 产品编号/客户型号 → 建目录并复制客户文件")
        self.root.minsize(700, 500)
        self._data: List[dict] = []
        self._out_root: Optional[Path] = None
        self._source_dir: Optional[Path] = None
        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # 第一行：选择 Excel
        row0 = ttk.Frame(main)
        row0.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(row0, text="选择 Excel 表格", command=self._select_excel).pack(side=tk.LEFT, padx=(0, 5))
        self.file_label = ttk.Label(row0, text="未选择文件", foreground="gray")
        self.file_label.pack(side=tk.LEFT)

        # 表格（可编辑）
        table_frame = ttk.LabelFrame(main, text="解析结果（双击单元格可修改）", padding=5)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.table = EditableTable(table_frame, columns=["产品编号", "客户型号"])
        self.table.pack(fill=tk.BOTH, expand=True)

        # 目录选择
        row1 = ttk.Frame(main)
        row1.pack(fill=tk.X, pady=5)
        ttk.Button(row1, text="选择生成目录位置（如 D:/测试/）", command=self._select_output).pack(side=tk.LEFT, padx=(0, 5))
        self.out_label = ttk.Label(row1, text="未选择", foreground="gray")
        self.out_label.pack(side=tk.LEFT)

        row2 = ttk.Frame(main)
        row2.pack(fill=tk.X, pady=5)
        ttk.Button(row2, text="选择客户文件所在目录", command=self._select_source).pack(side=tk.LEFT, padx=(0, 5))
        self.src_label = ttk.Label(row2, text="未选择", foreground="gray")
        self.src_label.pack(side=tk.LEFT)

        # 执行
        row3 = ttk.Frame(main)
        row3.pack(fill=tk.X, pady=5)
        ttk.Button(row3, text="开始执行", command=self._run).pack(side=tk.LEFT, padx=(0, 5))

        # 日志
        log_frame = ttk.LabelFrame(main, text="日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD)
        vsb = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=vsb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def _log(self, msg: str) -> None:
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)

    def _select_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Excel 表格（含「产品编号」「客户型号」列）",
            filetypes=[("Excel", "*.xlsx *.xls"), ("全部", "*.*")],
        )
        if not path:
            return
        self.file_label.config(text=Path(path).name, foreground="black")
        self._log(f"正在读取: {path}")
        self.root.update()
        try:
            data = parse_excel(path)
            self._data = data
            self.table.set_data(data)
            self._log(f"读取到 {len(data)} 行。请核对并可在表格中双击修改。")
        except Exception as e:
            self._log(f"读取失败: {e}")
            messagebox.showerror("读取失败", str(e))

    def _select_output(self) -> None:
        path = filedialog.askdirectory(title="选择生成目录的根位置（如 D:/测试/）")
        if path:
            self._out_root = Path(path)
            self.out_label.config(text=str(self._out_root), foreground="black")

    def _select_source(self) -> None:
        path = filedialog.askdirectory(title="选择客户文件所在目录")
        if path:
            self._source_dir = Path(path)
            self.src_label.config(text=str(self._source_dir), foreground="black")

    def _run(self) -> None:
        data = self.table.get_data()
        if not data:
            messagebox.showwarning("提示", "请先选择 Excel 并读取数据，或手动在表格中添加行。")
            return
        if not self._out_root or not self._out_root.exists():
            messagebox.showwarning("提示", "请选择「生成目录位置」。")
            return
        if not self._source_dir or not self._source_dir.exists():
            messagebox.showwarning("提示", "请选择「客户文件所在目录」。")
            return
        self._log("---------- 开始执行 ----------")
        try:
            run_pipeline(self._out_root, self._source_dir, data, self._log)
            self._log("---------- 执行结束 ----------")
            messagebox.showinfo("完成", "执行完成，请查看日志。")
        except Exception as e:
            self._log(f"[异常] {e}")
            messagebox.showerror("错误", str(e))


def main():
    app = App()
    app.root.mainloop()


if __name__ == "__main__":
    main()
