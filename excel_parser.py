# -*- coding: utf-8 -*-
"""
从 Excel 中读取「产品编号」「客户型号」两列。
"""
from __future__ import annotations

from pathlib import Path
from typing import List

try:
    import openpyxl
except ImportError:
    openpyxl = None  # type: ignore


COL_PRODUCT = "产品编号"
COL_CUSTOMER = "客户型号"


def parse_excel(path: str | Path) -> List[dict]:
    """
    读取 Excel 第一个工作表，按表头找到「产品编号」「客户型号」两列，返回行数据。
    返回 [{"产品编号": str, "客户型号": str}, ...]。
    """
    if openpyxl is None:
        raise RuntimeError("请先安装: pip install openpyxl")
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {path}")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    if ws is None:
        wb.close()
        return []
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    header = [str(c or "").strip() for c in rows[0]]
    col_product = _col_index(header, COL_PRODUCT)
    col_customer = _col_index(header, COL_CUSTOMER)
    if col_product < 0:
        col_product = 0
    if col_customer < 0:
        col_customer = 1
    result = []
    for row in rows[1:]:
        vals = list(row) if row else []
        product = str(vals[col_product] or "").strip() if col_product < len(vals) else ""
        customer = str(vals[col_customer] or "").strip() if col_customer < len(vals) else ""
        if product or customer:
            result.append({"产品编号": product, "客户型号": customer})
    return result


def _col_index(header: List[str], keyword: str) -> int:
    for i, h in enumerate(header):
        if keyword in (h or ""):
            return i
    return -1
