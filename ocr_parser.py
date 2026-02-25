# -*- coding: utf-8 -*-
"""
仅保留从「客户型号」提取客户编号的逻辑，供 file_ops 使用。
表格数据改为从 Excel 读取，见 excel_parser.py。
"""
from __future__ import annotations

import re


def extract_customer_id(customer_model: str) -> str:
    """
    从「客户型号」字符串中提取客户编号（破折号前的数字部分）。
    例如: "1116030298 - BG-G215V21500-ACLVBO-V002" -> "1116030298"
    """
    if not customer_model or not isinstance(customer_model, str):
        return ""
    s = customer_model.strip()
    if " - " in s:
        s = s.split(" - ", 1)[0].strip()
    m = re.match(r"^(\d+)", s)
    return m.group(1) if m else s
