# -*- coding: utf-8 -*-
"""
按产品编号创建目录，并按客户编号从源目录复制「名称包含客户编号」的压缩文件与 PDF 到 产品编号/YG。
"""
from __future__ import annotations

import os
import shutil
from pathlib import Path
from typing import Callable, List

from ocr_parser import extract_customer_id

# 视为压缩文件的扩展名（客户型号破折号前数字在文件名中即匹配）
ARCHIVE_EXTENSIONS = (".zip", ".rar", ".7z", ".tar.gz", ".tgz", ".tar")


def ensure_product_folders(
    root: Path,
    product_ids: List[str],
    log: Callable[[str], None],
) -> None:
    """
    在 root 下为每个 product_id 创建文件夹，并在其下创建 YG 子文件夹。
    若文件夹已存在则跳过并记录日志。
    """
    root = Path(root)
    for pid in product_ids:
        pid = (pid or "").strip()
        if not pid:
            continue
        folder = root / pid
        yg = folder / "YG"
        if folder.exists():
            log(f"[跳过] 目录已存在: {folder}")
        else:
            folder.mkdir(parents=True, exist_ok=True)
            log(f"[创建] {folder}")
        if yg.exists():
            log(f"[跳过] 目录已存在: {yg}")
        else:
            yg.mkdir(parents=True, exist_ok=True)
            log(f"[创建] {yg}")


def _is_archive(path: Path) -> bool:
    """是否为支持的压缩文件。"""
    name = path.name.lower()
    return any(name.endswith(ext) for ext in ARCHIVE_EXTENSIONS)


def _find_matching_archives(source_dir: Path, customer_id: str) -> List[Path]:
    """
    递归遍历 source_dir 及其子目录，收集名称包含 customer_id 的压缩文件路径。
    """
    source_dir = Path(source_dir).resolve()
    matching: List[Path] = []
    for root, _, files in os.walk(source_dir):
        root = Path(root)
        for f in files:
            if customer_id not in f:
                continue
            p = (root / f).resolve()
            if p.is_file() and _is_archive(p):
                matching.append(p)
    return matching


def _find_matching_pdfs(source_dir: Path, customer_id: str) -> List[Path]:
    """
    递归遍历 source_dir 及其子目录，收集名称包含 customer_id 的 PDF 文件路径。
    """
    source_dir = Path(source_dir).resolve()
    matching: List[Path] = []
    for root, _, files in os.walk(source_dir):
        root = Path(root)
        for f in files:
            if customer_id not in f or not f.lower().endswith(".pdf"):
                continue
            p = (root / f).resolve()
            if p.is_file():
                matching.append(p)
    return matching


def _yg_has_customer_archive_and_pdf(yg_dir: Path, customer_id: str) -> bool:
    """
    检查 YG 目录下是否已同时存在名称包含 customer_id 的压缩文件和 PDF。若都有则返回 True（可跳过复制）。
    """
    yg_dir = Path(yg_dir)
    if not yg_dir.is_dir():
        return False
    has_archive = False
    has_pdf = False
    for item in yg_dir.iterdir():
        if not item.is_file() or customer_id not in item.name:
            continue
        if _is_archive(item):
            has_archive = True
        elif item.suffix.lower() == ".pdf":
            has_pdf = True
        if has_archive and has_pdf:
            return True
    return has_archive and has_pdf


def copy_customer_files_to_yg(
    source_dir: Path,
    dest_yg_dir: Path,
    customer_id: str,
    log: Callable[[str], None],
) -> None:
    """
    在 source_dir 及其子目录中递归查找名称包含 customer_id 的压缩文件与 PDF，
    复制到 dest_yg_dir。若 YG 下已同时存在该客户的压缩文件和 PDF 则跳过。
    """
    source_dir = Path(source_dir).resolve()
    dest_yg_dir = Path(dest_yg_dir)
    customer_id = (customer_id or "").strip()
    if not customer_id:
        log(f"[跳过] 客户编号为空，目标: {dest_yg_dir}")
        return

    if not source_dir.exists():
        log(f"[错误] 源目录不存在: {source_dir}")
        return

    if _yg_has_customer_archive_and_pdf(dest_yg_dir, customer_id):
        log(f"[跳过] {dest_yg_dir} 下已存在客户压缩文件与客户PDF（客户编号 {customer_id}）")
        return

    archives = _find_matching_archives(source_dir, customer_id)
    pdfs = _find_matching_pdfs(source_dir, customer_id)
    copied_any = False
    for item in archives:
        if not item.is_file():
            continue
        dest = dest_yg_dir / item.name
        try:
            shutil.copy2(item, dest)
            log(f"[复制] 压缩文件 -> {dest_yg_dir.name}: {item.name} （来自 {item.parent}）")
            copied_any = True
        except Exception as e:
            log(f"[错误] 复制失败 {item}: {e}")
    for item in pdfs:
        if not item.is_file():
            continue
        dest = dest_yg_dir / item.name
        try:
            shutil.copy2(item, dest)
            log(f"[复制] PDF -> {dest_yg_dir.name}: {item.name} （来自 {item.parent}）")
            copied_any = True
        except Exception as e:
            log(f"[错误] 复制失败 {item}: {e}")

    if not copied_any:
        log(f"[无匹配] 客户编号 {customer_id} 在源目录及子目录中未找到对应压缩文件或PDF")


def run_pipeline(
    root: Path,
    source_dir: Path,
    rows: List[dict],
    log: Callable[[str], None],
) -> None:
    """
    根据解析结果执行：创建 产品编号/YG，再按每行的客户编号从 source_dir 复制到对应 YG。
    rows: [{"产品编号": str, "客户型号": str}, ...]
    """
    root = Path(root)
    source_dir = Path(source_dir)
    product_ids = list({(r.get("产品编号") or "").strip() for r in rows if (r.get("产品编号") or "").strip()})
    ensure_product_folders(root, product_ids, log)

    for r in rows:
        product_id = (r.get("产品编号") or "").strip()
        customer_model = r.get("客户型号") or ""
        customer_id = extract_customer_id(customer_model)
        if not product_id:
            log(f"[跳过] 产品编号为空，客户型号: {customer_model}")
            continue
        yg_dir = root / product_id / "YG"
        copy_customer_files_to_yg(source_dir, yg_dir, customer_id, log)
