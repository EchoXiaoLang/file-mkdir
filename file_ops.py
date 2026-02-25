# -*- coding: utf-8 -*-
"""
按产品编号创建目录，并按客户编号从源目录复制文件夹/PDF 到 产品编号/YG。
"""
from __future__ import annotations

import os
import shutil
from pathlib import Path
from typing import Callable, List, Tuple

from ocr_parser import extract_customer_id


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


def _find_matching_dirs_and_pdfs(
    source_dir: Path,
    customer_id: str,
) -> Tuple[List[Path], List[Path]]:
    """
    递归遍历 source_dir 及其子目录，收集名称包含 customer_id 的文件夹和 PDF 文件路径。
    返回 (匹配的文件夹列表, 匹配的 PDF 列表)。
    """
    source_dir = Path(source_dir).resolve()
    matching_dirs: List[Path] = []
    matching_pdfs: List[Path] = []

    for root, dirs, files in os.walk(source_dir):
        root = Path(root)
        for d in dirs:
            if customer_id in d:
                matching_dirs.append((root / d).resolve())
        for f in files:
            if f.lower().endswith(".pdf") and customer_id in f:
                matching_pdfs.append((root / f).resolve())
    return matching_dirs, matching_pdfs


def _topmost_dirs(dirs: List[Path]) -> List[Path]:
    """只保留「最顶层」匹配文件夹：若 A 是 B 的父目录则只保留 A，避免重复复制。"""
    dirs = list(set(dirs))
    result: List[Path] = []
    for d in dirs:
        # d 是否位于其它任意匹配目录之下（即其它是 d 的祖先）
        under_other = any(
            t != d and len(d.parts) > len(t.parts) and d.parts[: len(t.parts)] == t.parts
            for t in dirs
        )
        if not under_other:
            result.append(d)
    return result


def _pdf_under_any(pdf: Path, dirs: List[Path]) -> bool:
    """PDF 是否位于 dirs 中某个目录之下（复制该目录时已包含此 PDF）。"""
    for t in dirs:
        if len(pdf.parts) > len(t.parts) and pdf.parts[: len(t.parts)] == t.parts:
            return True
    return False


def _yg_has_customer_folder_and_pdf(yg_dir: Path, customer_id: str) -> bool:
    """
    检查 YG 目录下是否同时存在「名称包含 customer_id 的文件夹」和「名称包含 customer_id 的 PDF」。
    若都有则返回 True（可跳过复制）。
    """
    yg_dir = Path(yg_dir)
    if not yg_dir.is_dir():
        return False
    has_folder = False
    has_pdf = False
    for item in yg_dir.iterdir():
        if customer_id not in item.name:
            continue
        if item.is_dir():
            has_folder = True
        elif item.is_file() and item.suffix.lower() == ".pdf":
            has_pdf = True
        if has_folder and has_pdf:
            return True
    return has_folder and has_pdf


def copy_customer_files_to_yg(
    source_dir: Path,
    dest_yg_dir: Path,
    customer_id: str,
    log: Callable[[str], None],
) -> None:
    """
    在 source_dir 及其子目录中递归查找名称包含 customer_id 的文件夹和 PDF，
    复制到 dest_yg_dir。若 YG 目录已存在且其中已有该客户的文件夹和 PDF，则跳过并记日志。
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

    # 若 YG 下已有该客户的文件夹和 PDF，则跳过复制
    if _yg_has_customer_folder_and_pdf(dest_yg_dir, customer_id):
        log(f"[跳过] {dest_yg_dir} 下已存在客户文件夹与客户PDF（客户编号 {customer_id}）")
        return

    matching_dirs, matching_pdfs = _find_matching_dirs_and_pdfs(source_dir, customer_id)
    topmost_dirs = _topmost_dirs(matching_dirs)
    # 只复制不在「已复制文件夹」内的 PDF，避免重复
    pdfs_to_copy = [p for p in matching_pdfs if not _pdf_under_any(p, topmost_dirs)]

    copied_any = False
    for item in topmost_dirs:
        if not item.is_dir():
            continue
        dest = dest_yg_dir / item.name
        try:
            if dest.exists():
                shutil.rmtree(dest)
            shutil.copytree(item, dest)
            log(f"[复制] 文件夹 -> {dest_yg_dir.name}: {item.name} （来自 {item.parent}）")
            copied_any = True
        except Exception as e:
            log(f"[错误] 复制失败 {item}: {e}")

    for item in pdfs_to_copy:
        if not item.is_file():
            continue
        dest = dest_yg_dir / item.name
        try:
            shutil.copy2(item, dest)
            log(f"[复制] 文件 -> {dest_yg_dir.name}: {item.name} （来自 {item.parent}）")
            copied_any = True
        except Exception as e:
            log(f"[错误] 复制失败 {item}: {e}")

    if not copied_any:
        log(f"[无匹配] 客户编号 {customer_id} 在源目录及子目录中未找到对应文件夹/PDF")


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
