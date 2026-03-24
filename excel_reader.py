# -*- coding: utf-8 -*-
"""
excel_reader.py
Excel 讀取工具

支援：
- xlsx / xls
- .xls 自動轉 xlsx
- 自動找最佳 sheet
- 自動找最佳 header row
- 保留舊介面函式，避免現有程式壞掉
"""

from __future__ import annotations

import io
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from typing import Optional, List, Dict, Any, Tuple

import pandas as pd

from utils import normalize_columns, compact_text


# =========================================================
# 欄位關鍵字
# =========================================================
PO_CANDIDATES = [
    "PO#", "PO", "P/O", "訂單編號", "訂單號", "訂單號碼", "工單", "工單號", "單號",
    "ORDER NO", "Order No", "Order Number"
]

PART_CANDIDATES = [
    "Part No", "Part No.", "P/N", "PN", "料號", "品號", "客戶料號",
    "Cust. P / N", "LS P/N", "祥竑料號", "客戶品號", "產品料號", "產品編號", "Model"
]

QTY_CANDIDATES = [
    "Qty", "QTY", "Q'TY", "Order Q'TY (PCS)", "訂購量(PCS)", "訂單量(PCS)",
    "訂單量", "數量", "數量(PCS)", "PCS", "生產數量", "投產數", "訂單數量", "未出貨數量"
]

DUE_CANDIDATES = [
    "Factory Due Date", "工廠交期", "交貨日期", "交期", "出貨日期",
    "Required Ship date", "Required Ship Date", "confirmed DD", "confrimed DD",
    "預交日", "預定交期", "交貨期"
]

SHIP_DATE_CANDIDATES = [
    "Ship Date", "Ship date", "出貨日期", "交貨日期",
    "Required Ship date", "Required Ship Date", "confirmed DD", "confrimed DD"
]

WIP_CANDIDATES = [
    "WIP", "WIP Stage", "進度", "製程", "工序", "目前站別", "生產進度"
]

REMARK_CANDIDATES = [
    "Remark", "備註", "情況", "備註說明", "Note", "說明", "異常備註"
]

HEADER_KEYWORDS = (
    PO_CANDIDATES
    + PART_CANDIDATES
    + QTY_CANDIDATES
    + DUE_CANDIDATES
    + SHIP_DATE_CANDIDATES
    + WIP_CANDIDATES
    + REMARK_CANDIDATES
)


PROCESS_KEYWORDS = [
    "發料", "下料", "排版", "內層", "內乾", "內蝕", "黑化", "壓合", "壓板",
    "鑽孔", "沉銅", "一銅", "電鍍", "乾膜", "外層", "二銅", "二銅蝕刻",
    "AOI", "半測", "防焊", "文字", "噴錫", "化金", "OSP", "化銀",
    "成型", "V-CUT", "測試", "成檢", "包裝", "出貨", "庫存"
]


# =========================================================
# 資料結構
# =========================================================
@dataclass
class ExcelParseResult:
    df: pd.DataFrame
    sheet_name: Optional[str]
    header_row: Optional[int]
    score: int
    all_sheets: Optional[List[str]] = None
    meta: Optional[Dict[str, Any]] = None


# =========================================================
# 基本工具
# =========================================================
def _safe_str(v: Any) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def _clean_col_name(v: Any) -> str:
    s = _safe_str(v)
    s = s.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    s = " ".join(s.split())
    return s


def _normalize_header_list(cols) -> List[str]:
    return [_clean_col_name(c) for c in cols]


def _column_text_set(cols: List[str]) -> List[str]:
    out = []
    for c in cols:
        s = compact_text(c)
        if s:
            out.append(s.lower())
    return out


def _count_nonempty_rows(df: pd.DataFrame) -> int:
    if df is None or df.empty:
        return 0
    return int(df.dropna(how="all").shape[0])


def _count_nonempty_cols(df: pd.DataFrame) -> int:
    if df is None or df.empty:
        return 0
    return int(df.dropna(axis=1, how="all").shape[1])


# =========================================================
# Excel 讀取 / 轉檔
# =========================================================
def try_read_excel_bytes(file_bytes: bytes, header=0, sheet_name=0):
    bio = io.BytesIO(file_bytes)
    return pd.read_excel(bio, header=header, sheet_name=sheet_name)


def convert_xls_bytes_to_xlsx_bytes(file_bytes: bytes, original_name: str) -> bytes:
    suffix = os.path.splitext(original_name)[1].lower()
    if suffix != ".xls":
        return file_bytes

    with tempfile.TemporaryDirectory() as td:
        src_path = os.path.join(td, original_name)
        with open(src_path, "wb") as f:
            f.write(file_bytes)

        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            raise RuntimeError(
                "無法轉換 .xls：系統未安裝 libreoffice / soffice，且 pandas 讀取 .xls 也失敗。"
            )

        cmd = [soffice, "--headless", "--convert-to", "xlsx", "--outdir", td, src_path]
        result = subprocess.run(cmd, capture_output=True, text=True)

        xlsx_path = os.path.join(td, os.path.splitext(original_name)[0] + ".xlsx")
        if result.returncode != 0 or not os.path.exists(xlsx_path):
            raise RuntimeError(
                f".xls 轉檔失敗：{result.stderr or result.stdout or 'unknown error'}"
            )

        with open(xlsx_path, "rb") as f:
            return f.read()


def get_excel_file_obj(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    name = uploaded_file.name

    try:
        return pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception:
        if name.lower().endswith(".xls"):
            xlsx_bytes = convert_xls_bytes_to_xlsx_bytes(file_bytes, name)
            return pd.ExcelFile(io.BytesIO(xlsx_bytes))
        raise


# =========================================================
# 清理 / 評分
# =========================================================
def _clean_df_after_header(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    df = df.dropna(axis=0, how="all")
    df = df.dropna(axis=1, how="all")

    if df.empty:
        return pd.DataFrame()

    cleaned_cols = []
    for i, c in enumerate(df.columns):
        cs = _clean_col_name(c)
        cleaned_cols.append(cs if cs else f"UNNAMED_{i}")
    df.columns = cleaned_cols

    df = df[df.notna().any(axis=1)].copy()
    df = df.reset_index(drop=True)
    return df


def _header_keyword_score(cols: List[str]) -> int:
    if not cols:
        return 0

    score = 0
    cols_norm = _column_text_set(cols)

    for kw in HEADER_KEYWORDS:
        kw_norm = compact_text(kw).lower()

        if any(c == kw_norm for c in cols_norm):
            score += 6
        elif any(kw_norm in c for c in cols_norm):
            score += 3

    # 特別加權
    for bucket, weight in [
        (PO_CANDIDATES, 18),
        (PART_CANDIDATES, 15),
        (WIP_CANDIDATES, 12),
        (QTY_CANDIDATES, 8),
        (DUE_CANDIDATES, 8),
    ]:
        bucket_hit = False
        for kw in bucket:
            kw_norm = compact_text(kw).lower()
            if any(c == kw_norm or kw_norm in c for c in cols_norm):
                bucket_hit = True
                break
        if bucket_hit:
            score += weight

    # 製程欄位加分
    proc_hits = 0
    for p in PROCESS_KEYWORDS:
        pn = compact_text(p).lower()
        if any(c == pn or pn in c for c in cols_norm):
            proc_hits += 1
    if proc_hits >= 3:
        score += 10
    elif proc_hits >= 1:
        score += 4

    nonempty_cols = len([c for c in cols if _clean_col_name(c)])
    if nonempty_cols >= 4:
        score += 4
    if nonempty_cols >= 6:
        score += 6
    if nonempty_cols >= 10:
        score += 4

    unnamed_count = sum(
        1 for c in cols if compact_text(c).lower().startswith("unnamed")
    )
    score -= unnamed_count * 2

    purely_numeric = 0
    for c in cols:
        s = compact_text(c)
        if s.isdigit():
            purely_numeric += 1
    score -= purely_numeric * 2

    return score


def _data_density_score(df: pd.DataFrame) -> int:
    if df is None or df.empty:
        return 0

    score = 0
    rows = _count_nonempty_rows(df)
    cols = _count_nonempty_cols(df)

    if rows >= 3:
        score += 3
    if rows >= 8:
        score += 5
    if rows >= 15:
        score += 4

    if cols >= 4:
        score += 3
    if cols >= 6:
        score += 5
    if cols >= 10:
        score += 4

    return score


def _sheet_name_bonus(sheet_name: str) -> int:
    s = compact_text(sheet_name).lower()
    if not s:
        return 0

    bonus = 0
    keywords = [
        "ls", "wip", "progress", "進度", "流程", "工作", "生產", "report", "報表"
    ]
    for kw in keywords:
        if kw in s:
            bonus += 2
    return bonus


def _evaluate_candidate(df: pd.DataFrame, sheet_name: str) -> int:
    if df is None or df.empty:
        return -999

    cols = _normalize_header_list(df.columns.tolist())
    score = 0
    score += _header_keyword_score(cols)
    score += _data_density_score(df)
    score += _sheet_name_bonus(sheet_name)

    # 若完全沒有 PO / Part / WIP / 製程，降分
    cols_norm = _column_text_set(cols)

    has_po = any(compact_text(k).lower() in c or compact_text(k).lower() == c for k in PO_CANDIDATES for c in cols_norm)
    has_part = any(compact_text(k).lower() in c or compact_text(k).lower() == c for k in PART_CANDIDATES for c in cols_norm)
    has_wip = any(compact_text(k).lower() in c or compact_text(k).lower() == c for k in WIP_CANDIDATES for c in cols_norm)

    process_hits = 0
    for p in PROCESS_KEYWORDS:
        pn = compact_text(p).lower()
        if any(pn in c or pn == c for c in cols_norm):
            process_hits += 1

    if not has_po and not has_part:
        score -= 12
    if not has_wip and process_hits < 3:
        score -= 8

    return score


# =========================================================
# 舊介面：讀第一個非空 raw sheet
# =========================================================
def read_first_nonempty_sheet_raw(uploaded_file) -> Tuple[pd.DataFrame, Optional[str]]:
    xls = get_excel_file_obj(uploaded_file)

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            if df is not None and not df.empty and _count_nonempty_rows(df) > 0:
                return df, sheet
        except Exception:
            continue

    return pd.DataFrame(), None


# =========================================================
# 舊介面：指定 header 讀第一個可用 sheet
# =========================================================
def read_first_nonempty_sheet_with_header(uploaded_file, header=0) -> Tuple[pd.DataFrame, Optional[str]]:
    """
    保留舊介面，避免其他模組壞掉。
    這裡不是單純第一個非空，而是選出指定 header 下最像有效表的 sheet。
    """
    xls = get_excel_file_obj(uploaded_file)

    best_df = pd.DataFrame()
    best_sheet = None
    best_score = -999

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=header)
            df = _clean_df_after_header(df)
            if df.empty:
                continue

            norm_df = normalize_columns(df.copy())
            score = _evaluate_candidate(norm_df, sheet)

            if score > best_score:
                best_score = score
                best_df = norm_df
                best_sheet = sheet

        except Exception:
            continue

    if best_df is None or best_df.empty:
        return pd.DataFrame(), None

    return best_df, best_sheet


# =========================================================
# 新介面：自動找最佳 sheet + header
# =========================================================
def detect_best_sheet_and_header(uploaded_file, max_header_scan_rows: int = 8) -> ExcelParseResult:
    """
    掃描所有 sheet 與多個 header 列，找出最佳組合
    """
    xls = get_excel_file_obj(uploaded_file)

    best_df = pd.DataFrame()
    best_sheet = None
    best_header = None
    best_score = -999
    best_meta: Dict[str, Any] = {}

    for sheet in xls.sheet_names:
        try:
            raw_df = pd.read_excel(xls, sheet_name=sheet, header=None)
        except Exception:
            continue

        if raw_df is None or raw_df.empty or _count_nonempty_rows(raw_df) == 0:
            continue

        raw_df = raw_df.dropna(axis=1, how="all")
        if raw_df.empty:
            continue

        scan_rows = min(max_header_scan_rows, max(len(raw_df), 1))

        for header_row in range(scan_rows):
            try:
                df = pd.read_excel(xls, sheet_name=sheet, header=header_row)
                df = _clean_df_after_header(df)
                if df.empty:
                    continue

                original_cols = _normalize_header_list(df.columns.tolist())
                norm_df = normalize_columns(df.copy())
                score = _evaluate_candidate(norm_df, sheet)

                # 若 header_row 越後面，給少量懲罰，避免亂選太下面
                score -= header_row

                # 完全像數據表再加分
                if len(norm_df) >= 5:
                    score += 2

                if score > best_score:
                    best_score = score
                    best_df = norm_df.copy()
                    best_sheet = sheet
                    best_header = header_row
                    best_meta = {
                        "original_columns": original_cols,
                        "normalized_columns": list(norm_df.columns),
                        "sheet_name": sheet,
                    }

            except Exception:
                continue

    if best_df is None or best_df.empty:
        return ExcelParseResult(
            df=pd.DataFrame(),
            sheet_name=None,
            header_row=None,
            score=0,
            all_sheets=xls.sheet_names,
            meta={"reason": "no_valid_sheet_found"},
        )

    return ExcelParseResult(
        df=best_df,
        sheet_name=best_sheet,
        header_row=best_header,
        score=best_score,
        all_sheets=xls.sheet_names,
        meta=best_meta,
    )


def read_best_sheet_with_header(uploaded_file, max_header_scan_rows: int = 8):
    """
    簡化介面：
    回傳 (df, sheet_name, header_row, score)
    """
    result = detect_best_sheet_and_header(uploaded_file, max_header_scan_rows=max_header_scan_rows)
    return result.df, result.sheet_name, result.header_row, result.score
