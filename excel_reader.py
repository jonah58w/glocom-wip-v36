import io
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from typing import Optional, List, Dict, Any

import pandas as pd

from utils import normalize_columns
from config import (
    PO_CANDIDATES,
    PART_CANDIDATES,
    QTY_CANDIDATES,
    WIP_CANDIDATES,
    FACTORY_DUE_CANDIDATES,
    SHIP_DATE_CANDIDATES,
    REMARK_CANDIDATES,
    CUSTOMER_CANDIDATES,
)


HEADER_KEYWORDS = (
    PO_CANDIDATES
    + PART_CANDIDATES
    + QTY_CANDIDATES
    + WIP_CANDIDATES
    + FACTORY_DUE_CANDIDATES
    + SHIP_DATE_CANDIDATES
    + REMARK_CANDIDATES
    + CUSTOMER_CANDIDATES
)


@dataclass
class ExcelParseResult:
    df: pd.DataFrame
    sheet_name: Optional[str]
    header_row: Optional[int]
    score: int
    all_sheets: Optional[List[str]] = None
    meta: Optional[Dict[str, Any]] = None


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


def _safe_str(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def _normalize_header_list(cols) -> List[str]:
    return [str(c).strip().replace("\n", " ").replace("\r", " ") for c in cols]


def _keyword_score(cols: List[str]) -> int:
    score = 0
    cols_lower = [c.lower() for c in cols if c]

    for kw in HEADER_KEYWORDS:
        kw_lower = str(kw).strip().lower()
        if any(kw_lower == c for c in cols_lower):
            score += 5
        elif any(kw_lower in c for c in cols_lower):
            score += 3

    # 特別加權
    if any(any(k.lower() == c for c in cols_lower) for k in PO_CANDIDATES):
        score += 15
    if any(any(k.lower() == c for c in cols_lower) for k in PART_CANDIDATES):
        score += 12
    if any(any(k.lower() == c for c in cols_lower) for k in WIP_CANDIDATES):
        score += 10
    if any(any(k.lower() == c for c in cols_lower) for k in QTY_CANDIDATES):
        score += 8
    if any(any(k.lower() == c for c in cols_lower) for k in SHIP_DATE_CANDIDATES):
        score += 6

    # 欄位數太少通常不是正確表頭
    if len([c for c in cols if c]) >= 4:
        score += 5
    if len([c for c in cols if c]) >= 6:
        score += 5

    return score


def _data_density_score(df: pd.DataFrame) -> int:
    if df is None or df.empty:
        return 0

    non_empty_rows = df.dropna(how="all").shape[0]
    non_empty_cols = df.dropna(axis=1, how="all").shape[1]

    score = 0
    if non_empty_rows >= 3:
        score += 3
    if non_empty_rows >= 8:
        score += 5
    if non_empty_cols >= 4:
        score += 3
    if non_empty_cols >= 6:
        score += 5

    return score


def _clean_df_after_header(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")

    # 去掉欄名全空或 Unnamed
    cleaned_cols = []
    for c in df.columns:
        cs = _safe_str(c)
        if not cs:
            cleaned_cols.append("")
        else:
            cleaned_cols.append(cs)
    df.columns = cleaned_cols

    # 刪除完全空白列
    df = df[df.notna().any(axis=1)].copy()

    return df.reset_index(drop=True)


def read_first_nonempty_sheet_raw(uploaded_file):
    xls = get_excel_file_obj(uploaded_file)
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        if not df.empty and df.dropna(how="all").shape[0] > 0:
            return df, sheet
    return pd.DataFrame(), None


def read_first_nonempty_sheet_with_header(uploaded_file, header=0):
    """
    保留舊介面，避免其他檔案呼叫壞掉。
    但內部仍走較穩的讀法。
    """
    xls = get_excel_file_obj(uploaded_file)
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=header)
            df = _clean_df_after_header(df)
            df = normalize_columns(df)
            if not df.empty and df.dropna(how="all").shape[0] > 0 and df.shape[1] > 0:
                return df, sheet
        except Exception:
            continue
    return pd.DataFrame(), None


def detect_best_sheet_and_header(uploaded_file, max_header_scan_rows: int = 8) -> ExcelParseResult:
    """
    核心函式：
    - 掃全部 sheet
    - 每個 sheet 掃前幾列當 header
    - 自動找最佳表頭列
    """
    xls = get_excel_file_obj(uploaded_file)

    best_df = pd.DataFrame()
    best_sheet = None
    best_header = None
    best_score = -1

    for sheet in xls.sheet_names:
        try:
            raw_df = pd.read_excel(xls, sheet_name=sheet, header=None)
        except Exception:
            continue

        if raw_df is None or raw_df.empty or raw_df.dropna(how="all").empty:
            continue

        raw_df = raw_df.dropna(axis=1, how="all")
        header_candidates = min(max_header_scan_rows, max(len(raw_df), 1))

        for header_row in range(header_candidates):
            try:
                df = pd.read_excel(xls, sheet_name=sheet, header=header_row)
                df = _clean_df_after_header(df)
                if df.empty or df.shape[1] == 0:
                    continue

                original_cols = _normalize_header_list(df.columns.tolist())
                score = _keyword_score(original_cols) + _data_density_score(df)

                # 避免全是 unnamed / 數字欄名
                unnamed_count = sum(
                    1 for c in original_cols if (not c) or c.lower().startswith("unnamed")
                )
                if unnamed_count >= max(1, len(original_cols) // 2):
                    score -= 8

                # 看起來像正常資料表，多給一點分
                sample_rows = df.head(5).dropna(how="all")
                if not sample_rows.empty:
                    score += 2

                if score > best_score:
                    best_score = score
                    best_sheet = sheet
                    best_header = header_row
                    best_df = df.copy()

            except Exception:
                continue

    if best_df.empty:
        return ExcelParseResult(
            df=pd.DataFrame(),
            sheet_name=None,
            header_row=None,
            score=0,
            all_sheets=xls.sheet_names,
            meta={"reason": "no_valid_sheet_found"},
        )

    normalized_df = normalize_columns(best_df.copy())

    return ExcelParseResult(
        df=normalized_df,
        sheet_name=best_sheet,
        header_row=best_header,
        score=best_score,
        all_sheets=xls.sheet_names,
        meta={
            "original_columns": _normalize_header_list(best_df.columns.tolist()),
            "normalized_columns": _normalize_header_list(normalized_df.columns.tolist()),
        },
    )


def read_best_sheet_with_header(uploaded_file, max_header_scan_rows: int = 8):
    """
    給 app.py / parser 呼叫的簡單介面
    """
    result = detect_best_sheet_and_header(uploaded_file, max_header_scan_rows=max_header_scan_rows)
    return result.df, result.sheet_name, result.header_row, result.score
