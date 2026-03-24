# -*- coding: utf-8 -*-
"""
utils.py
共用工具函式

提供：
- 文字清洗
- DataFrame 欄名標準化
- 日期文字正規化
- 欄位匹配工具
- 標籤拆分
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import Any, Iterable, List, Optional

import pandas as pd


# =========================================================
# 基本文字工具
# =========================================================
def safe_text(value: Any) -> str:
    """
    安全轉字串：
    - None / NaN -> ""
    - 其他 -> strip 後字串
    """
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    if value is None:
        return ""

    return str(value).strip()


def compact_text(value: Any) -> str:
    """
    緊縮文字：
    - 去除前後空白
    - 換行 / tab 轉空白
    - 連續空白壓成 1 個
    - 去掉全形空白
    """
    s = safe_text(value)
    if not s:
        return ""

    s = s.replace("\u3000", " ")
    s = s.replace("\n", " ")
    s = s.replace("\r", " ")
    s = s.replace("\t", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def normalize_key(value: Any) -> str:
    """
    用於比對主鍵：
    - 去空白
    - 去 - _
    - 小寫化
    """
    s = compact_text(value)
    if not s:
        return ""
    return s.replace(" ", "").replace("-", "").replace("_", "").lower()


# =========================================================
# 欄名正規化
# =========================================================
def _normalize_column_name(name: Any, idx: int = 0) -> str:
    s = compact_text(name)

    if not s:
        return f"UNNAMED_{idx}"

    # 統一括號
    s = s.replace("（", "(").replace("）", ")")

    # 常見大小寫 / 空白 / 斜線差異修正
    replacements = {
        "cust. p / n": "Cust. P / N",
        "cust.p/n": "Cust. P / N",
        "cust p/n": "Cust. P / N",
        "ls p/n": "LS P/N",
        "l/s p/n": "LS P/N",
        "required ship date": "Required Ship date",
        "required shipdate": "Required Ship date",
        "ship date": "Ship Date",
        "shipdate": "Ship Date",
        "confirmed dd": "confirmed DD",
        "confrimed dd": "confrimed DD",
        "part no": "Part No",
        "part no.": "Part No.",
        "p/n": "P/N",
        "po": "PO",
        "p/o": "P/O",
        "qty": "Qty",
        "q'ty": "Q'TY",
        "remark": "Remark",
        "wip": "WIP",
    }

    low = s.lower()
    if low in replacements:
        return replacements[low]

    return s


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    DataFrame 欄名正規化
    - 去空白 / 換行
    - 空欄名改 UNNAMED_x
    - 重複欄名自動加尾碼
    """
    if df is None:
        return pd.DataFrame()

    if not isinstance(df, pd.DataFrame):
        return df

    df = df.copy()

    new_cols: List[str] = []
    seen = {}

    for i, col in enumerate(df.columns):
        c = _normalize_column_name(col, i)

        if c in seen:
            seen[c] += 1
            c = f"{c}_{seen[c]}"
        else:
            seen[c] = 0

        new_cols.append(c)

    df.columns = new_cols
    return df


# =========================================================
# 日期正規化
# =========================================================
def _excel_serial_to_date_str(value: float) -> Optional[str]:
    """
    Excel serial date 轉 YYYY-MM-DD
    """
    try:
        if value <= 0:
            return None
        base = datetime(1899, 12, 30)
        dt = base + pd.to_timedelta(float(value), unit="D")
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None


def normalize_due_text(value: Any) -> str:
    """
    將常見日期格式統一成 YYYY-MM-DD。
    若無法確定則保留原文字。
    支援：
    - datetime / Timestamp
    - Excel serial number
    - YYYY/MM/DD
    - YYYY-MM-DD
    - MM/DD/YYYY
    - YYYY.M.D
    """
    if value is None:
        return ""

    # pandas / python datetime
    if isinstance(value, (pd.Timestamp, datetime)):
        try:
            return pd.to_datetime(value).strftime("%Y-%m-%d")
        except Exception:
            return safe_text(value)

    s = compact_text(value)
    if not s:
        return ""

    # 數字可能是 Excel serial
    if re.fullmatch(r"\d+(\.\d+)?", s):
        try:
            num = float(s)
            if num > 20000:  # 大概率 Excel serial date
                converted = _excel_serial_to_date_str(num)
                if converted:
                    return converted
        except Exception:
            pass

    # 直接讓 pandas 嘗試解析
    try:
        parsed = pd.to_datetime(s, errors="raise")
        return parsed.strftime("%Y-%m-%d")
    except Exception:
        pass

    # 手動處理一些常見型態
    patterns = [
        r"^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$",
        r"^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$",
    ]

    m1 = re.match(patterns[0], s)
    if m1:
        y, m, d = m1.groups()
        try:
            return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
        except Exception:
            return s

    m2 = re.match(patterns[1], s)
    if m2:
        m, d, y = m2.groups()
        try:
            return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
        except Exception:
            return s

    return s


# =========================================================
# 欄位匹配工具
# =========================================================
def get_first_matching_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    """
    從 DataFrame 欄位中，找第一個匹配候選欄名的欄位
    先完全比對，再做包含比對
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return None

    cols = list(df.columns)
    cols_compact = {c: compact_text(c).lower() for c in cols}

    # 先完整比對
    for cand in candidates:
        cand_norm = compact_text(cand).lower()
        for c in cols:
            if cols_compact[c] == cand_norm:
                return c

    # 再包含比對
    for cand in candidates:
        cand_norm = compact_text(cand).lower()
        for c in cols:
            if cand_norm and cand_norm in cols_compact[c]:
                return c

    return None


def get_series_by_col(df: pd.DataFrame, col_name: Optional[str]) -> Optional[pd.Series]:
    """
    安全取得 Series
    """
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return None
    if not col_name:
        return None
    if col_name not in df.columns:
        return None
    return df[col_name]


# =========================================================
# 標籤工具
# =========================================================
def split_tags(value: Any) -> List[str]:
    """
    將標籤字串拆成 list
    支援：
    - list
    - "a,b,c"
    - "a; b; c"
    - 中文逗號 / 頓號 / 斜線
    """
    if value is None:
        return []

    if isinstance(value, list):
        return [safe_text(x) for x in value if safe_text(x)]

    text = safe_text(value)
    if not text:
        return []

    text = text.replace("；", ";")
    text = text.replace("，", ",")
    text = text.replace("、", ",")
    text = text.replace("/", ",")
    text = text.replace("|", ",")

    parts = [x.strip() for x in re.split(r"[;,]", text) if x.strip()]

    # 去重但保序
    out = []
    seen = set()
    for p in parts:
        if p not in seen:
            seen.add(p)
            out.append(p)

    return out


# =========================================================
# 其他輔助
# =========================================================
def is_meaningful_value(value: Any) -> bool:
    s = safe_text(value).lower()
    return s not in {"", "nan", "none", "null", "-", "--"}


def dataframe_has_content(df: pd.DataFrame) -> bool:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return False
    return df.dropna(how="all").shape[0] > 0
