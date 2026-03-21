# -*- coding: utf-8 -*-
import re
import pandas as pd


def safe_text(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def parse_tags_from_text(text):
    if not text:
        return []
    return [x.strip() for x in str(text).split(",") if x.strip()]


def split_tags(value):
    if value is None:
        return []
    try:
        if isinstance(value, float) and pd.isna(value):
            return []
    except Exception:
        pass
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    text = str(value).strip()
    if not text:
        return []
    return [x.strip() for x in text.split(",") if x.strip()]


def build_tags_value(tags, multi_select_mode=True):
    tags = [str(x).strip() for x in tags if str(x).strip()]
    if multi_select_mode:
        return tags
    return ", ".join(tags)


def safe_to_datetime(series):
    return pd.to_datetime(series, errors="coerce")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df.columns = [str(c).strip() for c in df.columns]
    return df


def get_first_matching_column(df: pd.DataFrame, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def get_series_by_col(df: pd.DataFrame, col_name: str):
    if not col_name or col_name not in df.columns:
        return None
    obj = df[col_name]
    if isinstance(obj, pd.DataFrame):
        return obj.iloc[:, 0]
    return obj


def normalize_match_text(v):
    s = safe_text(v).upper()
    s = s.replace("－", "-").replace("—", "-")
    s = re.sub(r"\s+", "", s)
    return s
