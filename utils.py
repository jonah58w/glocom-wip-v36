
# -*- coding: utf-8 -*-
import io
import re
import pandas as pd
import streamlit as st

DONE_WIP_SET = {"完成", "DONE", "COMPLETE", "COMPLETED", "FINISHED", "FINISH"}


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


def normalize_match_qty(v):
    s = safe_text(v).replace(",", "")
    if not s:
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def normalize_match_date(v):
    text = safe_text(v)
    if not text:
        return None
    try:
        dt = pd.to_datetime(text, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None


def normalize_wip_value(value: str) -> str:
    text = safe_text(value)
    if not text:
        return ""
    if text.upper() in DONE_WIP_SET or text in DONE_WIP_SET:
        return "完成"
    return text


def clean_part_no(text):
    val = safe_text(text)
    if not val:
        return ""
    return re.sub(r"\s*\(.*?\)\s*$", "", val).strip()


def normalize_due_date_text(value):
    text = safe_text(value)
    if not text:
        return ""
    if "=>" in text:
        text = text.split("=>")[-1].strip()
    try:
        dt = pd.to_datetime(text, errors="coerce")
        if pd.isna(dt):
            return text
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return text


def refresh_after_update():
    st.cache_data.clear()
    st.rerun()


def show_no_data_layout():
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Orders", 0)
    c2.metric("Production", 0)
    c3.metric("Shipping", 0)
    st.divider()
    st.warning("No data from Teable")


def wip_display_html(value: str) -> str:
    text = safe_text(value)
    lower = text.lower()

    if any(k in lower for k in ["完成"]) or text.upper() in DONE_WIP_SET:
        label = text or "完成"
        bg = "#065f46"
        fg = "#d1fae5"
    elif any(k in lower for k in ["ship", "shipping", "出貨"]):
        label = text or "Shipping"
        bg = "#14532d"
        fg = "#dcfce7"
    elif any(k in lower for k in ["pack", "包裝"]):
        label = text or "Packing"
        bg = "#166534"
        fg = "#dcfce7"
    elif any(k in lower for k in ["fqc", "qa", "inspection", "成檢", "測試"]):
        label = text or "Inspection"
        bg = "#854d0e"
        fg = "#fef3c7"
    elif any(k in lower for k in ["aoi", "drill", "route", "routing", "plating", "inner", "production", "防焊", "壓合", "外層", "內層", "成型"]):
        label = text or "Production"
        bg = "#9a3412"
        fg = "#ffedd5"
    elif any(k in lower for k in ["eng", "gerber", "cam", "eq"]):
        label = text or "Engineering"
        bg = "#1d4ed8"
        fg = "#dbeafe"
    elif any(k in lower for k in ["hold", "等待", "暫停"]):
        label = text or "On Hold"
        bg = "#7f1d1d"
        fg = "#fee2e2"
    else:
        label = text or "-"
        bg = "#374151"
        fg = "#f3f4f6"

    return f'<span class="wip-chip" style="background:{bg};color:{fg};">{label}</span>'


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
    bio.seek(0)
    return bio.getvalue()
