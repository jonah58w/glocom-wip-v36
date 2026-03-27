# -*- coding: utf-8 -*-
import io
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
import requests
import streamlit as st
from PIL import Image
import pytesseract
try:
    from excel_exporter import generate_quote_excel_v2
    EXCEL_EXPORTER_AVAILABLE = True
except Exception:
    EXCEL_EXPORTER_AVAILABLE = False
def generate_quote_excel_v2(*args, **kwargs):
    raise ModuleNotFoundError("excel_exporter.py not found")
from factory_progress_updater import (
    dedupe_import_df_by_key,
    classify_and_update_factory_row,
    build_manual_review_item,
    update_working_orders_local,
    normalize_wip_value,
    patch_record_by_id,
)
from reports import (
    show_new_orders_wip_report,
    show_sandy_internal_wip_report,
    show_sandy_sales_report,
    render_sales_detail_from_teable,
)
# ================================
# PAGE CONFIG
# ================================
st.set_page_config(
    page_title="GLOCOM Control Tower",
    page_icon="🏭",
    layout="wide"
)
# ================================
# SESSION STATE
# ================================
if "manual_review_queue" not in st.session_state:
    st.session_state.manual_review_queue = []
# ================================
# TEABLE CONFIG
# ================================
DEFAULT_TABLE_URL = "https://app.teable.ai/api/table/tbl6c05EPXYtJcZfeir/record"
TEABLE_WEB_URL = "https://app.teable.ai/base/bsedgLzbHjiK0XoZH01/table/tbl6c05EPXYtJcZfeir"
try:
    TEABLE_TOKEN = st.secrets.get("TEABLE_TOKEN", "")
except Exception:
    TEABLE_TOKEN = ""
try:
    TABLE_URL = st.secrets.get("TEABLE_TABLE_URL", DEFAULT_TABLE_URL)
except Exception:
    TABLE_URL = DEFAULT_TABLE_URL
try:
    TESSERACT_CMD = st.secrets.get("TESSERACT_CMD", "")
except Exception:
    TESSERACT_CMD = ""
if TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
HEADERS = {
    "Authorization": f"Bearer {TEABLE_TOKEN}",
    "Content-Type": "application/json"
}
# ================================
# FIELD CONFIG
# ================================
PO_CANDIDATES = [
    "PO#", "PO", "P/O", "訂單編號", "訂單號", "訂單號碼", "工單", "工單號", "單號"
]
CUSTOMER_CANDIDATES = [
    "Customer", "客戶", "客戶名稱"
]
PART_CANDIDATES = [
    "Part No", "Part No.", "P/N", "客戶料號", "Cust. P / N", "LS P/N",
    "料號", "品號", "成品料號", "產品料號"
]
QTY_CANDIDATES = [
    "Qty", "Order Q'TY (PCS)", "Order Q'TY (PCS)", "訂購量 (PCS)",
    "訂購量", "Q'TY", "數量", "PCS", "訂單量", "生產數量", "投產數"
]
FACTORY_CANDIDATES = [
    "Factory", "工廠", "廠編"
]
WIP_CANDIDATES = [
    "WIP", "WIP Stage", "進度", "製程", "工序", "目前站別", "生產進度"
]
FACTORY_DUE_CANDIDATES = [
    "Factory Due Date", "工廠交期", "交貨日期", "Required Ship date",
    "confrimed DD", "交期", "預交日", "預定交期", "交貨期"
]
SHIP_DATE_CANDIDATES = [
    "Ship Date", "Ship date", "出貨日期", "交貨日期", "Required Ship date", "confrimed DD"
]
REMARK_CANDIDATES = [
    "Remark", "備註", "情況", "備註說明", "Note", "說明", "異常備註"
]
CUSTOMER_TAG_CANDIDATES = [
    "Customer Remark Tags", "Customer Tags", "客戶備註標籤"
]
ORDER_DATE_CANDIDATES = [
    "客戶下單日期", "工廠下單日期", "下單日期", "Order Date", "PO Date", "Date",
    "訂單日期", "接單日期"
]
AMOUNT_ORDER_CANDIDATES = [
    "接單金額", "接單總金額", "Order Amount", "Order amount", "Order Total",
    "客戶金額", "銷售金額", "Sales Amount", "Quote Total", "Total Amount", "Amount",
    "INVOICE", "Invoice", "Invoice Amount", "Invoice Total"
]
AMOUNT_SHIP_CANDIDATES = [
    "出貨金額", "出貨總金額", "Shipment Amount", "Ship Amount", "Shipping Amount",
    "Invoice Amount", "Invoice Total", "出貨發票金額", "Invoice", "INVOICE"
]
MULTI_SELECT_MODE = True
TAG_OPTIONS = [
    "Working Gerber for Approval",
    "Engineering Question",
    "Payment Pending",
    "Remake in Process",
    "On Hold",
    "Partial Shipment",
    "Shipped",
    "Waiting Confirmation",
]
DONE_WIP_VALUES = {"完成", "DONE", "COMPLETE", "COMPLETED", "FINISHED", "FINISH"}
# ================================
# STYLE
# ================================
st.markdown(
    """
<style>
.portal-box {
padding: 18px 20px;
border: 1px solid rgba(120,120,120,.22);
border-radius: 16px;
background: rgba(255,255,255,.03);
margin-bottom: 14px;
}
.portal-title {
font-size: 1.2rem;
font-weight: 700;
margin-bottom: 4px;
}
.tag-chip {
display: inline-block;
padding: 4px 10px;
margin: 2px 6px 2px 0;
border-radius: 999px;
font-size: 0.82rem;
border: 1px solid rgba(120,120,120,.25);
background: rgba(255,255,255,.05);
}
.wip-chip {
display: inline-block;
padding: 4px 10px;
border-radius: 999px;
font-size: 0.82rem;
font-weight: 600;
}
</style>
""",
    unsafe_allow_html=True,
)
# ================================
# HELPERS
# ================================
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
def build_tags_value(tags):
    tags = [str(x).strip() for x in tags if str(x).strip()]
    if MULTI_SELECT_MODE:
        return tags
    return ", ".join(tags)
def parse_tags_from_text(text):
    if not text:
        return []
    return [x.strip() for x in str(text).split(",") if x.strip()]
def safe_text(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()
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
def wip_display_html(value: str) -> str:
    text = safe_text(value)
    lower = text.lower()
    if any(k in lower for k in ["完成"]) or text.upper() in DONE_WIP_VALUES:
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
def show_metrics(df: pd.DataFrame, wip_col: str | None):
    total_orders = len(df)
    shipping = 0
    if wip_col:
        wip_series = get_series_by_col(df, wip_col)
        if wip_series is not None:
            shipping = len(
                df[
                    wip_series.astype(str).str.contains(
                        "ship|shipping|出貨",
                        case=False,
                        na=False
                    )
                ]
            )
    production = total_orders - shipping
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Orders", total_orders)
    c2.metric("Production", production)
    c3.metric("Shipping", shipping)
def show_no_data_layout():
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Orders", 0)
    c2.metric("Production", 0)
    c3.metric("Shipping", 0)
    st.divider()
    st.warning("No data from Teable")
def parse_quick_text_line(line: str):
    parts = [x.strip() for x in line.split("|")]
    if not parts or not parts[0]:
        return None
    po = parts[0]
    wip = parts[1] if len(parts) > 1 else ""
    ship_date = parts[2] if len(parts) > 2 else ""
    tags_text = parts[3] if len(parts) > 3 else ""
    remark = parts[4] if len(parts) > 4 else ""
    return {
        "po": po,
        "wip": wip,
        "ship_date": ship_date,
        "tags": parse_tags_from_text(tags_text),
        "remark": remark,
    }
def refresh_after_update():
    st.cache_data.clear()
    st.rerun()
def customer_portal_columns(df, po_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col):
    return [c for c in [po_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col] if c and c in df.columns]
def parse_factory_text_report(text: str) -> pd.DataFrame:
    lines = [safe_text(x) for x in str(text).splitlines() if safe_text(x)]
    rows = []
    def guess_part_no(line: str, po_value: str, qty_value: str):
        s = line
        if po_value:
            s = re.sub(re.escape(po_value), " ", s, flags=re.IGNORECASE)
        if qty_value:
            s = re.sub(rf"{re.escape(qty_value)}\s*(PCS|PCS\.|PNL|PNLS|SETS)?", " ", s, flags=re.IGNORECASE)
        s = re.sub(r"^\s*\d+[\.)、-]?\s*", "", s)
        s = re.sub(r"(待出貨 | 已出貨 | 出貨 | 生產中 | 製作中 | 待包裝 | 包裝 | 測試 | 完成|HOLD|ON HOLD)", " ", s, flags=re.IGNORECASE)
        s = re.sub(r"\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?", " ", s)
        s = re.sub(r"20\d{2}[/-]\d{1,2}[/-]\d{1,2}", " ", s)
        s = s.replace("--", " ").replace("=>", " ")
        candidates = re.findall(r"[A-Za-z0-9][A-Za-z0-9_./\-]*(?:\s*\([^\)]*\))?", s)
        candidates = [c.strip(" -_:") for c in candidates if c.strip(" -_:")]
        for c in candidates:
            if len(c) >= 3 and re.search(r"\d", c):
                return c
        return candidates[0] if candidates else ""
    for line in lines:
        po = extract_po_from_text(line)
        if not po:
            continue
        qty_match = re.search(r"(\d[\d,]*)\s*(PCS|PCS\.|PNL|PNLS|SETS)?", line, flags=re.IGNORECASE)
        qty = qty_match.group(1) if qty_match else ""
        due = ""
        range_due = parse_mmdd_range_to_date(line)
        if range_due:
            due = range_due
        else:
            date_candidates = re.findall(r"20\d{2}[/-]\d{1,2}[/-]\d{1,2}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2}/\d{1,2}", line)
            if date_candidates:
                due = normalize_match_date(date_candidates[-1]) or date_candidates[-1]
        part_no = guess_part_no(line, po, qty)
        wip = infer_wip_from_text(line)
        if not wip:
            if any(k in line for k in ["待出貨", "已出貨", "出貨"]):
                wip = "Shipping"
            elif any(k in line for k in ["包裝"]):
                wip = "Packing"
            elif any(k in line for k in ["完成"]):
                wip = "完成"
            elif any(k in line for k in ["測試", "成檢"]):
                wip = "Inspection"
            elif any(k in line for k in ["生產", "製作", "鑽孔", "壓合", "防焊", "文字", "成型", "AOI"]):
                wip = "Production"
            elif any(k.lower() in line.lower() for k in ["hold", "on hold"]):
                wip = "On Hold"
        tags = infer_customer_tags_from_text(line)
        if wip == "Shipping" and "Shipped" not in tags:
            tags.append("Shipped")
        if wip == "On Hold" and "On Hold" not in tags:
            tags.append("On Hold")
        rows.append({
            "PO#": po,
            "Part No": part_no,
            "Qty": qty,
            "Factory Due Date": due,
            "Ship Date": due,
            "WIP": wip,
            "Remark": safe_text(line)[:300],
            "Customer Remark Tags": list(dict.fromkeys([t for t in tags if t]))
        })
    return normalize_columns(pd.DataFrame(rows))
# ================================
# MATCH / NORMALIZE HELPERS
# ================================
# moved to factory_progress_updater.py: normalize_match_text
# moved to factory_progress_updater.py: normalize_match_qty
# moved to factory_progress_updater.py: parse_mmdd_range_to_date
# moved to factory_progress_updater.py: normalize_match_date
# moved to factory_progress_updater.py: normalize_part_no
# moved to factory_progress_updater.py: normalize_wip_value
# moved to factory_progress_updater.py: build_record_for_match
# moved to factory_progress_updater.py: match_score
# moved to factory_progress_updater.py: build_teable_match_records
# moved to factory_progress_updater.py: find_best_match_by_4fields
# moved to factory_progress_updater.py: dedupe_import_df_by_key
# ================================
# OCR HELPERS
# ================================
def ocr_image_to_text(image: Image.Image) -> str:
    try:
        return pytesseract.image_to_string(image, lang="eng")
    except Exception as e:
        return f"OCR_ERROR: {e}"
def extract_po_from_text(text: str) -> str:
    patterns = [
        r"\bPO[-\s]?\d+\b",
        r"\bPO\d+\b",
        r"\bEW[-\s]?\d+\b",
        r"\bEW\d+\b",
        r"\b[A-Z]{1,4}\d{5,}(?:-\d+)?\b",
    ]
    for p in patterns:
        m = re.search(p, text, flags=re.IGNORECASE)
        if m:
            return m.group(0).replace(" ", "")
    return ""
def extract_date_from_text(text: str) -> str:
    patterns = [
        r"\b20\d{2}[-/]\d{1,2}[-/]\d{1,2}\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b",
        r"\b\d{1,2}/\d{1,2}\b",
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return m.group(0)
    return ""
def infer_wip_from_text(text: str) -> str:
    t = text.lower()
    if "complete" in t or "completed" in t or "finish" in t or "finished" in t or "完成" in text:
        return "完成"
    if "shipping" in t or "ship out" in t or "shipped" in t:
        return "Shipping"
    if "packing" in t or "packed" in t or "pack" in t or "包裝" in text:
        return "Packing"
    if "hold" in t or "on hold" in t:
        return "On Hold"
    if "remake" in t or "rework" in t:
        return "Remake in Process"
    if "gerber" in t:
        return "Engineering"
    if "eq" in t or "engineering question" in t:
        return "Engineering"
    if "fqc" in t or "qa" in t or "inspection" in t or "測試" in text or "成檢" in text:
        return "Inspection"
    if "aoi" in t:
        return "Production"
    if "drill" in t or "drilling" in t:
        return "Production"
    if "plating" in t:
        return "Production"
    if "routing" in t or "route" in t:
        return "Production"
    if "inner layer" in t or "inner" in t:
        return "Production"
    return ""
def infer_customer_tags_from_text(text: str):
    t = text.lower()
    tags = []
    if "working gerber" in t or "gerber for approval" in t:
        tags.append("Working Gerber for Approval")
    if "eq" in t or "engineering question" in t:
        tags.append("Engineering Question")
    if "payment" in t:
        tags.append("Payment Pending")
    if "remake" in t or "rework" in t:
        tags.append("Remake in Process")
    if "hold" in t:
        tags.append("On Hold")
    if "partial shipment" in t:
        tags.append("Partial Shipment")
    if "shipped" in t or "ship out" in t:
        tags.append("Shipped")
    if "waiting confirmation" in t or "await confirmation" in t:
        tags.append("Waiting Confirmation")
    return list(dict.fromkeys(tags))
def infer_remark_from_text(text: str) -> str:
    cleaned = " ".join(str(text).split())
    return cleaned[:300] if cleaned else ""
# ================================
# EXCEL / XLS HELPERS
# ================================
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
            raise RuntimeError("無法轉換 .xls：系統未安裝 libreoffice / soffice，且 pandas 讀取 .xls 也失敗。")
        cmd = [soffice, "--headless", "--convert-to", "xlsx", "--outdir", td, src_path]
        result = subprocess.run(cmd, capture_output=True, text=True)
        xlsx_path = os.path.join(td, os.path.splitext(original_name)[0] + ".xlsx")
        if result.returncode != 0 or not os.path.exists(xlsx_path):
            raise RuntimeError(f".xls 轉檔失敗：{result.stderr or result.stdout or 'unknown error'}")
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
def read_first_nonempty_sheet_raw(uploaded_file):
    xls = get_excel_file_obj(uploaded_file)
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        if not df.empty and df.dropna(how="all").shape[0] > 0:
            return df, sheet
    return pd.DataFrame(), None
def read_first_nonempty_sheet_with_header(uploaded_file, header=0):
    xls = get_excel_file_obj(uploaded_file)
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=header)
            df = normalize_columns(df)
            if not df.empty and df.dropna(how="all").shape[0] > 0 and df.shape[1] > 0:
                return df, sheet
        except Exception:
            continue
    return pd.DataFrame(), None
# ================================
# 西拓工序表解析
# ================================
def compact_text(x) -> str:
    return re.sub(r"\s+", "", safe_text(x))
def detect_xitop_header_row(raw_df: pd.DataFrame):
    for i in range(min(len(raw_df), 12)):
        row_text = "".join([compact_text(x) for x in raw_df.iloc[i].tolist()])
        if ("P/O" in row_text or "訂單號碼" in row_text) and ("工作流程計劃" in row_text or "交貨日期" in row_text or "客戶料號" in row_text):
            return i
    return None
def combine_header_cells(a, b):
    a1 = compact_text(a)
    b1 = compact_text(b)
    if a1 in {"", "nan"} and b1 in {"", "nan"}:
        return ""
    if b1 in {"", "nan"}:
        return a1
    if a1 in {"", "nan"}:
        return b1
    merged = a1 + b1
    replacements = {
        "訂單號碼 P/O": "P/O",
        "訂購量 (PCS)": "訂購量 (PCS)",
        "客戶料號": "客戶料號",
        "交貨日期": "交貨日期",
    }
    return replacements.get(merged, merged)
def looks_like_xitop_workflow(raw_df: pd.DataFrame) -> bool:
    if raw_df.empty:
        return False
    sample = raw_df.head(8).fillna("").astype(str)
    joined = "".join(sample.apply(lambda col: "".join(col), axis=1).tolist())
    flags = [
        "工作流程計劃" in joined,
        "P/O" in joined or "訂單號碼" in joined,
        "交貨" in joined and "日期" in joined,
        "成型" in joined or "測試" in joined or "防焊" in joined,
    ]
    return sum(flags) >= 2
def parse_xitop_workflow_report(uploaded_file) -> pd.DataFrame:
    raw_df, sheet_name = read_first_nonempty_sheet_raw(uploaded_file)
    if raw_df.empty:
        raise ValueError("西拓報表讀取失敗：工作表為空")
    header_row = detect_xitop_header_row(raw_df)
    if header_row is None:
        raise ValueError("無法辨識西拓報表表頭")
    second_header_row = header_row + 1 if header_row + 1 < len(raw_df) else header_row
    headers = []
    for idx in range(raw_df.shape[1]):
        top = raw_df.iloc[header_row, idx] if idx < raw_df.shape[1] else ""
        bot = raw_df.iloc[second_header_row, idx] if idx < raw_df.shape[1] else ""
        headers.append(combine_header_cells(top, bot) or f"COL_{idx}")
    data_start = second_header_row + 1
    df = raw_df.iloc[data_start:].copy()
    df.columns = headers
    df = df.dropna(how="all").reset_index(drop=True)
    clean_cols = []
    seen = {}
    for c in df.columns:
        c2 = compact_text(c) or "UNNAMED"
        count = seen.get(c2, 0)
        seen[c2] = count + 1
        clean_cols.append(c2 if count == 0 else f"{c2}_{count}")
    df.columns = clean_cols
    po_source = next((c for c in df.columns if "P/O" in c or "訂單號碼" in c), None)
    due_source = next((c for c in df.columns if "交貨日期" in c or c == "交貨"), None)
    part_source = next((c for c in df.columns if "客戶料號" in c or "料號" in c), None)
    qty_source = next((c for c in df.columns if "訂購量" in c), None)
    remark_source = next((c for c in df.columns if "備註" in c), None)
    if not po_source:
        raise ValueError("西拓報表解析失敗：找不到 P/O 欄位")
    process_order = [
        "下料", "內層", "壓合", "鑽孔", "一銅", "外層", "二銅蝕刻",
        "中檢測", "防焊", "文字", "化金", "無鉛", "有鉛", "OSP", "化錫", "化銀",
        "成型", "測試", "成檢", "包裝", "出貨"
    ]
    existing_process_cols = [p for p in process_order if p in df.columns]
    def is_filled(v):
        txt = compact_text(v)
        return txt not in {"", "nan", "None"}
    def infer_xitop_wip(row):
        last_step = ""
        for p in existing_process_cols:
            if is_filled(row.get(p, "")):
                last_step = p
        remark_text = safe_text(row.get(remark_source, "")) if remark_source else ""
        remark_lower = remark_text.lower()
        if "hold" in remark_lower or "暫停" in remark_text or "等待" in remark_text:
            return "On Hold", last_step
        if last_step == "出貨":
            return "Shipping", last_step
        if last_step == "包裝":
            return "Packing", last_step
        if last_step in ["成檢", "測試"]:
            return "Inspection", last_step
        if last_step:
            return "Production", last_step
        return "", last_step
    rows = []
    for _, row in df.iterrows():
        po_val = safe_text(row.get(po_source, ""))
        if not po_val:
            continue
        wip_val, last_step = infer_xitop_wip(row)
        due_val = safe_text(row.get(due_source, "")) if due_source else ""
        part_val = safe_text(row.get(part_source, "")) if part_source else ""
        qty_val = safe_text(row.get(qty_source, "")) if qty_source else ""
        remark_val = safe_text(row.get(remark_source, "")) if remark_source else ""
        tags = []
        if wip_val == "Shipping":
            tags.append("Shipped")
        if wip_val == "On Hold":
            tags.append("On Hold")
        extra_parts = []
        if last_step:
            extra_parts.append(f"Last process: {last_step}")
        if remark_val:
            extra_parts.append(remark_val)
        rows.append({
            "PO#": po_val,
            "Part No": part_val,
            "Qty": qty_val,
            "Factory Due Date": due_val,
            "Ship Date": due_val,
            "WIP": wip_val,
            "Remark": " | ".join([x for x in extra_parts if x])[:300],
            "Customer Remark Tags": tags,
            "_source_sheet": sheet_name or "",
            "_source_type": "xitop_workflow",
        })
    parsed_df = pd.DataFrame(rows)
    return normalize_columns(parsed_df)
# ================================
# LOAD DATA
# ================================
@st.cache_data(ttl=0)
def load_orders():
    if not TEABLE_TOKEN:
        return pd.DataFrame(), "NO_TOKEN", "TEABLE_TOKEN is empty"
    try:
        all_rows = []
        page_token = None
        last_status = 200
        last_text = ""
        while True:
            params = {
                "fieldKeyType": "name",
                "cellFormat": "text",
                "take": 1000,
            }
            if page_token:
                params["pageToken"] = page_token
            response = requests.get(
                TABLE_URL,
                headers=HEADERS,
                params=params,
                timeout=30
            )
            last_status = response.status_code
            last_text = response.text
            if response.status_code != 200:
                return pd.DataFrame(), response.status_code, response.text
            data = response.json()
            records = data.get("records", []) or []
            for rec in records:
                fields = rec.get("fields", {}) or {}
                fields["_record_id"] = rec.get("id", "")
                all_rows.append(fields)
            page_token = (
                data.get("pageToken")
                or data.get("nextPageToken")
                or data.get("next_page_token")
                or None
            )
            if not page_token:
                break
        df = pd.DataFrame(all_rows)
        df = normalize_columns(df)
        return df, last_status, last_text
    except Exception as e:
        return pd.DataFrame(), "EXCEPTION", str(e)
def find_record_id_by_po(df: pd.DataFrame, po_value: str, po_col: str | None):
    if df.empty or not po_col or po_col not in df.columns:
        return None
    po_series = get_series_by_col(df, po_col)
    if po_series is None:
        return None
    matched = df[po_series.astype(str).str.strip().str.lower() == str(po_value).strip().lower()]
    if matched.empty:
        return None
    if "_record_id" in matched.columns:
        return matched.iloc[0]["_record_id"]
    return None
def upsert_to_teable(current_df: pd.DataFrame, po_col_name: str, po_value: str, updates: dict):
    if not po_value:
        return False, "PO is empty"
    record_id = find_record_id_by_po(current_df, po_value=po_value, po_col=po_col_name)
    payload_fields = dict(updates)
    payload_fields[po_col_name] = po_value
    try:
        if record_id:
            r = requests.patch(
                f"{TABLE_URL}/{record_id}",
                headers=HEADERS,
                json={"record": {"fields": payload_fields}},
                timeout=30
            )
        else:
            r = requests.post(
                TABLE_URL,
                headers=HEADERS,
                json={"records": [{"fields": payload_fields}]},
                timeout=30
            )
        if r.status_code in (200, 201):
            return True, r.text
        return False, f"{r.status_code} | {r.text}"
    except Exception as e:
        return False, str(e)
# ================================
# EXCEL EXPORT HELPERS
# ================================
def _g(name_list, default=None):
    for name in name_list:
        if name in st.session_state:
            value = st.session_state.get(name)
            if value not in [None, ""]:
                return value
        if name in globals():
            value = globals().get(name)
            if value not in [None, ""]:
                return value
    return default
def _to_float(value, default=0.0):
    try:
        if value is None or value == "":
            return default
        return float(value)
    except Exception:
        return default
def _to_int(value, default=0):
    try:
        if value is None or value == "":
            return default
        return int(float(value))
    except Exception:
        return default
def _build_default_factory_rows():
    return [
        {"qty": 50, "unit_price": 2.10, "total": 105.0, "factory": "FCF", "remark": "sample"},
        {"qty": 200, "unit_price": 1.75, "total": 350.0, "factory": "FCF", "remark": ""},
        {"qty": 500, "unit_price": 1.52, "total": 760.0, "factory": "FCF", "remark": ""},
        {"qty": "", "unit_price": "", "total": "", "factory": "", "remark": ""},
        {"qty": "", "unit_price": "", "total": "", "factory": "", "remark": ""},
        {"qty": "", "unit_price": "", "total": "", "factory": "", "remark": ""},
        {"qty": "", "unit_price": "", "total": "", "factory": "", "remark": ""},
        {"qty": "", "unit_price": "", "total": "", "factory": "", "remark": ""},
    ]
def _build_default_customer_rows():
    return [
        {"qty": 50, "unit_price": 2.85, "freight": 0.15, "total": 150.0, "remark": "EXW China (CN)"},
        {"qty": 200, "unit_price": 2.28, "freight": 0.08, "total": 472.0, "remark": "EXW China (CN)"},
        {"qty": 500, "unit_price": 2.05, "freight": 0.05, "total": 1050.0, "remark": "EXW China (CN)"},
        {"qty": "", "unit_price": "", "freight": "", "total": "", "remark": ""},
        {"qty": "", "unit_price": "", "freight": "", "total": "", "remark": ""},
        {"qty": "", "unit_price": "", "freight": "", "total": "", "remark": ""},
        {"qty": "", "unit_price": "", "freight": "", "total": "", "remark": ""},
        {"qty": "", "unit_price": "", "freight": "", "total": "", "remark": ""},
    ]
def _normalize_rows(rows, row_type="factory"):
    if not isinstance(rows, list):
        return _build_default_factory_rows() if row_type == "factory" else _build_default_customer_rows()
    normalized = []
    for r in rows[:8]:
        if not isinstance(r, dict):
            continue
        if row_type == "factory":
            normalized.append({
                "qty": r.get("qty", ""),
                "unit_price": r.get("unit_price", ""),
                "total": r.get("total", ""),
                "factory": r.get("factory", r.get("factory_name", "")),
                "remark": r.get("remark", ""),
            })
        else:
            normalized.append({
                "qty": r.get("qty", ""),
                "unit_price": r.get("unit_price", ""),
                "freight": r.get("freight", r.get("ddu", "")),
                "total": r.get("total", ""),
                "remark": r.get("remark", ""),
            })
    while len(normalized) < 8:
        if row_type == "factory":
            normalized.append({"qty": "", "unit_price": "", "total": "", "factory": "", "remark": ""})
        else:
            normalized.append({"qty": "", "unit_price": "", "freight": "", "total": "", "remark": ""})
    return normalized
def build_quote_data_from_existing_app():
    customer_name = _g(["customer_name", "customer", "cust_name", "client_name"], "TEST CUSTOMER")
    part_no = _g(["part_no", "part_number", "pn", "mpn", "model_name"], "PCB-001")
    revision = _g(["revision", "rev", "version"], "A1")
    quote_date = _g(["quote_date"], datetime.now().strftime("%Y-%m-%d"))
    currency = _g(["currency", "quote_currency"], "USD")
    lead_time = _g(["lead_time", "lt", "delivery_time"], "3-4 weeks")
    sales_pic = _g(["sales_pic", "sales", "pic", "owner_name", "user_name"], "Jonah")
    internal_ref_note = _g(["internal_ref_note", "internal_note", "ref_note"], "Generated from app.py")
    layers = _to_int(_g(["layers", "layer_count", "pcb_layers"], 4), 4)
    material = _g(["material", "base_material"], "FR4 Tg170")
    tg = str(_g(["tg", "tg_value"], "170"))
    thickness = _g(["thickness", "board_thickness", "thickness_mm_text"], "1.6mm")
    board_x_mm = _to_float(_g(["board_x_mm", "x_mm", "size_x_mm", "finished_x_mm", "length_mm"], 120.50), 120.50)
    board_y_mm = _to_float(_g(["board_y_mm", "y_mm", "size_y_mm", "finished_y_mm", "width_mm"], 85.00), 85.00)
    outer_copper = _g(["outer_copper", "outer_cu", "cu_outer"], "1 oz")
    inner_copper = _g(["inner_copper", "inner_cu", "cu_inner"], "0.5 oz")
    surface_finish = _g(["surface_finish", "finish", "sf"], 'ENIG 2u"')
    solder_mask = _g(["solder_mask", "sm_color", "mask_color"], "Green")
    silkscreen = _g(["silkscreen", "legend_color", "sl_color"], "White")
    qty = _g(["qty", "quote_qty", "quantity"], "200 pcs")
    via_type = _g(["via_type"], "PTH")
    blind_buried_via = _g(["blind_buried_via", "bb_via"], "No")
    special_process = _g(["special_process", "special_note", "process_note"], "None")
    pcs_per_panel = _to_int(_g(["pcs_per_panel", "panel_qty", "array", "up"], 2), 2)
    panel_note = _g(["panel_note", "array_note"], f"{pcs_per_panel} up")
    panel_x_mm = _to_float(_g(["panel_x_mm", "pnl_x_mm", "working_panel_x_mm"], board_x_mm * 2 + 4), board_x_mm * 2 + 4)
    panel_y_mm = _to_float(_g(["panel_y_mm", "pnl_y_mm", "working_panel_y_mm"], board_y_mm * 2 + 2), board_y_mm * 2 + 2)
    sheet_x_mm = _to_int(_g(["sheet_x_mm", "material_sheet_x", "sheet_width_mm"], 1020), 1020)
    sheet_y_mm = _to_int(_g(["sheet_y_mm", "material_sheet_y", "sheet_length_mm"], 1220), 1220)
    arrangement = _g(["arrangement", "layout", "nesting_layout"], "4 x 3")
    panels_per_sheet = _to_int(_g(["panels_per_sheet", "panel_per_sheet"], 12), 12)
    pcs_per_sheet = _to_int(_g(["pcs_per_sheet"], panels_per_sheet * pcs_per_panel), panels_per_sheet * pcs_per_panel)
    best_direction = _g(["best_direction", "direction", "rotation"], "Rotate 0°")
    utilization_pct = _to_float(_g(["utilization_pct", "util_pct", "utilization"], 78.5), 78.5)
    single_board_size = f"{board_x_mm:.2f} x {board_y_mm:.2f} mm"
    panel_size = f"{panel_x_mm:.2f} x {panel_y_mm:.2f} mm"
    sheet_size = f"{sheet_x_mm} x {sheet_y_mm} mm"
    board_area = board_x_mm * board_y_mm
    sheet_area = sheet_x_mm * sheet_y_mm
    used_area = pcs_per_sheet * board_area
    sheet_area_formula = _g(["sheet_area_formula"], f"{sheet_x_mm} x {sheet_y_mm} = {sheet_area:,.2f} mm²")
    board_area_formula = _g(["board_area_formula"], f"{board_x_mm:.2f} x {board_y_mm:.2f} = {board_area:,.2f} mm²")
    used_area_formula = _g(["used_area_formula"], f"{pcs_per_sheet} x {board_area:,.2f} = {used_area:,.2f} mm²")
    utilization_formula = _g(["utilization_formula"], f"(PCS per sheet x board area) / sheet area = {utilization_pct:.1f}%")
    factory_rows = _normalize_rows(_g(["factory_rows", "factory_quote_rows"], None), "factory")
    customer_rows = _normalize_rows(_g(["customer_rows", "customer_quote_rows"], None), "customer")
    factory_unit_price = _g(["factory_unit_price", "factory_price", "supplier_price"], None)
    factory_total = _g(["factory_total", "factory_total_price"], None)
    factory_name = _g(["factory_name", "supplier_name"], "FCF")
    if factory_unit_price not in [None, ""] and factory_rows[0]["unit_price"] in [None, "", 0]:
        factory_rows[0]["qty"] = qty
        factory_rows[0]["unit_price"] = factory_unit_price
        factory_rows[0]["total"] = factory_total if factory_total not in [None, ""] else ""
        factory_rows[0]["factory"] = factory_name
    customer_unit_price = _g(["customer_unit_price", "unit_price", "quote_unit_price"], None)
    customer_total = _g(["customer_total", "quote_total", "total_amount"], None)
    freight = _g(["freight", "ddu", "shipping_fee"], None)
    incoterm = _g(["incoterm"], "EXW China (CN)")
    if customer_unit_price not in [None, ""] and customer_rows[0]["unit_price"] in [None, "", 0]:
        customer_rows[0]["qty"] = qty
        customer_rows[0]["unit_price"] = customer_unit_price
        customer_rows[0]["freight"] = freight if freight not in [None, ""] else ""
        customer_rows[0]["total"] = customer_total if customer_total not in [None, ""] else ""
        customer_rows[0]["remark"] = incoterm
    notes_text = _g(
        ["notes_text", "notes", "remark_text"],
        "Price based on new utilization calculation.\nFreight excluded.\nFinal spec subject to customer file confirmation."
    )
    quote_data = {
        "header": {
            "title": "PC B QUOTATION WORKSHEET",
            "customer_name": customer_name,
            "part_no": part_no,
            "revision": revision,
            "quote_date": quote_date,
            "currency": currency,
            "lead_time": lead_time,
            "sales_pic": sales_pic,
            "internal_ref_note": internal_ref_note,
        },
        "specification": {
            "section_title": "PCB SPECIFICATION",
            "layers": layers,
            "material": material,
            "tg": tg,
            "thickness": thickness,
            "board_size_x": f"{board_x_mm:.2f} mm",
            "board_size_y": f"{board_y_mm:.2f} mm",
            "outer_copper": outer_copper,
            "inner_copper": inner_copper,
            "surface_finish": surface_finish,
            "solder_mask": solder_mask,
            "silkscreen": silkscreen,
            "qty": qty,
            "via_type": via_type,
            "blind_buried_via": blind_buried_via,
            "special_process": special_process,
            "pcs_per_panel": pcs_per_panel,
            "panel_note": panel_note,
        },
        "utilization": {
            "section_title": "UTILIZATION ANALYSIS",
            "single_board_size": single_board_size,
            "panel_size": panel_size,
            "sheet_size": sheet_size,
            "arrangement": arrangement,
            "pcs_per_panel": pcs_per_panel,
            "panels_per_sheet": panels_per_sheet,
            "pcs_per_sheet": pcs_per_sheet,
            "best_direction": best_direction,
            "utilization_pct": f"{utilization_pct:.1f}%",
            "calculation_summary_title": "Calculation Summary",
            "sheet_area_formula": sheet_area_formula,
            "board_area_formula": board_area_formula,
            "used_area_formula": used_area_formula,
            "utilization_formula": utilization_formula,
        },
        "factory_quote": {
            "section_title": "FACTORY QUOTE",
            "headers": {
                "qty": "Qty",
                "unit_price": "Unit Price",
                "total": "Total",
                "factory": "Factory",
                "remark": "Remark",
            },
            "rows": factory_rows,
        },
        "customer_quote": {
            "section_title": "CUSTOMER QUOTE",
            "headers": {
                "qty": "Qty",
                "unit_price": "Unit Price",
                "freight": "Freight",
                "total": "Total",
                "remark": "Remark",
            },
            "rows": customer_rows,
        },
        "notes": {
            "section_title": "NOTES",
            "content": notes_text,
        },
    }
    return quote_data
def show_excel_quote_export():
    st.subheader("Excel Quote Export")
    template_path = "template.xlsx"
    if not Path(template_path).exists():
        st.error("找不到 template.xlsx，請先放到 app.py 同一層資料夾。")
        return
    base_quote = build_quote_data_from_existing_app()
    col1, col2, col3 = st.columns(3)
    with col1:
        export_customer = st.text_input("Customer", value=base_quote["header"]["customer_name"], key="export_customer")
    with col2:
        export_part_no = st.text_input("Part No.", value=base_quote["header"]["part_no"], key="export_part_no")
    with col3:
        export_revision = st.text_input("Revision", value=base_quote["header"]["revision"], key="export_revision")
    col4, col5, col6 = st.columns(3)
    with col4:
        export_currency = st.text_input("Currency", value=base_quote["header"]["currency"], key="export_currency")
    with col5:
        export_lead_time = st.text_input("Lead Time", value=base_quote["header"]["lead_time"], key="export_lead_time")
    with col6:
        export_sales_pic = st.text_input("Sales / PIC", value=base_quote["header"]["sales_pic"], key="export_sales_pic")
    col7, col8, col9 = st.columns(3)
    with col7:
        export_layers = st.number_input("Layers", min_value=1, value=int(base_quote["specification"]["layers"]), step=1, key="export_layers")
    with col8:
        bx_default = float(str(base_quote["specification"]["board_size_x"]).replace(" mm", ""))
        export_board_x = st.number_input("Board Size X (mm)", min_value=0.01, value=bx_default, step=0.01, key="export_board_x")
    with col9:
        by_default = float(str(base_quote["specification"]["board_size_y"]).replace(" mm", ""))
        export_board_y = st.number_input("Board Size Y (mm)", min_value=0.01, value=by_default, step=0.01, key="export_board_y")
    col10, col11, col12 = st.columns(3)
    with col10:
        panel_x_default = float(base_quote["utilization"]["panel_size"].split(" x ")[0])
        export_panel_x = st.number_input("Panel Size X (mm)", min_value=0.01, value=panel_x_default, step=0.01, key="export_panel_x")
    with col11:
        panel_y_default = float(base_quote["utilization"]["panel_size"].split(" x ")[1].replace(" mm", ""))
        export_panel_y = st.number_input("Panel Size Y (mm)", min_value=0.01, value=panel_y_default, step=0.01, key="export_panel_y")
    with col12:
        export_pcs_per_panel = st.number_input("PCS / Panel", min_value=1, value=int(base_quote["specification"]["pcs_per_panel"]), step=1, key="export_pcs_per_panel")
    col13, col14, col15 = st.columns(3)
    with col13:
        sheet_x_default = int(base_quote["utilization"]["sheet_size"].split(" x ")[0])
        export_sheet_x = st.number_input("Sheet Size X (mm)", min_value=1, value=sheet_x_default, step=1, key="export_sheet_x")
    with col14:
        sheet_y_default = int(base_quote["utilization"]["sheet_size"].split(" x ")[1].replace(" mm", ""))
        export_sheet_y = st.number_input("Sheet Size Y (mm)", min_value=1, value=sheet_y_default, step=1, key="export_sheet_y")
    with col15:
        util_default = float(str(base_quote["utilization"]["utilization_pct"]).replace("%", ""))
        export_util_pct = st.number_input("Utilization %", min_value=0.0, max_value=100.0, value=util_default, step=0.1, key="export_util_pct")
    col16, col17, col18 = st.columns(3)
    with col16:
        export_material = st.text_input("Material", value=base_quote["specification"]["material"], key="export_material")
    with col17:
        export_thickness = st.text_input("Thickness", value=base_quote["specification"]["thickness"], key="export_thickness")
    with col18:
        export_qty = st.text_input("Qty", value=base_quote["specification"]["qty"], key="export_qty")
    col19, col20, col21 = st.columns(3)
    with col19:
        export_outer_cu = st.text_input("Outer Cu", value=base_quote["specification"]["outer_copper"], key="export_outer_cu")
    with col20:
        export_inner_cu = st.text_input("Inner Cu", value=base_quote["specification"]["inner_copper"], key="export_inner_cu")
    with col21:
        export_surface_finish = st.text_input("Surface Finish", value=base_quote["specification"]["surface_finish"], key="export_surface_finish")
    col22, col23, col24 = st.columns(3)
    with col22:
        export_solder_mask = st.text_input("Solder Mask", value=base_quote["specification"]["solder_mask"], key="export_solder_mask")
    with col23:
        export_silkscreen = st.text_input("Silkscreen", value=base_quote["specification"]["silkscreen"], key="export_silkscreen")
    with col24:
        export_arrangement = st.text_input("Arrangement", value=base_quote["utilization"]["arrangement"], key="export_arrangement")
    col25, col26, col27 = st.columns(3)
    with col25:
        export_panels_per_sheet = st.number_input("Panels / Sheet", min_value=1, value=int(base_quote["utilization"]["panels_per_sheet"]), step=1, key="export_panels_per_sheet")
    with col26:
        export_pcs_per_sheet = st.number_input("PCS / Sheet", min_value=1, value=int(base_quote["utilization"]["pcs_per_sheet"]), step=1, key="export_pcs_per_sheet")
    with col27:
        export_best_direction = st.text_input("Best Direction", value=base_quote["utilization"]["best_direction"], key="export_best_direction")
    export_notes = st.text_area(
        "Notes",
        value=base_quote["notes"]["content"],
        height=120,
        key="export_notes",
    )
    with st.expander("Factory Quote Rows", expanded=False):
        st.caption("若你目前 app.py 已經有 factory_rows / customer_rows，就會自動帶入；這裡可再手動調整。")
        factory_rows = []
        for i, row in enumerate(base_quote["factory_quote"]["rows"], start=1):
            cols = st.columns(5)
            qty_i = cols[0].text_input(f"F{i} Qty", value=str(row.get("qty", "")), key=f"fq_qty_{i}")
            up_i = cols[1].text_input(f"F{i} Unit Price", value=str(row.get("unit_price", "")), key=f"fq_up_{i}")
            total_i = cols[2].text_input(f"F{i} Total", value=str(row.get("total", "")), key=f"fq_total_{i}")
            fac_i = cols[3].text_input(f"F{i} Factory", value=str(row.get("factory", "")), key=f"fq_factory_{i}")
            rmk_i = cols[4].text_input(f"F{i} Remark", value=str(row.get("remark", "")), key=f"fq_remark_{i}")
            factory_rows.append({
                "qty": qty_i,
                "unit_price": up_i,
                "total": total_i,
                "factory": fac_i,
                "remark": rmk_i,
            })
    with st.expander("Customer Quote Rows", expanded=False):
        customer_rows = []
        for i, row in enumerate(base_quote["customer_quote"]["rows"], start=1):
            cols = st.columns(5)
            qty_i = cols[0].text_input(f"C{i} Qty", value=str(row.get("qty", "")), key=f"cq_qty_{i}")
            up_i = cols[1].text_input(f"C{i} Unit Price", value=str(row.get("unit_price", "")), key=f"cq_up_{i}")
            freight_i = cols[2].text_input(f"C{i} Freight", value=str(row.get("freight", "")), key=f"cq_freight_{i}")
            total_i = cols[3].text_input(f"C{i} Total", value=str(row.get("total", "")), key=f"cq_total_{i}")
            rmk_i = cols[4].text_input(f"C{i} Remark", value=str(row.get("remark", "")), key=f"cq_remark_{i}")
            customer_rows.append({
                "qty": qty_i,
                "unit_price": up_i,
                "freight": freight_i,
                "total": total_i,
                "remark": rmk_i,
            })
    quote_data = {
        "header": {
            "title": "PC B QUOTATION WORKSHEET",
            "customer_name": export_customer,
            "part_no": export_part_no,
            "revision": export_revision,
            "quote_date": datetime.now().strftime("%Y-%m-%d"),
            "currency": export_currency,
            "lead_time": export_lead_time,
            "sales_pic": export_sales_pic,
            "internal_ref_note": "Generated from app.py",
        },
        "specification": {
            "section_title": "PCB SPECIFICATION",
            "layers": export_layers,
            "material": export_material,
            "tg": base_quote["specification"]["tg"],
            "thickness": export_thickness,
            "board_size_x": f"{export_board_x:.2f} mm",
            "board_size_y": f"{export_board_y:.2f} mm",
            "outer_copper": export_outer_cu,
            "inner_copper": export_inner_cu,
            "surface_finish": export_surface_finish,
            "solder_mask": export_solder_mask,
            "silkscreen": export_silkscreen,
            "qty": export_qty,
            "via_type": base_quote["specification"]["via_type"],
            "blind_buried_via": base_quote["specification"]["blind_buried_via"],
            "special_process": base_quote["specification"]["special_process"],
            "pcs_per_panel": export_pcs_per_panel,
            "panel_note": base_quote["specification"]["panel_note"],
        },
        "utilization": {
            "section_title": "UTILIZATION ANALYSIS",
            "single_board_size": f"{export_board_x:.2f} x {export_board_y:.2f} mm",
            "panel_size": f"{export_panel_x:.2f} x {export_panel_y:.2f} mm",
            "sheet_size": f"{int(export_sheet_x)} x {int(export_sheet_y)} mm",
            "arrangement": export_arrangement,
            "pcs_per_panel": export_pcs_per_panel,
            "panels_per_sheet": export_panels_per_sheet,
            "pcs_per_sheet": export_pcs_per_sheet,
            "best_direction": export_best_direction,
            "utilization_pct": f"{export_util_pct:.1f}%",
            "calculation_summary_title": "Calculation Summary",
            "sheet_area_formula": f"{int(export_sheet_x)} x {int(export_sheet_y)} = {export_sheet_x * export_sheet_y:,.2f} mm²",
            "board_area_formula": f"{export_board_x:.2f} x {export_board_y:.2f} = {export_board_x * export_board_y:,.2f} mm²",
            "used_area_formula": f"{export_pcs_per_sheet} x {(export_board_x * export_board_y):,.2f} = {(export_pcs_per_sheet * export_board_x * export_board_y):,.2f} mm²",
            "utilization_formula": f"(PCS per sheet x board area) / sheet area = {export_util_pct:.1f}%",
        },
        "factory_quote": {
            "section_title": "FACTORY QUOTE",
            "headers": {
                "qty": "Qty",
                "unit_price": "Unit Price",
                "total": "Total",
                "factory": "Factory",
                "remark": "Remark",
            },
            "rows": factory_rows,
        },
        "customer_quote": {
            "section_title": "CUSTOMER QUOTE",
            "headers": {
                "qty": "Qty",
                "unit_price": "Unit Price",
                "freight": "Freight",
                "total": "Total",
                "remark": "Remark",
            },
            "rows": customer_rows,
        },
        "notes": {
            "section_title": "NOTES",
            "content": export_notes,
        },
    }
    with st.expander("匯出前預覽資料", expanded=False):
        st.json(quote_data)
    default_filename = f"Quote_{export_part_no}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    export_filename = st.text_input("輸出檔名", value=default_filename, key="export_filename_v2")
    c1, c2, c3 = st.columns(3)
    clear_factory = c1.checkbox("先清空工廠報價區", value=False, key="clear_factory_v2")
    clear_customer = c2.checkbox("先清空客戶報價區", value=False, key="clear_customer_v2")
    clear_notes = c3.checkbox("先清空 Notes 區", value=False, key="clear_notes_v2")
    if st.button("產生 Excel 報價檔", key="generate_excel_quote_v2"):
        try:
            output_dir = "outputs"
            output_file = generate_quote_excel_v2(
                template_path=template_path,
                output_dir=output_dir,
                quote_data=quote_data,
                filename=export_filename,
                keep_vba=True,
                clear_factory_rows_first=clear_factory,
                clear_customer_rows_first=clear_customer,
                clear_notes_first=clear_notes,
            )
            st.success(f"Excel 已產生：{output_file}")
            with open(output_file, "rb") as f:
                st.download_button(
                    label="下載 Excel 報價檔",
                    data=f,
                    file_name=Path(output_file).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel_quote_v2",
                )
        except Exception as e:
            st.error(f"產生 Excel 失敗：{e}")
# ================================
# LOAD
# ================================
orders, api_status, api_text = load_orders()
if orders.empty:
    st.title("🏭 GLOCOM Control Tower")
    show_no_data_layout()
    st.stop()
# ================================
# DETECT KEY COLUMNS
# ================================
po_col = get_first_matching_column(orders, PO_CANDIDATES)
customer_col = get_first_matching_column(orders, CUSTOMER_CANDIDATES)
part_col = get_first_matching_column(orders, PART_CANDIDATES)
qty_col = get_first_matching_column(orders, QTY_CANDIDATES)
factory_col = get_first_matching_column(orders, FACTORY_CANDIDATES)
wip_col = get_first_matching_column(orders, WIP_CANDIDATES)
factory_due_col = get_first_matching_column(orders, FACTORY_DUE_CANDIDATES)
ship_date_col = get_first_matching_column(orders, SHIP_DATE_CANDIDATES)
remark_col = get_first_matching_column(orders, REMARK_CANDIDATES)
customer_tag_col = get_first_matching_column(orders, CUSTOMER_TAG_CANDIDATES)
# ================================
# CUSTOMER MODE
# ================================
query = st.query_params
customer_param = query.get("customer", None)
if customer_param:
    st.title("GLOCOM Order Status")
    st.caption("Customer WIP Progress")
    if not customer_col:
        st.error("Customer column not found")
        st.stop()
    customer_series = get_series_by_col(orders, customer_col)
    if customer_series is None:
        st.error("Customer data unavailable")
        st.stop()
    cust_orders = orders[
        customer_series.astype(str).str.strip().str.lower() == str(customer_param).strip().lower()
    ].copy()
    if cust_orders.empty:
        st.warning("No orders found")
        st.stop()
    show_metrics(cust_orders, wip_col)
    st.divider()
    for _, row in cust_orders.iterrows():
        po_val = safe_text(row.get(po_col, "")) if po_col else ""
        part_val = safe_text(row.get(part_col, "")) if part_col else ""
        qty_val = safe_text(row.get(qty_col, "")) if qty_col else ""
        wip_val = safe_text(row.get(wip_col, "")) if wip_col else ""
        ship_val = safe_text(row.get(ship_date_col, "")) if ship_date_col else ""
        remark_val = safe_text(row.get(remark_col, "")) if remark_col else ""
        tags_val = split_tags(row.get(customer_tag_col, "")) if customer_tag_col else []
        tag_html = "".join([f'<span class="tag-chip">{t}</span>' for t in tags_val]) if tags_val else '<span class="tag-chip">-</span>'
        st.markdown(
            f"""
<div class="portal-box">
<div class="portal-title">{po_val or '-'}</div>
<div style="display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;">
<div>
<div><strong>P/N</strong> : {part_val or '-'}</div>
<div><strong>Qty</strong> : {qty_val or '-'}</div>
</div>
<div>
<div><strong>WIP</strong> : {wip_display_html(wip_val)}</div>
<div style="margin-top:8px;"><strong>Ship Date</strong> : {ship_val or '-'}</div>
</div>
</div>
<div style="margin-top:12px;"><strong>Customer Remark Tags</strong> : {tag_html}</div>
<div style="margin-top:10px;"><strong>Remark</strong> : {remark_val or '-'}</div>
</div>
""",
            unsafe_allow_html=True,
        )
    portal_cols = customer_portal_columns(cust_orders, po_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col)
    csv_data = cust_orders[portal_cols].to_csv(index=False).encode("utf-8-sig")
    st.download_button("Download WIP CSV", data=csv_data, file_name=f"{customer_param}_wip.csv", mime="text/csv")
    st.stop()
# ================================
# INTERNAL MODE ONLY
# ================================
st.title("🏭 GLOCOM Control Tower")
st.caption("Internal PCB Production Monitoring System")
with st.expander("Debug"):
    st.write("API Status:", api_status)
    st.write("TABLE_URL:", TABLE_URL)
    st.write("Token loaded:", bool(TEABLE_TOKEN))
    st.write("Columns:", list(orders.columns) if not orders.empty else [])
    if isinstance(api_text, str):
        st.text(api_text[:1200])
st.sidebar.title("GLOCOM Internal")
st.sidebar.link_button("Open Teable", TEABLE_WEB_URL, use_container_width=True)
menu = st.sidebar.radio(
    "功能選單",
    [
        "Dashboard",
        "Factory Load",
        "Delayed Orders",
        "Shipment Forecast",
        "Orders",
        "新訂單 WIP",
        "Sandy 內部 WIP",
        "Sandy 銷貨底",
        "業績明細表",
        "Customer Preview",
        "Import / Update",
    ]
)
if st.sidebar.button("Refresh"):
    refresh_after_update()
st.sidebar.markdown("---")
st.sidebar.caption("完成案件請在 Teable 主 View 設定篩選：WIP ≠ 完成")
st.sidebar.caption("另建 Completed View：WIP = 完成")
# ================================
# INTERNAL HELPERS
# ================================
def show_factory_load(df: pd.DataFrame, f_col: str | None):
    st.subheader("🏭 Factory Load")
    if f_col:
        factory_series = get_series_by_col(df, f_col)
        if factory_series is not None:
            factory_summary = (
                factory_series.fillna("(blank)")
                .astype(str)
                .value_counts()
                .reset_index()
            )
            factory_summary.columns = [f_col, "Orders"]
            st.bar_chart(factory_summary.set_index(f_col))
            st.dataframe(factory_summary, use_container_width=True, height=400)
        else:
            st.info("No factory data")
    else:
        st.info("No factory column found")
def show_delayed_orders(df: pd.DataFrame):
    st.subheader("⚠️ Delayed Orders")
    if factory_due_col:
        temp = df.copy()
        due_series = get_series_by_col(temp, factory_due_col)
        if due_series is not None:
            temp["_FactoryDueDateParsed"] = safe_to_datetime(due_series)
            today = pd.Timestamp.today().normalize()
            delayed = temp[
                temp["_FactoryDueDateParsed"].notna() &
                (temp["_FactoryDueDateParsed"] < today)
            ].copy()
            if not delayed.empty:
                delayed["Delay Days"] = (today - delayed["_FactoryDueDateParsed"]).dt.days
                show_cols = [c for c in [po_col, customer_col, part_col, qty_col, factory_col, wip_col, factory_due_col] if c and c in delayed.columns]
                if "Delay Days" not in show_cols:
                    show_cols.append("Delay Days")
                st.dataframe(delayed[show_cols], use_container_width=True, height=520)
            else:
                st.success("No delayed orders")
        else:
            st.info("No factory due date data")
    else:
        st.info("No factory due date column")
def show_shipment_forecast(df: pd.DataFrame):
    st.subheader("📦 Shipment Forecast (Next 7 days)")
    if ship_date_col:
        temp = df.copy()
        ship_series = get_series_by_col(temp, ship_date_col)
        if ship_series is not None:
            temp["_ShipDateParsed"] = safe_to_datetime(ship_series)
            today = pd.Timestamp.today().normalize()
            next_7 = today + timedelta(days=7)
            forecast = temp[
                temp["_ShipDateParsed"].notna() &
                (temp["_ShipDateParsed"] >= today) &
                (temp["_ShipDateParsed"] <= next_7)
            ].copy()
            if not forecast.empty:
                show_cols = [c for c in [po_col, customer_col, part_col, qty_col, factory_col, wip_col, ship_date_col] if c and c in forecast.columns]
                st.dataframe(
                    forecast.sort_values("_ShipDateParsed")[show_cols],
                    use_container_width=True,
                    height=520
                )
            else:
                st.info("No shipment within next 7 days")
        else:
            st.info("No ship date data")
    else:
        st.info("No ship date column")
def make_unique_columns(columns):
    seen = {}
    new_cols = []
    for col in columns:
        col = str(col)
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
    return new_cols
def show_orders_table(df: pd.DataFrame):
    st.subheader("📋 Orders")
    display_df = df.copy()
    display_df.columns = make_unique_columns(display_df.columns)
    st.dataframe(display_df, use_container_width=True, height=520)
    csv_data = display_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Download Orders CSV",
        data=csv_data,
        file_name="glocom_orders.csv",
        mime="text/csv"
    )
# ================================
# INTERNAL VIEWS
# ================================
if menu == "Dashboard":
    show_metrics(orders, wip_col)
    st.divider()
    left, right = st.columns(2)
    with left:
        show_factory_load(orders, factory_col)
    with right:
        show_shipment_forecast(orders)
elif menu == "Factory Load":
    show_factory_load(orders, factory_col)
elif menu == "Delayed Orders":
    show_delayed_orders(orders)
elif menu == "Shipment Forecast":
    show_shipment_forecast(orders)
elif menu == "Orders":
    st.subheader("🔎 Filters")
    filtered = orders.copy()
    col1, col2 = st.columns(2)
    if customer_col:
        customer_series = get_series_by_col(filtered, customer_col)
        if customer_series is not None:
            customer_options = ["All"] + sorted([str(x) for x in customer_series.dropna().unique().tolist()])
            selected_customer = col1.selectbox("Customer", customer_options)
            if selected_customer != "All":
                filtered = filtered[get_series_by_col(filtered, customer_col).astype(str) == selected_customer]
    if wip_col:
        wip_series = get_series_by_col(filtered, wip_col)
        if wip_series is not None:
            wip_options = ["All"] + sorted([str(x) for x in wip_series.dropna().unique().tolist()])
            selected_wip = col2.selectbox("WIP Stage", wip_options)
            if selected_wip != "All":
                filtered = filtered[get_series_by_col(filtered, wip_col).astype(str) == selected_wip]
    show_orders_table(filtered)
elif menu == "新訂單 WIP":
    show_new_orders_wip_report(orders)
elif menu == "Sandy 內部 WIP":
    show_sandy_internal_wip_report(orders)
elif menu == "Sandy 銷貨底":
    show_sandy_sales_report(orders)
elif menu == "業績明細表":
    render_sales_detail_from_teable(orders)
elif menu == "Customer Preview":
    st.subheader("Customer Preview")
    st.caption("僅供內部預覽。客戶請直接使用 Teable View。")
    if not customer_col:
        st.error("Customer column not found in Teable data")
    else:
        customer_series = get_series_by_col(orders, customer_col)
        if customer_series is None:
            st.error("Customer data unavailable")
        else:
            customers = sorted([str(x).strip() for x in customer_series.dropna().unique().tolist() if str(x).strip()])
            if not customers:
                st.warning("No customers found")
            else:
                default_customer = "WESCO"
                default_index = customers.index(default_customer) if default_customer in customers else 0
                selected_customer = st.selectbox("Select customer to preview", customers, index=default_index)
                preview_df = orders[
                    customer_series.astype(str).str.strip().str.lower() == selected_customer.strip().lower()
                ].copy()
                if preview_df.empty:
                    st.warning("No orders found for this customer")
                else:
                    preview_cols = [c for c in [po_col, customer_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col] if c and c in preview_df.columns]
                    st.dataframe(preview_df[preview_cols], use_container_width=True, height=420)
                    st.info("客戶端連結功能已移除，請改用 Teable View 分享給客戶。")
elif menu == "Import / Update":
    st.subheader("Import / Update")
    if not po_col:
        st.error("PO column not found. 請確認 Teable 表有 PO# 欄位。")
        st.stop()
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Excel / CSV / TXT", "Manual Update", "Quick Text", "Email Text", "Image OCR"])
    # TAB 1
    with tab1:
        st.markdown("上傳工廠報表匯入 Teable。")
        st.caption("匯入比對規則：PO#、Part No、Qty、Factory Due Date 四項中，符合 3 項以上且唯一候選者自動覆蓋 WIP；符合 2 項以下或多筆候選者，不自動覆蓋，改列入待人工確認清單。")
        uploaded = st.file_uploader("Upload Excel / CSV / TXT", type=["xlsx", "xls", "csv", "txt"])
        if uploaded is not None:
            try:
                source_type = "standard"
                if uploaded.name.lower().endswith(".csv"):
                    import_df = pd.read_csv(uploaded)
                    import_df = normalize_columns(import_df)
                    source_type = "csv"
                elif uploaded.name.lower().endswith(".txt"):
                    text_content = uploaded.getvalue().decode("utf-8", errors="ignore")
                    import_df = parse_factory_text_report(text_content)
                    source_type = "txt_text_parse"
                else:
                    raw_df, _ = read_first_nonempty_sheet_raw(uploaded)
                    if looks_like_xitop_workflow(raw_df):
                        source_type = "xitop_workflow"
                        import_df = parse_xitop_workflow_report(uploaded)
                    else:
                        import_df, detected_sheet = read_first_nonempty_sheet_with_header(uploaded, header=0)
                        import_df = normalize_columns(import_df)
                        if import_df.empty:
                            raise ValueError("Excel file has no readable non-empty sheet.")
                        source_type = f"standard_excel:{detected_sheet}"
                st.info(f"Detected source type: {source_type}")
                st.dataframe(import_df, use_container_width=True, height=280)
                import_po_col = get_first_matching_column(import_df, PO_CANDIDATES)
                import_customer_col = get_first_matching_column(import_df, CUSTOMER_CANDIDATES)
                import_part_col = get_first_matching_column(import_df, PART_CANDIDATES)
                import_qty_col = get_first_matching_column(import_df, QTY_CANDIDATES)
                import_wip_col = get_first_matching_column(import_df, WIP_CANDIDATES)
                import_factory_due_col = get_first_matching_column(import_df, FACTORY_DUE_CANDIDATES)
                import_ship_col = get_first_matching_column(import_df, SHIP_DATE_CANDIDATES)
                import_remark_col = get_first_matching_column(import_df, REMARK_CANDIDATES)
                import_tag_col = get_first_matching_column(import_df, CUSTOMER_TAG_CANDIDATES)
                st.write("Detected import columns:")
                st.json({
                    "PO": import_po_col,
                    "Customer": import_customer_col,
                    "Part No": import_part_col,
                    "Qty": import_qty_col,
                    "WIP": import_wip_col,
                    "Factory Due Date": import_factory_due_col,
                    "Ship Date": import_ship_col,
                    "Remark": import_remark_col,
                    "Customer Remark Tags": import_tag_col,
                })
                deduped_import_df, duplicate_keys_in_file = dedupe_import_df_by_key(
                    import_df,
                    import_po_col,
                    import_part_col,
                    import_qty_col,
                    import_factory_due_col,
                )
                if duplicate_keys_in_file:
                    st.warning(f"同一批匯入檔案中有重複 key，已自動去重。重複筆數：{len(duplicate_keys_in_file)}")
                if st.button("Batch Update from File"):
                    if not import_wip_col:
                        st.error("匯入檔至少要能辨識出 WIP 欄位。")
                        st.stop()
                    ok_update_count = 0
                    manual_review_count = 0
                    skip_count = 0
                    fail_count = 0
                    logs = []
                    working_orders = orders.copy()
                    manual_review_items = []
                    for _, row in deduped_import_df.iterrows():
                        result = classify_and_update_factory_row(
                            current_df=working_orders,
                            teable_po_col=po_col,
                            teable_part_col=part_col,
                            teable_qty_col=qty_col,
                            teable_wip_col=wip_col,
                            teable_customer_col=customer_col,
                            teable_ship_date_col=ship_date_col,
                            teable_factory_due_col=factory_due_col,
                            teable_remark_col=remark_col,
                            teable_tag_col=customer_tag_col,
                            import_row=row,
                            import_po_col=import_po_col,
                            import_part_col=import_part_col,
                            import_qty_col=import_qty_col,
                            import_wip_col=import_wip_col,
                            import_customer_col=import_customer_col,
                            import_ship_col=import_ship_col,
                            import_factory_due_col=import_factory_due_col,
                            import_remark_col=import_remark_col,
                            import_tag_col=import_tag_col,
                            table_url=TABLE_URL,
                            headers=HEADERS,
                            done_wip_values=DONE_WIP_VALUES,
                            multi_select_mode=MULTI_SELECT_MODE,
                        )
                        po_value = safe_text(row.get(import_po_col, "")) if import_po_col else ""
                        part_value = safe_text(row.get(import_part_col, "")) if import_part_col else ""
                        qty_value = safe_text(row.get(import_qty_col, "")) if import_qty_col else ""
                        due_value = safe_text(row.get(import_factory_due_col, "")) if import_factory_due_col else ""
                        if result["success"] and result["action"] == "UPDATED":
                            ok_update_count += 1
                            logs.append(
                                f"[UPDATED] {po_value} | {part_value} | {qty_value} | {due_value} | matched {','.join(result['match_info'].get('matched_fields', []))}"
                            )
                            if result.get("record_id"):
                                working_orders = update_working_orders_local(
                                    working_orders,
                                    result["record_id"],
                                    result.get("payload_fields", {})
                                )
                        elif result["action"] == "MANUAL_REVIEW":
                            manual_review_count += 1
                            item = build_manual_review_item(
                                import_row=row,
                                import_po_col=import_po_col,
                                import_part_col=import_part_col,
                                import_qty_col=import_qty_col,
                                import_factory_due_col=import_factory_due_col,
                                import_wip_col=import_wip_col,
                                import_remark_col=import_remark_col,
                                match_info=result.get("match_info", {}),
                                reason=result.get("message", "需要人工確認"),
                                teable_po_col=po_col,
                                teable_part_col=part_col,
                                teable_qty_col=qty_col,
                                teable_factory_due_col=factory_due_col,
                                teable_wip_col=wip_col,
                            )
                            manual_review_items.append(item)
                            logs.append(f"[MANUAL REVIEW] {po_value} | {part_value} | {qty_value} | {due_value} -> {result.get('message', '')}")
                        elif result["action"] == "SKIP":
                            skip_count += 1
                            logs.append(f"[SKIP] {po_value} | {part_value} | {qty_value} | {due_value} -> {result.get('message', '')}")
                        else:
                            fail_count += 1
                            logs.append(f"[FAILED] {po_value} | {part_value} | {qty_value} | {due_value} -> {result.get('message', '')}")
                    st.session_state.manual_review_queue = manual_review_items
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Updated WIP", ok_update_count)
                    c2.metric("Manual Review", manual_review_count)
                    c3.metric("Skipped", skip_count)
                    c4.metric("Failed", fail_count)
                    st.success("Batch import finished.")
                    if logs:
                        st.text("\n".join(logs[:200]))
                    if manual_review_items:
                        st.warning("以下資料未自動覆蓋，已列入待人工確認。")
                        review_df = pd.DataFrame([
                            {
                                "PO#": x["PO#"],
                                "Part No": x["Part No"],
                                "Qty": x["Qty"],
                                "Factory Due Date": x["Factory Due Date"],
                                "New WIP": x["New WIP"],
                                "Reason": x["Reason"],
                                "Best Score": x["Best Score"],
                                "Matched Fields": x["Matched Fields"],
                            }
                            for x in manual_review_items
                        ])
                        st.dataframe(review_df, use_container_width=True, height=260)
                        csv_data = review_df.to_csv(index=False).encode("utf-8-sig")
                        st.download_button(
                            "Download Manual Review CSV",
                            data=csv_data,
                            file_name="manual_review_queue.csv",
                            mime="text/csv"
                        )
                    refresh_after_update()
            except Exception as e:
                st.error(f"Import failed: {e}")
    # TAB 2
    with tab2:
        st.markdown("手動更新單筆 WIP。")
        st.caption("若 Excel 匯入時無法確認，請參考下方待人工確認清單，再手動更新。")
        queue = st.session_state.get("manual_review_queue", [])
        if queue:
            st.markdown("### 待人工確認清單")
            review_df = pd.DataFrame([
                {
                    "PO#": x["PO#"],
                    "Part No": x["Part No"],
                    "Qty": x["Qty"],
                    "Factory Due Date": x["Factory Due Date"],
                    "New WIP": x["New WIP"],
                    "Reason": x["Reason"],
                    "Best Score": x["Best Score"],
                    "Matched Fields": x["Matched Fields"],
                }
                for x in queue
            ])
            st.dataframe(review_df, use_container_width=True, height=240)
            selected_idx = st.selectbox(
                "選擇一筆待人工確認資料帶入下方表單",
                options=list(range(len(queue))),
                format_func=lambda i: f"{queue[i]['PO#']} | {queue[i]['Part No']} | {queue[i]['Qty']} | {queue[i]['New WIP']}"
            )
            selected_item = queue[selected_idx]
            with st.expander("查看候選比對資料"):
                candidates = selected_item.get("Candidates", [])
                if candidates:
                    st.dataframe(pd.DataFrame(candidates), use_container_width=True, height=220)
                candidate_options = [c for c in candidates if c.get("record_id")]
                if candidate_options:
                    chosen_candidate_idx = st.selectbox(
                        "選擇候選 record_id 直接套用",
                        options=list(range(len(candidate_options))),
                        format_func=lambda i: f"{candidate_options[i].get('record_id','')} | score={candidate_options[i].get('score',0)} | {candidate_options[i].get('PO#','')} | {candidate_options[i].get('Part No','')}"
                    )
                    if st.button("套用到選定候選", key="apply_candidate_btn"):
                        chosen = candidate_options[chosen_candidate_idx]
                        payload_fields = {}
                        if wip_col and selected_item.get("New WIP"):
                            payload_fields[wip_col] = normalize_wip_value(selected_item.get("New WIP", ""))
                        if remark_col and selected_item.get("Remark"):
                            payload_fields[remark_col] = selected_item.get("Remark", "")
                        if factory_due_col and selected_item.get("Factory Due Date"):
                            payload_fields[factory_due_col] = selected_item.get("Factory Due Date", "")
                        ok, msg = patch_record_by_id(chosen.get("record_id", ""), payload_fields)
                        if ok:
                            st.success("已套用到選定候選")
                            refresh_after_update()
                        else:
                            st.error(msg)
                else:
                    st.info("沒有候選資料，請直接用 PO 手動更新。")
            default_po = selected_item.get("PO#", "")
            default_wip = selected_item.get("New WIP", "")
            default_remark = selected_item.get("Remark", "")
        else:
            default_po = ""
            default_wip = ""
            default_remark = ""
            st.info("目前沒有待人工確認清單。")
        with st.form("manual_update_form"):
            po_input = st.text_input("PO#", value=default_po, placeholder="例如：PO78310")
            wip_input = st.text_input("WIP", value=default_wip, placeholder="例如：Shipping")
            ship_input = st.text_input("Ship Date", placeholder="例如：2026-03-20")
            tags_input = st.multiselect("Customer Remark Tags", TAG_OPTIONS)
            remark_input = st.text_area("Remark", value=default_remark, placeholder="給客戶看的備註")
            submitted = st.form_submit_button("Update This PO")
            if submitted:
                if not po_input.strip():
                    st.error("PO# is required")
                else:
                    updates = {}
                    if wip_col and wip_input.strip():
                        updates[wip_col] = normalize_wip_value(wip_input.strip())
                    if ship_date_col and ship_input.strip():
                        updates[ship_date_col] = ship_input.strip()
                    if customer_tag_col:
                        updates[customer_tag_col] = build_tags_value(tags_input)
                    if remark_col:
                        updates[remark_col] = remark_input.strip()
                    success, msg = upsert_to_teable(
                        current_df=orders,
                        po_col_name=po_col,
                        po_value=po_input.strip(),
                        updates=updates
                    )
                    if success:
                        st.success(f"{po_input.strip()} updated successfully")
                        refresh_after_update()
                    else:
                        st.error(msg)
    # TAB 3
    with tab3:
        st.code(
            "PO78310 | Shipping | 2026-03-20 | Partial Shipment, Shipped | ready to ship\n"
            "PO78311 | On Hold |  | On Hold | waiting customer reply"
        )
        quick_text = st.text_area("Paste Quick Text", height=220)
        if st.button("Batch Update from Quick Text"):
            lines = [x.strip() for x in quick_text.splitlines() if x.strip()]
            ok_count = 0
            fail_count = 0
            logs = []
            for line in lines:
                parsed = parse_quick_text_line(line)
                if not parsed:
                    continue
                updates = {}
                if wip_col and parsed["wip"]:
                    updates[wip_col] = normalize_wip_value(parsed["wip"])
                if ship_date_col and parsed["ship_date"]:
                    updates[ship_date_col] = parsed["ship_date"]
                if customer_tag_col:
                    updates[customer_tag_col] = build_tags_value(parsed["tags"])
                if remark_col and parsed["remark"]:
                    updates[remark_col] = parsed["remark"]
                success, msg = upsert_to_teable(
                    current_df=orders,
                    po_col_name=po_col,
                    po_value=parsed["po"],
                    updates=updates
                )
                if success:
                    ok_count += 1
                else:
                    fail_count += 1
                    logs.append(f"{parsed['po']} -> {msg}")
            st.success(f"Quick text update finished. Success: {ok_count}, Failed: {fail_count}")
            if logs:
                st.text("\n".join(logs[:50]))
            refresh_after_update()
    # TAB 4
    with tab4:
        st.markdown("貼上工廠 Email / 純文字進度，自動解析後再批次更新。")
        st.caption("適合宏棋、優技等以 Email 文字提供 WIP 的工廠。")
        email_text = st.text_area("Paste Email / Text Report", height=240, key="email_text_report")
        if st.button("Parse Email Text"):
            parsed_email_df = parse_factory_text_report(email_text)
            if parsed_email_df.empty:
                st.warning("沒有辨識出有效資料。")
            else:
                st.session_state["parsed_email_df"] = parsed_email_df
                st.success(f"已辨識 {len(parsed_email_df)} 筆資料")
        parsed_email_df = st.session_state.get("parsed_email_df")
        if isinstance(parsed_email_df, pd.DataFrame) and not parsed_email_df.empty:
            st.dataframe(parsed_email_df, use_container_width=True, height=260)
            if st.button("Batch Update from Email Text"):
                temp_upload_df = parsed_email_df.copy()
                import_po_col = get_first_matching_column(temp_upload_df, PO_CANDIDATES)
                import_customer_col = get_first_matching_column(temp_upload_df, CUSTOMER_CANDIDATES)
                import_part_col = get_first_matching_column(temp_upload_df, PART_CANDIDATES)
                import_qty_col = get_first_matching_column(temp_upload_df, QTY_CANDIDATES)
                import_wip_col = get_first_matching_column(temp_upload_df, WIP_CANDIDATES)
                import_factory_due_col = get_first_matching_column(temp_upload_df, FACTORY_DUE_CANDIDATES)
                import_ship_col = get_first_matching_column(temp_upload_df, SHIP_DATE_CANDIDATES)
                import_remark_col = get_first_matching_column(temp_upload_df, REMARK_CANDIDATES)
                import_tag_col = get_first_matching_column(temp_upload_df, CUSTOMER_TAG_CANDIDATES)
                deduped_import_df, _ = dedupe_import_df_by_key(temp_upload_df, import_po_col, import_part_col, import_qty_col, import_factory_due_col)
                ok_update_count = 0
                manual_review_count = 0
                fail_count = 0
                logs = []
                working_orders = orders.copy()
                manual_review_items = list(st.session_state.get("manual_review_queue", []))
                for _, row in deduped_import_df.iterrows():
                    result = classify_and_update_factory_row(
                        current_df=working_orders,
                        teable_po_col=po_col,
                        teable_part_col=part_col,
                        teable_qty_col=qty_col,
                        teable_wip_col=wip_col,
                        teable_customer_col=customer_col,
                        teable_ship_date_col=ship_date_col,
                        teable_factory_due_col=factory_due_col,
                        teable_remark_col=remark_col,
                        teable_tag_col=customer_tag_col,
                        import_row=row,
                        import_po_col=import_po_col,
                        import_part_col=import_part_col,
                        import_qty_col=import_qty_col,
                        import_wip_col=import_wip_col,
                        import_customer_col=import_customer_col,
                        import_ship_col=import_ship_col,
                        import_factory_due_col=import_factory_due_col,
                        import_remark_col=import_remark_col,
                        import_tag_col=import_tag_col,
                        table_url=TABLE_URL,
                        headers=HEADERS,
                        done_wip_values=DONE_WIP_VALUES,
                        multi_select_mode=MULTI_SELECT_MODE,
                    )
                    po_value = safe_text(row.get(import_po_col, "")) if import_po_col else ""
                    part_value = safe_text(row.get(import_part_col, "")) if import_part_col else ""
                    qty_value = safe_text(row.get(import_qty_col, "")) if import_qty_col else ""
                    due_value = safe_text(row.get(import_factory_due_col, "")) if import_factory_due_col else ""
                    if result["success"] and result["action"] == "UPDATED":
                        ok_update_count += 1
                        logs.append(f"[UPDATED] {po_value} | {part_value} | {qty_value} | {due_value}")
                        if result.get("record_id"):
                            working_orders = update_working_orders_local(working_orders, result["record_id"], result.get("payload_fields", {}))
                    elif result["action"] == "MANUAL_REVIEW":
                        manual_review_count += 1
                        manual_review_items.append(
                            build_manual_review_item(
                                import_row=row,
                                import_po_col=import_po_col,
                                import_part_col=import_part_col,
                                import_qty_col=import_qty_col,
                                import_factory_due_col=import_factory_due_col,
                                import_wip_col=import_wip_col,
                                import_remark_col=import_remark_col,
                                match_info=result.get("match_info", {}),
                                reason=result.get("message", "需要人工確認"),
                                teable_po_col=po_col,
                                teable_part_col=part_col,
                                teable_qty_col=qty_col,
                                teable_factory_due_col=factory_due_col,
                                teable_wip_col=wip_col,
                            )
                        )
                    else:
                        fail_count += 1
                        logs.append(f"[FAILED] {po_value} | {part_value} | {qty_value} | {due_value} -> {result.get('message','')}")
                st.session_state.manual_review_queue = manual_review_items
                c1, c2, c3 = st.columns(3)
                c1.metric("Updated WIP", ok_update_count)
                c2.metric("Manual Review", manual_review_count)
                c3.metric("Failed", fail_count)
                if logs:
                    st.text("\n".join(logs[:120]))
                refresh_after_update()
    # TAB 5
    with tab5:
        st.markdown("上傳進度截圖，OCR 辨識後確認再更新 Teable。")
        uploaded_img = st.file_uploader("Upload PNG / JPG / JPEG", type=["png", "jpg", "jpeg"], key="ocr_uploader")
        if uploaded_img is not None:
            try:
                image = Image.open(uploaded_img)
                st.image(image, caption="Uploaded Image", use_container_width=True)
                ocr_text = ocr_image_to_text(image)
                st.text_area("OCR Raw Text", value=ocr_text, height=220)
                if str(ocr_text).startswith("OCR_ERROR:"):
                    st.error(ocr_text)
                else:
                    guessed_po = extract_po_from_text(ocr_text)
                    guessed_wip = infer_wip_from_text(ocr_text)
                    guessed_date = extract_date_from_text(ocr_text)
                    guessed_tags = infer_customer_tags_from_text(ocr_text)
                    guessed_remark = infer_remark_from_text(ocr_text)
                    st.markdown("### Parsed Result")
                    with st.form("ocr_update_form"):
                        po_input = st.text_input("PO#", value=guessed_po)
                        wip_input = st.text_input("WIP", value=guessed_wip)
                        ship_input = st.text_input("Ship Date", value=guessed_date)
                        tags_input = st.multiselect(
                            "Customer Remark Tags",
                            TAG_OPTIONS,
                            default=[t for t in guessed_tags if t in TAG_OPTIONS]
                        )
                        remark_input = st.text_area("Remark", value=guessed_remark, height=120)
                        submitted_ocr = st.form_submit_button("Update to Teable")
                        if submitted_ocr:
                            if not po_input.strip():
                                st.error("PO# is required")
                            else:
                                updates = {}
                                if wip_col and wip_input.strip():
                                    updates[wip_col] = normalize_wip_value(wip_input.strip())
                                if ship_date_col and ship_input.strip():
                                    updates[ship_date_col] = ship_input.strip()
                                if customer_tag_col:
                                    updates[customer_tag_col] = build_tags_value(tags_input)
                                if remark_col and remark_input.strip():
                                    updates[remark_col] = remark_input.strip()
                                success, msg = upsert_to_teable(
                                    current_df=orders,
                                    po_col_name=po_col,
                                    po_value=po_input.strip(),
                                    updates=updates
                                )
                                if success:
                                    st.success(f"{po_input.strip()} updated successfully from OCR")
                                    refresh_after_update()
                                else:
                                    st.error(msg)
            except Exception as e:
                st.error(f"Image OCR failed: {e}")
# Excel Quote Export removed from menu.
