# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st


# ==================================
# CONFIG
# ==================================
DEFAULT_TABLE_URL = "https://app.teable.ai/api/table/tbl6c05EPXYtJcZfeir/record"

try:
    TEABLE_TOKEN = st.secrets.get("TEABLE_TOKEN", "")
except Exception:
    TEABLE_TOKEN = ""

try:
    TABLE_URL = st.secrets.get("TEABLE_TABLE_URL", DEFAULT_TABLE_URL)
except Exception:
    TABLE_URL = DEFAULT_TABLE_URL

try:
    TEABLE_VIEW_ID = st.secrets.get("TEABLE_VIEW_ID", "")
except Exception:
    TEABLE_VIEW_ID = ""


# ==================================
# CANDIDATES
# ==================================
PO_CANDIDATES = [
    "PO", "PO#", "P/O", "訂單", "訂單號", "Order No", "Order", "PO No", "PO Number"
]
PART_CANDIDATES = [
    "Part", "Part No", "Part Number", "P/N", "料號", "品名", "Item", "Item No", "PN"
]
WIP_CANDIDATES = [
    "WIP", "進度", "Status", "Current WIP", "目前進度", "生產進度"
]
FACTORY_CANDIDATES = [
    "Factory", "工廠", "Vendor", "供應商"
]
CUSTOMER_CANDIDATES = [
    "Customer", "客戶", "Client"
]
QTY_CANDIDATES = [
    "Qty", "QTY", "Quantity", "數量"
]
SHIP_DATE_CANDIDATES = [
    "Ship Date", "Shipment Date", "交期", "出貨日", "出貨日期", "ETD"
]
FACTORY_DUE_CANDIDATES = [
    "Factory Due", "工廠交期", "Factory Due Date", "工廠預交"
]
REMARK_CANDIDATES = [
    "Remark", "Remarks", "備註", "說明", "Note", "Notes"
]
CUSTOMER_TAG_CANDIDATES = [
    "Customer Tag", "Customer Tags", "客戶標籤", "客戶備註標籤", "Tag", "Tags"
]

PROCESS_CANDIDATE_KEYWORDS = [
    "開料", "壓合", "鑽孔", "電鍍", "線路", "綠漆", "文字", "成型", "測試", "包裝", "出貨",
    "cut", "lam", "drill", "plating", "aoi", "mask", "silk", "routing", "test", "packing", "ship"
]


# ==================================
# BASIC HELPERS
# ==================================
def _safe_text(value: Any) -> str:
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def _normalize_key(text: str) -> str:
    return re.sub(r"[\s_\-#/]+", "", _safe_text(text)).lower()


def _find_first_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None

    normalized = {_normalize_key(c): c for c in df.columns}
    for cand in candidates:
        key = _normalize_key(cand)
        if key in normalized:
            return normalized[key]

    # 模糊比對
    for cand in candidates:
        key = _normalize_key(cand)
        for col in df.columns:
            if key and key in _normalize_key(col):
                return col
    return None


def _detect_process_columns(df: pd.DataFrame) -> List[str]:
    cols = []
    for c in df.columns:
        t = _safe_text(c).lower()
        if any(k in t for k in PROCESS_CANDIDATE_KEYWORDS):
            cols.append(c)
    return cols


def _extract_table_id(table_url: str) -> Optional[str]:
    if not table_url:
        return None

    m = re.search(r"/table/(tbl[a-zA-Z0-9]+)/record", table_url)
    if m:
        return m.group(1)

    m = re.search(r"/table/(tbl[a-zA-Z0-9]+)", table_url)
    if m:
        return m.group(1)

    m = re.search(r"(tbl[a-zA-Z0-9]+)", table_url)
    if m:
        return m.group(1)

    return None


def _headers() -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {TEABLE_TOKEN}",
        "Content-Type": "application/json",
    }


def _build_record_api_url() -> str:
    if TABLE_URL and "/api/table/" in TABLE_URL and TABLE_URL.rstrip("/").endswith("/record"):
        return TABLE_URL.rstrip("/")

    table_id = _extract_table_id(TABLE_URL or DEFAULT_TABLE_URL)
    if not table_id:
        table_id = _extract_table_id(DEFAULT_TABLE_URL)

    return f"https://app.teable.ai/api/table/{table_id}/record"


def _normalize_records_to_df(records: List[Dict[str, Any]]) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    for rec in records:
        row = {}
        fields = rec.get("fields", {}) or {}
        if isinstance(fields, dict):
            row.update(fields)

        row["record_id"] = rec.get("id", "")
        row["_record_id"] = rec.get("id", "")
        row["_createdTime"] = rec.get("createdTime", "")
        row["_lastModifiedTime"] = rec.get("lastModifiedTime", "")
        rows.append(row)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)

    # 去除欄名空白
    df.columns = [str(c).strip() for c in df.columns]

    return df


# ==================================
# LOAD ORDERS
# ==================================
def load_orders() -> Tuple[pd.DataFrame, str, str]:
    """
    回傳: (df, api_status, api_text)
    """
    if not TEABLE_TOKEN:
        return pd.DataFrame(), "error", "Missing TEABLE_TOKEN in Streamlit secrets."

    api_url = _build_record_api_url()
    if not api_url:
        return pd.DataFrame(), "error", f"Invalid TABLE_URL: {TABLE_URL}"

    all_records: List[Dict[str, Any]] = []
    take = 1000
    skip = 0

    debug_msgs = [
        f"api_url={api_url}",
        f"view_id={TEABLE_VIEW_ID or '(none)'}",
    ]

    try:
        while True:
            params = {
                "take": take,
                "skip": skip,
                "fieldKeyType": "name",
                "cellFormat": "text",
            }
            if TEABLE_VIEW_ID:
                params["viewId"] = TEABLE_VIEW_ID

            resp = requests.get(api_url, headers=_headers(), params=params, timeout=30)
            debug_msgs.append(f"GET {resp.url} -> {resp.status_code}")

            if resp.status_code != 200:
                return pd.DataFrame(), "error", "\n".join(debug_msgs + [resp.text[:2000]])

            data = resp.json()

            records = data.get("records", [])
            if not isinstance(records, list):
                return pd.DataFrame(), "error", "\n".join(debug_msgs + [f"Unexpected response JSON: {str(data)[:2000]}"])

            all_records.extend(records)

            if len(records) < take:
                break

            skip += take

        df = _normalize_records_to_df(all_records)

        if df.empty:
            return df, "ok", "\n".join(debug_msgs + ["Teable connected, but 0 records returned."])

        return df, "ok", "\n".join(debug_msgs + [f"Loaded records: {len(df)}"])

    except Exception as e:
        return pd.DataFrame(), "error", "\n".join(debug_msgs + [f"Exception: {e}"])


# ==================================
# MATCH / UPDATE HELPERS
# ==================================
def _normalize_match_value(value: Any) -> str:
    text = _safe_text(value)
    text = text.replace(" ", "").replace("-", "").replace("_", "").upper()
    return text


def _build_lookup_maps(current_df: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    result = {"po": {}, "part": {}}

    if current_df is None or current_df.empty:
        return result

    po_col = _find_first_col(current_df, PO_CANDIDATES)
    part_col = _find_first_col(current_df, PART_CANDIDATES)
    rec_col = "record_id" if "record_id" in current_df.columns else "_record_id"

    if rec_col not in current_df.columns:
        return result

    for _, row in current_df.iterrows():
        rec_id = _safe_text(row.get(rec_col))
        if not rec_id:
            continue

        if po_col and po_col in current_df.columns:
            po_val = _normalize_match_value(row.get(po_col))
            if po_val:
                result["po"][po_val] = rec_id

        if part_col and part_col in current_df.columns:
            part_val = _normalize_match_value(row.get(part_col))
            if part_val:
                result["part"][part_val] = rec_id

    return result


def _infer_wip_from_row(row: pd.Series, uploaded_df: pd.DataFrame, direct_wip_col: Optional[str]) -> str:
    if direct_wip_col and direct_wip_col in uploaded_df.columns:
        direct_val = _safe_text(row.get(direct_wip_col))
        if direct_val:
            return direct_val

    process_cols = _detect_process_columns(uploaded_df)
    last_done = ""

    for c in process_cols:
        v = _safe_text(row.get(c)).lower()
        if v in {"y", "yes", "ok", "done", "完成", "已完成", "v", "✓", "✔"}:
            last_done = c
        elif v and v not in {"", "n", "no", "x", "-", "未"}:
            last_done = c

    return last_done or ""


def _patch_record(record_id: str, fields: Dict[str, Any]) -> Tuple[bool, str]:
    api_url = f"{_build_record_api_url()}/{record_id}"
    body = {
        "record": {
            "fields": fields
        }
    }

    try:
        resp = requests.patch(api_url, headers=_headers(), json=body, timeout=30)
        if resp.status_code == 200:
            return True, "ok"
        return False, f"HTTP {resp.status_code}: {resp.text[:1000]}"
    except Exception as e:
        return False, str(e)


# ==================================
# BATCH UPDATE
# ==================================
def batch_update_wip_from_excel(
    current_df: pd.DataFrame,
    uploaded_df: pd.DataFrame,
    factory_name: str = "",
) -> Dict[str, Any]:
    results = {
        "success_count": 0,
        "failed_count": 0,
        "warnings": [],
        "details": [],
    }

    if uploaded_df is None or uploaded_df.empty:
        results["warnings"].append("uploaded_df is empty.")
        return results

    if current_df is None or current_df.empty:
        results["warnings"].append("current_df is empty, cannot match Teable records.")
        return results

    if not TEABLE_TOKEN:
        results["warnings"].append("Missing TEABLE_TOKEN in Streamlit secrets.")
        return results

    up_po_col = _find_first_col(uploaded_df, PO_CANDIDATES)
    up_part_col = _find_first_col(uploaded_df, PART_CANDIDATES)
    up_wip_col = _find_first_col(uploaded_df, WIP_CANDIDATES)

    cur_wip_col = _find_first_col(current_df, WIP_CANDIDATES)
    cur_factory_col = _find_first_col(current_df, FACTORY_CANDIDATES)

    if not up_po_col and not up_part_col:
        results["warnings"].append("Uploaded file cannot find PO or Part column.")
        return results

    if not cur_wip_col:
        results["warnings"].append("Teable current table cannot find WIP column.")
        return results

    lookup = _build_lookup_maps(current_df)

    for i, row in uploaded_df.iterrows():
        row_no = i + 2

        po_val_raw = row.get(up_po_col) if up_po_col else ""
        part_val_raw = row.get(up_part_col) if up_part_col else ""

        po_key = _normalize_match_value(po_val_raw)
        part_key = _normalize_match_value(part_val_raw)

        record_id = ""
        matched_by = ""

        if po_key and po_key in lookup["po"]:
            record_id = lookup["po"][po_key]
            matched_by = "PO"
        elif part_key and part_key in lookup["part"]:
            record_id = lookup["part"][part_key]
            matched_by = "PART"

        wip_value = _infer_wip_from_row(row, uploaded_df, up_wip_col)

        if not record_id:
            results["failed_count"] += 1
            results["details"].append({
                "row": row_no,
                "po": _safe_text(po_val_raw),
                "part": _safe_text(part_val_raw),
                "wip": wip_value,
                "error": "No matching Teable record.",
            })
            continue

        if not wip_value:
            results["failed_count"] += 1
            results["details"].append({
                "row": row_no,
                "po": _safe_text(po_val_raw),
                "part": _safe_text(part_val_raw),
                "error": "Cannot infer WIP from uploaded row.",
            })
            continue

        fields_to_update = {
            cur_wip_col: wip_value
        }

        if cur_factory_col and factory_name:
            fields_to_update[cur_factory_col] = factory_name

        ok, msg = _patch_record(record_id, fields_to_update)
        if ok:
            results["success_count"] += 1
            results["details"].append({
                "row": row_no,
                "po": _safe_text(po_val_raw),
                "part": _safe_text(part_val_raw),
                "wip": wip_value,
                "matched_by": matched_by,
                "status": "更新成功",
            })
        else:
            results["failed_count"] += 1
            results["details"].append({
                "row": row_no,
                "po": _safe_text(po_val_raw),
                "part": _safe_text(part_val_raw),
                "wip": wip_value,
                "error": msg,
            })

    return results
