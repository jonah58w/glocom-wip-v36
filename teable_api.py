# -*- coding: utf-8 -*-
"""
Teable API helpers for GLOCOM Control Tower
- 載入 Orders 主資料
- 從 Excel/CSV 更新 WIP 回 Teable
"""

from __future__ import annotations

import math
from typing import Any, Dict, List, Tuple, Optional

import pandas as pd
import requests
import streamlit as st

import config as cfg


# ==============================
# 基本設定
# ==============================
def _get_teable_token() -> str:
    token = ""
    if "TEABLE_TOKEN" in st.secrets:
        token = st.secrets["TEABLE_TOKEN"]
    elif hasattr(cfg, "TEABLE_TOKEN"):
        token = cfg.TEABLE_TOKEN
    return str(token).strip()


def _parse_table_view_from_url(url: str) -> Tuple[str, Optional[str]]:
    """
    解析 /table/{tableId}/view/{viewId} 或 /table/{tableId} 格式
    回傳 (table_id, view_id or None)
    """
    if not url:
        return "", None

    # 去掉 query string
    base = url.split("?", 1)[0].strip("/")

    parts = base.split("/")
    # 找到 'table' 的 index
    if "table" in parts:
        idx = parts.index("table")
        if idx + 1 < len(parts):
            table_id = parts[idx + 1]
        else:
            table_id = ""
    else:
        table_id = ""

    view_id = None
    if "view" in parts:
        vidx = parts.index("view")
        if vidx + 1 < len(parts):
            view_id = parts[vidx + 1]

    return table_id, view_id


def _build_record_api_url(table_id: str) -> str:
    """
    組 Teable record API URL
    官方格式: https://app.teable.io/api/table/{tableId}/record
    """
    base = getattr(cfg, "TEABLE_API_BASE", "https://app.teable.io/api")
    base = base.rstrip("/")
    return f"{base}/table/{table_id}/record"


def _teable_headers() -> Dict[str, str]:
    token = _get_teable_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


# ==============================
# 載入 Orders 主資料
# ==============================
def load_orders() -> Tuple[pd.DataFrame, str, str]:
    """
    從 Teable via /table/{tableId}/record 分頁載入全部訂單
    回傳 (df, status, text)
    status: "ok" / "error"
    text: debug / error 訊息
    """
    token = _get_teable_token()
    if not token:
        return pd.DataFrame(), "error", "TEABLE_TOKEN 未設定（st.secrets 或 config.py）"

    table_url = ""
    if "TEABLE_TABLE_URL" in st.secrets:
        table_url = st.secrets["TEABLE_TABLE_URL"]
    elif hasattr(cfg, "TEABLE_TABLE_URL"):
        table_url = cfg.TEABLE_TABLE_URL

    if not table_url:
        return pd.DataFrame(), "error", "TEABLE_TABLE_URL 未設定"

    table_id, view_id_from_url = _parse_table_view_from_url(table_url)
    if not table_id:
        return pd.DataFrame(), "error", f"無法從 TEABLE_TABLE_URL 解析 tableId: {table_url}"

    record_url = _build_record_api_url(table_id)

    # 分頁抓資料
    all_records: List[Dict[str, Any]] = []
    take = 200
    skip = 0
    page = 1
    debug_msgs: List[str] = []

    while True:
        params: Dict[str, Any] = {
            "take": take,
            "skip": skip,
        }
        # 優先使用 secrets/cfg 中的 View Id，如果沒有則用 URL 裡解析到的 viewId
        view_id = ""
        if "TEABLE_VIEW_ID" in st.secrets:
            view_id = st.secrets["TEABLE_VIEW_ID"]
        elif hasattr(cfg, "TEABLE_VIEW_ID"):
            view_id = cfg.TEABLE_VIEW_ID
        if view_id:
            params["viewId"] = view_id
        elif view_id_from_url:
            params["viewId"] = view_id_from_url

        try:
            resp = requests.get(record_url, headers=_teable_headers(), params=params, timeout=20)
            if resp.status_code != 200:
                return (
                    pd.DataFrame(),
                    "error",
                    f"Teable API 回應錯誤 ({resp.status_code}): {resp.text[:500]}",
                )

            data = resp.json()
            items = data.get("records", [])
            if not items:
                break

            all_records.extend(items)
            debug_msgs.append(f"page {page}: {len(items)} records")

            # 判斷是否還有下一頁
            total = data.get("total", None)
            if total is not None:
                if skip + take >= total:
                    break
            else:
                # 如果沒給 total，就用「這頁小於 take」當終止條件
                if len(items) < take:
                    break

            # 下一頁
            page += 1
            skip += take

            # 安全上限，避免無限 loop
            if page > 50:
                debug_msgs.append("達到 page>50，強制停止")
                break

        except Exception as e:
            return pd.DataFrame(), "error", f"Teable API 呼叫例外: {e}"

    if not all_records:
        return pd.DataFrame(), "ok", "Teable API 沒有回傳任何 record"

    # 轉 DataFrame：取每筆 record 的 fields
    rows: List[Dict[str, Any]] = []
    for rec in all_records:
        fields = rec.get("fields", {}) or {}
        # 保留 record_id 方便之後 PATCH 更新
        fields["_record_id"] = rec.get("id")
        rows.append(fields)

    df = pd.DataFrame(rows)
    debug_text = " | ".join(debug_msgs) or f"Loaded {len(df)} records from Teable"

    return df, "ok", debug_text


# ==============================
# 工廠進度 / WIP 更新
# ==============================
def _build_lookup_maps(current_df: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    """
    建立 PO / PartNo → record_id 的查詢 map
    """
    if current_df is None or current_df.empty:
        return {"by_po": {}, "by_part": {}}

    by_po: Dict[str, str] = {}
    by_part: Dict[str, str] = {}

    po_candidates = getattr(cfg, "PO_CANDIDATES", ["PO", "PO No", "PO#", "訂單編號"])
    part_candidates = getattr(cfg, "PART_CANDIDATES", ["Part", "Part No", "Part#", "料號"])
    record_id_cols = ["_record_id", "record_id", "Record Id"]

    def _first_col(cols: List[str]) -> Optional[str]:
        for c in cols:
            if c in current_df.columns:
                return c
        return None

    po_col = _first_col(po_candidates)
    part_col = _first_col(part_candidates)
    rec_col = _first_col(record_id_cols)

    if not rec_col:
        # 若沒有 record id，後面也沒辦法 PATCH
        return {"by_po": {}, "by_part": {}}

    if po_col:
        for _, row in current_df[[po_col, rec_col]].dropna().iterrows():
            key = str(row[po_col]).strip()
            if key:
                by_po[key] = str(row[rec_col])

    if part_col:
        for _, row in current_df[[part_col, rec_col]].dropna().iterrows():
            key = str(row[part_col]).strip()
            if key:
                by_part[key] = str(row[rec_col])

    return {"by_po": by_po, "by_part": by_part}


def _detect_process_columns(df_up: pd.DataFrame) -> List[str]:
    """
    偵測可能為製程 / WIP 相關欄位（給 fallback_import_update 預覽用）
    """
    if df_up is None or df_up.empty:
        return []
    cols = []
    for c in df_up.columns:
        cl = str(c).lower()
        if any(k in cl for k in ["wip", "process", "製程", "stage", "status", "狀態"]):
            cols.append(c)
    return cols


def _patch_record(table_id: str, record_id: str, fields: Dict[str, Any]) -> Tuple[bool, str]:
    """
    呼叫 Teable PATCH /table/{tableId}/record/{recordId}
    回傳 (success, error_message)
    """
    base = getattr(cfg, "TEABLE_API_BASE", "https://app.teable.io/api").rstrip("/")
    url = f"{base}/table/{table_id}/record/{record_id}"

    payload = {"fields": fields}

    try:
        resp = requests.patch(url, headers=_teable_headers(), json=payload, timeout=20)
        if resp.status_code not in (200, 201):
            return False, f"PATCH {resp.status_code}: {resp.text[:300]}"
        return True, ""
    except Exception as e:
        return False, str(e)


def batch_update_wip_from_excel(
    current_df: pd.DataFrame,
    uploaded_df: pd.DataFrame,
    factory_name: str = "",
) -> Dict[str, Any]:
    """
    從上傳的 Excel/CSV DataFrame 解析 WIP，對照 current_df 的 PO / Part，把對應 record PATCH 回 Teable
    回傳結果 dict，給前端顯示
    """
    token = _get_teable_token()
    if not token:
        return {"success_count": 0, "failed_count": 0, "warnings": ["TEABLE_TOKEN 未設定"], "details": []}

    table_url = ""
    if "TEABLE_TABLE_URL" in st.secrets:
        table_url = st.secrets["TEABLE_TABLE_URL"]
    elif hasattr(cfg, "TEABLE_TABLE_URL"):
        table_url = cfg.TEABLE_TABLE_URL

    table_id, _ = _parse_table_view_from_url(table_url)
    if not table_id:
        return {
            "success_count": 0,
            "failed_count": 0,
            "warnings": ["無法從 TEABLE_TABLE_URL 解析 tableId"],
            "details": [],
        }

    if current_df is None or current_df.empty:
        return {
            "success_count": 0,
            "failed_count": 0,
            "warnings": ["目前主資料 orders 為空，無法比對"],
            "details": [],
        }

    if uploaded_df is None or uploaded_df.empty:
        return {
            "success_count": 0,
            "failed_count": 0,
            "warnings": ["上傳檔案解析後沒有有效資料"],
            "details": [],
        }

    lookup = _build_lookup_maps(current_df)
    by_po = lookup["by_po"]
    by_part = lookup["by_part"]

    po_col = _first_match(uploaded_df, getattr(cfg, "PO_CANDIDATES", ["PO", "PO No", "PO#", "訂單編號"]))
    part_col = _first_match(
        uploaded_df, getattr(cfg, "PART_CANDIDATES", ["Part", "Part No", "Part#", "料號"])
    )
    wip_col = _first_match(uploaded_df, getattr(cfg, "WIP_CANDIDATES", ["WIP", "Status", "製程", "狀態"]))

    details: List[Dict[str, Any]] = []
    success_count = 0
    failed_count = 0
    warnings: List[str] = []

    if not po_col and not part_col:
        warnings.append("上傳檔案中未偵測到 PO 或 Part 欄，無法進行匹配。")
        return {
            "success_count": 0,
            "failed_count": 0,
            "warnings": warnings,
            "details": [],
        }

    if not wip_col:
        warnings.append("上傳檔案中未偵測到 WIP / Status 欄，暫不進行更新。")
        return {
            "success_count": 0,
            "failed_count": 0,
            "warnings": warnings,
            "details": [],
        }

    for idx, row in uploaded_df.iterrows():
        po_val = str(row[po_col]).strip() if po_col in uploaded_df.columns else ""
        part_val = str(row[part_col]).strip() if part_col in uploaded_df.columns else ""
        wip_val = str(row[wip_col]).strip() if wip_col in uploaded_df.columns else ""

        if not wip_val:
            continue

        record_id = None
        matched_by = None

        if po_val and po_val in by_po:
            record_id = by_po[po_val]
            matched_by = "PO"
        elif part_val and part_val in by_part:
            record_id = by_part[part_val]
            matched_by = "PART"

        if not record_id:
            details.append(
                {
                    "row": idx + 1,
                    "po": po_val,
                    "part": part_val,
                    "wip": wip_val,
                    "matched_by": None,
                    "status": "未匹配",
                    "error": "找不到對應的 Teable record",
                }
            )
            failed_count += 1
            continue

        # 組 PATCH 欄位
        fields_to_update: Dict[str, Any] = {}
        wip_field_name = getattr(cfg, "WIP_FIELD_NAME", None) or _guess_wip_field_name(current_df)
        if not wip_field_name:
            details.append(
                {
                    "row": idx + 1,
                    "po": po_val,
                    "part": part_val,
                    "wip": wip_val,
                    "matched_by": matched_by,
                    "status": "未更新",
                    "error": "無法推斷 Teable WIP 欄位名稱 (請在 config 設定 WIP_FIELD_NAME)",
                }
            )
            failed_count += 1
            continue

        fields_to_update[wip_field_name] = wip_val

        success, err = _patch_record(table_id, record_id, fields_to_update)
        if success:
            success_count += 1
            details.append(
                {
                    "row": idx + 1,
                    "po": po_val,
                    "part": part_val,
                    "wip": wip_val,
                    "matched_by": matched_by,
                    "status": "更新成功",
                    "error": "",
                }
            )
        else:
            failed_count += 1
            details.append(
                {
                    "row": idx + 1,
                    "po": po_val,
                    "part": part_val,
                    "wip": wip_val,
                    "matched_by": matched_by,
                    "status": "更新失敗",
                    "error": err,
                }
            )

    return {
        "success_count": success_count,
        "failed_count": failed_count,
        "warnings": warnings,
        "details": details,
    }


def _first_match(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _guess_wip_field_name(df: pd.DataFrame) -> Optional[str]:
    """
    從 current_df 欄位中推測 WIP 欄位名稱
    """
    candidates = getattr(cfg, "WIP_FIELD_CANDIDATES", ["WIP", "Status", "製程", "狀態"])
    for c in candidates:
        if c in df.columns:
            return c
    return None
