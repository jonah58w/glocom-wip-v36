# -*- coding: utf-8 -*-
"""
Teable API 模組 - 訂單讀取與進度更新
支援批量更新工廠 WIP 進度到 Teable 主表
"""

from __future__ import annotations

from typing import Optional, Dict, Any, Tuple, List

import pandas as pd
import requests

from config import (
    TABLE_URL,
    TEABLE_TOKEN,
    PO_CANDIDATES,
    PART_CANDIDATES,
    WIP_CANDIDATES,
)
from utils import normalize_columns, get_series_by_col


HEADERS = {
    "Authorization": f"Bearer {TEABLE_TOKEN}",
    "Content-Type": "application/json",
}


# =========================================
# 基本工具
# =========================================

def _clean_text(v: Any) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def _norm_key(v: Any) -> str:
    s = _clean_text(v)
    if not s:
        return ""
    return s.replace(" ", "").replace("-", "").replace("_", "").lower()


def _get_first_existing_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _valid_key_value(v: Any) -> bool:
    s = _clean_text(v)
    if not s:
        return False
    return s.lower() not in {
        "nan", "none", "null", "p/o", "po", "訂單號碼", "訂單編號", "料號", "part no", "p/n"
    }


# =========================================
# 讀取 Teable 主表
# =========================================

def load_orders() -> Tuple[pd.DataFrame, Any, str]:
    """從 Teable 載入訂單數據"""
    if not TEABLE_TOKEN:
        return pd.DataFrame(), "NO_TOKEN", "TEABLE_TOKEN is empty"

    try:
        response = requests.get(
            TABLE_URL,
            headers=HEADERS,
            params={
                "fieldKeyType": "name",
                "cellFormat": "text",
                "take": 1000,
            },
            timeout=30,
        )

        status = response.status_code
        text = response.text

        if status != 200:
            return pd.DataFrame(), status, text

        data = response.json()
        records = data.get("records", [])

        rows = []
        for rec in records:
            fields = rec.get("fields", {}) or {}
            fields["_record_id"] = rec.get("id", "")
            rows.append(fields)

        df = pd.DataFrame(rows)
        df = normalize_columns(df)

        return df, status, text

    except Exception as e:
        return pd.DataFrame(), "EXCEPTION", str(e)


# =========================================
# 記錄查找
# =========================================

def find_record_id_by_po(
    df: pd.DataFrame,
    po_value: str,
    po_col: Optional[str]
) -> Optional[str]:
    """根據指定 PO 欄位查找記錄 ID（舊介面保留）"""
    if df.empty or not po_col or po_col not in df.columns:
        return None

    po_series = get_series_by_col(df, po_col)
    if po_series is None:
        return None

    target = _norm_key(po_value)
    if not target:
        return None

    matched = df[po_series.astype(str).map(_norm_key) == target]

    if matched.empty:
        return None

    if "_record_id" in matched.columns:
        return str(matched.iloc[0]["_record_id"])

    return None


def find_record_id_by_part(
    df: pd.DataFrame,
    part_value: str,
    part_col: Optional[str]
) -> Optional[str]:
    """根據指定 Part 欄位查找記錄 ID"""
    if df.empty or not part_col or part_col not in df.columns:
        return None

    part_series = get_series_by_col(df, part_col)
    if part_series is None:
        return None

    target = _norm_key(part_value)
    if not target:
        return None

    matched = df[part_series.astype(str).map(_norm_key) == target]

    if matched.empty:
        return None

    if "_record_id" in matched.columns:
        return str(matched.iloc[0]["_record_id"])

    return None


def find_record_id_by_keys(
    current_df: pd.DataFrame,
    po_value: str = "",
    part_value: str = "",
) -> Tuple[Optional[str], Optional[str]]:
    """
    依序用：
    1. PO
    2. Part No / 料號
    查找 Teable record_id

    回傳:
    (record_id, matched_by)
    matched_by: "PO" / "PART" / None
    """
    if current_df is None or current_df.empty:
        return None, None

    current_po_col = _get_first_existing_column(current_df, PO_CANDIDATES)
    current_part_col = _get_first_existing_column(current_df, PART_CANDIDATES)

    if _valid_key_value(po_value) and current_po_col:
        rid = find_record_id_by_po(current_df, po_value, current_po_col)
        if rid:
            return rid, "PO"

    if _valid_key_value(part_value) and current_part_col:
        rid = find_record_id_by_part(current_df, part_value, current_part_col)
        if rid:
            return rid, "PART"

    return None, None


# =========================================
# API 更新
# =========================================

def patch_record_by_id(record_id: str, payload_fields: Dict[str, Any]) -> Tuple[bool, str]:
    """更新單條記錄"""
    try:
        r = requests.patch(
            f"{TABLE_URL}/{record_id}",
            headers=HEADERS,
            json={"record": {"fields": payload_fields}},
            timeout=30,
        )

        if r.status_code in (200, 201):
            return True, r.text

        return False, f"{r.status_code} | {r.text}"

    except Exception as e:
        return False, str(e)


def upsert_to_teable(
    current_df: pd.DataFrame,
    po_col_name: str,
    po_value: str,
    updates: Dict[str, Any]
) -> Tuple[bool, str]:
    """
    舊介面保留：
    以 PO 為主的更新或新增
    """
    if not _valid_key_value(po_value):
        return False, "PO is empty"

    # 這裡改為自動用 current_df 內真正的 PO 欄查找
    record_id, _ = find_record_id_by_keys(current_df=current_df, po_value=po_value)

    payload_fields = dict(updates)
    payload_fields[po_col_name] = po_value

    try:
        if record_id:
            r = requests.patch(
                f"{TABLE_URL}/{record_id}",
                headers=HEADERS,
                json={"record": {"fields": payload_fields}},
                timeout=30,
            )
        else:
            r = requests.post(
                TABLE_URL,
                headers=HEADERS,
                json={"records": [{"fields": payload_fields}]},
                timeout=30,
            )

        if r.status_code in (200, 201):
            return True, r.text

        return False, f"{r.status_code} | {r.text}"

    except Exception as e:
        return False, str(e)


def update_working_orders_local(
    working_orders: pd.DataFrame,
    record_id: str,
    payload_fields: Dict[str, Any]
) -> pd.DataFrame:
    """本地緩存更新"""
    if working_orders.empty or "_record_id" not in working_orders.columns:
        return working_orders

    mask = working_orders["_record_id"].astype(str) == str(record_id)

    if not mask.any():
        return working_orders

    for field_name, field_value in payload_fields.items():
        if field_name not in working_orders.columns:
            working_orders[field_name] = ""
        working_orders.loc[mask, field_name] = field_value

    return working_orders


# =========================================
# WIP 解析
# =========================================

def _detect_process_columns(df: pd.DataFrame) -> List[str]:
    """檢測是否為多列製程格式，返回製程列名列表"""
    if df is None or df.empty:
        return []

    process_keywords = [
        "下料", "壓合", "鑽孔", "一銅", "外層", "二銅", "防焊", "文字", "成型", "測試",
        "排版", "內層", "內測", "乾膜", "AOI", "半測", "化金", "OSP", "化銀", "包裝",
        "蝕刻", "檢測", "銅面處理", "無鉛", "有鉛", "化錫", "發料",
        "內乾", "內蝕", "黑化", "壓板", "沉銅", "電鍍", "噴錫", "V-CUT", "出貨"
    ]

    matched_cols = []
    for col in df.columns:
        col_str = _clean_text(col)
        if any(kw in col_str for kw in process_keywords):
            matched_cols.append(col_str)

    return matched_cols if len(matched_cols) >= 3 else []


def _parse_wip_from_process(row: pd.Series, process_cols: List[str]) -> Optional[str]:
    """
    從多列製程格式解析當前 WIP 狀態
    策略：
    - 由後往前找最後一個有效製程
    - 遇到包裝/出貨等則直接回傳較後段狀態
    """
    if not process_cols:
        return None

    invalid_values = {
        "", "-", "--", "0", "0.0", "nan", "none", "null", "pcs", "pcs.", "x"
    }

    last_active = None

    for col in reversed(process_cols):
        val = _clean_text(row.get(col, ""))
        if not val:
            continue
        if val.lower() in invalid_values:
            continue

        last_active = col
        break

    return last_active


# =========================================
# 批量更新
# =========================================

def batch_update_wip_from_excel(
    current_df: pd.DataFrame,
    uploaded_df: pd.DataFrame,
    factory_name: Optional[str] = None
) -> Dict[str, Any]:
    """
    從工廠 Excel / parser 輸出 批量更新 WIP 到 Teable

    支援：
    1. 用 PO 匹配
    2. 若沒有 PO，改用 Part No / 料號 匹配
    3. 若沒有直接 WIP 欄，嘗試用製程欄推算
    """
    results = {
        "success_count": 0,
        "failed_count": 0,
        "details": [],
        "warnings": [],
    }

    if uploaded_df is None or uploaded_df.empty:
        results["warnings"].append("上傳資料為空，無法更新")
        return results

    current_df = normalize_columns(current_df.copy()) if current_df is not None else pd.DataFrame()
    uploaded_df = normalize_columns(uploaded_df.copy())

    # 上傳資料的主鍵欄
    uploaded_po_col = _get_first_existing_column(uploaded_df, PO_CANDIDATES)
    uploaded_part_col = _get_first_existing_column(uploaded_df, PART_CANDIDATES)

    if not uploaded_po_col and not uploaded_part_col:
        results["warnings"].append("未找到 PO 或料號欄，請確認文件格式")
        return results

    # 上傳資料的 WIP 欄
    uploaded_wip_col = _get_first_existing_column(uploaded_df, WIP_CANDIDATES)

    # 製程欄偵測
    process_cols = _detect_process_columns(uploaded_df)

    # Teable 主表實際欄位
    current_wip_col = _get_first_existing_column(current_df, WIP_CANDIDATES)
    if not current_wip_col:
        current_wip_col = WIP_CANDIDATES[0] if WIP_CANDIDATES else "WIP"

    for idx, row in uploaded_df.iterrows():
        try:
            po_value = _clean_text(row.get(uploaded_po_col, "")) if uploaded_po_col else ""
            part_value = _clean_text(row.get(uploaded_part_col, "")) if uploaded_part_col else ""

            if not _valid_key_value(po_value) and not _valid_key_value(part_value):
                continue

            # 解析 WIP
            wip_status = None

            if uploaded_wip_col and uploaded_wip_col in uploaded_df.columns:
                raw_val = row.get(uploaded_wip_col, "")
                if pd.notna(raw_val) and _clean_text(raw_val):
                    wip_status = _clean_text(raw_val)

            if not wip_status and process_cols:
                wip_status = _parse_wip_from_process(row, process_cols)

            if not wip_status:
                results["warnings"].append(
                    f"行 {idx + 1}: "
                    f"PO={po_value or '-'} / PART={part_value or '-'} 無法解析 WIP 狀態"
                )
                continue

            # 查找 Teable 記錄
            record_id, matched_by = find_record_id_by_keys(
                current_df=current_df,
                po_value=po_value,
                part_value=part_value,
            )

            if not record_id:
                results["failed_count"] += 1
                results["details"].append({
                    "row": idx + 1,
                    "po": po_value,
                    "part": part_value,
                    "error": "Teable 中未找到對應記錄",
                    "wip": wip_status,
                    "matched_by": None,
                })
                continue

            # 準備更新欄位
            update_fields = {
                current_wip_col: wip_status
            }

            success, msg = patch_record_by_id(record_id, update_fields)

            if success:
                results["success_count"] += 1
                results["details"].append({
                    "row": idx + 1,
                    "po": po_value,
                    "part": part_value,
                    "status": "更新成功",
                    "wip": wip_status,
                    "matched_by": matched_by,
                })
            else:
                results["failed_count"] += 1
                results["details"].append({
                    "row": idx + 1,
                    "po": po_value,
                    "part": part_value,
                    "error": f"API 更新失敗: {msg}",
                    "wip": wip_status,
                    "matched_by": matched_by,
                })

        except Exception as e:
            results["failed_count"] += 1
            results["details"].append({
                "row": idx + 1,
                "po": _clean_text(row.get(uploaded_po_col, "N/A")) if uploaded_po_col else "",
                "part": _clean_text(row.get(uploaded_part_col, "N/A")) if uploaded_part_col else "",
                "error": str(e),
            })

    return results
