# -*- coding: utf-8 -*-
"""
Teable API 模組 - 訂單讀取與進度更新
支援批量更新工廠 WIP 進度到 Teable 主表
"""
from __future__ import annotations

import pandas as pd
import requests
from typing import Optional, Dict, Any, Tuple, List

from config import TABLE_URL, TEABLE_TOKEN, HEADERS, PO_CANDIDATES, WIP_CANDIDATES
from utils import normalize_columns, get_series_by_col


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
            fields = rec.get("fields", {})
            fields["_record_id"] = rec.get("id", "")
            rows.append(fields)

        df = pd.DataFrame(rows)
        df = normalize_columns(df)

        return df, status, text

    except Exception as e:
        return pd.DataFrame(), "EXCEPTION", str(e)


def find_record_id_by_po(
    df: pd.DataFrame, 
    po_value: str, 
    po_col: Optional[str]
) -> Optional[str]:
    """根據 PO 號碼查找 Teable 記錄 ID"""
    if df.empty or not po_col or po_col not in df.columns:
        return None

    po_series = get_series_by_col(df, po_col)
    if po_series is None:
        return None

    matched = df[
        po_series.astype(str).str.strip().str.lower()
        == str(po_value).strip().lower()
    ]

    if matched.empty:
        return None

    if "_record_id" in matched.columns:
        return str(matched.iloc[0]["_record_id"])

    return None


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
    """更新或新增記錄到 Teable"""
    if not po_value:
        return False, "PO is empty"

    record_id = find_record_id_by_po(current_df, po_value, po_col_name)

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


# ============ 新增：批量更新函數 ============

def batch_update_wip_from_excel(
    current_df: pd.DataFrame,
    uploaded_df: pd.DataFrame,
    factory_name: Optional[str] = None
) -> Dict[str, Any]:
    """
    從工廠 Excel 批量更新 WIP 到 Teable
    
    支援多種工廠格式：
    - 全興電子：多列製程，通過日期判斷當前進度
    - Profit Grand：直接 WIP 列
    - 祥竑電子：多列製程 + 數量判斷
    - 西拓電子：簡單 進度 列
    
    Returns:
        Dict with success_count, failed_count, details
    """
    results = {
        'success_count': 0,
        'failed_count': 0,
        'details': [],
        'warnings': []
    }
    
    # 自動檢測 PO 列
    po_col = None
    for candidate in PO_CANDIDATES:
        if candidate in uploaded_df.columns:
            po_col = candidate
            break
    
    if not po_col:
        results['warnings'].append("未找到 PO 列，請確認文件格式")
        return results
    
    # 檢測 WIP 列（直接格式）
    wip_col = None
    for candidate in WIP_CANDIDATES:
        if candidate in uploaded_df.columns:
            wip_col = candidate
            break
    
    # 檢測是否為多列製程格式（全興/祥竑類型）
    process_cols = _detect_process_columns(uploaded_df)
    
    for idx, row in uploaded_df.iterrows():
        try:
            po_value = str(row.get(po_col, "")).strip()
            if not po_value or po_value.lower() in ['nan', 'none', '', 'p/o', '訂單號碼']:
                continue
            
            # 提取 WIP 狀態
            wip_status = None
            
            if wip_col and wip_col in uploaded_df.columns:
                # 簡單格式：直接讀取 WIP 列
                raw_val = row.get(wip_col, "")
                if pd.notna(raw_val) and str(raw_val).strip():
                    wip_status = str(raw_val).strip()
            elif process_cols:
                # 多列製程格式：智能解析當前進度
                wip_status = _parse_wip_from_process(row, process_cols)
            
            if not wip_status:
                results['warnings'].append(f"行 {idx+1}: PO={po_value} 無法解析 WIP 狀態")
                continue
            
            # 查找 Teable 中的記錄
            record_id = find_record_id_by_po(current_df, po_value, po_col)
            
            if not record_id:
                results['failed_count'] += 1
                results['details'].append({
                    'row': idx + 1,
                    'po': po_value,
                    'error': 'Teable 中未找到對應記錄',
                    'wip': wip_status
                })
                continue
            
            # 準備更新數據
            update_fields = {}
            # 使用配置中的第一個 WIP 候選列名
            from config import WIP_CANDIDATES as CFG_WIP
            if CFG_WIP:
                update_fields[CFG_WIP[0]] = wip_status
            
            # 執行更新
            success, msg = patch_record_by_id(record_id, update_fields)
            
            if success:
                results['success_count'] += 1
                results['details'].append({
                    'row': idx + 1,
                    'po': po_value,
                    'status': '更新成功',
                    'wip': wip_status
                })
            else:
                results['failed_count'] += 1
                results['details'].append({
                    'row': idx + 1,
                    'po': po_value,
                    'error': f'API 更新失敗: {msg}',
                    'wip': wip_status
                })
                
        except Exception as e:
            results['failed_count'] += 1
            results['details'].append({
                'row': idx + 1,
                'error': str(e),
                'po': str(row.get(po_col, 'N/A'))
            })
    
    return results


def _detect_process_columns(df: pd.DataFrame) -> List[str]:
    """檢測是否為多列製程格式，返回製程列名列表"""
    # 常見製程關鍵詞
    process_keywords = [
        '下料', '壓合', '鑽孔', '一銅', '外層', '二銅', '防焊', '文字', '成型', '測試',
        '排版', '內層', '內測', '乾膜', 'AOI', '半測', '化金', 'OSP', '化銀', '包裝',
        '蝕刻', '檢測', '銅面處理', '無鉛', '有鉛', '化錫', '發料'
    ]
    
    matched_cols = []
    for col in df.columns:
        col_str = str(col).strip()
        if any(kw in col_str for kw in process_keywords):
            matched_cols.append(col_str)
    
    return matched_cols if len(matched_cols) >= 3 else []


def _parse_wip_from_process(row: pd.Series, process_cols: List[str]) -> Optional[str]:
    """從多列製程格式解析當前 WIP 狀態"""
    # 策略1：查找有日期/數值的最新製程（從後往前找）
    last_active = None
    for col in reversed(process_cols):
        val = str(row.get(col, "")).strip()
        # 排除空白、0、橫線、nan 等無效值
        if val and val.lower() not in ['nan', 'none', '', '-', '0', 'pcs', 'wip', 'wpnl']:
            # 如果是日期格式（如 03-04, 0311）或數字，視為該製程有進度
            if _is_date_like(val) or _is_number_like(val):
                last_active = col
                break
    
    if last_active:
        return f"進行中:{last_active}"
    
    # 策略2：查找空白前的最後一個有值製程
    for i, col in enumerate(process_cols):
        val = str(row.get(col, "")).strip()
        if not val or val.lower() in ['nan', 'none', '', '-', '0']:
            if i > 0:
                return f"已完成:{process_cols[i-1]}"
    
    return None


def _is_date_like(val: str) -> bool:
    """判斷是否為日期格式（如 03-04, 0311, 03/24）"""
    import re
    # 匹配 數字-數字, 數字/數字, 純數字(4位)
    return bool(re.match(r'^\d{2}[-/]\d{2}$|^\d{4}$|^\d{2}$', val.strip()))


def _is_number_like(val: str) -> bool:
    """判斷是否為數字（製程數量）"""
    try:
        float(val.replace(',', ''))
        return True
    except (ValueError, TypeError):
        return False
