# -*- coding: utf-8 -*-
"""
從 Control Tower 的 Teable 主表 (orders DataFrame) 撈取工廠 PO 資料。

設計原則:
- 以「西拓訂單編號」為主 key,撈出該編號的所有 row(可能多品項)
- 不新增 Teable 表,不依賴 customers.json
- 從西拓訂單編號開頭判斷 issuing_company:ET/G → GLOCOM,EW → EUSWAY
- 沿用 Control Tower 現有的 helper(safe_text 等)
"""

from __future__ import annotations
import re
from datetime import datetime, date
from typing import Optional

import pandas as pd

# Control Tower 主表的關鍵欄位名稱
# 從 主表.csv 確認的精確拼寫
COL_GLOCOM_PO = "西拓訂單編號"
COL_CUSTOMER = "客戶"
COL_CUSTOMER_PO = "PO#"
COL_PART_NO = "P/N"
COL_QTY = "Order Q'TY\n (PCS)"  # 注意:欄位名含換行
COL_FACTORY = "工廠"
COL_FACTORY_DUE = "工廠交期"
COL_SHIP_DATE = "Ship date"
COL_SHIP_TO = "Ship to"
COL_SHIP_VIA = "Ship via"
COL_FACTORY_NOTE = "工廠出貨事項"
COL_CUSTOMER_NOTE = "客戶要求注意事項"
COL_LAYER = "板\n層"
COL_NEW_OLD = "新/舊\n料號"
COL_AMOUNT_FACTORY = "銷貨金額"   # 給工廠的金額
COL_AMOUNT_CUSTOMER = "接單金額"   # 給客戶的金額
COL_FACTORY_PO_PDF = "Factory PO PDF"  # ← Jonah 要在 Teable 加的新欄位
COL_RECORD_ID = "_record_id"


def safe_text(v) -> str:
    """跟 Control Tower 的 safe_text 行為一致。"""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def parse_glocom_po_no(po_no: str) -> dict:
    """解析西拓訂單編號,推導出 issuing_company。

    Examples:
        ET1150029-01 → {"order_type": "ET", "issuing_company": "GLOCOM", ...}
        EW1150018-01 → {"order_type": "EW", "issuing_company": "EUSWAY", ...}
        G1150030-01  → {"order_type": "G",  "issuing_company": "GLOCOM", ...}
    """
    s = safe_text(po_no).upper()
    if not s:
        return {"order_type": "", "issuing_company": "GLOCOM"}

    if s.startswith("ET"):
        order_type = "ET"
    elif s.startswith("EW"):
        order_type = "EW"
    elif s.startswith("G"):
        order_type = "G"
    else:
        order_type = "?"

    issuing = "EUSWAY" if order_type == "EW" else "GLOCOM"
    return {
        "order_type": order_type,
        "issuing_company": issuing,
        "raw": po_no,
    }


def list_glocom_po_options(orders_df: pd.DataFrame) -> list[tuple[str, str]]:
    """掃描主表,列出所有可選的西拓訂單編號(去重)。

    Returns:
        [(po_no, display_label), ...]
        例如 [("ET1150029-01", "ET1150029-01 - 宏棋 - Kolff (1 品項)"), ...]
    """
    if orders_df.empty or COL_GLOCOM_PO not in orders_df.columns:
        return []

    # 取所有西拓訂單編號(非空)
    po_series = orders_df[COL_GLOCOM_PO].dropna().astype(str).str.strip()
    po_series = po_series[po_series != ""]

    options = []
    seen_po = set()
    for po in po_series:
        if po in seen_po:
            continue
        seen_po.add(po)

        rows = orders_df[orders_df[COL_GLOCOM_PO].astype(str).str.strip() == po]
        n_items = len(rows)

        # 取第一筆代表這張單
        first = rows.iloc[0]
        factory = safe_text(first.get(COL_FACTORY, ""))
        customer_raw = safe_text(first.get(COL_CUSTOMER, ""))
        # 客戶欄位常被塞物流訊息,取前面真正的客戶名(用空白或中文逗號斷)
        customer_clean = re.split(r"[\s ]{2,}|[,,]", customer_raw, maxsplit=1)[0].strip()
        if len(customer_clean) > 30:
            customer_clean = customer_clean[:30] + "..."

        item_str = f"{n_items} 品項" if n_items > 1 else "1 品項"
        label = f"{po} - {factory or '(無工廠)'} - {customer_clean or '(無客戶)'} ({item_str})"
        options.append((po, label))

    # 倒序(最新的在前)
    options.sort(key=lambda x: x[0], reverse=True)
    return options


def get_po_rows(orders_df: pd.DataFrame, po_no: str) -> pd.DataFrame:
    """撈出指定西拓訂單編號的所有 row。

    Returns:
        DataFrame,可能 0 筆 / 1 筆 / 多筆
    """
    if orders_df.empty or not po_no:
        return pd.DataFrame()
    if COL_GLOCOM_PO not in orders_df.columns:
        return pd.DataFrame()

    target = str(po_no).strip()
    return orders_df[
        orders_df[COL_GLOCOM_PO].astype(str).str.strip() == target
    ].copy().reset_index(drop=True)


def parse_due_date_to_iso(v) -> Optional[date]:
    """主表的工廠交期格式不一致(2026/04/29 / 4/29 / Apr. 29, 26),嘗試解析。"""
    s = safe_text(v)
    if not s:
        return None
    # 試多種格式
    for fmt in (
        "%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d",
        "%m/%d/%Y", "%m/%d",
        "%b. %d, %Y", "%b %d, %Y",
        "%b. %d, %y", "%b %d, %y",
    ):
        try:
            dt = datetime.strptime(s, fmt)
            # 短年(%y 或 %m/%d)補 today.year
            if dt.year < 2000:
                dt = dt.replace(year=datetime.now().year)
            return dt.date()
        except ValueError:
            continue
    # pandas 萬能 fallback
    try:
        ts = pd.to_datetime(s, errors="coerce")
        if pd.notna(ts):
            return ts.date()
    except Exception:
        pass
    return None


def parse_qty_int(v) -> int:
    """數量解析,支援 '504' / '500' / '1,200' 等格式。"""
    s = safe_text(v).replace(",", "")
    if not s:
        return 0
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return 0


def parse_unit_price_float(v) -> float:
    """單價解析,從金額/數量推算或取現有單價欄。"""
    s = safe_text(v).replace(",", "").replace("$", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def build_po_context(po_no: str, rows: pd.DataFrame, factory_master: dict) -> dict:
    """從多 row 組裝成一個完整的 PO 上下文,給 PDF generator 用。

    Args:
        po_no: 西拓訂單編號 ET1150029-01
        rows: 該編號的所有 Teable rows(可能多品項)
        factory_master: 從 factories.json 取得的工廠完整資料
    """
    if rows.empty:
        raise ValueError(f"No rows for PO: {po_no}")

    parsed_no = parse_glocom_po_no(po_no)

    # 取第一筆作為「PO 級別」資料的來源
    first = rows.iloc[0]

    # 客戶欄位清理(去掉物流訊息)
    customer_raw = safe_text(first.get(COL_CUSTOMER, ""))
    customer_clean = re.split(
        r"[\s ]{2,}|[,,] *\d+/\d+", customer_raw, maxsplit=1
    )[0].strip()

    # 採購日期:從工廠下單日期取
    order_date_raw = safe_text(first.get("工廠下單日期", "")) or \
                     safe_text(first.get("客戶下單日期", ""))
    order_date_obj = parse_due_date_to_iso(order_date_raw) or datetime.now().date()

    # 組品項列表
    items = []
    for _, row in rows.iterrows():
        qty = parse_qty_int(row.get(COL_QTY, 0))
        # 從金額反推單價(主表沒有單獨「單價」欄)
        amount_factory = parse_unit_price_float(row.get(COL_AMOUNT_FACTORY, 0))
        unit_price = (amount_factory / qty) if qty > 0 else 0.0

        due_date = parse_due_date_to_iso(row.get(COL_FACTORY_DUE, "")) or \
                   parse_due_date_to_iso(row.get(COL_SHIP_DATE, ""))

        # 規格 = 客戶要求注意事項 + 工廠出貨事項 拼接(因為主表沒有專門的規格欄)
        # 使用者在 UI 上可以再編輯
        spec_parts = []
        new_old = safe_text(row.get(COL_NEW_OLD, ""))
        layer = safe_text(row.get(COL_LAYER, ""))
        if new_old:
            spec_parts.append(f"{new_old}料號")
        if layer:
            try:
                layer_int = int(float(layer))
                spec_parts.append(f"{layer_int}L")
            except (ValueError, TypeError):
                pass
        cust_note = safe_text(row.get(COL_CUSTOMER_NOTE, ""))
        if cust_note:
            spec_parts.append(cust_note)
        spec_text = "\n".join(spec_parts)

        items.append({
            "part_number": safe_text(row.get(COL_PART_NO, "")),
            "spec_text": spec_text,
            "quantity": qty,
            "panel_qty": None,  # 主表沒有 panel,使用者 UI 可補
            "unit_price": unit_price,
            "amount": amount_factory if amount_factory > 0 else (qty * unit_price),
            "delivery_date": due_date,
            "delivery_note": "",
            "_record_id": safe_text(row.get(COL_RECORD_ID, "")),
            "_customer_po_no": safe_text(row.get(COL_CUSTOMER_PO, "")),
        })

    factory_short = safe_text(first.get(COL_FACTORY, ""))

    return {
        "po_no": po_no,
        "order_type": parsed_no["order_type"],
        "issuing_company": parsed_no["issuing_company"],
        "order_date": order_date_obj,
        "customer_name": customer_clean,
        "customer_po_no": safe_text(first.get(COL_CUSTOMER_PO, "")),
        "factory_short": factory_short,
        "factory": factory_master,  # 完整工廠資料
        "ship_to": safe_text(first.get(COL_SHIP_TO, "")),
        "ship_via": safe_text(first.get(COL_SHIP_VIA, "")),
        "factory_note": safe_text(first.get(COL_FACTORY_NOTE, "")),
        "items": items,
        "total_amount": sum(it["amount"] for it in items),
        "currency": factory_master.get("default_currency", "NT$") if factory_master else "NT$",
        "payment_terms": factory_master.get("default_payment_terms", "") if factory_master else "",
        "shipment_method": factory_master.get("default_shipment", "待通知") if factory_master else "待通知",
        "ship_to_default": factory_master.get("default_ship_to", "待通知") if factory_master else "待通知",
        "_record_ids": [it["_record_id"] for it in items if it["_record_id"]],
    }
