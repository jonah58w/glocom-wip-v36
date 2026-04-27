# -*- coding: utf-8 -*-
"""
建立工廠 PO 頁面 v3.5。

主要更新(v3.5):
- ★ 三層編號防撞機制(避免人工手 key 跟系統建單撞號):
  第 1 層【即時撈】:進頁面時直接呼叫 Teable API 撈最新編號(不靠側邊欄 Refresh)
  第 2 層【撞號警告】:產 PDF 前再撈一次,撞號就紅色警告擋下,自動換新編號
  第 3 層【寫入鎖】:寫回 Teable 前最後檢查一次,真的撞才往上跳
- ★ Step 3 加「🔄 重新計算編號」按鈕(不用整頁 reload)

v3.4 變更:
- 規格優先從 data/spec_history.json(歷史 RTF/DOCX)抓
- 找不到再從 Teable 主表「客戶要求注意事項」抓

v3.3 變更:
- 規格頂部加「舊料號 / 新料號」radio
- 選舊料號自動帶歷史注意事項

v3.2 變更:
- 工廠下拉用 sort_priority 排序

v3.1 變更:
- 拿掉「規格沒填→fallback 客戶描述」邏輯,避免規格欄重複印料號
- 規格沒填時顯示紅色錯誤,阻擋產 PDF / 寫回 Teable
"""

from __future__ import annotations
import json
import re
import traceback
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

from core.customer_po_parser import (
    parse_customer_po_from_pdf,
    POItem,
    ParsedPO,
)
from core.teable_query import (
    parse_glocom_po_no,
    safe_text,
    COL_GLOCOM_PO,
)


# ─── 設定 ────────────────────────────────────────
HERE = Path(__file__).parent
FACTORIES_JSON = HERE / "data" / "factories.json"
SPEC_HISTORY_JSON = HERE / "data" / "spec_history.json"


# 新料號完整規格的選項
SPEC_OPTIONS_FULL = {
    "Material": ["2L", "4L", "6L", "8L", "10L", "Aluminum", "FR4", "Rogers"],
    "Tg": ["Tg130", "Tg150", "Tg170", "Tg180"],
    "Board thickness": ["0.4mm", "0.6mm", "0.8mm", "1.0mm", "1.2mm", "1.6mm", "2.0mm", "2.4mm", "3.2mm"],
    "Copper": ["1oz/1oz", "1/2oz/1/2oz", "2oz/2oz", "3oz/3oz", "Ext: 1oz all layers"],
    "Surface Finish": ["ENIG", "Lead-free HASL", "HASL", "Immersion Gold", "OSP", "Hard Gold", "Immersion Silver"],
    "S/M": ["Green", "Matte Green", "Red", "Blue", "Black", "Matte Black", "White"],
    "S/L": ["White", "Black", "Yellow"],
}


# ─── 工廠主檔 ────────────────────────────────────
def _strip_placeholder_brackets(text: str) -> str:
    """剝掉字串中的 [請補...] 方框註記,例如『優技電子 [請補完整名]』→『優技電子』"""
    if not text:
        return ""
    return re.sub(r"\s*\[[^\]]*\]\s*", "", str(text)).strip()


def load_factories() -> dict:
    if not FACTORIES_JSON.exists():
        return {}
    try:
        with open(FACTORIES_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
        all_factories = data.get("factories", {})
        active = {k: v for k, v in all_factories.items() if v.get("is_active")}
        # 把 factory_name 的方框註記剝掉(避免印在簽名欄)
        for v in active.values():
            v["factory_name"] = _strip_placeholder_brackets(v.get("factory_name", ""))
        return active
    except Exception as e:
        st.error(f"factories.json 讀取失敗:{e}")
        return {}


# ─── 字首邏輯 ────────────────────────────────────
def derive_prefix(issuing_company: str, factory_region: str) -> str:
    issuing = (issuing_company or "").upper()
    region = (factory_region or "").lower()
    if issuing == "GLOCOM":
        return "G"
    if issuing == "EUSWAY":
        if region in ("china", "hk", "hong kong"):
            return "EW"
        return "ET"
    return "G"


def display_prefix(internal_prefix: str) -> str:
    return "GC" if internal_prefix == "G" else internal_prefix


def internal_prefix(display: str) -> str:
    return "G" if display == "GC" else display


# ─── ★ v3.5: 即時撈 Teable API ─────────────
def fetch_all_po_numbers_from_teable(table_url: str, headers: dict) -> set[str]:
    """
    直接打 Teable API 撈所有訂單編號(不靠側邊欄 Refresh 的 cache)。
    
    回傳:set of all 西拓訂單編號 字串
    """
    po_numbers = set()
    page_token = None
    pn = COL_GLOCOM_PO  # "西拓訂單編號"
    
    try:
        while True:
            params = {
                "fieldKeyType": "name",
                "cellFormat": "text",
                "take": 1000,
                "viewId": "viwGqqwdvy5WkaqfQuJ",  # 「API - All Records」
                "projection": [pn],  # 只撈這一欄省流量
            }
            if page_token:
                params["nextPageToken"] = page_token
            
            r = requests.get(table_url, headers=headers, params=params, timeout=30)
            if r.status_code != 200:
                break
            
            data = r.json()
            records = data.get("records", []) or []
            for rec in records:
                fields = rec.get("fields", {}) or {}
                val = str(fields.get(pn, "") or "").strip()
                if val:
                    po_numbers.add(val)
            
            page_token = data.get("nextPageToken")
            if not page_token:
                break
    except Exception as e:
        st.warning(f"⚠️ 即時撈 Teable 編號失敗:{e}(改用側邊欄快取)")
    
    return po_numbers


def calc_next_po_number(po_numbers_set: set[str], prefix: str) -> str:
    """
    計算下一個訂單編號。
    
    重點:流水號要**全局唯一**,不論字首 G/ET/EW/GC。
    例如 Teable 上有 G1150030-01,則下一個 ET 編號也不能用 1150030,
    必須是 ET1150031-01。
    
    Args:
        po_numbers_set: 所有現有訂單編號的 set
        prefix: 內部字首 (G / ET / EW)
    
    Returns: 下一個訂單編號,例如 G1150031-01
    """
    if not po_numbers_set:
        roc_year = datetime.now().year - 1911
        return f"{prefix}{roc_year}0001-01"

    pattern = re.compile(r"^[A-Z]+(\d{3})(\d{4})-(\d+)$")
    max_year_serial = (0, 0)

    for po in po_numbers_set:
        m = pattern.match(po.strip())
        if m:
            year = int(m.group(1))
            serial = int(m.group(2))
            if (year, serial) > max_year_serial:
                max_year_serial = (year, serial)

    if max_year_serial == (0, 0):
        roc_year = datetime.now().year - 1911
        return f"{prefix}{roc_year}0001-01"

    year, serial = max_year_serial
    return f"{prefix}{year}{serial + 1:04d}-01"


def calc_next_po_number_from_df(orders_df: pd.DataFrame, prefix: str) -> str:
    """從 DataFrame (側邊欄快取)算下一編號 - fallback 用"""
    if orders_df.empty or COL_GLOCOM_PO not in orders_df.columns:
        roc_year = datetime.now().year - 1911
        return f"{prefix}{roc_year}0001-01"
    po_set = set()
    for raw_po in orders_df[COL_GLOCOM_PO].dropna().astype(str):
        po_set.add(raw_po.strip())
    return calc_next_po_number(po_set, prefix)


# ─── Teable 寫入 ─────────────────────────────────
def create_teable_records(table_url: str, headers: dict, fields_list: list) -> dict:
    result = {"success": 0, "failed": 0, "errors": [], "created_ids": []}
    if not fields_list:
        return result

    payload = {
        "fieldKeyType": "name",
        "typecast": True,
        "records": [{"fields": f} for f in fields_list],
    }

    try:
        r = requests.post(table_url, headers=headers, json=payload, timeout=30)
        if r.status_code in (200, 201):
            data = r.json()
            records = data.get("records", []) or []
            result["success"] = len(records) if records else len(fields_list)
            result["created_ids"] = [rec.get("id", "") for rec in records]
        else:
            result["failed"] = len(fields_list)
            result["errors"].append(f"HTTP {r.status_code}: {r.text[:500]}")
    except Exception as e:
        result["failed"] = len(fields_list)
        result["errors"].append(f"Exception: {e}")

    return result


# ─── Teable 查同料號歷史 ─────────────────────────
def load_spec_history() -> dict:
    """讀 data/spec_history.json (從歷史 RTF/DOCX 訂單檔抽出來的注意事項對應表)"""
    if not SPEC_HISTORY_JSON.exists():
        return {}
    try:
        with open(SPEC_HISTORY_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("spec_history", {}) or {}
    except Exception as e:
        st.warning(f"spec_history.json 讀取失敗:{e}")
        return {}


def fetch_previous_spec(orders: pd.DataFrame, part_number: str, factory_short: str = ""):
    """
    查同 P/N 最近一筆訂單的規格。
    
    優先順序:
    1. data/spec_history.json (從歷史 RTF/DOCX 解析)
    2. Teable 主表「客戶要求注意事項」欄
    
    同工廠優先(如果有指定 factory_short)。
    """
    results = []
    pn_target = str(part_number).strip()
    if not pn_target:
        return results
    
    # ─── 1. 從 spec_history.json 查 ─
    spec_history = load_spec_history()
    if pn_target in spec_history:
        history_records = spec_history[pn_target].get("history", [])
        # 同工廠優先
        if factory_short:
            same_factory = [r for r in history_records if r.get("factory") == factory_short]
            other = [r for r in history_records if r.get("factory") != factory_short]
            ordered = same_factory + other
        else:
            ordered = history_records
        for r in ordered:
            results.append({
                "po_no": r.get("po", ""),
                "spec": r.get("spec_text", ""),
                "factory": r.get("factory", ""),
                "date": r.get("date", ""),
                "source": "history_file",
            })
    
    # ─── 2. 若 spec_history 沒命中,再查 Teable 主表 ─
    if not results and not orders.empty:
        pn_col = "P/N"
        spec_col = "客戶要求注意事項"
        po_col = COL_GLOCOM_PO
        date_col = "工廠下單日期"
        factory_col = "工廠"
        
        if pn_col in orders.columns:
            matches = orders[orders[pn_col].astype(str).str.strip() == pn_target].copy()
            if not matches.empty:
                if date_col in matches.columns:
                    matches[date_col + "_dt"] = pd.to_datetime(matches[date_col], errors="coerce")
                    matches = matches.sort_values(date_col + "_dt", ascending=False)
                
                seen_pos = set()
                for _, row in matches.iterrows():
                    po = str(row.get(po_col, "")).strip() if po_col in row.index else ""
                    if not po or po in seen_pos:
                        continue
                    seen_pos.add(po)
                    spec = str(row.get(spec_col, "")).strip() if spec_col in row.index else ""
                    if pd.isna(spec) or spec.lower() == "nan":
                        spec = ""
                    factory = str(row.get(factory_col, "")).strip() if factory_col in row.index else ""
                    date_str = str(row.get(date_col, "")).strip() if date_col in row.index else ""
                    results.append({
                        "po_no": po,
                        "spec": spec,
                        "factory": factory,
                        "date": date_str,
                        "source": "teable",
                    })
                    if len(results) >= 5:
                        break
    
    return results


# ─── 規格組字串 ──────────────────────────────────
def build_full_spec_oneline(selections: dict, extra_text: str) -> str:
    parts = []
    for cat, val in selections.items():
        if val:
            if cat in ("Tg",):
                parts.append(val)
            else:
                parts.append(f"{cat}: {val}")
    if extra_text and extra_text.strip():
        parts.append(extra_text.strip())
    return "; ".join(parts)


def build_po_context_from_new_order(
    new_po_no, parsed, factory_data, factory_short, issuing_company,
    factory_unit_prices, factory_due_dates, item_specs,
    purchase_responsible, order_date, is_revised,
):
    items = []
    for it in parsed.items:
        f_price = float(factory_unit_prices.get(it.part_number, 0.0))
        f_due = factory_due_dates.get(it.part_number)
        spec_text = (item_specs.get(it.part_number, "") or "").strip()
        items.append({
            "part_number": it.part_number,
            "spec_text": spec_text,
            "quantity": it.quantity,
            "panel_qty": None,
            "unit_price": f_price,
            "amount": round(f_price * it.quantity, 2),
            "delivery_date": f_due,
            "delivery_note": "",
        })

    return {
        "po_no": new_po_no,
        "order_type": parse_glocom_po_no(new_po_no).get("order_type", ""),
        "issuing_company": issuing_company,
        "order_date": order_date,
        "customer_name": parsed.customer_name,
        "customer_po_no": parsed.customer_po_no,
        "factory_short": factory_short,
        "factory": factory_data,
        "items": items,
        "total_amount": sum(it["amount"] for it in items),
        "currency": factory_data.get("default_currency", "NT$"),
        "payment_terms": factory_data.get("default_payment_terms", ""),
        "shipment_method": factory_data.get("default_shipment", "待通知"),
        "ship_to_default": factory_data.get("default_ship_to", "待通知"),
        "purchase_responsible": purchase_responsible,
        "is_revised": is_revised,
    }


def _reset_form():
    keys_to_clear = [k for k in list(st.session_state.keys())
                     if k.startswith("fpo_") or k == "customer_po_uploader"]
    for k in keys_to_clear:
        try:
            del st.session_state[k]
        except KeyError:
            pass


# ─── Step 5 規格輸入(每品項)──────────────────
def render_spec_input_for_item(idx: int, item, orders: pd.DataFrame) -> str:
    final_key = f"fpo_spec_final_{idx}"
    if final_key not in st.session_state:
        st.session_state[final_key] = ""

    type_key = f"fpo_spec_type_{idx}"
    if type_key not in st.session_state:
        st.session_state[type_key] = "舊料號"

    applied_old_default_key = f"fpo_spec_old_applied_{idx}"

    with st.container(border=True):
        st.markdown(f"**品項 {idx + 1}:`{item.part_number}`**(數量 {item.quantity:,})")

        cols_type = st.columns([2, 5])
        with cols_type[0]:
            new_type = st.radio(
                "料號類型", ["舊料號", "新料號"],
                key=type_key, horizontal=True,
            )

        if new_type == "舊料號":
            with cols_type[1]:
                st.caption("👉 系統會從歷史訂單檔(同料號最近一筆)自動帶入注意事項。請於最終規格框直接編輯。")

            if not st.session_state.get(applied_old_default_key, False):
                current_factory = st.session_state.get("fpo_factory_short", "")
                hist_hits = fetch_previous_spec(orders, item.part_number, factory_short=current_factory)
                hist_spec = ""
                hist_po = ""
                hist_factory = ""
                hist_source = ""
                if hist_hits:
                    for h in hist_hits:
                        if h.get("spec"):
                            hist_spec = h["spec"].strip()
                            hist_po = h.get("po_no", "")
                            hist_factory = h.get("factory", "")
                            hist_source = h.get("source", "")
                            break

                if hist_spec:
                    default_text = f"舊料號\n\n注意:\n{hist_spec}"
                else:
                    default_text = "舊料號"

                st.session_state[final_key] = default_text
                st.session_state[applied_old_default_key] = True
                st.session_state[f"_fpo_old_hist_po_{idx}"] = hist_po
                st.session_state[f"_fpo_old_hist_factory_{idx}"] = hist_factory
                st.session_state[f"_fpo_old_hist_source_{idx}"] = hist_source

            hist_po_shown = st.session_state.get(f"_fpo_old_hist_po_{idx}", "")
            hist_factory_shown = st.session_state.get(f"_fpo_old_hist_factory_{idx}", "")
            hist_source_shown = st.session_state.get(f"_fpo_old_hist_source_{idx}", "")
            if hist_po_shown:
                source_label = "歷史訂單檔" if hist_source_shown == "history_file" else "Teable 主表"
                factory_info = f" / {hist_factory_shown}" if hist_factory_shown else ""
                st.caption(
                    f"📌 注意事項來源:**{hist_po_shown}**{factory_info}({source_label})"
                )
            else:
                st.caption("⚠️ 沒找到同料號的歷史訂單,請手動填注意事項。")

        else:  # 新料號
            with cols_type[1]:
                st.caption("👉 新料號:用下方 3 個工具填完整規格,或直接在最終規格框打字。")
            st.session_state[applied_old_default_key] = False

        if new_type == "新料號":
            tab_a, tab_b, tab_c = st.tabs([
                "⚡ 從前一張同料號帶入",
                "📋 多列貼上(從 ERP)",
                "✏️ 7 選單 + 補充",
            ])

            with tab_a:
                st.caption(f"從 Teable 主表 / 歷史訂單檔查 P/N『{item.part_number}』")
                if st.button(f"🔍 查詢", key=f"fpo_query_btn_{idx}"):
                    hits = fetch_previous_spec(orders, item.part_number)
                    st.session_state[f"fpo_query_hits_{idx}"] = hits

                hits = st.session_state.get(f"fpo_query_hits_{idx}")
                if hits is not None:
                    if not hits:
                        st.warning(f"沒找到 P/N『{item.part_number}』的歷史訂單")
                    else:
                        st.info(f"找到 **{len(hits)}** 筆歷史訂單(由新到舊)")
                        for h_idx, h in enumerate(hits):
                            with st.expander(
                                f"📦 {h['po_no']}  |  {h['factory']}  |  {h['date']}",
                                expanded=(h_idx == 0),
                            ):
                                if h["spec"]:
                                    st.code(h["spec"], language=None)
                                    if st.button(
                                        f"⬇ 套用此規格",
                                        key=f"fpo_apply_hist_{idx}_{h_idx}",
                                        type="primary",
                                    ):
                                        st.session_state[final_key] = h["spec"]
                                        st.rerun()
                                else:
                                    st.caption("(此筆訂單沒填規格)")

            with tab_b:
                st.caption("從 ERP 一列一列複製貼上(最多 8 列),按下「組合」會用『; 』串成一列")
                paste_lines = []
                cols_paste = st.columns(2)
                for line_idx in range(8):
                    target = cols_paste[line_idx % 2]
                    with target:
                        line = st.text_input(
                            f"第 {line_idx + 1} 列",
                            key=f"fpo_paste_{idx}_{line_idx}",
                            placeholder="例如: 8L, FR4, 1.6mm, 1up",
                        )
                    if line and line.strip():
                        paste_lines.append(line.strip())

                cols_btn_b = st.columns([3, 1])
                with cols_btn_b[0]:
                    if paste_lines:
                        preview = "; ".join(paste_lines)
                        st.caption(f"預覽: `{preview}`")
                with cols_btn_b[1]:
                    if st.button(
                        "⬇ 組合並套用",
                        key=f"fpo_combine_paste_{idx}",
                        disabled=not paste_lines,
                        use_container_width=True,
                        type="primary",
                    ):
                        st.session_state[final_key] = "; ".join(paste_lines)
                        st.rerun()

            with tab_c:
                st.caption("勾選 7 項基本規格 + 補充其他要求。會組成一列。")
                cols_a = st.columns(4)
                cols_b = st.columns(4)
                selections = {}
                cats = list(SPEC_OPTIONS_FULL.keys())
                for i, cat in enumerate(cats):
                    target_col = cols_a[i] if i < 4 else cols_b[i - 4]
                    with target_col:
                        selections[cat] = st.selectbox(
                            cat, [""] + SPEC_OPTIONS_FULL[cat],
                            key=f"fpo_spec_full_{idx}_{cat}",
                        )

                extra = st.text_input(
                    "其他規格 / 補充",
                    placeholder="例如: 排版 4up panel; 須加西拓 UL Logo + Date Code (YYWW)",
                    key=f"fpo_spec_full_extra_{idx}",
                )

                preview = build_full_spec_oneline(selections, extra)
                cols_btn_c = st.columns([3, 1])
                with cols_btn_c[0]:
                    if preview:
                        st.caption(f"預覽: `{preview}`")
                with cols_btn_c[1]:
                    if st.button(
                        "⬇ 套用",
                        key=f"fpo_apply_full_{idx}",
                        disabled=not preview,
                        use_container_width=True,
                        type="primary",
                    ):
                        st.session_state[final_key] = preview
                        st.rerun()

        st.markdown("##### 📋 最終規格(印在 PO 上,可直接編輯)")
        if new_type == "舊料號":
            spec_placeholder = "舊料號\n注意:\n1. ...\n2. ..."
        else:
            spec_placeholder = "(尚未套用任何規格,可手動輸入或從上方工具套用)"

        st.text_area(
            "最終規格",
            key=final_key,
            label_visibility="collapsed",
            height=120,
            placeholder=spec_placeholder,
        )

    return st.session_state[final_key]


# ─── ★ v3.5: 即時撈 Teable 編號 + 計算下一編號 (有 cache) ─
def get_live_po_numbers(table_url: str, headers: dict, force_refresh: bool = False) -> set[str]:
    """
    即時撈 Teable 上所有訂單編號。30 秒 cache 避免每次 rerun 都打 API。
    
    Args:
        force_refresh: True 時強制重撈
    """
    cache_key = "_fpo_live_po_numbers"
    cache_time_key = "_fpo_live_po_numbers_t"
    
    now = datetime.now().timestamp()
    cached_time = st.session_state.get(cache_time_key, 0)
    cached_set = st.session_state.get(cache_key)
    
    if not force_refresh and cached_set is not None and (now - cached_time) < 30:
        return cached_set
    
    fresh_set = fetch_all_po_numbers_from_teable(table_url, headers)
    st.session_state[cache_key] = fresh_set
    st.session_state[cache_time_key] = now
    return fresh_set


# ─── 主入口 ───────────────────────────────────────
def render_factory_po_create_page(orders: pd.DataFrame, table_url: str, headers: dict):
    st.subheader("📝 建立工廠 PO(新流程)")
    st.caption("上傳客戶 PO PDF → 自動解析 → 選工廠/規格 → 產出工廠 PO PDF / 寫回 Teable")

    factories = load_factories()
    if not factories:
        st.error("data/factories.json 讀取失敗或無工廠資料")
        st.stop()

    # ─── Step 1 ─
    st.markdown("### Step 1. 上傳客戶 PO PDF")
    col_up, col_clear = st.columns([4, 1])
    with col_up:
        uploaded = st.file_uploader(
            "支援:WESCO / TIETO / GUDE / KCS / VORNE",
            type=["pdf"],
            key="customer_po_uploader",
        )
    with col_clear:
        st.write("")
        st.write("")
        if st.button("🗑️ 清除", use_container_width=True, key="fpo_clear_top"):
            _reset_form()
            st.rerun()

    if uploaded is None:
        st.info("👆 請上傳客戶 PO PDF。")
        st.stop()

    # ─── Step 2 ─
    try:
        with st.spinner("解析 PDF..."):
            pdf_bytes = uploaded.getvalue()
            parsed = parse_customer_po_from_pdf(pdf_bytes)
    except Exception as e:
        st.error(f"PDF 解析失敗:{e}")
        with st.expander("錯誤詳情"):
            st.code(traceback.format_exc())
        st.stop()

    st.markdown("### Step 2. 解析結果")
    parser_badge = {
        "WESCO": "🟢 WESCO 解析器", "TIETO": "🟢 TIETO 解析器",
        "GUDE": "🟢 GUDE 解析器", "KCS": "🟢 KCS 解析器",
        "VORNE": "🟢 VORNE 解析器", "GENERIC": "🟡 通用模式(欄位請手動補)",
    }
    st.caption(f"使用解析器:{parser_badge.get(parsed.parser_used, parsed.parser_used)}")
    if parsed.parse_warnings:
        for w in parsed.parse_warnings:
            st.warning(w)

    info_cols = st.columns(4)
    info_cols[0].metric("客戶", parsed.customer_name or "(未識別)")
    info_cols[1].metric("客戶 PO#", parsed.customer_po_no or "—")
    info_cols[2].metric("PO 日期", parsed.po_date or "—")
    info_cols[3].metric("總金額", f"{parsed.currency} {parsed.total_amount:,.2f}")

    if not parsed.items:
        st.error("沒有解析出品項,無法繼續。請檢查 PDF 是否正確。")
        with st.expander("📄 PDF 原始文字"):
            st.text(parsed.raw_text)
        st.stop()

    # 可編輯品項表
    st.markdown("#### 品項清單(可編輯)")
    items_df = pd.DataFrame([{
        "P/N": it.part_number, "描述": it.description, "數量": it.quantity,
        "客戶單價": it.unit_price, "客戶金額": it.amount,
        "客戶交期": it.delivery_date or "",
    } for it in parsed.items])
    edited_items_df = st.data_editor(
        items_df, use_container_width=True, hide_index=True,
        num_rows="fixed", key="fpo_items_editor",
    )
    for i, row in edited_items_df.iterrows():
        if i < len(parsed.items):
            parsed.items[i].part_number = str(row["P/N"]).strip()
            parsed.items[i].description = str(row["描述"]).strip()
            try: parsed.items[i].quantity = int(row["數量"])
            except: pass
            try: parsed.items[i].unit_price = float(row["客戶單價"])
            except: pass
            try: parsed.items[i].amount = float(row["客戶金額"])
            except: pass
            parsed.items[i].delivery_date = str(row["客戶交期"]).strip() or None

    # ─── Step 3 ─
    st.markdown("### Step 3. 工廠 / 發出公司 / 字首 / 編號")
    s3_cols = st.columns(3)
    factory_keys = sorted(
        factories.keys(),
        key=lambda k: (factories[k].get("sort_priority", 999), k),
    )
    with s3_cols[0]:
        factory_short = st.selectbox(
            "工廠 *", factory_keys, key="fpo_factory_short",
        )
    factory_data = factories[factory_short]
    factory_region = factory_data.get("region", "Taiwan")

    missing_fields = []
    for k in ["factory_name", "address", "phone", "default_currency", "default_payment_terms"]:
        v = factory_data.get(k, "")
        if not v or "[請補]" in str(v):
            missing_fields.append(k)
    if missing_fields:
        st.warning(f"⚠️ 工廠「{factory_short}」資料不完整,缺:{', '.join(missing_fields)}")

    with s3_cols[1]:
        default_issuing = parsed.issuing_company_detected or "GLOCOM"
        issuing_options = ["GLOCOM", "EUSWAY"]
        issuing_idx = issuing_options.index(default_issuing) if default_issuing in issuing_options else 0
        issuing_company = st.selectbox(
            "發出公司 *", issuing_options, index=issuing_idx, key="fpo_issuing",
        )

    auto_internal_prefix = derive_prefix(issuing_company, factory_region)
    auto_display = display_prefix(auto_internal_prefix)
    prefix_options = ["GC", "ET", "EW"]
    with s3_cols[2]:
        prefix_idx = prefix_options.index(auto_display) if auto_display in prefix_options else 0
        prefix_chosen = st.selectbox(
            "字首", prefix_options, index=prefix_idx, key="fpo_prefix",
        )

    chosen_internal = internal_prefix(prefix_chosen)

    # ─── ★ v3.5 第 1 層防護:即時撈 Teable 編號(不靠 Refresh) ─
    live_po_numbers = get_live_po_numbers(table_url, headers)
    new_po_no = calc_next_po_number(live_po_numbers, chosen_internal)

    # 顯示編號 + Refresh 按鈕
    cols_no = st.columns([2, 0.6, 1, 1])
    with cols_no[0]:
        st.success(f"📋 新訂單編號:**`{new_po_no}`**")
    with cols_no[1]:
        if st.button("🔄", key="fpo_refresh_no", help="重新從 Teable 撈最新編號"):
            get_live_po_numbers(table_url, headers, force_refresh=True)
            st.rerun()
    with cols_no[2]:
        order_date_input = st.date_input("採購日期", value=date.today(), key="fpo_order_date")
    with cols_no[3]:
        is_revised = st.checkbox("REVISED", value=False, key="fpo_revised")

    # 顯示 cache 狀態(讓使用者放心)
    cache_t = st.session_state.get("_fpo_live_po_numbers_t", 0)
    if cache_t:
        age_sec = int(datetime.now().timestamp() - cache_t)
        st.caption(
            f"🟢 Teable 即時資料(共 {len(live_po_numbers)} 筆訂單,{age_sec} 秒前撈) · "
            "30 秒 cache 內不重撈 · 按 🔄 強制刷新"
        )

    cols_pic = st.columns([1, 3])
    with cols_pic[0]:
        purchase_responsible = st.text_input("負責採購", value="Amy", key="fpo_pic")

    # ─── Step 4 ─
    st.markdown("### Step 4. 每品項:工廠單價 + 工廠交期")
    factory_unit_prices = {}
    factory_due_dates = {}
    for idx, it in enumerate(parsed.items):
        p_cols = st.columns([2, 1, 1, 2])
        with p_cols[0]:
            st.text_input("P/N", value=it.part_number, disabled=True, key=f"fpo_disp_pn_{idx}")
        with p_cols[1]:
            st.text_input("數量", value=str(it.quantity), disabled=True, key=f"fpo_disp_qty_{idx}")
        with p_cols[2]:
            f_price = st.number_input(
                "工廠單價", min_value=0.0, value=0.0, step=0.01, format="%.4f",
                key=f"fpo_fprice_{idx}",
            )
        with p_cols[3]:
            default_due = date.today()
            if it.delivery_date:
                try:
                    default_due = datetime.strptime(it.delivery_date, "%Y-%m-%d").date()
                except ValueError:
                    pass
            f_due = st.date_input("工廠交期", value=default_due, key=f"fpo_fdue_{idx}")
        factory_unit_prices[it.part_number] = f_price
        factory_due_dates[it.part_number] = f_due

    factory_total = sum(
        factory_unit_prices.get(it.part_number, 0.0) * it.quantity
        for it in parsed.items
    )
    st.metric("工廠總金額", f"{factory_data.get('default_currency', 'NT$')} {factory_total:,.2f}")

    # ─── Step 5: 每品項規格(3 Tab) ─
    st.markdown("### Step 5. 每品項產品規格")
    st.caption("每個品項各有 3 種輸入方式可選。規格會以「一列」印在工廠 PO 的『產品規格』欄。")

    item_specs = {}
    for idx, it in enumerate(parsed.items):
        spec_text = render_spec_input_for_item(idx, it, orders)
        item_specs[it.part_number] = spec_text

    # ─── Step 6: 兩個獨立按鈕 ─
    st.divider()
    st.markdown("### Step 6. 產出 / 寫回")
    st.caption("**產 PDF** 和 **寫回 Teable** 是分開的。建議先產 PDF 預覽 → OK 再寫回 Teable。")

    btn_cols = st.columns([2, 2, 1])
    with btn_cols[0]:
        do_pdf = st.button(
            "📄 產 PDF", type="primary",
            use_container_width=True, key="fpo_do_pdf",
        )
    with btn_cols[1]:
        do_writeback = st.button(
            "💾 寫回 Teable", type="secondary",
            use_container_width=True, key="fpo_do_writeback",
        )
    with btn_cols[2]:
        if st.button("🗑️ 全部清除", use_container_width=True, key="fpo_clear_bottom"):
            _reset_form()
            st.rerun()

    if not (do_pdf or do_writeback):
        st.stop()

    # ─── 共用驗證 ─
    zero_price_pns = [pn for pn, price in factory_unit_prices.items() if price == 0]
    if zero_price_pns:
        st.warning(f"⚠️ 工廠單價為 0 的品項:{', '.join(zero_price_pns)}")

    empty_spec_pns = [pn for pn, sp in item_specs.items() if not (sp or "").strip()]
    if empty_spec_pns:
        st.error(f"❌ 以下品項規格沒填:**{', '.join(empty_spec_pns)}**。"
                 "規格欄會印空白,請回 Step 5 填好再產 PDF / 寫回 Teable。")
        st.stop()

    # ─── ★ v3.5 第 2 層防護:產 PDF 前撞號檢查 ─
    if do_pdf:
        with st.spinner("最後檢查 Teable 編號..."):
            fresh_po_set = get_live_po_numbers(table_url, headers, force_refresh=True)
        
        if new_po_no in fresh_po_set:
            # 撞了!算新的並擋下
            corrected_po = calc_next_po_number(fresh_po_set, chosen_internal)
            st.error(
                f"❌ **編號撞號警告!** "
                f"剛剛偵測到 Teable 上已經有 **`{new_po_no}`**(可能是別人在你建單期間手動 key 進去)。\n\n"
                f"➡️ 系統建議改用 **`{corrected_po}`**。請按 🔄 重新整理頁面後再試一次。"
            )
            st.stop()

    if do_writeback:
        with st.spinner("最後檢查 Teable 編號..."):
            fresh_po_set = get_live_po_numbers(table_url, headers, force_refresh=True)
        
        if new_po_no in fresh_po_set:
            # ─── ★ v3.5 第 3 層:寫入鎖,撞了就拒絕 ─
            corrected_po = calc_next_po_number(fresh_po_set, chosen_internal)
            st.error(
                f"❌ **寫入鎖觸發!** "
                f"Teable 上已經有 **`{new_po_no}`**,拒絕寫入避免覆蓋。\n\n"
                f"➡️ 請按 🔄 重新整理編號(會跳到 **`{corrected_po}`**)後重試。"
            )
            st.stop()

    # ─── 產 PDF ─
    if do_pdf:
        po_ctx = build_po_context_from_new_order(
            new_po_no=new_po_no, parsed=parsed, factory_data=factory_data,
            factory_short=factory_short, issuing_company=issuing_company,
            factory_unit_prices=factory_unit_prices, factory_due_dates=factory_due_dates,
            item_specs=item_specs, purchase_responsible=purchase_responsible,
            order_date=order_date_input, is_revised=is_revised,
        )

        try:
            from core.pdf_generator import generate_po_files
            with st.spinner("產生 docx + PDF..."):
                result = generate_po_files(po_ctx)
            docx_path = result.get("docx_path")
            pdf_path = result.get("pdf_path")
            if result.get("error"):
                st.warning(result["error"])
        except FileNotFoundError as e:
            if "PO_EUSWAY" in str(e):
                st.error(
                    "❌ EUSWAY 模板未上傳:`templates/PO_EUSWAY.docx` 不存在。"
                )
            else:
                st.error(f"❌ 產 PDF 失敗:{e}")
            st.stop()
        except Exception as e:
            st.error(f"❌ 產 PDF 失敗:{e}")
            with st.expander("錯誤詳情"):
                st.code(traceback.format_exc())
            st.stop()

        st.success(f"✅ PDF / DOCX 已產生(訂單編號 {new_po_no})")

        d_cols = st.columns(2)
        if docx_path and Path(docx_path).exists():
            with d_cols[0]:
                with open(docx_path, "rb") as f:
                    st.download_button(
                        "📥 下載 DOCX", data=f.read(),
                        file_name=Path(docx_path).name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True, key="fpo_dl_docx",
                    )
        if pdf_path and Path(pdf_path).exists():
            with d_cols[1]:
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        "📥 下載 PDF", data=f.read(),
                        file_name=Path(pdf_path).name, mime="application/pdf",
                        use_container_width=True, key="fpo_dl_pdf",
                    )

        st.info("📝 PDF 確認 OK 後,可按上方「💾 寫回 Teable」把資料寫進主表。")

    # ─── 寫回 Teable ─
    if do_writeback:
        records_fields = []
        for it in parsed.items:
            f_price = factory_unit_prices.get(it.part_number, 0.0)
            f_due = factory_due_dates.get(it.part_number)
            f_due_str = f_due.strftime("%Y/%m/%d") if f_due else ""
            order_date_str = order_date_input.strftime("%Y/%m/%d")
            cust_date_str = parsed.po_date or order_date_str
            spec_text = (item_specs.get(it.part_number, "") or "").strip()

            records_fields.append({
                COL_GLOCOM_PO: new_po_no,
                "客戶": parsed.customer_name,
                "PO#": parsed.customer_po_no,
                "P/N": it.part_number,
                "Order Q'TY\n (PCS)": it.quantity,
                "工廠": factory_short,
                "工廠交期": f_due_str,
                "Ship date": f_due_str,
                "銷貨金額": round(f_price * it.quantity, 2),
                "接單金額": round(it.amount, 2),
                "客戶要求注意事項": spec_text,
                "工廠下單日期": order_date_str,
                "客戶下單日期": cust_date_str,
            })

        with st.spinner(f"寫入 {len(records_fields)} 筆 Teable record..."):
            wb_result = create_teable_records(table_url, headers, records_fields)

        if wb_result["success"] > 0:
            st.success(f"✅ 已寫入 Teable {wb_result['success']} 筆 record(訂單編號 {new_po_no})")
            # 寫成功後清掉 cache,下次進來會重撈
            for k in ["_fpo_live_po_numbers", "_fpo_live_po_numbers_t"]:
                if k in st.session_state:
                    del st.session_state[k]
            st.info("⚠️ 請按側邊欄 **Refresh** 才會看到主表更新。")
        if wb_result["failed"] > 0:
            st.error(f"❌ 寫入失敗 {wb_result['failed']} 筆")
            with st.expander("失敗原因"):
                for err in wb_result["errors"]:
                    st.text(err)
