# -*- coding: utf-8 -*-
"""
建立工廠 PO(新流程)頁面。

流程:
  1. 上傳客戶 PO PDF
  2. 自動解析 → 顯示客戶/PO#/品項/金額(可手動修正)
  3. 自動偵測「發出公司」(GLOCOM / EUSWAY) — 可手動覆蓋
  4. 後續(下一階段再做):選工廠 → 算字首 → 自動編號 → 產 PDF → 寫回 Teable

呼叫方式(在 app.py menu 加上 "建立工廠 PO" 後):
    elif menu == "建立工廠 PO":
        from factory_po_create_page import render_factory_po_create_page
        render_factory_po_create_page(orders, TABLE_URL, HEADERS)
"""

from __future__ import annotations
import streamlit as st
import pandas as pd

from core.customer_po_parser import (
    parse_customer_po,
    extract_text_from_pdf,
    POItem,
    ParsedPO,
)
from core.teable_query import (
    parse_glocom_po_no,
    safe_text,
    COL_GLOCOM_PO,
)


# ─── 字首邏輯 ────────────────────────────────────
# 發出公司 + 工廠地區 → 字首
# GLOCOM + Taiwan → GC
# EUSWAY + Taiwan → ET
# EUSWAY + China  → EW
# (HK 暫時當 China,因為 EW 樣本含 Profit Grand HK)
def derive_prefix(issuing_company: str, factory_region: str) -> str:
    issuing = (issuing_company or "").upper()
    region = (factory_region or "").lower()

    if issuing == "GLOCOM":
        return "G"  # 內部用 G,UI 顯示為 GC
    if issuing == "EUSWAY":
        if region in ("china", "hk", "hong kong"):
            return "EW"
        return "ET"
    return "G"


def display_prefix(internal_prefix: str) -> str:
    """G → GC,其他原樣。"""
    return "GC" if internal_prefix == "G" else internal_prefix


def calc_next_po_number(orders_df: pd.DataFrame, prefix: str) -> str:
    """從 Teable 主表抓同字首訂單,流水號 +1。

    西拓訂單編號格式:[字首][民國年3碼][流水號4碼]-[項次2碼]
    例如 G1150031-02 = G + 115 + 0031 + -02

    新單的流水號 = 同字首所有訂單中,流水號最大值 + 1
    新單預設項次 = -01
    """
    import re

    if orders_df.empty or COL_GLOCOM_PO not in orders_df.columns:
        return ""

    # 取所有同字首的訂單編號
    pattern = re.compile(rf"^{re.escape(prefix)}(\d{{3}})(\d{{4}})-(\d+)$")
    max_year_serial = (0, 0)  # (民國年, 流水號)

    for raw_po in orders_df[COL_GLOCOM_PO].dropna().astype(str):
        m = pattern.match(raw_po.strip())
        if m:
            year = int(m.group(1))
            serial = int(m.group(2))
            if (year, serial) > max_year_serial:
                max_year_serial = (year, serial)

    if max_year_serial == (0, 0):
        # 沒有同字首訂單,從目前民國年開始
        from datetime import datetime
        roc_year = datetime.now().year - 1911
        return f"{prefix}{roc_year}0001-01"

    year, serial = max_year_serial
    new_serial = serial + 1
    return f"{prefix}{year}{new_serial:04d}-01"


# ─── UI 主入口 ────────────────────────────────────
def render_factory_po_create_page(orders: pd.DataFrame, table_url: str, headers: dict):
    st.subheader("📝 建立工廠 PO(新流程)")
    st.caption("上傳客戶 PO PDF → 自動解析 → 後續產出工廠 PO")

    # 提示這頁是 in-progress
    st.info(
        "🚧 **目前可用功能**:PDF 解析、自動偵測客戶與發出公司、自動產生新訂單編號預覽。  \n"
        "**還沒做**:選工廠、規格自動帶入、產 PDF、寫回 Teable(下一階段)。"
    )

    # ─── Step 1: 上傳 PDF ──────────────────────────
    st.markdown("### Step 1. 上傳客戶 PO PDF")
    uploaded = st.file_uploader(
        "支援 PDF",
        type=["pdf"],
        key="customer_po_uploader",
        help="目前自動解析支援:WESCO / TIETO / GUDE / KCS。其他客戶會走通用模式,需要手動填欄位。",
    )

    if uploaded is None:
        st.stop()

    # ─── Step 2: 解析 ──────────────────────────────
    try:
        with st.spinner("解析 PDF..."):
            pdf_bytes = uploaded.getvalue()
            raw_text = extract_text_from_pdf(pdf_bytes)
            parsed = parse_customer_po(raw_text)
    except Exception as e:
        st.error(f"PDF 解析失敗:{e}")
        with st.expander("錯誤詳情"):
            import traceback
            st.code(traceback.format_exc())
        st.stop()

    # 顯示解析結果
    st.markdown("### Step 2. 解析結果")

    if parsed.parse_warnings:
        for w in parsed.parse_warnings:
            st.warning(w)

    # 解析器標籤
    parser_badge = {
        "WESCO": "🟢 WESCO 解析器",
        "TIETO": "🟢 TIETO 解析器",
        "GUDE":  "🟢 GUDE 解析器",
        "KCS":   "🟢 KCS 解析器",
        "GENERIC": "🟡 通用模式(欄位請手動補)",
    }
    st.caption(f"使用解析器:{parser_badge.get(parsed.parser_used, parsed.parser_used)}")

    # 主檔欄位
    info_cols = st.columns(4)
    info_cols[0].metric("客戶", parsed.customer_name or "(未識別)")
    info_cols[1].metric("客戶 PO#", parsed.customer_po_no or "—")
    info_cols[2].metric("PO 日期", parsed.po_date or "—")
    info_cols[3].metric("總金額", f"{parsed.currency} {parsed.total_amount:,.2f}")

    # 品項表
    st.markdown("#### 品項清單")
    if parsed.items:
        items_df = pd.DataFrame([{
            "Line": it.line,
            "P/N": it.part_number,
            "描述": it.description,
            "數量": it.quantity,
            "單價": it.unit_price,
            "金額": it.amount,
            "交期": it.delivery_date or "",
        } for it in parsed.items])
        st.dataframe(items_df, use_container_width=True, hide_index=True)
    else:
        st.info("未解析出品項,請參考下方原始文字手動補。")

    # 其他欄位
    other_cols = st.columns(2)
    with other_cols[0]:
        st.text_input("付款條件", value=parsed.payment_terms, key="po_payment_terms")
    with other_cols[1]:
        st.text_input("出貨方式 (Ship Via)", value=parsed.ship_via, key="po_ship_via")

    # ─── Step 3: 發出公司 + 字首 ──────────────────
    st.markdown("### Step 3. 發出公司與字首")

    issuing_cols = st.columns(3)

    with issuing_cols[0]:
        # 預設用解析的,使用者可改
        default_issuing = parsed.issuing_company_detected or "GLOCOM"
        issuing_options = ["GLOCOM", "EUSWAY"]
        issuing_idx = issuing_options.index(default_issuing) if default_issuing in issuing_options else 0
        issuing_company = st.selectbox(
            "發出公司",
            issuing_options,
            index=issuing_idx,
            key="po_issuing_company",
            help=f"PDF 上偵測到:{parsed.issuing_company_detected}",
        )

    with issuing_cols[1]:
        # 工廠地區先用簡單下拉(下一階段會接 factories.json)
        factory_region = st.selectbox(
            "工廠地區",
            ["Taiwan", "China"],
            key="po_factory_region",
            help="GLOCOM 一律 Taiwan;EUSWAY 大陸工廠選 China,台灣工廠選 Taiwan。",
        )

    # 算字首
    auto_prefix = derive_prefix(issuing_company, factory_region)
    auto_prefix_display = display_prefix(auto_prefix)

    with issuing_cols[2]:
        # 字首使用者可覆蓋
        prefix_options = ["GC", "ET", "EW"]
        prefix_idx = prefix_options.index(auto_prefix_display) if auto_prefix_display in prefix_options else 0
        prefix_chosen = st.selectbox(
            "字首",
            prefix_options,
            index=prefix_idx,
            key="po_prefix",
            help=f"自動推算:{auto_prefix_display}(可覆蓋)",
        )

    # 算新編號
    internal_prefix = "G" if prefix_chosen == "GC" else prefix_chosen
    new_po_no = calc_next_po_number(orders, internal_prefix)

    if new_po_no:
        st.success(f"自動產生新訂單編號:**`{new_po_no}`**")
    else:
        st.warning("無法產生新編號(主表沒同字首訂單,且無法判斷民國年)")

    # ─── Step 4: 之後再做 ──────────────────────────
    st.markdown("### Step 4. 下一步(尚未實作)")
    st.markdown("""
    後續流程:
    1. 選工廠(從 `factories.json` 載入,讀出工廠 region 自動更新字首)
    2. 規格從前張同料號訂單帶入(可改)
    3. 工廠單價、工廠出貨日(手動填)
    4. 產出工廠 PO PDF
    5. 寫回 Teable(建立新 record)
    """)

    # ─── 原始文字對照 ─────────────────────────────
    with st.expander("📄 PDF 原始文字(對照用)"):
        st.text(raw_text)
