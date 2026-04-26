# -*- coding: utf-8 -*-
"""
Factory PO 頁面模組,給 GLOCOM Control Tower 使用。

呼叫方式(在 app.py 中):
    elif menu == "Factory PO":
        from factory_po_page import render_factory_po_page
        render_factory_po_page(orders, TABLE_URL, HEADERS)

依賴:
- core.teable_query (從主表 DataFrame 撈 PO 資料)
- core.factory_master (factories.json)
- core.pdf_generator (產 docx + PDF)
- core.teable_writeback (回寫 PDF URL)

階段 1 變更(2026-04-26):
- 新增三個下拉聯動篩選:類型 / 客戶 / 工廠
- 類型篩選 UI 上 G 改顯示為 GC(內部資料仍維持 G)
"""

import streamlit as st
import pandas as pd
from datetime import date
from pathlib import Path

from core.teable_query import (
    list_glocom_po_options,
    get_po_rows,
    build_po_context,
    parse_glocom_po_no,
    safe_text,
    COL_GLOCOM_PO,
    COL_PART_NO,
    COL_QTY,
    COL_FACTORY,
    COL_FACTORY_DUE,
    COL_AMOUNT_FACTORY,
)
from core.factory_master import (
    get_factory,
    has_complete_data,
)
from core.pdf_generator import generate_po_files
from core.teable_writeback import write_pdf_url_to_records


# ─── 客戶欄位探測 ─────────────────────────────────
# core.teable_query 沒有 export COL_CUSTOMER,所以這裡用候選名單探測。
# 若你的 Teable 表用其他名稱,請加到這個 list。
CUSTOMER_COL_CANDIDATES = ["Customer", "客戶", "客戶名稱"]


def _find_customer_col(df: pd.DataFrame) -> str | None:
    for c in CUSTOMER_COL_CANDIDATES:
        if c in df.columns:
            return c
    return None


def _ui_label_to_order_type(label: str) -> str:
    """UI selectbox 顯示的標籤,轉回 parse_glocom_po_no 內部的 order_type。

    UI 上 G 顯示為 GC,但內部資料(訂單編號 G1150031-02)仍是 G。
    """
    if label == "GC":
        return "G"
    return label


def _display_order_type(order_type: str) -> str:
    """內部 order_type 轉成 UI 顯示。"""
    if order_type == "G":
        return "GC"
    return order_type


def render_factory_po_page(orders: pd.DataFrame, table_url: str, headers: dict):
    """主頁面渲染。

    Args:
        orders: Control Tower 已載入的主表 DataFrame
        table_url: Teable record API URL(從 app.py 帶入)
        headers: 含 Authorization 的 dict
    """
    st.subheader("🏭 Factory PO Generator")
    st.caption("依「西拓訂單編號」從主表撈出資料 → 產出工廠 PO PDF → 回寫 Teable")

    if orders.empty:
        st.warning("主表沒有資料,無法選單號。")
        return

    if COL_GLOCOM_PO not in orders.columns:
        st.error(f"主表缺少必要欄位「{COL_GLOCOM_PO}」")
        return

    # 偵測客戶欄位
    customer_col = _find_customer_col(orders)

    # ─── 三個下拉聯動篩選 ──────────────────────────
    st.markdown("### 1. 篩選與選擇西拓訂單編號")

    # filtered 會被三個篩選條件依序縮小範圍,最後再丟給 list_glocom_po_options
    filtered = orders.copy()

    col_type, col_customer, col_factory = st.columns(3)

    # 1) 類型篩選 (UI 上 G 顯示為 GC)
    with col_type:
        type_filter = st.selectbox(
            "類型篩選",
            ["全部", "GC", "ET", "EW"],
            key="po_type_filter",
            help="GC = GLOCOM 給台灣工廠 / ET = EUSWAY 給台灣工廠 / EW = EUSWAY 給大陸工廠",
        )

    if type_filter != "全部":
        actual_order_type = _ui_label_to_order_type(type_filter)
        type_mask = filtered[COL_GLOCOM_PO].apply(
            lambda x: parse_glocom_po_no(safe_text(x)).get("order_type", "") == actual_order_type
        )
        filtered = filtered[type_mask]

    # 2) 客戶篩選
    with col_customer:
        if customer_col:
            customer_values = sorted({
                str(x).strip()
                for x in filtered[customer_col].dropna().tolist()
                if str(x).strip()
            })
            customer_options = ["全部"] + customer_values
            customer_filter = st.selectbox(
                "客戶",
                customer_options,
                key="po_customer_filter",
            )
        else:
            st.selectbox(
                "客戶 (找不到欄位)",
                ["全部"],
                disabled=True,
                key="po_customer_filter_disabled",
            )
            customer_filter = "全部"

    if customer_filter != "全部" and customer_col:
        filtered = filtered[
            filtered[customer_col].astype(str).str.strip() == customer_filter
        ]

    # 3) 工廠篩選
    with col_factory:
        if COL_FACTORY in filtered.columns:
            factory_values = sorted({
                str(x).strip()
                for x in filtered[COL_FACTORY].dropna().tolist()
                if str(x).strip()
            })
            factory_options = ["全部"] + factory_values
            factory_filter = st.selectbox(
                "工廠",
                factory_options,
                key="po_factory_filter",
            )
        else:
            st.selectbox(
                "工廠 (找不到欄位)",
                ["全部"],
                disabled=True,
                key="po_factory_filter_disabled",
            )
            factory_filter = "全部"

    if factory_filter != "全部" and COL_FACTORY in filtered.columns:
        filtered = filtered[
            filtered[COL_FACTORY].astype(str).str.strip() == factory_filter
        ]

    # ─── 從過濾後 orders 撈 PO options ─────────────
    po_options = list_glocom_po_options(filtered)

    if not po_options:
        st.info("依目前篩選條件沒有任何訂單。請調整上方篩選器。")
        return

    selected_idx = st.selectbox(
        f"西拓訂單編號 (共 {len(po_options)} 個可選)",
        range(len(po_options)),
        format_func=lambda i: po_options[i][1],
        key="factory_po_select",
    )

    po_no = po_options[selected_idx][0]

    # ─── 撈出該編號的所有 row ─────────────────────────
    # 注意:get_po_rows 用原始 orders(不是 filtered),確保抓得到完整品項
    rows = get_po_rows(orders, po_no)
    if rows.empty:
        st.error(f"找不到 {po_no} 的資料")
        return

    parsed = parse_glocom_po_no(po_no)

    st.divider()
    st.markdown(f"### 2. 訂單資訊 — `{po_no}`")

    info_cols = st.columns(4)
    info_cols[0].metric("品項數", len(rows))
    info_cols[1].metric("訂單類型", _display_order_type(parsed["order_type"]))
    info_cols[2].metric("開單抬頭", parsed["issuing_company"])
    factory_short = safe_text(rows.iloc[0][COL_FACTORY]) if COL_FACTORY in rows.columns else ""
    info_cols[3].metric("工廠", factory_short or "(無)")

    # ─── 工廠主檔檢查 ────────────────────────────────
    factory = get_factory(factory_short) if factory_short else None
    if not factory:
        st.warning(
            f"⚠️ 工廠「{factory_short}」不在 factories.json,"
            f"PDF 將印出空殼資料。請先補完 data/factories.json。"
        )
    else:
        is_complete, missing = has_complete_data(factory)
        if not is_complete:
            st.warning(
                f"⚠️ 工廠「{factory_short}」資料不完整(缺:{', '.join(missing)})。"
                f"PDF 仍可產生,但這些欄位會印 [請補]。"
            )

    # ─── 顯示品項表 ────────────────────────────────
    st.markdown("#### 品項清單(主表撈到的原始資料)")
    show_cols = [c for c in [COL_PART_NO, COL_QTY, COL_FACTORY_DUE, COL_AMOUNT_FACTORY] if c in rows.columns]
    if show_cols:
        st.dataframe(rows[show_cols], use_container_width=True, hide_index=True)

    # ─── 採購日期 / REVISED 選項 ─────────────────────
    st.divider()
    st.markdown("### 3. 產出選項")
    opt_cols = st.columns(3)
    with opt_cols[0]:
        order_date = st.date_input("採購日期", value=date.today(), key="po_date")
    with opt_cols[1]:
        purchase_responsible = st.text_input("負責採購", value="Amy", key="po_resp")
    with opt_cols[2]:
        is_revised = st.checkbox("REVISED(修訂版)", value=False, key="po_revised")

    # ─── 產生 + 回寫 ─────────────────────────────────
    st.divider()
    st.markdown("### 4. 產生 PDF")

    action_cols = st.columns([1, 1, 2])
    with action_cols[0]:
        gen_clicked = st.button("📄 產生 PDF", type="primary", use_container_width=True)
    with action_cols[1]:
        gen_writeback_clicked = st.button(
            "📄 產生 + 回寫 Teable",
            use_container_width=True,
            disabled=True,
            help="Phase 2 功能(Teable Attachment 上傳邏輯尚未實作)。目前請手動把 PDF 拖到 Teable 主表的 Factory PO PDF 欄位。",
        )

    if gen_clicked or gen_writeback_clicked:
        try:
            # 用空殼工廠資料兜底
            fac_for_ctx = factory or {
                "factory_name": factory_short or "(請補)",
                "address": "[請補]",
                "contact_person": "[請補]",
                "phone": "[請補]",
                "fax": "[請補]",
                "default_currency": "NT$",
                "default_payment_terms": "[請補]",
                "default_shipment": "待通知",
                "default_ship_to": "待通知",
            }

            po_ctx = build_po_context(po_no, rows, fac_for_ctx)
            # 帶入 UI 上的選項
            po_ctx["order_date"] = order_date
            po_ctx["purchase_responsible"] = purchase_responsible
            po_ctx["is_revised"] = is_revised

            with st.spinner("產生 docx 與 PDF..."):
                result = generate_po_files(po_ctx)

            if result.get("error"):
                st.warning(result["error"])

            docx_path = result["docx_path"]
            pdf_path = result.get("pdf_path")

            st.success(f"✓ {po_no} 已產生")

            # 提供下載
            with open(docx_path, "rb") as f:
                st.download_button(
                    "下載 DOCX",
                    data=f.read(),
                    file_name=docx_path.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_docx",
                )
            if pdf_path and Path(pdf_path).exists():
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        "下載 PDF",
                        data=f.read(),
                        file_name=pdf_path.name,
                        mime="application/pdf",
                        key="dl_pdf",
                    )

            # 回寫 Teable
            if gen_writeback_clicked:
                record_ids = po_ctx.get("_record_ids", [])
                if not record_ids:
                    st.warning("找不到對應的 record_id,無法回寫 Teable")
                else:
                    pdf_url_to_write = (
                        str(pdf_path) if pdf_path else str(docx_path)
                    )
                    with st.spinner(f"回寫 {len(record_ids)} 筆 Teable record..."):
                        wb_result = write_pdf_url_to_records(
                            record_ids=record_ids,
                            pdf_url=pdf_url_to_write,
                            table_url=table_url,
                            headers=headers,
                        )
                    st.write(
                        f"回寫結果:成功 {wb_result['success']} 筆 / "
                        f"失敗 {wb_result['failed']} 筆"
                    )
                    if wb_result["errors"]:
                        with st.expander("回寫錯誤明細"):
                            for err in wb_result["errors"][:20]:
                                st.text(err)

                    if wb_result["success"] > 0:
                        st.info("⚠️ 回寫成功,請按側邊欄 Refresh 看到主表更新後的資料")

        except Exception as e:
            st.error(f"產生失敗:{e}")
            import traceback
            with st.expander("錯誤詳情"):
                st.code(traceback.format_exc())
