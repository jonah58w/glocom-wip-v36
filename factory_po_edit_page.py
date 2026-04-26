# -*- coding: utf-8 -*-
"""
編輯既有訂單頁面。

功能:
- 輸入西拓訂單編號 → 載入 Teable 主表上同訂單編號的所有 records
- 顯示可編輯欄位(客戶/PO#/P/N/數量/工廠/交期/銷貨金額/接單金額/規格/日期)
- 不存進 Teable 的欄位:負責採購、單價(PDF 顯示用,從金額/數量反推)
- 兩個獨立按鈕:【📄 產 PDF】+【💾 更新 Teable】
- 支援 REVISED 印章

設計原則:
- 不改 Teable schema(主表沒「工廠單價」欄,PDF 單價即時計算)
- PATCH API 直接覆蓋原 record(用 _record_id 識別)
- 訂單編號是 key,不可編輯
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

from core.teable_query import COL_GLOCOM_PO, parse_glocom_po_no, safe_text


HERE = Path(__file__).parent
FACTORIES_JSON = HERE / "data" / "factories.json"


# ─── 工廠主檔(剝括號註記)─────────────────
def _strip_placeholder_brackets(text: str) -> str:
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
        for v in active.values():
            v["factory_name"] = _strip_placeholder_brackets(v.get("factory_name", ""))
        return active
    except Exception as e:
        st.error(f"factories.json 讀取失敗:{e}")
        return {}


# ─── 共用工具 ────────────────────────────
def _safe_str(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def _safe_int(v, default=0) -> int:
    try:
        if v is None or v == "":
            return default
        return int(float(str(v).strip()))
    except Exception:
        return default


def _safe_float(v, default=0.0) -> float:
    try:
        if v is None or v == "":
            return default
        return float(str(v).strip())
    except Exception:
        return default


def _parse_date_flexible(v):
    """嘗試多種格式解析日期,失敗回傳 None"""
    s = _safe_str(v)
    if not s:
        return None
    # 移除附加文字(例如「VORNE 4/23 10:51 PM 已送達...」)
    s = s.split(" ")[0] if " " in s else s
    fmts = ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%b. %d, %y", "%b %d, %Y"]
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _clean_customer_name(raw: str) -> str:
    """『VORNE 4/23 10:51 PM 已送達...』→『VORNE』"""
    s = _safe_str(raw)
    if not s:
        return ""
    # 取第一個空白前的字
    m = re.match(r"^([A-Z0-9]+)", s)
    if m:
        return m.group(1)
    return s.split()[0] if s else ""


# ─── Teable API ─────────────────────────
def patch_teable_record(table_url: str, headers: dict, record_id: str, fields: dict) -> dict:
    """PATCH 單一 record(用 record_id 識別)"""
    if not record_id:
        return {"success": False, "error": "record_id 為空"}

    url = f"{table_url}/{record_id}"
    payload = {
        "fieldKeyType": "name",
        "typecast": True,
        "record": {"fields": fields},
    }
    try:
        r = requests.patch(url, headers=headers, json=payload, timeout=30)
        if r.status_code in (200, 201):
            return {"success": True, "data": r.json()}
        return {"success": False, "error": f"HTTP {r.status_code}: {r.text[:300]}"}
    except Exception as e:
        return {"success": False, "error": f"Exception: {e}"}


# ─── 載入訂單 ───────────────────────────
def load_order_records(orders: pd.DataFrame, po_no: str) -> pd.DataFrame:
    """從 orders DataFrame 篩出指定西拓訂單編號的所有品項"""
    if orders.empty or COL_GLOCOM_PO not in orders.columns:
        return pd.DataFrame()
    target = po_no.strip()
    matches = orders[orders[COL_GLOCOM_PO].astype(str).str.strip() == target].copy()
    return matches.reset_index(drop=True)


# ─── 規格輸入 (從 create page 借用 3 Tab 設計但簡化) ─────
SPEC_OPTIONS_FULL = {
    "Material": ["2L", "4L", "6L", "8L", "10L", "Aluminum", "FR4", "Rogers"],
    "Tg": ["Tg130", "Tg150", "Tg170", "Tg180"],
    "Board thickness": ["0.4mm", "0.6mm", "0.8mm", "1.0mm", "1.2mm", "1.6mm", "2.0mm", "2.4mm", "3.2mm"],
    "Copper": ["1oz/1oz", "1/2oz/1/2oz", "2oz/2oz", "3oz/3oz", "Ext: 1oz all layers"],
    "Surface Finish": ["ENIG", "Lead-free HASL", "HASL", "Immersion Gold", "OSP", "Hard Gold", "Immersion Silver"],
    "S/M": ["Green", "Matte Green", "Red", "Blue", "Black", "Matte Black", "White"],
    "S/L": ["White", "Black", "Yellow"],
}


def _build_full_spec_oneline(selections: dict, extra: str) -> str:
    parts = []
    for cat, val in selections.items():
        if val:
            parts.append(val if cat == "Tg" else f"{cat}: {val}")
    if extra and extra.strip():
        parts.append(extra.strip())
    return "; ".join(parts)


def render_spec_editor(idx: int, current_spec: str, part_no: str) -> str:
    """規格編輯區(3 Tab + 最終文字框)"""
    final_key = f"fpoe_spec_final_{idx}"
    if final_key not in st.session_state:
        st.session_state[final_key] = current_spec or ""

    tab_a, tab_b, tab_c = st.tabs([
        "📋 多列貼上(從 ERP)",
        "✏️ 完整規格(7選單)",
        "📝 簡短規格",
    ])

    with tab_a:
        st.caption("從 ERP 一列一列複製貼上,組合成一列")
        paste_lines = []
        cols_paste = st.columns(2)
        for line_idx in range(8):
            target = cols_paste[line_idx % 2]
            with target:
                line = st.text_input(
                    f"第 {line_idx + 1} 列",
                    key=f"fpoe_paste_{idx}_{line_idx}",
                    placeholder="例如: 8L, FR4, 1.6mm, 1up",
                )
            if line and line.strip():
                paste_lines.append(line.strip())

        if paste_lines:
            preview = "; ".join(paste_lines)
            st.caption(f"預覽: `{preview}`")
            if st.button("⬇ 組合並套用", key=f"fpoe_combine_{idx}", type="primary"):
                st.session_state[final_key] = preview
                st.rerun()

    with tab_b:
        cols_a = st.columns(4)
        cols_b = st.columns(4)
        selections = {}
        cats = list(SPEC_OPTIONS_FULL.keys())
        for i, cat in enumerate(cats):
            target_col = cols_a[i] if i < 4 else cols_b[i - 4]
            with target_col:
                selections[cat] = st.selectbox(
                    cat, [""] + SPEC_OPTIONS_FULL[cat],
                    key=f"fpoe_spec_full_{idx}_{cat}",
                )
        extra = st.text_input(
            "其他規格 / 補充",
            key=f"fpoe_spec_full_extra_{idx}",
        )
        preview = _build_full_spec_oneline(selections, extra)
        if preview:
            st.caption(f"預覽: `{preview}`")
            if st.button("⬇ 套用", key=f"fpoe_apply_full_{idx}", type="primary"):
                st.session_state[final_key] = preview
                st.rerun()

    with tab_c:
        short = st.text_input(
            "簡短規格(直接輸入一列)",
            key=f"fpoe_spec_short_{idx}",
            placeholder='例如: 舊料號;S/M: Red',
        )
        if short.strip():
            if st.button("⬇ 套用", key=f"fpoe_apply_short_{idx}", type="primary"):
                st.session_state[final_key] = short.strip()
                st.rerun()

    # 最終可編輯文字框
    st.markdown("##### 📋 最終規格(印在 PO 上,可直接編輯)")
    st.text_area(
        "spec_final",
        key=final_key,
        label_visibility="collapsed",
        height=68,
    )

    return st.session_state[final_key]


# ─── 主入口 ──────────────────────────────
def render_factory_po_edit_page(orders: pd.DataFrame, table_url: str, headers: dict):
    st.subheader("✏️ 編輯既有訂單")
    st.caption(
        "用西拓訂單編號載入 → 編輯 → 重產 PDF 或更新 Teable。"
        "**產 PDF** 跟 **更新 Teable** 是分開動作。"
    )

    factories = load_factories()
    if not factories:
        st.error("data/factories.json 讀取失敗")
        st.stop()

    # ─── 載入訂單 ─
    cols_load = st.columns([3, 1, 1])
    with cols_load[0]:
        po_no_input = st.text_input(
            "西拓訂單編號",
            placeholder="例如: G1150027-01",
            key="fpoe_po_no_input",
        )
    with cols_load[1]:
        st.write("")
        st.write("")
        do_load = st.button("🔍 載入", use_container_width=True, key="fpoe_load_btn")
    with cols_load[2]:
        st.write("")
        st.write("")
        if st.button("🗑️ 清除", use_container_width=True, key="fpoe_clear_btn"):
            keys_to_clear = [k for k in list(st.session_state.keys()) if k.startswith("fpoe_")]
            for k in keys_to_clear:
                del st.session_state[k]
            st.rerun()

    if do_load:
        st.session_state.fpoe_loaded_po = po_no_input.strip()

    loaded_po = st.session_state.get("fpoe_loaded_po", "")
    if not loaded_po:
        st.info("👆 輸入西拓訂單編號 → 按【🔍 載入】開始編輯。")
        st.stop()

    # ─── 查詢 records ─
    matches = load_order_records(orders, loaded_po)
    if matches.empty:
        st.error(f"❌ 找不到訂單編號『{loaded_po}』。請按 Refresh 重新載入主表後再試。")
        st.stop()

    n_items = len(matches)
    st.success(f"✅ 載入訂單『{loaded_po}』,共 **{n_items}** 筆品項")

    # ─── 共用欄位(整張單共用) ─
    st.markdown("### 📋 訂單共用設定")
    cols_common = st.columns(3)
    first_row = matches.iloc[0]

    # 客戶名稱(從第一筆取,清掉附加文字)
    raw_customer = _safe_str(first_row.get("客戶", ""))
    cleaned_customer = _clean_customer_name(raw_customer)
    with cols_common[0]:
        customer_name = st.text_input(
            "客戶",
            value=cleaned_customer,
            key="fpoe_customer",
        )

    # 客戶 PO#
    raw_cpo = _safe_str(first_row.get("PO#", ""))
    if raw_cpo.endswith(".0"):
        raw_cpo = raw_cpo[:-2]
    with cols_common[1]:
        customer_po_no = st.text_input(
            "客戶 PO#",
            value=raw_cpo,
            key="fpoe_customer_po",
        )

    # 工廠(整單共用)
    factory_keys = sorted(factories.keys())
    raw_factory = _safe_str(first_row.get("工廠", ""))
    factory_idx = factory_keys.index(raw_factory) if raw_factory in factory_keys else 0
    with cols_common[2]:
        factory_short = st.selectbox(
            "工廠",
            factory_keys,
            index=factory_idx,
            key="fpoe_factory",
        )
    factory_data = factories[factory_short]

    # 採購日期 + 客戶下單日期 + 負責採購 + REVISED
    cols_dates = st.columns(4)
    purchase_date = _parse_date_flexible(first_row.get("工廠下單日期", "")) or date.today()
    with cols_dates[0]:
        order_date_input = st.date_input(
            "採購日期(工廠下單)",
            value=purchase_date,
            key="fpoe_order_date",
        )
    cust_order_date = _parse_date_flexible(first_row.get("客戶下單日期", "")) or purchase_date
    with cols_dates[1]:
        customer_order_date = st.date_input(
            "客戶下單日期",
            value=cust_order_date,
            key="fpoe_cust_date",
        )
    with cols_dates[2]:
        purchase_responsible = st.text_input(
            "負責採購(只印 PDF)",
            value="Amy",
            key="fpoe_pic",
        )
    with cols_dates[3]:
        is_revised = st.checkbox(
            "REVISED(蓋紅章)",
            value=False,
            key="fpoe_revised",
        )

    # 發出公司 + 字首(用於 PDF 模板選擇)
    cols_issuing = st.columns(3)
    parsed_po = parse_glocom_po_no(loaded_po) or {}
    detected_prefix = parsed_po.get("prefix", "G")
    if detected_prefix == "G":
        issuing_default = "GLOCOM"
    elif detected_prefix in ("ET", "EW"):
        issuing_default = "EUSWAY"
    else:
        issuing_default = "GLOCOM"

    with cols_issuing[0]:
        issuing_company = st.selectbox(
            "發出公司",
            ["GLOCOM", "EUSWAY"],
            index=0 if issuing_default == "GLOCOM" else 1,
            key="fpoe_issuing",
        )

    # ─── 各品項編輯 ─
    st.markdown("### 📦 品項清單(每筆都可編輯)")

    # 收集每品項的編輯結果
    edited_items = []

    for idx, row in matches.iterrows():
        record_id = _safe_str(row.get("_record_id", ""))
        with st.container(border=True):
            st.markdown(f"#### 品項 {idx + 1} `(record_id: {record_id[:12]}...)`")

            # 第 1 列:P/N + 數量
            c1 = st.columns(2)
            with c1[0]:
                pn = st.text_input(
                    "P/N",
                    value=_safe_str(row.get("P/N", "")),
                    key=f"fpoe_pn_{idx}",
                )
            with c1[1]:
                raw_qty = _safe_str(row.get("Order Q'TY\n (PCS)", ""))
                qty_val = _safe_int(raw_qty, 0)
                qty = st.number_input(
                    "數量 (PCS)",
                    value=qty_val,
                    min_value=0,
                    step=1,
                    key=f"fpoe_qty_{idx}",
                )

            # 第 2 列:工廠交期 + Ship Date
            c2 = st.columns(2)
            with c2[0]:
                fdue = _parse_date_flexible(row.get("工廠交期", "")) or date.today()
                factory_due = st.date_input(
                    "工廠交期",
                    value=fdue,
                    key=f"fpoe_fdue_{idx}",
                )
            with c2[1]:
                sdate = _parse_date_flexible(row.get("Ship date", "")) or factory_due
                ship_date = st.date_input(
                    "Ship Date",
                    value=sdate,
                    key=f"fpoe_sdate_{idx}",
                )

            # 第 3 列:銷貨金額 + 接單金額
            c3 = st.columns(2)
            sale_amt = _safe_float(row.get("銷貨金額", 0))
            order_amt = _safe_float(row.get("接單金額", 0))
            with c3[0]:
                new_sale = st.number_input(
                    "銷貨金額(工廠端)",
                    value=sale_amt,
                    min_value=0.0,
                    step=0.01,
                    format="%.2f",
                    key=f"fpoe_sale_{idx}",
                )
            with c3[1]:
                new_order = st.number_input(
                    "接單金額(客戶端)",
                    value=order_amt,
                    min_value=0.0,
                    step=0.01,
                    format="%.2f",
                    key=f"fpoe_order_{idx}",
                )

            # 顯示反推單價(只供 PDF 顯示用,不存)
            unit_price_display = (new_sale / qty) if qty > 0 else 0
            st.caption(
                f"💡 反推工廠單價 = 銷貨金額 ÷ 數量 = "
                f"**{unit_price_display:,.4f}**(只用於印 PDF,不存 Teable)"
            )

            # 規格(3 Tab)
            st.markdown("**規格 (印在 PO 上的『產品規格』欄)**")
            current_spec = _safe_str(row.get("客戶要求注意事項", ""))
            spec_text = render_spec_editor(idx, current_spec, pn)

            # 收集這品項的編輯結果
            edited_items.append({
                "record_id": record_id,
                "part_number": pn,
                "quantity": qty,
                "factory_due": factory_due,
                "ship_date": ship_date,
                "sale_amount": new_sale,
                "order_amount": new_order,
                "unit_price_display": unit_price_display,
                "spec_text": spec_text,
            })

    # ─── 動作按鈕 ─
    st.divider()
    st.markdown("### 動作")
    st.caption("**📄 產 PDF** 不會動 Teable;**💾 更新 Teable** 直接覆蓋原 record(不保留歷史版本)。")

    btn_cols = st.columns([2, 2, 1])
    with btn_cols[0]:
        do_pdf = st.button(
            "📄 產 PDF",
            type="primary",
            use_container_width=True,
            key="fpoe_do_pdf",
        )
    with btn_cols[1]:
        do_update = st.button(
            "💾 更新 Teable",
            type="secondary",
            use_container_width=True,
            key="fpoe_do_update",
        )
    with btn_cols[2]:
        if st.button("🚪 退出編輯", use_container_width=True, key="fpoe_exit"):
            keys_to_clear = [k for k in list(st.session_state.keys()) if k.startswith("fpoe_")]
            for k in keys_to_clear:
                del st.session_state[k]
            st.rerun()

    if not (do_pdf or do_update):
        st.stop()

    # 共用驗證
    empty_specs = [it["part_number"] for it in edited_items if not (it["spec_text"] or "").strip()]
    if empty_specs:
        st.error(f"❌ 以下品項規格沒填:**{', '.join(empty_specs)}**。請填妥再操作。")
        st.stop()

    # ─── 產 PDF ─
    if do_pdf:
        # 組 po_ctx
        po_items = []
        total_amount = 0.0
        for it in edited_items:
            po_items.append({
                "part_number": it["part_number"],
                "spec_text": it["spec_text"],
                "quantity": it["quantity"],
                "panel_qty": None,
                "unit_price": it["unit_price_display"],
                "amount": round(it["sale_amount"], 2),
                "delivery_date": it["factory_due"],
                "delivery_note": "",
            })
            total_amount += it["sale_amount"]

        po_ctx = {
            "po_no": loaded_po,
            "order_type": parsed_po.get("order_type", ""),
            "issuing_company": issuing_company,
            "order_date": order_date_input,
            "customer_name": customer_name,
            "customer_po_no": customer_po_no,
            "factory_short": factory_short,
            "factory": factory_data,
            "items": po_items,
            "total_amount": round(total_amount, 2),
            "currency": factory_data.get("default_currency", "NT$"),
            "payment_terms": factory_data.get("default_payment_terms", ""),
            "shipment_method": factory_data.get("default_shipment", "待通知"),
            "ship_to_default": factory_data.get("default_ship_to", "待通知"),
            "purchase_responsible": purchase_responsible,
            "is_revised": is_revised,
        }

        try:
            from core.pdf_generator import generate_po_files
            with st.spinner("產生 docx + PDF..."):
                result = generate_po_files(po_ctx)
            docx_path = result.get("docx_path")
            pdf_path = result.get("pdf_path")
            if result.get("error"):
                st.warning(result["error"])
        except Exception as e:
            st.error(f"❌ 產 PDF 失敗:{e}")
            with st.expander("錯誤詳情"):
                st.code(traceback.format_exc())
            st.stop()

        st.success(f"✅ PDF / DOCX 已產生(訂單編號 {loaded_po}{'(REVISED)' if is_revised else ''})")

        d_cols = st.columns(2)
        if docx_path and Path(docx_path).exists():
            with d_cols[0]:
                with open(docx_path, "rb") as f:
                    st.download_button(
                        "📥 下載 DOCX",
                        data=f.read(),
                        file_name=Path(docx_path).name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key="fpoe_dl_docx",
                    )
        if pdf_path and Path(pdf_path).exists():
            with d_cols[1]:
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        "📥 下載 PDF",
                        data=f.read(),
                        file_name=Path(pdf_path).name,
                        mime="application/pdf",
                        use_container_width=True,
                        key="fpoe_dl_pdf",
                    )

        st.info("📝 PDF 確認 OK 後,可按【💾 更新 Teable】把資料寫回主表。")

    # ─── 更新 Teable ─
    if do_update:
        # 確認 dialog 不能用 st.dialog (太新) — 用 expander+確認鈕的方式
        st.warning("⚠️ 即將把以下變更直接覆蓋到 Teable 主表(原資料會被取代,不保留歷史)。")

        # 顯示 diff 摘要
        st.markdown("**將更新的欄位**")
        diff_rows = []
        for it in edited_items:
            diff_rows.append({
                "P/N": it["part_number"],
                "數量": it["quantity"],
                "工廠": factory_short,
                "工廠交期": it["factory_due"].strftime("%Y/%m/%d"),
                "Ship Date": it["ship_date"].strftime("%Y/%m/%d"),
                "銷貨金額": f"{it['sale_amount']:,.2f}",
                "接單金額": f"{it['order_amount']:,.2f}",
                "規格": (it["spec_text"][:50] + "...") if len(it["spec_text"]) > 50 else it["spec_text"],
            })
        st.dataframe(pd.DataFrame(diff_rows), use_container_width=True, hide_index=True)

        if not st.session_state.get("fpoe_confirm_update", False):
            if st.button("✅ 確認覆蓋 Teable", type="primary", key="fpoe_confirm_btn"):
                st.session_state.fpoe_confirm_update = True
                st.rerun()
            st.stop()

        # 執行 PATCH
        success_count = 0
        fail_count = 0
        errors = []

        progress = st.progress(0, text="更新中...")

        for i, it in enumerate(edited_items):
            if not it["record_id"]:
                fail_count += 1
                errors.append(f"品項 {i+1}: record_id 缺失")
                continue

            fields_payload = {
                "客戶": customer_name,
                "PO#": customer_po_no,
                "P/N": it["part_number"],
                "Order Q'TY\n (PCS)": it["quantity"],
                "工廠": factory_short,
                "工廠交期": it["factory_due"].strftime("%Y/%m/%d"),
                "Ship date": it["ship_date"].strftime("%Y/%m/%d"),
                "銷貨金額": round(it["sale_amount"], 2),
                "接單金額": round(it["order_amount"], 2),
                "客戶要求注意事項": it["spec_text"],
                "工廠下單日期": order_date_input.strftime("%Y/%m/%d"),
                "客戶下單日期": customer_order_date.strftime("%Y/%m/%d"),
            }

            r = patch_teable_record(table_url, headers, it["record_id"], fields_payload)
            if r.get("success"):
                success_count += 1
            else:
                fail_count += 1
                errors.append(f"品項 {i+1} (P/N {it['part_number']}): {r.get('error', '')}")

            progress.progress((i + 1) / len(edited_items), text=f"已更新 {i + 1}/{len(edited_items)}")

        progress.empty()
        st.session_state.fpoe_confirm_update = False  # reset

        if success_count > 0:
            st.success(f"✅ 已更新 {success_count} 筆 record")
            st.info("⚠️ 請按側邊欄 **Refresh** 才會看到主表更新。")
        if fail_count > 0:
            st.error(f"❌ 失敗 {fail_count} 筆")
            with st.expander("失敗詳情"):
                for e in errors:
                    st.text(e)
