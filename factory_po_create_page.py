# -*- coding: utf-8 -*-
"""
建立工廠 PO 頁面 v2。

主要更新(v2):
- Step 5 規格改成「新料號完整規格」+「舊料號簡短規格」雙模式
- 規格輸出改為「一列」(用分號 ; 串接,不換行),不再拉長頁面
- Step 6 確認鍵拆成兩個獨立按鈕:【產 PDF】+【寫回 Teable】
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


def load_factories() -> dict:
    if not FACTORIES_JSON.exists():
        return {}
    try:
        with open(FACTORIES_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
        all_factories = data.get("factories", {})
        return {k: v for k, v in all_factories.items() if v.get("is_active")}
    except Exception as e:
        st.error(f"factories.json 讀取失敗:{e}")
        return {}


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


def calc_next_po_number(orders_df: pd.DataFrame, prefix: str) -> str:
    if orders_df.empty or COL_GLOCOM_PO not in orders_df.columns:
        roc_year = datetime.now().year - 1911
        return f"{prefix}{roc_year}0001-01"

    pattern = re.compile(rf"^{re.escape(prefix)}(\d{{3}})(\d{{4}})-(\d+)$")
    max_year_serial = (0, 0)

    for raw_po in orders_df[COL_GLOCOM_PO].dropna().astype(str):
        m = pattern.match(raw_po.strip())
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


def build_po_context_from_new_order(
    new_po_no, parsed, factory_data, factory_short, issuing_company,
    factory_unit_prices, factory_due_dates, item_specs,
    purchase_responsible, order_date, is_revised,
):
    """item_specs: dict {part_no: spec_text} 每品項各自的規格(一列字串)"""
    items = []
    for it in parsed.items:
        f_price = float(factory_unit_prices.get(it.part_number, 0.0))
        f_due = factory_due_dates.get(it.part_number)
        spec_text = item_specs.get(it.part_number, "") or it.description
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


def build_full_spec_oneline(selections: dict, extra_text: str) -> str:
    """新料號完整規格 → 一列字串 (用 ; 串接)
    例: Material: 4L; Tg170; Board thickness: 1.6mm; ...
    """
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


# ─── 主入口 ───────────────────────────────────────
def render_factory_po_create_page(orders: pd.DataFrame, table_url: str, headers: dict):
    st.subheader("📝 建立工廠 PO(新流程)")
    st.caption("上傳客戶 PO PDF → 自動解析 → 選工廠/規格 → 產出工廠 PO PDF / 寫回 Teable")

    factories = load_factories()
    if not factories:
        st.error("data/factories.json 讀取失敗或無工廠資料")
        st.stop()

    # ─── Step 1: 上傳 PDF ─
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

    # ─── Step 2: 解析 ─
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

    # ─── Step 3: 工廠 + 字首 + 編號 ─
    st.markdown("### Step 3. 工廠 / 發出公司 / 字首 / 編號")
    s3_cols = st.columns(3)

    factory_keys = sorted(factories.keys())
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
    new_po_no = calc_next_po_number(orders, chosen_internal)

    cols_no = st.columns([2, 1, 1])
    with cols_no[0]:
        st.success(f"📋 新訂單編號:**`{new_po_no}`**")
    with cols_no[1]:
        order_date_input = st.date_input("採購日期", value=date.today(), key="fpo_order_date")
    with cols_no[2]:
        is_revised = st.checkbox("REVISED", value=False, key="fpo_revised")
    cols_pic = st.columns([1, 3])
    with cols_pic[0]:
        purchase_responsible = st.text_input("負責採購", value="Amy", key="fpo_pic")

    # ─── Step 4: 工廠單價 + 出貨日 ─
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

    # ─── Step 5: 每品項規格(雙模式)─
    st.markdown("### Step 5. 每品項產品規格")
    st.caption("規格會印成「一列」放在工廠 PO 的『產品規格』欄。請選擇填寫模式。")

    item_specs = {}  # {part_no: spec_text}

    for idx, it in enumerate(parsed.items):
        with st.container(border=True):
            st.markdown(f"**品項 {idx + 1}:`{it.part_number}`**(數量 {it.quantity})")

            mode = st.radio(
                "規格類型",
                ["新料號(完整規格)", "舊料號(簡短規格)"],
                horizontal=True,
                key=f"fpo_spec_mode_{idx}",
            )

            if mode == "新料號(完整規格)":
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
                    "其他規格 / 補充(會接在後面,可空)",
                    placeholder="例如: 排版 4up panel; 須加西拓 UL Logo + Date Code (YYWW)",
                    key=f"fpo_spec_full_extra_{idx}",
                )
                spec_text = build_full_spec_oneline(selections, extra)

            else:  # 舊料號模式
                spec_text = st.text_input(
                    "舊料號規格(直接輸入一列)",
                    placeholder="例如: 舊料號;4up, 但規格須改為 Tg170; ENIG: 2u\"",
                    key=f"fpo_spec_short_{idx}",
                )

            if spec_text:
                st.success(f"📋 印在 PO 上:`{spec_text}`")
            else:
                st.info("⚠️ 規格還沒填,印出來會是空的(或 fallback 客戶描述)")

            item_specs[it.part_number] = spec_text

    # ─── Step 6: 兩個獨立按鈕 ─
    st.divider()
    st.markdown("### Step 6. 產出 / 寫回")
    st.caption("**產 PDF** 和 **寫回 Teable** 是分開的,可以分別執行。建議先按產 PDF 預覽 → OK 再按寫回 Teable。")

    btn_cols = st.columns([2, 2, 1])
    with btn_cols[0]:
        do_pdf = st.button(
            "📄 產 PDF",
            type="primary",
            use_container_width=True,
            key="fpo_do_pdf",
        )
    with btn_cols[1]:
        do_writeback = st.button(
            "💾 寫回 Teable",
            type="secondary",
            use_container_width=True,
            key="fpo_do_writeback",
        )
    with btn_cols[2]:
        if st.button("🗑️ 全部清除", use_container_width=True, key="fpo_clear_bottom"):
            _reset_form()
            st.rerun()

    if not (do_pdf or do_writeback):
        st.stop()

    zero_price_pns = [pn for pn, price in factory_unit_prices.items() if price == 0]
    if zero_price_pns:
        st.warning(f"⚠️ 工廠單價為 0 的品項:{', '.join(zero_price_pns)}")

    empty_spec_pns = [pn for pn, sp in item_specs.items() if not sp.strip()]
    if empty_spec_pns:
        st.warning(f"⚠️ 規格沒填的品項:{', '.join(empty_spec_pns)}(會用客戶 PO 描述當 fallback)")

    # ─── 產 PDF 流程 ─
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
                    "如需產 EW 字首 PO,請先建立並上傳 PO_EUSWAY.docx。"
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

    # ─── 寫回 Teable 流程 ─
    if do_writeback:
        records_fields = []
        for it in parsed.items:
            f_price = factory_unit_prices.get(it.part_number, 0.0)
            f_due = factory_due_dates.get(it.part_number)
            f_due_str = f_due.strftime("%Y/%m/%d") if f_due else ""
            order_date_str = order_date_input.strftime("%Y/%m/%d")
            cust_date_str = parsed.po_date or order_date_str
            spec_text = item_specs.get(it.part_number, "") or it.description

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
            st.info("⚠️ 請按側邊欄 **Refresh** 才會看到主表更新。")
        if wb_result["failed"] > 0:
            st.error(f"❌ 寫入失敗 {wb_result['failed']} 筆")
            with st.expander("失敗原因"):
                for err in wb_result["errors"]:
                    st.text(err)
