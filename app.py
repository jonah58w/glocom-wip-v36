import pandas as pd
import streamlit as st
from PIL import Image

from config import *
from utils import *
from teable_api import load_orders, upsert_to_teable, patch_record_by_id, update_working_orders_local
from match_engine import build_record_for_match, build_teable_match_records, find_best_match_by_4fields, dedupe_import_df_by_key
from factory_parsers import read_import_dataframe
from text_ocr_parsers import (
    parse_quick_text_line, parse_email_text_to_rows, ocr_image_to_text, extract_po_from_text,
    extract_date_from_text, infer_wip_from_text, infer_customer_tags_from_text, infer_remark_from_text,
)
from reports import *

st.set_page_config(page_title="GLOCOM Control Tower", page_icon="🏭", layout="wide")
if "manual_review_queue" not in st.session_state:
    st.session_state.manual_review_queue = []

st.markdown("""
<style>
.portal-box {padding:18px 20px;border:1px solid rgba(120,120,120,.22);border-radius:16px;background:rgba(255,255,255,.03);margin-bottom:14px;}
.portal-title {font-size:1.2rem;font-weight:700;margin-bottom:4px;}
.tag-chip {display:inline-block;padding:4px 10px;margin:2px 6px 2px 0;border-radius:999px;font-size:0.82rem;border:1px solid rgba(120,120,120,.25);background:rgba(255,255,255,.05);}
.wip-chip {display:inline-block;padding:4px 10px;border-radius:999px;font-size:0.82rem;font-weight:600;}
</style>
""", unsafe_allow_html=True)


def refresh_after_update():
    st.cache_data.clear()
    st.rerun()


@st.cache_data(ttl=60)
def cached_load_orders():
    return load_orders()


orders, api_status, api_text = cached_load_orders()
if orders.empty:
    st.title("🏭 GLOCOM Control Tower")
    show_no_data_layout()
    st.stop()

po_col = get_first_matching_column(orders, PO_CANDIDATES)
customer_col = get_first_matching_column(orders, CUSTOMER_CANDIDATES)
part_col = get_first_matching_column(orders, PART_CANDIDATES)
qty_col = get_first_matching_column(orders, QTY_CANDIDATES)
factory_col = get_first_matching_column(orders, FACTORY_CANDIDATES)
wip_col = get_first_matching_column(orders, WIP_CANDIDATES)
factory_due_col = get_first_matching_column(orders, FACTORY_DUE_CANDIDATES)
ship_date_col = get_first_matching_column(orders, SHIP_DATE_CANDIDATES)
remark_col = get_first_matching_column(orders, REMARK_CANDIDATES)
customer_tag_col = get_first_matching_column(orders, CUSTOMER_TAG_CANDIDATES)

query = st.query_params
customer_param = query.get("customer", None)
if customer_param:
    st.title("GLOCOM Order Status")
    st.caption("Customer WIP Progress")
    if not customer_col:
        st.error("Customer column not found")
        st.stop()
    customer_series = get_series_by_col(orders, customer_col)
    cust_orders = orders[customer_series.astype(str).str.strip().str.lower() == str(customer_param).strip().lower()].copy()
    if cust_orders.empty:
        st.warning("No orders found")
        st.stop()
    show_metrics(cust_orders, wip_col)
    st.divider()
    render_customer_portal(cust_orders, po_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col)
    portal_cols = [c for c in [po_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col] if c and c in cust_orders.columns]
    csv_data = cust_orders[portal_cols].to_csv(index=False).encode("utf-8-sig")
    st.download_button("Download WIP CSV", data=csv_data, file_name=f"{customer_param}_wip.csv", mime="text/csv")
    st.stop()

st.title("🏭 GLOCOM Control Tower")
st.caption("Internal PCB Production Monitoring System")
with st.expander("Debug"):
    st.write("API Status:", api_status)
    st.write("TABLE_URL:", TABLE_URL)
    st.write("Token loaded:", bool(TEABLE_TOKEN))
    st.write("Columns:", list(orders.columns))
    if isinstance(api_text, str):
        st.text(api_text[:1200])

st.sidebar.title("GLOCOM Internal")
st.sidebar.link_button("Open Teable", TEABLE_WEB_URL, use_container_width=True)
menu = st.sidebar.radio("功能選單", ["Dashboard", "Factory Load", "Delayed Orders", "Shipment Forecast", "Orders", "Customer Preview", "Sandy 內部 WIP", "Sandy 銷貨底", "新訂單 WIP", "Import / Update"])
if st.sidebar.button("Refresh"):
    refresh_after_update()


def classify_and_update_factory_row(current_df, import_row, import_cols):
    imported_rec = build_record_for_match(
        po_value=import_row.get(import_cols["po"], "") if import_cols["po"] else "",
        part_value=import_row.get(import_cols["part"], "") if import_cols["part"] else "",
        qty_value=import_row.get(import_cols["qty"], "") if import_cols["qty"] else "",
        due_value=import_row.get(import_cols["due"], "") if import_cols["due"] else "",
    )
    present_count = sum([bool(imported_rec["po"]), bool(imported_rec["part_no"]), imported_rec["qty"] is not None, bool(imported_rec["factory_due_date"])])
    if present_count < 3:
        return {"success": False, "action": "MANUAL_REVIEW", "message": "匯入資料可用比對欄位不足 3 項", "match_info": {"score": present_count, "matched_fields": [], "candidates": []}}
    match_result = find_best_match_by_4fields(imported_rec, build_teable_match_records(current_df, po_col, part_col, qty_col, factory_due_col))
    if match_result["status"] != "matched":
        return {"success": False, "action": "MANUAL_REVIEW", "message": "比對不足 3 項相同，或有多筆候選", "match_info": match_result}
    target = match_result["target"]
    record_id = target.get("_record_id", "")
    payload_fields = {}
    if wip_col and import_cols["wip"]:
        v = normalize_wip_value(import_row.get(import_cols["wip"], ""))
        if v:
            payload_fields[wip_col] = v
    if customer_col and import_cols["customer"]:
        v = safe_text(import_row.get(import_cols["customer"], ""))
        if v:
            payload_fields[customer_col] = v
    if ship_date_col and import_cols["ship"]:
        v = normalize_due_text(import_row.get(import_cols["ship"], ""))
        if v:
            payload_fields[ship_date_col] = v
    if factory_due_col and import_cols["due"]:
        v = normalize_due_text(import_row.get(import_cols["due"], ""))
        if v:
            payload_fields[factory_due_col] = v
    if remark_col and import_cols["remark"]:
        v = safe_text(import_row.get(import_cols["remark"], ""))
        if v:
            payload_fields[remark_col] = v
    if customer_tag_col and import_cols["tag"]:
        raw = import_row.get(import_cols["tag"], "")
        tags = raw if isinstance(raw, list) else parse_tags_from_text(raw)
        if tags:
            payload_fields[customer_tag_col] = build_tags_value(tags)
    if not payload_fields:
        return {"success": False, "action": "SKIP", "message": "沒有可更新欄位", "match_info": match_result}
    success, msg = patch_record_by_id(record_id, payload_fields)
    return {"success": success, "action": "UPDATED" if success else "FAILED", "message": msg, "match_info": match_result, "record_id": record_id, "payload_fields": payload_fields}


def build_manual_review_item(import_row, import_cols, match_info, reason):
    item = {
        "PO#": safe_text(import_row.get(import_cols["po"], "")) if import_cols["po"] else "",
        "Part No": safe_text(import_row.get(import_cols["part"], "")) if import_cols["part"] else "",
        "Qty": safe_text(import_row.get(import_cols["qty"], "")) if import_cols["qty"] else "",
        "Factory Due Date": safe_text(import_row.get(import_cols["due"], "")) if import_cols["due"] else "",
        "New WIP": safe_text(import_row.get(import_cols["wip"], "")) if import_cols["wip"] else "",
        "Remark": safe_text(import_row.get(import_cols["remark"], "")) if import_cols["remark"] else "",
        "Reason": reason,
        "Best Score": match_info.get("score", 0),
        "Matched Fields": ", ".join(match_info.get("matched_fields", [])),
        "Candidates": [],
    }
    for c in match_info.get("candidates", [])[:5]:
        rec = c.get("record", {})
        src = rec.get("_source_row")
        item["Candidates"].append({
            "record_id": rec.get("_record_id", ""),
            "score": c.get("score", 0),
            "matched_fields": ", ".join(c.get("matched_fields", [])),
            "PO#": safe_text(src.get(po_col, "")) if src is not None and po_col else "",
            "Part No": safe_text(src.get(part_col, "")) if src is not None and part_col else "",
            "Qty": safe_text(src.get(qty_col, "")) if src is not None and qty_col else "",
            "Factory Due Date": safe_text(src.get(factory_due_col, "")) if src is not None and factory_due_col else "",
            "WIP": safe_text(src.get(wip_col, "")) if src is not None and wip_col else "",
        })
    return item


if menu == "Dashboard":
    show_metrics(orders, wip_col)
    st.divider()
    left, right = st.columns(2)
    with left:
        show_factory_load(orders, factory_col)
    with right:
        show_shipment_forecast(orders, ship_date_col, po_col, customer_col, part_col, qty_col, factory_col, wip_col)
elif menu == "Factory Load":
    show_factory_load(orders, factory_col)
elif menu == "Delayed Orders":
    show_delayed_orders(orders, factory_due_col, po_col, customer_col, part_col, qty_col, factory_col, wip_col)
elif menu == "Shipment Forecast":
    show_shipment_forecast(orders, ship_date_col, po_col, customer_col, part_col, qty_col, factory_col, wip_col)
elif menu == "Orders":
    st.subheader("🔎 Filters")
    filtered = orders.copy()
    col1, col2 = st.columns(2)
    if customer_col:
        customer_series = get_series_by_col(filtered, customer_col)
        customer_options = ["All"] + sorted([str(x) for x in customer_series.dropna().unique().tolist()]) if customer_series is not None else ["All"]
        selected_customer = col1.selectbox("Customer", customer_options)
        if selected_customer != "All":
            filtered = filtered[get_series_by_col(filtered, customer_col).astype(str) == selected_customer]
    if wip_col:
        wip_series = get_series_by_col(filtered, wip_col)
        wip_options = ["All"] + sorted([str(x) for x in wip_series.dropna().unique().tolist()]) if wip_series is not None else ["All"]
        selected_wip = col2.selectbox("WIP Stage", wip_options)
        if selected_wip != "All":
            filtered = filtered[get_series_by_col(filtered, wip_col).astype(str) == selected_wip]
    show_orders_table(filtered)
elif menu == "Customer Preview":
    st.subheader("Customer Preview")
    st.caption("僅供內部預覽。客戶請直接使用 Teable View。")
    if not customer_col:
        st.error("Customer column not found in Teable data")
    else:
        customer_series = get_series_by_col(orders, customer_col)
        customers = sorted([str(x).strip() for x in customer_series.dropna().unique().tolist() if str(x).strip()]) if customer_series is not None else []
        if not customers:
            st.warning("No customers found")
        else:
            selected_customer = st.selectbox("Select customer to preview", customers)
            preview_df = orders[customer_series.astype(str).str.strip().str.lower() == selected_customer.strip().lower()].copy()
            preview_cols = [c for c in [po_col, customer_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col] if c and c in preview_df.columns]
            st.dataframe(preview_df[preview_cols], use_container_width=True, height=420)
elif menu == "Sandy 內部 WIP":
    show_sandy_internal_wip_report(orders)
elif menu == "Sandy 銷貨底":
    show_sandy_shipment_report(orders)
elif menu == "新訂單 WIP":
    show_new_orders_wip_report(orders)
else:
    st.subheader("Import / Update")
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Excel / CSV / TXT", "Manual Update", "Quick Text", "Email Text", "Image OCR"])
    with tab1:
        st.caption("匯入比對規則：PO#、Part No、Qty、Factory Due Date 四項中，符合 3 項以上且唯一候選者自動覆蓋 WIP；其餘進待人工確認。")
        uploaded = st.file_uploader("Upload Excel / CSV / TXT", type=["xlsx", "xls", "csv", "txt"])
        if uploaded is not None:
            try:
                import_df, source_type = read_import_dataframe(uploaded)
                st.info(f"Detected source type: {source_type}")
                st.dataframe(import_df, use_container_width=True, height=280)
                import_cols = {
                    "po": get_first_matching_column(import_df, PO_CANDIDATES),
                    "customer": get_first_matching_column(import_df, CUSTOMER_CANDIDATES),
                    "part": get_first_matching_column(import_df, PART_CANDIDATES),
                    "qty": get_first_matching_column(import_df, QTY_CANDIDATES),
                    "wip": get_first_matching_column(import_df, WIP_CANDIDATES),
                    "due": get_first_matching_column(import_df, FACTORY_DUE_CANDIDATES),
                    "ship": get_first_matching_column(import_df, SHIP_DATE_CANDIDATES),
                    "remark": get_first_matching_column(import_df, REMARK_CANDIDATES),
                    "tag": get_first_matching_column(import_df, CUSTOMER_TAG_CANDIDATES),
                }
                st.json(import_cols)
                deduped_import_df, duplicate_keys_in_file = dedupe_import_df_by_key(import_df, import_cols["po"], import_cols["part"], import_cols["qty"], import_cols["due"])
                if duplicate_keys_in_file:
                    st.warning(f"同一批匯入檔案中有重複 key，已自動去重。重複筆數：{len(duplicate_keys_in_file)}")
                if st.button("Batch Update from File"):
                    if not import_cols["wip"]:
                        st.error("匯入檔至少要能辨識出 WIP 欄位。")
                        st.stop()
                    ok_update_count = manual_review_count = skip_count = fail_count = 0
                    logs = []
                    working_orders = orders.copy()
                    manual_review_items = []
                    for _, row in deduped_import_df.iterrows():
                        result = classify_and_update_factory_row(working_orders, row, import_cols)
                        po_value = safe_text(row.get(import_cols["po"], "")) if import_cols["po"] else ""
                        part_value = safe_text(row.get(import_cols["part"], "")) if import_cols["part"] else ""
                        qty_value = safe_text(row.get(import_cols["qty"], "")) if import_cols["qty"] else ""
                        due_value = safe_text(row.get(import_cols["due"], "")) if import_cols["due"] else ""
                        if result["success"] and result["action"] == "UPDATED":
                            ok_update_count += 1
                            logs.append(f"[UPDATED] {po_value} | {part_value} | {qty_value} | {due_value}")
                            if result.get("record_id"):
                                working_orders = update_working_orders_local(working_orders, result["record_id"], result.get("payload_fields", {}))
                        elif result["action"] == "MANUAL_REVIEW":
                            manual_review_count += 1
                            manual_review_items.append(build_manual_review_item(row, import_cols, result.get("match_info", {}), result.get("message", "需要人工確認")))
                            logs.append(f"[MANUAL REVIEW] {po_value} | {part_value} | {qty_value} | {due_value}")
                        elif result["action"] == "SKIP":
                            skip_count += 1
                            logs.append(f"[SKIP] {po_value} | {part_value} | {qty_value} | {due_value} -> {result.get('message', '')}")
                        else:
                            fail_count += 1
                            logs.append(f"[FAILED] {po_value} | {part_value} | {qty_value} | {due_value} -> {result.get('message', '')}")
                    st.session_state.manual_review_queue = manual_review_items
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Updated WIP", ok_update_count)
                    c2.metric("Manual Review", manual_review_count)
                    c3.metric("Skipped", skip_count)
                    c4.metric("Failed", fail_count)
                    if logs:
                        st.text("\n".join(logs[:200]))
            except Exception as e:
                st.error(f"Import failed: {e}")
    with tab2:
        queue = st.session_state.get("manual_review_queue", [])
        default_po = default_wip = default_remark = ""
        if queue:
            review_df = pd.DataFrame([{k: x[k] for k in ["PO#", "Part No", "Qty", "Factory Due Date", "New WIP", "Reason", "Best Score", "Matched Fields"]} for x in queue])
            st.dataframe(review_df, use_container_width=True, height=240)
            selected_idx = st.selectbox("選擇一筆待人工確認資料", options=list(range(len(queue))), format_func=lambda i: f"{queue[i]['PO#']} | {queue[i]['Part No']} | {queue[i]['Qty']} | {queue[i]['New WIP']}")
            selected_item = queue[selected_idx]
            default_po, default_wip, default_remark = selected_item.get("PO#", ""), selected_item.get("New WIP", ""), selected_item.get("Remark", "")
            candidates = selected_item.get("Candidates", [])
            if candidates:
                st.dataframe(pd.DataFrame(candidates), use_container_width=True, height=180)
                chosen = st.selectbox("套用到候選 record_id", options=[""] + [c["record_id"] for c in candidates])
                if chosen and st.button("Apply to Selected Candidate"):
                    payload = {}
                    if wip_col and default_wip:
                        payload[wip_col] = normalize_wip_value(default_wip)
                    if remark_col and default_remark:
                        payload[remark_col] = default_remark
                    success, msg = patch_record_by_id(chosen, payload)
                    if success:
                        st.success("Applied to selected candidate")
                        refresh_after_update()
                    else:
                        st.error(msg)
        with st.form("manual_update_form"):
            po_input = st.text_input("PO#", value=default_po)
            wip_input = st.text_input("WIP", value=default_wip)
            ship_input = st.text_input("Ship Date")
            tags_input = st.multiselect("Customer Remark Tags", TAG_OPTIONS)
            remark_input = st.text_area("Remark", value=default_remark)
            submitted = st.form_submit_button("Update This PO")
        if submitted:
            updates = {}
            if wip_col and wip_input.strip():
                updates[wip_col] = normalize_wip_value(wip_input.strip())
            if ship_date_col and ship_input.strip():
                updates[ship_date_col] = ship_input.strip()
            if customer_tag_col:
                updates[customer_tag_col] = build_tags_value(tags_input)
            if remark_col:
                updates[remark_col] = remark_input.strip()
            success, msg = upsert_to_teable(orders, po_col, po_input.strip(), updates)
            if success:
                st.success(f"{po_input.strip()} updated successfully")
                refresh_after_update()
            else:
                st.error(msg)
    with tab3:
        st.code("PO78310 | Shipping | 2026-03-20 | Partial Shipment, Shipped | ready to ship")
        quick_text = st.text_area("Paste Quick Text", height=220)
        if st.button("Batch Update from Quick Text"):
            lines = [x.strip() for x in quick_text.splitlines() if x.strip()]
            ok_count = fail_count = 0
            logs = []
            for line in lines:
                parsed = parse_quick_text_line(line)
                if not parsed:
                    continue
                updates = {}
                if wip_col and parsed["wip"]:
                    updates[wip_col] = normalize_wip_value(parsed["wip"])
                if ship_date_col and parsed["ship_date"]:
                    updates[ship_date_col] = parsed["ship_date"]
                if customer_tag_col:
                    updates[customer_tag_col] = build_tags_value(parsed["tags"])
                if remark_col and parsed["remark"]:
                    updates[remark_col] = parsed["remark"]
                success, msg = upsert_to_teable(orders, po_col, parsed["po"], updates)
                if success:
                    ok_count += 1
                else:
                    fail_count += 1
                    logs.append(f"{parsed['po']} -> {msg}")
            st.success(f"Quick text update finished. Success: {ok_count}, Failed: {fail_count}")
            if logs:
                st.text("\n".join(logs[:50]))
    with tab4:
        email_text = st.text_area("Paste Email / Text Content", height=220)
        if st.button("Parse Email Text"):
            rows = parse_email_text_to_rows(email_text)
            if rows:
                st.dataframe(pd.DataFrame(rows), use_container_width=True, height=260)
            else:
                st.warning("No parsable rows found")
    with tab5:
        uploaded_img = st.file_uploader("Upload PNG / JPG / JPEG", type=["png", "jpg", "jpeg"], key="ocr_uploader")
        if uploaded_img is not None:
            image = Image.open(uploaded_img)
            st.image(image, caption="Uploaded Image", use_container_width=True)
            ocr_text = ocr_image_to_text(image)
            st.text_area("OCR Raw Text", value=ocr_text, height=220)
            if not str(ocr_text).startswith("OCR_ERROR:"):
                guessed_po = extract_po_from_text(ocr_text)
                guessed_wip = infer_wip_from_text(ocr_text)
                guessed_date = extract_date_from_text(ocr_text)
                guessed_tags = infer_customer_tags_from_text(ocr_text)
                guessed_remark = infer_remark_from_text(ocr_text)
                with st.form("ocr_update_form"):
                    po_input = st.text_input("PO#", value=guessed_po)
                    wip_input = st.text_input("WIP", value=guessed_wip)
                    ship_input = st.text_input("Ship Date", value=guessed_date)
                    tags_input = st.multiselect("Customer Remark Tags", TAG_OPTIONS, default=[t for t in guessed_tags if t in TAG_OPTIONS])
                    remark_input = st.text_area("Remark", value=guessed_remark, height=120)
                    submitted_ocr = st.form_submit_button("Update to Teable")
                if submitted_ocr:
                    updates = {}
                    if wip_col and wip_input.strip():
                        updates[wip_col] = normalize_wip_value(wip_input.strip())
                    if ship_date_col and ship_input.strip():
                        updates[ship_date_col] = ship_input.strip()
                    if customer_tag_col:
                        updates[customer_tag_col] = build_tags_value(tags_input)
                    if remark_col and remark_input.strip():
                        updates[remark_col] = remark_input.strip()
                    success, msg = upsert_to_teable(orders, po_col, po_input.strip(), updates)
                    if success:
                        st.success(f"{po_input.strip()} updated successfully from OCR")
                        refresh_after_update()
                    else:
                        st.error(msg)

st.caption("Auto refresh cache: 60 seconds")
