
# -*- coding: utf-8 -*-
import pandas as pd
import streamlit as st
from PIL import Image

from config import *
from utils import *
from teable_api import (
    load_orders,
    upsert_to_teable,
    patch_record_by_id,
    update_working_orders_local,
)
from match_engine import (
    build_record_for_match,
    build_teable_match_records,
    find_best_match_by_4fields,
    dedupe_import_df_by_key,
)
from factory_parsers import read_import_dataframe
from text_ocr_parsers import (
    parse_quick_text_line,
    parse_email_text_to_rows,
    ocr_image_to_text,
    extract_po_from_text,
    extract_date_from_text,
    infer_wip_from_text,
    infer_customer_tags_from_text,
    infer_remark_from_text,
)
from reports import (
    show_dashboard_report,
    show_factory_load_report,
    show_delayed_orders_report,
    show_shipment_forecast_report,
    show_orders_report,
    show_customer_preview_report,
    show_sandy_internal_wip_report,
    show_sandy_shipment_report,
    show_new_orders_wip_report,
)

st.markdown(GLOBAL_STYLE, unsafe_allow_html=True)

if "manual_review_queue" not in st.session_state:
    st.session_state.manual_review_queue = []


def build_manual_review_item(import_row, import_cols, match_info, reason, teable_cols):
    import_po_col = import_cols.get("po")
    import_part_col = import_cols.get("part")
    import_qty_col = import_cols.get("qty")
    import_due_col = import_cols.get("factory_due")
    import_wip_col = import_cols.get("wip")
    import_remark_col = import_cols.get("remark")

    po_value = safe_text(import_row.get(import_po_col, "")) if import_po_col else ""
    part_value = safe_text(import_row.get(import_part_col, "")) if import_part_col else ""
    qty_value = safe_text(import_row.get(import_qty_col, "")) if import_qty_col else ""
    due_value = safe_text(import_row.get(import_due_col, "")) if import_due_col else ""
    wip_value = safe_text(import_row.get(import_wip_col, "")) if import_wip_col else ""
    remark_value = safe_text(import_row.get(import_remark_col, "")) if import_remark_col else ""

    candidate_desc = []
    for c in match_info.get("candidates", [])[:5]:
        rec = c.get("record", {})
        src_row = rec.get("_source_row")
        candidate_desc.append({
            "record_id": rec.get("_record_id", ""),
            "score": c.get("score", 0),
            "matched_fields": ", ".join(c.get("matched_fields", [])),
            "PO#": safe_text(src_row.get(teable_cols.get("po", ""), "")) if src_row is not None and teable_cols.get("po") else "",
            "Part No": safe_text(src_row.get(teable_cols.get("part", ""), "")) if src_row is not None and teable_cols.get("part") else "",
            "Qty": safe_text(src_row.get(teable_cols.get("qty", ""), "")) if src_row is not None and teable_cols.get("qty") else "",
            "Factory Due Date": safe_text(src_row.get(teable_cols.get("factory_due", ""), "")) if src_row is not None and teable_cols.get("factory_due") else "",
            "WIP": safe_text(src_row.get(teable_cols.get("wip", ""), "")) if src_row is not None and teable_cols.get("wip") else "",
        })

    return {
        "PO#": po_value,
        "Part No": part_value,
        "Qty": qty_value,
        "Factory Due Date": due_value,
        "New WIP": wip_value,
        "Remark": remark_value,
        "Reason": reason,
        "Best Score": match_info.get("score", 0),
        "Matched Fields": ", ".join(match_info.get("matched_fields", [])),
        "Candidates": candidate_desc,
    }


def classify_and_update_factory_row(current_df, teable_cols, import_row, import_cols):
    po_value = safe_text(import_row.get(import_cols.get("po", ""), "")) if import_cols.get("po") else ""
    part_value = clean_part_no(safe_text(import_row.get(import_cols.get("part", ""), ""))) if import_cols.get("part") else ""
    qty_value = safe_text(import_row.get(import_cols.get("qty", ""), "")) if import_cols.get("qty") else ""
    due_value = normalize_due_date_text(import_row.get(import_cols.get("factory_due", ""), "")) if import_cols.get("factory_due") else ""
    wip_value = normalize_wip_value(safe_text(import_row.get(import_cols.get("wip", ""), ""))) if import_cols.get("wip") else ""

    customer_value = safe_text(import_row.get(import_cols.get("customer", ""), "")) if import_cols.get("customer") else ""
    ship_value = normalize_due_date_text(import_row.get(import_cols.get("ship", ""), "")) if import_cols.get("ship") else ""
    remark_value = safe_text(import_row.get(import_cols.get("remark", ""), "")) if import_cols.get("remark") else ""

    raw_tags = import_row.get(import_cols.get("tags", ""), "") if import_cols.get("tags") else ""
    tags_value = raw_tags if isinstance(raw_tags, list) else parse_tags_from_text(raw_tags)

    imported_rec = build_record_for_match(
        po_value=po_value,
        part_value=part_value,
        qty_value=qty_value,
        due_value=due_value,
    )

    present_count = sum([
        bool(imported_rec["po"]),
        bool(imported_rec["part_no"]),
        imported_rec["qty"] is not None,
        bool(imported_rec["factory_due_date"]),
    ])

    if present_count < 3:
        return {
            "success": False,
            "action": "MANUAL_REVIEW",
            "message": "匯入資料可用比對欄位不足 3 項",
            "payload_fields": {},
            "match_info": {"score": present_count, "matched_fields": [], "candidates": []},
        }

    teable_match_records = build_teable_match_records(
        current_df=current_df,
        teable_po_col=teable_cols.get("po"),
        teable_part_col=teable_cols.get("part"),
        teable_qty_col=teable_cols.get("qty"),
        teable_factory_due_col=teable_cols.get("factory_due"),
    )

    match_result = find_best_match_by_4fields(imported_rec, teable_match_records)

    if match_result["status"] != "matched":
        return {
            "success": False,
            "action": "MANUAL_REVIEW",
            "message": "比對不足 3 項相同，或有多筆候選",
            "payload_fields": {},
            "match_info": match_result,
        }

    target = match_result["target"]
    record_id = target.get("_record_id", "")
    if not record_id:
        return {
            "success": False,
            "action": "MANUAL_REVIEW",
            "message": "找到候選但 record_id 缺失",
            "payload_fields": {},
            "match_info": match_result,
        }

    payload_fields = {}
    if teable_cols.get("wip") and wip_value:
        payload_fields[teable_cols["wip"]] = wip_value
    if teable_cols.get("customer") and customer_value:
        payload_fields[teable_cols["customer"]] = customer_value
    if teable_cols.get("ship") and ship_value:
        payload_fields[teable_cols["ship"]] = ship_value
    if teable_cols.get("factory_due") and due_value:
        payload_fields[teable_cols["factory_due"]] = due_value
    if teable_cols.get("remark") and remark_value:
        payload_fields[teable_cols["remark"]] = remark_value
    if teable_cols.get("tags") and tags_value:
        payload_fields[teable_cols["tags"]] = build_tags_value(tags_value, MULTI_SELECT_MODE)

    if not payload_fields:
        return {
            "success": False,
            "action": "SKIP",
            "message": "沒有可更新欄位",
            "payload_fields": {},
            "match_info": match_result,
        }

    success, msg = patch_record_by_id(record_id, payload_fields)
    return {
        "success": success,
        "action": "UPDATED" if success else "FAILED",
        "message": msg if success else f"更新失敗: {msg}",
        "payload_fields": payload_fields,
        "match_info": match_result,
        "record_id": record_id,
    }


def render_manual_update_tab(orders, detected_cols):
    st.markdown("手動更新單筆 WIP。")
    st.caption("若 Excel / Email / OCR 匯入時無法確認，請參考下方待人工確認清單，再手動更新。")

    queue = st.session_state.get("manual_review_queue", [])
    default_po = ""
    default_wip = ""
    default_remark = ""

    if queue:
        st.markdown("### 待人工確認清單")
        review_df = pd.DataFrame([
            {
                "PO#": x["PO#"],
                "Part No": x["Part No"],
                "Qty": x["Qty"],
                "Factory Due Date": x["Factory Due Date"],
                "New WIP": x["New WIP"],
                "Reason": x["Reason"],
                "Best Score": x["Best Score"],
                "Matched Fields": x["Matched Fields"],
            }
            for x in queue
        ])
        st.dataframe(review_df, use_container_width=True, height=240)

        selected_idx = st.selectbox(
            "選擇一筆待人工確認資料帶入下方表單",
            options=list(range(len(queue))),
            format_func=lambda i: f"{queue[i]['PO#']} | {queue[i]['Part No']} | {queue[i]['Qty']} | {queue[i]['New WIP']}",
        )
        selected_item = queue[selected_idx]

        with st.expander("查看候選比對資料"):
            candidates = selected_item.get("Candidates", [])
            if candidates:
                st.dataframe(pd.DataFrame(candidates), use_container_width=True, height=220)
            else:
                st.info("沒有候選資料，請直接用 PO 手動更新。")

        default_po = selected_item.get("PO#", "")
        default_wip = selected_item.get("New WIP", "")
        default_remark = selected_item.get("Remark", "")

        candidate_options = [c for c in selected_item.get("Candidates", []) if c.get("record_id")]
        if candidate_options:
            pick = st.selectbox(
                "或直接套用到指定候選 record_id",
                options=list(range(len(candidate_options))),
                format_func=lambda i: f"{candidate_options[i]['record_id']} | score={candidate_options[i]['score']} | {candidate_options[i]['PO#']} | {candidate_options[i]['Part No']}",
                key="manual_candidate_pick",
            )
            if st.button("套用到選定候選", key="apply_selected_candidate"):
                record_id = candidate_options[pick]["record_id"]
                payload_fields = {}
                if detected_cols.get("wip") and default_wip:
                    payload_fields[detected_cols["wip"]] = normalize_wip_value(default_wip)
                if detected_cols.get("remark") and default_remark:
                    payload_fields[detected_cols["remark"]] = default_remark
                success, msg = patch_record_by_id(record_id, payload_fields)
                if success:
                    st.success("已套用到指定候選。")
                    refresh_after_update()
                else:
                    st.error(msg)
    else:
        st.info("目前沒有待人工確認清單。")

    with st.form("manual_update_form"):
        po_input = st.text_input("PO#", value=default_po, placeholder="例如：PO78310")
        wip_input = st.text_input("WIP", value=default_wip, placeholder="例如：Shipping")
        ship_input = st.text_input("Ship Date", placeholder="例如：2026-03-20")
        tags_input = st.multiselect("Customer Remark Tags", TAG_OPTIONS)
        remark_input = st.text_area("Remark", value=default_remark, placeholder="給客戶看的備註")
        submitted = st.form_submit_button("Update This PO")

    if submitted:
        if not po_input.strip():
            st.error("PO# is required")
            return

        updates = {}
        if detected_cols.get("wip") and wip_input.strip():
            updates[detected_cols["wip"]] = normalize_wip_value(wip_input.strip())
        if detected_cols.get("ship") and ship_input.strip():
            updates[detected_cols["ship"]] = ship_input.strip()
        if detected_cols.get("tags"):
            updates[detected_cols["tags"]] = build_tags_value(tags_input, MULTI_SELECT_MODE)
        if detected_cols.get("remark"):
            updates[detected_cols["remark"]] = remark_input.strip()

        success, msg = upsert_to_teable(
            current_df=orders,
            po_col_name=detected_cols["po"],
            po_value=po_input.strip(),
            updates=updates,
        )
        if success:
            st.success(f"{po_input.strip()} updated successfully")
            refresh_after_update()
        else:
            st.error(msg)


def render_import_update(orders, detected_cols):
    st.subheader("Import / Update")

    if not detected_cols.get("po"):
        st.error("PO column not found. 請確認 Teable 表有 PO# 欄位。")
        st.stop()

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Excel / CSV", "Manual Update", "Quick Text", "Email Text", "Image OCR"])

    with tab1:
        st.markdown("上傳工廠報表匯入 Teable。")
        st.caption("匯入比對規則：PO#、Part No、Qty、Factory Due Date 四項中，符合 3 項以上且唯一候選者自動覆蓋 WIP；符合 2 項以下或多筆候選者，不自動覆蓋，改列入待人工確認清單。")
        uploaded = st.file_uploader("Upload Excel / CSV", type=["xlsx", "xls", "csv"], key="file_uploader")

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
                    "factory_due": get_first_matching_column(import_df, FACTORY_DUE_CANDIDATES),
                    "ship": get_first_matching_column(import_df, SHIP_DATE_CANDIDATES),
                    "remark": get_first_matching_column(import_df, REMARK_CANDIDATES),
                    "tags": get_first_matching_column(import_df, CUSTOMER_TAG_CANDIDATES),
                }

                st.write("Detected import columns:")
                st.json(import_cols)

                deduped_import_df, duplicate_keys_in_file = dedupe_import_df_by_key(
                    import_df,
                    import_cols.get("po"),
                    import_cols.get("part"),
                    import_cols.get("qty"),
                    import_cols.get("factory_due"),
                )
                if duplicate_keys_in_file:
                    st.warning(f"同一批匯入檔案中有重複 key，已自動去重。重複筆數：{len(duplicate_keys_in_file)}")

                if st.button("Batch Update from File", key="batch_update_from_file"):
                    if not import_cols.get("wip"):
                        st.error("匯入檔至少要能辨識出 WIP 欄位。")
                        st.stop()

                    ok_update_count = 0
                    manual_review_count = 0
                    skip_count = 0
                    fail_count = 0
                    logs = []

                    working_orders = orders.copy()
                    manual_review_items = []

                    for _, row in deduped_import_df.iterrows():
                        result = classify_and_update_factory_row(
                            current_df=working_orders,
                            teable_cols=detected_cols,
                            import_row=row,
                            import_cols=import_cols,
                        )

                        po_value = safe_text(row.get(import_cols.get("po", ""), "")) if import_cols.get("po") else ""
                        part_value = safe_text(row.get(import_cols.get("part", ""), "")) if import_cols.get("part") else ""
                        qty_value = safe_text(row.get(import_cols.get("qty", ""), "")) if import_cols.get("qty") else ""
                        due_value = safe_text(row.get(import_cols.get("factory_due", ""), "")) if import_cols.get("factory_due") else ""

                        if result["success"] and result["action"] == "UPDATED":
                            ok_update_count += 1
                            logs.append(
                                f"[UPDATED] {po_value} | {part_value} | {qty_value} | {due_value} | matched {','.join(result['match_info'].get('matched_fields', []))}"
                            )
                            if result.get("record_id"):
                                working_orders = update_working_orders_local(
                                    working_orders,
                                    result["record_id"],
                                    result.get("payload_fields", {}),
                                )

                        elif result["action"] == "MANUAL_REVIEW":
                            manual_review_count += 1
                            item = build_manual_review_item(
                                import_row=row,
                                import_cols=import_cols,
                                match_info=result.get("match_info", {}),
                                reason=result.get("message", "需要人工確認"),
                                teable_cols=detected_cols,
                            )
                            manual_review_items.append(item)
                            logs.append(f"[MANUAL REVIEW] {po_value} | {part_value} | {qty_value} | {due_value} -> {result.get('message', '')}")

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

                    st.success("Batch import finished.")
                    if logs:
                        st.text("\n".join(logs[:200]))

                    if manual_review_items:
                        st.warning("以下資料未自動覆蓋，已列入待人工確認。")
                        review_df = pd.DataFrame([
                            {
                                "PO#": x["PO#"],
                                "Part No": x["Part No"],
                                "Qty": x["Qty"],
                                "Factory Due Date": x["Factory Due Date"],
                                "New WIP": x["New WIP"],
                                "Reason": x["Reason"],
                                "Best Score": x["Best Score"],
                                "Matched Fields": x["Matched Fields"],
                            }
                            for x in manual_review_items
                        ])
                        st.dataframe(review_df, use_container_width=True, height=260)

                    refresh_after_update()
            except Exception as e:
                st.error(f"Import failed: {e}")

    with tab2:
        render_manual_update_tab(orders, detected_cols)

    with tab3:
        st.code(
            "PO78310 | Shipping | 2026-03-20 | Partial Shipment, Shipped | ready to ship\n"
            "PO78311 | On Hold |  | On Hold | waiting customer reply"
        )
        quick_text = st.text_area("Paste Quick Text", height=220)

        if st.button("Batch Update from Quick Text", key="batch_quick_text"):
            lines = [x.strip() for x in quick_text.splitlines() if x.strip()]
            ok_count = 0
            fail_count = 0
            logs = []

            for line in lines:
                parsed = parse_quick_text_line(line)
                if not parsed:
                    continue

                updates = {}
                if detected_cols.get("wip") and parsed.get("wip"):
                    updates[detected_cols["wip"]] = normalize_wip_value(parsed["wip"])
                if detected_cols.get("ship") and parsed.get("ship_date"):
                    updates[detected_cols["ship"]] = parsed["ship_date"]
                if detected_cols.get("tags"):
                    updates[detected_cols["tags"]] = build_tags_value(parsed.get("tags", []), MULTI_SELECT_MODE)
                if detected_cols.get("remark") and parsed.get("remark"):
                    updates[detected_cols["remark"]] = parsed["remark"]

                success, msg = upsert_to_teable(
                    current_df=orders,
                    po_col_name=detected_cols["po"],
                    po_value=parsed["po"],
                    updates=updates,
                )
                if success:
                    ok_count += 1
                else:
                    fail_count += 1
                    logs.append(f"{parsed['po']} -> {msg}")

            st.success(f"Quick text update finished. Success: {ok_count}, Failed: {fail_count}")
            if logs:
                st.text("\n".join(logs[:50]))
            refresh_after_update()

    with tab4:
        st.markdown("貼上 Email 文字，系統會自動解析出多筆 PO / WIP 資料。")
        email_text = st.text_area("Paste Email Text", height=240)

        if st.button("Parse Email Text", key="parse_email_text"):
            try:
                parsed_rows = parse_email_text_to_rows(email_text)
                parsed_df = pd.DataFrame(parsed_rows)
                if parsed_df.empty:
                    st.warning("沒有解析到可用資料。")
                else:
                    st.dataframe(parsed_df, use_container_width=True, height=260)
            except Exception as e:
                st.error(f"Email text parse failed: {e}")

    with tab5:
        st.markdown("上傳進度截圖，OCR 辨識後確認再更新 Teable。")
        uploaded_img = st.file_uploader("Upload PNG / JPG / JPEG", type=["png", "jpg", "jpeg"], key="ocr_uploader")

        if uploaded_img is not None:
            try:
                image = Image.open(uploaded_img)
                st.image(image, caption="Uploaded Image", use_container_width=True)

                ocr_text = ocr_image_to_text(image)
                st.text_area("OCR Raw Text", value=ocr_text, height=220)

                if str(ocr_text).startswith("OCR_ERROR:"):
                    st.error(ocr_text)
                else:
                    guessed_po = extract_po_from_text(ocr_text)
                    guessed_wip = infer_wip_from_text(ocr_text)
                    guessed_date = extract_date_from_text(ocr_text)
                    guessed_tags = infer_customer_tags_from_text(ocr_text)
                    guessed_remark = infer_remark_from_text(ocr_text)

                    with st.form("ocr_update_form"):
                        po_input = st.text_input("PO#", value=guessed_po)
                        wip_input = st.text_input("WIP", value=guessed_wip)
                        ship_input = st.text_input("Ship Date", value=guessed_date)
                        tags_input = st.multiselect(
                            "Customer Remark Tags",
                            TAG_OPTIONS,
                            default=[t for t in guessed_tags if t in TAG_OPTIONS],
                        )
                        remark_input = st.text_area("Remark", value=guessed_remark, height=120)
                        submitted_ocr = st.form_submit_button("Update to Teable")

                    if submitted_ocr:
                        if not po_input.strip():
                            st.error("PO# is required")
                        else:
                            updates = {}
                            if detected_cols.get("wip") and wip_input.strip():
                                updates[detected_cols["wip"]] = normalize_wip_value(wip_input.strip())
                            if detected_cols.get("ship") and ship_input.strip():
                                updates[detected_cols["ship"]] = ship_input.strip()
                            if detected_cols.get("tags"):
                                updates[detected_cols["tags"]] = build_tags_value(tags_input, MULTI_SELECT_MODE)
                            if detected_cols.get("remark") and remark_input.strip():
                                updates[detected_cols["remark"]] = remark_input.strip()

                            success, msg = upsert_to_teable(
                                current_df=orders,
                                po_col_name=detected_cols["po"],
                                po_value=po_input.strip(),
                                updates=updates,
                            )
                            if success:
                                st.success(f"{po_input.strip()} updated successfully from OCR")
                                refresh_after_update()
                            else:
                                st.error(msg)
            except Exception as e:
                st.error(f"Image OCR failed: {e}")


orders, api_status, api_text = load_orders()

if orders.empty:
    st.title("🏭 GLOCOM Control Tower")
    show_no_data_layout()
    st.stop()

detected_cols = {
    "po": get_first_matching_column(orders, PO_CANDIDATES),
    "customer": get_first_matching_column(orders, CUSTOMER_CANDIDATES),
    "part": get_first_matching_column(orders, PART_CANDIDATES),
    "qty": get_first_matching_column(orders, QTY_CANDIDATES),
    "factory": get_first_matching_column(orders, FACTORY_CANDIDATES),
    "wip": get_first_matching_column(orders, WIP_CANDIDATES),
    "factory_due": get_first_matching_column(orders, FACTORY_DUE_CANDIDATES),
    "ship": get_first_matching_column(orders, SHIP_DATE_CANDIDATES),
    "remark": get_first_matching_column(orders, REMARK_CANDIDATES),
    "tags": get_first_matching_column(orders, CUSTOMER_TAG_CANDIDATES),
    "merge_date": get_first_matching_column(orders, MERGE_DATE_CANDIDATES),
    "order_date": get_first_matching_column(orders, ORDER_DATE_CANDIDATES),
    "factory_order_date": get_first_matching_column(orders, FACTORY_ORDER_DATE_CANDIDATES),
    "changed_due_date": get_first_matching_column(orders, CHANGED_DUE_DATE_CANDIDATES),
}

query = st.query_params
customer_param = query.get("customer", None)

if customer_param:
    st.title("GLOCOM Order Status")
    st.caption("Customer WIP Progress")
    show_customer_preview_report(
        orders=orders,
        detected_cols=detected_cols,
        selected_customer=str(customer_param),
        portal_mode=True,
    )
    st.stop()

st.title("🏭 GLOCOM Control Tower")
st.caption("Internal PCB Production Monitoring System")

with st.expander("Debug"):
    st.write("API Status:", api_status)
    st.write("TABLE_URL:", TABLE_URL)
    st.write("Token loaded:", bool(TEABLE_TOKEN))
    st.write("Columns:", list(orders.columns) if not orders.empty else [])
    if isinstance(api_text, str):
        st.text(api_text[:1200])

st.sidebar.title("GLOCOM Internal")
st.sidebar.link_button("Open Teable", TEABLE_WEB_URL, use_container_width=True)

menu = st.sidebar.radio(
    "功能選單",
    [
        "Dashboard",
        "Factory Load",
        "Delayed Orders",
        "Shipment Forecast",
        "Orders",
        "Customer Preview",
        "Sandy 內部 WIP",
        "Sandy 銷貨底",
        "新訂單 WIP",
        "Import / Update",
    ],
)

if st.sidebar.button("Refresh"):
    refresh_after_update()

st.sidebar.markdown("---")
st.sidebar.caption("完成案件請在 Teable 主 View 設定篩選：WIP ≠ 完成")
st.sidebar.caption("另建 Completed View：WIP = 完成")

if menu == "Dashboard":
    show_dashboard_report(orders, detected_cols)
elif menu == "Factory Load":
    show_factory_load_report(orders, detected_cols)
elif menu == "Delayed Orders":
    show_delayed_orders_report(orders, detected_cols)
elif menu == "Shipment Forecast":
    show_shipment_forecast_report(orders, detected_cols)
elif menu == "Orders":
    show_orders_report(orders, detected_cols)
elif menu == "Customer Preview":
    show_customer_preview_report(orders, detected_cols)
elif menu == "Sandy 內部 WIP":
    show_sandy_internal_wip_report(orders, detected_cols)
elif menu == "Sandy 銷貨底":
    show_sandy_shipment_report(orders, detected_cols)
elif menu == "新訂單 WIP":
    show_new_orders_wip_report(orders, detected_cols)
elif menu == "Import / Update":
    render_import_update(orders, detected_cols)

st.caption("Auto refresh cache: 60 seconds")
