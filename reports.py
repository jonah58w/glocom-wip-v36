import io
import re

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

from utils import safe_text, split_tags, get_series_by_col, safe_to_datetime, get_first_matching_column


def wip_display_html(value: str) -> str:
    text = safe_text(value)
    lower = text.lower()
    if any(k in lower for k in ["完成"]) or text.upper() in {"DONE", "COMPLETE", "COMPLETED", "FINISHED", "FINISH"}:
        label, bg, fg = text or "完成", "#065f46", "#d1fae5"
    elif any(k in lower for k in ["ship", "shipping", "shipment", "出貨"]):
        label, bg, fg = text or "Shipping", "#14532d", "#dcfce7"
    elif any(k in lower for k in ["pack", "包裝"]):
        label, bg, fg = text or "Packing", "#166534", "#dcfce7"
    elif any(k in lower for k in ["fqc", "qa", "inspection", "成檢", "測試"]):
        label, bg, fg = text or "Inspection", "#854d0e", "#fef3c7"
    elif any(k in lower for k in ["aoi", "drill", "route", "routing", "plating", "inner", "production", "防焊", "壓合", "外層", "內層", "成型"]):
        label, bg, fg = text or "Production", "#9a3412", "#ffedd5"
    elif any(k in lower for k in ["eng", "gerber", "cam", "eq"]):
        label, bg, fg = text or "Engineering", "#1d4ed8", "#dbeafe"
    elif any(k in lower for k in ["hold", "等待", "暫停"]):
        label, bg, fg = text or "On Hold", "#7f1d1d", "#fee2e2"
    else:
        label, bg, fg = text or "-", "#374151", "#f3f4f6"
    return f'<span class="wip-chip" style="background:{bg};color:{fg};">{label}</span>'


def show_metrics(df: pd.DataFrame, wip_col: str | None):
    total_orders = len(df)
    shipping = 0
    if wip_col:
        wip_series = get_series_by_col(df, wip_col)
        if wip_series is not None:
            shipping = len(df[wip_series.astype(str).str.contains("ship|shipping|shipment|出貨", case=False, na=False)])
    production = total_orders - shipping
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Orders", total_orders)
    c2.metric("Production", production)
    c3.metric("Shipping", shipping)


def show_no_data_layout():
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Orders", 0)
    c2.metric("Production", 0)
    c3.metric("Shipping", 0)
    st.divider()
    st.warning("No data from Teable")


def show_factory_load(df: pd.DataFrame, f_col: str | None):
    st.subheader("🏭 Factory Load")
    if f_col:
        factory_series = get_series_by_col(df, f_col)
        if factory_series is not None:
            factory_summary = factory_series.fillna("(blank)").astype(str).value_counts().reset_index()
            factory_summary.columns = [f_col, "Orders"]
            st.bar_chart(factory_summary.set_index(f_col))
            st.dataframe(factory_summary, use_container_width=True, height=400)
            return
    st.info("No factory column found")


def show_delayed_orders(df: pd.DataFrame, factory_due_col, po_col, customer_col, part_col, qty_col, factory_col, wip_col):
    st.subheader("⚠️ Delayed Orders")
    if not factory_due_col:
        st.info("No factory due date column")
        return
    temp = df.copy()
    due_series = get_series_by_col(temp, factory_due_col)
    if due_series is None:
        st.info("No factory due date data")
        return
    temp["_FactoryDueDateParsed"] = safe_to_datetime(due_series)
    today = pd.Timestamp.today().normalize()
    delayed = temp[temp["_FactoryDueDateParsed"].notna() & (temp["_FactoryDueDateParsed"] < today)].copy()
    if delayed.empty:
        st.success("No delayed orders")
        return
    delayed["Delay Days"] = (today - delayed["_FactoryDueDateParsed"]).dt.days
    show_cols = [c for c in [po_col, customer_col, part_col, qty_col, factory_col, wip_col, factory_due_col] if c and c in delayed.columns]
    if "Delay Days" not in show_cols:
        show_cols.append("Delay Days")
    st.dataframe(delayed[show_cols], use_container_width=True, height=520)


def show_shipment_forecast(df: pd.DataFrame, ship_date_col, po_col, customer_col, part_col, qty_col, factory_col, wip_col):
    st.subheader("📦 Shipment Forecast (Next 7 days)")
    if not ship_date_col:
        st.info("No ship date column")
        return
    temp = df.copy()
    ship_series = get_series_by_col(temp, ship_date_col)
    if ship_series is None:
        st.info("No ship date data")
        return
    temp["_ShipDateParsed"] = safe_to_datetime(ship_series)
    today = pd.Timestamp.today().normalize()
    next_7 = today + pd.Timedelta(days=7)
    forecast = temp[temp["_ShipDateParsed"].notna() & (temp["_ShipDateParsed"] >= today) & (temp["_ShipDateParsed"] <= next_7)].copy()
    if forecast.empty:
        st.info("No shipment within next 7 days")
        return
    show_cols = [c for c in [po_col, customer_col, part_col, qty_col, factory_col, wip_col, ship_date_col] if c and c in forecast.columns]
    st.dataframe(forecast.sort_values("_ShipDateParsed")[show_cols], use_container_width=True, height=520)


def show_orders_table(df: pd.DataFrame):
    st.subheader("📋 Orders")
    st.dataframe(df, use_container_width=True, height=520)
    csv_data = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Download Orders CSV", data=csv_data, file_name="glocom_orders.csv", mime="text/csv")


def render_customer_portal(cust_orders, po_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col):
    for _, row in cust_orders.iterrows():
        po_val = safe_text(row.get(po_col, "")) if po_col else ""
        part_val = safe_text(row.get(part_col, "")) if part_col else ""
        qty_val = safe_text(row.get(qty_col, "")) if qty_col else ""
        wip_val = safe_text(row.get(wip_col, "")) if wip_col else ""
        ship_val = safe_text(row.get(ship_date_col, "")) if ship_date_col else ""
        remark_val = safe_text(row.get(remark_col, "")) if remark_col else ""
        tags_val = split_tags(row.get(customer_tag_col, "")) if customer_tag_col else []
        tag_html = "".join([f'<span class="tag-chip">{t}</span>' for t in tags_val]) if tags_val else '<span class="tag-chip">-</span>'
        st.markdown(f"""
        <div class="portal-box">
            <div class="portal-title">{po_val or '-'}</div>
            <div style="display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;">
                <div>
                    <div><strong>P/N</strong> : {part_val or '-'}</div>
                    <div><strong>Qty</strong> : {qty_val or '-'}</div>
                </div>
                <div>
                    <div><strong>WIP</strong> : {wip_display_html(wip_val)}</div>
                    <div style="margin-top:8px;"><strong>Ship Date</strong> : {ship_val or '-'}</div>
                </div>
            </div>
            <div style="margin-top:12px;"><strong>Customer Remark Tags</strong> : {tag_html}</div>
            <div style="margin-top:10px;"><strong>Remark</strong> : {remark_val or '-'}</div>
        </div>
        """, unsafe_allow_html=True)


MERGE_DATE_CANDIDATES = ["併貨日期(限內部使用)", "併貨日期\n(限內部使用)", "併貨日期", "Merge Date"]
CUSTOMER_ORDER_DATE_CANDIDATES = ["客戶下單日期", "客戶下\n單日期"]
FACTORY_ORDER_DATE_CANDIDATES = ["工廠下單日期", "工廠下\n單日期"]
SHIPMENT_STATUS_CANDIDATES = ["已出貨+C:CＶ", "已出貨+C:C\nＶ", "SHIPMENT"]


def _clean_col_name(name: str) -> str:
    return re.sub(r"\s+", "", safe_text(name))


def find_column_fuzzy(df: pd.DataFrame, candidates):
    if df is None or df.empty:
        return None
    direct = get_first_matching_column(df, candidates)
    if direct:
        return direct
    clean_map = {_clean_col_name(c): c for c in df.columns}
    for cand in candidates:
        key = _clean_col_name(cand)
        if key in clean_map:
            return clean_map[key]
    for col in df.columns:
        ckey = _clean_col_name(col)
        for cand in candidates:
            ckey2 = _clean_col_name(cand)
            if ckey2 == ckey or ckey2 in ckey or ckey in ckey2:
                return col
    return None


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Report") -> bytes:
    output = io.BytesIO()
    export_df = df.copy()
    for col in export_df.columns:
        if pd.api.types.is_datetime64_any_dtype(export_df[col]):
            export_df[col] = export_df[col].dt.strftime("%Y-%m-%d")
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        ws = writer.book[sheet_name[:31]]
        from openpyxl.utils import get_column_letter
        for idx, column in enumerate(export_df.columns, start=1):
            max_len = max([len(str(column))] + [len(safe_text(x)) for x in export_df[column].head(200)])
            ws.column_dimensions[get_column_letter(idx)].width = min(max(12, max_len + 2), 40)
    return output.getvalue()


def dataframe_to_pdf_bytes(df: pd.DataFrame, title: str) -> bytes:
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), leftMargin=18, rightMargin=18, topMargin=18, bottomMargin=18)
    styles = getSampleStyleSheet()
    story = [Paragraph(title, styles["Title"]), Spacer(1, 8)]
    export_df = df.copy().fillna("")
    max_rows = min(len(export_df), 200)
    export_df = export_df.head(max_rows)
    for col in export_df.columns:
        export_df[col] = export_df[col].map(lambda x: safe_text(x))
    headers = [safe_text(c) for c in export_df.columns]
    rows = [headers] + export_df.values.tolist()
    rows = [[str(x)[:60] for x in row] for row in rows]
    table = Table(rows, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9e2f3")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f7f7f7")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(table)
    if len(df) > max_rows:
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"Only first {max_rows} rows exported to PDF.", styles["Italic"]))
    doc.build(story)
    return output.getvalue()


def _prepare_report_export_ui(title: str, report_df: pd.DataFrame, excel_name: str, pdf_name: str):
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "Download Excel",
            data=dataframe_to_excel_bytes(report_df, sheet_name=title[:31]),
            file_name=excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "Download PDF",
            data=dataframe_to_pdf_bytes(report_df, title),
            file_name=pdf_name,
            mime="application/pdf",
            use_container_width=True,
        )


def _sort_by_date_col(df: pd.DataFrame, date_col: str | None, ascending=True):
    out = df.copy()
    if date_col and date_col in out.columns:
        out["_sort_date"] = pd.to_datetime(out[date_col], errors="coerce")
        out = out.sort_values("_sort_date", ascending=ascending, na_position="last").drop(columns=["_sort_date"])
    return out


def _selected_existing_columns(df: pd.DataFrame, cols):
    return [c for c in cols if c and c in df.columns]


def show_sandy_internal_wip_report(df: pd.DataFrame):
    st.subheader("Sandy 內部 WIP")
    merge_col = find_column_fuzzy(df, MERGE_DATE_CANDIDATES)
    ship_col = find_column_fuzzy(df, ["Ship date", "Ship Date", "出貨日期", "出貨日期(公式)"])
    due_change_col = find_column_fuzzy(df, ["交期(更改)"])
    note_col = find_column_fuzzy(df, ["Note", "情況", "備註", "工廠提醒事項"])
    required_cols = [
        find_column_fuzzy(df, CUSTOMER_ORDER_DATE_CANDIDATES),
        find_column_fuzzy(df, FACTORY_ORDER_DATE_CANDIDATES),
        find_column_fuzzy(df, ["Customer", "客戶"]),
        find_column_fuzzy(df, ["PO#", "PO"]),
        find_column_fuzzy(df, ["P/N", "Part No"]),
        find_column_fuzzy(df, ["Order Q'TY (PCS)", "Qty", "Order Q'TY\n(PCS)"]),
        find_column_fuzzy(df, ["Dock"]),
        ship_col,
        find_column_fuzzy(df, ["WIP"]),
        find_column_fuzzy(df, ["工廠交期", "Factory Due Date"]),
        due_change_col,
        find_column_fuzzy(df, ["工廠"]),
        merge_col,
        note_col,
        find_column_fuzzy(df, ["Ship to"]),
        find_column_fuzzy(df, ["Ship via"]),
        find_column_fuzzy(df, ["Tracking No."]),
    ]
    report_df = df[_selected_existing_columns(df, required_cols)].copy()
    show_only_open = st.toggle("Hide completed / shipped", value=True, key="sandy_internal_hide_done")
    if show_only_open:
        wip_col = find_column_fuzzy(report_df, ["WIP"])
        if wip_col:
            report_df = report_df[~report_df[wip_col].astype(str).str.contains("完成|shipment|shipping|出貨", case=False, na=False)]
    report_df = _sort_by_date_col(report_df, merge_col, ascending=True)
    st.caption("依併貨日期排序，可匯出 Excel / PDF。")
    st.dataframe(report_df, use_container_width=True, height=560)
    _prepare_report_export_ui("Sandy Internal WIP", report_df, "sandy_internal_wip.xlsx", "sandy_internal_wip.pdf")


def show_sandy_shipment_report(df: pd.DataFrame):
    st.subheader("Sandy 銷貨底")
    merge_col = find_column_fuzzy(df, MERGE_DATE_CANDIDATES)
    wip_col = find_column_fuzzy(df, ["WIP"])
    shipment_mark_col = find_column_fuzzy(df, SHIPMENT_STATUS_CANDIDATES)
    required_cols = [
        find_column_fuzzy(df, ["Customer", "客戶"]),
        find_column_fuzzy(df, ["PO#", "PO"]),
        find_column_fuzzy(df, ["P/N", "Part No"]),
        find_column_fuzzy(df, ["Order Q'TY (PCS)", "Qty", "Order Q'TY\n(PCS)"]),
        find_column_fuzzy(df, ["Dock"]),
        find_column_fuzzy(df, ["Ship date", "Ship Date", "出貨日期", "出貨日期(公式)"]),
        wip_col,
        find_column_fuzzy(df, ["工廠交期", "Factory Due Date"]),
        find_column_fuzzy(df, ["交期(更改)"]),
        find_column_fuzzy(df, ["工廠"]),
        merge_col,
        find_column_fuzzy(df, ["Ship to"]),
        find_column_fuzzy(df, ["Ship via"]),
        find_column_fuzzy(df, ["Tracking No."]),
        shipment_mark_col,
        find_column_fuzzy(df, ["Note", "情況", "備註"]),
    ]
    report_df = df[_selected_existing_columns(df, required_cols)].copy()
    mask = pd.Series([False] * len(report_df), index=report_df.index)
    if wip_col and wip_col in report_df.columns:
        mask = mask | report_df[wip_col].astype(str).str.contains("shipment|shipping|出貨", case=False, na=False)
    if shipment_mark_col and shipment_mark_col in report_df.columns:
        mask = mask | report_df[shipment_mark_col].astype(str).str.strip().ne("")
    report_df = report_df[mask].copy()
    report_df = _sort_by_date_col(report_df, merge_col, ascending=True)
    st.caption("依併貨日期排序，只顯示 SHIPMENT / 出貨相關資料，可匯出 Excel / PDF。")
    st.dataframe(report_df, use_container_width=True, height=560)
    _prepare_report_export_ui("Sandy Shipment Report", report_df, "sandy_shipment_report.xlsx", "sandy_shipment_report.pdf")


def show_new_orders_wip_report(df: pd.DataFrame):
    st.subheader("新訂單 WIP")
    customer_order_col = find_column_fuzzy(df, CUSTOMER_ORDER_DATE_CANDIDATES)
    merge_col = find_column_fuzzy(df, MERGE_DATE_CANDIDATES)
    today_default = pd.Timestamp.today().normalize().date()
    filter_date = st.date_input("客戶下單日期", value=today_default, key="new_order_filter_date")
    required_cols = [
        customer_order_col,
        find_column_fuzzy(df, FACTORY_ORDER_DATE_CANDIDATES),
        find_column_fuzzy(df, ["Customer", "客戶"]),
        find_column_fuzzy(df, ["PO#", "PO"]),
        find_column_fuzzy(df, ["P/N", "Part No"]),
        find_column_fuzzy(df, ["Order Q'TY (PCS)", "Qty", "Order Q'TY\n(PCS)"]),
        find_column_fuzzy(df, ["Dock"]),
        find_column_fuzzy(df, ["Ship date", "Ship Date"]),
        find_column_fuzzy(df, ["WIP"]),
        find_column_fuzzy(df, ["工廠交期", "Factory Due Date"]),
        find_column_fuzzy(df, ["交期(更改)"]),
        find_column_fuzzy(df, ["工廠"]),
        merge_col,
        find_column_fuzzy(df, ["工廠提醒事項"]),
        find_column_fuzzy(df, ["客戶要求注意事項"]),
        find_column_fuzzy(df, ["Working Gerber Approval"]),
        find_column_fuzzy(df, ["Engineering Question"]),
        find_column_fuzzy(df, ["Pricing & Qty issue"]),
        find_column_fuzzy(df, ["T/T"]),
        find_column_fuzzy(df, ["文件"]),
    ]
    report_df = df[_selected_existing_columns(df, required_cols)].copy()
    if customer_order_col and customer_order_col in report_df.columns:
        order_dates = pd.to_datetime(report_df[customer_order_col], errors="coerce").dt.date
        report_df = report_df[order_dates == filter_date].copy()
    change_due_col = find_column_fuzzy(report_df, ["交期(更改)"])
    if change_due_col and change_due_col in report_df.columns:
        report_df["工廠交期已變更"] = report_df[change_due_col].astype(str).str.strip().ne("")
    if merge_col and merge_col in report_df.columns:
        report_df["有併貨日"] = report_df[merge_col].astype(str).str.strip().ne("")
    report_df = _sort_by_date_col(report_df, merge_col, ascending=True)
    st.caption("顯示當天下單的訂單，並顯示工廠交期 / 交期更改 / 併貨日期，可匯出 Excel / PDF。")
    st.dataframe(report_df, use_container_width=True, height=560)
    _prepare_report_export_ui("New Orders WIP", report_df, "new_orders_wip.xlsx", "new_orders_wip.pdf")
