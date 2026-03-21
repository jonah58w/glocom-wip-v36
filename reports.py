# -*- coding: utf-8 -*-
from io import BytesIO

import pandas as pd
import streamlit as st

import utils as u
import config as cfg


# ================================
# Common helpers
# ================================
def _existing_cols(df, cols):
    return [c for c in cols if c and c in df.columns]


def _safe_series(df, col_name):
    if col_name and col_name in df.columns:
        s = u.get_series_by_col(df, col_name)
        if s is not None:
            return s
    return pd.Series([""] * len(df), index=df.index)


def _safe_text_series(df, col_name):
    s = _safe_series(df, col_name)
    return s.astype(str).fillna("").str.strip()


def _safe_dt_series(df, col_name):
    if not col_name or col_name not in df.columns:
        return pd.Series([pd.NaT] * len(df), index=df.index)
    return pd.to_datetime(df[col_name], errors="coerce")


def _download_excel(df, file_name, sheet_name="Report"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    st.download_button(
        "Download Excel",
        data=output.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _download_pdf(df, file_name, title="Report"):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet

        pdf_buffer = BytesIO()
        doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(A4), leftMargin=18, rightMargin=18, topMargin=18, bottomMargin=18)
        styles = getSampleStyleSheet()

        safe_df = df.fillna("").astype(str).copy()
        table_data = [list(safe_df.columns)] + safe_df.values.tolist()

        table = Table(table_data, repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#333333")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                    ("FONTSIZE", (0, 0), (-1, -1), 7),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5f5")]),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ]
            )
        )

        story = [
            Paragraph(title, styles["Title"]),
            Spacer(1, 8),
            table,
        ]
        doc.build(story)

        st.download_button(
            "Download PDF",
            data=pdf_buffer.getvalue(),
            file_name=file_name,
            mime="application/pdf",
        )
    except Exception as e:
        st.error(f"PDF 匯出失敗: {e}")


def _show_export_buttons(df, excel_name, pdf_name, title):
    c1, c2 = st.columns(2)
    with c1:
        _download_excel(df, excel_name, sheet_name=title)
    with c2:
        _download_pdf(df, pdf_name, title=title)


# ================================
# Dashboard
# ================================
def show_dashboard_report(df, wip_col=None, factory_col=None, ship_date_col=None, **kwargs):
    total_orders = len(df)
    shipping = 0

    if wip_col and wip_col in df.columns:
        wip_series = _safe_text_series(df, wip_col)
        shipping = len(df[wip_series.str.contains("ship|shipping|出貨", case=False, na=False)])

    production = total_orders - shipping

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Orders", total_orders)
    c2.metric("Production", production)
    c3.metric("Shipping", shipping)

    st.divider()

    left, right = st.columns(2)
    with left:
        show_factory_load_report(df, factory_col=factory_col)
    with right:
        show_shipment_forecast_report(df, ship_date_col=ship_date_col, po_col=kwargs.get("po_col"), customer_col=kwargs.get("customer_col"), part_col=kwargs.get("part_col"), qty_col=kwargs.get("qty_col"), factory_col=factory_col, wip_col=wip_col)


# ================================
# Factory Load
# ================================
def show_factory_load_report(df, factory_col=None, **kwargs):
    st.subheader("🏭 Factory Load")

    if not factory_col or factory_col not in df.columns:
        st.info("No factory column found")
        return

    factory_series = _safe_text_series(df, factory_col)
    if factory_series is None:
        st.info("No factory data")
        return

    factory_summary = (
        factory_series.replace("", "(blank)")
        .value_counts()
        .reset_index()
    )
    factory_summary.columns = [factory_col, "Orders"]

    st.bar_chart(factory_summary.set_index(factory_col))
    st.dataframe(factory_summary, use_container_width=True, height=400)


# ================================
# Delayed Orders
# ================================
def show_delayed_orders_report(df, factory_due_col=None, po_col=None, customer_col=None, part_col=None, qty_col=None, factory_col=None, wip_col=None, **kwargs):
    st.subheader("⚠️ Delayed Orders")

    if not factory_due_col or factory_due_col not in df.columns:
        st.info("No factory due date column")
        return

    temp = df.copy()
    temp["_FactoryDueDateParsed"] = pd.to_datetime(temp[factory_due_col], errors="coerce")
    today = pd.Timestamp.today().normalize()

    delayed = temp[
        temp["_FactoryDueDateParsed"].notna()
        & (temp["_FactoryDueDateParsed"] < today)
    ].copy()

    if delayed.empty:
        st.success("No delayed orders")
        return

    delayed["Delay Days"] = (today - delayed["_FactoryDueDateParsed"]).dt.days
    show_cols = _existing_cols(delayed, [po_col, customer_col, part_col, qty_col, factory_col, wip_col, factory_due_col, "Delay Days"])
    st.dataframe(delayed[show_cols], use_container_width=True, height=520)


# ================================
# Shipment Forecast
# ================================
def show_shipment_forecast_report(df, ship_date_col=None, po_col=None, customer_col=None, part_col=None, qty_col=None, factory_col=None, wip_col=None, **kwargs):
    st.subheader("📦 Shipment Forecast (Next 7 days)")

    if not ship_date_col or ship_date_col not in df.columns:
        st.info("No ship date column")
        return

    temp = df.copy()
    temp["_ShipDateParsed"] = pd.to_datetime(temp[ship_date_col], errors="coerce")
    today = pd.Timestamp.today().normalize()
    next_7 = today + pd.Timedelta(days=7)

    forecast = temp[
        temp["_ShipDateParsed"].notna()
        & (temp["_ShipDateParsed"] >= today)
        & (temp["_ShipDateParsed"] <= next_7)
    ].copy()

    if forecast.empty:
        st.info("No shipment within next 7 days")
        return

    show_cols = _existing_cols(forecast, [po_col, customer_col, part_col, qty_col, factory_col, wip_col, ship_date_col])
    st.dataframe(forecast.sort_values("_ShipDateParsed")[show_cols], use_container_width=True, height=520)


# ================================
# Orders
# ================================
def show_orders_report(df, customer_col=None, wip_col=None, **kwargs):
    st.subheader("📋 Orders")

    filtered = df.copy()
    c1, c2 = st.columns(2)

    if customer_col and customer_col in filtered.columns:
        customer_series = _safe_text_series(filtered, customer_col)
        customer_options = ["All"] + sorted([x for x in customer_series.unique().tolist() if x])
        selected_customer = c1.selectbox("Customer", customer_options)
        if selected_customer != "All":
            filtered = filtered[_safe_text_series(filtered, customer_col) == selected_customer]

    if wip_col and wip_col in filtered.columns:
        wip_series = _safe_text_series(filtered, wip_col)
        wip_options = ["All"] + sorted([x for x in wip_series.unique().tolist() if x])
        selected_wip = c2.selectbox("WIP Stage", wip_options)
        if selected_wip != "All":
            filtered = filtered[_safe_text_series(filtered, wip_col) == selected_wip]

    st.dataframe(filtered, use_container_width=True, height=520)
    csv_data = filtered.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Download Orders CSV", data=csv_data, file_name="glocom_orders.csv", mime="text/csv")


# ================================
# Customer Preview
# ================================
def show_customer_preview_report(df, customer_col=None, po_col=None, part_col=None, qty_col=None, wip_col=None, ship_date_col=None, customer_tag_col=None, remark_col=None, **kwargs):
    st.subheader("Customer Preview")
    st.caption("僅供內部預覽。客戶請直接使用 Teable View。")

    if not customer_col or customer_col not in df.columns:
        st.error("Customer column not found in Teable data")
        return

    customer_series = _safe_text_series(df, customer_col)
    customers = sorted([x for x in customer_series.unique().tolist() if x])

    if not customers:
        st.warning("No customers found")
        return

    selected_customer = st.selectbox("Select customer to preview", customers)
    preview_df = df[customer_series.str.lower() == selected_customer.strip().lower()].copy()

    if preview_df.empty:
        st.warning("No orders found for this customer")
        return

    preview_cols = _existing_cols(preview_df, [po_col, customer_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col])
    st.dataframe(preview_df[preview_cols], use_container_width=True, height=420)


# ================================
# Sandy 內部 WIP
# ================================
def show_sandy_internal_wip_report(
    df,
    po_col=None,
    customer_col=None,
    part_col=None,
    qty_col=None,
    ship_date_col=None,
    wip_col=None,
    factory_due_col=None,
    changed_due_date_col=None,
    merge_date_col=None,
    factory_col=None,
    remark_col=None,
    **kwargs,
):
    st.subheader("Sandy 內部 WIP")
    st.caption("依併貨日期排序，可匯出 Excel / PDF。")

    work = df.copy()

    if merge_date_col and merge_date_col in work.columns:
        work["_merge_sort"] = pd.to_datetime(work[merge_date_col], errors="coerce")
        work = work.sort_values("_merge_sort", ascending=True)

    show_cols = _existing_cols(
        work,
        [
            customer_col,
            po_col,
            part_col,
            qty_col,
            ship_date_col,
            wip_col,
            factory_due_col,
            changed_due_date_col,
            merge_date_col,
            factory_col,
            remark_col,
        ],
    )

    if not show_cols:
        st.warning("沒有可顯示欄位。")
        return

    export_df = work[show_cols].copy()
    st.dataframe(export_df, use_container_width=True, height=520)
    _show_export_buttons(export_df, "Sandy_內部WIP.xlsx", "Sandy_內部WIP.pdf", "Sandy 內部 WIP")


# ================================
# Sandy 銷貨底
# ================================
def show_sandy_shipment_report(
    df,
    po_col=None,
    customer_col=None,
    part_col=None,
    qty_col=None,
    ship_date_col=None,
    wip_col=None,
    factory_due_col=None,
    changed_due_date_col=None,
    merge_date_col=None,
    factory_col=None,
    remark_col=None,
    **kwargs,
):
    st.subheader("Sandy 銷貨底")
    st.caption("只顯示 SHIPMENT / Shipping / 出貨，依併貨日期排序，可匯出 Excel / PDF。")

    work = df.copy()

    if wip_col and wip_col in work.columns:
        wip_s = _safe_text_series(work, wip_col)
        work = work[wip_s.str.contains("SHIPMENT|Shipping|出貨", case=False, na=False)].copy()

    if merge_date_col and merge_date_col in work.columns:
        work["_merge_sort"] = pd.to_datetime(work[merge_date_col], errors="coerce")
        work = work.sort_values("_merge_sort", ascending=True)

    show_cols = _existing_cols(
        work,
        [
            customer_col,
            po_col,
            part_col,
            qty_col,
            ship_date_col,
            wip_col,
            factory_due_col,
            changed_due_date_col,
            merge_date_col,
            factory_col,
            remark_col,
        ],
    )

    if work.empty:
        st.warning("目前沒有 SHIPMENT / Shipping / 出貨 資料。")
        return

    export_df = work[show_cols].copy()
    st.dataframe(export_df, use_container_width=True, height=520)
    _show_export_buttons(export_df, "Sandy_銷貨底.xlsx", "Sandy_銷貨底.pdf", "Sandy 銷貨底")


# ================================
# 新訂單 WIP
# 改成：抓取任何 工廠交期 / 交期更改 / 併貨日期 有變更者
# ================================
def show_new_orders_wip_report(
    df,
    po_col=None,
    customer_col=None,
    part_col=None,
    qty_col=None,
    wip_col=None,
    ship_date_col=None,
    factory_due_col=None,
    remark_col=None,
    merge_date_col=None,
    order_date_col=None,
    factory_order_date_col=None,
    changed_due_date_col=None,
    **kwargs,
):
    st.subheader("新訂單 WIP")
    st.caption("顯示任何 工廠交期 / 交期更改 / 併貨日期 有異動者，可匯出 Excel / PDF。")

    work = df.copy()

    due_s = _safe_text_series(work, factory_due_col)
    changed_due_s = _safe_text_series(work, changed_due_date_col)
    merge_s = _safe_text_series(work, merge_date_col)

    # 任一異動條件成立就顯示
    cond_changed_due_has_value = changed_due_s != ""
    cond_merge_has_value = merge_s != ""
    cond_due_diff = (
        (due_s != "")
        & (changed_due_s != "")
        & (due_s != changed_due_s)
    )

    filtered = work[
        cond_changed_due_has_value
        | cond_merge_has_value
        | cond_due_diff
    ].copy()

    if merge_date_col and merge_date_col in filtered.columns:
        filtered["_merge_sort"] = pd.to_datetime(filtered[merge_date_col], errors="coerce")
    else:
        filtered["_merge_sort"] = pd.NaT

    if changed_due_date_col and changed_due_date_col in filtered.columns:
        filtered["_changed_due_sort"] = pd.to_datetime(filtered[changed_due_date_col], errors="coerce")
    else:
        filtered["_changed_due_sort"] = pd.NaT

    if factory_due_col and factory_due_col in filtered.columns:
        filtered["_factory_due_sort"] = pd.to_datetime(filtered[factory_due_col], errors="coerce")
    else:
        filtered["_factory_due_sort"] = pd.NaT

    filtered = filtered.sort_values(
        ["_merge_sort", "_changed_due_sort", "_factory_due_sort"],
        ascending=True,
        na_position="last",
    )

    show_cols = _existing_cols(
        filtered,
        [
            order_date_col,
            factory_order_date_col,
            customer_col,
            po_col,
            part_col,
            qty_col,
            ship_date_col,
            wip_col,
            factory_due_col,
            changed_due_date_col,
            merge_date_col,
            remark_col,
        ],
    )

    if filtered.empty:
        st.warning("目前沒有符合條件的異動訂單。")
        return

    export_df = filtered[show_cols].copy()
    st.dataframe(export_df, use_container_width=True, height=520)
    _show_export_buttons(export_df, "新訂單WIP.xlsx", "新訂單WIP.pdf", "新訂單 WIP")
