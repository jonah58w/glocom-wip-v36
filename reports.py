# -*- coding: utf-8 -*-
from io import BytesIO

import pandas as pd
import streamlit as st

import utils as u
import config as cfg


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


def _norm_name(s):
    s = str(s).strip()
    s = s.replace("\n", "")
    s = s.replace(" ", "")
    s = s.replace("（", "(").replace("）", ")")
    return s.lower()


def _find_col(df, candidates, fallback=None):
    if fallback and fallback in df.columns:
        return fallback

    actual_cols = list(df.columns)
    direct_map = {_norm_name(c): c for c in actual_cols}

    for c in candidates:
        if c in actual_cols:
            return c

    for c in candidates:
        nc = _norm_name(c)
        if nc in direct_map:
            return direct_map[nc]

    return None


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


def _show_export_buttons(df, excel_name, title):
    _download_excel(df, excel_name, sheet_name=title)


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
        show_shipment_forecast_report(
            df,
            ship_date_col=ship_date_col,
            po_col=kwargs.get("po_col"),
            customer_col=kwargs.get("customer_col"),
            part_col=kwargs.get("part_col"),
            qty_col=kwargs.get("qty_col"),
            factory_col=factory_col,
            wip_col=wip_col,
        )


def show_factory_load_report(df, factory_col=None, **kwargs):
    st.subheader("🏭 Factory Load")

    if not factory_col or factory_col not in df.columns:
        st.info("No factory column found")
        return

    factory_series = _safe_text_series(df, factory_col)
    factory_summary = factory_series.replace("", "(blank)").value_counts().reset_index()
    factory_summary.columns = [factory_col, "Orders"]

    st.bar_chart(factory_summary.set_index(factory_col))
    st.dataframe(factory_summary, use_container_width=True, height=400)


def show_delayed_orders_report(
    df,
    factory_due_col=None,
    po_col=None,
    customer_col=None,
    part_col=None,
    qty_col=None,
    factory_col=None,
    wip_col=None,
    **kwargs,
):
    st.subheader("⚠️ Delayed Orders")

    if not factory_due_col or factory_due_col not in df.columns:
        st.info("No factory due date column")
        return

    temp = df.copy()
    temp["_FactoryDueDateParsed"] = pd.to_datetime(temp[factory_due_col], errors="coerce")
    today = pd.Timestamp.today().normalize()

    delayed = temp[temp["_FactoryDueDateParsed"].notna() & (temp["_FactoryDueDateParsed"] < today)].copy()

    if delayed.empty:
        st.success("No delayed orders")
        return

    delayed["Delay Days"] = (today - delayed["_FactoryDueDateParsed"]).dt.days
    show_cols = _existing_cols(delayed, [po_col, customer_col, part_col, qty_col, factory_col, wip_col, factory_due_col, "Delay Days"])
    st.dataframe(delayed[show_cols], use_container_width=True, height=520)


def show_shipment_forecast_report(
    df,
    ship_date_col=None,
    po_col=None,
    customer_col=None,
    part_col=None,
    qty_col=None,
    factory_col=None,
    wip_col=None,
    **kwargs,
):
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


def show_customer_preview_report(
    df,
    customer_col=None,
    po_col=None,
    part_col=None,
    qty_col=None,
    wip_col=None,
    ship_date_col=None,
    customer_tag_col=None,
    remark_col=None,
    **kwargs,
):
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

    wesco_matches = [x for x in customers if "wesco" in str(x).strip().lower()]
    default_customer = wesco_matches[0] if wesco_matches else customers[0]
    default_idx = customers.index(default_customer)

    selected_customer = st.selectbox(
        "Select customer to preview",
        customers,
        index=default_idx,
        key="customer_preview_select_v2",
    )

    preview_df = df[customer_series.str.lower() == selected_customer.strip().lower()].copy()

    if preview_df.empty:
        st.warning("No orders found for this customer")
        return

    preview_cols = _existing_cols(preview_df, [po_col, customer_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col])
    st.dataframe(preview_df[preview_cols], use_container_width=True, height=420)


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
    st.caption("依併貨日期排序。")

    work = df.copy()
    merge_date_col = _find_col(work, getattr(cfg, "MERGE_DATE_CANDIDATES", []), merge_date_col)

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
    _show_export_buttons(export_df, "Sandy_內部WIP.xlsx", "Sandy 內部 WIP")


def show_sandy_shipment_report(
    df,
    sales_shipment_df=None,
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
    st.caption("優先顯示 Sandy需要的銷貨底.xlsx / 銷貨底 工作表。")

    work = sales_shipment_df.copy() if isinstance(sales_shipment_df, pd.DataFrame) and not sales_shipment_df.empty else df.copy()

    customer_col = _find_col(work, ["客戶", "Customer"], customer_col)
    po_col = _find_col(work, ["PO#"], po_col)
    part_col = _find_col(work, ["P/N"], part_col)
    qty_col = _find_col(work, ["Order Q'TY (PCS)", "Order Q'TY\n (PCS)", "Order Q'TY\n(PCS)", "QTY", "Qty"], qty_col)
    ship_date_col = _find_col(work, ["Ship date", "出貨日期"], ship_date_col)
    wip_col = _find_col(work, ["WIP"], wip_col)
    factory_due_col = _find_col(work, ["工廠交期"], factory_due_col)
    changed_due_date_col = _find_col(work, ["交期(更改)", "交期\n (更改)"], changed_due_date_col)
    merge_date_col = _find_col(work, ["併貨日期(限內部使用)", "併貨日期\n (限內部使用)"], merge_date_col)
    factory_col = _find_col(work, ["工廠"], factory_col)
    remark_col = _find_col(work, ["Note", "Remark"], remark_col)

    if wip_col and wip_col in work.columns:
        wip_s = _safe_text_series(work, wip_col)
        work = work[wip_s.str.contains("SHIPMENT|Shipping|出貨", case=False, na=False)].copy()

    if merge_date_col and merge_date_col in work.columns:
        work["_merge_sort"] = pd.to_datetime(work[merge_date_col], errors="coerce")
        work = work.sort_values("_merge_sort", ascending=True)
    elif ship_date_col and ship_date_col in work.columns:
        work["_ship_sort"] = pd.to_datetime(work[ship_date_col], errors="coerce")
        work = work.sort_values("_ship_sort", ascending=True)

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
    _show_export_buttons(export_df, "Sandy_銷貨底.xlsx", "Sandy 銷貨底")


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
    st.caption("顯示工廠交期 / 交期更改 / 併貨日期 有異動者。")

    work = df.copy()

    order_date_col = _find_col(work, getattr(cfg, "ORDER_DATE_CANDIDATES", []), order_date_col)
    factory_order_date_col = _find_col(work, getattr(cfg, "FACTORY_ORDER_DATE_CANDIDATES", []), factory_order_date_col)
    merge_date_col = _find_col(work, getattr(cfg, "MERGE_DATE_CANDIDATES", []), merge_date_col)
    changed_due_date_col = _find_col(work, getattr(cfg, "CHANGED_DUE_DATE_CANDIDATES", []), changed_due_date_col)
    factory_due_col = _find_col(work, getattr(cfg, "FACTORY_DUE_CANDIDATES", []), factory_due_col)

    due_s = _safe_text_series(work, factory_due_col) if factory_due_col else pd.Series([""] * len(work), index=work.index)
    changed_due_s = _safe_text_series(work, changed_due_date_col) if changed_due_date_col else pd.Series([""] * len(work), index=work.index)
    merge_s = _safe_text_series(work, merge_date_col) if merge_date_col else pd.Series([""] * len(work), index=work.index)

    cond_changed_due_has_value = changed_due_s != ""
    cond_merge_has_value = merge_s != ""
    cond_due_diff = (due_s != "") & (changed_due_s != "") & (due_s != changed_due_s)

    filtered = work[cond_changed_due_has_value | cond_merge_has_value | cond_due_diff].copy()

    if merge_date_col and merge_date_col in filtered.columns:
        filtered["_merge_sort"] = pd.to_datetime(filtered[merge_date_col], errors="coerce")
    else:
        filtered["_merge_sort"] = pd.NaT

    if changed_due_date_col and changed_due_date_col in filtered.columns:
        filtered["_changed_due_sort"] = pd.to_datetime(filtered[changed_due_date_col], errors="coerce")
    else:
        filtered["_changed_due_sort"] = pd.NaT

    filtered = filtered.sort_values(["_merge_sort", "_changed_due_sort"], ascending=[True, True])

    if filtered.empty:
        st.info("目前沒有新訂單 / 異動交期資料。")
        return

    show_cols = _existing_cols(
        filtered,
        [
            customer_col,
            po_col,
            part_col,
            qty_col,
            wip_col,
            ship_date_col,
            factory_due_col,
            changed_due_date_col,
            merge_date_col,
            order_date_col,
            factory_order_date_col,
            remark_col,
        ],
    )

    export_df = filtered[show_cols].copy()
    st.dataframe(export_df, use_container_width=True, height=520)
    _show_export_buttons(export_df, "新訂單WIP.xlsx", "新訂單 WIP")
