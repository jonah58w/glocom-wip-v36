# -*- coding: utf-8 -*-
import inspect
import pandas as pd
import streamlit as st

import config as cfg
import utils as u
import teable_api
import reports
from sales_report import render_sales_report_page


# ================================
# STYLE
# ================================
st.markdown(
    """
    <style>
    .portal-box {
        padding: 18px 20px;
        border: 1px solid rgba(120,120,120,.22);
        border-radius: 16px;
        background: rgba(255,255,255,.03);
        margin-bottom: 14px;
    }
    .portal-title {
        font-size: 1.2rem;
        font-weight: 700;
        margin-bottom: 4px;
    }
    .tag-chip {
        display: inline-block;
        padding: 4px 10px;
        margin: 2px 6px 2px 0;
        border-radius: 999px;
        font-size: 0.82rem;
        border: 1px solid rgba(120,120,120,.25);
        background: rgba(255,255,255,.05);
    }
    .wip-chip {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        font-size: 0.82rem;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ================================
# HELPERS
# ================================
def refresh_after_update():
    st.cache_data.clear()
    st.rerun()


def split_tags(value):
    return u.split_tags(value)


def wip_display_html(value: str) -> str:
    text = u.safe_text(value)
    lower = text.lower()

    if any(k in lower for k in ["完成"]) or text.upper() in cfg.DONE_WIP_VALUES:
        label = text or "完成"
        bg = "#065f46"
        fg = "#d1fae5"
    elif any(k in lower for k in ["ship", "shipping", "出貨"]):
        label = text or "Shipping"
        bg = "#14532d"
        fg = "#dcfce7"
    elif any(k in lower for k in ["pack", "包裝"]):
        label = text or "Packing"
        bg = "#166534"
        fg = "#dcfce7"
    elif any(k in lower for k in ["fqc", "qa", "inspection", "成檢", "測試"]):
        label = text or "Inspection"
        bg = "#854d0e"
        fg = "#fef3c7"
    elif any(k in lower for k in ["aoi", "drill", "route", "routing", "plating", "inner", "production", "防焊", "壓合", "外層", "內層", "成型"]):
        label = text or "Production"
        bg = "#9a3412"
        fg = "#ffedd5"
    elif any(k in lower for k in ["eng", "gerber", "cam", "eq"]):
        label = text or "Engineering"
        bg = "#1d4ed8"
        fg = "#dbeafe"
    elif any(k in lower for k in ["hold", "等待", "暫停"]):
        label = text or "On Hold"
        bg = "#7f1d1d"
        fg = "#fee2e2"
    else:
        label = text or "-"
        bg = "#374151"
        fg = "#f3f4f6"

    return f'<span class="wip-chip" style="background:{bg};color:{fg};">{label}</span>'


def show_metrics(df: pd.DataFrame, wip_col: str | None):
    total_orders = len(df)
    shipping = 0

    if wip_col and wip_col in df.columns:
        wip_series = u.get_series_by_col(df, wip_col)
        if wip_series is not None:
            shipping = len(
                df[
                    wip_series.astype(str).str.contains(
                        "ship|shipping|出貨",
                        case=False,
                        na=False,
                    )
                ]
            )

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


def customer_portal_columns(df, po_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col):
    return [
        c for c in [
            po_col,
            part_col,
            qty_col,
            wip_col,
            ship_date_col,
            customer_tag_col,
            remark_col,
        ]
        if c and c in df.columns
    ]


def _try_call(func, **kwargs):
    try:
        sig = inspect.signature(func)
        accepted = {k: v for k, v in kwargs.items() if k in sig.parameters}
        return func(**accepted)
    except TypeError:
        return False


def call_report_function(possible_names, **kwargs):
    """
    依名稱清單找 reports 裡存在的函式，並自動只傳它需要的參數。
    這樣就算 reports.py 函式簽名有些微不同，也比較不容易炸。
    """
    for name in possible_names:
        func = getattr(reports, name, None)
        if callable(func):
            return _try_call(func, **kwargs)
    return False


def fallback_import_update_page(**kwargs):
    st.subheader("Import / Update")
    st.caption("保留工廠進度輸入功能。若 reports.py 已提供完整頁面，會優先使用；這裡是內建 fallback。")

    uploaded = st.file_uploader(
        "上傳工廠進度檔案（Excel / CSV / TXT）",
        type=["xlsx", "xls", "csv", "txt"],
        accept_multiple_files=False,
    )

    if not uploaded:
        st.info("請上傳工廠進度檔案。")
        return

    name = uploaded.name.lower()
    parsed_df = None
    parse_notes = []

    # 1) 優先嘗試 reports.py / 其他模組既有函式
    module_func_candidates = []
    for module_name in ["reports", "excel_reader", "factory_parsers", "text_ocr_parsers"]:
        try:
            mod = __import__(module_name)
        except Exception:
            continue
        for func_name in [
            "parse_factory_file",
            "parse_factory_progress_file",
            "read_factory_file",
            "read_excel_file",
            "load_excel_file",
            "parse_uploaded_file",
        ]:
            func = getattr(mod, func_name, None)
            if callable(func):
                module_func_candidates.append((module_name, func_name, func))

    for module_name, func_name, func in module_func_candidates:
        try:
            result = _try_call(
                func,
                uploaded_file=uploaded,
                file=uploaded,
                filename=uploaded.name,
                df=kwargs.get("df"),
                orders=kwargs.get("orders"),
            )
            if isinstance(result, pd.DataFrame):
                parsed_df = result
                parse_notes.append(f"使用 {module_name}.{func_name} 成功解析")
                break
            if isinstance(result, tuple) and result and isinstance(result[0], pd.DataFrame):
                parsed_df = result[0]
                parse_notes.append(f"使用 {module_name}.{func_name} 成功解析")
                break
        except Exception as e:
            parse_notes.append(f"{module_name}.{func_name} 失敗：{e}")

    # 2) 內建 fallback 讀檔
    if parsed_df is None:
        try:
            if name.endswith(".csv"):
                parsed_df = pd.read_csv(uploaded)
                parse_notes.append("使用 pandas.read_csv 讀取")
            elif name.endswith(".txt"):
                content = uploaded.read().decode("utf-8", errors="ignore")
                lines = [line.strip() for line in content.splitlines() if line.strip()]
                parsed_df = pd.DataFrame({"raw_text": lines})
                parse_notes.append("以文字行模式匯入 TXT")
            else:
                parsed_df = pd.read_excel(uploaded)
                parse_notes.append("使用 pandas.read_excel 讀取")
        except Exception as e:
            st.error(f"讀取檔案失敗：{e}")
            return

    if parse_notes:
        with st.expander("解析資訊", expanded=False):
            for note in parse_notes:
                st.write("-", note)

    st.success(f"已讀取 {uploaded.name}")
    st.write("欄位：", list(parsed_df.columns))
    st.dataframe(parsed_df, use_container_width=True, height=420)

    csv_data = parsed_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "下載解析結果 CSV",
        data=csv_data,
        file_name=f"parsed_{uploaded.name.rsplit('.', 1)[0]}.csv",
        mime="text/csv",
        use_container_width=True,
    )

    st.info("這個 fallback 版先恢復上傳與解析顯示功能；若你的既有 reports.py 匯入頁存在，系統仍會優先用既有版本。")


# ================================
# LOAD DATA
# ================================
try:
    orders, api_status, api_text = teable_api.load_orders()
except Exception as e:
    orders = pd.DataFrame()
    api_status = "EXCEPTION"
    api_text = str(e)


# ================================
# DETECT KEY COLUMNS
# ================================
po_col = u.get_first_matching_column(orders, cfg.PO_CANDIDATES)
customer_col = u.get_first_matching_column(orders, cfg.CUSTOMER_CANDIDATES)
part_col = u.get_first_matching_column(orders, cfg.PART_CANDIDATES)
qty_col = u.get_first_matching_column(orders, cfg.QTY_CANDIDATES)
factory_col = u.get_first_matching_column(orders, cfg.FACTORY_CANDIDATES)
wip_col = u.get_first_matching_column(orders, cfg.WIP_CANDIDATES)
factory_due_col = u.get_first_matching_column(orders, cfg.FACTORY_DUE_CANDIDATES)
ship_date_col = u.get_first_matching_column(orders, cfg.SHIP_DATE_CANDIDATES)
remark_col = u.get_first_matching_column(orders, cfg.REMARK_CANDIDATES)
customer_tag_col = u.get_first_matching_column(orders, cfg.CUSTOMER_TAG_CANDIDATES)
merge_date_col = u.get_first_matching_column(orders, cfg.MERGE_DATE_CANDIDATES)
order_date_col = u.get_first_matching_column(orders, cfg.ORDER_DATE_CANDIDATES)
factory_order_date_col = u.get_first_matching_column(orders, cfg.FACTORY_ORDER_DATE_CANDIDATES)
changed_due_date_col = u.get_first_matching_column(orders, cfg.CHANGED_DUE_DATE_CANDIDATES)


# ================================
# CUSTOMER MODE
# ================================
query = st.query_params
customer_param = query.get("customer", None)

if customer_param:
    st.title("GLOCOM Order Status")
    st.caption("Customer WIP Progress")

    if not customer_col:
        st.error("Customer column not found")
        st.stop()

    customer_series = u.get_series_by_col(orders, customer_col)
    if customer_series is None:
        st.error("Customer data unavailable")
        st.stop()

    cust_orders = orders[
        customer_series.astype(str).str.strip().str.lower()
        == str(customer_param).strip().lower()
    ].copy()

    if cust_orders.empty:
        st.warning("No orders found")
        st.stop()

    show_metrics(cust_orders, wip_col)
    st.divider()

    for _, row in cust_orders.iterrows():
        po_val = u.safe_text(row.get(po_col, "")) if po_col else ""
        part_val = u.safe_text(row.get(part_col, "")) if part_col else ""
        qty_val = u.safe_text(row.get(qty_col, "")) if qty_col else ""
        wip_val = u.safe_text(row.get(wip_col, "")) if wip_col else ""
        ship_val = u.safe_text(row.get(ship_date_col, "")) if ship_date_col else ""
        remark_val = u.safe_text(row.get(remark_col, "")) if remark_col else ""
        tags_val = split_tags(row.get(customer_tag_col, "")) if customer_tag_col else []

        tag_html = "".join([f'<span class="tag-chip">{t}</span>' for t in tags_val]) if tags_val else '<span class="tag-chip">-</span>'

        st.markdown(
            f"""
            <div class="portal-box">
                <div class="portal-title">{po_val or "-"}</div>
                <div style="display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;">
                    <div>
                        <div><strong>P/N</strong> : {part_val or "-"}</div>
                        <div><strong>Qty</strong> : {qty_val or "-"}</div>
                    </div>
                    <div>
                        <div><strong>WIP</strong> : {wip_display_html(wip_val)}</div>
                        <div style="margin-top:8px;"><strong>Ship Date</strong> : {ship_val or "-"}</div>
                    </div>
                </div>
                <div style="margin-top:12px;"><strong>Customer Remark Tags</strong> : {tag_html}</div>
                <div style="margin-top:10px;"><strong>Remark</strong> : {remark_val or "-"}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    portal_cols = customer_portal_columns(
        cust_orders,
        po_col,
        part_col,
        qty_col,
        wip_col,
        ship_date_col,
        customer_tag_col,
        remark_col,
    )
    csv_data = cust_orders[portal_cols].to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        "Download WIP CSV",
        data=csv_data,
        file_name=f"{customer_param}_wip.csv",
        mime="text/csv",
    )
    st.stop()


# ================================
# INTERNAL MODE
# ================================
st.title("🏭 GLOCOM Control Tower")
st.caption("Internal PCB Production Monitoring System")

with st.expander("Debug"):
    st.write("API Status:", api_status)
    st.write("TABLE_URL:", cfg.TABLE_URL)
    st.write("Token loaded:", bool(cfg.TEABLE_TOKEN))
    st.write("Columns:", list(orders.columns) if not orders.empty else [])
    if isinstance(api_text, str):
        st.text(api_text[:1200])

if orders.empty:
    show_no_data_layout()
    st.stop()

st.sidebar.title("GLOCOM Internal")
st.sidebar.link_button("Open Teable", cfg.TEABLE_WEB_URL, use_container_width=True)

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
        "業績明細表",
        "Import / Update",
    ],
)

if st.sidebar.button("Refresh"):
    refresh_after_update()

st.sidebar.markdown("---")
st.sidebar.caption("完成案件請在 Teable 主 View 設定篩選：WIP ≠ 完成")
st.sidebar.caption("另建 Completed View：WIP = 完成")


# ================================
# FALLBACK VIEWS
# ================================
def fallback_factory_load(df):
    st.subheader("🏭 Factory Load")
    if factory_col and factory_col in df.columns:
        factory_series = u.get_series_by_col(df, factory_col)
        if factory_series is not None:
            factory_summary = (
                factory_series.fillna("(blank)")
                .astype(str)
                .value_counts()
                .reset_index()
            )
            factory_summary.columns = [factory_col, "Orders"]
            st.bar_chart(factory_summary.set_index(factory_col))
            st.dataframe(factory_summary, use_container_width=True, height=400)
            return
    st.info("No factory data")


def fallback_delayed_orders(df):
    st.subheader("⚠️ Delayed Orders")
    if not factory_due_col or factory_due_col not in df.columns:
        st.info("No factory due date column")
        return

    temp = df.copy()
    due_series = u.get_series_by_col(temp, factory_due_col)
    if due_series is None:
        st.info("No factory due date data")
        return

    temp["_FactoryDueDateParsed"] = u.safe_to_datetime(due_series)
    today = pd.Timestamp.today().normalize()

    delayed = temp[
        temp["_FactoryDueDateParsed"].notna()
        & (temp["_FactoryDueDateParsed"] < today)
    ].copy()

    if delayed.empty:
        st.success("No delayed orders")
        return

    delayed["Delay Days"] = (today - delayed["_FactoryDueDateParsed"]).dt.days
    show_cols = [
        c for c in [po_col, customer_col, part_col, qty_col, factory_col, wip_col, factory_due_col]
        if c and c in delayed.columns
    ]
    if "Delay Days" not in show_cols:
        show_cols.append("Delay Days")
    st.dataframe(delayed[show_cols], use_container_width=True, height=520)


def fallback_shipment_forecast(df):
    st.subheader("📦 Shipment Forecast (Next 7 days)")
    if not ship_date_col or ship_date_col not in df.columns:
        st.info("No ship date column")
        return

    temp = df.copy()
    ship_series = u.get_series_by_col(temp, ship_date_col)
    if ship_series is None:
        st.info("No ship date data")
        return

    temp["_ShipDateParsed"] = u.safe_to_datetime(ship_series)
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

    show_cols = [
        c for c in [po_col, customer_col, part_col, qty_col, factory_col, wip_col, ship_date_col]
        if c and c in forecast.columns
    ]
    st.dataframe(
        forecast.sort_values("_ShipDateParsed")[show_cols],
        use_container_width=True,
        height=520,
    )


def fallback_orders(df):
    st.subheader("📋 Orders")
    filtered = df.copy()

    col1, col2 = st.columns(2)

    if customer_col and customer_col in filtered.columns:
        customer_series = u.get_series_by_col(filtered, customer_col)
        if customer_series is not None:
            customer_options = ["All"] + sorted([str(x) for x in customer_series.dropna().unique().tolist()])
            selected_customer = col1.selectbox("Customer", customer_options)
            if selected_customer != "All":
                filtered = filtered[
                    u.get_series_by_col(filtered, customer_col).astype(str) == selected_customer
                ]

    if wip_col and wip_col in filtered.columns:
        wip_series = u.get_series_by_col(filtered, wip_col)
        if wip_series is not None:
            wip_options = ["All"] + sorted([str(x) for x in wip_series.dropna().unique().tolist()])
            selected_wip = col2.selectbox("WIP Stage", wip_options)
            if selected_wip != "All":
                filtered = filtered[
                    u.get_series_by_col(filtered, wip_col).astype(str) == selected_wip
                ]

    st.dataframe(filtered, use_container_width=True, height=520)

    csv_data = filtered.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Download Orders CSV",
        data=csv_data,
        file_name="glocom_orders.csv",
        mime="text/csv",
    )


# ================================
# ROUTING
# ================================
common_kwargs = dict(
    df=orders,
    orders=orders,
    po_col=po_col,
    customer_col=customer_col,
    part_col=part_col,
    qty_col=qty_col,
    factory_col=factory_col,
    wip_col=wip_col,
    factory_due_col=factory_due_col,
    ship_date_col=ship_date_col,
    remark_col=remark_col,
    customer_tag_col=customer_tag_col,
    merge_date_col=merge_date_col,
    order_date_col=order_date_col,
    factory_order_date_col=factory_order_date_col,
    changed_due_date_col=changed_due_date_col,
)


if menu == "Dashboard":
    ok = call_report_function(
        ["show_dashboard_report"],
        **common_kwargs,
    )
    if ok is False:
        show_metrics(orders, wip_col)
        st.divider()
        left, right = st.columns(2)
        with left:
            fallback_factory_load(orders)
        with right:
            fallback_shipment_forecast(orders)

elif menu == "Factory Load":
    ok = call_report_function(["show_factory_load_report"], **common_kwargs)
    if ok is False:
        fallback_factory_load(orders)

elif menu == "Delayed Orders":
    ok = call_report_function(["show_delayed_orders_report"], **common_kwargs)
    if ok is False:
        fallback_delayed_orders(orders)

elif menu == "Shipment Forecast":
    ok = call_report_function(["show_shipment_forecast_report"], **common_kwargs)
    if ok is False:
        fallback_shipment_forecast(orders)

elif menu == "Orders":
    ok = call_report_function(["show_orders_report"], **common_kwargs)
    if ok is False:
        fallback_orders(orders)

elif menu == "Customer Preview":
    ok = call_report_function(["show_customer_preview_report"], default_customer="WESCO", **common_kwargs)
    if ok is False:
        st.subheader("Customer Preview")
        st.caption("預設顯示 WESCO，但仍保留下拉選項。客戶請直接使用 Teable View。")
        if not customer_col or customer_col not in orders.columns:
            st.error("Customer column not found in Teable data")
        else:
            customer_series = u.get_series_by_col(orders, customer_col)
            if customer_series is None:
                st.error("Customer data unavailable")
            else:
                customers = sorted([str(x).strip() for x in customer_series.dropna().unique().tolist() if str(x).strip()])
                if not customers:
                    st.warning("No customers found")
                else:
                    target_default = "WESCO"
                    default_idx = 0
                    for idx, name in enumerate(customers):
                        if name.strip().lower() == target_default.lower():
                            default_idx = idx
                            break
                    selected_customer = st.selectbox("Select customer to preview", customers, index=default_idx)
                    preview_df = orders[
                        customer_series.astype(str).str.strip().str.lower()
                        == selected_customer.strip().lower()
                    ].copy()
                    if preview_df.empty:
                        st.warning("No orders found for this customer")
                    else:
                        show_metrics(preview_df, wip_col)
                        preview_cols = [
                            c for c in [po_col, customer_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col]
                            if c and c in preview_df.columns
                        ]
                        st.dataframe(preview_df[preview_cols], use_container_width=True, height=420)

elif menu == "Sandy 內部 WIP":
    ok = call_report_function(["show_sandy_internal_wip_report"], **common_kwargs)
    if ok is False:
        st.warning("reports.py 尚未提供 Sandy 內部 WIP 報表函式。")

elif menu == "Sandy 銷貨底":
    ok = call_report_function(["show_sandy_shipment_report"], **common_kwargs)
    if ok is False:
        st.warning("reports.py 尚未提供 Sandy 銷貨底報表函式。")

elif menu == "新訂單 WIP":
    ok = call_report_function(["show_new_orders_wip_report"], **common_kwargs)
    if ok is False:
        st.warning("reports.py 尚未提供 新訂單 WIP 報表函式。")

elif menu == "業績明細表":
    render_sales_report_page(
        df=orders,
        orders=orders,
        po_col=po_col,
        customer_col=customer_col,
        part_col=part_col,
        qty_col=qty_col,
        factory_col=factory_col,
        wip_col=wip_col,
        ship_date_col=ship_date_col,
        remark_col=remark_col,
        order_date_col=order_date_col,
    )

elif menu == "Import / Update":
    ok = call_report_function(
        [
            "show_import_update_page",
            "show_import_update_report",
            "render_import_update_page",
            "render_import_update_report",
        ],
        **common_kwargs,
    )
    if ok is False:
        fallback_import_update_page(**common_kwargs)

st.caption("Auto refresh cache: 60 seconds")
