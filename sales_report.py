import io
from datetime import datetime
import pandas as pd
import streamlit as st


def _pick_col(df, candidates):
    for c in candidates:
        if c and c in df.columns:
            return c
    return None


def _safe_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)


def render_sales_report_page(
    df=None,
    orders=None,
    po_col=None,
    customer_col=None,
    part_col=None,
    qty_col=None,
    factory_col=None,
    wip_col=None,
    ship_date_col=None,
    remark_col=None,
    order_date_col=None,
):
    source_df = orders if orders is not None and not orders.empty else df

    st.subheader("業績明細表")
    st.caption("沿用目前 WIP 主資料，不額外要求新的 Teable Sales Table ID。")

    if source_df is None or source_df.empty:
        st.warning("目前沒有可用資料")
        return

    work = source_df.copy()

    amount_col = _pick_col(work, ["出貨金額", "金額", "Amount", "Sales", "Shipment Amount"])
    hold_col = _pick_col(work, ["HOLD金額", "HOLD", "Hold Amount"])
    discount_col = _pick_col(work, ["折讓", "Discount"])
    net_col = _pick_col(work, ["淨出貨金額", "Net", "Net Amount"])

    if not customer_col:
        customer_col = _pick_col(work, ["客戶", "Customer", "客戶名稱"])
    if not factory_col:
        factory_col = _pick_col(work, ["工廠", "Factory"])
    if not order_date_col:
        order_date_col = _pick_col(work, ["日期", "Date", "出貨日期", "Ship Date"])
    if not ship_date_col:
        ship_date_col = _pick_col(work, ["出貨日期", "Ship Date", "日期", "Date"])

    report_month = st.text_input("報表月份 (YYYY-MM)", value=datetime.now().strftime("%Y-%m"))
    company_name = st.text_input("子表名稱 / 公司名稱", value="")
    currency_symbol = st.text_input("幣別符號", value="NT$")

    if not order_date_col and not ship_date_col:
        st.error("找不到日期欄位，至少要有 日期 / 出貨日期 / Ship Date 其中一欄")
        return

    date_col = order_date_col or ship_date_col
    work["_report_date"] = pd.to_datetime(work[date_col], errors="coerce")
    work["_month"] = work["_report_date"].dt.strftime("%Y-%m")

    month_df = work[work["_month"] == report_month].copy()

    if amount_col:
        month_df["_amount"] = _safe_num(month_df[amount_col])
    else:
        month_df["_amount"] = 0

    if hold_col:
        month_df["_hold"] = _safe_num(month_df[hold_col])
    else:
        month_df["_hold"] = 0

    if discount_col:
        month_df["_discount"] = _safe_num(month_df[discount_col])
    else:
        month_df["_discount"] = 0

    if net_col:
        month_df["_net"] = _safe_num(month_df[net_col])
    else:
        month_df["_net"] = month_df["_amount"] - month_df["_hold"] - month_df["_discount"]

    total_amount = float(month_df["_amount"].sum())
    total_net = float(month_df["_net"].sum())
    customer_count = month_df[customer_col].nunique() if customer_col and customer_col in month_df.columns else 0
    factory_count = month_df[factory_col].nunique() if factory_col and factory_col in month_df.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("總業績", f"{currency_symbol} {total_amount:,.0f}")
    c2.metric("客戶數", int(customer_count))
    c3.metric("廠商數", int(factory_count))
    c4.metric("淨出貨", f"{currency_symbol} {total_net:,.0f}")

    if customer_col and customer_col in month_df.columns:
        by_customer = (
            month_df.groupby(customer_col)["_net"]
            .sum()
            .sort_values(ascending=False)
            .head(10)
        )
        left, right = st.columns(2)
        with left:
            st.markdown("**客戶業績比較**")
            st.bar_chart(by_customer)
        with right:
            st.markdown("**業績佔比**")
            pct_df = by_customer.reset_index()
            pct_df.columns = ["客戶", "淨出貨金額"]
            total = pct_df["淨出貨金額"].sum()
            pct_df["佔比%"] = (pct_df["淨出貨金額"] / total * 100).round(2) if total else 0
            st.dataframe(pct_df, use_container_width=True, hide_index=True)

    if date_col:
        daily = (
            month_df.groupby(month_df["_report_date"].dt.strftime("%Y-%m-%d"))["_net"]
            .sum()
            .reset_index()
        )
        daily.columns = ["日期", "淨出貨金額"]
        if not daily.empty:
            st.markdown("**每日趨勢**")
            st.line_chart(daily.set_index("日期"))

    show_cols = [
        c for c in [
            order_date_col, ship_date_col, customer_col, factory_col,
            po_col, part_col, qty_col, amount_col, hold_col, discount_col, net_col, remark_col
        ]
        if c and c in month_df.columns
    ]

    st.markdown("**業績明細表**")
    st.dataframe(month_df[show_cols], use_container_width=True, hide_index=True)

    csv_data = month_df[show_cols].to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "下載業績明細 CSV",
        data=csv_data,
        file_name=f"sales_report_{report_month}.csv",
        mime="text/csv",
        use_container_width=True,
    )
