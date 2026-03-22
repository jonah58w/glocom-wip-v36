from __future__ import annotations

import re
from datetime import datetime
from typing import Any

import pandas as pd
import streamlit as st


def _pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        c = normalized.get(str(cand).strip().lower())
        if c:
            return c
    return None


def _keyword_pick(df: pd.DataFrame, keywords: list[str]) -> str | None:
    for col in df.columns:
        c = str(col).strip().lower()
        if any(k in c for k in keywords):
            return col
    return None


def _num_like_cols(df: pd.DataFrame) -> list[str]:
    out = []
    for col in df.columns:
        s = _to_num_series(df[col])
        if s.notna().sum() and float(s.abs().sum()) > 0:
            out.append(col)
    return out


def _to_num_series(series: pd.Series) -> pd.Series:
    s = series.astype(str)
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace("US$", "", regex=False)
    s = s.str.replace("USD", "", regex=False)
    s = s.str.replace("NT$", "", regex=False)
    s = s.str.replace("$", "", regex=False)
    s = s.str.replace("￥", "", regex=False)
    s = s.str.replace("¥", "", regex=False)
    s = s.str.replace("€", "", regex=False)
    s = s.str.extract(r"([-+]?\d*\.?\d+)", expand=False)
    return pd.to_numeric(s, errors="coerce").fillna(0)


def _normalize_month_text(text: str) -> str:
    text = (text or "").strip()
    m = re.match(r"^(\d{4})-(\d{1,2})$", text)
    if not m:
        return datetime.now().strftime("%Y-%m")
    y, mm = m.groups()
    return f"{y}-{int(mm):02d}"


def _fmt_money(val: float, symbol: str = "US$") -> str:
    return f"{symbol} {val:,.0f}"


def _safe_series(df: pd.DataFrame, col: str | None) -> pd.Series:
    if col and col in df.columns:
        return df[col]
    return pd.Series([0] * len(df), index=df.index)


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
    **kwargs: Any,
):
    source_df = orders if isinstance(orders, pd.DataFrame) and not orders.empty else df

    st.subheader("業績明細表")
    st.caption("只統計所選月份；預設幣別為美金；可手動修正欄位對應。")

    if not isinstance(source_df, pd.DataFrame) or source_df.empty:
        st.warning("目前沒有可用資料")
        return

    work = source_df.copy()

    if not customer_col or customer_col not in work.columns:
        customer_col = _pick_col(work, ["客戶", "Customer", "客戶名稱"])
    if not factory_col or factory_col not in work.columns:
        factory_col = _pick_col(work, ["工廠", "Factory", "Vendor"])
    if not order_date_col or order_date_col not in work.columns:
        order_date_col = _pick_col(work, ["Order date", "Order Date", "日期", "接單日期", "建立日期"])
    if not ship_date_col or ship_date_col not in work.columns:
        ship_date_col = _pick_col(work, ["Ship date", "Ship Date", "出貨日期", "日期"])

    order_amount_candidates = [
        "Order Amount", "Order Amt", "Order Value", "接單金額", "接單總額", "Amount", "Sales Amount", "USD Amount", "Total Amount", "Order USD", "Sales USD",
    ]
    ship_amount_candidates = [
        "Shipment Amount", "Ship Amount", "出貨金額", "出貨總額", "Invoice Amount", "Net Shipment", "淨出貨金額", "Net Amount", "Shipment USD", "Invoice USD",
    ]
    unit_price_candidates = [
        "Unit Price", "Price", "單價", "USD/PCS", "USD Price", "Sell Price", "Sales Price", "Price USD",
    ]
    hold_candidates = ["HOLD金額", "Hold Amount", "HOLD", "Hold"]
    discount_candidates = ["折讓", "Discount", "Discount Amount"]

    auto_order_amount_col = _pick_col(work, order_amount_candidates) or _keyword_pick(work, ["order amount", "接單金額", "sales amount", "amount", "usd"])
    auto_ship_amount_col = _pick_col(work, ship_amount_candidates) or _keyword_pick(work, ["shipment amount", "ship amount", "出貨金額", "invoice", "net shipment", "淨出貨"])
    auto_unit_price_col = _pick_col(work, unit_price_candidates) or _keyword_pick(work, ["unit price", "單價", "price"])
    auto_hold_col = _pick_col(work, hold_candidates) or _keyword_pick(work, ["hold"])
    auto_discount_col = _pick_col(work, discount_candidates) or _keyword_pick(work, ["discount", "折讓"])

    numeric_cols = _num_like_cols(work)
    default_month = datetime.now().strftime("%Y-%m")
    c1, c2, c3 = st.columns(3)
    with c1:
        report_month_input = st.text_input("報表月份 (YYYY-MM)", value=default_month)
    with c2:
        company_name = st.text_input("子表名稱 / 公司名稱", value="")
    with c3:
        currency_symbol = st.text_input("幣別符號", value="US$")

    report_month = _normalize_month_text(report_month_input)

    with st.expander("欄位偵測", expanded=True):
        cols = ["(無)"] + list(work.columns)
        date_default = ship_date_col or order_date_col or "(無)"
        order_default = auto_order_amount_col or "(無)"
        ship_default = auto_ship_amount_col or "(無)"
        unit_default = auto_unit_price_col or "(無)"
        hold_default = auto_hold_col or "(無)"
        discount_default = auto_discount_col or "(無)"

        d1, d2, d3, d4, d5, d6 = st.columns(6)
        date_col = d1.selectbox("日期欄位", cols, index=cols.index(date_default) if date_default in cols else 0)
        order_amount_col = d2.selectbox("接單金額欄位", cols, index=cols.index(order_default) if order_default in cols else 0)
        ship_amount_col = d3.selectbox("出貨金額欄位", cols, index=cols.index(ship_default) if ship_default in cols else 0)
        unit_price_col = d4.selectbox("單價欄位", cols, index=cols.index(unit_default) if unit_default in cols else 0)
        hold_col = d5.selectbox("HOLD欄位", cols, index=cols.index(hold_default) if hold_default in cols else 0)
        discount_col = d6.selectbox("折讓欄位", cols, index=cols.index(discount_default) if discount_default in cols else 0)
        st.write(f"客戶欄位：{customer_col or '(無)'} / 工廠欄位：{factory_col or '(無)'} / 數量欄位：{qty_col or '(無)'}")
        st.write("可用數值欄位：", numeric_cols[:20])

    if date_col == "(無)":
        st.error("請指定日期欄位")
        return

    work["_report_date"] = pd.to_datetime(work[date_col], errors="coerce")
    work["_month"] = work["_report_date"].dt.strftime("%Y-%m")
    month_df = work[work["_month"] == report_month].copy()

    if month_df.empty:
        st.warning(f"{report_month} 沒有資料。")
        return

    qty_series = _to_num_series(_safe_series(month_df, qty_col)) if qty_col and qty_col in month_df.columns else pd.Series([0] * len(month_df), index=month_df.index)
    unit_price_series = _to_num_series(_safe_series(month_df, unit_price_col)) if unit_price_col != "(無)" else pd.Series([0] * len(month_df), index=month_df.index)

    if order_amount_col != "(無)":
        month_df["_order_amount"] = _to_num_series(month_df[order_amount_col])
    elif unit_price_col != "(無)" and qty_col and qty_col in month_df.columns:
        month_df["_order_amount"] = qty_series * unit_price_series
    else:
        month_df["_order_amount"] = 0

    if ship_amount_col != "(無)":
        month_df["_ship_amount"] = _to_num_series(month_df[ship_amount_col])
    elif unit_price_col != "(無)" and qty_col and qty_col in month_df.columns:
        month_df["_ship_amount"] = qty_series * unit_price_series
    else:
        month_df["_ship_amount"] = 0

    month_df["_hold"] = _to_num_series(month_df[hold_col]) if hold_col != "(無)" else 0
    month_df["_discount"] = _to_num_series(month_df[discount_col]) if discount_col != "(無)" else 0
    month_df["_net_ship"] = month_df["_ship_amount"] - month_df["_hold"] - month_df["_discount"]

    order_total = float(month_df["_order_amount"].sum())
    ship_total = float(month_df["_ship_amount"].sum())
    net_ship_total = float(month_df["_net_ship"].sum())
    customer_count = int(month_df[customer_col].astype(str).replace("", pd.NA).dropna().nunique()) if customer_col and customer_col in month_df.columns else 0
    factory_count = int(month_df[factory_col].astype(str).replace("", pd.NA).dropna().nunique()) if factory_col and factory_col in month_df.columns else 0

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("接單金額", _fmt_money(order_total, currency_symbol))
    m2.metric("出貨金額", _fmt_money(ship_total, currency_symbol))
    m3.metric("淨出貨", _fmt_money(net_ship_total, currency_symbol))
    m4.metric("客戶數", customer_count)
    m5.metric("廠商數", factory_count)

    if order_total == 0 and ship_total == 0:
        st.warning("目前接單/出貨金額仍為 0。請在『欄位偵測』中手動指定接單金額、出貨金額，或指定單價欄位搭配數量欄位計算。")

    if customer_col and customer_col in month_df.columns:
        by_customer = (
            month_df.groupby(customer_col)[["_order_amount", "_ship_amount", "_net_ship"]]
            .sum()
            .sort_values(["_order_amount", "_ship_amount"], ascending=False)
            .reset_index()
        )
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**客戶接單金額比較**")
            st.bar_chart(by_customer.set_index(customer_col)[["_order_amount"]].head(10))
        with c2:
            st.markdown("**客戶出貨金額比較**")
            st.bar_chart(by_customer.set_index(customer_col)[["_ship_amount"]].head(10))
        display_df = by_customer.rename(columns={
            customer_col: "客戶",
            "_order_amount": "接單金額",
            "_ship_amount": "出貨金額",
            "_net_ship": "淨出貨",
        }).copy()
        for col in ["接單金額", "出貨金額", "淨出貨"]:
            display_df[col] = display_df[col].map(lambda x: _fmt_money(float(x), currency_symbol))
        st.dataframe(display_df, use_container_width=True, hide_index=True)

    daily = (
        month_df.groupby(month_df["_report_date"].dt.strftime("%Y-%m-%d"))[["_order_amount", "_ship_amount", "_net_ship"]]
        .sum()
        .reset_index()
    )
    if not daily.empty:
        daily.columns = ["日期", "接單金額", "出貨金額", "淨出貨"]
        st.markdown("**每日趨勢**")
        st.line_chart(daily.set_index("日期")[["接單金額", "出貨金額", "淨出貨"]])

    show_cols = [
        c for c in [date_col, po_col, customer_col, factory_col, part_col, qty_col, unit_price_col if unit_price_col != '(無)' else None, wip_col, ship_date_col, remark_col]
        if c and c in month_df.columns
    ]
    for c in [order_amount_col, ship_amount_col, hold_col, discount_col]:
        if c != "(無)" and c in month_df.columns and c not in show_cols:
            show_cols.append(c)

    st.markdown("**業績明細表**")
    st.dataframe(month_df[show_cols], use_container_width=True, hide_index=True, height=420)

    csv_data = month_df[show_cols].to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "下載業績明細 CSV",
        data=csv_data,
        file_name=f"sales_report_{report_month}.csv",
        mime="text/csv",
        use_container_width=True,
    )
