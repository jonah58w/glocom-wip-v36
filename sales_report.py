# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from datetime import datetime
from typing import Any, Optional

import pandas as pd
import streamlit as st


def _pick_existing(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for c in candidates:
        hit = lower_map.get(c.strip().lower())
        if hit is not None:
            return hit
    return None


def _parse_number_series(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.replace(",", "", regex=False)
    s = s.str.replace("US$", "", regex=False)
    s = s.str.replace("USD", "", regex=False)
    s = s.str.replace("NT$", "", regex=False)
    s = s.str.replace("$", "", regex=False)
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(r"[^0-9.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")


def _candidate_numeric_columns(df: pd.DataFrame) -> list[str]:
    cols: list[str] = []
    for c in df.columns:
        try:
            parsed = _parse_number_series(df[c])
            if parsed.notna().sum() > 0:
                cols.append(c)
        except Exception:
            pass
    return cols


def _normalize_month(text: str) -> str:
    text = (text or "").strip()
    m = re.match(r"^(\d{4})-(\d{1,2})$", text)
    if not m:
        return datetime.now().strftime("%Y-%m")
    y, mm = m.groups()
    return f"{y}-{int(mm):02d}"


def _fmt_money(v: float, symbol: str = "US$") -> str:
    return f"{symbol} {v:,.0f}"


def render_sales_report_page(
    df: Optional[pd.DataFrame] = None,
    orders: Optional[pd.DataFrame] = None,
    po_col: Optional[str] = None,
    customer_col: Optional[str] = None,
    part_col: Optional[str] = None,
    qty_col: Optional[str] = None,
    factory_col: Optional[str] = None,
    wip_col: Optional[str] = None,
    ship_date_col: Optional[str] = None,
    remark_col: Optional[str] = None,
    order_date_col: Optional[str] = None,
    **kwargs: Any,
):
    source_df = orders if isinstance(orders, pd.DataFrame) and not orders.empty else df

    st.subheader("業績明細表")
    st.caption("只統計所選月份；預設幣別為美金。")

    if source_df is None or source_df.empty:
        st.warning("沒有可用資料")
        return

    work = source_df.copy()

    if customer_col not in work.columns:
        customer_col = _pick_existing(work, ["客戶", "customer", "customer name"])
    if factory_col not in work.columns:
        factory_col = _pick_existing(work, ["工廠", "factory"])
    if qty_col not in work.columns:
        qty_col = _pick_existing(work, ["Order Q'TY (PCS)", "Qty", "Quantity", "數量", "pcs"])
    date_col = ship_date_col if ship_date_col in work.columns else None
    if not date_col:
        date_col = order_date_col if order_date_col in work.columns else None
    if not date_col:
        date_col = _pick_existing(work, ["Ship date", "Ship Date", "出貨日期", "日期", "Date", "Order Date"])

    numeric_cols = _candidate_numeric_columns(work)

    order_amount_auto = _pick_existing(work, [
        "接單金額", "order amount", "order amt", "sales amount", "amount", "usd amount", "total amount",
        "net amount", "invoice amount", "金額"
    ])
    ship_amount_auto = _pick_existing(work, [
        "出貨金額", "shipment amount", "ship amount", "invoice amount", "net shipment", "淨出貨金額"
    ])
    unit_price_auto = _pick_existing(work, [
        "單價", "unit price", "price", "usd/pcs", "pcs price", "us$", "usd"
    ])
    hold_auto = _pick_existing(work, ["hold", "hold amount", "HOLD金額"])
    discount_auto = _pick_existing(work, ["折讓", "discount", "discount amount"])

    month_default = datetime.now().strftime("%Y-%m")
    c1, c2, c3 = st.columns(3)
    report_month = _normalize_month(c1.text_input("報表月份 (YYYY-MM)", value=month_default))
    c2.text_input("子表名稱 / 公司名稱", value="")
    currency_symbol = c3.text_input("幣別符號", value="US$")

    with st.expander("欄位偵測"):
        if not date_col:
            st.error("找不到日期欄位")
            return
        date_options = [c for c in work.columns]
        numeric_options = ["(無)"] + numeric_cols
        d1, d2, d3, d4, d5, d6 = st.columns(6)
        date_col = d1.selectbox("日期欄位", date_options, index=date_options.index(date_col) if date_col in date_options else 0)
        order_amount_col = d2.selectbox("接單金額欄位", numeric_options, index=numeric_options.index(order_amount_auto) if order_amount_auto in numeric_options else 0)
        ship_amount_col = d3.selectbox("出貨金額欄位", numeric_options, index=numeric_options.index(ship_amount_auto) if ship_amount_auto in numeric_options else 0)
        unit_price_col = d4.selectbox("單價欄位", numeric_options, index=numeric_options.index(unit_price_auto) if unit_price_auto in numeric_options else 0)
        hold_col = d5.selectbox("HOLD欄位", numeric_options, index=numeric_options.index(hold_auto) if hold_auto in numeric_options else 0)
        discount_col = d6.selectbox("折讓欄位", numeric_options, index=numeric_options.index(discount_auto) if discount_auto in numeric_options else 0)
        st.write(f"客戶欄位：{customer_col or '-'} / 工廠欄位：{factory_col or '-'} / 數量欄位：{qty_col or '-'}")

    work["_report_date"] = pd.to_datetime(work[date_col], errors="coerce")
    work["_month"] = work["_report_date"].dt.strftime("%Y-%m")
    month_df = work[work["_month"] == report_month].copy()

    def parse_col(colname: str) -> pd.Series:
        if not colname or colname == "(無)" or colname not in month_df.columns:
            return pd.Series([0.0] * len(month_df), index=month_df.index)
        return _parse_number_series(month_df[colname]).fillna(0.0)

    qty_series = parse_col(qty_col) if qty_col in month_df.columns else pd.Series([0.0] * len(month_df), index=month_df.index)
    unit_price_series = parse_col(unit_price_col)
    order_amount_series = parse_col(order_amount_col)
    ship_amount_series = parse_col(ship_amount_col)
    hold_series = parse_col(hold_col)
    discount_series = parse_col(discount_col)

    if float(order_amount_series.sum()) == 0 and float(unit_price_series.sum()) != 0 and float(qty_series.sum()) != 0:
        order_amount_series = unit_price_series * qty_series
    if float(ship_amount_series.sum()) == 0 and float(unit_price_series.sum()) != 0 and float(qty_series.sum()) != 0:
        ship_amount_series = unit_price_series * qty_series

    month_df["_order_amount"] = order_amount_series
    month_df["_ship_amount"] = ship_amount_series
    month_df["_hold"] = hold_series
    month_df["_discount"] = discount_series
    month_df["_net"] = month_df["_ship_amount"] - month_df["_hold"] - month_df["_discount"]

    total_order = float(month_df["_order_amount"].sum())
    total_ship = float(month_df["_ship_amount"].sum())
    total_net = float(month_df["_net"].sum())

    customer_count = int(month_df[customer_col].astype(str).replace("", pd.NA).dropna().nunique()) if customer_col and customer_col in month_df.columns else 0
    factory_count = int(month_df[factory_col].astype(str).replace("", pd.NA).dropna().nunique()) if factory_col and factory_col in month_df.columns else 0

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("接單金額", _fmt_money(total_order, currency_symbol))
    m2.metric("出貨金額", _fmt_money(total_ship, currency_symbol))
    m3.metric("淨出貨", _fmt_money(total_net, currency_symbol))
    m4.metric("客戶數", customer_count)
    m5.metric("廠商數", factory_count)

    if customer_col and customer_col in month_df.columns:
        by_customer = (
            month_df.groupby(customer_col)["_ship_amount"]
            .sum()
            .reset_index()
            .sort_values("_ship_amount", ascending=False)
        )
        left, right = st.columns(2)
        with left:
            st.markdown("**客戶業績比較**")
            if by_customer.empty:
                st.info("本月份沒有資料")
            else:
                chart_df = by_customer.set_index(customer_col)
                st.bar_chart(chart_df)
        with right:
            st.markdown("**業績佔比**")
            if by_customer.empty:
                st.info("本月份沒有資料")
            else:
                show = by_customer.copy()
                show.columns = ["客戶", "出貨金額"]
                total = show["出貨金額"].sum()
                show["佔比%"] = (show["出貨金額"] / total * 100).round(2) if total else 0
                show["出貨金額"] = show["出貨金額"].map(lambda x: _fmt_money(x, currency_symbol))
                st.dataframe(show, use_container_width=True, hide_index=True, height=340)

    detail_cols = [c for c in [date_col, po_col, customer_col, part_col, qty_col, factory_col, wip_col, remark_col] if c and c in month_df.columns]
    extra_cols = []
    for col in [order_amount_col, ship_amount_col, unit_price_col, hold_col, discount_col]:
        if col and col != "(無)" and col in month_df.columns and col not in detail_cols:
            extra_cols.append(col)
    st.markdown("**業績明細**")
    st.dataframe(month_df[detail_cols + extra_cols], use_container_width=True, height=420, hide_index=True)
