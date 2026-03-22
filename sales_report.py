# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from datetime import datetime
from typing import Any

import pandas as pd
import streamlit as st


AMOUNT_TOKENS = [
    "amount", "amt", "金額", "總額", "usd", "us$", "sales", "revenue", "invoice", "價值", "貨款"
]
ORDER_AMOUNT_TOKENS = [
    "order amount", "order amt", "接單金額", "接單總額", "訂單金額", "訂單總額", "sales amount", "booking", "bookings"
]
SHIP_AMOUNT_TOKENS = [
    "shipment amount", "ship amount", "出貨金額", "出貨總額", "invoice amount", "銷貨金額", "淨出貨", "net shipment", "net sales"
]
UNIT_PRICE_TOKENS = [
    "unit price", "price", "單價", "usd/pcs", "usd/pc", "單價(usd)", "price usd", "pcs price"
]
HOLD_TOKENS = ["hold", "hold金額", "hold amount"]
DISCOUNT_TOKENS = ["discount", "折讓", "折讓金額"]
DATE_TOKENS = ["ship date", "date", "出貨日期", "日期", "invoice date", "order date"]


def _normalize_month_text(text: str) -> str:
    text = (text or "").strip()
    m = re.match(r"^(\d{4})-(\d{1,2})$", text)
    if not m:
        return datetime.now().strftime("%Y-%m")
    y, mm = m.groups()
    return f"{y}-{int(mm):02d}"


def _safe_series(df: pd.DataFrame, col: str | None) -> pd.Series | None:
    if not col or col not in df.columns:
        return None
    return df[col]


def _to_num(s: pd.Series | None) -> pd.Series:
    if s is None:
        return pd.Series(dtype="float64")
    text = s.astype(str).str.replace(",", "", regex=False)
    text = text.str.replace("US$", "", regex=False)
    text = text.str.replace("USD", "", regex=False)
    text = text.str.replace("NT$", "", regex=False)
    text = text.str.replace("$", "", regex=False)
    return pd.to_numeric(text, errors="coerce").fillna(0.0)


def _match_score(name: str, tokens: list[str]) -> int:
    lower = str(name).strip().lower()
    return sum(1 for t in tokens if t in lower)


def _pick_best_col(columns: list[str], preferred_tokens: list[str], fallback_tokens: list[str] | None = None) -> str | None:
    scored = []
    for c in columns:
        score = _match_score(c, preferred_tokens)
        if fallback_tokens:
            score = max(score, _match_score(c, fallback_tokens))
        if score > 0:
            scored.append((score, c))
    if not scored:
        return None
    scored.sort(key=lambda x: (-x[0], x[1]))
    return scored[0][1]


def _all_column_options(df: pd.DataFrame) -> list[str]:
    return ["(無)"] + [str(c) for c in df.columns]


def _default_index(options: list[str], value: str | None) -> int:
    if value and value in options:
        return options.index(value)
    return 0


def _select_col(label: str, options: list[str], default_value: str | None, key: str) -> str | None:
    idx = _default_index(options, default_value)
    selected = st.selectbox(label, options, index=idx, key=key)
    return None if selected == "(無)" else selected


def render_sales_report_page(
    df: pd.DataFrame | None = None,
    orders: pd.DataFrame | None = None,
    po_col: str | None = None,
    customer_col: str | None = None,
    part_col: str | None = None,
    qty_col: str | None = None,
    factory_col: str | None = None,
    wip_col: str | None = None,
    ship_date_col: str | None = None,
    remark_col: str | None = None,
    order_date_col: str | None = None,
    **kwargs: Any,
):
    source_df = orders if isinstance(orders, pd.DataFrame) and not orders.empty else df

    st.subheader("業績明細表")
    st.caption("只統計所選月份；預設幣別為美金。")

    if source_df is None or source_df.empty:
        st.warning("目前沒有可用資料")
        return

    work = source_df.copy()
    columns = [str(c) for c in work.columns]

    auto_date_col = ship_date_col if ship_date_col in work.columns else _pick_best_col(columns, DATE_TOKENS)
    auto_order_amount_col = _pick_best_col(columns, ORDER_AMOUNT_TOKENS, AMOUNT_TOKENS)
    auto_ship_amount_col = _pick_best_col(columns, SHIP_AMOUNT_TOKENS, AMOUNT_TOKENS)
    auto_unit_price_col = _pick_best_col(columns, UNIT_PRICE_TOKENS)
    auto_hold_col = _pick_best_col(columns, HOLD_TOKENS)
    auto_discount_col = _pick_best_col(columns, DISCOUNT_TOKENS)

    current_month = datetime.now().strftime("%Y-%m")
    c1, c2, c3 = st.columns(3)
    report_month = _normalize_month_text(c1.text_input("報表月份 (YYYY-MM)", value=current_month))
    company_name = c2.text_input("子表名稱 / 公司名稱", value="")
    currency_symbol = c3.text_input("幣別符號", value="US$")

    with st.expander("欄位偵測", expanded=True):
        opts = _all_column_options(work)
        d1, d2, d3 = st.columns(3)
        date_col = _select_col("日期欄位", opts, auto_date_col, "sales_date_col")
        order_amount_col = _select_col("接單金額欄位", opts, auto_order_amount_col, "sales_order_amount_col")
        ship_amount_col = _select_col("出貨金額欄位", opts, auto_ship_amount_col, "sales_ship_amount_col")
        d4, d5, d6 = st.columns(3)
        unit_price_col = _select_col("單價欄位", opts, auto_unit_price_col, "sales_unit_price_col")
        hold_col = _select_col("HOLD欄位", opts, auto_hold_col, "sales_hold_col")
        discount_col = _select_col("折讓欄位", opts, auto_discount_col, "sales_discount_col")

    if not date_col:
        st.error("找不到日期欄位")
        return

    work["_report_date"] = pd.to_datetime(work[date_col], errors="coerce")
    work["_month"] = work["_report_date"].dt.strftime("%Y-%m")
    month_df = work[work["_month"] == report_month].copy()

    qty_series = _to_num(_safe_series(month_df, qty_col)) if qty_col else pd.Series([0.0] * len(month_df), index=month_df.index)
    unit_price_series = _to_num(_safe_series(month_df, unit_price_col)) if unit_price_col else pd.Series([0.0] * len(month_df), index=month_df.index)
    order_amount_series = _to_num(_safe_series(month_df, order_amount_col)) if order_amount_col else pd.Series([0.0] * len(month_df), index=month_df.index)
    ship_amount_series = _to_num(_safe_series(month_df, ship_amount_col)) if ship_amount_col else pd.Series([0.0] * len(month_df), index=month_df.index)
    hold_series = _to_num(_safe_series(month_df, hold_col)) if hold_col else pd.Series([0.0] * len(month_df), index=month_df.index)
    discount_series = _to_num(_safe_series(month_df, discount_col)) if discount_col else pd.Series([0.0] * len(month_df), index=month_df.index)

    if order_amount_series.sum() == 0 and unit_price_col and qty_col:
        order_amount_series = unit_price_series * qty_series
    if ship_amount_series.sum() == 0 and unit_price_col and qty_col:
        ship_amount_series = unit_price_series * qty_series

    month_df["_order_amount"] = order_amount_series
    month_df["_ship_amount"] = ship_amount_series
    month_df["_net_amount"] = ship_amount_series - hold_series - discount_series

    customer_count = 0
    factory_count = 0
    if customer_col and customer_col in month_df.columns:
        customer_count = month_df[customer_col].astype(str).str.strip().replace("", pd.NA).dropna().nunique()
    if factory_col and factory_col in month_df.columns:
        factory_count = month_df[factory_col].astype(str).str.strip().replace("", pd.NA).dropna().nunique()

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("接單金額", f"{currency_symbol} {month_df['_order_amount'].sum():,.0f}")
    k2.metric("出貨金額", f"{currency_symbol} {month_df['_ship_amount'].sum():,.0f}")
    k3.metric("淨出貨", f"{currency_symbol} {month_df['_net_amount'].sum():,.0f}")
    k4.metric("客戶數", int(customer_count))
    k5.metric("廠商數", int(factory_count))

    if customer_col and customer_col in month_df.columns:
        by_customer = (
            month_df.groupby(customer_col)["_ship_amount"]
            .sum()
            .sort_values(ascending=False)
            .head(10)
        )
        left, right = st.columns(2)
        with left:
            st.markdown("**客戶業績比較**")
            if by_customer.empty:
                st.info("本月無金額資料")
            else:
                st.bar_chart(by_customer)
        with right:
            st.markdown("**業績佔比**")
            if by_customer.empty:
                st.info("本月無金額資料")
            else:
                ratio_df = by_customer.reset_index()
                ratio_df.columns = ["客戶", "出貨金額"]
                total_ship = ratio_df["出貨金額"].sum()
                ratio_df["佔比%"] = (ratio_df["出貨金額"] / total_ship * 100).round(2) if total_ship else 0
                st.dataframe(ratio_df, use_container_width=True, hide_index=True)

    detail_cols = [c for c in [date_col, po_col, customer_col, part_col, qty_col, factory_col, wip_col, remark_col] if c and c in month_df.columns]
    detail = month_df[detail_cols].copy() if detail_cols else pd.DataFrame(index=month_df.index)
    detail["接單金額"] = month_df["_order_amount"].values
    detail["出貨金額"] = month_df["_ship_amount"].values
    detail["淨出貨"] = month_df["_net_amount"].values

    st.markdown("**業績明細**")
    st.dataframe(detail, use_container_width=True, height=420, hide_index=True)
