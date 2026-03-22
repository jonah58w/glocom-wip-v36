from __future__ import annotations

import re
from datetime import datetime
from typing import Any

import pandas as pd
import streamlit as st


def _normalize_month_text(text: str) -> str:
    text = (text or "").strip()
    m = re.match(r"^(\d{4})-(\d{1,2})$", text)
    if not m:
        return datetime.now().strftime("%Y-%m")
    y, mm = m.groups()
    return f"{y}-{int(mm):02d}"


def _safe_numeric_series(series: pd.Series) -> pd.Series:
    s = series.fillna("").astype(str)
    s = (
        s.str.replace(",", "", regex=False)
        .str.replace("US$", "", regex=False)
        .str.replace("USD", "", regex=False)
        .str.replace("NT$", "", regex=False)
        .str.replace("\u3000", "", regex=False)
        .str.replace(" ", "", regex=False)
    )
    s = s.str.extract(r"([-+]?[0-9]*\.?[0-9]+)", expand=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    lowered = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.strip().lower() in lowered:
            return lowered[cand.strip().lower()]
    for col in df.columns:
        c = str(col).strip().lower()
        for cand in candidates:
            if cand.strip().lower() in c:
                return col
    return None


def _series_text(df: pd.DataFrame, col: str | None) -> pd.Series:
    if not col or col not in df.columns:
        return pd.Series([""] * len(df), index=df.index)
    return df[col].fillna("").astype(str).str.strip()


def _date_series(df: pd.DataFrame, col: str | None) -> pd.Series:
    if not col or col not in df.columns:
        return pd.Series(pd.NaT, index=df.index)
    return pd.to_datetime(df[col], errors="coerce")


def _money(v: float, symbol: str = "US$") -> str:
    return f"{symbol} {v:,.0f}"


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
    st.caption("只統計所選月份；預設幣別為美金。")

    if source_df is None or not isinstance(source_df, pd.DataFrame) or source_df.empty:
        st.warning("目前沒有可統計資料")
        return

    work = source_df.copy()
    default_month = datetime.now().strftime("%Y-%m")

    col1, col2, col3 = st.columns(3)
    report_month_input = col1.text_input("報表月份 (YYYY-MM)", value=default_month)
    report_month = _normalize_month_text(report_month_input)
    col2.text_input("子表名稱 / 公司名稱", value="", key="sales_company_name")
    currency_symbol = col3.text_input("幣別符號", value="US$", key="sales_currency_symbol")

    date_candidates = [
        *( [ship_date_col] if ship_date_col else [] ),
        *( [order_date_col] if order_date_col else [] ),
        "Ship date", "Ship Date", "出貨日期", "日期", "Date", "Order Date", "下單日期",
    ]
    order_amount_candidates = [
        "接單金額", "金額(USD)", "Order Amount", "Order Amt", "Amount", "Sales Amount", "Total Amount",
        "USD Amount", "訂單金額", "Amount(USD)", "Order USD",
    ]
    ship_amount_candidates = [
        "出貨金額", "淨出貨金額", "Shipment Amount", "Ship Amount", "Net Shipment", "Invoice Amount",
        "銷貨金額", "Shipping Amount", "Net Amount",
    ]
    unit_price_candidates = [
        "單價", "Price", "Unit Price", "單價(USD)", "USD Price", "Price USD", "Selling Price",
    ]
    hold_candidates = ["HOLD", "HOLD金額", "Hold", "Hold Amount"]
    discount_candidates = ["折讓", "銷貨折讓", "Discount"]

    all_cols = ["(無)"] + [str(c) for c in work.columns]

    detected_date_col = _pick_col(work, date_candidates)
    detected_order_amount_col = _pick_col(work, order_amount_candidates)
    detected_ship_amount_col = _pick_col(work, ship_amount_candidates)
    detected_unit_price_col = _pick_col(work, unit_price_candidates)
    detected_hold_col = _pick_col(work, hold_candidates)
    detected_discount_col = _pick_col(work, discount_candidates)

    with st.expander("欄位偵測", expanded=False):
        c1, c2, c3 = st.columns(3)
        selected_date_col = c1.selectbox("日期欄位", all_cols, index=all_cols.index(str(detected_date_col)) if detected_date_col in work.columns else 0)
        selected_order_amount_col = c2.selectbox("接單金額欄位", all_cols, index=all_cols.index(str(detected_order_amount_col)) if detected_order_amount_col in work.columns else 0)
        selected_ship_amount_col = c3.selectbox("出貨金額欄位", all_cols, index=all_cols.index(str(detected_ship_amount_col)) if detected_ship_amount_col in work.columns else 0)
        c4, c5, c6 = st.columns(3)
        selected_unit_price_col = c4.selectbox("單價欄位", all_cols, index=all_cols.index(str(detected_unit_price_col)) if detected_unit_price_col in work.columns else 0)
        selected_hold_col = c5.selectbox("HOLD欄位", all_cols, index=all_cols.index(str(detected_hold_col)) if detected_hold_col in work.columns else 0)
        selected_discount_col = c6.selectbox("折讓欄位", all_cols, index=all_cols.index(str(detected_discount_col)) if detected_discount_col in work.columns else 0)

    date_col = None if selected_date_col == "(無)" else selected_date_col
    order_amount_col = None if selected_order_amount_col == "(無)" else selected_order_amount_col
    ship_amount_col = None if selected_ship_amount_col == "(無)" else selected_ship_amount_col
    unit_price_col = None if selected_unit_price_col == "(無)" else selected_unit_price_col
    hold_col = None if selected_hold_col == "(無)" else selected_hold_col
    discount_col = None if selected_discount_col == "(無)" else selected_discount_col

    customer_col = customer_col if customer_col in work.columns else _pick_col(work, ["客戶", "Customer", "Cust"])
    factory_col = factory_col if factory_col in work.columns else _pick_col(work, ["工廠", "Factory", "Vendor"])
    qty_col = qty_col if qty_col in work.columns else _pick_col(work, ["Order Q'TY (PCS)", "Qty", "Quantity", "數量"])

    if not date_col:
        st.warning("請在欄位偵測中選擇日期欄位")
        return

    work["_report_date"] = _date_series(work, date_col)
    work["_month"] = work["_report_date"].dt.strftime("%Y-%m")
    month_df = work[work["_month"] == report_month].copy()

    if month_df.empty:
        st.warning("所選月份沒有資料")
        return

    qty_series = _safe_numeric_series(month_df[qty_col]) if qty_col and qty_col in month_df.columns else pd.Series(0.0, index=month_df.index)
    order_amount_series = _safe_numeric_series(month_df[order_amount_col]) if order_amount_col and order_amount_col in month_df.columns else pd.Series(0.0, index=month_df.index)
    ship_amount_series = _safe_numeric_series(month_df[ship_amount_col]) if ship_amount_col and ship_amount_col in month_df.columns else pd.Series(0.0, index=month_df.index)
    unit_price_series = _safe_numeric_series(month_df[unit_price_col]) if unit_price_col and unit_price_col in month_df.columns else pd.Series(0.0, index=month_df.index)
    hold_series = _safe_numeric_series(month_df[hold_col]) if hold_col and hold_col in month_df.columns else pd.Series(0.0, index=month_df.index)
    discount_series = _safe_numeric_series(month_df[discount_col]) if discount_col and discount_col in month_df.columns else pd.Series(0.0, index=month_df.index)

    calc_from_unit = unit_price_series * qty_series
    if float(order_amount_series.sum()) == 0 and float(calc_from_unit.sum()) > 0:
        order_amount_series = calc_from_unit
    if float(ship_amount_series.sum()) == 0 and float(calc_from_unit.sum()) > 0:
        ship_amount_series = calc_from_unit

    net_ship_series = ship_amount_series - hold_series - discount_series

    month_df["_order_amount"] = order_amount_series
    month_df["_ship_amount"] = ship_amount_series
    month_df["_net_ship"] = net_ship_series

    customer_series = _series_text(month_df, customer_col) if customer_col else pd.Series([""] * len(month_df), index=month_df.index)
    factory_series = _series_text(month_df, factory_col) if factory_col else pd.Series([""] * len(month_df), index=month_df.index)

    total_order_amount = float(month_df["_order_amount"].sum())
    total_ship_amount = float(month_df["_ship_amount"].sum())
    total_net_ship = float(month_df["_net_ship"].sum())
    customer_count = int(customer_series.replace("", pd.NA).dropna().nunique())
    factory_count = int(factory_series.replace("", pd.NA).dropna().nunique())

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("接單金額", _money(total_order_amount, currency_symbol))
    m2.metric("出貨金額", _money(total_ship_amount, currency_symbol))
    m3.metric("淨出貨", _money(total_net_ship, currency_symbol))
    m4.metric("客戶數", customer_count)
    m5.metric("廠商數", factory_count)

    by_customer = pd.DataFrame({
        "客戶": customer_series,
        "出貨金額": month_df["_ship_amount"],
        "淨出貨": month_df["_net_ship"],
    })
    by_customer = by_customer.groupby("客戶", dropna=False)[["出貨金額", "淨出貨"]].sum().reset_index()
    by_customer = by_customer[by_customer["客戶"].astype(str).str.strip() != ""]
    by_customer = by_customer.sort_values("出貨金額", ascending=False)

    left, right = st.columns(2)
    with left:
        st.markdown("**客戶業績比較**")
        if by_customer.empty:
            st.info("無資料")
        else:
            st.bar_chart(by_customer.head(10).set_index("客戶")[["出貨金額"]])

    with right:
        st.markdown("**業績佔比**")
        if by_customer.empty:
            st.info("無資料")
        else:
            pct = by_customer.head(10).copy()
            total = float(pct["出貨金額"].sum())
            pct["佔比%"] = (pct["出貨金額"] / total * 100).round(2) if total else 0
            st.dataframe(pct[["客戶", "出貨金額", "佔比%"]], use_container_width=True, hide_index=True)

    show_cols = [c for c in [date_col, po_col, customer_col, part_col, qty_col, factory_col, wip_col, ship_date_col, remark_col] if c and c in month_df.columns]
    st.markdown("**業績明細**")
    show_df = month_df[show_cols].copy() if show_cols else month_df.copy()
    show_df.insert(len(show_df.columns), "接單金額", month_df["_order_amount"].values)
    show_df.insert(len(show_df.columns), "出貨金額", month_df["_ship_amount"].values)
    show_df.insert(len(show_df.columns), "淨出貨", month_df["_net_ship"].values)
    st.dataframe(show_df, use_container_width=True, hide_index=True, height=420)
