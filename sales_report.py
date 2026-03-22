# -*- coding: utf-8 -*-
from __future__ import annotations

import io
from datetime import datetime
from typing import Any, Optional

import pandas as pd
import streamlit as st


def _pick_col(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand is None:
            continue
        key = str(cand).strip().lower()
        if key in cols_lower:
            return cols_lower[key]
    for c in df.columns:
        cl = str(c).strip().lower()
        for cand in candidates:
            if cand and str(cand).strip().lower() in cl:
                return c
    return None



def _to_num(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=float)
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("US$", "", regex=False)
        .str.replace("USD", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("NT$", "", regex=False)
        .str.replace("\u00a0", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)



def _fmt_money(value: float, symbol: str = "US$") -> str:
    return f"{symbol} {value:,.0f}"



def _month_str(ts: Any) -> str:
    dt = pd.to_datetime(ts, errors="coerce")
    if pd.isna(dt):
        return ""
    return dt.strftime("%Y-%m")



def _build_excel(
    order_detail: pd.DataFrame,
    ship_detail: pd.DataFrame,
    by_customer: pd.DataFrame,
    report_month: str,
    currency_symbol: str,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cover = pd.DataFrame(
            {
                "項目": ["報表月份", "接單金額", "出貨金額", "客戶數", "廠商數"],
                "值": [
                    report_month,
                    _fmt_money(float(order_detail["接單金額"].sum()) if "接單金額" in order_detail.columns else 0, currency_symbol),
                    _fmt_money(float(ship_detail["出貨金額"].sum()) if "出貨金額" in ship_detail.columns else 0, currency_symbol),
                    int(by_customer["客戶"].nunique()) if not by_customer.empty else 0,
                    int(order_detail["工廠"].replace("", pd.NA).dropna().nunique()) if "工廠" in order_detail.columns else 0,
                ],
            }
        )
        cover.to_excel(writer, index=False, sheet_name="業績圖表總覽")
        order_detail.to_excel(writer, index=False, sheet_name="接單明細")
        ship_detail.to_excel(writer, index=False, sheet_name="出貨明細")
        by_customer.to_excel(writer, index=False, sheet_name="客戶彙總")
    output.seek(0)
    return output.getvalue()



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
    **kwargs,
):
    source_df = orders if isinstance(orders, pd.DataFrame) and not orders.empty else df

    st.subheader("業績明細表")
    st.caption("以所選月份統計全部客戶全部訂單，接單金額與出貨金額分開計算。")

    if source_df is None or not isinstance(source_df, pd.DataFrame) or source_df.empty:
        st.warning("目前沒有可用資料")
        return

    work = source_df.copy()

    detected_customer_col = customer_col if customer_col in work.columns else _pick_col(work, ["客戶", "customer", "customer name"])
    detected_factory_col = factory_col if factory_col in work.columns else _pick_col(work, ["工廠", "factory", "vendor", "supplier"])
    detected_order_date_col = order_date_col if order_date_col in work.columns else _pick_col(work, ["接單日期", "order date", "日期", "date", "merge date"])
    detected_ship_date_col = ship_date_col if ship_date_col in work.columns else _pick_col(work, ["出貨日期", "ship date", "shipment date", "shipping date", "日期", "date"])
    detected_order_amt_col = _pick_col(work, ["接單金額(usd)", "order amount usd", "order usd", "amount usd", "接單金額", "order amount", "amount"])
    detected_ship_amt_col = _pick_col(work, ["出貨金額(usd)", "ship amount usd", "shipment amount usd", "shipping usd", "出貨金額", "shipment amount", "ship amount"])

    now_month = datetime.now().strftime("%Y-%m")
    report_month = st.text_input("報表月份 (YYYY-MM)", value=now_month, key="sales_report_month")
    c1, c2 = st.columns(2)
    with c1:
        currency_symbol = st.text_input("幣別符號", value="US$", key="sales_currency_symbol")
    with c2:
        st.text_input("子表名稱 / 公司名稱", value="", key="sales_company_name")

    with st.expander("欄位偵測", expanded=True):
        cols = ["(無)"] + list(work.columns)
        def idx(v):
            return cols.index(v) if v in cols else 0

        col_a, col_b = st.columns(2)
        with col_a:
            selected_customer_col = st.selectbox("客戶欄位", cols, index=idx(detected_customer_col), key="sales_customer_col")
            selected_factory_col = st.selectbox("工廠欄位", cols, index=idx(detected_factory_col), key="sales_factory_col")
            selected_order_date_col = st.selectbox("接單日期欄位", cols, index=idx(detected_order_date_col), key="sales_order_date_col")
            selected_order_amt_col = st.selectbox("接單金額欄位", cols, index=idx(detected_order_amt_col), key="sales_order_amt_col")
        with col_b:
            selected_ship_date_col = st.selectbox("出貨日期欄位", cols, index=idx(detected_ship_date_col), key="sales_ship_date_col")
            selected_ship_amt_col = st.selectbox("出貨金額欄位", cols, index=idx(detected_ship_amt_col), key="sales_ship_amt_col")
            st.write(f"PO欄位：{po_col or '(無)'}")
            st.write(f"P/N欄位：{part_col or '(無)'}")

    customer_col = None if selected_customer_col == "(無)" else selected_customer_col
    factory_col = None if selected_factory_col == "(無)" else selected_factory_col
    order_date_col = None if selected_order_date_col == "(無)" else selected_order_date_col
    ship_date_col = None if selected_ship_date_col == "(無)" else selected_ship_date_col
    order_amt_col = None if selected_order_amt_col == "(無)" else selected_order_amt_col
    ship_amt_col = None if selected_ship_amt_col == "(無)" else selected_ship_amt_col

    work["_order_month"] = pd.to_datetime(work[order_date_col], errors="coerce").dt.strftime("%Y-%m") if order_date_col else ""
    work["_ship_month"] = pd.to_datetime(work[ship_date_col], errors="coerce").dt.strftime("%Y-%m") if ship_date_col else ""
    work["_order_amount"] = _to_num(work[order_amt_col]) if order_amt_col else 0.0
    work["_ship_amount"] = _to_num(work[ship_amt_col]) if ship_amt_col else 0.0

    order_df = work[work["_order_month"] == report_month].copy() if order_date_col else work.iloc[0:0].copy()
    ship_df = work[work["_ship_month"] == report_month].copy() if ship_date_col else work.iloc[0:0].copy()

    total_order_amount = float(order_df["_order_amount"].sum()) if not order_df.empty else 0.0
    total_ship_amount = float(ship_df["_ship_amount"].sum()) if not ship_df.empty else 0.0
    customer_count = int(pd.concat([
        order_df[customer_col] if customer_col and customer_col in order_df.columns else pd.Series(dtype=object),
        ship_df[customer_col] if customer_col and customer_col in ship_df.columns else pd.Series(dtype=object),
    ]).astype(str).replace("", pd.NA).dropna().nunique())
    factory_count = int(pd.concat([
        order_df[factory_col] if factory_col and factory_col in order_df.columns else pd.Series(dtype=object),
        ship_df[factory_col] if factory_col and factory_col in ship_df.columns else pd.Series(dtype=object),
    ]).astype(str).replace("", pd.NA).dropna().nunique())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("接單金額", _fmt_money(total_order_amount, currency_symbol))
    m2.metric("出貨金額", _fmt_money(total_ship_amount, currency_symbol))
    m3.metric("客戶數", customer_count)
    m4.metric("廠商數", factory_count)

    if customer_col and customer_col in work.columns:
        by_customer_order = (
            order_df.groupby(customer_col, dropna=False)["_order_amount"]
            .sum()
            .reset_index()
            .rename(columns={customer_col: "客戶", "_order_amount": "接單金額"})
        )
        by_customer_ship = (
            ship_df.groupby(customer_col, dropna=False)["_ship_amount"]
            .sum()
            .reset_index()
            .rename(columns={customer_col: "客戶", "_ship_amount": "出貨金額"})
        )
        by_customer = by_customer_order.merge(by_customer_ship, on="客戶", how="outer").fillna(0)
        by_customer = by_customer.sort_values(["接單金額", "出貨金額"], ascending=False)
    else:
        by_customer = pd.DataFrame(columns=["客戶", "接單金額", "出貨金額"])

    left, right = st.columns(2)
    with left:
        st.markdown("**客戶接單金額比較**")
        if by_customer.empty:
            st.info("本月沒有可統計資料")
        else:
            chart_df = by_customer.head(10).set_index("客戶")[["接單金額"]]
            st.bar_chart(chart_df)
    with right:
        st.markdown("**客戶出貨金額比較**")
        if by_customer.empty:
            st.info("本月沒有可統計資料")
        else:
            chart_df = by_customer.head(10).set_index("客戶")[["出貨金額"]]
            st.bar_chart(chart_df)

    order_detail_cols = [c for c in [order_date_col, po_col, customer_col, part_col, qty_col, factory_col, order_amt_col, remark_col] if c and c in order_df.columns]
    ship_detail_cols = [c for c in [ship_date_col, po_col, customer_col, part_col, qty_col, factory_col, ship_amt_col, wip_col, remark_col] if c and c in ship_df.columns]

    order_detail = order_df[order_detail_cols].copy() if order_detail_cols else pd.DataFrame()
    ship_detail = ship_df[ship_detail_cols].copy() if ship_detail_cols else pd.DataFrame()
    if order_amt_col and order_amt_col in order_detail.columns:
        order_detail = order_detail.rename(columns={order_amt_col: "接單金額"})
    if ship_amt_col and ship_amt_col in ship_detail.columns:
        ship_detail = ship_detail.rename(columns={ship_amt_col: "出貨金額"})
    if customer_col and customer_col in order_detail.columns:
        order_detail = order_detail.rename(columns={customer_col: "客戶"})
    if customer_col and customer_col in ship_detail.columns:
        ship_detail = ship_detail.rename(columns={customer_col: "客戶"})
    if factory_col and factory_col in order_detail.columns:
        order_detail = order_detail.rename(columns={factory_col: "工廠"})
    if factory_col and factory_col in ship_detail.columns:
        ship_detail = ship_detail.rename(columns={factory_col: "工廠"})

    t1, t2, t3 = st.tabs(["接單明細", "出貨明細", "客戶彙總"])
    with t1:
        st.dataframe(order_detail, use_container_width=True, height=420)
    with t2:
        st.dataframe(ship_detail, use_container_width=True, height=420)
    with t3:
        st.dataframe(by_customer, use_container_width=True, height=420, hide_index=True)

    excel_bytes = _build_excel(order_detail, ship_detail, by_customer, report_month, currency_symbol)
    st.download_button(
        "下載 Excel 業績明細表",
        data=excel_bytes,
        file_name=f"sales_report_{report_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
