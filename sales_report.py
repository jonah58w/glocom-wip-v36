from __future__ import annotations

import io
from datetime import datetime
from typing import Any, Optional

import pandas as pd
import streamlit as st

try:
    import utils as u
except Exception:
    u = None


def _series(df: pd.DataFrame, col: Optional[str]):
    if not col or col not in df.columns:
        return None
    if u and hasattr(u, "get_series_by_col"):
        s = u.get_series_by_col(df, col)
        if s is not None:
            return s
    return df[col]


def _pick_col(df: pd.DataFrame, candidates: list[str], preferred: Optional[str] = None) -> Optional[str]:
    if preferred and preferred in df.columns:
        return preferred
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for c in candidates:
        if c.strip().lower() in lower_map:
            return lower_map[c.strip().lower()]
    for real in df.columns:
        real_l = str(real).strip().lower()
        for cand in candidates:
            if cand.strip().lower() in real_l:
                return real
    return None


def _safe_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("NT$", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(" ", "", regex=False),
        errors="coerce",
    ).fillna(0)


def _fmt_money(value: float, symbol: str = "NT$") -> str:
    return f"{symbol}{value:,.0f}"


def _build_excel_bytes(df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="業績明細表", index=False)
    return out.getvalue()


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
    st.caption("沿用目前 WIP 主資料，統計所有客戶所有訂單金額。")

    if source_df is None or not isinstance(source_df, pd.DataFrame) or source_df.empty:
        st.warning("目前沒有可用資料")
        return

    work = source_df.copy()

    detected_customer_col = _pick_col(work, ["客戶", "customer", "客戶名稱"], customer_col)
    detected_factory_col = _pick_col(work, ["工廠", "factory", "vendor", "廠商"], factory_col)
    detected_date_col = _pick_col(work, ["日期", "date", "出貨日期", "ship date", "order date", "交期"], ship_date_col or order_date_col)
    detected_amount_col = _pick_col(work, ["金額", "amount", "出貨金額", "sales amount", "shipment amount", "未稅金額", "銷貨金額", "總金額"])
    detected_hold_col = _pick_col(work, ["hold金額", "hold amount", "hold", "保留金額"])
    detected_discount_col = _pick_col(work, ["折讓", "discount", "折扣", "discount amount"])
    detected_net_col = _pick_col(work, ["淨出貨金額", "net amount", "net", "實際金額", "淨額"])

    report_month = st.text_input("報表月份 (YYYY-MM)", value=datetime.now().strftime("%Y-%m"))
    company_name = st.text_input("子表名稱 / 公司名稱", value="")
    currency_symbol = st.text_input("幣別符號", value="NT$")

    with st.expander("欄位偵測", expanded=False):
        cols = [""] + list(work.columns)
        customer_col_use = st.selectbox("客戶欄位", cols, index=cols.index(detected_customer_col) if detected_customer_col in cols else 0)
        factory_col_use = st.selectbox("工廠欄位", cols, index=cols.index(detected_factory_col) if detected_factory_col in cols else 0)
        date_col_use = st.selectbox("日期欄位", cols, index=cols.index(detected_date_col) if detected_date_col in cols else 0)
        amount_col_use = st.selectbox("總業績欄位", cols, index=cols.index(detected_amount_col) if detected_amount_col in cols else 0)
        hold_col_use = st.selectbox("HOLD 金額欄位", cols, index=cols.index(detected_hold_col) if detected_hold_col in cols else 0)
        discount_col_use = st.selectbox("折讓欄位", cols, index=cols.index(detected_discount_col) if detected_discount_col in cols else 0)
        net_col_use = st.selectbox("淨出貨欄位", cols, index=cols.index(detected_net_col) if detected_net_col in cols else 0)

    customer_col_use = customer_col_use or None
    factory_col_use = factory_col_use or None
    date_col_use = date_col_use or None
    amount_col_use = amount_col_use or None
    hold_col_use = hold_col_use or None
    discount_col_use = discount_col_use or None
    net_col_use = net_col_use or None

    if date_col_use and date_col_use in work.columns:
        work["_report_date"] = pd.to_datetime(_series(work, date_col_use), errors="coerce")
        work["_month"] = work["_report_date"].dt.strftime("%Y-%m")
        month_df = work[work["_month"] == report_month].copy()
        if month_df.empty:
            month_df = work.copy()
            st.info("指定月份沒有資料，已改為顯示全部資料。")
    else:
        month_df = work.copy()
        st.info("找不到日期欄位，已改為統計全部資料。")

    month_df["_amount"] = _safe_num(_series(month_df, amount_col_use)) if amount_col_use else 0
    month_df["_hold"] = _safe_num(_series(month_df, hold_col_use)) if hold_col_use else 0
    month_df["_discount"] = _safe_num(_series(month_df, discount_col_use)) if discount_col_use else 0

    if net_col_use:
        month_df["_net"] = _safe_num(_series(month_df, net_col_use))
    else:
        month_df["_net"] = month_df["_amount"] - month_df["_hold"] - month_df["_discount"]

    if customer_col_use and customer_col_use in month_df.columns:
        customer_series = _series(month_df, customer_col_use).astype(str).str.strip()
    else:
        customer_series = pd.Series(["(blank)"] * len(month_df), index=month_df.index)

    if factory_col_use and factory_col_use in month_df.columns:
        factory_series = _series(month_df, factory_col_use).astype(str).str.strip()
    else:
        factory_series = pd.Series(["(blank)"] * len(month_df), index=month_df.index)

    total_amount = float(month_df["_amount"].sum())
    total_net = float(month_df["_net"].sum())
    customer_count = int(customer_series.replace("", pd.NA).dropna().nunique())
    factory_count = int(factory_series.replace("", pd.NA).dropna().nunique())

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("總業績", _fmt_money(total_amount, currency_symbol))
    k2.metric("客戶數", customer_count)
    k3.metric("廠商數", factory_count)
    k4.metric("淨出貨", _fmt_money(total_net, currency_symbol))

    by_customer = (
        pd.DataFrame({"客戶": customer_series, "淨出貨金額": month_df["_net"]})
        .groupby("客戶", dropna=False)["淨出貨金額"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    by_factory = (
        pd.DataFrame({"工廠": factory_series, "淨出貨金額": month_df["_net"]})
        .groupby("工廠", dropna=False)["淨出貨金額"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    left, right = st.columns(2)
    with left:
        st.markdown("**客戶業績比較**")
        if by_customer.empty:
            st.write("無資料")
        else:
            st.bar_chart(by_customer.head(10).set_index("客戶"))
    with right:
        st.markdown("**業績佔比**")
        if by_customer.empty:
            st.write("無資料")
        else:
            pct_df = by_customer.head(10).copy()
            total = float(pct_df["淨出貨金額"].sum())
            pct_df["佔比%"] = 0 if total == 0 else (pct_df["淨出貨金額"] / total * 100).round(2)
            pct_df["淨出貨金額"] = pct_df["淨出貨金額"].map(lambda x: _fmt_money(x, currency_symbol))
            st.dataframe(pct_df, use_container_width=True, hide_index=True)

    if "_report_date" in month_df.columns and month_df["_report_date"].notna().any():
        daily = month_df.groupby(month_df["_report_date"].dt.strftime("%Y-%m-%d"))["_net"].sum().reset_index()
        daily.columns = ["日期", "淨出貨金額"]
        st.markdown("**每日趨勢**")
        st.line_chart(daily.set_index("日期"))

    tabs = st.tabs(["業績明細表", "客戶/工廠彙總", "RawData"])
    with tabs[0]:
        show_cols = [
            c for c in [date_col_use, customer_col_use, factory_col_use, po_col, part_col, qty_col, amount_col_use, hold_col_use, discount_col_use, net_col_use, wip_col, remark_col]
            if c and c in month_df.columns
        ]
        if not show_cols:
            show_cols = list(month_df.columns)
        st.dataframe(month_df[show_cols], use_container_width=True, hide_index=True)
        st.download_button(
            "下載業績明細 Excel",
            data=_build_excel_bytes(month_df[show_cols]),
            file_name=f"sales_report_{report_month or 'all'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**客戶彙總**")
            st.dataframe(by_customer, use_container_width=True, hide_index=True)
        with c2:
            st.markdown("**工廠彙總**")
            st.dataframe(by_factory, use_container_width=True, hide_index=True)
    with tabs[2]:
        st.dataframe(source_df, use_container_width=True, hide_index=True)

    st.caption(
        f"目前統計欄位：日期={date_col_use or '-'} / 總業績={amount_col_use or '-'} / HOLD={hold_col_use or '-'} / 折讓={discount_col_use or '-'} / 淨出貨={net_col_use or '自動計算'}"
    )
