from __future__ import annotations

import io
from datetime import datetime
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

MONEY_HINTS = [
    "淨出貨金額", "出貨金額", "銷貨金額", "業績金額", "開票金額", "invoice amount",
    "shipment amount", "ship amount", "sales amount", "revenue", "net amount", "net sales",
    "amount", "amt", "營業額", "金額", "invoice", "sales",
]
NET_HINTS = ["淨出貨金額", "net amount", "net sales", "淨額", "net"]
HOLD_HINTS = ["hold金額", "hold amount", "hold", "暫停金額"]
DISCOUNT_HINTS = ["折讓", "discount", "credit note", "allowance"]
DATE_HINTS = ["出貨日期", "ship date", "invoice date", "日期", "date", "開票日", "order date"]


def _norm(v: Any) -> str:
    return str(v or "").strip().lower().replace("_", " ")


def _safe_num(series: pd.Series | None) -> pd.Series:
    if series is None:
        return pd.Series(dtype=float)
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("NT$", "", regex=False)
        .str.replace("US$", "", regex=False)
        .str.replace("USD", "", regex=False)
        .str.replace("TWD", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.strip(),
        errors="coerce",
    ).fillna(0)


def _fmt_money(v: float, symbol: str = "NT$") -> str:
    return f"{symbol} {float(v):,.0f}"


def _candidate_columns(df: pd.DataFrame, hints: list[str]) -> list[str]:
    scored = []
    for col in df.columns:
        name = _norm(col)
        score = 0
        for h in hints:
            h2 = h.lower()
            if name == h2:
                score += 100
            elif h2 in name:
                score += 20
        num_sum = float(_safe_num(df[col]).abs().sum()) if len(df) else 0.0
        if score > 0 or num_sum > 0:
            scored.append((score, num_sum, col))
    scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return [x[2] for x in scored]


def _pick_date_col(df: pd.DataFrame, ship_date_col: str | None, order_date_col: str | None) -> str | None:
    for c in [ship_date_col, order_date_col]:
        if c and c in df.columns:
            return c
    cand = _candidate_columns(df, DATE_HINTS)
    return cand[0] if cand else None


def _build_excel(month_df: pd.DataFrame, show_cols: list[str], by_customer: pd.DataFrame, daily: pd.DataFrame,
                 report_month: str, company_name: str, currency_symbol: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "業績圖表總覽"
    ws["A1"] = f"{report_month} 業績報表"
    ws["A2"] = f"子表 {company_name}" if company_name else "全部客戶"
    ws["A1"].font = Font(size=18, bold=True)

    total_amount = float(month_df["_amount"].sum()) if "_amount" in month_df.columns else 0.0
    total_net = float(month_df["_net"].sum()) if "_net" in month_df.columns else 0.0
    customer_count = int(month_df["_customer"].replace("", pd.NA).dropna().nunique()) if "_customer" in month_df.columns else 0
    factory_count = int(month_df["_factory"].replace("", pd.NA).dropna().nunique()) if "_factory" in month_df.columns else 0

    ws["A4"], ws["C4"], ws["E4"], ws["G4"] = "總業績", "客戶數", "廠商數", "淨出貨"
    ws["A5"], ws["C5"], ws["E5"], ws["G5"] = total_amount, customer_count, factory_count, total_net
    ws["A5"].number_format = f'"{currency_symbol}" #,##0'
    ws["G5"].number_format = f'"{currency_symbol}" #,##0'

    if not by_customer.empty:
        ws["A8"], ws["B8"] = "客戶", "淨出貨金額"
        for i, (_, row) in enumerate(by_customer.iterrows(), start=9):
            ws[f"A{i}"] = row["客戶"]
            ws[f"B{i}"] = float(row["淨出貨金額"])
            ws[f"B{i}"].number_format = f'"{currency_symbol}" #,##0'
        data = Reference(ws, min_col=2, min_row=8, max_row=8 + len(by_customer))
        cats = Reference(ws, min_col=1, min_row=9, max_row=8 + len(by_customer))
        bar = BarChart(); bar.title = "客戶業績比較"; bar.height = 8; bar.width = 12; bar.legend = None
        bar.add_data(data, titles_from_data=True); bar.set_categories(cats); ws.add_chart(bar, "D8")
        donut = DoughnutChart(); donut.title = "業績佔比"; donut.holeSize = 55; donut.height = 8; donut.width = 10
        donut.add_data(data, titles_from_data=True); donut.set_categories(cats); ws.add_chart(donut, "Q8")

    if not daily.empty:
        start = 28
        ws[f"A{start}"], ws[f"B{start}"] = "日期", "淨出貨金額"
        for i, (_, row) in enumerate(daily.iterrows(), start=start + 1):
            ws[f"A{i}"] = row["日期"]
            ws[f"B{i}"] = float(row["淨出貨金額"])
            ws[f"B{i}"].number_format = f'"{currency_symbol}" #,##0'
        data = Reference(ws, min_col=2, min_row=start, max_row=start + len(daily))
        cats = Reference(ws, min_col=1, min_row=start + 1, max_row=start + len(daily))
        line = LineChart(); line.title = "每日淨出貨趨勢"; line.height = 8; line.width = 18
        line.add_data(data, titles_from_data=True); line.set_categories(cats); ws.add_chart(line, "D28")

    sh = wb.create_sheet("業績明細表")
    for c_idx, col in enumerate(show_cols, start=1):
        sh.cell(row=1, column=c_idx, value=col).font = Font(bold=True)
    for r_idx, (_, row) in enumerate(month_df.iterrows(), start=2):
        for c_idx, col in enumerate(show_cols, start=1):
            sh.cell(row=r_idx, column=c_idx, value="" if pd.isna(row.get(col)) else row.get(col))
    for i in range(1, len(show_cols) + 1):
        sh.column_dimensions[get_column_letter(i)].width = 16

    out = io.BytesIO(); wb.save(out); return out.getvalue()


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
    source_df = orders if orders is not None and not orders.empty else df
    st.subheader("業績明細表")
    st.caption("統計全部客戶全部訂單；可手動修正欄位對應。")

    if source_df is None or source_df.empty:
        st.warning("目前沒有可用資料")
        return

    work = source_df.copy()
    default_month = datetime.now().strftime("%Y-%m")
    if ship_date_col and ship_date_col in work.columns:
        dt = pd.to_datetime(work[ship_date_col], errors="coerce")
        if dt.notna().any():
            default_month = dt.dropna().max().strftime("%Y-%m")
    elif order_date_col and order_date_col in work.columns:
        dt = pd.to_datetime(work[order_date_col], errors="coerce")
        if dt.notna().any():
            default_month = dt.dropna().max().strftime("%Y-%m")

    c1, c2, c3 = st.columns(3)
    report_month = c1.text_input("報表月份 (YYYY-MM)", value=default_month)
    company_name = c2.text_input("子表名稱 / 公司名稱", value="")
    currency_symbol = c3.text_input("幣別符號", value="NT$")

    date_candidates = ["(不使用)"] + ([ship_date_col] if ship_date_col and ship_date_col in work.columns else []) + ([order_date_col] if order_date_col and order_date_col in work.columns else [])
    for c in _candidate_columns(work, DATE_HINTS):
        if c not in date_candidates:
            date_candidates.append(c)

    money_candidates = ["(自動偵測)"] + [c for c in _candidate_columns(work, MONEY_HINTS) if c]
    net_candidates = ["(自動偵測/計算)"] + [c for c in _candidate_columns(work, NET_HINTS) if c]
    hold_candidates = ["(無)"] + [c for c in _candidate_columns(work, HOLD_HINTS) if c]
    discount_candidates = ["(無)"] + [c for c in _candidate_columns(work, DISCOUNT_HINTS) if c]

    with st.expander("欄位偵測", expanded=True):
        s1, s2, s3, s4, s5 = st.columns(5)
        date_pick = s1.selectbox("日期欄位", date_candidates, index=1 if len(date_candidates) > 1 else 0)
        amount_pick = s2.selectbox("總業績欄位", money_candidates, index=0)
        net_pick = s3.selectbox("淨出貨欄位", net_candidates, index=0)
        hold_pick = s4.selectbox("HOLD欄位", hold_candidates, index=0)
        discount_pick = s5.selectbox("折讓欄位", discount_candidates, index=0)
        st.caption(f"客戶欄位：{customer_col or '-'} / 工廠欄位：{factory_col or '-'}")

    date_col = None if date_pick == "(不使用)" else date_pick
    if not date_col:
        date_col = _pick_date_col(work, ship_date_col, order_date_col)

    amount_col = None if amount_pick == "(自動偵測)" else amount_pick
    if not amount_col:
        tmp = _candidate_columns(work, MONEY_HINTS)
        amount_col = tmp[0] if tmp else None

    net_col = None if net_pick == "(自動偵測/計算)" else net_pick
    if not net_col:
        tmp = [c for c in _candidate_columns(work, NET_HINTS) if c != amount_col]
        net_col = tmp[0] if tmp else None

    hold_col = None if hold_pick == "(無)" else hold_pick
    if hold_col is None:
        tmp = [c for c in _candidate_columns(work, HOLD_HINTS) if c not in [amount_col, net_col]]
        hold_col = tmp[0] if tmp else None

    discount_col = None if discount_pick == "(無)" else discount_pick
    if discount_col is None:
        tmp = [c for c in _candidate_columns(work, DISCOUNT_HINTS) if c not in [amount_col, net_col, hold_col]]
        discount_col = tmp[0] if tmp else None

    if date_col and date_col in work.columns:
        work["_report_date"] = pd.to_datetime(work[date_col], errors="coerce")
        work["_month"] = work["_report_date"].dt.strftime("%Y-%m")
        month_df = work[work["_month"] == report_month].copy()
        if month_df.empty:
            month_df = work.copy()
            st.info(f"{report_month} 沒有資料，已改為統計全部資料。")
    else:
        month_df = work.copy()
        st.info("未使用日期篩選，改為統計全部資料。")

    month_df["_customer"] = month_df[customer_col].astype(str).str.strip() if customer_col and customer_col in month_df.columns else ""
    month_df["_factory"] = month_df[factory_col].astype(str).str.strip() if factory_col and factory_col in month_df.columns else ""
    month_df["_amount"] = _safe_num(month_df[amount_col]) if amount_col and amount_col in month_df.columns else 0
    month_df["_hold"] = _safe_num(month_df[hold_col]) if hold_col and hold_col in month_df.columns else 0
    month_df["_discount"] = _safe_num(month_df[discount_col]) if discount_col and discount_col in month_df.columns else 0
    if net_col and net_col in month_df.columns:
        month_df["_net"] = _safe_num(month_df[net_col])
    else:
        month_df["_net"] = month_df["_amount"] - month_df["_hold"] - month_df["_discount"]

    total_amount = float(month_df["_amount"].sum())
    total_net = float(month_df["_net"].sum())
    customer_count = int(month_df["_customer"].replace("", pd.NA).dropna().nunique())
    factory_count = int(month_df["_factory"].replace("", pd.NA).dropna().nunique())

    by_customer = pd.DataFrame(columns=["客戶", "淨出貨金額", "佔比%"])
    if customer_col and customer_col in month_df.columns:
        by_customer = (
            month_df.groupby(customer_col)["_net"].sum().sort_values(ascending=False).head(10).reset_index()
        )
        by_customer.columns = ["客戶", "淨出貨金額"]
        total = float(by_customer["淨出貨金額"].sum())
        by_customer["佔比%"] = by_customer["淨出貨金額"].div(total).mul(100).round(2) if total else 0.0

    daily = pd.DataFrame(columns=["日期", "淨出貨金額"])
    if "_report_date" in month_df.columns and month_df["_report_date"].notna().any():
        daily = (
            month_df.groupby(month_df["_report_date"].dt.strftime("%Y-%m-%d"))["_net"].sum().reset_index()
        )
        daily.columns = ["日期", "淨出貨金額"]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("總業績", _fmt_money(total_amount, currency_symbol))
    k2.metric("客戶數", customer_count)
    k3.metric("廠商數", factory_count)
    k4.metric("淨出貨", _fmt_money(total_net, currency_symbol))

    l1, l2 = st.columns(2)
    with l1:
        st.markdown("**客戶業績比較**")
        if not by_customer.empty:
            st.bar_chart(by_customer.set_index("客戶")[["淨出貨金額"]])
        else:
            st.info("無客戶彙總資料")
    with l2:
        st.markdown("**業績佔比**")
        st.dataframe(by_customer, use_container_width=True, hide_index=True)

    st.markdown("**每日趨勢**")
    if not daily.empty:
        st.line_chart(daily.set_index("日期"))
    else:
        st.info("無日期趨勢資料")

    show_cols = [
        c for c in [order_date_col, ship_date_col, customer_col, factory_col, po_col, part_col, qty_col,
                    amount_col, hold_col, discount_col, net_col, remark_col, wip_col]
        if c and c in month_df.columns
    ]
    st.markdown("**業績明細表**")
    st.dataframe(month_df[show_cols] if show_cols else month_df, use_container_width=True, hide_index=True, height=420)

    excel_bytes = _build_excel(month_df, show_cols or list(month_df.columns), by_customer[["客戶", "淨出貨金額"]] if not by_customer.empty else by_customer,
                               daily, report_month, company_name, currency_symbol)
    st.download_button(
        "下載 Excel 業績圖表報表",
        data=excel_bytes,
        file_name=f"sales_report_{report_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
