from __future__ import annotations

from datetime import datetime
import io
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


def _pick_col(df, candidates):
    for c in candidates:
        if c and c in df.columns:
            return c
    return None


def _safe_num(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=float)
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.replace("NT$", "", regex=False).str.replace("$", "", regex=False),
        errors="coerce",
    ).fillna(0.0)


def _fmt_money(value: float, currency_symbol: str = "NT$") -> str:
    return f"{currency_symbol} {value:,.0f}"


def _build_excel(month_df, show_cols, by_customer, daily, report_month, company_name, currency_symbol):
    wb = Workbook()
    ws = wb.active
    ws.title = "業績圖表總覽"

    ws["A1"] = f"{report_month} 業績報表"
    ws["A2"] = f"子表 {company_name}" if company_name else "子表"
    ws["A1"].font = Font(size=18, bold=True)
    ws["A2"].font = Font(size=11, bold=False)

    total_amount = float(month_df["_amount"].sum()) if "_amount" in month_df.columns else 0
    total_net = float(month_df["_net"].sum()) if "_net" in month_df.columns else 0
    customer_count = int(month_df["_customer_name"].replace("", pd.NA).dropna().nunique()) if "_customer_name" in month_df.columns else 0
    factory_count = int(month_df["_factory_name"].replace("", pd.NA).dropna().nunique()) if "_factory_name" in month_df.columns else 0

    labels = [("A4", "總業績"), ("C4", "客戶數"), ("E4", "廠商數"), ("G4", "淨出貨")]
    for cell, label in labels:
        ws[cell] = label
        ws[cell].font = Font(bold=True)

    ws["A5"] = total_amount
    ws["C5"] = customer_count
    ws["E5"] = factory_count
    ws["G5"] = total_net
    ws["A5"].number_format = f'"{currency_symbol}" #,##0'
    ws["G5"].number_format = f'"{currency_symbol}" #,##0'

    ws["A8"] = "客戶業績比較"
    ws["F8"] = "業績佔比"
    ws["A8"].font = Font(bold=True)
    ws["F8"].font = Font(bold=True)

    top_customer = by_customer.head(10).copy()
    if top_customer.empty:
        top_customer = pd.DataFrame({"客戶": ["N/A"], "淨出貨金額": [0.0]})

    ws["A9"] = "客戶"
    ws["B9"] = "淨出貨金額"
    for i, (_, row) in enumerate(top_customer.iterrows(), start=10):
        ws[f"A{i}"] = row["客戶"]
        ws[f"B{i}"] = float(row["淨出貨金額"])
        ws[f"B{i}"].number_format = f'"{currency_symbol}" #,##0'

    bar = BarChart()
    bar.title = "客戶業績比較"
    data = Reference(ws, min_col=2, min_row=9, max_row=9 + len(top_customer))
    cats = Reference(ws, min_col=1, min_row=10, max_row=9 + len(top_customer))
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 8
    bar.width = 12
    bar.legend = None
    ws.add_chart(bar, "A11")

    pie = PieChart()
    pie.title = "業績佔比"
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(cats)
    pie.height = 8
    pie.width = 10
    ws.add_chart(pie, "F11")

    start_row = 30
    ws[f"A{start_row}"] = "每日趨勢"
    ws[f"A{start_row}"].font = Font(bold=True)
    ws[f"A{start_row+1}"] = "日期"
    ws[f"B{start_row+1}"] = "淨出貨金額"
    if daily.empty:
        daily = pd.DataFrame({"日期": ["N/A"], "淨出貨金額": [0.0]})
    for i, (_, row) in enumerate(daily.iterrows(), start=start_row + 2):
        ws[f"A{i}"] = row["日期"]
        ws[f"B{i}"] = float(row["淨出貨金額"])
        ws[f"B{i}"].number_format = f'"{currency_symbol}" #,##0'

    line = LineChart()
    line.title = "每日淨出貨趨勢"
    d2 = Reference(ws, min_col=2, min_row=start_row + 1, max_row=start_row + 1 + len(daily))
    c2 = Reference(ws, min_col=1, min_row=start_row + 2, max_row=start_row + 1 + len(daily))
    line.add_data(d2, titles_from_data=True)
    line.set_categories(c2)
    line.height = 8
    line.width = 14
    ws.add_chart(line, "D30")

    detail = wb.create_sheet("業績明細表")
    for idx, col in enumerate(show_cols, start=1):
        detail.cell(row=1, column=idx, value=col).font = Font(bold=True)
    for r, (_, row) in enumerate(month_df.iterrows(), start=2):
        for c, col in enumerate(show_cols, start=1):
            detail.cell(row=r, column=c, value="" if pd.isna(row[col]) else row[col])
    for col_idx in range(1, len(show_cols) + 1):
        detail.column_dimensions[get_column_letter(col_idx)].width = 18

    def _write_df_sheet(name, df):
        sh = wb.create_sheet(name)
        if df.empty:
            sh["A1"] = "無資料"
            return
        for c, col in enumerate(df.columns, start=1):
            sh.cell(row=1, column=c, value=col).font = Font(bold=True)
        for r, (_, row) in enumerate(df.iterrows(), start=2):
            for c, col in enumerate(df.columns, start=1):
                sh.cell(row=r, column=c, value="" if pd.isna(row[col]) else row[col])
        for col_idx in range(1, len(df.columns) + 1):
            sh.column_dimensions[get_column_letter(col_idx)].width = 18

    _write_df_sheet("客戶彙總", by_customer)
    _write_df_sheet("每日明細", daily)
    _write_df_sheet("RawData", month_df)

    out = io.BytesIO()
    wb.save(out)
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

    c1, c2, c3 = st.columns(3)
    with c1:
        report_month = st.text_input("報表月份 (YYYY-MM)", value=datetime.now().strftime("%Y-%m"))
    with c2:
        company_name = st.text_input("子表名稱 / 公司名稱", value="WESCO")
    with c3:
        currency_symbol = st.text_input("幣別符號", value="NT$")

    date_col = order_date_col or ship_date_col
    if not date_col:
        st.error("找不到日期欄位，至少要有 日期 / 出貨日期 / Ship Date 其中一欄")
        return

    work["_report_date"] = pd.to_datetime(work[date_col], errors="coerce")
    work["_month"] = work["_report_date"].dt.strftime("%Y-%m")
    work["_customer_name"] = work[customer_col].astype(str).fillna("").str.strip() if customer_col and customer_col in work.columns else ""
    work["_factory_name"] = work[factory_col].astype(str).fillna("").str.strip() if factory_col and factory_col in work.columns else ""

    month_df = work[work["_month"] == report_month].copy()

    month_df["_amount"] = _safe_num(month_df[amount_col]) if amount_col else 0.0
    month_df["_hold"] = _safe_num(month_df[hold_col]) if hold_col else 0.0
    month_df["_discount"] = _safe_num(month_df[discount_col]) if discount_col else 0.0
    month_df["_net"] = _safe_num(month_df[net_col]) if net_col else month_df["_amount"] - month_df["_hold"] - month_df["_discount"]

    total_amount = float(month_df["_amount"].sum())
    total_net = float(month_df["_net"].sum())
    customer_count = month_df["_customer_name"].replace("", pd.NA).dropna().nunique() if "_customer_name" in month_df.columns else 0
    factory_count = month_df["_factory_name"].replace("", pd.NA).dropna().nunique() if "_factory_name" in month_df.columns else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("總業績", _fmt_money(total_amount, currency_symbol))
    m2.metric("客戶數", int(customer_count))
    m3.metric("廠商數", int(factory_count))
    m4.metric("淨出貨", _fmt_money(total_net, currency_symbol))

    by_customer = pd.DataFrame(columns=["客戶", "淨出貨金額"])
    if customer_col and customer_col in month_df.columns:
        by_customer = (
            month_df.groupby(customer_col)["_net"]
            .sum()
            .reset_index()
            .rename(columns={customer_col: "客戶", "_net": "淨出貨金額"})
            .sort_values("淨出貨金額", ascending=False)
        )

    left, right = st.columns(2)
    with left:
        st.markdown("**客戶業績比較**")
        if by_customer.empty:
            st.info("無資料")
        else:
            st.bar_chart(by_customer.head(10).set_index("客戶"))
    with right:
        st.markdown("**業績佔比**")
        if by_customer.empty:
            st.info("無資料")
        else:
            pct_df = by_customer.head(10).copy()
            total = pct_df["淨出貨金額"].sum()
            pct_df["佔比%"] = ((pct_df["淨出貨金額"] / total) * 100).round(2) if total else 0
            display_df = pct_df.copy()
            display_df["淨出貨金額"] = display_df["淨出貨金額"].map(lambda x: _fmt_money(x, currency_symbol))
            st.dataframe(display_df, use_container_width=True, hide_index=True)

    daily = pd.DataFrame(columns=["日期", "淨出貨金額"])
    if not month_df.empty:
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
            po_col, part_col, qty_col, amount_col, hold_col, discount_col, net_col, remark_col, wip_col
        ]
        if c and c in month_df.columns
    ]

    st.markdown("**業績明細表**")
    st.dataframe(month_df[show_cols] if show_cols else month_df, use_container_width=True, hide_index=True, height=420)

    excel_bytes = _build_excel(
        month_df=month_df,
        show_cols=show_cols if show_cols else list(month_df.columns),
        by_customer=by_customer,
        daily=daily,
        report_month=report_month,
        company_name=company_name,
        currency_symbol=currency_symbol,
    )
    st.download_button(
        "下載 Excel 業績圖表報表",
        data=excel_bytes,
        file_name=f"sales_report_{report_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
