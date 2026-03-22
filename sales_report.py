from __future__ import annotations

import io
from datetime import datetime
from typing import Any, Iterable

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


MONEY_KEYWORDS = [
    "淨出貨金額", "出貨金額", "銷貨金額", "業績金額", "開票金額", "invoice amount",
    "shipment amount", "ship amount", "amount", "amt", "sales amount", "revenue",
    "total amount", "net amount", "net sales", "營業額", "金額", "sales", "invoice",
]
HOLD_KEYWORDS = ["hold金額", "hold amount", "hold", "暫停金額"]
DISCOUNT_KEYWORDS = ["折讓", "discount", "credit note", "allowance"]
NET_KEYWORDS = ["淨出貨金額", "net amount", "net sales", "淨額", "net"]
DATE_KEYWORDS = ["出貨日期", "ship date", "日期", "date", "invoice date", "開票日", "order date"]


def _normalize_name(name: Any) -> str:
    return str(name or "").strip().lower().replace("_", " ")


def _safe_num(series: pd.Series) -> pd.Series:
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


def _find_best_column(df: pd.DataFrame, preferred: Iterable[str], exclude: Iterable[str] | None = None) -> str | None:
    exclude = set(exclude or [])
    cols = [c for c in df.columns if c not in exclude]
    if not cols:
        return None

    # 1) exact-ish keyword score
    scored = []
    for col in cols:
        norm = _normalize_name(col)
        score = sum(10 for kw in preferred if kw.lower() == norm)
        score += sum(3 for kw in preferred if kw.lower() in norm)
        num_sum = _safe_num(df[col]).abs().sum()
        scored.append((score, num_sum, col))
    scored.sort(reverse=True)
    if scored and scored[0][0] > 0:
        return scored[0][2]

    # 2) fallback to numeric column with highest sum
    numeric_scored = []
    for col in cols:
        num_sum = _safe_num(df[col]).abs().sum()
        numeric_scored.append((num_sum, col))
    numeric_scored.sort(reverse=True)
    if numeric_scored and numeric_scored[0][0] > 0:
        return numeric_scored[0][1]
    return None


def _pick_date_col(df: pd.DataFrame, order_date_col: str | None, ship_date_col: str | None) -> str | None:
    for c in [ship_date_col, order_date_col]:
        if c and c in df.columns:
            return c
    for col in df.columns:
        norm = _normalize_name(col)
        if any(k in norm for k in DATE_KEYWORDS):
            return col
    return None


def _fmt_money(value: float, currency_symbol: str = "NT$") -> str:
    return f"{currency_symbol} {value:,.0f}"


def build_excel_report(
    month_df: pd.DataFrame,
    report_month: str,
    company_name: str,
    currency_symbol: str,
    customer_col: str | None,
    factory_col: str | None,
    display_cols: list[str],
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "業績圖表總覽"

    ws["A1"] = f"{report_month} 業績報表"
    ws["A2"] = f"子表 {company_name}" if company_name else "子表"
    ws["A1"].font = Font(size=18, bold=True)
    ws["A2"].font = Font(size=11, bold=False)

    total_amount = float(month_df["_amount"].sum()) if "_amount" in month_df.columns else 0.0
    total_net = float(month_df["_net"].sum()) if "_net" in month_df.columns else 0.0
    customer_count = int(month_df[customer_col].astype(str).replace("", pd.NA).dropna().nunique()) if customer_col and customer_col in month_df.columns else 0
    factory_count = int(month_df[factory_col].astype(str).replace("", pd.NA).dropna().nunique()) if factory_col and factory_col in month_df.columns else 0

    ws["A4"] = "總業績"
    ws["C4"] = "客戶數"
    ws["E4"] = "廠商數"
    ws["G4"] = "淨出貨"
    ws["A5"] = total_amount
    ws["C5"] = customer_count
    ws["E5"] = factory_count
    ws["G5"] = total_net
    ws["A5"].number_format = f'"{currency_symbol}" #,##0'
    ws["G5"].number_format = f'"{currency_symbol}" #,##0'

    if customer_col and customer_col in month_df.columns:
        by_customer = (
            month_df.groupby(customer_col)["_net"]
            .sum()
            .sort_values(ascending=False)
            .head(10)
            .reset_index()
        )
        ws["A8"] = "客戶"
        ws["B8"] = "淨出貨金額"
        for i, (_, row) in enumerate(by_customer.iterrows(), start=9):
            ws[f"A{i}"] = row[customer_col]
            ws[f"B{i}"] = float(row["_net"])
            ws[f"B{i}"].number_format = f'"{currency_symbol}" #,##0'

        if len(by_customer) > 0:
            bar = BarChart()
            bar.title = "客戶業績比較"
            bar.height = 8
            bar.width = 12
            data = Reference(ws, min_col=2, min_row=8, max_row=8 + len(by_customer))
            cats = Reference(ws, min_col=1, min_row=9, max_row=8 + len(by_customer))
            bar.add_data(data, titles_from_data=True)
            bar.set_categories(cats)
            bar.legend = None
            ws.add_chart(bar, "D8")

            donut = DoughnutChart()
            donut.title = "業績佔比"
            donut.holeSize = 55
            donut.height = 8
            donut.width = 10
            donut.add_data(data, titles_from_data=True)
            donut.set_categories(cats)
            ws.add_chart(donut, "Q8")

    if "_report_date" in month_df.columns and "_net" in month_df.columns:
        daily = (
            month_df.groupby(month_df["_report_date"].dt.strftime("%Y-%m-%d"))["_net"]
            .sum()
            .reset_index()
        )
        if not daily.empty:
            start = 28
            ws[f"A{start}"] = "日期"
            ws[f"B{start}"] = "淨出貨金額"
            for i, (_, row) in enumerate(daily.iterrows(), start=start + 1):
                ws[f"A{i}"] = row.iloc[0]
                ws[f"B{i}"] = float(row.iloc[1])
                ws[f"B{i}"].number_format = f'"{currency_symbol}" #,##0'
            line = LineChart()
            line.title = "每日淨出貨趨勢"
            line.height = 8
            line.width = 18
            data = Reference(ws, min_col=2, min_row=start, max_row=start + len(daily))
            cats = Reference(ws, min_col=1, min_row=start + 1, max_row=start + len(daily))
            line.add_data(data, titles_from_data=True)
            line.set_categories(cats)
            ws.add_chart(line, "D28")

    sh = wb.create_sheet("業績明細表")
    for c_idx, col in enumerate(display_cols, start=1):
        sh.cell(row=1, column=c_idx, value=col).font = Font(bold=True)
    for r_idx, (_, row) in enumerate(month_df.iterrows(), start=2):
        for c_idx, col in enumerate(display_cols, start=1):
            sh.cell(row=r_idx, column=c_idx, value="" if pd.isna(row.get(col)) else row.get(col))
    for i in range(1, len(display_cols) + 1):
        sh.column_dimensions[get_column_letter(i)].width = 16

    raw = wb.create_sheet("RawData")
    for c_idx, col in enumerate(month_df.columns, start=1):
        raw.cell(row=1, column=c_idx, value=col).font = Font(bold=True)
    for r_idx, (_, row) in enumerate(month_df.iterrows(), start=2):
        for c_idx, col in enumerate(month_df.columns, start=1):
            raw.cell(row=r_idx, column=c_idx, value="" if pd.isna(row.get(col)) else row.get(col))

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
    report_month = st.text_input("報表月份 (YYYY-MM)", value=datetime.now().strftime("%Y-%m"))
    company_name = st.text_input("子表名稱 / 公司名稱", value="WESCO")
    currency_symbol = st.text_input("幣別符號", value="NT$")

    date_col = _pick_date_col(work, order_date_col, ship_date_col)
    if not date_col:
        st.error("找不到日期欄位，至少要有 日期 / 出貨日期 / Ship Date 其中一欄")
        return

    work["_report_date"] = pd.to_datetime(work[date_col], errors="coerce")
    work["_month"] = work["_report_date"].dt.strftime("%Y-%m")
    month_df = work[work["_month"] == report_month].copy()

    if month_df.empty:
        st.warning(f"{report_month} 沒有資料")
        return

    amount_col = _find_best_column(month_df, MONEY_KEYWORDS)
    net_col = _find_best_column(month_df, NET_KEYWORDS, exclude=[amount_col] if amount_col else None)
    hold_col = _find_best_column(month_df, HOLD_KEYWORDS, exclude=[c for c in [amount_col, net_col] if c])
    discount_col = _find_best_column(month_df, DISCOUNT_KEYWORDS, exclude=[c for c in [amount_col, net_col, hold_col] if c])

    month_df["_amount"] = _safe_num(month_df[amount_col]) if amount_col else 0
    month_df["_hold"] = _safe_num(month_df[hold_col]) if hold_col else 0
    month_df["_discount"] = _safe_num(month_df[discount_col]) if discount_col else 0
    if net_col:
        month_df["_net"] = _safe_num(month_df[net_col])
    elif amount_col:
        month_df["_net"] = month_df["_amount"] - month_df["_hold"] - month_df["_discount"]
    else:
        # 若找不到金額欄位，避免誤導，不硬算成 0 而是保留為 0 並提示偵測結果
        month_df["_net"] = 0

    total_amount = float(month_df["_amount"].sum())
    total_net = float(month_df["_net"].sum())
    customer_count = month_df[customer_col].astype(str).replace("", pd.NA).dropna().nunique() if customer_col and customer_col in month_df.columns else 0
    factory_count = month_df[factory_col].astype(str).replace("", pd.NA).dropna().nunique() if factory_col and factory_col in month_df.columns else 0

    with st.expander("欄位偵測", expanded=False):
        st.write({
            "date_col": date_col,
            "amount_col": amount_col,
            "net_col": net_col,
            "hold_col": hold_col,
            "discount_col": discount_col,
        })
        if not amount_col and not net_col:
            st.warning("目前未成功偵測到金額欄位，所以總業績 / 淨出貨會顯示 0。請確認原表是否有金額欄位，或把欄位名稱貼給我再幫你對應。")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("總業績", _fmt_money(total_amount, currency_symbol))
    c2.metric("客戶數", int(customer_count))
    c3.metric("廠商數", int(factory_count))
    c4.metric("淨出貨", _fmt_money(total_net, currency_symbol))

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
            if not by_customer.empty:
                st.bar_chart(by_customer)
            else:
                st.info("無資料")
        with right:
            st.markdown("**業績佔比**")
            if not by_customer.empty:
                pct_df = by_customer.reset_index()
                pct_df.columns = ["客戶", "淨出貨金額"]
                total = pct_df["淨出貨金額"].sum()
                pct_df["佔比%"] = (pct_df["淨出貨金額"] / total * 100).round(2) if total else 0
                pct_df["淨出貨金額"] = pct_df["淨出貨金額"].map(lambda x: _fmt_money(x, currency_symbol))
                st.dataframe(pct_df, use_container_width=True, hide_index=True)
            else:
                st.info("無資料")

    daily = (
        month_df.groupby(month_df["_report_date"].dt.strftime("%Y-%m-%d"))["_net"]
        .sum()
        .reset_index()
    )
    daily.columns = ["日期", "淨出貨金額"]
    if not daily.empty:
        st.markdown("**每日趨勢**")
        st.line_chart(daily.set_index("日期"))

    display_cols = [
        c for c in [
            order_date_col, ship_date_col, customer_col, factory_col,
            po_col, part_col, qty_col, amount_col, hold_col, discount_col, net_col, wip_col, remark_col
        ] if c and c in month_df.columns
    ]
    display_cols = list(dict.fromkeys(display_cols))

    st.markdown("**業績明細表**")
    st.dataframe(month_df[display_cols], use_container_width=True, hide_index=True, height=460)

    excel_bytes = build_excel_report(
        month_df=month_df,
        report_month=report_month,
        company_name=company_name,
        currency_symbol=currency_symbol,
        customer_col=customer_col,
        factory_col=factory_col,
        display_cols=display_cols,
    )
    st.download_button(
        "下載 Excel 業績圖表報表",
        data=excel_bytes,
        file_name=f"sales_report_{report_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
