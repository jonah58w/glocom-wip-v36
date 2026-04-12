# -*- coding: utf-8 -*-
from __future__ import annotations

from io import BytesIO
from typing import Optional

import pandas as pd
import streamlit as st

# ── 新增：OBU 模組 ─────────────────────────────────────────────────────────────
try:
    from obu_page import render_customs_price_tab, render_obu_calc_tab
    _OBU_AVAILABLE = True
except ImportError:
    _OBU_AVAILABLE = False


def _norm(s: str) -> str:
    s = str(s).strip()
    s = s.replace("\n", "")
    s = s.replace(" ", "")
    s = s.replace("（", "(").replace("）", ")")
    return s.lower()


def _pick_col(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    if df is None or df.empty:
        return None

    norm_map = {_norm(c): c for c in df.columns}

    for cand in candidates:
        nc = _norm(cand)
        if nc in norm_map:
            return norm_map[nc]

    best = None
    best_score = -1
    for col in df.columns:
        ncol = _norm(col)
        score = 0
        for cand in candidates:
            nc = _norm(cand)
            if nc and nc in ncol:
                score = max(score, len(nc))
        if score > best_score:
            best_score = score
            best = col

    return best if best_score > 0 else None


def _to_num(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=float)

    s = series.astype(str).fillna("")
    s = s.str.replace(",", "", regex=False)
    for token in ["US$", "USD", "$", "NT$", "EUR", "¥"]:
        s = s.str.replace(token, "", regex=False)
    s = s.str.replace("(", "-", regex=False).str.replace(")", "", regex=False).str.strip()
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _fmt_money(v: float, symbol: str) -> str:
    fv = float(v)
    return f"{symbol} {fv:,.1f}" if abs(fv - round(fv)) > 1e-9 else f"{symbol} {fv:,.0f}"


def _download_excel(df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="業績明細")
    bio.seek(0)
    return bio.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# 主函式：render_sales_report_page
# 原有業績明細邏輯放在 Tab 1，新增 Tab 2（報關單價）與 Tab 3（匯 HK OBU 金額）
# ══════════════════════════════════════════════════════════════════════════════

def render_sales_report_page(
    df=None,
    orders=None,
    sales_df=None,
    po_col=None,
    customer_col=None,
    part_col=None,
    qty_col=None,
    factory_col=None,
    ship_date_col=None,
    order_date_col=None,
    remark_col=None,
    **kwargs,
):
    st.subheader("業績明細表")
    st.caption("業績明細表 v5 — 含報關單價核算 & 匯 HK OBU 金額")

    # ── Tab 結構 ──────────────────────────────────────────────────────────────
    if _OBU_AVAILABLE:
        tab1, tab2, tab3 = st.tabs([
            "📊 業績明細",
            "🏷️ 報關單價",
            "🏦 匯 HK OBU 金額",
        ])
    else:
        tab1 = st.container()
        tab2 = tab3 = None
        st.warning("⚠️ obu_page.py / customs_price.py 未找到，OBU 功能停用。")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 1：業績明細（原始邏輯，完全保留）
    # ══════════════════════════════════════════════════════════════════════════
    with tab1:
        _render_sales_detail_tab(
            df=df,
            orders=orders,
            sales_df=sales_df,
            po_col=po_col,
            customer_col=customer_col,
            part_col=part_col,
            qty_col=qty_col,
            factory_col=factory_col,
            ship_date_col=ship_date_col,
            order_date_col=order_date_col,
            remark_col=remark_col,
        )

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2：報關單價核算
    # ══════════════════════════════════════════════════════════════════════════
    if _OBU_AVAILABLE and tab2 is not None:
        with tab2:
            render_customs_price_tab()

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 3：匯 HK OBU 金額計算
    # ══════════════════════════════════════════════════════════════════════════
    if _OBU_AVAILABLE and tab3 is not None:
        with tab3:
            render_obu_calc_tab()


# ─────────────────────────────────────────────────────────────────────────────
# 原業績明細邏輯（抽成獨立函式）
# ─────────────────────────────────────────────────────────────────────────────

def _render_sales_detail_tab(
    df=None,
    orders=None,
    sales_df=None,
    po_col=None,
    customer_col=None,
    part_col=None,
    qty_col=None,
    factory_col=None,
    ship_date_col=None,
    order_date_col=None,
    remark_col=None,
):
    src = sales_df if isinstance(sales_df, pd.DataFrame) and not sales_df.empty else (
        orders if isinstance(orders, pd.DataFrame) and not orders.empty else df
    )

    if src is None or src.empty:
        st.warning("目前沒有可用資料")
        return

    work = src.copy()

    real_customer_col  = _pick_col(work, ["Customer", "客戶", "公司名稱"])
    real_po_col        = _pick_col(work, ["PO#"])
    real_part_col      = _pick_col(work, ["P/N"])
    real_qty_col       = _pick_col(work, ["Order Q'TY (PCS)", "Order Q'TY\n(PCS)",
                                          "Order Q'TY\n (PCS)", "QTY", "Qty"])
    real_factory_col   = _pick_col(work, ["工廠", "Factory"])
    real_ship_date_col = _pick_col(work, ["Ship date", "出貨日期", "Date"])
    real_remark_col    = _pick_col(work, ["工廠提醒事項", "Note", "Remark"])
    real_order_amt_col = _pick_col(work, ["INVOICE", "Invoice", "invoice amount",
                                          "出貨金額", "Amount"])
    real_ship_amt_col  = _pick_col(work, ["INVOICE", "Invoice", "invoice amount",
                                          "出貨金額", "Amount"])
    real_tooling_col   = _pick_col(work, ["TOOLING", "Tooling"])

    c1, c2, c3 = st.columns(3)
    report_month   = c1.text_input("報表月份 (YYYY-MM)", value="2026-04",
                                   key="sales_report_month_v5")
    company_name   = c2.text_input("子表名稱 / 公司名稱", value="WESCO",
                                   key="sales_report_company_v5")
    currency_symbol = c3.text_input("幣別符號", value="US$",
                                    key="sales_report_currency_v5")

    with st.expander("欄位偵測", expanded=False):
        st.write({
            "all_columns":   list(work.columns),
            "customer_col":  real_customer_col,
            "po_col":        real_po_col,
            "part_col":      real_part_col,
            "qty_col":       real_qty_col,
            "factory_col":   real_factory_col,
            "ship_date_col": real_ship_date_col,
            "order_amt_col": real_order_amt_col,
            "ship_amt_col":  real_ship_amt_col,
            "tooling_col":   real_tooling_col,
        })

    if not real_ship_date_col or real_ship_date_col not in work.columns:
        st.error("找不到 Ship date 欄位")
        return

    if not real_customer_col or real_customer_col not in work.columns:
        st.error("找不到 Customer 欄位")
        return

    if not real_ship_amt_col or real_ship_amt_col not in work.columns:
        st.error("找不到 INVOICE 欄位")
        return

    work["_ship_date"] = pd.to_datetime(work[real_ship_date_col], errors="coerce")
    work["_month"]     = work["_ship_date"].dt.strftime("%Y-%m")

    month_df = work[work["_month"] == report_month].copy()

    if company_name.strip():
        month_df = month_df[
            month_df[real_customer_col].astype(str).str.strip().str.upper()
            == company_name.strip().upper()
        ].copy()

    invoice = (_to_num(month_df[real_ship_amt_col])
               if real_ship_amt_col in month_df.columns
               else pd.Series(0.0, index=month_df.index))
    tooling = (_to_num(month_df[real_tooling_col])
               if real_tooling_col and real_tooling_col in month_df.columns
               else pd.Series(0.0, index=month_df.index))

    month_df["_order_amt"] = invoice + tooling
    month_df["_ship_amt"]  = invoice
    month_df["_net"]       = invoice

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("接單金額",  _fmt_money(month_df["_order_amt"].sum(), currency_symbol))
    k2.metric("出貨金額",  _fmt_money(month_df["_ship_amt"].sum(),  currency_symbol))
    k3.metric("淨出貨",    _fmt_money(month_df["_net"].sum(),       currency_symbol))
    k4.metric(
        "客戶數",
        int(month_df[real_customer_col].astype(str)
            .replace("", pd.NA).dropna().nunique())
        if not month_df.empty else 0
    )
    k5.metric(
        "廠商數",
        int(month_df[real_factory_col].astype(str)
            .replace("", pd.NA).dropna().nunique())
        if real_factory_col and real_factory_col in month_df.columns
           and not month_df.empty else 0
    )

    by_customer = pd.DataFrame(columns=["客戶", "出貨金額", "佔比%"])
    if not month_df.empty:
        by_customer = (
            month_df.groupby(real_customer_col, dropna=False)["_ship_amt"]
            .sum().reset_index()
        )
        by_customer.columns = ["客戶", "出貨金額"]
        total_ship = float(by_customer["出貨金額"].sum())
        by_customer["佔比%"] = (
            (by_customer["出貨金額"] / total_ship * 100).round(2)
            if total_ship else 0.0
        )

    left, right = st.columns(2)
    with left:
        st.markdown("**客戶業績比較**")
        if not by_customer.empty:
            st.bar_chart(by_customer.set_index("客戶")[["出貨金額"]])
        else:
            st.info("沒有可用資料")

    with right:
        st.markdown("**業績佔比**")
        if not by_customer.empty:
            show_pct = by_customer.copy()
            show_pct["出貨金額"] = show_pct["出貨金額"].map(
                lambda x: _fmt_money(x, currency_symbol)
            )
            st.dataframe(show_pct, use_container_width=True,
                         hide_index=True, height=320)
        else:
            st.info("沒有資料")

    if month_df.empty:
        st.info("目前沒有符合月份與公司條件的資料")
        return

    month_df["接單金額"] = month_df["_order_amt"]
    month_df["出貨金額"] = month_df["_ship_amt"]
    month_df["淨出貨"]   = month_df["_net"]

    show_cols = [
        c for c in [
            real_ship_date_col, real_po_col, real_customer_col, real_part_col,
            real_qty_col, real_factory_col, real_remark_col,
            "接單金額", "出貨金額", "淨出貨",
        ]
        if c and c in month_df.columns
    ]

    st.markdown("**明細**")
    st.dataframe(month_df[show_cols], use_container_width=True,
                 height=420, hide_index=True)

    st.download_button(
        "下載明細 CSV",
        month_df[show_cols].to_csv(index=False).encode("utf-8-sig"),
        "sales_report.csv", "text/csv",
    )
    st.download_button(
        "下載明細 Excel",
        _download_excel(month_df[show_cols]),
        "sales_report.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
