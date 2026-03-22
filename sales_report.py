# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Optional

import pandas as pd
import streamlit as st


def _norm(s: str) -> str:
    return str(s).strip().lower()


def _pick_col(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    exact = {_norm(c): c for c in df.columns}
    for cand in candidates:
        if _norm(cand) in exact:
            return exact[_norm(cand)]
    best = None
    best_score = -1
    for c in df.columns:
        cl = _norm(c)
        score = 0
        for cand in candidates:
            cd = _norm(cand)
            if cd and cd in cl:
                score = max(score, len(cd))
        if score > best_score:
            best_score = score
            best = c
    return best if best_score > 0 else None


def _to_num(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=float)
    s = series.astype(str).str.replace(',', '', regex=False)
    for token in ['US$', 'USD', '$', 'NT$', 'EUR', '¥']:
        s = s.str.replace(token, '', regex=False)
    s = s.str.replace('(', '-', regex=False).str.replace(')', '', regex=False).str.strip()
    return pd.to_numeric(s, errors='coerce').fillna(0.0)


def _fmt_money(v: float, symbol: str) -> str:
    return f"{symbol} {float(v):,.0f}"


def _to_month(series: pd.Series) -> pd.Series:
    dt = pd.to_datetime(series, errors='coerce')
    return dt.dt.strftime('%Y-%m').fillna('')


def _download_excel(df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='業績明細')
    bio.seek(0)
    return bio.getvalue()


def render_sales_report_page(df=None, orders=None, po_col=None, customer_col=None, part_col=None,
                             qty_col=None, factory_col=None, ship_date_col=None, order_date_col=None,
                             remark_col=None, **kwargs):
    src = orders if isinstance(orders, pd.DataFrame) and not orders.empty else df

    st.subheader('業績明細表')
    st.caption('只統計所選月份；預設幣別為美金。')

    if src is None or src.empty:
        st.warning('目前沒有可用資料')
        return

    work = src.copy()
    customer_col = customer_col if customer_col in work.columns else _pick_col(work, ['客戶', 'customer'])
    factory_col = factory_col if factory_col in work.columns else _pick_col(work, ['工廠', 'factory'])
    qty_col = qty_col if qty_col in work.columns else _pick_col(work, ['qty', 'quantity', "order q'ty", 'pcs'])

    date_default = None
    for c in [ship_date_col, order_date_col, 'Ship date', 'Ship Date', '出貨日期', '日期', 'Date']:
        if c in work.columns:
            date_default = c
            break

    order_amt_auto = _pick_col(work, ['接單金額', 'order amount', 'order amt', 'sales amount', 'amount', 'total amount', 'order total', 'usd amount'])
    ship_amt_auto = _pick_col(work, ['出貨金額', 'shipment amount', 'ship amount', 'invoice amount', 'net shipment', 'net amount', 'sales usd', 'revenue'])
    unit_price_auto = _pick_col(work, ['單價', 'unit price', 'price', 'unit usd', 'usd/pcs', 'us$/pcs'])
    hold_auto = _pick_col(work, ['hold', 'hold amount', 'hold金額'])
    discount_auto = _pick_col(work, ['折讓', 'discount', 'discount amount'])

    all_cols = ['(無)'] + [str(c) for c in work.columns]

    def idx_for(col):
        return all_cols.index(col) if col in all_cols else 0

    c1, c2, c3 = st.columns(3)
    report_month = c1.text_input('報表月份 (YYYY-MM)', value=datetime.now().strftime('%Y-%m'))

    # 只改項目 1：預設公司改成 WESCO
    company_name = c2.text_input('子表名稱 / 公司名稱', value='WESCO')

    currency_symbol = c3.text_input('幣別符號', value='US$')

    with st.expander('欄位偵測'):
        d1, d2, d3 = st.columns(3)
        date_col = d1.selectbox('日期欄位', all_cols, index=idx_for(date_default))
        order_amt_col = d2.selectbox('接單金額欄位', all_cols, index=idx_for(order_amt_auto))
        ship_amt_col = d3.selectbox('出貨金額欄位', all_cols, index=idx_for(ship_amt_auto))
        e1, e2, e3 = st.columns(3)
        unit_price_col = e1.selectbox('單價欄位', all_cols, index=idx_for(unit_price_auto))
        hold_col = e2.selectbox('HOLD欄位', all_cols, index=idx_for(hold_auto))
        discount_col = e3.selectbox('折讓欄位', all_cols, index=idx_for(discount_auto))

    if date_col == '(無)':
        st.error('請在欄位偵測中指定日期欄位')
        return

    work['_month'] = _to_month(work[date_col])
    month_df = work[work['_month'] == report_month].copy()

    # 只改項目 2：實際套用公司篩選，避免業績明細抓錯或顯示 0
    if company_name.strip():
        candidate_company_cols = []

        if customer_col and customer_col in month_df.columns:
            candidate_company_cols.append(customer_col)

        for col in ['子表名稱', '公司名稱', '客戶', 'Customer', 'customer']:
            if col in month_df.columns and col not in candidate_company_cols:
                candidate_company_cols.append(col)

        filtered_df = pd.DataFrame()
        matched_col = None

        for col in candidate_company_cols:
            temp = month_df[
                month_df[col].astype(str).str.strip().str.upper()
                == company_name.strip().upper()
            ].copy()
            if not temp.empty:
                filtered_df = temp
                matched_col = col
                break

        if matched_col is not None:
            month_df = filtered_df

    if month_df.empty:
        st.warning('所選月份或公司沒有資料')

    qty = _to_num(month_df[qty_col]) if qty_col and qty_col in month_df.columns else pd.Series(0.0, index=month_df.index)
    unit_price = _to_num(month_df[unit_price_col]) if unit_price_col != '(無)' and unit_price_col in month_df.columns else pd.Series(0.0, index=month_df.index)

    if order_amt_col != '(無)' and order_amt_col in month_df.columns:
        month_df['_order_amt'] = _to_num(month_df[order_amt_col])
    else:
        month_df['_order_amt'] = unit_price * qty

    if ship_amt_col != '(無)' and ship_amt_col in month_df.columns:
        month_df['_ship_amt'] = _to_num(month_df[ship_amt_col])
    else:
        month_df['_ship_amt'] = unit_price * qty

    hold = _to_num(month_df[hold_col]) if hold_col != '(無)' and hold_col in month_df.columns else pd.Series(0.0, index=month_df.index)
    discount = _to_num(month_df[discount_col]) if discount_col != '(無)' and discount_col in month_df.columns else pd.Series(0.0, index=month_df.index)
    month_df['_net'] = month_df['_ship_amt'] - hold - discount

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric('接單金額', _fmt_money(month_df['_order_amt'].sum(), currency_symbol))
    k2.metric('出貨金額', _fmt_money(month_df['_ship_amt'].sum(), currency_symbol))
    k3.metric('淨出貨', _fmt_money(month_df['_net'].sum(), currency_symbol))
    cust_count = int(month_df[customer_col].astype(str).replace('', pd.NA).dropna().nunique()) if customer_col and customer_col in month_df.columns else 0
    fac_count = int(month_df[factory_col].astype(str).replace('', pd.NA).dropna().nunique()) if factory_col and factory_col in month_df.columns else 0
    k4.metric('客戶數', cust_count)
    k5.metric('廠商數', fac_count)

    by_customer = pd.DataFrame(columns=['客戶', '出貨金額', '佔比%'])
    if customer_col and customer_col in month_df.columns:
        by_customer = month_df.groupby(customer_col, dropna=False)['_ship_amt'].sum().reset_index()
        by_customer.columns = ['客戶', '出貨金額']
        by_customer = by_customer.sort_values('出貨金額', ascending=False)
        total_ship = float(by_customer['出貨金額'].sum())
        by_customer['佔比%'] = (by_customer['出貨金額'] / total_ship * 100).round(2) if total_ship else 0.0

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('**客戶業績比較**')
        if not by_customer.empty:
            st.bar_chart(by_customer.set_index('客戶')[['出貨金額']])
        else:
            st.info('沒有可用金額欄位，請在欄位偵測指定金額或單價欄位。')
    with c2:
        st.markdown('**業績佔比**')
        if not by_customer.empty:
            show_pct = by_customer.copy()
            show_pct['出貨金額'] = show_pct['出貨金額'].map(lambda x: _fmt_money(x, currency_symbol))
            st.dataframe(show_pct, use_container_width=True, hide_index=True, height=320)
        else:
            st.info('沒有資料')

    show_cols = [c for c in [date_col, po_col, customer_col, part_col, qty_col, factory_col, remark_col] if c and c in month_df.columns]
    if '_order_amt' in month_df.columns:
        month_df['接單金額'] = month_df['_order_amt']
        show_cols.append('接單金額')
    if '_ship_amt' in month_df.columns:
        month_df['出貨金額'] = month_df['_ship_amt']
        show_cols.append('出貨金額')
    if '_net' in month_df.columns:
        month_df['淨出貨'] = month_df['_net']
        show_cols.append('淨出貨')

    st.markdown('**明細**')
    if show_cols:
        st.dataframe(month_df[show_cols], use_container_width=True, height=420, hide_index=True)
        st.download_button('下載明細 CSV', month_df[show_cols].to_csv(index=False).encode('utf-8-sig'), 'sales_report.csv', 'text/csv')
        st.download_button('下載明細 Excel', _download_excel(month_df[show_cols]), 'sales_report.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
