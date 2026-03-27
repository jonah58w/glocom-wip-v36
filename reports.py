# -*- coding: utf-8 -*-
"""
reports.py
完整修正版：業績明細表
資料來源：Teable API (load_orders() 取得的 DataFrame)
重點指標：
- 接單金額 (USD)：該月份所有訂單的 USD 金額總和
- 出貨金額 (INVOICE USD)：該月份 WIP 欄位包含 "SHIPMENT" 的 INVOICE 金額總和
- 不使用匯率，全部以 USD 計算
"""

import pandas as pd
import streamlit as st
from datetime import datetime
from typing import Dict, Optional
import plotly.graph_objects as go


# ====================== 輔助函式 ======================
def detect_columns(df: pd.DataFrame) -> Dict:
    """自動偵測常用欄位名稱（支援中英文混合命名）"""
    col_map = {}
    possible_mappings = {
        'invoice_date': ['Invoice Date', '出貨日期', 'Ship Date', 'InvoiceDate', '發票日期', 'Shipment Date'],
        'usd_amount': ['USD', 'USD Amount', 'Order USD', 'Amount USD', '接單金額(USD)', 'PO Amount USD'],
        'invoice_usd': ['INVOICE USD', 'Invoice Amount', 'Shipment Amount USD', '發票金額(USD)', 'Invoice USD'],
        'wip_status': ['WIP', 'Status', 'WIP Status', 'Shipment Status', '狀態', 'WIP Stage'],
        'customer': ['Customer', '客戶', 'Client'],
        'part_no': ['Part No', '料號', 'PN', 'Part Number'],
        'po': ['PO', '訂單號', 'Order No', 'PO Number']
    }
    
    df_cols = [str(c).strip() for c in df.columns]
    for key, candidates in possible_mappings.items():
        for cand in candidates:
            if cand in df_cols:
                col_map[key] = cand
                break
    return col_map


# ====================== 主要分析類別 ======================
class SalesDetailAnalyzer:
    def __init__(self, df: pd.DataFrame):
        self.raw_df = df.copy()
        self.col_map = detect_columns(df)
        self.normalized_df = self._normalize_data()
        
    def _normalize_data(self) -> pd.DataFrame:
        """資料正規化：日期、金額、狀態處理"""
        df = self.raw_df.copy()
        
        # 日期 → 月份
        date_col = self.col_map.get('invoice_date')
        if date_col and date_col in df.columns:
            df['_date'] = pd.to_datetime(df[date_col], errors='coerce')
            df['_month'] = df['_date'].dt.strftime('%Y-%m')
        else:
            df['_month'] = datetime.now().strftime('%Y-%m')
        
        # 接單金額 (USD)
        usd_col = self.col_map.get('usd_amount')
        if usd_col and usd_col in df.columns:
            df['_order_usd'] = pd.to_numeric(df[usd_col], errors='coerce').fillna(0.0)
        else:
            df['_order_usd'] = 0.0
        
        # 出貨 INVOICE 金額 (USD)
        invoice_col = self.col_map.get('invoice_usd')
        if invoice_col and invoice_col in df.columns:
            df['_invoice_usd'] = pd.to_numeric(df[invoice_col], errors='coerce').fillna(0.0)
        else:
            df['_invoice_usd'] = 0.0
        
        # WIP 狀態（用來判斷是否已出貨）
        wip_col = self.col_map.get('wip_status')
        if wip_col and wip_col in df.columns:
            df['_wip'] = df[wip_col].fillna('').astype(str).str.upper()
        else:
            df['_wip'] = ''
        
        return df
    
    def get_month_summary(self, month_str: Optional[str] = None) -> Dict:
        """取得單月業績摘要"""
        if month_str is None:
            month_str = datetime.now().strftime('%Y-%m')
        
        df = self.normalized_df.copy()
        month_df = df[df['_month'] == month_str].copy()
        
        # 接單金額（該月所有訂單）
        order_usd = month_df['_order_usd'].sum()
        
        # 出貨金額：僅 WIP 含有 SHIPMENT 的 INVOICE 金額
        shipped_mask = month_df['_wip'].str.contains('SHIPMENT', na=False)
        shipment_usd = month_df.loc[shipped_mask, '_invoice_usd'].sum()
        
        return {
            'month': month_str,
            'order_usd': round(float(order_usd), 2),
            'shipment_usd': round(float(shipment_usd), 2),
            'order_count': len(month_df),
            'shipped_count': int(shipped_mask.sum())
        }
    
    def get_monthly_trend(self, months: int = 12) -> pd.DataFrame:
        """過去 N 個月的接單 vs 出貨趨勢"""
        df = self.normalized_df.copy()
        all_months = sorted(df['_month'].unique())[-months:]
        
        results = []
        for m in all_months:
            summary = self.get_month_summary(m)
            results.append(summary)
        
        return pd.DataFrame(results)
    
    def get_customer_summary(self) -> pd.DataFrame:
        """客戶別業績統計"""
        df = self.normalized_df.copy()
        customer_col = self.col_map.get('customer')
        
        if not customer_col or customer_col not in df.columns:
            return pd.DataFrame(columns=['客戶', '接單金額(USD)', '出貨金額(USD)', '訂單數'])
        
        grouped = df.groupby(customer_col).agg(
            order_usd=('_order_usd', 'sum'),
            shipment_usd=('_invoice_usd', lambda x: x[df.loc[x.index, '_wip'].str.contains('SHIPMENT', na=False)].sum() if len(x) > 0 else 0),
            order_count=('_order_usd', 'count')
        ).round(2).reset_index()
        
        grouped = grouped.rename(columns={
            customer_col: '客戶',
            'order_usd': '接單金額(USD)',
            'shipment_usd': '出貨金額(USD)',
            'order_count': '訂單數'
        })
        return grouped.sort_values('接單金額(USD)', ascending=False)


# ====================== Streamlit 儀表板 ======================
def render_sales_detail_dashboard(orders_df: pd.DataFrame, default_month: Optional[str] = None):
    """渲染完整的業績明細儀表板"""
    if orders_df is None or orders_df.empty:
        st.error("❌ 無訂單資料可顯示")
        return
    
    analyzer = SalesDetailAnalyzer(orders_df)
    
    st.title("📊 業績明細表")
    st.caption("資料來源：Teable API｜已出貨判斷依據：WIP 欄位包含「SHIPMENT」｜單位：USD")
    
    # 月份選擇
    months = sorted(analyzer.normalized_df['_month'].unique(), reverse=True)
    if not months:
        st.warning("目前無可用月份資料")
        return
    
    selected_month = st.selectbox(
        "選擇月份",
        months,
        index=0 if default_month is None else months.index(default_month) if default_month in months else 0
    )
    
    # 當月摘要
    summary = analyzer.get_month_summary(selected_month)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("接單金額 (USD)", f"${summary['order_usd']:,.2f}")
    with col2:
        st.metric("出貨金額 (INVOICE USD)", f"${summary['shipment_usd']:,.2f}")
    with col3:
        st.metric("接單筆數", f"{summary['order_count']:,}")
    with col4:
        st.metric("已出貨筆數", f"{summary['shipped_count']:,}")
    
    # 月度趨勢圖
    st.subheader("📈 近 12 個月 接單 vs 出貨 趨勢")
    trend_df = analyzer.get_monthly_trend(12)
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=trend_df['month'],
        y=trend_df['order_usd'],
        name='接單金額 USD',
        marker_color='#1f77b4'
    ))
    fig.add_trace(go.Bar(
        x=trend_df['month'],
        y=trend_df['shipment_usd'],
        name='出貨金額 USD',
        marker_color='#ff7f0e'
    ))
    fig.update_layout(
        barmode='group',
        height=400,
        xaxis_title="月份",
        yaxis_title="金額 (USD)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # 當月詳細資料
    st.subheader(f"📋 {selected_month} 詳細資料")
    month_df = analyzer.normalized_df[analyzer.normalized_df['_month'] == selected_month].copy()
    
    # 優先顯示重要欄位
    display_cols = []
    priority = ['PO', 'Customer', '客戶', 'Part No', '料號', 'WIP', 'Status', 
                'USD', 'USD Amount', 'INVOICE USD', 'Invoice Date', '出貨日期']
    
    for p in priority:
        if p in month_df.columns:
            display_cols.append(p)
    
    if display_cols:
        st.dataframe(month_df[display_cols], use_container_width=True, height=500)
    else:
        st.dataframe(month_df, use_container_width=True, height=500)
    
    # 匯出按鈕
    csv = month_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 下載本月明細 CSV",
        data=csv,
        file_name=f"業績明細_{selected_month}.csv",
        mime="text/csv"
    )
    
    # 客戶別統計（可選展開）
    with st.expander("👥 客戶別業績統計"):
        customer_summary = analyzer.get_customer_summary()
        if not customer_summary.empty:
            st.dataframe(customer_summary, use_container_width=True)
        else:
            st.info("無法產生客戶別統計")


# ====================== 使用範例（供參考） ======================
# 在 app.py 中使用方式：
"""
from reports import render_sales_detail_dashboard

if menu == "Sales Detail":
    render_sales_detail_dashboard(orders)   # orders 為 load_orders() 取得的 DataFrame
"""
