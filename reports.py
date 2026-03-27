# -*- coding: utf-8 -*-
"""
reports.py
完整修正版（2026/3/27 最終版） - 完全依照目前 APP 截圖與附件格式
資料來源：Teable API (load_orders() 取得的 DataFrame)
重點修正：
1. 金額欄位統一使用「銷貨金額(USD)」作為接單與出貨金額來源
2. 已出貨 = WIP 欄位包含 "SHIPMENT" 的金額總和
3. 預計出貨 = 本月非 SHIPMENT 的金額總和
4. 接單金額 = 本月全部金額總和（與月銷售合計相同）
5. 月銷售合計 = 已出貨 + 預計出貨
6. 新增「依工廠別統計」、「已出貨明細」、「預計出貨明細」
7. 完全匹配 PDF 截圖與附件圖片的顯示格式與數字
"""

import pandas as pd
import streamlit as st
from datetime import datetime
from typing import Dict, Optional
import plotly.graph_objects as go


# ====================== 輔助函式 ======================
def detect_columns(df: pd.DataFrame) -> Dict:
    """自動偵測 Teable 實際欄位名稱（已針對目前 APP 截圖最佳化）"""
    col_map = {}
    possible_mappings = {
        'date': ['日期', 'Invoice Date', '出貨日期', 'Ship Date', '發票日期', 'Shipment Date'],
        'amount_usd': ['銷貨金額(USD)', 'USD', 'Amount USD', '銷貨金額', 'Invoice Amount', 'Shipment Amount USD', '金額(USD)'],
        'wip_status': ['WIP', 'Status', 'WIP Status', 'Shipment Status', '狀態', 'WIP Stage'],
        'customer': ['客戶', 'Customer', 'Client'],
        'po': ['PO#', 'PO', '訂單號', 'Order No', 'PO Number'],
        'pn': ['P/N', 'Part No', '料號', 'PN', 'Part Number'],
        'factory': ['工廠', 'Factory'],
        'qty': ['QTY', '數量', 'Qty']
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
        df = self.raw_df.copy()
        
        # 日期 → 月份
        date_col = self.col_map.get('date')
        if date_col and date_col in df.columns:
            df['_date'] = pd.to_datetime(df[date_col], errors='coerce')
            df['_month'] = df['_date'].dt.strftime('%Y-%m')
        else:
            df['_month'] = datetime.now().strftime('%Y-%m')
        
        # 金額欄位（統一使用銷貨金額(USD)）
        amount_col = self.col_map.get('amount_usd')
        if amount_col and amount_col in df.columns:
            df['_amount_usd'] = pd.to_numeric(df[amount_col], errors='coerce').fillna(0.0)
        else:
            df['_amount_usd'] = 0.0
        
        # WIP 狀態
        wip_col = self.col_map.get('wip_status')
        if wip_col and wip_col in df.columns:
            df['_wip'] = df[wip_col].fillna('').astype(str).str.upper()
        else:
            df['_wip'] = ''
        
        # 其他欄位保留原名稱
        return df
    
    def get_month_summary(self, month_str: Optional[str] = None) -> Dict:
        """單月摘要（完全匹配 PDF 截圖的四個指標）"""
        if month_str is None:
            month_str = datetime.now().strftime('%Y-%m')
        
        df = self.normalized_df.copy()
        month_df = df[df['_month'] == month_str].copy()
        
        total_usd = month_df['_amount_usd'].sum()                     # 接單金額 / 月銷售合計
        shipped_mask = month_df['_wip'].str.contains('SHIPMENT', na=False)
        shipped_usd = month_df.loc[shipped_mask, '_amount_usd'].sum()  # 已出貨金額
        pending_usd = total_usd - shipped_usd                          # 預計出貨金額
        
        return {
            'month': month_str,
            'order_usd': round(float(total_usd), 2),      # 接單金額
            'shipped_usd': round(float(shipped_usd), 2),  # 已出貨金額 (SHIPMENT)
            'pending_usd': round(float(pending_usd), 2),  # 預計出貨金額
            'total_usd': round(float(total_usd), 2),      # 月銷售合計
            'order_count': len(month_df),
            'shipped_count': int(shipped_mask.sum())
        }
    
    def get_monthly_trend(self, months: int = 12) -> pd.DataFrame:
        """近12個月月銷貨趨勢"""
        df = self.normalized_df.copy()
        all_months = sorted(df['_month'].unique())[-months:]
        results = [self.get_month_summary(m) for m in all_months]
        return pd.DataFrame(results)
    
    def get_factory_summary(self, month_str: Optional[str] = None) -> pd.DataFrame:
        """依工廠別統計（完全匹配附件圖片格式）"""
        if month_str is None:
            month_str = datetime.now().strftime('%Y-%m')
        df = self.normalized_df.copy()
        month_df = df[df['_month'] == month_str].copy()
        factory_col = self.col_map.get('factory')
        
        if not factory_col or factory_col not in month_df.columns:
            return pd.DataFrame(columns=['工廠', '訂單數', '銷貨金額(USD)'])
        
        grouped = month_df.groupby(factory_col).agg(
            訂單數=('_amount_usd', 'count'),
            銷貨金額=('_amount_usd', 'sum')
        ).round(2).reset_index()
        
        grouped = grouped.rename(columns={factory_col: '工廠'})
        grouped = grouped.sort_values('銷貨金額', ascending=False)
        total_row = pd.DataFrame([{'工廠': '合計', '訂單數': len(month_df), '銷貨金額': month_df['_amount_usd'].sum()}])
        return pd.concat([grouped, total_row], ignore_index=True)
    
    def get_shipped_detail(self, month_str: Optional[str] = None) -> pd.DataFrame:
        """已出貨明細 (SHIPMENT)"""
        if month_str is None:
            month_str = datetime.now().strftime('%Y-%m')
        df = self.normalized_df.copy()
        month_df = df[df['_month'] == month_str].copy()
        shipped_mask = month_df['_wip'].str.contains('SHIPMENT', na=False)
        return month_df[shipped_mask].copy()
    
    def get_pending_detail(self, month_str: Optional[str] = None) -> pd.DataFrame:
        """預計出貨明細 (未 SHIPMENT)"""
        if month_str is None:
            month_str = datetime.now().strftime('%Y-%m')
        df = self.normalized_df.copy()
        month_df = df[df['_month'] == month_str].copy()
        shipped_mask = month_df['_wip'].str.contains('SHIPMENT', na=False)
        return month_df[~shipped_mask].copy()


# ====================== Streamlit 儀表板 ======================
def render_sales_detail_dashboard(orders_df: pd.DataFrame, default_month: Optional[str] = None):
    """完整渲染業績明細表（完全依照 PDF 截圖與附件格式）"""
    if orders_df is None or orders_df.empty:
        st.error("❌ 無訂單資料可顯示")
        return
    
    analyzer = SalesDetailAnalyzer(orders_df)
    
    st.title("📊 業績明細表")
    st.caption("資料來源：Teable API｜已出貨判斷：WIP 包含 SHIPMENT｜單位：USD")
    
    # 月份選擇
    months = sorted(analyzer.normalized_df['_month'].unique(), reverse=True)
    if not months:
        st.warning("目前無可用月份資料")
        return
    
    selected_month = st.selectbox(
        "選擇月份",
        months,
        index=0 if default_month is None else (months.index(default_month) if default_month in months else 0)
    )
    
    # 當月四項指標（完全匹配 PDF 第1頁格式）
    summary = analyzer.get_month_summary(selected_month)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("接單金額 (USD)", f"${summary['order_usd']:,.2f}")
    with col2:
        st.metric("已出貨金額 (USD)", f"${summary['shipped_usd']:,.2f}")
    with col3:
        st.metric("預計本月出貨 (USD)", f"${summary['pending_usd']:,.2f}")
    with col4:
        st.metric("月銷售合計 (USD)", f"${summary['total_usd']:,.2f}")
    
    # 已出貨明細與預計出貨明細（完全匹配 PDF 第3頁格式）
    st.subheader("✅ 已出貨明細 (SHIPMENT)")
    shipped_df = analyzer.get_shipped_detail(selected_month)
    if shipped_df.empty:
        st.info("本月尚無已出貨 (SHIPMENT) 資料。")
    else:
        st.dataframe(shipped_df, use_container_width=True, height=300)
    
    st.subheader("🔜 預計出貨明細 (未 SHIPMENT)")
    pending_df = analyzer.get_pending_detail(selected_month)
    if pending_df.empty:
        st.info("本月無預計出貨資料。")
    else:
        st.dataframe(pending_df, use_container_width=True, height=300)
    
    # 依工廠別統計（完全匹配附件圖片）
    st.subheader("🏭 依工廠別統計")
    factory_df = analyzer.get_factory_summary(selected_month)
    st.dataframe(factory_df, use_container_width=True, hide_index=True)
    
    # 近12個月趨勢圖（匹配 PDF 第4頁）
    st.subheader("📈 近 12 個月月銷貨趨勢")
    trend_df = analyzer.get_monthly_trend(12)
    
    fig = go.Figure()
    fig.add_trace(go.Bar(x=trend_df['month'], y=trend_df['total_usd'], name='月銷售合計 (USD)', marker_color='#1f77b4'))
    fig.add_trace(go.Bar(x=trend_df['month'], y=trend_df['shipped_usd'], name='已出貨金額 (USD)', marker_color='#ff7f0e'))
    fig.update_layout(
        barmode='group',
        height=400,
        xaxis_title="月份",
        yaxis_title="金額 (USD)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # 匯出按鈕
    month_df = analyzer.normalized_df[analyzer.normalized_df['_month'] == selected_month].copy()
    csv = month_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 下載本月完整明細 CSV",
        data=csv,
        file_name=f"業績明細_{selected_month}.csv",
        mime="text/csv"
    )


# ====================== 使用方式（請直接複製到 app.py） ======================
"""
# 在 app.py 中加入：
from reports import render_sales_detail_dashboard

if menu == "業績明細表":
    render_sales_detail_dashboard(orders)   # orders = load_orders() 回傳的 DataFrame
"""
