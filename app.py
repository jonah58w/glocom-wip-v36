# -*- coding: utf-8 -*-
"""
GLOCOM Control Tower - PCB Production Monitoring System
客戶端 WIP 查詢 + 內部生產監控 + 工廠進度批量更新
"""
from __future__ import annotations

import inspect
import os
from typing import Optional, List, Any, Dict
import pandas as pd
import streamlit as st
import config as cfg
import utils as u
import teable_api
import reports
from sales_report import render_sales_report_page

# ==================================
# PAGE CONFIG
# ==================================
st.set_page_config(
    page_title="GLOCOM Control Tower",
    page_icon="🏭",
    layout="wide",
)

# ==================================
# STYLE
# ==================================
st.markdown(
    """
    <style>
    .portal-box {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        padding: 20px;
        margin: 12px 0;
        color: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .portal-title {
        font-size: 1.2em;
        font-weight: bold;
        margin-bottom: 12px;
        padding-bottom: 8px;
        border-bottom: 2px solid rgba(255,255,255,0.3);
    }
    .wip-chip {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 16px;
        font-size: 0.85em;
        font-weight: 600;
    }
    .tag-chip {
        display: inline-block;
        background: rgba(255,255,255,0.2);
        padding: 2px 8px;
        border-radius: 8px;
        margin: 2px;
        font-size: 0.8em;
    }
    .update-success { background: #065f46; color: #d1fae5; }
    .update-fail { background: #7f1d1d; color: #fee2e2; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==================================
# HELPERS
# ==================================

def refresh_after_update() -> None:
    """清除緩存並重新運行"""
    st.cache_data.clear()
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.rerun()

def split_tags(value: Any) -> List[str]:
    """分割標籤"""
    return u.split_tags(value)

def wip_display_html(value: str) -> str:
    """生成 WIP 狀態 HTML"""
    text = u.safe_text(value)
    lower = text.lower()
    
    if any(k in lower for k in ["完成"]) or text.upper() in getattr(cfg, "DONE_WIP_VALUES", []):
        label = text or "完成"
        bg = "#065f46"
        fg = "#d1fae5"
    elif any(k in lower for k in ["ship", "shipping", "出貨"]):
        label = text or "Shipping"
        bg = "#14532d"
        fg = "#dcfce7"
    elif any(k in lower for k in ["pack", "包裝"]):
        label = text or "Packing"
        bg = "#166534"
        fg = "#dcfce7"
    elif any(k in lower for k in ["fqc", "qa", "inspection", "成檢", "測試"]):
        label = text or "Inspection"
        bg = "#854d0e"
        fg = "#fef3c7"
    elif any(
        k in lower
        for k in [
            "aoi", "drill", "route", "routing", "plating", "inner", "production",
            "防焊", "壓合", "外層", "內層", "成型"
        ]
    ):
        label = text or "Production"
        bg = "#9a3412"
        fg = "#ffedd5"
    elif any(k in lower for k in ["eng", "gerber", "cam", "eq"]):
        label = text or "Engineering"
        bg = "#1d4ed8"
        fg = "#dbeafe"
    elif any(k in lower for k in ["hold", "等待", "暫停"]):
        label = text or "On Hold"
        bg = "#7f1d1d"
        fg = "#fee2e2"
    else:
        label = text or "-"
        bg = "#374151"
        fg = "#f3f4f6"
    
    return f'<span class="wip-chip" style="background:{bg};color:{fg};">{label}</span>'

def show_metrics(df: pd.DataFrame, wip_col: Optional[str]) -> None:
    """顯示指標卡片"""
    total_orders = len(df)
    shipping = 0
    
    if wip_col and wip_col in df.columns:
        wip_series = u.get_series_by_col(df, wip_col)
        if wip_series is not None:
            shipping = len(
                df[
                    wip_series.astype(str).str.contains(
                        "ship|shipping|出貨",
                        case=False,
                        na=False,
                    )
                ]
            )
    
    production = total_orders - shipping
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Orders", total_orders)
    c2.metric("Production", production)
    c3.metric("Shipping", shipping)

def show_no_data_layout() -> None:
    """顯示無數據佈局"""
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Orders", 0)
    c2.metric("Production", 0)
    c3.metric("Shipping", 0)
    st.divider()
    st.warning("No data from Teable")

def customer_portal_columns(
    df: pd.DataFrame,
    po_col: Optional[str],
    part_col: Optional[str],
    qty_col: Optional[str],
    wip_col: Optional[str],
    ship_date_col: Optional[str],
    customer_tag_col: Optional[str],
    remark_col: Optional[str]
) -> List[str]:
    """獲取客戶門戶顯示的列"""
    return [
        c for c in [
            po_col, part_col, qty_col, wip_col,
            ship_date_col, customer_tag_col, remark_col,
        ]
        if c and c in df.columns
    ]

def call_report_function(possible_names: List[str], **kwargs: Any) -> Any:
    """調用報告函數"""
    for name in possible_names:
        func = getattr(reports, name, None)
        if callable(func):
            sig = inspect.signature(func)
            has_var_kw = any(p.kind == inspect.Parameter.VAR_KEYWORD for p in sig.parameters.values())
            if has_var_kw:
                return func(**kwargs)
            accepted = {k: v for k, v in kwargs.items() if k in sig.parameters}
            return func(**accepted)
    return False

@st.cache_data(ttl=60)
def load_sales_workbook(path: str):
    """載入銷貨底工作簿"""
    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到銷貨底檔案: {path}")
    
    xls = pd.ExcelFile(path)
    
    if "主表" not in xls.sheet_names:
        raise ValueError(f"Excel 缺少工作表: 主表，現有工作表: {xls.sheet_names}")
    
    main_df = pd.read_excel(path, sheet_name="主表")
    shipment_df = pd.read_excel(path, sheet_name="銷貨底") if "銷貨底" in xls.sheet_names else pd.DataFrame()
    return main_df, shipment_df

# ==================================
# LOAD DATA
# ==================================

try:
    orders, api_status, api_text = teable_api.load_orders()
except Exception as e:
    orders = pd.DataFrame()
    api_status = "EXCEPTION"
    api_text = str(e)

SALES_BASE_PATH = "Sandy需要的銷貨底.xlsx"
try:
    sales_df, sales_shipment_df = load_sales_workbook(SALES_BASE_PATH)
    sales_error_text = ""
except Exception as e:
    sales_df, sales_shipment_df = pd.DataFrame(), pd.DataFrame()
    sales_error_text = str(e)

# ==================================
# DETECT KEY COLUMNS
# ==================================

po_col = u.get_first_matching_column(orders, cfg.PO_CANDIDATES)
customer_col = u.get_first_matching_column(orders, cfg.CUSTOMER_CANDIDATES)
part_col = u.get_first_matching_column(orders, cfg.PART_CANDIDATES)
qty_col = u.get_first_matching_column(orders, cfg.QTY_CANDIDATES)
factory_col = u.get_first_matching_column(orders, cfg.FACTORY_CANDIDATES)
wip_col = u.get_first_matching_column(orders, cfg.WIP_CANDIDATES)
factory_due_col = u.get_first_matching_column(orders, cfg.FACTORY_DUE_CANDIDATES)
ship_date_col = u.get_first_matching_column(orders, cfg.SHIP_DATE_CANDIDATES)
remark_col = u.get_first_matching_column(orders, cfg.REMARK_CANDIDATES)
customer_tag_col = u.get_first_matching_column(orders, cfg.CUSTOMER_TAG_CANDIDATES)
merge_date_col = u.get_first_matching_column(orders, cfg.MERGE_DATE_CANDIDATES)
order_date_col = u.get_first_matching_column(orders, cfg.ORDER_DATE_CANDIDATES)
factory_order_date_col = u.get_first_matching_column(orders, cfg.FACTORY_ORDER_DATE_CANDIDATES)
changed_due_date_col = u.get_first_matching_column(orders, cfg.CHANGED_DUE_DATE_CANDIDATES)

# ==================================
# CUSTOMER MODE
# ==================================

query = st.query_params
customer_param = query.get("customer", None)

if customer_param:
    st.title("GLOCOM Order Status")
    st.caption("Customer WIP Progress")
    
    if not customer_col:
        st.error("Customer column not found")
        st.stop()
    
    customer_series = u.get_series_by_col(orders, customer_col)
    if customer_series is None:
        st.error("Customer data unavailable")
        st.stop()
    
    cust_orders = orders[
        customer_series.astype(str).str.strip().str.lower()
        == str(customer_param).strip().lower()
    ].copy()
    
    if cust_orders.empty:
        st.warning("No orders found")
        st.stop()
    
    show_metrics(cust_orders, wip_col)
    st.divider()
    
    for _, row in cust_orders.iterrows():
        po_val = u.safe_text(row.get(po_col, "")) if po_col else ""
        part_val = u.safe_text(row.get(part_col, "")) if part_col else ""
        qty_val = u.safe_text(row.get(qty_col, "")) if qty_col else ""
        wip_val = u.safe_text(row.get(wip_col, "")) if wip_col else ""
        ship_val = u.safe_text(row.get(ship_date_col, "")) if ship_date_col else ""
        remark_val = u.safe_text(row.get(remark_col, "")) if remark_col else ""
        tags_val = split_tags(row.get(customer_tag_col, "")) if customer_tag_col else []
        
        tag_html = "".join([f'<span class="tag-chip">{t}</span>' for t in tags_val]) if tags_val else '<span class="tag-chip">-</span>'
        
        st.markdown(
            f"""
            <div class="portal-box">
                <div class="portal-title">{po_val or "-"}</div>
                <div style="display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;">
                    <div>
                        <div><strong>P/N</strong> : {part_val or "-"}</div>
                        <div><strong>Qty</strong> : {qty_val or "-"}</div>
                    </div>
                    <div>
                        <div><strong>WIP</strong> : {wip_display_html(wip_val)}</div>
                        <div style="margin-top:8px;"><strong>Ship Date</strong> : {ship_val or "-"}</div>
                    </div>
                </div>
                <div style="margin-top:12px;"><strong>Customer Remark Tags</strong> : {tag_html}</div>
                <div style="margin-top:10px;"><strong>Remark</strong> : {remark_val or "-"}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    
    portal_cols = customer_portal_columns(
        cust_orders, po_col, part_col, qty_col, wip_col,
        ship_date_col, customer_tag_col, remark_col,
    )
    csv_data = cust_orders[portal_cols].to_csv(index=False).encode("utf-8-sig")
    
    st.download_button(
        "Download WIP CSV",
        data=csv_data,
        file_name=f"{customer_param}_wip.csv",
        mime="text/csv",
    )
    st.stop()

# ==================================
# INTERNAL MODE
# ==================================

st.title("🏭 GLOCOM Control Tower")
st.caption("Internal PCB Production Monitoring System")

with st.expander("🔧 Debug Info"):
    st.write("API Status:", api_status)
    st.write("TABLE_URL:", cfg.TABLE_URL)
    st.write("Token loaded:", bool(cfg.TEABLE_TOKEN))
    st.write("Columns:", list(orders.columns) if not orders.empty else [])
    st.write("SALES_BASE_PATH:", SALES_BASE_PATH)
    st.write("Sales workbook exists:", os.path.exists(SALES_BASE_PATH))
    if sales_error_text:
        st.error(f"銷貨底 Excel 載入失敗: {sales_error_text}")
    if isinstance(api_text, str):
        st.text(api_text[:800])

if orders.empty:
    show_no_data_layout()
    st.stop()

st.sidebar.title("GLOCOM Internal")
st.sidebar.link_button("🔗 Open Teable", cfg.TEABLE_WEB_URL, use_container_width=True)

menu = st.sidebar.radio(
    "📋 功能選單",
    [
        "Dashboard",
        "Factory Load",
        "Delayed Orders",
        "Shipment Forecast",
        "Orders",
        "Customer Preview",
        "Sandy 內部 WIP",
        "Sandy 銷貨底",
        "新訂單 WIP",
        "業績明細表",
        "📤 Import / Update",
    ],
)

if st.sidebar.button("🔄 Refresh"):
    refresh_after_update()

st.sidebar.markdown("---")
st.sidebar.caption("✅ 完成案件請在 Teable 主 View 設定篩選：WIP ≠ 完成")
st.sidebar.caption("📁 另建 Completed View：WIP = 完成")

# ==================================
# IMPORT / UPDATE WITH BATCH SYNC
# ==================================

def _detect_factory_name(filename: str) -> str:
    """從檔案名識別工廠"""
    name_lower = filename.lower()
    if '全興' in name_lower or 'quanxing' in name_lower or '203-' in name_lower:
        return "全興電子"
    elif 'profit' in name_lower or 'pg' in name_lower or 'glocom-pg' in name_lower:
        return "Profit Grand"
    elif '祥竑' in name_lower or 'xianghong' in name_lower:
        return "祥竑電子"
    elif '西拓' in name_lower or 'xituo' in name_lower:
        return "西拓電子"
    elif 'star' in name_lower or '星辰' in name_lower or '115' in name_lower:
        return "星晨電路"
    return "未知工廠"


def _display_update_results(results: Dict[str, Any]) -> None:
    """顯示更新結果"""
    # 匯總統計
    c1, c2, c3 = st.columns(3)
    c1.metric("✅ 成功", results['success_count'])
    c2.metric("❌ 失敗", results['failed_count'])
    c3.metric("⚠️ 警告", len(results.get('warnings', [])))
    
    # 顯示警告
    if results.get('warnings'):
        with st.expander("⚠️ 處理警告"):
            for w in results['warnings']:
                st.warning(w)
    
    # 顯示失敗詳情
    failed = [d for d in results['details'] if d.get('error')]
    if failed:
        with st.expander("❌ 失敗明細"):
            for item in failed[:10]:
                st.error(f"行 {item.get('row')}: PO={item.get('po')} - {item.get('error')}")
        if len(failed) > 10:
            st.caption(f"... 還有 {len(failed) - 10} 筆失敗記錄")
    
    # 顯示成功詳情
    success = [d for d in results['details'] if d.get('status') == '更新成功']
    if success:
        with st.expander("✅ 成功明細 (點擊展開)"):
            for item in success[:20]:
                st.success(f"✓ PO: {item.get('po')} → WIP: {item.get('wip')}")


def fallback_import_update() -> None:
    """後備導入/更新功能 - 支援批量更新到 Teable"""
    st.subheader("📤 Import / Update")
    st.caption("工廠進度輸入工具：檔案上傳、圖片截圖、貼上文字、手工輸入 + 批量同步到 Teable")
    
    tab1, tab2, tab3, tab4 = st.tabs(["📁 檔案上傳", "🖼️ 圖片截圖", "📋 貼上文字", "✏️ 手工輸入"])
    
    with tab1:
        uploaded = st.file_uploader(
            "上傳工廠進度檔案",
            type=["xlsx", "xls", "csv", "txt"],
            key="factory_upload_file",
        )
        
        if uploaded is not None:
            name = uploaded.name.lower()
            factory_name = _detect_factory_name(name)
            st.caption(f"🏭 識別工廠: **{factory_name}**")
            
            try:
                # 讀取數據
                if name.endswith((".xlsx", ".xls")):
                    df_up = pd.read_excel(uploaded)
                elif name.endswith(".csv"):
                    df_up = pd.read_csv(uploaded)
                else:
                    text = uploaded.getvalue().decode("utf-8", errors="ignore")
                    st.text_area("文字內容", value=text, height=260, key="factory_text_preview")
                    st.stop()
                
                # 顯示預覽
                st.success(f"✅ 已讀取 **{len(df_up)}** 筆記錄")
                st.dataframe(df_up.head(10), use_container_width=True, height=300)
                
                # 欄位匹配預覽
                with st.expander("📋 欄位匹配預覽"):
                    po_match = u.get_first_matching_column(df_up, cfg.PO_CANDIDATES)
                    wip_match = u.get_first_matching_column(df_up, cfg.WIP_CANDIDATES)
                    process_cols = teable_api._detect_process_columns(df_up)
                    st.write(f"**PO 列**: `{po_match}`" if po_match else "**PO 列**: ❌ 未找到")
                    st.write(f"**WIP 列**: `{wip_match}`" if wip_match else "**WIP 列**: ℹ️ 將嘗試多列製程解析")
                    if process_cols:
                        st.write(f"**製程列檢測**: {len(process_cols)} 個 → {process_cols[:5]}...")
                
                # 更新按鈕
                col1, col2 = st.columns([1, 4])
                with col1:
                    if st.button("📤 更新到 Teable", type="primary", key="update_teable_btn"):
                        with st.spinner("🔄 正在批量更新到 Teable..."):
                            # 調用批量更新函數
                            results = teable_api.batch_update_wip_from_excel(
                                current_df=orders,
                                uploaded_df=df_up,
                                factory_name=factory_name
                            )
                            
                            # 顯示結果
                            _display_update_results(results)
                            
                            # 自動刷新緩存
                            if results['success_count'] > 0:
                                st.balloons()
                                st.success("✨ 更新完成！頁面將自動刷新...")
                                refresh_after_update()
                
            except Exception as e:
                st.error(f"❌ 讀取失敗：{e}")
                import traceback
                with st.expander("🔍 錯誤詳情"):
                    st.code(traceback.format_exc())
    
    with tab2:
        image_file = st.file_uploader(
            "上傳截圖 / 圖片",
            type=["png", "jpg", "jpeg", "webp"],
            key="factory_image_upload",
        )
        if image_file is not None:
            st.image(image_file, caption="已上傳截圖")
            try:
                import pytesseract
                from PIL import Image
                
                img = Image.open(image_file)
                text = pytesseract.image_to_string(img, lang="eng+chi_tra")
                st.text_area("OCR 辨識文字", value=text, height=260, key="factory_ocr_text")
            except Exception:
                st.info("ℹ️ 目前環境未啟用 OCR，可改用『貼上文字』或『手工輸入』。")
    
    with tab3:
        pasted = st.text_area("貼上工廠進度文字", height=280, key="factory_pasted_text")
        if pasted.strip():
            st.text_area("預覽", value=pasted, height=280, key="factory_pasted_preview")
    
    with tab4:
        rows = st.number_input("手工輸入列數", min_value=1, max_value=50, value=5, step=1, key="factory_manual_rows")
        manual_df = pd.DataFrame(
            {
                "PO#": [""] * rows,
                "Customer": [""] * rows,
                "P/N": [""] * rows,
                "QTY": [""] * rows,
                "WIP": [""] * rows,
                "Ship date": [""] * rows,
                "Remark": [""] * rows,
            }
        )
        edited = st.data_editor(
            manual_df,
            use_container_width=True,
            num_rows="fixed",
            key="factory_manual_editor",
        )
        st.download_button(
            "⬇️ 下載手工輸入 CSV",
            data=edited.to_csv(index=False).encode("utf-8-sig"),
            file_name="factory_manual_input.csv",
            mime="text/csv",
        )

# ==================================
# ROUTING
# ==================================

common_kwargs = dict(
    df=orders,
    orders=orders,
    sales_df=sales_df,
    sales_shipment_df=sales_shipment_df,
    po_col=po_col,
    customer_col=customer_col,
    part_col=part_col,
    qty_col=qty_col,
    factory_col=factory_col,
    wip_col=wip_col,
    factory_due_col=factory_due_col,
    ship_date_col=ship_date_col,
    remark_col=remark_col,
    customer_tag_col=customer_tag_col,
    merge_date_col=merge_date_col,
    order_date_col=order_date_col,
    factory_order_date_col=factory_order_date_col,
    changed_due_date_col=changed_due_date_col,
)

if menu == "Dashboard":
    ok = call_report_function(["show_dashboard_report"], **common_kwargs)
    if ok is False:
        show_metrics(orders, wip_col)
elif menu == "Factory Load":
    ok = call_report_function(["show_factory_load_report"], **common_kwargs)
    if ok is False:
        st.info("Factory Load fallback not enabled.")
elif menu == "Delayed Orders":
    ok = call_report_function(["show_delayed_orders_report"], **common_kwargs)
    if ok is False:
        st.info("Delayed Orders fallback not enabled.")
elif menu == "Shipment Forecast":
    ok = call_report_function(["show_shipment_forecast_report"], **common_kwargs)
    if ok is False:
        st.info("Shipment Forecast fallback not enabled.")
elif menu == "Orders":
    ok = call_report_function(["show_orders_report"], **common_kwargs)
    if ok is False:
        st.dataframe(orders, use_container_width=True)
elif menu == "Customer Preview":
    ok = call_report_function(["show_customer_preview_report"], **common_kwargs)
    if ok is False:
        st.info("Customer Preview fallback not enabled.")
elif menu == "Sandy 內部 WIP":
    ok = call_report_function(["show_sandy_internal_wip_report"], **common_kwargs)
    if ok is False:
        st.info("Sandy 內部 WIP fallback not enabled.")
elif menu == "Sandy 銷貨底":
    ok = call_report_function(["show_sandy_shipment_report"], **common_kwargs)
    if ok is False:
        st.info("Sandy 銷貨底 fallback not enabled.")
elif menu == "新訂單 WIP":
    ok = call_report_function(["show_new_orders_wip_report"], **common_kwargs)
    if ok is False:
        st.info("新訂單 WIP fallback not enabled.")
elif menu == "業績明細表":
    render_sales_report_page(**common_kwargs)
elif menu == "📤 Import / Update":
    ok = call_report_function(["show_import_update_page", "show_import_update_report"], **common_kwargs)
    if ok is False:
        fallback_import_update()

st.caption("🔄 Auto refresh cache: 60 seconds | 📊 Last updated: " + pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"))
