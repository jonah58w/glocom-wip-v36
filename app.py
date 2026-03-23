-- coding: utf-8 --

import inspect
import os
import pandas as pd
import streamlit as st
import config as cfg
import utils as u
import teable_api
import reports
from sales_report import render_sales_report_page

================================


STYLE


================================

st.markdown(
"""

""",
unsafe_allow_html=True,
)

================================


HELPERS


================================

def refresh_after_update():
st.cache_data.clear()
for k in list(st.session_state.keys()):
del st.session_state[k]
st.rerun()
def split_tags(value):
return u.split_tags(value)
def wip_display_html(value: str) -> str:
text = u.safe_text(value)
lower = text.lower()
```
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
 
```
def show_metrics(df: pd.DataFrame, wip_col: str | None):
total_orders = len(df)
shipping = 0
```
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
 
```
def show_no_data_layout():
c1, c2, c3 = st.columns(3)
c1.metric("Total Orders", 0)
c2.metric("Production", 0)
c3.metric("Shipping", 0)
st.divider()
st.warning("No data from Teable")
def customer_portal_columns(df, po_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col):
return [
c for c in [
po_col,
part_col,
qty_col,
wip_col,
ship_date_col,
customer_tag_col,
remark_col,
]
if c and c in df.columns
]
def call_report_function(possible_names, **kwargs):
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
if not os.path.exists(path):
raise FileNotFoundError(f"找不到銷貨底檔案: {path}")
```
xls = pd.ExcelFile(path)
 
if "主表" not in xls.sheet_names:
    raise ValueError(f"Excel 缺少工作表: 主表，現有工作表: {xls.sheet_names}")
 
main_df = pd.read_excel(path, sheet_name="主表")
shipment_df = pd.read_excel(path, sheet_name="銷貨底") if "銷貨底" in xls.sheet_names else pd.DataFrame()
return main_df, shipment_df
 
```

================================


LOAD DATA


================================

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

================================


DETECT KEY COLUMNS


================================

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

================================


CUSTOMER MODE


================================

query = st.query_params
customer_param = query.get("customer", None)
if customer_param:
st.title("GLOCOM Order Status")
st.caption("Customer WIP Progress")
```
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
    cust_orders,
    po_col,
    part_col,
    qty_col,
    wip_col,
    ship_date_col,
    customer_tag_col,
    remark_col,
)
csv_data = cust_orders[portal_cols].to_csv(index=False).encode("utf-8-sig")
 
st.download_button(
    "Download WIP CSV",
    data=csv_data,
    file_name=f"{customer_param}_wip.csv",
    mime="text/csv",
)
st.stop()
 
```

================================


INTERNAL MODE


================================

st.title("🏭 GLOCOM Control Tower")
st.caption("Internal PCB Production Monitoring System")
with st.expander("Debug"):
st.write("API Status:", api_status)
st.write("TABLE_URL:", cfg.TABLE_URL)
st.write("Token loaded:", bool(cfg.TEABLE_TOKEN))
st.write("Columns:", list(orders.columns) if not orders.empty else [])
st.write("SALES_BASE_PATH:", SALES_BASE_PATH)
st.write("Sales workbook path exists:", os.path.exists(SALES_BASE_PATH))
st.write("sales_df loaded:", bool(isinstance(sales_df, pd.DataFrame) and not sales_df.empty))
st.write("sales_shipment_df loaded:", bool(isinstance(sales_shipment_df, pd.DataFrame) and not sales_shipment_df.empty))
if isinstance(sales_df, pd.DataFrame) and not sales_df.empty:
st.write("sales_df columns:", list(sales_df.columns))
st.dataframe(sales_df.head(5), use_container_width=True)
if sales_error_text:
st.error(f"銷貨底 Excel 載入失敗: {sales_error_text}")
if isinstance(api_text, str):
st.text(api_text[:1200])
if orders.empty:
show_no_data_layout()
st.stop()
st.sidebar.title("GLOCOM Internal")
st.sidebar.link_button("Open Teable", cfg.TEABLE_WEB_URL, use_container_width=True)
menu = st.sidebar.radio(
"功能選單",
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
"Import / Update",
],
)
if st.sidebar.button("Refresh"):
refresh_after_update()
st.sidebar.markdown("---")
st.sidebar.caption("完成案件請在 Teable 主 View 設定篩選：WIP ≠ 完成")
st.sidebar.caption("另建 Completed View：WIP = 完成")

================================


IMPORT / UPDATE FALLBACK


================================

def fallback_import_update():
st.subheader("Import / Update")
st.caption("工廠進度輸入工具：檔案上傳、圖片截圖、貼上文字、手工輸入。")
```
tab1, tab2, tab3, tab4 = st.tabs(["檔案上傳", "圖片截圖", "貼上文字", "手工輸入"])
 
with tab1:
    uploaded = st.file_uploader(
        "上傳工廠進度檔案",
        type=["xlsx", "xls", "csv", "txt"],
        key="factory_upload_file",
    )
    if uploaded is not None:
        name = uploaded.name.lower()
        try:
            if name.endswith((".xlsx", ".xls")):
                df_up = pd.read_excel(uploaded)
                st.success(f"已讀取 {len(df_up)} 筆")
                st.dataframe(df_up, use_container_width=True, height=420)
            elif name.endswith(".csv"):
                df_up = pd.read_csv(uploaded)
                st.success(f"已讀取 {len(df_up)} 筆")
                st.dataframe(df_up, use_container_width=True, height=420)
            else:
                text = uploaded.getvalue().decode("utf-8", errors="ignore")
                st.text_area("文字內容", value=text, height=260, key="factory_text_preview")
        except Exception as e:
            st.error(f"讀取失敗：{e}")
 
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
            text = pytesseract.image_to_string(img, lang="eng")
            st.text_area("OCR 辨識文字", value=text, height=260, key="factory_ocr_text")
        except Exception:
            st.info("目前環境未啟用 OCR，可改用下方『貼上文字』或『手工輸入』。")
 
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
        "下載手工輸入 CSV",
        data=edited.to_csv(index=False).encode("utf-8-sig"),
        file_name="factory_manual_input.csv",
        mime="text/csv",
    )
 
```

================================


ROUTING


================================

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
elif menu == "Import / Update":
ok = call_report_function(["show_import_update_page", "show_import_update_report"], **common_kwargs)
if ok is False:
fallback_import_update()
st.caption("Auto refresh cache: 60 seconds")
