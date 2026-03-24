# -*- coding: utf-8 -*-
"""
GLOCOM Control Tower - PCB Production Monitoring System
客戶端 WIP 查詢 + 內部生產監控 + 工廠進度批量更新
"""

from __future__ import annotations

import inspect
import os
from typing import Optional, List, Any, Dict, Tuple

import pandas as pd
import streamlit as st

import config as cfg
import utils as u
import teable_api
import reports
from sales_report import render_sales_report_page
from factory_parsers import read_import_dataframe


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
        background: rgba(255,255,255,0.12);
        padding: 2px 8px;
        border-radius: 8px;
        margin: 2px;
        font-size: 0.8em;
    }
    .small-muted { color: #9ca3af; font-size: 0.9em; }
    .debug-box {
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 10px;
        padding: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==================================
# HELPERS
# ==================================
def refresh_after_update() -> None:
    try:
        st.cache_data.clear()
    except Exception:
        pass
    st.rerun()


def split_tags(value: Any) -> List[str]:
    if hasattr(u, "split_tags"):
        return u.split_tags(value)
    if value is None:
        return []
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    text = str(value).strip()
    if not text:
        return []
    for sep in ["；", ";", "，", "/", "|"]:
        text = text.replace(sep, ",")
    return [x.strip() for x in text.split(",") if x.strip()]


def safe_text(value: Any) -> str:
    if hasattr(u, "safe_text"):
        return u.safe_text(value)
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def get_first_matching_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if hasattr(u, "get_first_matching_column"):
        return u.get_first_matching_column(df, candidates)
    if df is None or df.empty:
        return None
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _cfg_list(name: str, default: Optional[List[str]] = None) -> List[str]:
    value = getattr(cfg, name, None)
    if isinstance(value, list):
        return value
    return default or []


def wip_display_html(value: str) -> str:
    text = safe_text(value)
    lower = text.lower()
    done_values = {str(x).upper() for x in getattr(cfg, "DONE_WIP_VALUES", set())}

    if any(k in lower for k in ["完成"]) or text.upper() in done_values:
        label, bg, fg = text or "完成", "#065f46", "#d1fae5"
    elif any(k in lower for k in ["ship", "shipping", "出貨"]):
        label, bg, fg = text or "Shipping", "#14532d", "#dcfce7"
    elif any(k in lower for k in ["pack", "包裝"]):
        label, bg, fg = text or "Packing", "#1d4ed8", "#dbeafe"
    elif any(k in lower for k in ["inspec", "qa", "fqc", "iqc", "檢", "測試"]):
        label, bg, fg = text or "Inspection", "#92400e", "#fef3c7"
    elif any(k in lower for k in ["eng", "gerber", "工程"]):
        label, bg, fg = text or "Engineering", "#6d28d9", "#ede9fe"
    elif any(k in lower for k in ["hold", "暫停", "等待"]):
        label, bg, fg = text or "On Hold", "#7f1d1d", "#fee2e2"
    else:
        label, bg, fg = text or "Production", "#0f766e", "#ccfbf1"

    return f'<span class="wip-chip" style="background:{bg};color:{fg};">{label}</span>'


def show_no_data_layout(api_status: Any = "", api_text: Any = "") -> None:
    st.markdown(
        """
        <div class="portal-box">
            <div class="portal-title">GLOCOM Control Tower</div>
            <div>目前未讀取到 Teable 主資料，但系統仍可使用 Import / Update。</div>
            <div style="margin-top:8px;">請確認：</div>
            <ul>
                <li>Streamlit secrets 已設定 TEABLE_TOKEN</li>
                <li>TABLE_URL / TEABLE_TABLE_URL 是否正確</li>
                <li>Teable API 是否可正常連線</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if api_status or api_text:
        with st.expander("🔍 Teable 連線回傳"):
            st.write("api_status:", api_status)
            st.text(str(api_text)[:2000])


def normalize_orders_result(raw: Any) -> Tuple[pd.DataFrame, str, str]:
    """
    兼容 teable_api.load_orders() 的多種回傳型態：
    1) DataFrame
    2) (df, status, text)
    3) dict with keys: df/orders/data, status/api_status, text/api_text/message
    4) 其他型態 -> 空 DataFrame + 說明
    """
    empty_df = pd.DataFrame()

    if isinstance(raw, pd.DataFrame):
        return raw.copy(), "ok", ""

    if isinstance(raw, tuple):
        df = empty_df
        status = ""
        text = ""

        if len(raw) >= 1 and isinstance(raw[0], pd.DataFrame):
            df = raw[0].copy()
        if len(raw) >= 2:
            status = safe_text(raw[1])
        if len(raw) >= 3:
            text = safe_text(raw[2])

        return df, status, text

    if isinstance(raw, dict):
        df = raw.get("df")
        if not isinstance(df, pd.DataFrame):
            df = raw.get("orders")
        if not isinstance(df, pd.DataFrame):
            df = raw.get("data")
        if not isinstance(df, pd.DataFrame):
            df = empty_df

        status = safe_text(raw.get("status") or raw.get("api_status") or "")
        text = safe_text(raw.get("text") or raw.get("api_text") or raw.get("message") or "")
        return df.copy() if isinstance(df, pd.DataFrame) else empty_df, status, text

    return empty_df, "unknown", f"Unsupported load_orders() return type: {type(raw)}"


def show_metrics(orders: pd.DataFrame, wip_col: Optional[str]) -> None:
    st.subheader("📊 Dashboard")

    if orders is None or orders.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("總筆數", 0)
        c2.metric("進行中", 0)
        c3.metric("完成", 0)
        c4.metric("On Hold", 0)
        st.info("目前沒有資料。")
        return

    total = len(orders)
    done_count = 0
    hold_count = 0

    if wip_col and wip_col in orders.columns:
        s = orders[wip_col].fillna("").astype(str).str.strip()
        done_values = {str(x).upper() for x in getattr(cfg, "DONE_WIP_VALUES", set())}
        done_count = int((s.str.upper().isin(done_values) | s.str.contains("完成", na=False)).sum())
        hold_count = int(s.str.contains("hold|暫停|等待", case=False, na=False).sum())

    active = max(total - done_count - hold_count, 0)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("總筆數", total)
    c2.metric("進行中", active)
    c3.metric("完成", done_count)
    c4.metric("On Hold", hold_count)

    if wip_col and wip_col in orders.columns:
        st.markdown("### WIP 分布")
        vc = (
            orders[wip_col]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        vc = vc[vc != ""].value_counts().head(20)
        if not vc.empty:
            st.bar_chart(vc)
        else:
            st.info("目前沒有 WIP 分布資料。")


def call_report_function(candidate_names: List[str], **kwargs) -> Optional[bool]:
    alias_map = {
        "df": ["orders", "data", "dataset"],
        "orders": ["df", "data", "dataset"],
        "data": ["df", "orders", "dataset"],
        "dataset": ["df", "orders", "data"],
    }

    for name in candidate_names:
        fn = getattr(reports, name, None)
        if callable(fn):
            try:
                sig = inspect.signature(fn)
                accepted = {}

                for k, v in kwargs.items():
                    if k in sig.parameters:
                        accepted[k] = v

                for param_name in sig.parameters:
                    if param_name in accepted:
                        continue
                    if param_name in alias_map:
                        for source_name in alias_map[param_name]:
                            if source_name in kwargs:
                                accepted[param_name] = kwargs[source_name]
                                break

                result = fn(**accepted)
                if result is False:
                    return False
                return True
            except Exception as e:
                st.error(f"{name} 執行失敗：{e}")
                return False
    return False


def _detect_factory_name(filename: str) -> str:
    name_lower = filename.lower()
    if "全興" in filename or "quanxing" in name_lower or "203-" in name_lower:
        return "全興電子"
    elif "profit" in name_lower or "glocom-pg" in name_lower or "pg" in name_lower:
        return "Profit Grand"
    elif "祥竑" in filename or "xianghong" in name_lower:
        return "祥竑電子"
    elif "西拓" in filename or "xituo" in name_lower:
        return "西拓電子"
    elif "star" in name_lower or "星辰" in filename or "115" in name_lower:
        return "星晨電路"
    elif "優技" in filename or "yoji" in name_lower:
        return "優技"
    elif "宏棋" in filename or "hongqi" in name_lower:
        return "宏棋"
    return "未知工廠"


def _display_update_results(results: Dict[str, Any]) -> None:
    c1, c2, c3 = st.columns(3)
    c1.metric("✅ 成功", results.get("success_count", 0))
    c2.metric("❌ 失敗", results.get("failed_count", 0))
    c3.metric("⚠️ 警告", len(results.get("warnings", [])))

    warnings = results.get("warnings", [])
    if warnings:
        with st.expander("⚠️ 處理警告"):
            for w in warnings:
                st.warning(w)

    failed = [d for d in results.get("details", []) if d.get("error")]
    if failed:
        with st.expander("❌ 失敗明細"):
            for item in failed[:20]:
                po_text = item.get("po") or "-"
                part_text = item.get("part") or "-"
                st.error(f"行 {item.get('row')}: PO={po_text} / PART={part_text} - {item.get('error')}")

    success = [d for d in results.get("details", []) if d.get("status") == "更新成功"]
    if success:
        with st.expander("✅ 成功明細"):
            for item in success[:30]:
                po_text = item.get("po") or "-"
                part_text = item.get("part") or "-"
                matched_by = item.get("matched_by") or "-"
                st.success(f"✓ PO: {po_text} / PART: {part_text} → WIP: {item.get('wip')} 〔match: {matched_by}〕")


def show_orders_page(orders: pd.DataFrame) -> None:
    st.subheader("📋 Orders")

    if orders is None or orders.empty:
        st.info("目前沒有資料。")
        return

    df = orders.copy()

    cols = st.columns(4)

    if customer_col and customer_col in df.columns:
        customer_options = ["全部"] + sorted([x for x in df[customer_col].dropna().astype(str).unique() if x.strip()])
        sel_customer = cols[0].selectbox("Customer", customer_options, index=0)
        if sel_customer != "全部":
            df = df[df[customer_col].astype(str) == sel_customer]

    if factory_col and factory_col in df.columns:
        factory_options = ["全部"] + sorted([x for x in df[factory_col].dropna().astype(str).unique() if x.strip()])
        sel_factory = cols[1].selectbox("Factory", factory_options, index=0)
        if sel_factory != "全部":
            df = df[df[factory_col].astype(str) == sel_factory]

    if wip_col and wip_col in df.columns:
        wip_options = ["全部"] + sorted([x for x in df[wip_col].dropna().astype(str).unique() if x.strip()])
        sel_wip = cols[2].selectbox("WIP", wip_options, index=0)
        if sel_wip != "全部":
            df = df[df[wip_col].astype(str) == sel_wip]

    keyword = cols[3].text_input("搜尋 PO / Part / Remark")

    if keyword:
        mask = pd.Series(False, index=df.index)
        for c in [po_col, part_col, remark_col]:
            if c and c in df.columns:
                mask = mask | df[c].astype(str).str.contains(keyword, case=False, na=False)
        df = df[mask]

    st.caption(f"共 {len(df)} 筆")

    display_cols = []
    for c in [customer_col, po_col, part_col, qty_col, factory_col, wip_col, factory_due_col, ship_date_col, remark_col]:
        if c and c in df.columns and c not in display_cols:
            display_cols.append(c)

    if not display_cols:
        display_cols = list(df.columns)

    st.dataframe(df[display_cols], use_container_width=True, height=560)


def show_customer_preview_page(orders: pd.DataFrame) -> None:
    st.subheader("👀 Customer Preview")

    if orders is None or orders.empty:
        st.info("目前沒有資料。")
        return

    df = orders.copy()
    qp = st.query_params.get("customer", "")
    if isinstance(qp, list):
        qp = qp[0] if qp else ""

    if customer_col and customer_col in df.columns:
        customer_options = sorted([x for x in df[customer_col].dropna().astype(str).unique() if x.strip()])
        if customer_options:
            default_ix = customer_options.index(qp) if qp in customer_options else 0
            selected_customer = st.selectbox("Customer", customer_options, index=default_ix)
            df = df[df[customer_col].astype(str) == selected_customer]

    display_cols = []
    for c in [po_col, part_col, qty_col, wip_col, ship_date_col, customer_tag_col, remark_col]:
        if c and c in df.columns and c not in display_cols:
            display_cols.append(c)

    if not display_cols:
        display_cols = list(df.columns)

    st.caption(f"顯示 {len(df)} 筆")
    st.dataframe(df[display_cols], use_container_width=True, height=560)


def fallback_import_update(orders: pd.DataFrame) -> None:
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
            factory_name = _detect_factory_name(uploaded.name)
            st.caption(f"🏭 識別工廠: **{factory_name}**")

            try:
                df_up, parse_mode = read_import_dataframe(uploaded)

                if df_up is None or df_up.empty:
                    st.warning("⚠️ 解析後沒有有效資料")
                else:
                    st.success(f"✅ 已解析 **{len(df_up)}** 筆記錄 | 模式: **{parse_mode}**")
                    st.dataframe(df_up.head(20), use_container_width=True, height=320)

                    with st.expander("📋 欄位匹配預覽"):
                        po_match = get_first_matching_column(df_up, _cfg_list("PO_CANDIDATES"))
                        part_match = get_first_matching_column(df_up, _cfg_list("PART_CANDIDATES"))
                        wip_match = get_first_matching_column(df_up, _cfg_list("WIP_CANDIDATES"))
                        process_cols = []
                        if hasattr(teable_api, "_detect_process_columns"):
                            try:
                                process_cols = teable_api._detect_process_columns(df_up)
                            except Exception:
                                process_cols = []

                        st.write(f"**PO 列**: `{po_match}`" if po_match else "**PO 列**: ❌ 未找到")
                        st.write(f"**Part / 料號列**: `{part_match}`" if part_match else "**Part / 料號列**: ❌ 未找到")
                        st.write(f"**WIP 列**: `{wip_match}`" if wip_match else "**WIP 列**: ℹ️ 將嘗試多列製程解析")

                        if process_cols:
                            st.write(f"**製程列檢測**: {len(process_cols)} 個 → {process_cols[:8]}")

                    if st.button("📤 更新到 Teable", type="primary", key="update_teable_btn"):
                        with st.spinner("🔄 正在批量更新到 Teable..."):
                            results = teable_api.batch_update_wip_from_excel(
                                current_df=orders if isinstance(orders, pd.DataFrame) else pd.DataFrame(),
                                uploaded_df=df_up,
                                factory_name=factory_name,
                            )
                            _display_update_results(results)

                            if results.get("success_count", 0) > 0:
                                st.success("✨ 更新完成，頁面刷新中...")
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
            st.image(image_file, caption="已上傳截圖", use_container_width=True)
            try:
                import pytesseract
                from PIL import Image

                img = Image.open(image_file)
                text = pytesseract.image_to_string(img, lang="eng+chi_tra")
                st.text_area("OCR 結果", value=text, height=260, key="ocr_result_preview")
            except Exception as e:
                st.error(f"OCR 失敗：{e}")

    with tab3:
        pasted = st.text_area("請貼上工廠進度文字", height=260, key="factory_text_input")
        if pasted:
            st.text_area("文字預覽", value=pasted, height=220, key="factory_text_preview_2")
            st.info("這一版先保留貼上與預覽；之後可再接文字直解析更新 Teable。")

    with tab4:
        st.info("手工輸入模式保留。下一版可再接成直接寫入 Teable。")


# ==================================
# DATA LOADERS
# ==================================
@st.cache_data(ttl=60, show_spinner=False)
def cached_load_orders():
    raw = teable_api.load_orders()
    return normalize_orders_result(raw)


@st.cache_data(ttl=60, show_spinner=False)
def cached_load_sales_data():
    sales_df = pd.DataFrame()
    sales_shipment_df = pd.DataFrame()
    sales_error_text = ""
    sales_base_path = getattr(cfg, "SALES_BASE_PATH", "")

    candidate_loaders = [
        "load_sales_data",
        "get_sales_data",
        "read_sales_data",
        "load_sales_workbook",
    ]

    for loader_name in candidate_loaders:
        loader = getattr(reports, loader_name, None)
        if callable(loader):
            try:
                result = loader()
                if isinstance(result, tuple):
                    if len(result) >= 1 and isinstance(result[0], pd.DataFrame):
                        sales_df = result[0]
                    if len(result) >= 2 and isinstance(result[1], pd.DataFrame):
                        sales_shipment_df = result[1]
                    if len(result) >= 3 and isinstance(result[2], str):
                        sales_error_text = result[2]
                elif isinstance(result, pd.DataFrame):
                    sales_df = result
                break
            except Exception as e:
                sales_error_text = str(e)
                break

    return sales_df, sales_shipment_df, sales_error_text, sales_base_path


# ==================================
# LOAD DATA
# ==================================
orders, api_status, api_text = cached_load_orders()
sales_df, sales_shipment_df, sales_error_text, SALES_BASE_PATH = cached_load_sales_data()

po_col = get_first_matching_column(orders, _cfg_list("PO_CANDIDATES"))
customer_col = get_first_matching_column(orders, _cfg_list("CUSTOMER_CANDIDATES"))
part_col = get_first_matching_column(orders, _cfg_list("PART_CANDIDATES"))
qty_col = get_first_matching_column(orders, _cfg_list("QTY_CANDIDATES"))
factory_col = get_first_matching_column(orders, _cfg_list("FACTORY_CANDIDATES"))
wip_col = get_first_matching_column(orders, _cfg_list("WIP_CANDIDATES"))
factory_due_col = get_first_matching_column(orders, _cfg_list("FACTORY_DUE_CANDIDATES"))
ship_date_col = get_first_matching_column(orders, _cfg_list("SHIP_DATE_CANDIDATES"))
remark_col = get_first_matching_column(orders, _cfg_list("REMARK_CANDIDATES"))
customer_tag_col = get_first_matching_column(orders, _cfg_list("CUSTOMER_TAG_CANDIDATES"))

merge_date_col = get_first_matching_column(
    orders,
    getattr(cfg, "MERGE_DATE_CANDIDATES", ["Merge Date", "合併日期", "併單日期"]),
)
order_date_col = get_first_matching_column(
    orders,
    getattr(cfg, "ORDER_DATE_CANDIDATES", ["Order Date", "PO DATE", "下單日期"]),
)
factory_order_date_col = get_first_matching_column(
    orders,
    getattr(cfg, "FACTORY_ORDER_DATE_CANDIDATES", ["Factory Order Date", "工廠下單日期"]),
)
changed_due_date_col = get_first_matching_column(
    orders,
    getattr(cfg, "CHANGED_DUE_DATE_CANDIDATES", ["Changed Due Date", "更改交期", "新交期"]),
)


# ==================================
# SIDEBAR
# ==================================
st.sidebar.title("GLOCOM Internal")
if hasattr(cfg, "TEABLE_WEB_URL") and cfg.TEABLE_WEB_URL:
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
# DEBUG
# ==================================
with st.expander("🔧 Debug / System Info", expanded=False):
    st.write("api_status:", api_status)
    st.write("orders rows:", 0 if orders is None else len(orders))
    st.write("orders type:", type(orders).__name__)
    st.write("orders columns:", list(orders.columns) if isinstance(orders, pd.DataFrame) and not orders.empty else [])
    st.write("Detected PO col:", po_col)
    st.write("Detected WIP col:", wip_col)
    st.write("SALES_BASE_PATH:", SALES_BASE_PATH)
    st.write("Sales workbook exists:", os.path.exists(SALES_BASE_PATH) if SALES_BASE_PATH else False)
    if sales_error_text:
        st.error(f"銷貨底 Excel 載入失敗: {sales_error_text}")
    if isinstance(api_text, str) and api_text.strip():
        st.text(api_text[:2000])


# ==================================
# COMMON KWARGS
# ==================================
common_kwargs = dict(
    orders=orders,
    df=orders,
    data=orders,
    dataset=orders,
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
    split_tags=split_tags,
    safe_text=safe_text,
    wip_display_html=wip_display_html,
)


# ==================================
# ROUTER
# ==================================
if menu == "Dashboard":
    if orders is None or orders.empty:
        show_metrics(orders, wip_col)
        show_no_data_layout(api_status, api_text)
    else:
        ok = call_report_function(["show_dashboard_report"], **common_kwargs)
        if ok is False:
            show_metrics(orders, wip_col)

elif menu == "Factory Load":
    if orders is None or orders.empty:
        st.subheader("🏭 Factory Load")
        st.info("目前沒有資料。")
        show_no_data_layout(api_status, api_text)
    else:
        ok = call_report_function(["show_factory_load_report"], **common_kwargs)
        if ok is False:
            st.subheader("🏭 Factory Load")
            st.info("Factory Load fallback not enabled.")

elif menu == "Delayed Orders":
    if orders is None or orders.empty:
        st.subheader("⏰ Delayed Orders")
        st.info("目前沒有資料。")
        show_no_data_layout(api_status, api_text)
    else:
        ok = call_report_function(["show_delayed_orders_report"], **common_kwargs)
        if ok is False:
            st.info("Delayed Orders fallback not enabled.")

elif menu == "Shipment Forecast":
    if orders is None or orders.empty:
        st.subheader("🚚 Shipment Forecast")
        st.info("目前沒有資料。")
        show_no_data_layout(api_status, api_text)
    else:
        ok = call_report_function(["show_shipment_forecast_report"], **common_kwargs)
        if ok is False:
            st.info("Shipment Forecast fallback not enabled.")

elif menu == "Orders":
    show_orders_page(orders)
    if orders is None or orders.empty:
        show_no_data_layout(api_status, api_text)

elif menu == "Customer Preview":
    show_customer_preview_page(orders)
    if orders is None or orders.empty:
        show_no_data_layout(api_status, api_text)

elif menu == "Sandy 內部 WIP":
    if orders is None or orders.empty:
        st.subheader("🧾 Sandy 內部 WIP")
        st.info("目前沒有資料。")
        show_no_data_layout(api_status, api_text)
    else:
        ok = call_report_function(["show_sandy_internal_wip_report"], **common_kwargs)
        if ok is False:
            st.info("Sandy 內部 WIP fallback not enabled.")

elif menu == "Sandy 銷貨底":
    ok = call_report_function(["show_sandy_shipment_report"], **common_kwargs)
    if ok is False:
        st.info("Sandy 銷貨底 fallback not enabled.")

elif menu == "新訂單 WIP":
    if orders is None or orders.empty:
        st.subheader("🆕 新訂單 WIP")
        st.info("目前沒有資料。")
        show_no_data_layout(api_status, api_text)
    else:
        ok = call_report_function(["show_new_orders_wip_report"], **common_kwargs)
        if ok is False:
            st.info("新訂單 WIP fallback not enabled.")

elif menu == "業績明細表":
    try:
        render_sales_report_page(**common_kwargs)
    except Exception as e:
        st.error(f"業績明細表載入失敗：{e}")

elif menu == "📤 Import / Update":
    fallback_import_update(orders)

st.caption("🔄 Auto refresh cache: 60 seconds | 📊 Last updated: " + pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"))
