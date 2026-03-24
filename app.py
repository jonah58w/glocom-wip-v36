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
    .small-muted { color: #6b7280; font-size: 0.9em; }
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


def normalize_date_series(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def fmt_date(v: Any) -> str:
    try:
        dt = pd.to_datetime(v, errors="coerce")
        if pd.isna(dt):
            return ""
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return safe_text(v)


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


def show_no_data_layout() -> None:
    st.markdown(
        """
        <div class="portal-box">
            <div class="portal-title">GLOCOM Control Tower</div>
            <div>目前未讀取到 Teable 資料。</div>
            <div style="margin-top:8px;">請確認：</div>
            <ul>
                <li>Streamlit secrets 已設定 TEABLE_TOKEN / TEABLE_TABLE_URL</li>
                <li>TABLE_URL 是否正確（/table/{tableId}/view/{viewId}）</li>
                <li>Teable API 是否可正常連線</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )


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
    elif "profit" in name_lower or "pg" in name_lower or "glocom-pg" in name_lower:
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


def _safe_dataframe(df: pd.DataFrame, cols: List[Optional[str]]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    final_cols = []
    for c in cols:
        if c and c in df.columns and c not in final_cols:
            final_cols.append(c)
    if not final_cols:
        return df.copy()
    return df[final_cols].copy()


def _is_done_value(x: Any) -> bool:
    text = safe_text(x)
    if not text:
        return False
    done_values = {str(v).upper() for v in getattr(cfg, "DONE_WIP_VALUES", set())}
    return text.upper() in done_values or ("完成" in text)


def _is_hold_value(x: Any) -> bool:
    text = safe_text(x).lower()
    return any(k in text for k in ["hold", "暫停", "等待"])


def _get_active_orders(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or not wip_col or wip_col not in df.columns:
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
    mask = ~df[wip_col].apply(_is_done_value)
    return df[mask].copy()


def _get_delayed_orders(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    date_col = factory_due_col or ship_date_col
    if not date_col or date_col not in df.columns:
        return pd.DataFrame()

    temp = df.copy()
    temp["_due_dt"] = pd.to_datetime(temp[date_col], errors="coerce")
    temp = temp[temp["_due_dt"].notna()].copy()

    if wip_col and wip_col in temp.columns:
        temp = temp[~temp[wip_col].apply(_is_done_value)].copy()

    today = pd.Timestamp.today().normalize()
    temp = temp[temp["_due_dt"] < today].copy()
    temp = temp.sort_values("_due_dt")
    return temp


def _get_new_orders(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    date_col = order_date_col or factory_order_date_col or merge_date_col
    if not date_col or date_col not in df.columns:
        return pd.DataFrame()

    temp = df.copy()
    temp["_new_dt"] = pd.to_datetime(temp[date_col], errors="coerce")
    temp = temp[temp["_new_dt"].notna()].copy()

    cutoff = pd.Timestamp.today().normalize() - pd.Timedelta(days=14)
    temp = temp[temp["_new_dt"] >= cutoff].copy()
    temp = temp.sort_values("_new_dt", ascending=False)
    return temp


def _get_shipment_forecast(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    date_col = ship_date_col or factory_due_col
    if not date_col or date_col not in df.columns:
        return pd.DataFrame()

    temp = df.copy()
    temp["_ship_dt"] = pd.to_datetime(temp[date_col], errors="coerce")
    temp = temp[temp["_ship_dt"].notna()].copy()

    today = pd.Timestamp.today().normalize()
    end_date = today + pd.Timedelta(days=30)

    if wip_col and wip_col in temp.columns:
        temp = temp[~temp[wip_col].apply(_is_done_value)].copy()

    temp = temp[(temp["_ship_dt"] >= today) & (temp["_ship_dt"] <= end_date)].copy()
    temp = temp.sort_values("_ship_dt")
    return temp


# ==================================
# PAGE FUNCTIONS
# ==================================
def show_dashboard_page(orders: pd.DataFrame) -> None:
    st.subheader("📊 Dashboard")

    if orders is None or orders.empty:
        st.info("目前沒有資料。")
        return

    total = len(orders)
    done_count = 0
    hold_count = 0

    if wip_col and wip_col in orders.columns:
        done_count = int(orders[wip_col].apply(_is_done_value).sum())
        hold_count = int(orders[wip_col].apply(_is_hold_value).sum())

    active = max(total - done_count, 0)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("總筆數", total)
    c2.metric("進行中", active)
    c3.metric("完成", done_count)
    c4.metric("On Hold", hold_count)

    if wip_col and wip_col in orders.columns:
        st.markdown("### WIP 分布")
        vc = orders[wip_col].fillna("").astype(str).str.strip()
        vc = vc[vc != ""].value_counts().head(15)
        if not vc.empty:
            st.bar_chart(vc)

    date_col = ship_date_col or factory_due_col
    if date_col and date_col in orders.columns:
        temp = orders.copy()
        temp["_ship_dt"] = pd.to_datetime(temp[date_col], errors="coerce")
        temp = temp[temp["_ship_dt"].notna()].copy()
        if not temp.empty:
            st.markdown("### 出貨 / 交期趨勢")
            trend = temp.groupby(temp["_ship_dt"].dt.date).size()
            st.line_chart(trend)


def show_factory_load_page(orders: pd.DataFrame) -> None:
    st.subheader("🏭 Factory Load")

    if orders is None or orders.empty:
        st.info("目前沒有資料。")
        return

    df = _get_active_orders(orders)

    if factory_col and factory_col in df.columns:
        st.markdown("### 工廠在製筆數")
        vc = df[factory_col].fillna("未指定").astype(str).value_counts()
        st.bar_chart(vc)

        view = vc.reset_index()
        view.columns = ["Factory", "Active Orders"]
        st.dataframe(view, use_container_width=True, height=320)
    else:
        st.info("找不到工廠欄位。")
        st.dataframe(_safe_dataframe(df, [po_col, part_col, qty_col, wip_col, ship_date_col]), use_container_width=True)


def show_delayed_orders_page(orders: pd.DataFrame) -> None:
    st.subheader("⏰ Delayed Orders")

    delayed = _get_delayed_orders(orders)
    if delayed.empty:
        st.info("目前沒有 Delay 案件。")
        return

    date_col = factory_due_col or ship_date_col
    delayed["_Delay Days"] = (pd.Timestamp.today().normalize() - delayed["_due_dt"]).dt.days

    cols = [customer_col, po_col, part_col, qty_col, factory_col, wip_col, date_col, remark_col]
    view = _safe_dataframe(delayed, cols)
    if "_Delay Days" in delayed.columns:
        view["Delay Days"] = delayed["_Delay Days"].values

    st.metric("Delay 筆數", len(delayed))
    st.dataframe(view, use_container_width=True, height=560)


def show_shipment_forecast_page(orders: pd.DataFrame) -> None:
    st.subheader("🚚 Shipment Forecast")

    forecast = _get_shipment_forecast(orders)
    if forecast.empty:
        st.info("未來 30 天目前沒有可顯示的出貨預估。")
        return

    date_col = ship_date_col or factory_due_col
    forecast["_Ship Date"] = forecast["_ship_dt"].dt.date

    summary = forecast.groupby("_Ship Date").size().reset_index(name="Count")
    st.markdown("### 未來 30 天出貨分布")
    st.bar_chart(summary.set_index("_Ship Date")["Count"])

    cols = [customer_col, po_col, part_col, qty_col, factory_col, wip_col, date_col, remark_col]
    st.dataframe(_safe_dataframe(forecast, cols), use_container_width=True, height=560)


def show_orders_page(orders: pd.DataFrame) -> None:
    st.subheader("📋 Orders")

    if orders is None or orders.empty:
        st.info("目前沒有資料。")
        return

    df = orders.copy()
    filter_cols = st.columns(4)

    if customer_col and customer_col in df.columns:
        customer_options = ["全部"] + sorted(
            [x for x in df[customer_col].dropna().astype(str).unique() if x.strip()]
        )
        sel_customer = filter_cols[0].selectbox("Customer", customer_options, index=0)
        if sel_customer != "全部":
            df = df[df[customer_col].astype(str) == sel_customer]

    if factory_col and factory_col in df.columns:
        factory_options = ["全部"] + sorted(
            [x for x in df[factory_col].dropna().astype(str).unique() if x.strip()]
        )
        sel_factory = filter_cols[1].selectbox("Factory", factory_options, index=0)
        if sel_factory != "全部":
            df = df[df[factory_col].astype(str) == sel_factory]

    if wip_col and wip_col in df.columns:
        wip_options = ["全部"] + sorted(
            [x for x in df[wip_col].dropna().astype(str).unique() if x.strip()]
        )
        sel_wip = filter_cols[2].selectbox("WIP", wip_options, index=0)
        if sel_wip != "全部":
            df = df[df[wip_col].astype(str) == sel_wip]

    keyword = filter_cols[3].text_input("搜尋 PO / Part / Remark")

    if keyword:
        mask = pd.Series(False, index=df.index)
        for c in [po_col, part_col, remark_col]:
            if c and c in df.columns:
                mask = mask | df[c].astype(str).str.contains(keyword, case=False, na=False)
        df = df[mask]

    st.caption(f"共 {len(df)} 筆")

    display_cols = []
    for c in [
        customer_col,
        po_col,
        part_col,
        qty_col,
        factory_col,
        wip_col,
        factory_due_col,
        ship_date_col,
        remark_col,
    ]:
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
    customer_from_query = st.query_params.get("customer", "")
    if isinstance(customer_from_query, list) and customer_from_query:
        customer_from_query = customer_from_query[0]

    if customer_col and customer_col in df.columns:
        customer_options = sorted(
            [x for x in df[customer_col].dropna().astype(str).unique() if x.strip()]
        )
        if customer_from_query and customer_from_query in customer_options:
            default_ix = customer_options.index(customer_from_query)
        else:
            default_ix = 0 if customer_options else None

        if customer_options:
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


def show_sandy_internal_wip_page(orders: pd.DataFrame) -> None:
    st.subheader("🧾 Sandy 內部 WIP")

    if orders is None or orders.empty:
        st.info("目前沒有資料。")
        return

    df = _get_active_orders(orders)

    if customer_col and customer_col in df.columns:
        keywords = ["SANDY", "Sandy", "sandy"]
        mask = pd.Series(False, index=df.index)
        for kw in keywords:
            mask = mask | df[customer_col].astype(str).str.contains(kw, case=False, na=False)
        sandy_df = df[mask].copy()
        if not sandy_df.empty:
            df = sandy_df

    cols = [customer_col, po_col, part_col, qty_col, factory_col, wip_col, factory_due_col, ship_date_col, remark_col]
    st.caption(f"顯示 {len(df)} 筆")
    st.dataframe(_safe_dataframe(df, cols), use_container_width=True, height=560)


def show_sandy_shipment_page(sales_df: pd.DataFrame, sales_shipment_df: pd.DataFrame) -> None:
    st.subheader("📦 Sandy 銷貨底")

    df = pd.DataFrame()
    if isinstance(sales_shipment_df, pd.DataFrame) and not sales_shipment_df.empty:
        df = sales_shipment_df.copy()
    elif isinstance(sales_df, pd.DataFrame) and not sales_df.empty:
        df = sales_df.copy()

    if df.empty:
        st.info("目前沒有銷貨底資料。")
        return

    st.caption(f"顯示 {len(df)} 筆")
    st.dataframe(df, use_container_width=True, height=560)


def show_new_orders_wip_page(orders: pd.DataFrame) -> None:
    st.subheader("🆕 新訂單 WIP")

    df = _get_new_orders(orders)

    if df.empty:
        st.info("近 14 天目前沒有新訂單資料。")
        return

    date_col = order_date_col or factory_order_date_col or merge_date_col
    cols = [date_col, customer_col, po_col, part_col, qty_col, factory_col, wip_col, ship_date_col, remark_col]

    st.metric("新訂單筆數", len(df))
    st.dataframe(_safe_dataframe(df, cols), use_container_width=True, height=560)


def show_import_update_page(orders: pd.DataFrame) -> None:
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
                        detect_process_fn = getattr(teable_api, "_detect_process_columns", None)
                        if callable(detect_process_fn):
                            try:
                                process_cols = detect_process_fn(df_up)
                            except Exception:
                                process_cols = []

                        st.write(f"**PO 列**: `{po_match}`" if po_match else "**PO 列**: ❌ 未找到")
                        st.write(f"**Part / 料號列**: `{part_match}`" if part_match else "**Part / 料號列**: ❌ 未找到")
                        st.write(f"**WIP 列**: `{wip_match}`" if wip_match else "**WIP 列**: ℹ️ 將嘗試多列製程解析")

                        if process_cols:
                            st.write(f"**製程列檢測**: {len(process_cols)} 個 → {process_cols[:8]}")

                        if not po_match and not part_match:
                            st.warning("這份檔案目前沒有找到 PO 或料號欄，更新時可能失敗。")

                    if st.button("📤 更新到 Teable", type="primary", key="update_teable_btn"):
                        with st.spinner("🔄 正在批量更新到 Teable..."):
                            results = teable_api.batch_update_wip_from_excel(
                                current_df=orders,
                                uploaded_df=df_up,
                                factory_name=factory_name,
                            )
                            _display_update_results(results)

                            if results.get("success_count", 0) > 0:
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
            st.caption("已接收貼上文字，可先複製成 .txt 上傳，或之後再串接 text parser 即時更新。")
            st.text_area("文字預覽", value=pasted, height=220, key="factory_text_preview_2")

    with tab4:
        st.info("手工輸入模式保留。若您要，我下一版可幫您把手工輸入直接接成可更新 Teable。")


# ==================================
# DATA LOADERS
# ==================================
@st.cache_data(ttl=60, show_spinner=False)
def cached_load_orders():
    return teable_api.load_orders()


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

with st.expander("🔧 Debug / System Info", expanded=False):
    st.write("Orders loaded:", 0 if orders is None else len(orders))
    st.write(
        "Orders columns:",
        list(orders.columns) if isinstance(orders, pd.DataFrame) and not orders.empty else [],
    )
    st.write("Detected PO col:", po_col)
    st.write("Detected WIP col:", wip_col)
    st.write("SALES_BASE_PATH:", SALES_BASE_PATH)
    st.write("Sales workbook exists:", os.path.exists(SALES_BASE_PATH) if SALES_BASE_PATH else False)
    if sales_error_text:
        st.error(f"銷貨底 Excel 載入失敗: {sales_error_text}")
    if isinstance(api_text, str):
        st.text(api_text[:800])

if orders is None or orders.empty:
    show_no_data_layout()

# ==================================
# SIDEBAR
# ==================================
st.sidebar.title("GLOCOM Internal")
if hasattr(cfg, "TEABLE_WEB_URL"):
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
    show_dashboard_page(orders)

elif menu == "Factory Load":
    show_factory_load_page(orders)

elif menu == "Delayed Orders":
    show_delayed_orders_page(orders)

elif menu == "Shipment Forecast":
    show_shipment_forecast_page(orders)

elif menu == "Orders":
    show_orders_page(orders)

elif menu == "Customer Preview":
    show_customer_preview_page(orders)

elif menu == "Sandy 內部 WIP":
    show_sandy_internal_wip_page(orders)

elif menu == "Sandy 銷貨底":
    show_sandy_shipment_page(sales_df, sales_shipment_df)

elif menu == "新訂單 WIP":
    show_new_orders_wip_page(orders)

elif menu == "業績明細表":
    try:
        render_sales_report_page(**common_kwargs)
    except Exception as e:
        st.error(f"業績明細表載入失敗：{e}")

elif menu == "📤 Import / Update":
    show_import_update_page(orders)

st.caption(
    "🔄 Auto refresh cache: 60 seconds | 📊 Last updated: "
    + pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")
)
