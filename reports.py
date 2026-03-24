# -*- coding: utf-8 -*-
"""
reports.py
GLOCOM Control Tower 各頁面報表

相容 app.py 傳入的 common_kwargs。
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional

import pandas as pd
import streamlit as st


# =========================================================
# 基本工具
# =========================================================
def _safe_text(v: Any) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def _col(df: pd.DataFrame, col_name: Optional[str], default: str = "") -> pd.Series:
    if df is None or df.empty or not col_name or col_name not in df.columns:
        return pd.Series([default] * (0 if df is None else len(df)))
    return df[col_name]


def _copy_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame()
    return df.copy()


def _normalize_date_series(series: pd.Series) -> pd.Series:
    if series is None or len(series) == 0:
        return pd.Series(dtype="datetime64[ns]")
    return pd.to_datetime(series, errors="coerce")


def _today() -> pd.Timestamp:
    return pd.Timestamp.now().normalize()


def _contains_any(text: str, keywords: List[str]) -> bool:
    low = _safe_text(text).lower()
    return any(k.lower() in low for k in keywords)


def _is_done_wip(value: Any) -> bool:
    txt = _safe_text(value).upper()
    if not txt:
        return False
    done_values = {"完成", "DONE", "COMPLETE", "COMPLETED", "FINISHED", "FINISH", "CLOSED", "結案"}
    if txt in done_values:
        return True
    return "完成" in _safe_text(value)


def _is_shipping_wip(value: Any) -> bool:
    txt = _safe_text(value).lower()
    return any(k in txt for k in ["ship", "shipping", "出貨"])


def _is_packing_wip(value: Any) -> bool:
    txt = _safe_text(value).lower()
    return any(k in txt for k in ["pack", "packing", "包裝"])


def _is_inspection_wip(value: Any) -> bool:
    txt = _safe_text(value).lower()
    return any(k in txt for k in ["inspection", "inspect", "qa", "fqc", "iqc", "oqc", "檢", "測試", "成檢"])


def _is_hold_row(wip: Any, remark: Any, tags: Any) -> bool:
    combined = " | ".join([
        _safe_text(wip),
        _safe_text(remark),
        ", ".join(tags) if isinstance(tags, list) else _safe_text(tags),
    ]).lower()
    return any(k in combined for k in ["hold", "on hold", "暫停", "待料", "waiting", "pending"])


def _split_tags(value: Any, split_tags=None) -> List[str]:
    if callable(split_tags):
        try:
            return split_tags(value)
        except Exception:
            pass

    if value is None:
        return []
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    text = _safe_text(value)
    if not text:
        return []
    text = text.replace("；", ";").replace("，", ",").replace("、", ",").replace("/", ",").replace("|", ",")
    out = []
    seen = set()
    for p in [x.strip() for x in text.replace(";", ",").split(",") if x.strip()]:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


def _prepare_base_df(
    orders: pd.DataFrame,
    po_col: Optional[str],
    customer_col: Optional[str],
    part_col: Optional[str],
    qty_col: Optional[str],
    factory_col: Optional[str],
    wip_col: Optional[str],
    factory_due_col: Optional[str],
    ship_date_col: Optional[str],
    remark_col: Optional[str],
    customer_tag_col: Optional[str],
    split_tags=None,
) -> pd.DataFrame:
    df = _copy_df(orders)
    if df.empty:
        return pd.DataFrame(columns=[
            "PO#", "Customer", "Part No", "Qty", "Factory", "WIP",
            "Factory Due Date", "Ship Date", "Remark", "Customer Remark Tags"
        ])

    out = pd.DataFrame()
    out["PO#"] = _col(df, po_col, "")
    out["Customer"] = _col(df, customer_col, "")
    out["Part No"] = _col(df, part_col, "")
    out["Qty"] = _col(df, qty_col, "")
    out["Factory"] = _col(df, factory_col, "")
    out["WIP"] = _col(df, wip_col, "")
    out["Factory Due Date"] = _col(df, factory_due_col, "")
    out["Ship Date"] = _col(df, ship_date_col, "")
    out["Remark"] = _col(df, remark_col, "")
    out["Customer Remark Tags Raw"] = _col(df, customer_tag_col, "")

    out["Customer Remark Tags"] = out["Customer Remark Tags Raw"].apply(lambda x: _split_tags(x, split_tags))
    out["Factory Due Date_dt"] = _normalize_date_series(out["Factory Due Date"])
    out["Ship Date_dt"] = _normalize_date_series(out["Ship Date"])

    out["Qty_num"] = pd.to_numeric(out["Qty"], errors="coerce")
    out["Is Done"] = out["WIP"].apply(_is_done_wip)
    out["Is Shipping"] = out["WIP"].apply(_is_shipping_wip)
    out["Is Packing"] = out["WIP"].apply(_is_packing_wip)
    out["Is Inspection"] = out["WIP"].apply(_is_inspection_wip)
    out["Is Hold"] = out.apply(lambda r: _is_hold_row(r.get("WIP", ""), r.get("Remark", ""), r.get("Customer Remark Tags", [])), axis=1)

    return out


def _render_basic_table(df: pd.DataFrame, title: str, height: int = 520) -> None:
    st.subheader(title)
    if df is None or df.empty:
        st.info("目前沒有資料。")
        return
    st.dataframe(df, use_container_width=True, height=height)


def _format_preview_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()
    if "Customer Remark Tags" in out.columns:
        out["Customer Remark Tags"] = out["Customer Remark Tags"].apply(
            lambda x: ", ".join(x) if isinstance(x, list) else _safe_text(x)
        )

    # 移除內部輔助欄
    drop_cols = [c for c in ["Factory Due Date_dt", "Ship Date_dt", "Qty_num", "Is Done", "Is Shipping", "IsPacking", "Is Inspection", "Is Hold", "Customer Remark Tags Raw"] if c in out.columns]
    out = out.drop(columns=drop_cols, errors="ignore")
    return out


def _make_customer_preview_df(df: pd.DataFrame) -> pd.DataFrame:
    keep = [c for c in ["PO#", "Part No", "Qty", "WIP", "Ship Date", "Customer Remark Tags", "Remark"] if c in df.columns]
    out = df[keep].copy()
    if "Customer Remark Tags" in out.columns:
        out["Customer Remark Tags"] = out["Customer Remark Tags"].apply(lambda x: ", ".join(x) if isinstance(x, list) else _safe_text(x))
    return out


# =========================================================
# Dashboard
# =========================================================
def show_dashboard_report(**kwargs) -> None:
    orders = kwargs.get("orders")
    split_tags = kwargs.get("split_tags")
    base = _prepare_base_df(
        orders=orders,
        po_col=kwargs.get("po_col"),
        customer_col=kwargs.get("customer_col"),
        part_col=kwargs.get("part_col"),
        qty_col=kwargs.get("qty_col"),
        factory_col=kwargs.get("factory_col"),
        wip_col=kwargs.get("wip_col"),
        factory_due_col=kwargs.get("factory_due_col"),
        ship_date_col=kwargs.get("ship_date_col"),
        remark_col=kwargs.get("remark_col"),
        customer_tag_col=kwargs.get("customer_tag_col"),
        split_tags=split_tags,
    )

    st.subheader("📊 Dashboard")

    total = len(base)
    active = int((~base["Is Done"]).sum()) if not base.empty else 0
    done = int(base["Is Done"].sum()) if not base.empty else 0
    hold = int(base["Is Hold"].sum()) if not base.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("總筆數", total)
    c2.metric("進行中", active)
    c3.metric("完成", done)
    c4.metric("On Hold", hold)

    if base.empty:
        st.info("目前沒有資料。")
        return

    colA, colB = st.columns(2)

    with colA:
        st.markdown("### WIP 分布")
        vc = base["WIP"].fillna("").astype(str).str.strip()
        vc = vc[vc != ""].value_counts().head(15)
        if not vc.empty:
            st.bar_chart(vc)
        else:
            st.info("沒有 WIP 資料。")

    with colB:
        st.markdown("### 工廠分布")
        vc = base["Factory"].fillna("").astype(str).str.strip()
        vc = vc[vc != ""].value_counts().head(15)
        if not vc.empty:
            st.bar_chart(vc)
        else:
            st.info("沒有工廠資料。")

    st.markdown("### 即將到期 / 已逾期")
    today = _today()
    tmp = base.copy()
    tmp["Due Status"] = ""
    tmp.loc[tmp["Factory Due Date_dt"] < today, "Due Status"] = "Overdue"
    tmp.loc[(tmp["Factory Due Date_dt"] >= today) & (tmp["Factory Due Date_dt"] <= today + pd.Timedelta(days=7)), "Due Status"] = "Due in 7 Days"

    due_view = tmp[(tmp["Due Status"] != "") & (~tmp["Is Done"])].copy()
    due_view = due_view.sort_values(by=["Factory Due Date_dt", "Customer", "PO#"], na_position="last")
    _render_basic_table(_format_preview_df(due_view).head(100), "到期提醒", height=360)


# =========================================================
# Factory Load
# =========================================================
def show_factory_load_report(**kwargs) -> None:
    orders = kwargs.get("orders")
    split_tags = kwargs.get("split_tags")
    base = _prepare_base_df(
        orders=orders,
        po_col=kwargs.get("po_col"),
        customer_col=kwargs.get("customer_col"),
        part_col=kwargs.get("part_col"),
        qty_col=kwargs.get("qty_col"),
        factory_col=kwargs.get("factory_col"),
        wip_col=kwargs.get("wip_col"),
        factory_due_col=kwargs.get("factory_due_col"),
        ship_date_col=kwargs.get("ship_date_col"),
        remark_col=kwargs.get("remark_col"),
        customer_tag_col=kwargs.get("customer_tag_col"),
        split_tags=split_tags,
    )

    st.subheader("🏭 Factory Load")

    if base.empty:
        st.info("目前沒有資料。")
        return

    active = base[~base["Is Done"]].copy()

    if active.empty:
        st.info("目前沒有進行中的訂單。")
        return

    summary = (
        active.groupby("Factory", dropna=False)
        .agg(
            Orders=("PO#", "count"),
            Qty_Total=("Qty_num", "sum"),
            Hold_Count=("Is Hold", "sum"),
            Inspection_Count=("Is Inspection", "sum"),
            Packing_Count=("Is Packing", "sum"),
            Shipping_Count=("Is Shipping", "sum"),
        )
        .reset_index()
        .sort_values(["Orders", "Qty_Total"], ascending=[False, False], na_position="last")
    )

    st.dataframe(summary, use_container_width=True, height=360)

    st.markdown("### 依工廠檢視明細")
    factories = sorted([x for x in active["Factory"].dropna().astype(str).unique() if x.strip()])
    selected_factory = st.selectbox("選擇工廠", ["全部"] + factories)

    view = active.copy()
    if selected_factory != "全部":
        view = view[view["Factory"].astype(str) == selected_factory]

    view = view.sort_values(by=["Factory", "Factory Due Date_dt", "Customer", "PO#"], na_position="last")
    _render_basic_table(_format_preview_df(view), "工廠明細", height=420)


# =========================================================
# Delayed Orders
# =========================================================
def show_delayed_orders_report(**kwargs) -> None:
    orders = kwargs.get("orders")
    split_tags = kwargs.get("split_tags")
    base = _prepare_base_df(
        orders=orders,
        po_col=kwargs.get("po_col"),
        customer_col=kwargs.get("customer_col"),
        part_col=kwargs.get("part_col"),
        qty_col=kwargs.get("qty_col"),
        factory_col=kwargs.get("factory_col"),
        wip_col=kwargs.get("wip_col"),
        factory_due_col=kwargs.get("factory_due_col"),
        ship_date_col=kwargs.get("ship_date_col"),
        remark_col=kwargs.get("remark_col"),
        customer_tag_col=kwargs.get("customer_tag_col"),
        split_tags=split_tags,
    )

    st.subheader("⏰ Delayed Orders")

    if base.empty:
        st.info("目前沒有資料。")
        return

    today = _today()
    delayed = base[(~base["Is Done"]) & (base["Factory Due Date_dt"].notna()) & (base["Factory Due Date_dt"] < today)].copy()
    delayed["Delay Days"] = (today - delayed["Factory Due Date_dt"]).dt.days

    c1, c2 = st.columns(2)
    c1.metric("逾期筆數", len(delayed))
    c2.metric("平均延遲天數", round(delayed["Delay Days"].mean(), 1) if not delayed.empty else 0)

    if delayed.empty:
        st.success("目前沒有逾期訂單。")
        return

    view = delayed.sort_values(by=["Delay Days", "Factory Due Date_dt"], ascending=[False, True], na_position="last")
    _render_basic_table(_format_preview_df(view), "逾期明細", height=520)


# =========================================================
# Shipment Forecast
# =========================================================
def show_shipment_forecast_report(**kwargs) -> None:
    orders = kwargs.get("orders")
    split_tags = kwargs.get("split_tags")
    base = _prepare_base_df(
        orders=orders,
        po_col=kwargs.get("po_col"),
        customer_col=kwargs.get("customer_col"),
        part_col=kwargs.get("part_col"),
        qty_col=kwargs.get("qty_col"),
        factory_col=kwargs.get("factory_col"),
        wip_col=kwargs.get("wip_col"),
        factory_due_col=kwargs.get("factory_due_col"),
        ship_date_col=kwargs.get("ship_date_col"),
        remark_col=kwargs.get("remark_col"),
        customer_tag_col=kwargs.get("customer_tag_col"),
        split_tags=split_tags,
    )

    st.subheader("🚚 Shipment Forecast")

    if base.empty:
        st.info("目前沒有資料。")
        return

    today = _today()
    future = base[(~base["Is Done"]) & (base["Ship Date_dt"].notna())].copy()
    future = future[future["Ship Date_dt"] >= today - pd.Timedelta(days=3)]

    if future.empty:
        st.info("目前沒有可預測的出貨資料。")
        return

    future["Ship Week"] = future["Ship Date_dt"].dt.strftime("%Y-%W")
    weekly = (
        future.groupby("Ship Week")
        .agg(Orders=("PO#", "count"), Qty_Total=("Qty_num", "sum"))
        .reset_index()
        .sort_values("Ship Week")
    )

    st.markdown("### 週別預測")
    st.dataframe(weekly, use_container_width=True, height=260)

    st.markdown("### 近期出貨明細")
    detail = future.sort_values(by=["Ship Date_dt", "Customer", "PO#"], na_position="last")
    _render_basic_table(_format_preview_df(detail), "Shipment Forecast Detail", height=480)


# =========================================================
# Orders
# =========================================================
def show_orders_report(**kwargs) -> None:
    orders = kwargs.get("orders")
    split_tags = kwargs.get("split_tags")
    base = _prepare_base_df(
        orders=orders,
        po_col=kwargs.get("po_col"),
        customer_col=kwargs.get("customer_col"),
        part_col=kwargs.get("part_col"),
        qty_col=kwargs.get("qty_col"),
        factory_col=kwargs.get("factory_col"),
        wip_col=kwargs.get("wip_col"),
        factory_due_col=kwargs.get("factory_due_col"),
        ship_date_col=kwargs.get("ship_date_col"),
        remark_col=kwargs.get("remark_col"),
        customer_tag_col=kwargs.get("customer_tag_col"),
        split_tags=split_tags,
    )

    st.subheader("📋 Orders")

    if base.empty:
        st.info("目前沒有資料。")
        return

    keyword = st.text_input("搜尋 PO / 客戶 / Part No / 工廠", "")
    status_filter = st.selectbox("狀態", ["全部", "進行中", "完成", "On Hold", "Inspection", "Packing", "Shipping"])

    view = base.copy()

    if keyword.strip():
        kw = keyword.strip().lower()
        mask = (
            view["PO#"].astype(str).str.lower().str.contains(kw, na=False)
            | view["Customer"].astype(str).str.lower().str.contains(kw, na=False)
            | view["Part No"].astype(str).str.lower().str.contains(kw, na=False)
            | view["Factory"].astype(str).str.lower().str.contains(kw, na=False)
        )
        view = view[mask]

    if status_filter == "進行中":
        view = view[~view["Is Done"]]
    elif status_filter == "完成":
        view = view[view["Is Done"]]
    elif status_filter == "On Hold":
        view = view[view["Is Hold"]]
    elif status_filter == "Inspection":
        view = view[view["Is Inspection"]]
    elif status_filter == "Packing":
        view = view[view["Is Packing"]]
    elif status_filter == "Shipping":
        view = view[view["Is Shipping"]]

    view = view.sort_values(by=["Factory Due Date_dt", "Ship Date_dt", "Customer", "PO#"], na_position="last")
    _render_basic_table(_format_preview_df(view), "Orders Detail", height=560)


# =========================================================
# Customer Preview
# =========================================================
def show_customer_preview_report(**kwargs) -> None:
    orders = kwargs.get("orders")
    split_tags = kwargs.get("split_tags")
    base = _prepare_base_df(
        orders=orders,
        po_col=kwargs.get("po_col"),
        customer_col=kwargs.get("customer_col"),
        part_col=kwargs.get("part_col"),
        qty_col=kwargs.get("qty_col"),
        factory_col=kwargs.get("factory_col"),
        wip_col=kwargs.get("wip_col"),
        factory_due_col=kwargs.get("factory_due_col"),
        ship_date_col=kwargs.get("ship_date_col"),
        remark_col=kwargs.get("remark_col"),
        customer_tag_col=kwargs.get("customer_tag_col"),
        split_tags=split_tags,
    )

    st.subheader("👀 Customer Preview")

    if base.empty:
        st.info("目前沒有資料。")
        return

    customer_param = st.query_params.get("customer", "")
    all_customers = sorted([x for x in base["Customer"].dropna().astype(str).unique() if x.strip()])

    default_index = 0
    options = ["全部"] + all_customers
    if customer_param and customer_param in all_customers:
        default_index = options.index(customer_param)

    selected_customer = st.selectbox("選擇客戶", options, index=default_index)

    view = base.copy()
    if selected_customer != "全部":
        view = view[view["Customer"].astype(str) == selected_customer]

    preview = _make_customer_preview_df(view)
    preview = preview.sort_values(by=["Ship Date", "PO#", "Part No"], na_position="last")
    st.dataframe(preview, use_container_width=True, height=520)

    csv = preview.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ 下載目前客戶預覽 CSV",
        data=csv,
        file_name="customer_preview.csv",
        mime="text/csv",
    )


# =========================================================
# Sandy 內部 WIP
# =========================================================
def show_sandy_internal_wip_report(**kwargs) -> None:
    orders = kwargs.get("orders")
    split_tags = kwargs.get("split_tags")
    base = _prepare_base_df(
        orders=orders,
        po_col=kwargs.get("po_col"),
        customer_col=kwargs.get("customer_col"),
        part_col=kwargs.get("part_col"),
        qty_col=kwargs.get("qty_col"),
        factory_col=kwargs.get("factory_col"),
        wip_col=kwargs.get("wip_col"),
        factory_due_col=kwargs.get("factory_due_col"),
        ship_date_col=kwargs.get("ship_date_col"),
        remark_col=kwargs.get("remark_col"),
        customer_tag_col=kwargs.get("customer_tag_col"),
        split_tags=split_tags,
    )

    st.subheader("🧾 Sandy 內部 WIP")

    if base.empty:
        st.info("目前沒有資料。")
        return

    active = base[~base["Is Done"]].copy()
    active = active.sort_values(by=["Factory Due Date_dt", "Factory", "Customer", "PO#"], na_position="last")
    _render_basic_table(_format_preview_df(active), "內部 WIP", height=560)


# =========================================================
# Sandy 銷貨底
# =========================================================
def show_sandy_shipment_report(**kwargs) -> None:
    sales_shipment_df = kwargs.get("sales_shipment_df")
    st.subheader("📦 Sandy 銷貨底")

    df = _copy_df(sales_shipment_df)
    if df.empty:
        st.info("目前沒有銷貨底資料。")
        return

    st.dataframe(df, use_container_width=True, height=560)


# =========================================================
# 新訂單 WIP
# =========================================================
def show_new_orders_wip_report(**kwargs) -> None:
    orders = kwargs.get("orders")
    order_date_col = kwargs.get("order_date_col")
    split_tags = kwargs.get("split_tags")

    base = _prepare_base_df(
        orders=orders,
        po_col=kwargs.get("po_col"),
        customer_col=kwargs.get("customer_col"),
        part_col=kwargs.get("part_col"),
        qty_col=kwargs.get("qty_col"),
        factory_col=kwargs.get("factory_col"),
        wip_col=kwargs.get("wip_col"),
        factory_due_col=kwargs.get("factory_due_col"),
        ship_date_col=kwargs.get("ship_date_col"),
        remark_col=kwargs.get("remark_col"),
        customer_tag_col=kwargs.get("customer_tag_col"),
        split_tags=split_tags,
    )

    st.subheader("🆕 新訂單 WIP")

    if base.empty:
        st.info("目前沒有資料。")
        return

    raw = _copy_df(orders)
    if order_date_col and order_date_col in raw.columns:
        base["Order Date"] = pd.to_datetime(raw[order_date_col], errors="coerce")
    else:
        base["Order Date"] = pd.NaT

    days = st.slider("近幾天新訂單", min_value=3, max_value=60, value=14, step=1)
    cutoff = _today() - pd.Timedelta(days=days)

    view = base[(base["Order Date"].notna()) & (base["Order Date"] >= cutoff)].copy()
    view = view.sort_values(by=["Order Date", "Customer", "PO#"], ascending=[False, True, True], na_position="last")

    c1, c2 = st.columns(2)
    c1.metric("新訂單筆數", len(view))
    c2.metric("進行中筆數", int((~view["Is Done"]).sum()) if not view.empty else 0)

    _render_basic_table(_format_preview_df(view), "新訂單明細", height=520)


# =========================================================
# Import / Update 頁
# =========================================================
def show_import_update_page(**kwargs) -> None:
    st.subheader("📤 Import / Update")
    st.info("此頁已由 app.py 的 fallback_import_update() 接手處理。")


def show_import_update_report(**kwargs) -> None:
    st.subheader("📤 Import / Update")
    st.info("此頁已由 app.py 的 fallback_import_update() 接手處理。")


# =========================================================
# 可選：銷貨資料 loader（讓 app.py 的 cached_load_sales_data() 可呼叫）
# =========================================================
def load_sales_data():
    """
    保留相容介面。
    若你未來要接 Sandy 銷貨底 Excel，可在這裡補實作。
    目前先回傳空 DataFrame，避免 app.py 報錯。
    """
    return pd.DataFrame(), pd.DataFrame(), ""
