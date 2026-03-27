# -*- coding: utf-8 -*-
import re
from datetime import datetime

import pandas as pd
import streamlit as st

# ================================
# FIELD CANDIDATES
# ================================
PO_CANDIDATES = [
    "PO#", "PO", "P/O", "訂單編號", "訂單號", "訂單號碼", "工單", "工單號", "單號"
]
CUSTOMER_CANDIDATES = [
    "Customer", "客戶", "客戶名稱"
]
PART_CANDIDATES = [
    "Part No", "Part No.", "P/N", "客戶料號", "Cust. P / N", "LS P/N",
    "料號", "品號", "成品料號", "產品料號"
]
QTY_CANDIDATES = [
    "Qty", "Order Q'TY (PCS)", "Order Q'TY (PCS)", "訂購量 (PCS)",
    "訂購量", "Q'TY", "數量", "PCS", "訂單量", "生產數量", "投產數"
]
FACTORY_CANDIDATES = [
    "Factory", "工廠", "廠編"
]
WIP_CANDIDATES = [
    "WIP", "WIP Stage", "進度", "製程", "工序", "目前站別", "生產進度"
]
FACTORY_DUE_CANDIDATES = [
    "Factory Due Date", "工廠交期", "交貨日期", "Required Ship date",
    "confrimed DD", "交期", "預交日", "預定交期", "交貨期"
]
SHIP_DATE_CANDIDATES = [
    "Ship Date", "Ship date", "出貨日期", "交貨日期", "Required Ship date", "confrimed DD"
]
REMARK_CANDIDATES = [
    "Remark", "備註", "情況", "備註說明", "Note", "說明", "異常備註"
]
ORDER_DATE_CANDIDATES = [
    "客戶下單日期", "工廠下單日期", "下單日期", "Order Date", "PO Date", "Date",
    "訂單日期", "接單日期"
]
AMOUNT_ORDER_CANDIDATES = [
    "接單金額", "接單總金額", "Order Amount", "Order amount", "Order Total",
    "客戶金額", "銷售金額", "Sales Amount", "Quote Total", "Total Amount", "Amount",
    "INVOICE", "Invoice", "Invoice Amount", "Invoice Total"
]
AMOUNT_SHIP_CANDIDATES = [
    "出貨金額", "出貨總金額", "Shipment Amount", "Ship Amount", "Shipping Amount",
    "Invoice Amount", "Invoice Total", "出貨發票金額", "Invoice", "INVOICE"
]

# ================================
# SHARED HELPERS
# ================================
def get_series_by_col(df: pd.DataFrame, col_name: str):
    if not col_name or col_name not in df.columns:
        return None
    obj = df[col_name]
    if isinstance(obj, pd.DataFrame):
        return obj.iloc[:, 0]
    return obj

def col_candidates(*names):
    return [str(x).strip() for x in names if str(x).strip()]
SANDY_NEW_ORDER_SPECS = [
    ("客戶下單日期", col_candidates("客戶下單日期", "客戶下單日期", "客戶下單日期")),
    ("工廠下單日期", col_candidates("工廠下單日期", "工廠下單日期", "工廠下單日期")),
    ("客戶", CUSTOMER_CANDIDATES + ["Customer"]),
    ("PO#", PO_CANDIDATES),
    ("P/N", PART_CANDIDATES),
    ("Order Q'TY (PCS)", QTY_CANDIDATES + ["Order Q'TY(PCS)", "Order QTY (PCS)"]),
    ("Dock", col_candidates("Dock")),
    ("Ship date", SHIP_DATE_CANDIDATES),
    ("WIP", WIP_CANDIDATES),
    ("工廠交期", FACTORY_DUE_CANDIDATES),
    ("交期 (更改)", col_candidates("交期 (更改)", "交期 (更改)", "交期 (更改)", "交期 (更改)")),
    ("出貨日期", col_candidates("出貨日期")),
    ("工廠", FACTORY_CANDIDATES),
    ("工廠提醒事項", col_candidates("工廠提醒事項")),
    ("併貨日期 (限內部使用)", col_candidates("併貨日期 (限內部使用)", "併貨日期 (限內部使用)", "併貨日期 (限內部使用)")),
    ("情況", REMARK_CANDIDATES),
    ("客戶要求注意事項", col_candidates("客戶要求注意事項")),
    ("Ship to", col_candidates("Ship to")),
    ("Ship via", col_candidates("Ship via", " Ship via")),
    ("箱數", col_candidates("箱數", "CTNS", "CTN")),
    ("重量", col_candidates("重量", "Weight", "KGs")),
    ("重貨優惠", col_candidates("重貨優惠", "重貨優惠", "重貨優惠")),
    ("Working Gerber Approval", col_candidates("Working Gerber Approval", "Working Gerber Approval", "Working Gerber Approval")),
    ("Engineering Question", col_candidates("Engineering Question", "Engineering Question", "Engineering Question")),
    ("Pricing & Qty issue", col_candidates("Pricing & Qty issue", "Pricing & Qty issue", "Pricing & Qty issue")),
    ("T/T", col_candidates("T/T")),
    ("工廠出貨事項", col_candidates("工廠出貨事項", "工廠出貨注意事項")),
    ("文件", col_candidates("文件")),
    ("新/舊料號", col_candidates("新/舊料號", "新/舊料號")),
    ("板層", col_candidates("板層", "板層")),
    ("西拓訂單編號", col_candidates("西拓訂單編號", "西拓訂單編號")),
]
SANDY_INTERNAL_WIP_SPECS = [
    ("Customer", CUSTOMER_CANDIDATES + ["Customer"]),
    ("PO#", PO_CANDIDATES),
    ("P/N", PART_CANDIDATES),
    ("Q'TY (PCS)", QTY_CANDIDATES + ["Order QTY (PCS)"]),
    ("Dock", col_candidates("Dock")),
    ("Ship date", SHIP_DATE_CANDIDATES),
    ("WIP", WIP_CANDIDATES),
    ("出貨狀況 (限內部使用)", col_candidates("出貨狀況 (限內部使用)", "出貨狀況 (限內部使用)")),
    ("進度狀況", col_candidates("進度狀況", "進度狀況")),
    ("工廠交期", FACTORY_DUE_CANDIDATES),
    ("交期 (更改)", col_candidates("交期 (更改)", "交期 (更改)", "交期 (更改)", "交期 (更改)")),
    ("出貨日期", col_candidates("出貨日期")),
    ("工廠", FACTORY_CANDIDATES),
    ("工廠提醒事項", col_candidates("工廠提醒事項")),
    ("併貨日期 (限內部使用)", col_candidates("併貨日期 (限內部使用)", "併貨日期 (限內部使用)", "併貨日期 (限內部使用)")),
    ("客戶要求注意事項", col_candidates("客戶要求注意事項", "客戶要求注意事項")),
    ("Ship to", col_candidates("Ship to")),
    ("Ship via", col_candidates("Ship via", " Ship via")),
    ("CTN", col_candidates("CTN", "CTNS", "箱數")),
    ("KGs", col_candidates("KGs", "Weight", "重量")),
    ("重貨優惠", col_candidates("重貨優惠", "重貨優惠", "重貨優惠")),
    ("物流 Booking", col_candidates("物流 Booking", "物流 Booking", "物流 Booking")),
    ("更改 Booking", col_candidates("更改 Booking", "更改 Booking")),
    ("工廠入倉單", col_candidates("工廠入倉單", "工廠入倉單")),
    ("Working Gerber Approval", col_candidates("Working Gerber Approval", "Working Gerber Approval", "Working Gerber Approval")),
    ("Engineering Question", col_candidates("Engineering Question", "Engineering Question", "Engineering Question")),
    ("Pricing & Qty issue", col_candidates("Pricing & Qty issue", "Pricing & Qty issue", "Pricing & Qty issue")),
    ("Ocean Handling Charge (FOB TW)", col_candidates("Ocean Handling Charge (FOB TW)", "Ocean Handling Charge (FOB TW)", "Ocean Handling Charge (FOB TW)")),
    ("T/T", col_candidates("T/T")),
    ("Note", col_candidates("Note", "情況", "Remark", "備註")),
    ("新/舊料號", col_candidates("新/舊料號", "新/舊料號")),
    ("板層", col_candidates("板層", "板層")),
    ("工廠出貨注意事項", col_candidates("工廠出貨注意事項", "工廠出貨事項", "工廠出貨注意事項")),
    ("快遞出貨注意事項", col_candidates("快遞出貨注意事項", "快遞出貨注意事項")),
    ("西拓訂單編號", col_candidates("西拓訂單編號", "西拓訂單編號")),
    ("出貨報告", col_candidates("出貨報告", "出貨報告")),
    ("MADE IN USA", col_candidates("MADE IN USA", "MADE IN USA", "MADE IN USA")),
    ("工廠重量", col_candidates("工廠重量", "工廠重量")),
    ("文件", col_candidates("文件")),
    ("包裝明細", col_candidates("包裝明細", "包裝明細")),
    ("樣板需求", col_candidates("樣板需求", "樣板需求")),
    ("發票", col_candidates("發票")),
]
SANDY_SALES_BASE_SPECS = [
    ("客戶", CUSTOMER_CANDIDATES + ["Customer"]),
    ("PO#", PO_CANDIDATES),
    ("P/N", PART_CANDIDATES),
    ("Order Q'TY (PCS)", QTY_CANDIDATES + ["Order QTY (PCS)"]),
    ("Dock", col_candidates("Dock")),
    ("Ship date", SHIP_DATE_CANDIDATES),
    ("WIP", WIP_CANDIDATES),
    ("工廠交期", FACTORY_DUE_CANDIDATES),
    ("交期 (更改)", col_candidates("交期 (更改)", "交期 (更改)", "交期 (更改)", "交期 (更改)")),
    ("併貨日期 (限內部使用)", col_candidates("併貨日期 (限內部使用)", "併貨日期 (限內部使用)", "併貨日期 (限內部使用)")),
    ("工廠", FACTORY_CANDIDATES),
    ("Ship to", col_candidates("Ship to")),
    ("Ship via", col_candidates("Ship via", " Ship via")),
    ("Tracking No.", col_candidates("Tracking No.", "Tracking No")),
    ("Note", col_candidates("Note", "情況", "Remark", "備註")),
]
def normalize_col_key(col_name):
    s = str(col_name or "")
    s = s.replace("\n", "")
    s = re.sub(r"\s+", "", s)
    return s.strip().lower()
def first_existing_column(df: pd.DataFrame, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    normalized_map = {}
    for col in df.columns:
        key = normalize_col_key(col)
        if key not in normalized_map:
            normalized_map[key] = col
    for c in candidates:
        key = normalize_col_key(c)
        if key in normalized_map:
            return normalized_map[key]
    return None
def first_existing_series(df: pd.DataFrame, candidates):
    src = first_existing_column(df, candidates)
    if not src:
        return None, None
    series = get_series_by_col(df, src)
    return src, series
def build_teable_view_df(source_df: pd.DataFrame, specs):
    view_df = pd.DataFrame(index=source_df.index)
    mapping = {}
    for out_name, candidates in specs:
        src, series = first_existing_series(source_df, candidates)
        mapping[out_name] = src
        if series is not None:
            view_df[out_name] = series
        else:
            view_df[out_name] = ""
    view_df.columns = make_unique_columns(view_df.columns)
    return view_df, mapping
def apply_customer_filter(display_df: pd.DataFrame, customer_col_name: str, default_customer: str | None, key_prefix: str):
    if customer_col_name not in display_df.columns:
        return display_df
    customer_values = sorted(
        [str(x).strip() for x in display_df[customer_col_name].dropna().unique().tolist() if str(x).strip()]
    )
    if not customer_values:
        return display_df
    if default_customer and default_customer in customer_values:
        default_index = customer_values.index(default_customer) + 1
    else:
        default_index = 0
    selected_customer = st.selectbox(
        "客戶篩選",
        ["全部"] + customer_values,
        index=default_index,
        key=f"{key_prefix}_customer_filter",
    )
    if selected_customer != "全部":
        display_df = display_df[
            display_df[customer_col_name].astype(str).str.strip().str.lower()
            == selected_customer.strip().lower()
        ].copy()
    return display_df
def normalize_status_text(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series("", index=pd.RangeIndex(0))
    return series.fillna("").astype(str).str.strip().str.lower()
def _today_normalized() -> pd.Timestamp:
    return pd.Timestamp.today().normalize()
def _series_nonblank(series: pd.Series | None, index_like) -> pd.Series:
    if series is None:
        return pd.Series(False, index=index_like)
    s = series.fillna("").astype(str).str.strip()
    return s.ne("")
def _resolve_wip_series(df: pd.DataFrame) -> pd.Series:
    if "WIP" in df.columns:
        exact = get_series_by_col(df, "WIP")
        if exact is not None:
            return exact
    src, series = first_existing_series(df, ["WIP"] + WIP_CANDIDATES)
    if series is not None:
        return series
    if 'wip_col' in globals() and wip_col:
        fallback = get_series_by_col(df, wip_col)
        if fallback is not None:
            return fallback
    return pd.Series("", index=df.index)
def _wip_exclude_mask(df: pd.DataFrame) -> pd.Series:
    wip_norm = normalize_status_text(_resolve_wip_series(df))
    return wip_norm.str.contains(r"\b(shipment|cancelled|cancell|cancel)\b|取消", na=False).fillna(False)
def parse_mixed_date_series(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype="datetime64[ns]")
    s = series.astype(str).str.strip()
    s = s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    out = pd.to_datetime(s, errors="coerce")
    mask = out.isna()
    if mask.any():
        s2 = s[mask].str.replace(".", "", regex=False)
        out.loc[mask] = pd.to_datetime(s2, errors="coerce")
    mask = out.isna()
    if mask.any():
        def _parse_one(v):
            txt = str(v).strip()
            if not txt or txt.lower() in {"nan", "none"}:
                return pd.NaT
            txt = txt.replace(".", "").replace("  ", " ")
            patterns = [
                "%Y-%m-%d", "%Y/%m/%d",
                "%b %d,%y", "%b %d, %y",
                "%B %d,%y", "%B %d, %y",
                "%m/%d/%y", "%m/%d/%Y",
            ]
            for fmt in patterns:
                try:
                    return pd.Timestamp(datetime.strptime(txt, fmt))
                except Exception:
                    pass
            try:
                return pd.to_datetime(txt, errors="coerce")
            except Exception:
                return pd.NaT
        out.loc[mask] = s[mask].apply(_parse_one)
    return out
def build_subset_mask_new_order_today(df: pd.DataFrame) -> pd.Series:
    today = _today_normalized()
    order_date_col = first_existing_column(
        df,
        ORDER_DATE_CANDIDATES + ["客戶下單日期", "工廠下單日期", "下單日期", "接單日期"]
    )
    order_dt = parse_mixed_date_series(get_series_by_col(df, order_date_col)) if order_date_col else pd.Series(pd.NaT, index=df.index)
    order_today = order_dt.dt.normalize() == today
    due_col = first_existing_column(
        df,
        SHIP_DATE_CANDIDATES + ["交期 (更改)", "交期 (更改)", "交期", "客戶交期", "預交日", "預定交期", "交貨期"]
    )
    due_dt = parse_mixed_date_series(get_series_by_col(df, due_col)) if due_col else pd.Series(pd.NaT, index=df.index)
    due_today = due_dt.dt.normalize() == today
    change_col = first_existing_column(
        df,
        ["更改 Booking", "更改 Booking", "Ship via change", "出貨方式變更"]
    )
    change_flag = _series_nonblank(get_series_by_col(df, change_col) if change_col else None, df.index)
    exclude_flag = _wip_exclude_mask(df)
    mask = (order_today.fillna(False) | due_today.fillna(False) | change_flag.fillna(False)) & (~exclude_flag)
    return mask.fillna(False)
def build_subset_mask_unshipped(df: pd.DataFrame) -> pd.Series:
    exclude_flag = _wip_exclude_mask(df)
    ship_date_name = first_existing_column(df, SHIP_DATE_CANDIDATES + ["出貨日期", "出貨日期 (公式)"])
    ship_s = get_series_by_col(df, ship_date_name) if ship_date_name else None
    ship_dt = parse_mixed_date_series(ship_s) if ship_s is not None else pd.Series(pd.NaT, index=df.index)
    if ship_dt.isna().all():
        due_date_name = first_existing_column(df, FACTORY_DUE_CANDIDATES)
        due_s = get_series_by_col(df, due_date_name) if due_date_name else None
        ship_dt = parse_mixed_date_series(due_s) if due_s is not None else pd.Series(pd.NaT, index=df.index)
    year_2026_flag = ship_dt.dt.year == 2026
    mask = (~exclude_flag) & year_2026_flag.fillna(False)
    return mask.fillna(False)
def build_subset_mask_shipment_current_month(df: pd.DataFrame) -> pd.Series:
    today = pd.Timestamp.today()
    current_month = today.strftime("%Y-%m")
    wip_norm = normalize_status_text(_resolve_wip_series(df))
    shipment_flag = wip_norm.str.contains(r"\bshipment\b", na=False)
    cancel_flag = wip_norm.str.contains(r"\b(cancel|cancell|cancelled)\b|取消", na=False)
    ship_date_name = first_existing_column(df, SHIP_DATE_CANDIDATES + ["出貨日期", "出貨日期 (公式)"])
    ship_s = get_series_by_col(df, ship_date_name) if ship_date_name else None
    ship_dt = parse_mixed_date_series(ship_s) if ship_s is not None else pd.Series(pd.NaT, index=df.index)
    month_flag = ship_dt.dt.strftime("%Y-%m") == current_month
    mask = shipment_flag & (~cancel_flag) & month_flag.fillna(False)
    return mask.fillna(False)
def build_subset_mask(source_df: pd.DataFrame, subset_mode: str | None = None) -> pd.Series:
    if subset_mode == "new_order_today":
        return build_subset_mask_new_order_today(source_df)
    if subset_mode == "unshipped":
        return build_subset_mask_unshipped(source_df)
    if subset_mode == "shipment_only":
        return build_subset_mask_shipment_current_month(source_df)
    return pd.Series(True, index=source_df.index).fillna(False)
def render_teable_subset_table(
    title: str,
    source_df: pd.DataFrame,
    specs,
    default_customer: str | None = None,
    csv_name: str | None = None,
    caption: str | None = None,
    subset_mode: str | None = None,
):
    st.subheader(title)
    if source_df is None or source_df.empty:
        st.warning("Teable 主表目前沒有資料。")
        return
    filtered_source = source_df.copy()
    if subset_mode:
        filtered_source = filtered_source[build_subset_mask(filtered_source, subset_mode)].copy()
    display_df, mapping = build_teable_view_df(filtered_source, specs)
    if default_customer:
        customer_display_col = "客戶" if "客戶" in display_df.columns else ("Customer" if "Customer" in display_df.columns else None)
        if customer_display_col:
            display_df = apply_customer_filter(display_df, customer_display_col, default_customer, title)
    if caption:
        st.caption(caption)
    else:
        if subset_mode == "new_order_today":
            st.caption("資料來源：Teable 主表即時欄位（當天下單新單，或當天客戶交期，或有出貨方式變更；已排除 shipment / cancelled）")
        elif subset_mode == "unshipped":
            st.caption("資料來源：Teable 主表即時欄位（Sandy 內部 WIP：已扣除 WIP 為 shipment / cancelled，且只顯示 2026 年数据）")
        elif subset_mode == "shipment_only":
            st.caption("資料來源：Teable 主表即時欄位（Sandy 銷貨底：只顯示當月出貨者）")
        else:
            st.caption("資料來源：Teable 主表即時欄位")
    st.dataframe(display_df, use_container_width=True, height=520)
    out_name = csv_name or f"{title}.csv"
    st.download_button(
        f"下載 {out_name}",
        data=display_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=out_name,
        mime="text/csv",
        key=f"download_{title}"
    )
def parse_numeric_series(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=float)
    s = series.astype(str).fillna("").str.strip()
    replacements = [
        ("US$", ""), ("USD$", ""), ("NT$", ""), ("HK$", ""),
        ("USD", ""), ("US", ""), ("NTD", ""), ("TWD", ""), ("RMB", ""),
        ("$", ""), (",", ""), ("nan", ""), ("None", ""),
    ]
    for old, new in replacements:
        s = s.str.replace(old, new, regex=False)
    s = s.str.extract(r"([-+]?\d*\.?\d+)", expand=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)
def find_amount_column(df: pd.DataFrame, candidates):
    src = first_existing_column(df, candidates)
    if src:
        return src
    normalized_cols = {normalize_col_key(c): c for c in df.columns}
    candidate_keys = [normalize_col_key(c) for c in candidates]
    for key, col in normalized_cols.items():
        for ck in candidate_keys:
            if ck and ck in key:
                return col
    for col in df.columns:
        key = normalize_col_key(col)
        if any(token in key for token in ["金額", "amount"]):
            return col
    return None
def safe_display_subset(df: pd.DataFrame, columns):
    out = pd.DataFrame(index=df.index)
    for col in columns:
        if col in df.columns:
            out[col] = get_series_by_col(df, col)
    out.columns = make_unique_columns(out.columns)
    return out
# ================================
# ✅ 業績明細表 (修正版)
# ================================
def render_sales_detail_from_teable(source_df: pd.DataFrame):
    """
    業績明細表
    規則：
    - 已出貨：WIP == SHIPMENT，且「出貨日期」落在當月
    - 預計出貨：WIP != SHIPMENT，且「Ship date」落在當月
    - 接單金額：依「客戶下單日期」落在當月，金額取「接單金額」
    - 銷貨合計：已出貨 + 預計出貨
    """
    import datetime as _dt
    from pathlib import Path

    st.subheader("📊 業績明細表")

    if source_df is None or source_df.empty:
        st.warning("Teable 主表目前沒有資料。")
        return

    st.caption("資料來源：Teable 主表即時欄位（全客戶）")

    # ── 欄位偵測 ──────────────────────────────────────
    customer_col_local = first_existing_column(source_df, CUSTOMER_CANDIDATES + ["Customer"])
    factory_col_local = first_existing_column(source_df, FACTORY_CANDIDATES)
    po_col_local = first_existing_column(source_df, PO_CANDIDATES)
    pn_col_local = first_existing_column(source_df, PART_CANDIDATES)
    qty_col_local = first_existing_column(source_df, QTY_CANDIDATES + ["Order QTY (PCS)"])
    order_date_col = first_existing_column(source_df, ORDER_DATE_CANDIDATES)
    ship_plan_col = first_existing_column(source_df, ["Ship date", "Ship Date"] + SHIP_DATE_CANDIDATES)
    ship_actual_col = first_existing_column(source_df, ["出貨日期", "出貨日期_排序", "Actual Ship Date"])
    wip_col_local = first_existing_column(source_df, WIP_CANDIDATES)
    invoice_col = first_existing_column(source_df, ["銷貨金額", "Shipment Amount", "Ship Amount", "Invoice Amount", "Invoice", "INVOICE"])
    order_amt_col = first_existing_column(source_df, ["接單金額", "Order Amount"] + AMOUNT_ORDER_CANDIDATES)

    def _parse_usd(series):
        if series is None:
            return pd.Series(0.0, index=source_df.index)
        s = (
            series.astype(str)
            .str.replace("US$", "", regex=False)
            .str.replace("USD", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        return pd.to_numeric(s, errors="coerce").fillna(0.0)

    invoice_series = _parse_usd(get_series_by_col(source_df, invoice_col) if invoice_col else None)
    order_amt_series = _parse_usd(get_series_by_col(source_df, order_amt_col) if order_amt_col else None)
    if order_amt_col is None:
        order_amt_series = invoice_series.copy()

    order_dates = parse_mixed_date_series(get_series_by_col(source_df, order_date_col)) if order_date_col else pd.Series(pd.NaT, index=source_df.index)
    ship_plan_dates = parse_mixed_date_series(get_series_by_col(source_df, ship_plan_col)) if ship_plan_col else pd.Series(pd.NaT, index=source_df.index)
    ship_actual_dates = parse_mixed_date_series(get_series_by_col(source_df, ship_actual_col)) if ship_actual_col else pd.Series(pd.NaT, index=source_df.index)

    if wip_col_local:
        wip_series = get_series_by_col(source_df, wip_col_local).fillna("").astype(str).str.strip().str.upper()
    else:
        wip_series = pd.Series("", index=source_df.index)

    is_shipment = wip_series.eq("SHIPMENT")
    is_cancelled = wip_series.str.contains("CANCEL", na=False)

    # 月份選單：以實際出貨日期 / 預計 ship date / 接單日期共同組成
    actual_months = [m for m in ship_actual_dates.dt.strftime("%Y-%m").dropna().tolist() if m and m != "NaT"]
    plan_months = [m for m in ship_plan_dates.dt.strftime("%Y-%m").dropna().tolist() if m and m != "NaT"]
    order_months = [m for m in order_dates.dt.strftime("%Y-%m").dropna().tolist() if m and m != "NaT"]
    all_months = sorted(set(actual_months + plan_months + order_months), reverse=True)
    if not all_months:
        st.warning("找不到有效的日期資料。")
        return

    current_month = _dt.datetime.now().strftime("%Y-%m")
    default_idx = all_months.index(current_month) if current_month in all_months else 0
    selected_month = st.selectbox(
        "📅 選擇統計月份",
        all_months,
        index=default_idx,
        format_func=lambda m: f"{m[:4]} 年 {int(m[5:7])} 月",
        key="sales_detail_month_teable",
    )
    mon_str = selected_month[5:7]

    actual_month_mask = ship_actual_dates.dt.strftime("%Y-%m") == selected_month
    plan_month_mask = ship_plan_dates.dt.strftime("%Y-%m") == selected_month
    order_month_mask = order_dates.dt.strftime("%Y-%m") == selected_month

    # 已出貨看「出貨日期」；預計出貨看「Ship date」
    shipped_mask = is_shipment & actual_month_mask.fillna(False)
    forecast_mask = (~is_shipment) & (~is_cancelled) & plan_month_mask.fillna(False)
    order_mask = order_month_mask.fillna(False)

    # 歷史補登
    manual_path = Path("sales_manual_history.csv")
    manual_df = pd.DataFrame()
    manual_shipped_usd = manual_forecast_usd = manual_order_usd = 0.0
    if manual_path.exists():
        try:
            manual_df = pd.read_csv(manual_path)
            if not manual_df.empty:
                m_date_col = first_existing_column(manual_df, ["日期", "Date"])
                m_type_col = first_existing_column(manual_df, ["類型", "Type"])
                m_amt_col = first_existing_column(manual_df, ["金額(USD)", "金額", "Amount", "Amount(USD)"])
                if m_date_col and m_type_col and m_amt_col:
                    m_dates = parse_mixed_date_series(manual_df[m_date_col])
                    m_month_mask = m_dates.dt.strftime("%Y-%m") == selected_month
                    m_type = manual_df[m_type_col].fillna("").astype(str).str.strip()
                    m_amt = _parse_usd(manual_df[m_amt_col])
                    manual_shipped_usd = m_amt[m_month_mask & m_type.isin(["已出貨", "SHIPMENT"])].sum()
                    manual_forecast_usd = m_amt[m_month_mask & m_type.isin(["預計出貨", "Forecast", "QA"])].sum()
                    manual_order_usd = m_amt[m_month_mask & m_type.isin(["接單", "Order"])].sum()
        except Exception:
            pass

    shipped_usd = invoice_series[shipped_mask].sum() + manual_shipped_usd
    forecast_usd = invoice_series[forecast_mask].sum() + manual_forecast_usd
    order_usd = order_amt_series[order_mask].sum() + manual_order_usd
    total_usd = shipped_usd + forecast_usd

    today_str = pd.Timestamp.now().strftime("%Y/%m/%d")
    st.markdown(
        f"<h3 style='text-align:center;margin-bottom:4px;'>{int(mon_str)}月 業績明細表</h3>"
        f"<p style='text-align:right;color:gray;margin-top:0;'>{today_str}</p>",
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📥 接單金額 (USD)", f"${order_usd:,.2f}" if order_usd else "—")
    c2.metric("✅ 已出貨 (USD)", f"${shipped_usd:,.2f}" if shipped_usd else "—")
    c3.metric("🔜 預計出貨 (USD)", f"${forecast_usd:,.2f}" if forecast_usd else "—")
    c4.metric("📊 銷貨合計 (USD)", f"${total_usd:,.2f}" if total_usd else "—")

    qa_mask = forecast_mask & wip_series.isin(["QA", "QC", "FQC", "INSPECTION"])
    qa_sum = invoice_series[qa_mask].sum()
    if qa_sum > 0:
        qa_names = ""
        if customer_col_local:
            names = [
                str(n).split()[0]
                for n in get_series_by_col(source_df, customer_col_local)[qa_mask].dropna().unique()
            ]
            qa_names = "、".join(names)
        st.info(
            f"💡 QA 中（預計近期出貨）：{qa_names} ｜ US${qa_sum:,.2f}\n\n"
            f"出貨後 WIP 更新為 SHIPMENT，{int(mon_str)}月銷貨合計將增至 **US${total_usd:,.2f}**"
        )

    st.markdown("---")

    def _clean_factory(s):
        return re.sub(r'\s+', '', str(s)) if s else ""

    def _clean_customer(s):
        m = re.match(r'^([A-Za-z0-9_\-]+)', str(s))
        return m.group(1) if m else (str(s).split()[0] if str(s).split() else str(s))

    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown(f"#### 🏭 工廠別（已出貨 {int(mon_str)}月）")
        if shipped_mask.any() and factory_col_local:
            fac_s = get_series_by_col(source_df, factory_col_local)[shipped_mask]
            amt_s = invoice_series[shipped_mask]
            fdf = pd.DataFrame({"工廠": fac_s.apply(_clean_factory), "金額": amt_s})
            grp = fdf.groupby("工廠")["金額"].agg(訂單數="count", 銷貨金額_USD="sum").reset_index()
            grp = grp.sort_values("銷貨金額_USD", ascending=False)
            grp["銷貨金額 (USD)"] = grp["銷貨金額_USD"].apply(lambda x: f"${x:,.2f}")
            grp = grp[["工廠", "訂單數", "銷貨金額 (USD)"]]
            total_row = pd.DataFrame([{"工廠": "合計", "訂單數": int(shipped_mask.sum()), "銷貨金額 (USD)": f"${invoice_series[shipped_mask].sum():,.2f}"}])
            grp = pd.concat([grp, total_row], ignore_index=True)
            st.dataframe(grp, use_container_width=True, hide_index=True, height=260)
        else:
            st.info("本月無已出貨資料。")

    with col_right:
        st.markdown(f"#### 👥 客戶別（已出貨 {int(mon_str)}月）")
        if shipped_mask.any() and customer_col_local:
            cust_s = get_series_by_col(source_df, customer_col_local)[shipped_mask]
            amt_s = invoice_series[shipped_mask]
            cdf = pd.DataFrame({"客戶": cust_s.apply(_clean_customer), "金額": amt_s})
            grp2 = cdf.groupby("客戶")["金額"].agg(訂單數="count", 銷貨金額_USD="sum").reset_index()
            grp2 = grp2.sort_values("銷貨金額_USD", ascending=False)
            grp2["銷貨金額 (USD)"] = grp2["銷貨金額_USD"].apply(lambda x: f"${x:,.2f}")
            grp2 = grp2[["客戶", "訂單數", "銷貨金額 (USD)"]]
            total_row2 = pd.DataFrame([{"客戶": "合計", "訂單數": int(shipped_mask.sum()), "銷貨金額 (USD)": f"${invoice_series[shipped_mask].sum():,.2f}"}])
            grp2 = pd.concat([grp2, total_row2], ignore_index=True)
            st.dataframe(grp2, use_container_width=True, hide_index=True, height=260)
        else:
            st.info("本月無已出貨資料。")

    st.markdown(f"#### 📈 出貨累計走勢（{int(mon_str)}月，USD）")
    if shipped_mask.any():
        daily_df = pd.DataFrame({"日期": ship_actual_dates[shipped_mask], "金額": invoice_series[shipped_mask]}).dropna(subset=["日期"]).sort_values("日期")
        daily_df["累計出貨(USD)"] = daily_df["金額"].cumsum()
        daily_df["日"] = daily_df["日期"].dt.strftime("%m/%d")
        st.line_chart(daily_df.set_index("日")[["累計出貨(USD)"]], height=220)
    else:
        st.info("本月尚無出貨資料。")

    st.markdown(f"#### ✅ 已出貨明細（SHIPMENT，{int(mon_str)}月）")
    if shipped_mask.any():
        show_cols = [c for c in [customer_col_local, po_col_local, pn_col_local, factory_col_local, ship_actual_col, invoice_col] if c and c in source_df.columns]
        view = source_df.loc[shipped_mask, show_cols].copy() if show_cols else source_df.loc[shipped_mask].copy()
        if customer_col_local and customer_col_local in view.columns:
            view[customer_col_local] = view[customer_col_local].apply(_clean_customer)
        if factory_col_local and factory_col_local in view.columns:
            view[factory_col_local] = view[factory_col_local].apply(_clean_factory)
        if ship_actual_col and ship_actual_col in view.columns:
            view = view.sort_values(ship_actual_col, ascending=False, na_position="last")
        st.dataframe(view, use_container_width=True, hide_index=True, height=320)
        st.download_button(
            f"⬇️ 下載已出貨明細 CSV ({selected_month})",
            data=view.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"shipment_{selected_month}.csv",
            mime="text/csv",
            key="download_shipped_detail_csv",
        )
    else:
        st.info("本月尚無已出貨（SHIPMENT）資料。")

    if forecast_mask.any() or (not manual_df.empty and manual_forecast_usd > 0):
        st.markdown(f"#### 🔜 預計出貨（非 SHIPMENT，{int(mon_str)}月）")
        show_cols2 = [c for c in [customer_col_local, po_col_local, wip_col_local, factory_col_local, ship_plan_col, invoice_col] if c and c in source_df.columns]
        view2 = source_df.loc[forecast_mask, show_cols2].copy() if show_cols2 else source_df.loc[forecast_mask].copy()
        if customer_col_local and customer_col_local in view2.columns:
            view2[customer_col_local] = view2[customer_col_local].apply(_clean_customer)
        if factory_col_local and factory_col_local in view2.columns:
            view2[factory_col_local] = view2[factory_col_local].apply(_clean_factory)
        if not manual_df.empty and manual_forecast_usd > 0:
            m_date_col = first_existing_column(manual_df, ["日期", "Date"])
            m_type_col = first_existing_column(manual_df, ["類型", "Type"])
            m_amt_col = first_existing_column(manual_df, ["金額(USD)", "金額", "Amount", "Amount(USD)"])
            m_customer_col = first_existing_column(manual_df, ["客戶", "Customer"])
            m_factory_col = first_existing_column(manual_df, ["工廠", "Factory"])
            if m_date_col and m_type_col and m_amt_col:
                m_dates = parse_mixed_date_series(manual_df[m_date_col])
                m_mask = (m_dates.dt.strftime("%Y-%m") == selected_month) & manual_df[m_type_col].fillna("").astype(str).str.strip().isin(["預計出貨", "Forecast", "QA"])
                if m_mask.any():
                    manual_view = pd.DataFrame({
                        customer_col_local or "客戶": manual_df.loc[m_mask, m_customer_col] if m_customer_col else "",
                        wip_col_local or "WIP": "手動補登",
                        factory_col_local or "工廠": manual_df.loc[m_mask, m_factory_col] if m_factory_col else "",
                        ship_plan_col or "Ship date": manual_df.loc[m_mask, m_date_col],
                        invoice_col or "銷貨金額": manual_df.loc[m_mask, m_amt_col],
                    })
                    view2 = pd.concat([view2, manual_view], ignore_index=True)
        st.dataframe(view2, use_container_width=True, hide_index=True, height=220)

    st.markdown("---")
    st.markdown("#### 📅 月度出貨走勢（近 12 個月，USD）")
    monthly_mask = is_shipment & ship_actual_dates.notna()
    if monthly_mask.any():
        monthly_df = pd.DataFrame({"月份": ship_actual_dates[monthly_mask].dt.to_period("M").astype(str), "金額": invoice_series[monthly_mask]})
        monthly_grp = monthly_df.groupby("月份", as_index=False)["金額"].sum().sort_values("月份").tail(12).set_index("月份")
        monthly_grp.columns = ["出貨金額(USD)"]
        st.bar_chart(monthly_grp, height=240)
    else:
        st.info("無足夠歷史資料。")

    if order_mask.any():
        st.markdown(f"#### 📥 接單明細（{int(mon_str)}月）")
        show_cols3 = [c for c in [customer_col_local, po_col_local, pn_col_local, factory_col_local, order_date_col, order_amt_col] if c and c in source_df.columns]
        view3 = source_df.loc[order_mask, show_cols3].copy() if show_cols3 else source_df.loc[order_mask].copy()
        if customer_col_local and customer_col_local in view3.columns:
            view3[customer_col_local] = view3[customer_col_local].apply(_clean_customer)
        if factory_col_local and factory_col_local in view3.columns:
            view3[factory_col_local] = view3[factory_col_local].apply(_clean_factory)
        st.dataframe(view3, use_container_width=True, hide_index=True, height=220)
        st.download_button(
            f"⬇️ 下載接單明細 CSV ({selected_month})",
            data=view3.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"order_{selected_month}.csv",
            mime="text/csv",
            key="download_sales_detail_csv_teable",
        )

def show_new_orders_wip_report(source_df: pd.DataFrame):
    render_teable_subset_table(
        title="📄 新訂單 WIP",
        source_df=source_df,
        specs=SANDY_NEW_ORDER_SPECS,
        default_customer=None,
        csv_name="新訂單 WIP.csv",
        subset_mode="new_order_today",
    )


def show_sandy_internal_wip_report(source_df: pd.DataFrame):
    render_teable_subset_table(
        title="📄 Sandy 內部 WIP",
        source_df=source_df,
        specs=SANDY_INTERNAL_WIP_SPECS,
        default_customer=None,
        csv_name="Sandy 內部 WIP.csv",
        subset_mode="unshipped",
    )


def show_sandy_sales_report(source_df: pd.DataFrame):
    render_teable_subset_table(
        title="📄 Sandy 銷貨底",
        source_df=source_df,
        specs=SANDY_SALES_BASE_SPECS,
        default_customer=None,
        csv_name="Sandy 銷貨底.csv",
        subset_mode="shipment_only",
    )
