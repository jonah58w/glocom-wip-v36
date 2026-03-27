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

def make_unique_columns(columns):
    seen = {}
    result = []
    for col in list(columns):
        name = str(col)
        count = seen.get(name, 0)
        if count == 0:
            result.append(name)
        else:
            result.append(f"{name}_{count+1}")
        seen[name] = count + 1
    return result

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
    s = s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "NaT": pd.NA})
    out = pd.to_datetime(s, errors="coerce")

    # Excel serial date support
    numeric = pd.to_numeric(s, errors="coerce")
    excel_mask = out.isna() & numeric.notna()
    if excel_mask.any():
        out.loc[excel_mask] = pd.to_datetime(numeric.loc[excel_mask], unit="D", origin="1899-12-30", errors="coerce")

    mask = out.isna()
    if mask.any():
        s2 = (
            s[mask]
            .str.replace(".", "", regex=False)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )
        out.loc[mask] = pd.to_datetime(s2, errors="coerce")

    mask = out.isna()
    if mask.any():
        def _parse_one(v):
            txt = str(v).strip()
            if not txt or txt.lower() in {"nan", "none", "nat"}:
                return pd.NaT
            txt = re.sub(r"\s+", " ", txt.replace(".", "")).strip()
            patterns = [
                "%Y-%m-%d", "%Y/%m/%d", "%y-%m-%d", "%y/%m/%d",
                "%b %d, %y", "%b %d,%y", "%B %d, %y", "%B %d,%y",
                "%m/%d/%y", "%m/%d/%Y", "%m-%d-%y", "%m-%d-%Y",
                "%d-%b-%y", "%d-%b-%Y",
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

    return pd.to_datetime(out, errors="coerce")
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

    actual_col = first_existing_column(df, ["出貨日期", "出貨日期_排序"])
    planned_col = first_existing_column(df, ["Ship date", "Ship Date", "預計出貨日", "客戶交期"] + SHIP_DATE_CANDIDATES)

    actual_dt = parse_mixed_date_series(get_series_by_col(df, actual_col)) if actual_col else pd.Series(pd.NaT, index=df.index)
    planned_dt = parse_mixed_date_series(get_series_by_col(df, planned_col)) if planned_col else pd.Series(pd.NaT, index=df.index)

    effective_dt = actual_dt.copy()
    if len(effective_dt) != len(df):
        effective_dt = pd.Series(pd.NaT, index=df.index)
    effective_dt.loc[effective_dt.isna()] = planned_dt.loc[effective_dt.isna()]

    month_flag = effective_dt.dt.strftime("%Y-%m") == current_month
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
    1) 接單金額：依接單日期落在所選月份統計
    2) 已確認出貨：WIP=SHIPMENT，優先用「出貨日期」，沒有則回退「Ship date」
    3) 預計本月出貨：WIP!=SHIPMENT，優先用「Ship date」，沒有則回退「出貨日期」
    4) 金額：優先抓「銷貨金額」，若空白或 0，回補「接單金額」
    5) 支援手動補登「今天以前」的歷史資料，直接併入月摘要、工廠/客戶統計、每日統計與明細表
    """
    import datetime as _dt
    from pathlib import Path as _Path

    st.subheader("📊 業績明細表")

    if source_df is None or source_df.empty:
        st.warning("Teable 主表目前沒有資料。")
        return

    st.caption("資料來源：Teable 主表即時欄位（全客戶）＋ 歷史補登資料")

    FX_NTD_PER_USD = 31.5
    MANUAL_HISTORY_FILE = _Path("sales_manual_history.csv")

    def _pick_col(*candidate_groups):
        merged = []
        for group in candidate_groups:
            merged.extend(list(group))
        return first_existing_column(source_df, merged)

    def _parse_date_from_col(col_name):
        if not col_name:
            return pd.Series(pd.NaT, index=source_df.index)
        return parse_mixed_date_series(get_series_by_col(source_df, col_name))

    def _clean_factory(v):
        txt = str(v).strip()
        if txt.lower() in {"nan", "none"}:
            return ""
        return re.sub(r"\s+", "", txt)

    def _clean_customer(v):
        txt = str(v).strip()
        if txt.lower() in {"nan", "none"}:
            return ""
        m = re.match(r"^([A-Za-z0-9_\-]+)", txt)
        return m.group(1) if m else (txt.split()[0] if txt.split() else txt)

    def _fmt_usd(v):
        return f"${float(v):,.2f}"

    def _usd_to_ntd_10k(v):
        return round((float(v) * FX_NTD_PER_USD) / 10000.0, 2)

    def _safe_series(col_name, default=""):
        if not col_name:
            return pd.Series(default, index=source_df.index)
        s = get_series_by_col(source_df, col_name)
        if s is None:
            return pd.Series(default, index=source_df.index)
        return s

    def _read_manual_history() -> pd.DataFrame:
        cols = ["月份", "日期", "類型", "客戶", "工廠", "PO#", "P/N", "QTY", "WIP", "金額(USD)", "備註"]
        if MANUAL_HISTORY_FILE.exists():
            try:
                df = pd.read_csv(MANUAL_HISTORY_FILE, dtype=str).fillna("")
            except Exception:
                df = pd.DataFrame(columns=cols)
        else:
            df = pd.DataFrame(columns=cols)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        return df[cols].copy()

    def _normalize_manual_history(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame(columns=["月份", "日期", "類型", "客戶", "工廠", "PO#", "P/N", "QTY", "WIP", "金額(USD)", "備註"])
        out = df.copy().fillna("")
        out["日期"] = parse_mixed_date_series(out["日期"])
        out["月份"] = out["日期"].dt.to_period("M").astype(str).where(out["日期"].notna(), out["月份"].astype(str).str.strip())
        out["類型"] = out["類型"].astype(str).str.strip().replace({"": "已出貨"})
        out["金額(USD)"] = parse_numeric_series(out["金額(USD)"])
        out["QTY"] = out["QTY"].astype(str)
        return out

    def _manual_rows_to_standard(df_manual: pd.DataFrame) -> pd.DataFrame:
        if df_manual is None or df_manual.empty:
            return pd.DataFrame(columns=["來源", "客戶", "工廠", "PO#", "P/N", "QTY", "WIP", "日期", "金額"])
        out = pd.DataFrame({
            "來源": "手動補登",
            "客戶": df_manual["客戶"].astype(str).map(_clean_customer),
            "工廠": df_manual["工廠"].astype(str).map(_clean_factory),
            "PO#": df_manual["PO#"].astype(str),
            "P/N": df_manual["P/N"].astype(str),
            "QTY": df_manual["QTY"].astype(str),
            "WIP": df_manual["WIP"].astype(str),
            "日期": pd.to_datetime(df_manual["日期"], errors="coerce"),
            "金額": parse_numeric_series(df_manual["金額(USD)"]),
        })
        return out

    customer_col_local = _pick_col(CUSTOMER_CANDIDATES, ["Customer"])
    factory_col_local = _pick_col(FACTORY_CANDIDATES)
    po_col_local = _pick_col(PO_CANDIDATES)
    pn_col_local = _pick_col(PART_CANDIDATES)
    qty_col_local = _pick_col(QTY_CANDIDATES, ["Order QTY (PCS)", "Order QTY", "QTY"])
    wip_col_local = _pick_col(WIP_CANDIDATES)

    order_date_col = _pick_col(ORDER_DATE_CANDIDATES, ["客戶下單日期", "工廠下單日期", "下單日期", "接單日期"])
    actual_ship_col = _pick_col(["出貨日期", "出貨日期_排序"], ["出貨日期"])
    planned_ship_col = _pick_col(["Ship date", "Ship Date", "預計出貨日", "客戶交期"], ["Ship date", "Ship Date"])

    sales_amt_col = find_amount_column(
        source_df,
        ["銷貨金額", "出貨金額", "出貨發票金額", "Invoice Amount", "Invoice Total", "Invoice", "INVOICE", "發票"] + AMOUNT_SHIP_CANDIDATES,
    )
    order_amt_col = find_amount_column(
        source_df,
        ["接單金額", "接單總金額", "Order Amount", "Order Total", "客戶金額", "Total Amount", "Amount"] + AMOUNT_ORDER_CANDIDATES,
    )

    order_dates = _parse_date_from_col(order_date_col)
    actual_ship_dates = _parse_date_from_col(actual_ship_col)
    planned_ship_dates = _parse_date_from_col(planned_ship_col)

    shipped_dates = actual_ship_dates.copy()
    if len(shipped_dates) != len(source_df):
        shipped_dates = pd.Series(pd.NaT, index=source_df.index)
    shipped_dates.loc[shipped_dates.isna()] = planned_ship_dates.loc[shipped_dates.isna()]

    forecast_dates = planned_ship_dates.copy()
    if len(forecast_dates) != len(source_df):
        forecast_dates = pd.Series(pd.NaT, index=source_df.index)
    forecast_dates.loc[forecast_dates.isna()] = actual_ship_dates.loc[forecast_dates.isna()]

    ship_amt_series = parse_numeric_series(get_series_by_col(source_df, sales_amt_col) if sales_amt_col else None)
    order_amt_series = parse_numeric_series(get_series_by_col(source_df, order_amt_col) if order_amt_col else None)
    if len(ship_amt_series) != len(source_df):
        ship_amt_series = pd.Series(0.0, index=source_df.index)
    else:
        ship_amt_series.index = source_df.index
    if len(order_amt_series) != len(source_df):
        order_amt_series = pd.Series(0.0, index=source_df.index)
    else:
        order_amt_series.index = source_df.index

    sales_value_series = ship_amt_series.where(ship_amt_series.ne(0), order_amt_series).fillna(0.0)
    order_value_series = order_amt_series.where(order_amt_series.ne(0), ship_amt_series).fillna(0.0)

    if wip_col_local:
        wip_series = get_series_by_col(source_df, wip_col_local).fillna("").astype(str).str.strip().str.upper()
    else:
        wip_series = pd.Series("", index=source_df.index)

    is_shipment = wip_series.eq("SHIPMENT")
    is_cancelled = wip_series.str.contains(r"CANCEL|取消", na=False)

    order_periods = order_dates.dt.to_period("M")
    actual_periods = actual_ship_dates.dt.to_period("M")
    planned_periods = planned_ship_dates.dt.to_period("M")
    shipped_periods = shipped_dates.dt.to_period("M")
    forecast_periods = forecast_dates.dt.to_period("M")

    manual_history_raw = _read_manual_history()
    manual_history = _normalize_manual_history(manual_history_raw)
    manual_periods = set()
    if not manual_history.empty:
        manual_periods = {
            pd.Period(m, freq="M")
            for m in manual_history["月份"].astype(str)
            if re.match(r"^\d{4}-\d{2}$", str(m))
        }

    all_periods = sorted(
        set(order_periods.dropna().tolist())
        | set(actual_periods.dropna().tolist())
        | set(planned_periods.dropna().tolist())
        | manual_periods,
        reverse=True,
    )
    if not all_periods:
        st.warning("找不到有效的日期資料。")
        return

    current_period = pd.Period(_dt.datetime.now().strftime("%Y-%m"), freq="M")
    default_idx = all_periods.index(current_period) if current_period in all_periods else 0

    selected_period = st.selectbox(
        "📅 選擇統計月份",
        all_periods,
        index=default_idx,
        format_func=lambda p: f"{p.year} 年 {p.month} 月",
        key="sales_detail_month_teable",
    )
    selected_month_str = str(selected_period)

    with st.expander("Debug：業績明細表欄位偵測", expanded=False):
        debug_df = pd.DataFrame(
            [
                ("接單日期欄", order_date_col),
                ("實際出貨日期欄", actual_ship_col),
                ("預計出貨日期欄", planned_ship_col),
                ("接單金額欄", order_amt_col),
                ("銷貨金額欄", sales_amt_col),
                ("WIP 欄", wip_col_local),
                ("客戶欄", customer_col_local),
                ("工廠欄", factory_col_local),
            ],
            columns=["項目", "偵測結果"],
        )
        st.dataframe(debug_df, use_container_width=True, hide_index=True)
        st.caption("已出貨使用：出貨日期優先；預計出貨使用：Ship date 優先；金額使用：銷貨金額優先，空白才回補接單金額。")

    st.markdown("#### 📝 歷史補登（今天以前的金額）")
    st.caption("若今天以前的資料還沒進 Teable，可直接在下面 key in，儲存後會自動併入本月業績明細表。")
    today_cutoff = pd.Timestamp.today().normalize()

    manual_edit_df = manual_history_raw.copy()
    if manual_edit_df.empty:
        manual_edit_df = pd.DataFrame(columns=["月份", "日期", "類型", "客戶", "工廠", "PO#", "P/N", "QTY", "WIP", "金額(USD)", "備註"])
    month_rows = manual_edit_df[manual_edit_df["月份"].astype(str).eq(selected_month_str)].copy()
    if month_rows.empty:
        month_rows = pd.DataFrame([{
            "月份": selected_month_str,
            "日期": "",
            "類型": "已出貨",
            "客戶": "",
            "工廠": "",
            "PO#": "",
            "P/N": "",
            "QTY": "",
            "WIP": "SHIPMENT",
            "金額(USD)": "",
            "備註": "",
        }])

    edited_rows = st.data_editor(
        month_rows,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        key=f"manual_sales_editor_{selected_month_str}",
        column_config={
            "月份": st.column_config.TextColumn(disabled=True),
            "日期": st.column_config.TextColumn(help="請輸入 YYYY-MM-DD 或 M/D/YY"),
            "類型": st.column_config.SelectboxColumn(options=["接單", "已出貨", "預計出貨"]),
            "金額(USD)": st.column_config.NumberColumn(format="%.2f"),
        },
    )
    csave1, csave2 = st.columns([1, 4])
    with csave1:
        if st.button("💾 儲存本月補登", key=f"save_manual_sales_{selected_month_str}"):
            save_df = manual_edit_df[~manual_edit_df["月份"].astype(str).eq(selected_month_str)].copy()
            edited_rows = edited_rows.copy().fillna("")
            edited_rows["月份"] = selected_month_str
            edited_rows = edited_rows[
                (edited_rows["日期"].astype(str).str.strip() != "")
                | (edited_rows["金額(USD)"].astype(str).str.strip() != "")
            ]
            save_df = pd.concat([save_df, edited_rows], ignore_index=True)
            save_df.to_csv(MANUAL_HISTORY_FILE, index=False, encoding="utf-8-sig")
            st.success(f"已儲存 {selected_period.month} 月補登資料。")
            st.rerun()
    with csave2:
        st.caption(f"補登資料儲存在：{MANUAL_HISTORY_FILE.name}")

    manual_history = _normalize_manual_history(_read_manual_history())
    manual_month = manual_history[manual_history["月份"].astype(str).eq(selected_month_str)].copy()
    if not manual_month.empty:
        manual_month = manual_month[manual_month["日期"].notna()].copy()
        manual_month = manual_month[manual_month["日期"].dt.normalize().le(today_cutoff)].copy()
        manual_month["類型"] = manual_month["類型"].astype(str).str.strip()
        manual_month["WIP"] = manual_month["WIP"].astype(str).str.strip()
        manual_month.loc[manual_month["WIP"].eq("") & manual_month["類型"].eq("已出貨"), "WIP"] = "SHIPMENT"
        manual_month.loc[manual_month["WIP"].eq("") & manual_month["類型"].eq("預計出貨"), "WIP"] = "QA"

    shipped_mask = is_shipment & (shipped_periods == selected_period).fillna(False) & (~is_cancelled)
    forecast_mask = (~is_shipment) & (forecast_periods == selected_period).fillna(False) & (~is_cancelled)
    order_mask = (order_periods == selected_period).fillna(False) & (~is_cancelled)

    teable_shipped_usd = float(sales_value_series[shipped_mask].sum())
    teable_forecast_usd = float(sales_value_series[forecast_mask].sum())
    teable_order_usd = float(order_value_series[order_mask].sum())

    manual_order_df = manual_month[manual_month["類型"].eq("接單")].copy() if not manual_month.empty else pd.DataFrame()
    manual_shipped_df = manual_month[manual_month["類型"].eq("已出貨")].copy() if not manual_month.empty else pd.DataFrame()
    manual_forecast_df = manual_month[manual_month["類型"].eq("預計出貨")].copy() if not manual_month.empty else pd.DataFrame()

    manual_order_usd = float(manual_order_df["金額(USD)"].sum()) if not manual_order_df.empty else 0.0
    manual_shipped_usd = float(manual_shipped_df["金額(USD)"].sum()) if not manual_shipped_df.empty else 0.0
    manual_forecast_usd = float(manual_forecast_df["金額(USD)"].sum()) if not manual_forecast_df.empty else 0.0

    shipped_usd = teable_shipped_usd + manual_shipped_usd
    forecast_usd = teable_forecast_usd + manual_forecast_usd
    order_usd = teable_order_usd + manual_order_usd
    total_usd = shipped_usd + forecast_usd

    shipped_count = int(shipped_mask.sum()) + len(manual_shipped_df)
    forecast_count = int(forecast_mask.sum()) + len(manual_forecast_df)
    order_count = int(order_mask.sum()) + len(manual_order_df)

    today_str = pd.Timestamp.now().strftime("%Y/%m/%d")
    st.markdown(
        f"<h3 style='text-align:center;margin-bottom:4px;'>{selected_period.month}月 業績明細表</h3>"
        f"<p style='text-align:right;color:gray;margin-top:0;'>{today_str}</p>",
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📥 接單金額 (USD)", _fmt_usd(order_usd) if order_usd else "—", delta=f"{order_count} 筆")
    c2.metric("✅ 已確認出貨 (USD)", _fmt_usd(shipped_usd) if shipped_usd else "—", delta=f"{shipped_count} 筆")
    c3.metric("🔜 預計本月出貨 (USD)", _fmt_usd(forecast_usd) if forecast_usd else "—", delta=f"{forecast_count} 筆")
    c4.metric("📊 月銷貨合計 (USD)", _fmt_usd(total_usd) if total_usd else "—")

    st.markdown(
        "\n".join([
            f"- 已確認出貨（SHIPMENT）: US${shipped_usd:,.2f}",
            f"- 預計 {selected_period.month} 月出貨（未 SHIPMENT）: US${forecast_usd:,.2f}",
            f"- {selected_period.month} 月份銷貨金額總計: US${total_usd:,.2f}",
        ])
    )
    if manual_shipped_usd or manual_forecast_usd or manual_order_usd:
        st.info(f"本月已併入手動補登：接單 US${manual_order_usd:,.2f}、已出貨 US${manual_shipped_usd:,.2f}、預計出貨 US${manual_forecast_usd:,.2f}。")

    qa_mask = forecast_mask & wip_series.isin(["QA", "QC", "FQC", "INSPECTION"])
    if qa_mask.any():
        qa_sum = float(sales_value_series[qa_mask].sum())
        qa_dates = forecast_dates[qa_mask].dropna().sort_values()
        qa_date_txt = qa_dates.iloc[0].strftime("%-m/%-d") if len(qa_dates) else ""
        qa_names = "、".join(
            [_clean_customer(x) for x in _safe_series(customer_col_local)[qa_mask].dropna().astype(str).unique().tolist()][:5]
        )
        if qa_sum > 0:
            st.info(f"💡 QA 中（預計{qa_date_txt}出貨）：{qa_names}｜US${qa_sum:,.2f}\n\n出貨後 WIP 更新為 SHIPMENT，{selected_period.month}月銷貨合計將增至 US${total_usd:,.2f}。")

    # Prepare standard detail rows
    teable_shipped_detail = pd.DataFrame({
        "來源": "Teable",
        "客戶": _safe_series(customer_col_local)[shipped_mask].map(_clean_customer),
        "工廠": _safe_series(factory_col_local)[shipped_mask].map(_clean_factory),
        "PO#": _safe_series(po_col_local)[shipped_mask].astype(str),
        "P/N": _safe_series(pn_col_local)[shipped_mask].astype(str),
        "QTY": _safe_series(qty_col_local)[shipped_mask].astype(str),
        "WIP": wip_series[shipped_mask].astype(str),
        "日期": shipped_dates[shipped_mask],
        "金額": sales_value_series[shipped_mask],
    })
    teable_forecast_detail = pd.DataFrame({
        "來源": "Teable",
        "客戶": _safe_series(customer_col_local)[forecast_mask].map(_clean_customer),
        "工廠": _safe_series(factory_col_local)[forecast_mask].map(_clean_factory),
        "PO#": _safe_series(po_col_local)[forecast_mask].astype(str),
        "P/N": _safe_series(pn_col_local)[forecast_mask].astype(str),
        "QTY": _safe_series(qty_col_local)[forecast_mask].astype(str),
        "WIP": wip_series[forecast_mask].astype(str),
        "日期": forecast_dates[forecast_mask],
        "金額": sales_value_series[forecast_mask],
    })
    manual_shipped_detail = _manual_rows_to_standard(manual_shipped_df)
    manual_forecast_detail = _manual_rows_to_standard(manual_forecast_df)

    combined_month_total = pd.concat([teable_shipped_detail, teable_forecast_detail, manual_shipped_detail, manual_forecast_detail], ignore_index=True)

    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown(f"#### 🏭 依工廠別統計（{selected_period.month}月銷貨）")
        if not combined_month_total.empty:
            grp = combined_month_total.groupby("工廠", dropna=False).agg(訂單數=("金額", "count"), 銷貨金額_USD=("金額", "sum")).reset_index()
            grp["工廠"] = grp["工廠"].replace({"": "(空白)"})
            grp = grp.sort_values("銷貨金額_USD", ascending=False)
            total_row = pd.DataFrame([{"工廠": "合計", "訂單數": int(grp["訂單數"].sum()), "銷貨金額_USD": float(grp["銷貨金額_USD"].sum())}])
            grp = pd.concat([grp, total_row], ignore_index=True)
            grp["銷貨金額(USD)"] = grp["銷貨金額_USD"].map(_fmt_usd)
            st.dataframe(grp[["工廠", "訂單數", "銷貨金額(USD)"]], use_container_width=True, hide_index=True, height=320)
        else:
            st.info("本月無工廠銷貨資料。")

    with col_right:
        st.markdown(f"#### 👥 依客戶別統計（{selected_period.month}月銷貨）")
        if not combined_month_total.empty:
            grp2 = combined_month_total.groupby("客戶", dropna=False).agg(訂單數=("金額", "count"), 銷貨金額_USD=("金額", "sum")).reset_index()
            grp2["客戶"] = grp2["客戶"].replace({"": "(空白)"})
            grp2 = grp2.sort_values("銷貨金額_USD", ascending=False)
            grp2["銷貨金額(USD)"] = grp2["銷貨金額_USD"].map(_fmt_usd)
            st.dataframe(grp2[["客戶", "訂單數", "銷貨金額(USD)"]], use_container_width=True, hide_index=True, height=320)
        else:
            st.info("本月無客戶銷貨資料。")

    st.markdown(f"#### 📆 依日期統計（{selected_period.month}月銷貨）")
    if not combined_month_total.empty:
        daily_df = combined_month_total.dropna(subset=["日期"]).copy()
        if not daily_df.empty:
            daily_df = (
                daily_df.groupby(pd.to_datetime(daily_df["日期"]).dt.normalize())
                .agg(訂單數=("金額", "count"), 銷貨金額_USD=("金額", "sum"))
                .reset_index()
                .rename(columns={"日期": "日期"})
                .sort_values("日期")
            )
            daily_df["台幣估算(萬)"] = daily_df["銷貨金額_USD"].map(_usd_to_ntd_10k)
            daily_chart = daily_df.copy()
            daily_chart["日期"] = pd.to_datetime(daily_chart["日期"]).dt.strftime("%m/%d")
            daily_chart["累計出貨(USD)"] = daily_df["銷貨金額_USD"].cumsum()
            st.line_chart(daily_chart.set_index("日期")[["累計出貨(USD)"]], height=240)
            daily_fmt = daily_df.copy()
            daily_fmt["日期"] = pd.to_datetime(daily_fmt["日期"]).dt.strftime("%Y-%m-%d")
            daily_fmt["銷貨金額(USD)"] = daily_fmt["銷貨金額_USD"].map(_fmt_usd)
            daily_fmt["台幣估算(萬)"] = daily_fmt["台幣估算(萬)"].map(lambda x: f"{x:,.2f}")
            st.dataframe(daily_fmt[["日期", "訂單數", "銷貨金額(USD)", "台幣估算(萬)"]], use_container_width=True, hide_index=True, height=260)
        else:
            st.info("本月有銷貨金額，但沒有可解析的日期。")
    else:
        st.info("本月尚無銷貨資料。")

    st.markdown(f"#### ✅ 已出貨明細（SHIPMENT，{selected_period.month}月）")
    shipped_view = pd.concat([teable_shipped_detail, manual_shipped_detail], ignore_index=True)
    if not shipped_view.empty:
        shipped_view = shipped_view.copy()
        shipped_view["日期"] = pd.to_datetime(shipped_view["日期"]).dt.strftime("%Y-%m-%d")
        shipped_view["銷貨金額(USD)"] = shipped_view["金額"].map(_fmt_usd)
        shipped_view = shipped_view.sort_values("日期", ascending=False, na_position="last")
        st.dataframe(shipped_view[["來源", "客戶", "PO#", "P/N", "QTY", "工廠", "日期", "銷貨金額(USD)"]], use_container_width=True, hide_index=True, height=320)
    else:
        st.info("本月尚無已出貨（SHIPMENT）資料。")

    st.markdown(f"#### 🔜 預計出貨明細（未 SHIPMENT，{selected_period.month}月）")
    forecast_view = pd.concat([teable_forecast_detail, manual_forecast_detail], ignore_index=True)
    if not forecast_view.empty:
        forecast_view = forecast_view.copy()
        forecast_view["日期"] = pd.to_datetime(forecast_view["日期"]).dt.strftime("%Y-%m-%d")
        forecast_view["銷貨金額(USD)"] = forecast_view["金額"].map(_fmt_usd)
        forecast_view = forecast_view.sort_values(["日期", "WIP"], ascending=[True, True], na_position="last")
        st.dataframe(forecast_view[["來源", "客戶", "PO#", "P/N", "QTY", "WIP", "工廠", "日期", "銷貨金額(USD)"]], use_container_width=True, hide_index=True, height=260)
    else:
        st.info("本月無預計出貨資料。")

    st.markdown("#### 📈 近 12 個月月銷貨趨勢")
    teable_trend = pd.concat([
        pd.DataFrame({"月份": shipped_periods[shipped_mask | ((is_shipment) & (~is_cancelled))], "金額(USD)": sales_value_series[shipped_mask | ((is_shipment) & (~is_cancelled))], "類型": "已出貨"}),
        pd.DataFrame({"月份": forecast_periods[(~is_shipment) & (~is_cancelled)], "金額(USD)": sales_value_series[(~is_shipment) & (~is_cancelled)], "類型": "月銷貨總計"}),
    ], ignore_index=True)
    teable_trend = teable_trend.dropna(subset=["月份"])
    if not manual_history.empty:
        mh = manual_history.copy()
        mh = mh[mh["月份"].astype(str).str.match(r"^\d{4}-\d{2}$", na=False)].copy()
        if not mh.empty:
            mh["月份"] = mh["月份"].astype(str).apply(lambda m: pd.Period(m, freq="M"))
            manual_trend = pd.DataFrame({
                "月份": mh["月份"],
                "金額(USD)": parse_numeric_series(mh["金額(USD)"]),
                "類型": mh["類型"].replace({"預計出貨": "月銷貨總計", "已出貨": "已出貨", "接單": "接單"}),
            })
            teable_trend = pd.concat([teable_trend, manual_trend[manual_trend["類型"].isin(["已出貨", "月銷貨總計"])]], ignore_index=True)

    if not teable_trend.empty:
        shipped_trend = teable_trend[teable_trend["類型"].eq("已出貨")].groupby("月份")["金額(USD)"].sum().rename("已出貨金額_USD")
        total_trend = teable_trend.groupby("月份")["金額(USD)"].sum().rename("月銷貨總計_USD")
        trend = pd.concat([shipped_trend, total_trend], axis=1).fillna(0.0).reset_index().sort_values("月份").tail(12)
        trend_chart = trend.copy()
        trend_chart["月份"] = trend_chart["月份"].astype(str)
        st.bar_chart(trend_chart.set_index("月份")[["已出貨金額_USD", "月銷貨總計_USD"]], height=260)
        trend_fmt = trend.copy()
        trend_fmt["月份"] = trend_fmt["月份"].astype(str)
        trend_fmt["已出貨金額(USD)"] = trend_fmt["已出貨金額_USD"].map(_fmt_usd)
        trend_fmt["月銷貨總計(USD)"] = trend_fmt["月銷貨總計_USD"].map(_fmt_usd)
        st.dataframe(trend_fmt[["月份", "已出貨金額(USD)", "月銷貨總計(USD)"]], use_container_width=True, hide_index=True)
    else:
        st.info("沒有足夠資料可顯示近 12 個月趨勢。")


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
