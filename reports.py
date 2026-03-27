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
    核心規則：
    1) 接單金額：依接單日期落在所選月份統計
    2) 已確認出貨：WIP=SHIPMENT，優先用「出貨日期」，沒有則回退「Ship date」
    3) 預計本月出貨：WIP!=SHIPMENT，但 Ship date / 出貨日期 落在該月
    4) 金額：優先抓「銷貨金額」，若空白或 0，回補「接單金額」
    5) 額外把附圖需要的月摘要、依工廠統計、依日期統計都加進表格
    """
    import datetime as _dt

    st.subheader("📊 業績明細表")

    if source_df is None or source_df.empty:
        st.warning("Teable 主表目前沒有資料。")
        return

    st.caption("資料來源：Teable 主表即時欄位（全客戶）")

    FX_NTD_PER_USD = 31.5

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

    customer_col_local = _pick_col(CUSTOMER_CANDIDATES, ["Customer"])
    factory_col_local = _pick_col(FACTORY_CANDIDATES)
    po_col_local = _pick_col(PO_CANDIDATES)
    pn_col_local = _pick_col(PART_CANDIDATES)
    qty_col_local = _pick_col(QTY_CANDIDATES, ["Order QTY (PCS)", "Order QTY", "QTY"])
    wip_col_local = _pick_col(WIP_CANDIDATES)

    order_date_col = _pick_col(ORDER_DATE_CANDIDATES, ["客戶下單日期", "工廠下單日期", "下單日期", "接單日期"])
    actual_ship_col = _pick_col(["出貨日期", "出貨日期_排序"], SHIP_DATE_CANDIDATES)
    planned_ship_col = _pick_col(["Ship date", "Ship Date", "預計出貨日", "客戶交期"], SHIP_DATE_CANDIDATES)
    effective_ship_col = actual_ship_col or planned_ship_col

    sales_amt_col = find_amount_column(
        source_df,
        ["銷貨金額", "出貨金額", "出貨發票金額", "Invoice Amount", "Invoice Total", "Invoice", "INVOICE", "發票"]
        + AMOUNT_SHIP_CANDIDATES
    )
    order_amt_col = find_amount_column(
        source_df,
        ["接單金額", "接單總金額", "Order Amount", "Order Total", "客戶金額", "Total Amount", "Amount"]
        + AMOUNT_ORDER_CANDIDATES
    )

    order_dates = _parse_date_from_col(order_date_col)
    actual_ship_dates = _parse_date_from_col(actual_ship_col)
    planned_ship_dates = _parse_date_from_col(planned_ship_col)

    effective_ship_dates = actual_ship_dates.copy()
    missing_ship_date = effective_ship_dates.isna()
    effective_ship_dates.loc[missing_ship_date] = planned_ship_dates.loc[missing_ship_date]

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

    ship_periods = effective_ship_dates.dt.to_period("M")
    order_periods = order_dates.dt.to_period("M")
    all_periods = sorted(set(ship_periods.dropna().tolist()) | set(order_periods.dropna().tolist()), reverse=True)

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

    ship_month_mask = (ship_periods == selected_period).fillna(False)
    order_month_mask = (order_periods == selected_period).fillna(False)

    shipped_mask = is_shipment & ship_month_mask & (~is_cancelled)
    forecast_mask = (~is_shipment) & ship_month_mask & (~is_cancelled)
    order_mask = order_month_mask & (~is_cancelled)
    month_total_mask = shipped_mask | forecast_mask

    shipped_usd = float(sales_value_series[shipped_mask].sum())
    forecast_usd = float(sales_value_series[forecast_mask].sum())
    order_usd = float(order_value_series[order_mask].sum())
    total_usd = shipped_usd + forecast_usd

    shipped_count = int(shipped_mask.sum())
    forecast_count = int(forecast_mask.sum())
    order_count = int(order_mask.sum())

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

    qa_mask = forecast_mask & wip_series.isin(["QA", "QC", "FQC", "INSPECTION"])
    if qa_mask.any():
        qa_sum = float(sales_value_series[qa_mask].sum())
        qa_ship_dates = effective_ship_dates[qa_mask].dropna().sort_values()
        qa_date_txt = qa_ship_dates.iloc[0].strftime("%-m/%-d") if len(qa_ship_dates) else ""
        qa_names = ""
        if customer_col_local:
            names = [_clean_customer(n) for n in get_series_by_col(source_df, customer_col_local)[qa_mask].dropna().unique()]
            qa_names = "、".join([n for n in names if n])
        st.info(
            f"{qa_names} 這筆明天（{qa_date_txt}）出貨，屆時 WIP 更新為 SHIPMENT 後，"
            f"{selected_period.month}月份總銷貨金額即為 US${total_usd:,.2f}。"
        )

    st.markdown("#### 📋 月摘要表")
    summary_df = pd.DataFrame([
        {"分類": "接單金額", "訂單數": order_count, "金額(USD)": order_usd, "換算台幣(萬)": _usd_to_ntd_10k(order_usd)},
        {"分類": "已確認出貨 (SHIPMENT)", "訂單數": shipped_count, "金額(USD)": shipped_usd, "換算台幣(萬)": _usd_to_ntd_10k(shipped_usd)},
        {"分類": "預計本月出貨", "訂單數": forecast_count, "金額(USD)": forecast_usd, "換算台幣(萬)": _usd_to_ntd_10k(forecast_usd)},
        {"分類": f"{selected_period.month}月份銷貨合計", "訂單數": shipped_count + forecast_count, "金額(USD)": total_usd, "換算台幣(萬)": _usd_to_ntd_10k(total_usd)},
    ])
    summary_df_fmt = summary_df.copy()
    summary_df_fmt["金額(USD)"] = summary_df_fmt["金額(USD)"].map(_fmt_usd)
    summary_df_fmt["換算台幣(萬)"] = summary_df_fmt["換算台幣(萬)"].map(lambda x: f"{x:,.2f}")
    st.dataframe(summary_df_fmt, use_container_width=True, hide_index=True)

    with st.expander("Debug：業績明細表欄位偵測", expanded=False):
        dbg = pd.DataFrame([
            {"項目": "客戶", "欄位": customer_col_local or ""},
            {"項目": "工廠", "欄位": factory_col_local or ""},
            {"項目": "PO#", "欄位": po_col_local or ""},
            {"項目": "P/N", "欄位": pn_col_local or ""},
            {"項目": "QTY", "欄位": qty_col_local or ""},
            {"項目": "接單日期", "欄位": order_date_col or ""},
            {"項目": "實際出貨日期", "欄位": actual_ship_col or ""},
            {"項目": "預計出貨日期", "欄位": planned_ship_col or ""},
            {"項目": "有效出貨日期", "欄位": effective_ship_col or ""},
            {"項目": "WIP", "欄位": wip_col_local or ""},
            {"項目": "銷貨金額", "欄位": sales_amt_col or ""},
            {"項目": "接單金額", "欄位": order_amt_col or ""},
        ])
        st.dataframe(dbg, use_container_width=True, hide_index=True)

    st.markdown("---")

    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown(f"#### 🏭 依工廠別統計（{selected_period.month}月銷貨）")
        if month_total_mask.any() and factory_col_local:
            fdf = pd.DataFrame({
                "工廠": get_series_by_col(source_df, factory_col_local)[month_total_mask].apply(_clean_factory),
                "金額": sales_value_series[month_total_mask],
            })
            grp = (
                fdf.groupby("工廠")
                .agg(訂單數=("金額", "count"), 銷貨金額_USD=("金額", "sum"))
                .reset_index()
                .sort_values("銷貨金額_USD", ascending=False)
            )
            total_row = pd.DataFrame([{
                "工廠": "合計",
                "訂單數": int(grp["訂單數"].sum()) if not grp.empty else 0,
                "銷貨金額_USD": float(grp["銷貨金額_USD"].sum()) if not grp.empty else 0.0,
            }])
            grp = pd.concat([grp, total_row], ignore_index=True)
            grp["銷貨金額(USD)"] = grp["銷貨金額_USD"].map(_fmt_usd)
            st.dataframe(grp[["工廠", "訂單數", "銷貨金額(USD)"]], use_container_width=True, hide_index=True, height=320)
        else:
            st.info("本月無工廠銷貨資料。")

    with col_right:
        st.markdown(f"#### 👥 依客戶別統計（{selected_period.month}月銷貨）")
        if month_total_mask.any() and customer_col_local:
            cdf = pd.DataFrame({
                "客戶": get_series_by_col(source_df, customer_col_local)[month_total_mask].apply(_clean_customer),
                "金額": sales_value_series[month_total_mask],
            })
            grp2 = (
                cdf.groupby("客戶")
                .agg(訂單數=("金額", "count"), 銷貨金額_USD=("金額", "sum"))
                .reset_index()
                .sort_values("銷貨金額_USD", ascending=False)
            )
            grp2["銷貨金額(USD)"] = grp2["銷貨金額_USD"].map(_fmt_usd)
            st.dataframe(grp2[["客戶", "訂單數", "銷貨金額(USD)"]], use_container_width=True, hide_index=True, height=320)
        else:
            st.info("本月無客戶銷貨資料。")

    st.markdown(f"#### 📆 依日期統計（{selected_period.month}月銷貨）")
    if month_total_mask.any():
        daily_df = pd.DataFrame({
            "日期": effective_ship_dates[month_total_mask],
            "金額(USD)": sales_value_series[month_total_mask],
        }).dropna(subset=["日期"])
        if not daily_df.empty:
            daily_df = (
                daily_df.groupby(daily_df["日期"].dt.normalize())
                .agg(訂單數=("金額(USD)", "count"), 銷貨金額_USD=("金額(USD)", "sum"))
                .reset_index()
                .sort_values("日期")
            )
            daily_df["台幣估算(萬)"] = daily_df["銷貨金額_USD"].map(_usd_to_ntd_10k)

            cumulative = daily_df.copy()
            cumulative["累計出貨(USD)"] = cumulative["銷貨金額_USD"].cumsum()
            chart_df = cumulative.copy()
            chart_df["日期"] = pd.to_datetime(chart_df["日期"]).dt.strftime("%m/%d")
            st.line_chart(chart_df.set_index("日期")[["累計出貨(USD)"]], height=240)

            daily_df_fmt = daily_df.copy()
            daily_df_fmt["日期"] = pd.to_datetime(daily_df_fmt["日期"]).dt.strftime("%Y-%m-%d")
            daily_df_fmt["銷貨金額(USD)"] = daily_df_fmt["銷貨金額_USD"].map(_fmt_usd)
            daily_df_fmt["台幣估算(萬)"] = daily_df_fmt["台幣估算(萬)"].map(lambda x: f"{x:,.2f}")
            st.dataframe(daily_df_fmt[["日期", "訂單數", "銷貨金額(USD)", "台幣估算(萬)"]], use_container_width=True, hide_index=True, height=260)
        else:
            st.info("本月有銷貨金額，但沒有可解析的日期。")
    else:
        st.info("本月尚無銷貨資料。")

    st.markdown(f"#### ✅ 已出貨明細（SHIPMENT，{selected_period.month}月）")
    if shipped_mask.any():
        view = pd.DataFrame(index=source_df.index[shipped_mask])
        if customer_col_local:
            view["客戶"] = get_series_by_col(source_df, customer_col_local)[shipped_mask].apply(_clean_customer)
        if po_col_local:
            view["PO#"] = get_series_by_col(source_df, po_col_local)[shipped_mask]
        if pn_col_local:
            view["P/N"] = get_series_by_col(source_df, pn_col_local)[shipped_mask]
        if qty_col_local:
            view["QTY"] = get_series_by_col(source_df, qty_col_local)[shipped_mask]
        if factory_col_local:
            view["工廠"] = get_series_by_col(source_df, factory_col_local)[shipped_mask].apply(_clean_factory)
        view["日期"] = effective_ship_dates[shipped_mask].dt.strftime("%Y-%m-%d")
        view["銷貨金額(USD)"] = sales_value_series[shipped_mask].map(_fmt_usd)
        view = view.sort_values("日期", ascending=False, na_position="last")
        st.dataframe(view, use_container_width=True, hide_index=True, height=320)
    else:
        st.info("本月尚無已出貨（SHIPMENT）資料。")

    st.markdown(f"#### 🔜 預計出貨明細（未 SHIPMENT，{selected_period.month}月）")
    if forecast_mask.any():
        view2 = pd.DataFrame(index=source_df.index[forecast_mask])
        if customer_col_local:
            view2["客戶"] = get_series_by_col(source_df, customer_col_local)[forecast_mask].apply(_clean_customer)
        if po_col_local:
            view2["PO#"] = get_series_by_col(source_df, po_col_local)[forecast_mask]
        if pn_col_local:
            view2["P/N"] = get_series_by_col(source_df, pn_col_local)[forecast_mask]
        if qty_col_local:
            view2["QTY"] = get_series_by_col(source_df, qty_col_local)[forecast_mask]
        view2["WIP"] = wip_series[forecast_mask]
        if factory_col_local:
            view2["工廠"] = get_series_by_col(source_df, factory_col_local)[forecast_mask].apply(_clean_factory)
        view2["日期"] = effective_ship_dates[forecast_mask].dt.strftime("%Y-%m-%d")
        view2["銷貨金額(USD)"] = sales_value_series[forecast_mask].map(_fmt_usd)
        view2 = view2.sort_values(["日期", "WIP"], ascending=[True, True], na_position="last")
        st.dataframe(view2, use_container_width=True, hide_index=True, height=260)
    else:
        st.info("本月無預計出貨資料。")

    st.markdown("#### 📈 近 12 個月月銷貨趨勢")
    monthly_base = pd.DataFrame({
        "月份": ship_periods[~is_cancelled],
        "金額(USD)": sales_value_series[~is_cancelled],
        "已出貨": is_shipment[~is_cancelled],
    }).dropna(subset=["月份"])
    if not monthly_base.empty:
        shipped_trend = (
            monthly_base[monthly_base["已出貨"]]
            .groupby("月份")["金額(USD)"]
            .sum()
            .rename("已出貨金額_USD")
        )
        total_trend = (
            monthly_base.groupby("月份")["金額(USD)"]
            .sum()
            .rename("月銷貨總計_USD")
        )
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
