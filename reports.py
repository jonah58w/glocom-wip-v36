# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
import re
import pandas as pd
import streamlit as st

# ================================
# FIELD CANDIDATES
# ================================
PO_CANDIDATES = ["PO#", "PO", "P/O", "訂單編號", "訂單號", "訂單號碼", "工單", "工單號", "單號"]
CUSTOMER_CANDIDATES = ["Customer", "客戶", "客戶名稱"]
PART_CANDIDATES = [
    "Part No", "Part No.", "P/N", "客戶料號", "Cust. P / N", "LS P/N",
    "料號", "品號", "成品料號", "產品料號"
]
QTY_CANDIDATES = [
    "Qty", "Order Q'TY (PCS)", "Order Q'TY\n (PCS)", "Order Q'TY (PCS)", "訂購量 (PCS)",
    "訂購量", "Q'TY", "數量", "PCS", "訂單量", "生產數量", "投產數"
]
FACTORY_CANDIDATES = ["Factory", "工廠", "廠編"]
WIP_CANDIDATES = ["WIP", "WIP Stage", "進度", "製程", "工序", "目前站別", "生產進度"]
FACTORY_DUE_CANDIDATES = [
    "Factory Due Date", "工廠交期", "交貨日期", "Required Ship date",
    "confrimed DD", "交期", "預交日", "預定交期", "交貨期"
]
SHIP_DATE_CANDIDATES = ["Ship Date", "Ship date", "出貨日期", "交貨日期", "Required Ship date", "confrimed DD"]
REMARK_CANDIDATES = ["Remark", "備註", "情況", "備註說明", "Note", "說明", "異常備註"]
ORDER_DATE_CANDIDATES = ["客戶下單日期", "工廠下單日期", "下單日期", "Order Date", "PO Date", "Date", "訂單日期", "接單日期"]
AMOUNT_ORDER_CANDIDATES = [
    "接單金額", "接單總金額", "Order Amount", "Order amount", "Order Total",
    "客戶金額", "Sales Amount", "Quote Total", "Total Amount", "Amount"
]
AMOUNT_SHIP_CANDIDATES = [
    "銷貨金額", "出貨金額", "出貨總金額", "Shipment Amount", "Ship Amount", "Shipping Amount",
    "Invoice Amount", "Invoice Total", "出貨發票金額", "Invoice", "INVOICE"
]
ACTUAL_SHIP_DATE_CANDIDATES = ["出貨日期_排序", "出貨日期", "Actual Ship Date", "Actual ship date"]
PLANNED_SHIP_DATE_CANDIDATES = ["Ship date", "Ship Date", "Required Ship date", "confrimed DD"]

# PO Cancelled 相關關鍵字
CANCELLED_KEYWORDS = [
    "PO CANCELLED", "PO CANCELED", "CANCELLATION", "CANCELLED", "CANCELED", "CANCEL",
]

MANUAL_HISTORY_PATH = Path("sales_manual_history.csv")

# ── 補登 CSV 欄位對照表（新格式中文 ↔ 內部英文標準名）──────────────────────
MANUAL_COL_MAP = {
    "月份": "month", "日期": "date", "類型": "type",
    "客戶": "customer", "工廠": "factory",
    "PO#": "po", "P/N": "pn", "QTY": "qty", "WIP": "wip",
    "金額(USD)": "amount", "備註": "note",
    # 舊格式（英文直接對應）
    "date": "date", "type": "type", "customer": "customer", "factory": "factory",
    "po": "po", "pn": "pn", "qty": "qty", "wip": "wip", "amount": "amount", "note": "note",
}
MANUAL_REQUIRED_COLS = ["date", "type", "customer", "factory", "po", "pn", "qty", "wip", "amount", "note"]


# ================================
# UTILITIES
# ================================

def _norm(text: str) -> str:
    s = str(text or "")
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def make_unique_columns(columns):
    seen = {}
    out = []
    for c in list(columns):
        name = str(c)
        n = seen.get(name, 0)
        out.append(name if n == 0 else f"{name}_{n + 1}")
        seen[name] = n + 1
    return out


def col_candidates(*names):
    return [str(x).strip() for x in names if str(x).strip()]


def get_series_by_col(df: pd.DataFrame, col_name: str | None):
    if df is None or df.empty or not col_name or col_name not in df.columns:
        return None
    obj = df[col_name]
    return obj.iloc[:, 0] if isinstance(obj, pd.DataFrame) else obj


def find_col(df: pd.DataFrame, candidates):
    if df is None or df.empty:
        return None
    norm_map = {_norm(c): c for c in df.columns}
    for cand in candidates:
        if _norm(cand) in norm_map:
            return norm_map[_norm(cand)]
    for cand in candidates:
        n = _norm(cand)
        for col in df.columns:
            if n and n in _norm(col):
                return col
    return None


def parse_amount_series(series: pd.Series | None) -> pd.Series:
    if series is None:
        return pd.Series(dtype="float64")
    s = series.astype(str).fillna("").str.strip()
    s = s.str.replace(r"[^0-9.\-]", "", regex=True)
    s = s.replace({"": None, ".": None, "-": None})
    return pd.to_numeric(s, errors="coerce")


def parse_mixed_date_series(series: pd.Series | None) -> pd.Series:
    """
    解析混合格式日期，優先順序：
      1. Excel serial number
      2. YYYY-MM-DD（最可靠，避免歧義）
      3. "Mar. 27, 26" 等月份縮寫格式（兩位年份）
      4. YY-MM-DD（兩位年份，26 → 2026）
      5. pandas fallback
    """
    if series is None:
        return pd.Series(dtype="datetime64[ns]")

    s = series.astype(str).fillna("").str.strip()
    s = s.replace({"": None, "nan": None, "NaT": None, "None": None})
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")

    # 1. Excel serial
    nums = pd.to_numeric(s, errors="coerce")
    mask_num = nums.notna() & (nums > 20000) & (nums < 80000)
    if mask_num.any():
        out.loc[mask_num] = pd.to_datetime(
            nums.loc[mask_num], unit="D", origin="1899-12-30", errors="coerce"
        )

    # 2. YYYY-MM-DD
    rem = out.isna() & s.notna()
    if rem.any():
        out.loc[rem] = pd.to_datetime(s.loc[rem], format="%Y-%m-%d", errors="coerce")

    # 3. "Mar. 27, 26" / "Mar 27, 26"
    rem = out.isna() & s.notna()
    if rem.any():
        cleaned = (
            s.loc[rem]
            .str.replace(".", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )
        out.loc[rem] = pd.to_datetime(cleaned, format="%b %d %y", errors="coerce")

    # 4. YY-MM-DD（e.g. 26-03-27 → 2026-03-27）
    rem = out.isna() & s.notna()
    if rem.any():
        out.loc[rem] = pd.to_datetime(s.loc[rem], format="%y-%m-%d", errors="coerce")

    # 5. fallback
    rem = out.isna() & s.notna()
    if rem.any():
        out.loc[rem] = pd.to_datetime(s.loc[rem], errors="coerce")

    return out


def _is_cancelled(wip_series: pd.Series) -> pd.Series:
    """WIP 含有取消相關關鍵字 → True"""
    upper = wip_series.astype(str).str.strip().str.upper()
    mask = pd.Series(False, index=wip_series.index)
    for kw in CANCELLED_KEYWORDS:
        mask = mask | upper.str.contains(kw.upper(), regex=False, na=False)
    return mask


# ================================
# VIEW SPECS
# ================================

SANDY_NEW_ORDER_SPECS = [
    ("客戶下單日期", ORDER_DATE_CANDIDATES),
    ("工廠下單日期", col_candidates("工廠下單日期")),
    ("客戶", CUSTOMER_CANDIDATES + ["Customer"]),
    ("PO#", PO_CANDIDATES),
    ("P/N", PART_CANDIDATES),
    ("Order Q'TY (PCS)", QTY_CANDIDATES + ["Order QTY (PCS)"]),
    ("Dock", col_candidates("Dock")),
    ("Ship date", PLANNED_SHIP_DATE_CANDIDATES),
    ("WIP", WIP_CANDIDATES),
    ("工廠交期", FACTORY_DUE_CANDIDATES),
    ("交期 (更改)", col_candidates("交期 (更改)", "交期\n (更改)")),
    ("出貨日期", ACTUAL_SHIP_DATE_CANDIDATES),
    ("工廠", FACTORY_CANDIDATES),
    ("工廠提醒事項", col_candidates("工廠提醒事項")),
    ("併貨日期 (限內部使用)", col_candidates("併貨日期 (限內部使用)", "併貨日期\n (限內部使用)")),
    ("情況", REMARK_CANDIDATES),
    ("客戶要求注意事項", col_candidates("客戶要求注意事項")),
    ("Ship to", col_candidates("Ship to")),
    ("Ship via", col_candidates("Ship via")),
    ("箱數", col_candidates("箱數", "CTNS", "CTN")),
    ("重量", col_candidates("重量", "Weight", "KGs")),
    ("重貨優惠", col_candidates("重貨優惠", "重貨\n 優惠")),
    ("Pricing & Qty issue", col_candidates("Pricing & Qty issue", "Pricing\n &\n Qty issue")),
    ("T/T", col_candidates("T/T")),
    ("工廠出貨事項", col_candidates("工廠出貨事項", "工廠出貨注意事項")),
    ("新/舊料號", col_candidates("新/舊料號", "新/舊\n料號")),
    ("板層", col_candidates("板層", "板\n層")),
]

SANDY_INTERNAL_WIP_SPECS = [
    ("Customer", CUSTOMER_CANDIDATES + ["Customer"]),
    ("PO#", PO_CANDIDATES),
    ("P/N", PART_CANDIDATES),
    ("Q'TY (PCS)", QTY_CANDIDATES + ["Order QTY (PCS)"]),
    ("Dock", col_candidates("Dock")),
    ("Ship date", PLANNED_SHIP_DATE_CANDIDATES),
    ("WIP", WIP_CANDIDATES),
    ("出貨狀況 (限內部使用)", col_candidates("出貨狀況 (限內部使用)")),
    ("進度狀況", col_candidates("進度狀況")),
    ("工廠交期", FACTORY_DUE_CANDIDATES),
    ("交期 (更改)", col_candidates("交期 (更改)", "交期\n (更改)")),
    ("出貨日期", ACTUAL_SHIP_DATE_CANDIDATES),
    ("工廠", FACTORY_CANDIDATES),
    ("工廠提醒事項", col_candidates("工廠提醒事項")),
    ("併貨日期 (限內部使用)", col_candidates("併貨日期 (限內部使用)", "併貨日期\n (限內部使用)")),
    ("客戶要求注意事項", col_candidates("客戶要求注意事項")),
    ("Ship to", col_candidates("Ship to")),
    ("Ship via", col_candidates("Ship via")),
    ("CTN", col_candidates("CTN", "CTNS", "箱數")),
    ("KGs", col_candidates("KGs", "重量")),
    ("Pricing & Qty issue", col_candidates("Pricing & Qty issue", "Pricing\n &\n Qty issue")),
    ("T/T", col_candidates("T/T")),
    ("Note", REMARK_CANDIDATES),
]

SANDY_SALES_SPECS = [
    ("客戶", CUSTOMER_CANDIDATES),
    ("PO#", PO_CANDIDATES),
    ("P/N", PART_CANDIDATES),
    ("Q'TY (PCS)", QTY_CANDIDATES),
    ("工廠", FACTORY_CANDIDATES),
    ("出貨日期", ACTUAL_SHIP_DATE_CANDIDATES),
    ("Ship date", PLANNED_SHIP_DATE_CANDIDATES),
    ("WIP", WIP_CANDIDATES),
    ("接單金額", AMOUNT_ORDER_CANDIDATES),
    ("銷貨金額", AMOUNT_SHIP_CANDIDATES),
    ("Note", REMARK_CANDIDATES),
]


# ================================
# BUILD VIEW DF
# ================================

def build_teable_view_df(source_df: pd.DataFrame, specs):
    mapping = {}
    out = pd.DataFrame(index=source_df.index)
    for display_name, candidates in specs:
        src = find_col(source_df, candidates)
        if src:
            out[display_name] = get_series_by_col(source_df, src)
            mapping[display_name] = src
        else:
            out[display_name] = ""
            mapping[display_name] = None
    out.columns = make_unique_columns(out.columns)
    return out, mapping


# ================================
# CORE FILTER
# ================================

def build_subset_mask(source_df: pd.DataFrame, subset_mode: str) -> pd.Series:
    """
    new_order_today  → 客戶下單日期 == 今天  OR  工廠下單日期 == 今天

    unshipped        → WIP != SHIPMENT
                       AND WIP 不含 Cancelled/Cancellation 關鍵字
                       AND 年份 == 當年（用客戶下單日期 > 工廠下單日期 > Ship date 判斷）
                       （三者都無日期 → 保留，避免漏單）

    shipment_only    → WIP == SHIPMENT  AND  出貨日期 == 今天
    """
    idx = source_df.index
    today = pd.Timestamp.today().normalize()
    current_year = today.year

    # ── WIP ──────────────────────────────────────────────────────────────
    wip_col = find_col(source_df, WIP_CANDIDATES)
    wip_raw = (
        get_series_by_col(source_df, wip_col).astype(str).str.strip()
        if wip_col else pd.Series("", index=idx)
    )
    wip_upper = wip_raw.str.upper()

    # ════════════════════════════════════════════════════════════════════
    # 新訂單 WIP
    # ════════════════════════════════════════════════════════════════════
    if subset_mode == "new_order_today":
        cust_order_col = find_col(source_df, ["客戶下單日期"])
        fact_order_col = find_col(source_df, ["工廠下單日期"])

        cust_dates = (
            parse_mixed_date_series(get_series_by_col(source_df, cust_order_col))
            if cust_order_col else pd.Series(pd.NaT, index=idx)
        )
        fact_dates = (
            parse_mixed_date_series(get_series_by_col(source_df, fact_order_col))
            if fact_order_col else pd.Series(pd.NaT, index=idx)
        )

        cust_today = cust_dates.dt.normalize().eq(today).fillna(False)
        fact_today = fact_dates.dt.normalize().eq(today).fillna(False)
        return cust_today | fact_today

    # ════════════════════════════════════════════════════════════════════
    # Sandy 內部 WIP
    # ════════════════════════════════════════════════════════════════════
    if subset_mode == "unshipped":
        # 條件1：不是 SHIPMENT
        not_shipment = ~wip_upper.eq("SHIPMENT")
        # 條件2：不含取消關鍵字
        not_cancelled = ~_is_cancelled(wip_raw)

        # 條件3：限當年度
        cust_order_col = find_col(source_df, ["客戶下單日期"])
        fact_order_col = find_col(source_df, ["工廠下單日期"])
        ship_col       = find_col(source_df, PLANNED_SHIP_DATE_CANDIDATES)

        cust_dates = (
            parse_mixed_date_series(get_series_by_col(source_df, cust_order_col))
            if cust_order_col else pd.Series(pd.NaT, index=idx)
        )
        fact_dates = (
            parse_mixed_date_series(get_series_by_col(source_df, fact_order_col))
            if fact_order_col else pd.Series(pd.NaT, index=idx)
        )
        ship_dates = (
            parse_mixed_date_series(get_series_by_col(source_df, ship_col))
            if ship_col else pd.Series(pd.NaT, index=idx)
        )

        # 依優先序取年份：客戶下單 → 工廠下單 → Ship date
        year_series = (
            cust_dates.dt.year
            .where(cust_dates.notna(), fact_dates.dt.year)
            .where(cust_dates.notna() | fact_dates.notna(), ship_dates.dt.year)
        )

        # 年份等於當年，或三欄都沒有日期（保留）
        year_ok = year_series.isna() | year_series.eq(current_year)

        return not_shipment & not_cancelled & year_ok

    # ════════════════════════════════════════════════════════════════════
    # Sandy 銷貨底
    # ════════════════════════════════════════════════════════════════════
    if subset_mode == "shipment_only":
        is_shipment = wip_upper.eq("SHIPMENT")

        actual_col = find_col(source_df, ACTUAL_SHIP_DATE_CANDIDATES)
        actual_dates = (
            parse_mixed_date_series(get_series_by_col(source_df, actual_col))
            if actual_col else pd.Series(pd.NaT, index=idx)
        )
        shipped_today = actual_dates.dt.normalize().eq(today).fillna(False)

        return is_shipment & shipped_today

    return pd.Series(True, index=idx)


# ================================
# RENDER TABLE
# ================================

def render_teable_subset_table(title: str, source_df: pd.DataFrame, specs, subset_mode: str):
    st.subheader(title)
    mask = build_subset_mask(source_df, subset_mode)
    filtered = source_df[mask].copy()
    view_df, _ = build_teable_view_df(filtered, specs)
    st.caption(f"共 {len(view_df)} 筆")
    st.dataframe(view_df, use_container_width=True, hide_index=True)


# ================================
# MANUAL HISTORY（支援新舊兩種 CSV 格式）
# ================================

def _normalize_manual_df(raw: pd.DataFrame) -> pd.DataFrame:
    """將補登 CSV 欄位名稱統一轉換為內部標準欄位名稱（支援中文/英文）"""
    renamed = {}
    for col in raw.columns:
        std = MANUAL_COL_MAP.get(str(col).strip())
        if std:
            renamed[col] = std
    df = raw.rename(columns=renamed)
    for c in MANUAL_REQUIRED_COLS:
        if c not in df.columns:
            df[c] = ""
    keep = MANUAL_REQUIRED_COLS + (["month"] if "month" in df.columns else [])
    df = df[[c for c in keep if c in df.columns]]
    df["parsed_date"] = parse_mixed_date_series(df["date"])
    df["amount_num"]  = parse_amount_series(df["amount"])
    return df


def _load_manual_history() -> pd.DataFrame:
    if MANUAL_HISTORY_PATH.exists():
        try:
            raw = pd.read_csv(MANUAL_HISTORY_PATH)
            return _normalize_manual_df(raw)
        except Exception:
            pass
    empty = pd.DataFrame(columns=MANUAL_REQUIRED_COLS)
    empty["parsed_date"] = pd.Series(dtype="datetime64[ns]")
    empty["amount_num"]  = pd.Series(dtype="float64")
    return empty


def _save_manual_history(df: pd.DataFrame):
    out = df.copy()
    for c in ["parsed_date", "amount_num"]:
        if c in out.columns:
            out = out.drop(columns=[c])
    # 儲存用中文欄位名（與現有 CSV 格式一致）
    out = out.rename(columns={
        "date": "日期", "type": "類型", "customer": "客戶", "factory": "工廠",
        "po": "PO#", "pn": "P/N", "qty": "QTY", "wip": "WIP",
        "amount": "金額(USD)", "note": "備註",
    })
    out.to_csv(MANUAL_HISTORY_PATH, index=False, encoding="utf-8-sig")


def _render_manual_input(selected_period: pd.Period):
    with st.expander("歷史補登（今天以前的金額）"):
        manual_df = _load_manual_history()

        dl_df = manual_df.drop(columns=[c for c in ["parsed_date", "amount_num"] if c in manual_df.columns])
        st.download_button(
            "下載補登 CSV",
            dl_df.to_csv(index=False, encoding="utf-8-sig"),
            file_name="sales_manual_history.csv",
            mime="text/csv",
        )

        up = st.file_uploader("匯入補登 CSV", type=["csv"], key="manual_history_csv")
        if up is not None:
            try:
                imported_raw = pd.read_csv(up)
                manual_df = _normalize_manual_df(imported_raw)
                _save_manual_history(manual_df)
                st.success("補登 CSV 已匯入")
            except Exception as e:
                st.error(f"匯入失敗：{e}")

        show_df = manual_df.drop(columns=[c for c in ["parsed_date", "amount_num"] if c in manual_df.columns])
        st.dataframe(show_df, use_container_width=True, hide_index=True)

        with st.form("manual_history_form"):
            c1, c2, c3 = st.columns(3)
            date_val = c1.date_input("日期", value=selected_period.start_time.date())
            typ      = c2.selectbox("類型", ["接單", "已出貨", "預計出貨"])
            amount   = c3.text_input("金額(USD)", value="")
            c4, c5, c6 = st.columns(3)
            customer = c4.text_input("客戶", value="")
            factory  = c5.text_input("工廠", value="")
            po       = c6.text_input("PO#", value="")
            c7, c8, c9 = st.columns(3)
            pn  = c7.text_input("P/N", value="")
            qty = c8.text_input("QTY", value="")
            wip = c9.text_input("WIP", value="")
            note = st.text_input("備註", value="")
            if st.form_submit_button("儲存本月補登"):
                row = {
                    "date": pd.Timestamp(date_val).strftime("%Y-%m-%d"),
                    "type": typ, "customer": customer, "factory": factory,
                    "po": po, "pn": pn, "qty": qty, "wip": wip,
                    "amount": amount, "note": note,
                }
                drop_cols = [c for c in ["parsed_date", "amount_num"] if c in manual_df.columns]
                manual_df = pd.concat(
                    [manual_df.drop(columns=drop_cols), pd.DataFrame([row])],
                    ignore_index=True
                )
                _save_manual_history(manual_df)
                st.success("已儲存補登資料，重新整理頁面後會重新計算。")
        st.caption(f"補登資料儲存在：{MANUAL_HISTORY_PATH.name}")


# ================================
# 業績明細表
# ================================

def render_sales_detail_from_teable(source_df: pd.DataFrame):
    st.subheader("業績明細表")
    if source_df is None or source_df.empty:
        st.info("目前沒有資料。")
        return

    order_col     = find_col(source_df, ORDER_DATE_CANDIDATES)
    actual_col    = find_col(source_df, ACTUAL_SHIP_DATE_CANDIDATES)
    plan_col      = find_col(source_df, PLANNED_SHIP_DATE_CANDIDATES)
    ship_amt_col  = find_col(source_df, AMOUNT_SHIP_CANDIDATES)
    order_amt_col = find_col(source_df, AMOUNT_ORDER_CANDIDATES)
    customer_col  = find_col(source_df, CUSTOMER_CANDIDATES)
    factory_col   = find_col(source_df, FACTORY_CANDIDATES)
    po_col        = find_col(source_df, PO_CANDIDATES)
    pn_col        = find_col(source_df, PART_CANDIDATES)
    qty_col       = find_col(source_df, QTY_CANDIDATES)
    wip_col       = find_col(source_df, WIP_CANDIDATES)

    order_dates  = parse_mixed_date_series(get_series_by_col(source_df, order_col))
    actual_dates = parse_mixed_date_series(get_series_by_col(source_df, actual_col))
    if actual_col != find_col(source_df, ["出貨日期"]):
        fallback_actual = parse_mixed_date_series(
            get_series_by_col(source_df, find_col(source_df, ["出貨日期"]))
        )
        actual_dates = actual_dates.where(actual_dates.notna(), fallback_actual)
    plan_dates = parse_mixed_date_series(get_series_by_col(source_df, plan_col))

    ship_amount  = parse_amount_series(get_series_by_col(source_df, ship_amt_col))
    order_amount = parse_amount_series(get_series_by_col(source_df, order_amt_col))
    amount = ship_amount.where(ship_amount.notna() & (ship_amount != 0), order_amount)

    wip       = get_series_by_col(source_df, wip_col).astype(str).str.upper().str.strip() if wip_col else pd.Series("", index=source_df.index)
    customers = get_series_by_col(source_df, customer_col).astype(str).fillna("") if customer_col else pd.Series("", index=source_df.index)
    factories = get_series_by_col(source_df, factory_col).astype(str).fillna("") if factory_col else pd.Series("", index=source_df.index)
    pos       = get_series_by_col(source_df, po_col).astype(str).fillna("") if po_col else pd.Series("", index=source_df.index)
    pns       = get_series_by_col(source_df, pn_col).astype(str).fillna("") if pn_col else pd.Series("", index=source_df.index)
    qtys      = get_series_by_col(source_df, qty_col).astype(str).fillna("") if qty_col else pd.Series("", index=source_df.index)

    periods = set()
    for s in [order_dates, actual_dates, plan_dates]:
        periods |= {p for p in s.dt.to_period("M").dropna().unique().tolist()}
    manual_df = _load_manual_history()
    periods |= {p for p in manual_df["parsed_date"].dt.to_period("M").dropna().unique().tolist()}

    if not periods:
        st.info("找不到可用月份。")
        return

    periods = sorted(periods)
    selected = st.selectbox(
        "月份", periods, index=len(periods) - 1,
        format_func=lambda p: f"{p.year}-{p.month:02d}"
    )

    _render_manual_input(selected)
    manual_df    = _load_manual_history()
    manual_month = manual_df[manual_df["parsed_date"].dt.to_period("M") == selected].copy()

    is_shipment   = wip.eq("SHIPMENT")
    actual_mask   = actual_dates.dt.to_period("M") == selected
    plan_mask     = plan_dates.dt.to_period("M") == selected
    order_mask    = order_dates.dt.to_period("M") == selected
    shipped_mask  = is_shipment & actual_mask
    forecast_mask = (~is_shipment) & plan_mask

    shipped_df = pd.DataFrame({
        "日期": actual_dates[shipped_mask].dt.strftime("%Y-%m-%d"),
        "客戶": customers[shipped_mask], "工廠": factories[shipped_mask],
        "PO#": pos[shipped_mask], "P/N": pns[shipped_mask],
        "QTY": qtys[shipped_mask], "WIP": wip[shipped_mask],
        "金額(USD)": amount[shipped_mask],
    }).reset_index(drop=True)

    forecast_df = pd.DataFrame({
        "日期": plan_dates[forecast_mask].dt.strftime("%Y-%m-%d"),
        "客戶": customers[forecast_mask], "工廠": factories[forecast_mask],
        "PO#": pos[forecast_mask], "P/N": pns[forecast_mask],
        "QTY": qtys[forecast_mask], "WIP": wip[forecast_mask],
        "金額(USD)": amount[forecast_mask],
    }).reset_index(drop=True)

    order_total    = float(order_amount[order_mask].fillna(0).sum())
    shipped_total  = float(shipped_df["金額(USD)"].fillna(0).sum())
    forecast_total = float(forecast_df["金額(USD)"].fillna(0).sum())

    if not manual_month.empty:
        def _add(typ):
            return float(manual_month.loc[manual_month["type"].eq(typ), "amount_num"].fillna(0).sum())
        def _rows(typ):
            m = manual_month[manual_month["type"].eq(typ)]
            return pd.DataFrame({
                "日期": m["date"], "客戶": m["customer"], "工廠": m["factory"],
                "PO#": m["po"], "P/N": m["pn"], "QTY": m["qty"],
                "WIP": m["wip"], "金額(USD)": m["amount_num"],
            })
        order_total    += _add("接單")
        shipped_total  += _add("已出貨")
        forecast_total += _add("預計出貨")
        if _add("已出貨"):
            shipped_df  = pd.concat([shipped_df,  _rows("已出貨")],  ignore_index=True)
        if _add("預計出貨"):
            forecast_df = pd.concat([forecast_df, _rows("預計出貨")], ignore_index=True)

    month_total = shipped_total + forecast_total

    st.markdown(f"### {selected.month}月 業績明細表")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("接單金額 (USD)",     f"${order_total:,.2f}",    delta=f"{int(order_mask.sum() + manual_month['type'].eq('接單').sum())} 筆")
    c2.metric("已確認出貨 (USD)",   f"${shipped_total:,.2f}",  delta=f"{len(shipped_df)} 筆")
    c3.metric("預計本月出貨 (USD)", f"${forecast_total:,.2f}", delta=f"{len(forecast_df)} 筆")
    c4.metric("月銷貨合計 (USD)",   f"${month_total:,.2f}")

    st.markdown(f"- 已確認出貨（SHIPMENT）：US${shipped_total:,.2f}")
    st.markdown(f"- 預計{selected.month}月出貨（未 SHIPMENT）：US${forecast_total:,.2f}")
    st.markdown(f"- {selected.month}月份銷貨金額總計：US${month_total:,.2f}")

    if not forecast_df.empty:
        first = forecast_df.iloc[0]
        try:
            date_str = pd.to_datetime(first["日期"]).strftime("%-m/%-d")
        except Exception:
            date_str = str(first["日期"])
        st.info(
            f"💡 {first['WIP']} 中（預計 {date_str} 出貨）："
            f"{first['客戶']} | US${float(first['金額(USD)']):,.2f}\n\n"
            f"出貨後 WIP 更新為 SHIPMENT，{selected.month}月銷貨合計將增至 US${month_total:,.2f}。"
        )

    left, right = st.columns(2)
    with left:
        st.markdown(f"#### 🏭 依工廠別統計（{selected.month}月銷貨）")
        if shipped_df.empty:
            st.info("本月無已出貨資料。")
        else:
            fac = shipped_df.groupby("工廠", dropna=False).agg(
                訂單數=("PO#", "count"), 銷貨金額USD=("金額(USD)", "sum")
            ).reset_index()
            fac.columns = ["工廠", "訂單數", "銷貨金額(USD)"]
            fac.loc[len(fac)] = ["合計", int(fac["訂單數"].sum()), float(fac["銷貨金額(USD)"].sum())]
            st.dataframe(fac, use_container_width=True, hide_index=True)
    with right:
        st.markdown(f"#### 👥 依客戶別統計（{selected.month}月銷貨）")
        if shipped_df.empty:
            st.info("本月無已出貨資料。")
        else:
            cus = shipped_df.groupby("客戶", dropna=False).agg(
                訂單數=("PO#", "count"), 銷貨金額USD=("金額(USD)", "sum")
            ).reset_index()
            cus.columns = ["客戶", "訂單數", "銷貨金額(USD)"]
            cus.loc[len(cus)] = ["合計", int(cus["訂單數"].sum()), float(cus["銷貨金額(USD)"].sum())]
            st.dataframe(cus, use_container_width=True, hide_index=True)

    st.markdown("#### 已出貨明細")
    st.dataframe(shipped_df, use_container_width=True, hide_index=True)
    st.markdown("#### 預計出貨明細")
    st.dataframe(forecast_df, use_container_width=True, hide_index=True)

    monthly = []
    for p in sorted(periods)[-12:]:
        s_amt = float(amount[is_shipment & (actual_dates.dt.to_period("M") == p)].fillna(0).sum())
        f_amt = float(amount[(~is_shipment) & (plan_dates.dt.to_period("M") == p)].fillna(0).sum())
        mm = manual_df[manual_df["parsed_date"].dt.to_period("M") == p]
        s_amt += float(mm.loc[mm["type"].eq("已出貨"),   "amount_num"].fillna(0).sum())
        f_amt += float(mm.loc[mm["type"].eq("預計出貨"), "amount_num"].fillna(0).sum())
        monthly.append({"月份": str(p), "已出貨": s_amt, "預計出貨": f_amt, "銷貨合計": s_amt + f_amt})
    st.markdown("#### 近 12 個月月銷貨趨勢")
    st.dataframe(pd.DataFrame(monthly), use_container_width=True, hide_index=True)

    with st.expander("Debug：業績明細表欄位偵測"):
        st.json({
            "order_col": order_col, "actual_col": actual_col, "plan_col": plan_col,
            "ship_amt_col": ship_amt_col, "order_amt_col": order_amt_col,
            "customer_col": customer_col, "factory_col": factory_col, "wip_col": wip_col,
            "selected_month": str(selected),
            "shipped_total": shipped_total, "forecast_total": forecast_total, "month_total": month_total,
        })


# ================================
# PUBLIC ENTRY POINTS
# ================================

def show_new_orders_wip_report(source_df: pd.DataFrame):
    render_teable_subset_table("📄 新訂單 WIP", source_df, SANDY_NEW_ORDER_SPECS, "new_order_today")


def show_sandy_internal_wip_report(source_df: pd.DataFrame):
    render_teable_subset_table("📄 Sandy 內部 WIP", source_df, SANDY_INTERNAL_WIP_SPECS, "unshipped")


def show_sandy_sales_report(source_df: pd.DataFrame):
    render_teable_subset_table("📄 Sandy 銷貨底", source_df, SANDY_SALES_SPECS, "shipment_only")
