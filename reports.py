# -*- coding: utf-8 -*-
"""
reports.py  ── GLOCOM Control Tower 業績明細表
修正重點（v2）：
  1. 手動補登改存 Teable（獨立 table），解決 Streamlit Cloud 重啟後資料消失問題
  2. shipped_df / forecast_df 合併時金額欄位型別統一（均為 float）
  3. 近 12 個月趨勢正確加入補登金額（含非當月歷史資料）
  4. 依工廠 / 客戶統計同步加入補登已出貨資料
"""

from __future__ import annotations

from pathlib import Path
import re
import pandas as pd
import requests
import streamlit as st

try:
    import config as cfg
except ImportError:
    cfg = None  # type: ignore

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

CANCELLED_KEYWORDS = [
    "PO CANCELLED", "PO CANCELED", "CANCELLATION", "CANCELLED", "CANCELED", "CANCEL",
]

# ── 補登 CSV 欄位對照表 ──────────────────────────────────────────────────────
MANUAL_COL_MAP = {
    "月份": "month", "日期": "date", "類型": "type",
    "客戶": "customer", "工廠": "factory",
    "PO#": "po", "P/N": "pn", "QTY": "qty", "WIP": "wip",
    "金額(USD)": "amount", "備註": "note",
    "date": "date", "type": "type", "customer": "customer", "factory": "factory",
    "po": "po", "pn": "pn", "qty": "qty", "wip": "wip", "amount": "amount", "note": "note",
}
MANUAL_REQUIRED_COLS = ["date", "type", "customer", "factory", "po", "pn", "qty", "wip", "amount", "note"]

# ── 本機備援路徑（Teable 不可用時才用） ────────────────────────────────────
MANUAL_HISTORY_PATH = Path("sales_manual_history.csv")


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
    if series is None:
        return pd.Series(dtype="datetime64[ns]")

    s = series.astype(str).fillna("").str.strip()
    s = s.replace({"": None, "nan": None, "NaT": None, "None": None})
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")

    nums = pd.to_numeric(s, errors="coerce")
    mask_num = nums.notna() & (nums > 20000) & (nums < 80000)
    if mask_num.any():
        out.loc[mask_num] = pd.to_datetime(
            nums.loc[mask_num], unit="D", origin="1899-12-30", errors="coerce"
        )

    rem = out.isna() & s.notna()
    if rem.any():
        out.loc[rem] = pd.to_datetime(s.loc[rem], format="%Y-%m-%d", errors="coerce")

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

    rem = out.isna() & s.notna()
    if rem.any():
        out.loc[rem] = pd.to_datetime(s.loc[rem], format="%y-%m-%d", errors="coerce")

    rem = out.isna() & s.notna()
    if rem.any():
        out.loc[rem] = pd.to_datetime(s.loc[rem], errors="coerce")

    return out


def _is_cancelled(wip_series: pd.Series) -> pd.Series:
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
    idx = source_df.index
    today = pd.Timestamp.today().normalize()
    current_year = today.year

    wip_col = find_col(source_df, WIP_CANDIDATES)
    wip_raw = (
        get_series_by_col(source_df, wip_col).astype(str).str.strip()
        if wip_col else pd.Series("", index=idx)
    )
    wip_upper = wip_raw.str.upper()

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

    if subset_mode == "unshipped":
        not_shipment = ~wip_upper.eq("SHIPMENT")
        not_cancelled = ~_is_cancelled(wip_raw)

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

        year_series = (
            cust_dates.dt.year
            .where(cust_dates.notna(), fact_dates.dt.year)
            .where(cust_dates.notna() | fact_dates.notna(), ship_dates.dt.year)
        )

        year_ok = year_series.isna() | year_series.eq(current_year)
        return not_shipment & not_cancelled & year_ok

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
# MANUAL HISTORY — Teable 優先，本機備援
# ================================

def _get_manual_table_id() -> str:
    """從 secrets 或 config 取得補登用的 Teable Table ID"""
    if "TEABLE_MANUAL_TABLE_ID" in st.secrets:
        return str(st.secrets["TEABLE_MANUAL_TABLE_ID"]).strip()
    if cfg and hasattr(cfg, "TEABLE_MANUAL_TABLE_ID"):
        return str(cfg.TEABLE_MANUAL_TABLE_ID).strip()
    return ""


def _get_teable_token() -> str:
    if "TEABLE_TOKEN" in st.secrets:
        return str(st.secrets["TEABLE_TOKEN"]).strip()
    if cfg and hasattr(cfg, "TEABLE_TOKEN"):
        return str(cfg.TEABLE_TOKEN).strip()
    return ""


def _get_teable_api_base() -> str:
    if cfg and hasattr(cfg, "TEABLE_API_BASE"):
        return str(cfg.TEABLE_API_BASE).rstrip("/")
    return "https://app.teable.io/api"


def _teable_headers() -> dict:
    return {
        "Authorization": f"Bearer {_get_teable_token()}",
        "Content-Type": "application/json",
    }


# ── Teable 補登讀取 ──────────────────────────────────────────────────────────

@st.cache_data(ttl=60, show_spinner=False)
def _load_manual_from_teable() -> pd.DataFrame:
    """
    從 Teable 補登 table 讀取全部補登記錄
    回傳標準化後的 DataFrame（含 parsed_date, amount_num）
    """
    table_id = _get_manual_table_id()
    if not table_id:
        return _empty_manual_df()

    base = _get_teable_api_base()
    url = f"{base}/table/{table_id}/record"

    all_rows = []
    take, skip = 200, 0
    for _ in range(50):
        try:
            resp = requests.get(
                url, headers=_teable_headers(),
                params={"take": take, "skip": skip}, timeout=15
            )
            if resp.status_code != 200:
                break
            data = resp.json()
            items = data.get("records", [])
            if not items:
                break
            for rec in items:
                fields = rec.get("fields", {}) or {}
                fields["_record_id"] = rec.get("id", "")
                all_rows.append(fields)
            total = data.get("total")
            if total is not None and skip + take >= total:
                break
            if len(items) < take:
                break
            skip += take
        except Exception:
            break

    if not all_rows:
        return _empty_manual_df()

    raw = pd.DataFrame(all_rows)
    return _normalize_manual_df(raw)


def _save_manual_to_teable(row: dict) -> tuple[bool, str]:
    """
    新增一筆補登記錄到 Teable
    row: 內部標準欄位格式 {date, type, customer, factory, po, pn, qty, wip, amount, note}
    """
    table_id = _get_manual_table_id()
    if not table_id:
        return False, "TEABLE_MANUAL_TABLE_ID 未設定"

    base = _get_teable_api_base()
    url = f"{base}/table/{table_id}/record"

    # 轉回 Teable 中文欄位名
    fields = {
        "日期":     row.get("date", ""),
        "類型":     row.get("type", ""),
        "客戶":     row.get("customer", ""),
        "工廠":     row.get("factory", ""),
        "PO#":      row.get("po", ""),
        "P/N":      row.get("pn", ""),
        "QTY":      row.get("qty", ""),
        "WIP":      row.get("wip", ""),
        "金額(USD)": str(row.get("amount", "")),
        "備註":     row.get("note", ""),
    }
    payload = {"records": [{"fields": fields}]}

    try:
        resp = requests.post(url, headers=_teable_headers(), json=payload, timeout=15)
        if resp.status_code in (200, 201):
            return True, ""
        return False, f"HTTP {resp.status_code}: {resp.text[:300]}"
    except Exception as e:
        return False, str(e)


def _delete_manual_from_teable(record_id: str) -> tuple[bool, str]:
    """從 Teable 刪除指定補登記錄"""
    table_id = _get_manual_table_id()
    if not table_id:
        return False, "TEABLE_MANUAL_TABLE_ID 未設定"

    base = _get_teable_api_base()
    url = f"{base}/table/{table_id}/record/{record_id}"
    try:
        resp = requests.delete(url, headers=_teable_headers(), timeout=15)
        if resp.status_code in (200, 204):
            return True, ""
        return False, f"HTTP {resp.status_code}: {resp.text[:300]}"
    except Exception as e:
        return False, str(e)


# ── 本機備援（Teable 未設定時） ───────────────────────────────────────────────

def _empty_manual_df() -> pd.DataFrame:
    empty = pd.DataFrame(columns=MANUAL_REQUIRED_COLS + ["_record_id"])
    empty["parsed_date"] = pd.Series(dtype="datetime64[ns]")
    empty["amount_num"]  = pd.Series(dtype="float64")
    return empty


def _normalize_manual_df(raw: pd.DataFrame) -> pd.DataFrame:
    """統一轉換補登 CSV / Teable 欄位名稱為內部標準名稱"""
    renamed = {}
    for col in raw.columns:
        std = MANUAL_COL_MAP.get(str(col).strip())
        if std:
            renamed[col] = std
    df = raw.rename(columns=renamed)
    for c in MANUAL_REQUIRED_COLS:
        if c not in df.columns:
            df[c] = ""
    keep = MANUAL_REQUIRED_COLS + [c for c in ["month", "_record_id"] if c in df.columns]
    df = df[[c for c in keep if c in df.columns]]
    df["parsed_date"] = parse_mixed_date_series(df["date"])
    # ★ 修正：amount_num 強制轉 float，避免後續 concat 後型別錯誤
    df["amount_num"] = parse_amount_series(df["amount"])
    return df


def _load_manual_history() -> pd.DataFrame:
    """
    優先從 Teable 讀；若 TEABLE_MANUAL_TABLE_ID 未設定，則退回本機 CSV。
    """
    table_id = _get_manual_table_id()
    if table_id:
        return _load_manual_from_teable()

    # 本機備援
    if MANUAL_HISTORY_PATH.exists():
        try:
            raw = pd.read_csv(MANUAL_HISTORY_PATH)
            return _normalize_manual_df(raw)
        except Exception:
            pass
    return _empty_manual_df()


def _save_manual_history_local(df: pd.DataFrame):
    """本機備援儲存（只在 Teable 未設定時使用）"""
    out = df.copy()
    for c in ["parsed_date", "amount_num", "_record_id"]:
        if c in out.columns:
            out = out.drop(columns=[c])
    out = out.rename(columns={
        "date": "日期", "type": "類型", "customer": "客戶", "factory": "工廠",
        "po": "PO#", "pn": "P/N", "qty": "QTY", "wip": "WIP",
        "amount": "金額(USD)", "note": "備註",
    })
    out.to_csv(MANUAL_HISTORY_PATH, index=False, encoding="utf-8-sig")


# ================================
# MANUAL INPUT UI
# ================================

def _render_manual_input(selected_period: pd.Period):
    table_id = _get_manual_table_id()
    storage_label = "Teable" if table_id else f"本機 ({MANUAL_HISTORY_PATH.name})"

    with st.expander(f"歷史補登（今天以前的金額）— 儲存至 {storage_label}"):
        manual_df = _load_manual_history()

        # 顯示現有補登
        show_cols = [c for c in MANUAL_REQUIRED_COLS if c in manual_df.columns]
        st.dataframe(manual_df[show_cols], use_container_width=True, hide_index=True)

        # ── 如果是本機模式，提供 CSV 上傳/下載 ───────────────────────────
        if not table_id:
            dl_df = manual_df[[c for c in show_cols]]
            st.download_button(
                "下載補登 CSV",
                dl_df.rename(columns={
                    "date": "日期", "type": "類型", "customer": "客戶", "factory": "工廠",
                    "po": "PO#", "pn": "P/N", "qty": "QTY", "wip": "WIP",
                    "amount": "金額(USD)", "note": "備註",
                }).to_csv(index=False, encoding="utf-8-sig"),
                file_name="sales_manual_history.csv",
                mime="text/csv",
            )
            up = st.file_uploader("匯入補登 CSV", type=["csv"], key="manual_history_csv")
            if up is not None:
                try:
                    imported_raw = pd.read_csv(up)
                    imported = _normalize_manual_df(imported_raw)
                    _save_manual_history_local(imported)
                    st.success("補登 CSV 已匯入，請重新整理頁面。")
                except Exception as e:
                    st.error(f"匯入失敗：{e}")

        # ── 刪除功能（Teable 模式才顯示，本機無 record_id） ──────────────
        if table_id and "_record_id" in manual_df.columns and not manual_df.empty:
            del_options = {
                f"{r['date']} | {r['type']} | {r['customer']} | ${r['amount']}": r["_record_id"]
                for _, r in manual_df.iterrows()
                if r.get("_record_id")
            }
            if del_options:
                del_label = st.selectbox("選擇要刪除的補登記錄", ["（不刪除）"] + list(del_options.keys()))
                if st.button("刪除選取記錄") and del_label != "（不刪除）":
                    ok, err = _delete_manual_from_teable(del_options[del_label])
                    if ok:
                        st.success("已刪除，重新整理頁面後生效。")
                        _load_manual_from_teable.clear()
                    else:
                        st.error(f"刪除失敗：{err}")

        # ── 新增表單 ──────────────────────────────────────────────────────
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

            if st.form_submit_button("儲存補登"):
                row = {
                    "date": pd.Timestamp(date_val).strftime("%Y-%m-%d"),
                    "type": typ, "customer": customer, "factory": factory,
                    "po": po, "pn": pn, "qty": qty, "wip": wip,
                    "amount": amount, "note": note,
                }
                if table_id:
                    ok, err = _save_manual_to_teable(row)
                    if ok:
                        _load_manual_from_teable.clear()  # 清快取
                        st.success("已儲存到 Teable，重新整理頁面後會重新計算。")
                    else:
                        st.error(f"儲存失敗：{err}")
                else:
                    # 本機備援
                    drop_cols = [c for c in ["parsed_date", "amount_num", "_record_id"] if c in manual_df.columns]
                    manual_df = pd.concat(
                        [manual_df.drop(columns=drop_cols), pd.DataFrame([row])],
                        ignore_index=True
                    )
                    _save_manual_history_local(manual_df)
                    st.success("已儲存補登資料，重新整理頁面後會重新計算。")


# ================================
# 業績明細表（核心）
# ================================

def render_sales_detail_from_teable(source_df: pd.DataFrame):
    st.subheader("業績明細表")
    if source_df is None or source_df.empty:
        st.info("目前沒有資料。")
        return

    # ── 欄位偵測 ─────────────────────────────────────────────────────────────
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

    # ── 解析日期與金額 ────────────────────────────────────────────────────────
    order_dates  = parse_mixed_date_series(get_series_by_col(source_df, order_col))
    actual_dates = parse_mixed_date_series(get_series_by_col(source_df, actual_col))
    # 出貨日期 fallback
    fallback_actual_col = find_col(source_df, ["出貨日期"])
    if fallback_actual_col and fallback_actual_col != actual_col:
        fallback_actual = parse_mixed_date_series(get_series_by_col(source_df, fallback_actual_col))
        actual_dates = actual_dates.where(actual_dates.notna(), fallback_actual)
    plan_dates = parse_mixed_date_series(get_series_by_col(source_df, plan_col))

    ship_amount  = parse_amount_series(get_series_by_col(source_df, ship_amt_col))
    order_amount = parse_amount_series(get_series_by_col(source_df, order_amt_col))
    # 金額優先取銷貨金額，其次接單金額
    amount = ship_amount.where(ship_amount.notna() & (ship_amount != 0), order_amount)

    wip       = get_series_by_col(source_df, wip_col).astype(str).str.upper().str.strip() if wip_col else pd.Series("", index=source_df.index)
    customers = get_series_by_col(source_df, customer_col).astype(str).fillna("") if customer_col else pd.Series("", index=source_df.index)
    factories = get_series_by_col(source_df, factory_col).astype(str).fillna("") if factory_col else pd.Series("", index=source_df.index)
    pos       = get_series_by_col(source_df, po_col).astype(str).fillna("") if po_col else pd.Series("", index=source_df.index)
    pns       = get_series_by_col(source_df, pn_col).astype(str).fillna("") if pn_col else pd.Series("", index=source_df.index)
    qtys      = get_series_by_col(source_df, qty_col).astype(str).fillna("") if qty_col else pd.Series("", index=source_df.index)

    # ── 月份清單（Teable 資料 + 補登資料） ───────────────────────────────────
    periods: set = set()
    for s in [order_dates, actual_dates, plan_dates]:
        periods |= {p for p in s.dt.to_period("M").dropna().unique().tolist()}

    # ★ 修正：先讀補登，加入補登月份到可選清單
    manual_df = _load_manual_history()
    if not manual_df.empty:
        periods |= {p for p in manual_df["parsed_date"].dt.to_period("M").dropna().unique().tolist()}

    if not periods:
        st.info("找不到可用月份。")
        return

    periods = sorted(periods)
    selected = st.selectbox(
        "月份", periods, index=len(periods) - 1,
        format_func=lambda p: f"{p.year}-{p.month:02d}"
    )

    # ── 補登 UI ──────────────────────────────────────────────────────────────
    _render_manual_input(selected)

    # 補登 UI 可能新增/刪除，重新載入
    manual_df    = _load_manual_history()
    manual_month = manual_df[manual_df["parsed_date"].dt.to_period("M") == selected].copy()

    # ── Teable 資料過濾 ───────────────────────────────────────────────────────
    is_shipment   = wip.eq("SHIPMENT")
    actual_mask   = actual_dates.dt.to_period("M") == selected
    plan_mask     = plan_dates.dt.to_period("M") == selected
    order_mask    = order_dates.dt.to_period("M") == selected
    shipped_mask  = is_shipment & actual_mask
    forecast_mask = (~is_shipment) & plan_mask

    # ★ 修正：amount 欄統一用 float，避免 concat 後型別錯誤
    shipped_df = pd.DataFrame({
        "日期":      actual_dates[shipped_mask].dt.strftime("%Y-%m-%d"),
        "客戶":      customers[shipped_mask].values,
        "工廠":      factories[shipped_mask].values,
        "PO#":       pos[shipped_mask].values,
        "P/N":       pns[shipped_mask].values,
        "QTY":       qtys[shipped_mask].values,
        "WIP":       wip[shipped_mask].values,
        "金額(USD)": amount[shipped_mask].astype(float).values,
    }).reset_index(drop=True)

    forecast_df = pd.DataFrame({
        "日期":      plan_dates[forecast_mask].dt.strftime("%Y-%m-%d"),
        "客戶":      customers[forecast_mask].values,
        "工廠":      factories[forecast_mask].values,
        "PO#":       pos[forecast_mask].values,
        "P/N":       pns[forecast_mask].values,
        "QTY":       qtys[forecast_mask].values,
        "WIP":       wip[forecast_mask].values,
        "金額(USD)": amount[forecast_mask].astype(float).values,
    }).reset_index(drop=True)

    # ── 計算 Teable 小計 ──────────────────────────────────────────────────────
    order_total    = float(order_amount[order_mask].fillna(0).sum())
    shipped_total  = float(shipped_df["金額(USD)"].fillna(0).sum())
    forecast_total = float(forecast_df["金額(USD)"].fillna(0).sum())

    # ── 加入補登金額 ─────────────────────────────────────────────────────────
    if not manual_month.empty:
        def _manual_sum(typ: str) -> float:
            return float(manual_month.loc[manual_month["type"] == typ, "amount_num"].fillna(0).sum())

        def _manual_rows(typ: str) -> pd.DataFrame:
            m = manual_month[manual_month["type"] == typ].copy()
            return pd.DataFrame({
                "日期":      m["date"].values,
                "客戶":      m["customer"].values,
                "工廠":      m["factory"].values,
                "PO#":       m["po"].values,
                "P/N":       m["pn"].values,
                "QTY":       m["qty"].values,
                "WIP":       m["wip"].values,
                # ★ 修正：amount_num 已是 float，直接使用
                "金額(USD)": m["amount_num"].astype(float).values,
            })

        order_total    += _manual_sum("接單")
        shipped_total  += _manual_sum("已出貨")
        forecast_total += _manual_sum("預計出貨")

        # ★ 修正：concat 前確保欄位與型別一致
        if _manual_sum("已出貨") > 0:
            shipped_df = pd.concat(
                [shipped_df, _manual_rows("已出貨")], ignore_index=True
            )
        if _manual_sum("預計出貨") > 0:
            forecast_df = pd.concat(
                [forecast_df, _manual_rows("預計出貨")], ignore_index=True
            )

    month_total = shipped_total + forecast_total

    # ── 指標顯示 ─────────────────────────────────────────────────────────────
    st.markdown(f"### {selected.month}月 業績明細表")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("接單金額 (USD)",     f"${order_total:,.2f}",    delta=f"{int(order_mask.sum() + (manual_month['type'] == '接單').sum())} 筆")
    c2.metric("已確認出貨 (USD)",   f"${shipped_total:,.2f}",  delta=f"{len(shipped_df)} 筆")
    c3.metric("預計本月出貨 (USD)", f"${forecast_total:,.2f}", delta=f"{len(forecast_df)} 筆")
    c4.metric("月銷貨合計 (USD)",   f"${month_total:,.2f}")

    st.markdown(f"- 已確認出貨（SHIPMENT）：US${shipped_total:,.2f}")
    st.markdown(f"- 預計{selected.month}月出貨（QA 中）：US${forecast_total:,.2f}")
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

    # ── 依工廠 / 客戶統計 ────────────────────────────────────────────────────
    left, right = st.columns(2)
    with left:
        st.markdown(f"#### 🏭 依工廠別統計（{selected.month}月銷貨）")
        if shipped_df.empty:
            st.info("本月無已出貨資料。")
        else:
            fac = shipped_df.groupby("工廠", dropna=False).agg(
                訂單數=("PO#", "count"),
                銷貨金額=("金額(USD)", "sum")
            ).reset_index()
            fac.columns = ["工廠", "訂單數", "銷貨金額(USD)"]
            total_row = pd.DataFrame([["合計", int(fac["訂單數"].sum()), float(fac["銷貨金額(USD)"].sum())]],
                                     columns=fac.columns)
            fac = pd.concat([fac, total_row], ignore_index=True)
            st.dataframe(fac, use_container_width=True, hide_index=True)

    with right:
        st.markdown(f"#### 👥 依客戶別統計（{selected.month}月銷貨）")
        if shipped_df.empty:
            st.info("本月無已出貨資料。")
        else:
            cus = shipped_df.groupby("客戶", dropna=False).agg(
                訂單數=("PO#", "count"),
                銷貨金額=("金額(USD)", "sum")
            ).reset_index()
            cus.columns = ["客戶", "訂單數", "銷貨金額(USD)"]
            total_row = pd.DataFrame([["合計", int(cus["訂單數"].sum()), float(cus["銷貨金額(USD)"].sum())]],
                                     columns=cus.columns)
            cus = pd.concat([cus, total_row], ignore_index=True)
            st.dataframe(cus, use_container_width=True, hide_index=True)

    # ── 明細表 ────────────────────────────────────────────────────────────────
    st.markdown("#### 已出貨明細")
    st.dataframe(shipped_df, use_container_width=True, hide_index=True)
    st.markdown("#### 預計出貨明細")
    st.dataframe(forecast_df, use_container_width=True, hide_index=True)

    # ── 近 12 個月趨勢 ────────────────────────────────────────────────────────
    # ★ 修正：全部月份（不限選取月）都正確加入補登金額
    monthly = []
    all_periods_12 = sorted(periods)[-12:]
    for p in all_periods_12:
        s_amt = float(amount[is_shipment & (actual_dates.dt.to_period("M") == p)].fillna(0).sum())
        f_amt = float(amount[(~is_shipment) & (plan_dates.dt.to_period("M") == p)].fillna(0).sum())

        # 加入該月補登
        mm = manual_df[manual_df["parsed_date"].dt.to_period("M") == p]
        if not mm.empty:
            s_amt += float(mm.loc[mm["type"] == "已出貨",   "amount_num"].fillna(0).sum())
            f_amt += float(mm.loc[mm["type"] == "預計出貨", "amount_num"].fillna(0).sum())

        monthly.append({
            "月份":   str(p),
            "已出貨": round(s_amt, 2),
            "預計出貨": round(f_amt, 2),
            "銷貨合計": round(s_amt + f_amt, 2),
        })

    st.markdown("#### 近 12 個月月銷貨趨勢")
    st.dataframe(pd.DataFrame(monthly), use_container_width=True, hide_index=True)

    # ── Debug ─────────────────────────────────────────────────────────────────
    with st.expander("Debug：業績明細表欄位偵測"):
        st.json({
            "order_col": order_col, "actual_col": actual_col, "plan_col": plan_col,
            "ship_amt_col": ship_amt_col, "order_amt_col": order_amt_col,
            "customer_col": customer_col, "factory_col": factory_col, "wip_col": wip_col,
            "selected_month": str(selected),
            "teable_shipped": float(amount[is_shipment & actual_mask].fillna(0).sum()),
            "manual_shipped": float(manual_month.loc[manual_month["type"] == "已出貨", "amount_num"].fillna(0).sum()) if not manual_month.empty else 0,
            "shipped_total": shipped_total,
            "forecast_total": forecast_total,
            "month_total": month_total,
            "manual_table_id": _get_manual_table_id() or "(本機備援)",
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
