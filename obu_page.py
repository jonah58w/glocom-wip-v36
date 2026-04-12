# -*- coding: utf-8 -*-
"""
obu_page.py
報關單價核算 & 匯 HK OBU 金額 — Streamlit Tab 渲染模組

被 sales_report.py 的 render_sales_report_page() 呼叫。
資料來源：Teable 業績明細表 tblvy34oHUUw5PDlqd8
"""

from __future__ import annotations

from datetime import date
from typing import Optional

import pandas as pd
import requests
import streamlit as st

from customs_price import (
    load_price_db,
    decide_customs_price,
    confirm_and_save,
    calc_new_price,
    get_price_for_pn,
)

# ─────────────────────────────────────────
# 業績明細 / OBU 資料表設定
# ─────────────────────────────────────────
OBU_TABLE_ID = "tblvy34oHUUw5PDlqd8"
OBU_VIEW_ID  = "viwjlnlRnVPxN2P0D1K"
TEABLE_BASE  = "https://app.teable.io/api"   # 注意: teable.io (API) vs teable.ai (UI)

# ─────────────────────────────────────────
# 欄位候選名稱（依實際 Teable 欄位模糊比對）
# ─────────────────────────────────────────
_CAND = {
    "invoice":     ["Eusway Invoice", "西拓訂單編號", "Invoice", "PO#"],
    "pn":          ["P/N", "客戶料號", "Part No", "PN"],
    "qty":         ["Order Q'TY (PCS)", "Order Q'TY\n(PCS)", "Order Q'TY\n (PCS)",
                    "出貨數量", "Qty", "QTY"],
    "ship_date":   ["併貨日期", "出貨日期", "Ship date", "Ship Date"],
    "factory_ntd": ["工廠單價", "下單工廠金額(NTD)", "工廠金額"],
    "exrate":      ["海關匯率", "出貨匯率", "接單匯率"],
    "customer":    ["客戶", "Customer"],
    "random_usd":  ["美金出貨 (USD)", "接單金額 (USD)", "美金出貨(USD)",
                    "接單金額(USD)", "INVOICE", "Invoice Amount"],
    "tooling":     ["TOOLING", "Tooling"],
}


def _pick(df: pd.DataFrame, key: str) -> Optional[str]:
    """模糊比對欄位名稱"""
    if df is None or df.empty:
        return None
    norm = lambda s: str(s).strip().replace("\n", "").replace(" ", "").lower()
    col_map = {norm(c): c for c in df.columns}
    for cand in _CAND.get(key, []):
        nc = norm(cand)
        if nc in col_map:
            return col_map[nc]
    # 含substring match
    for cand in _CAND.get(key, []):
        nc = norm(cand)
        for orig_norm, orig in col_map.items():
            if nc and nc in orig_norm:
                return orig
    return None


def _to_float(v) -> float:
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return 0.0


def _dedup_cols(cols: list) -> list:
    """Rename duplicate column names: col, col_2, col_3 …"""
    seen: dict = {}
    out = []
    for c in cols:
        if c not in seen:
            seen[c] = 0
            out.append(c)
        else:
            seen[c] += 1
            out.append(f"{c}_{seen[c] + 1}")
    return out


def _fmt(v: float) -> str:
    return f"${v:,.2f}" if v else "—"


# ─────────────────────────────────────────
# 從 Teable 抓業績明細表資料
# ─────────────────────────────────────────
@st.cache_data(ttl=180, show_spinner=False)
def _fetch_obu_table(token: str) -> pd.DataFrame:
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url     = f"{TEABLE_BASE}/table/{OBU_TABLE_ID}/record"
    rows, skip = [], 0

    while True:
        params = {"viewId": OBU_VIEW_ID, "take": 500, "skip": skip,
                  "fieldKeyType": "name", "cellFormat": "text"}
        try:
            r = requests.get(url, headers=headers, params=params, timeout=20)
            r.raise_for_status()
        except Exception as e:
            st.error(f"OBU 資料載入失敗：{e}")
            return pd.DataFrame()

        data    = r.json()
        records = data.get("records", [])
        if not records:
            break
        for rec in records:
            row = rec.get("fields", {})
            row["_id"] = rec.get("id", "")
            rows.append(row)
        if len(records) < 500:
            break
        skip += 500

    df = pd.DataFrame(rows)
    if not df.empty:
        df.columns = _dedup_cols([str(c).strip() for c in df.columns])
    return df


def _get_token() -> str:
    try:
        return st.secrets.get("TEABLE_TOKEN", "")
    except Exception:
        return ""


# ─────────────────────────────────────────
# 月份篩選器（回傳篩選後的 DataFrame）
# ─────────────────────────────────────────
def _month_filtered(df: pd.DataFrame, key_suffix: str = "") -> pd.DataFrame:
    date_col = _pick(df, "ship_date")
    if not date_col:
        # show a hint about what columns exist
        avail = [c for c in df.columns if not c.startswith("_")]
        st.caption(f"⚠️ 找不到「併貨日期」欄，顯示全部。現有欄位前10個：{avail[:10]}")
        return df

    df = df.copy()
    df["_date_parsed"] = pd.to_datetime(df[date_col], errors="coerce")
    months = (
        df["_date_parsed"].dropna()
        .dt.to_period("M").unique()
        .sort_values(ascending=False)
    )
    opts = [str(m) for m in months]
    if not opts:
        return df

    sel = st.selectbox("📅 選擇月份", opts, index=0, key=f"obu_month_{key_suffix}")
    return df[df["_date_parsed"].dt.to_period("M").astype(str) == sel].copy()


# ─────────────────────────────────────────
# Tab 2：報關單價核算
# ─────────────────────────────────────────
def render_customs_price_tab():
    st.markdown("##### 🏷️ 報關單價核算")
    st.caption(
        "依工廠單價 × (1 + 6%) ÷ 海關匯率計算建議報關單價，"
        "與上次記錄比對後決定用舊或新單價。確認後寫入本機 `customs_prices.json`。"
    )

    token = _get_token()
    if not token:
        st.error("未設定 TEABLE_TOKEN")
        return

    with st.spinner("載入業績明細資料…"):
        df_raw = _fetch_obu_table(token)

    if df_raw.empty:
        st.warning("業績明細表無資料")
        return

    df = _month_filtered(df_raw, "customs")

    inv_col  = _pick(df, "invoice")
    pn_col   = _pick(df, "pn")
    fp_col   = _pick(df, "factory_ntd")
    er_col   = _pick(df, "exrate")
    date_col = _pick(df, "ship_date")

    # ── 只看 ET 出貨 ──────────────────────
    if inv_col:
        et_df = df[df[inv_col].astype(str).str.upper().str.startswith("ET")].copy()
    else:
        st.warning("找不到 Invoice 欄（Eusway Invoice / 西拓訂單編號），顯示全部")
        et_df = df.copy()

    if et_df.empty:
        st.info("本期無 ET 出貨")
        return

    db = load_price_db()

    # deduplicate by P/N（取最新出貨日）
    if date_col and date_col in et_df.columns:
        et_df = et_df.sort_values(date_col, ascending=False)
    pn_rows = et_df.drop_duplicates(subset=[pn_col]).copy() if pn_col else et_df

    st.caption(f"共 {len(pn_rows)} 個不重複料號（ET 帳戶）")

    updated = False
    for _, row in pn_rows.iterrows():
        pn          = str(row.get(pn_col, "")).strip() if pn_col else ""
        factory_ntd = _to_float(row.get(fp_col, 0)) if fp_col else 0.0
        ex_rate     = _to_float(row.get(er_col, 0)) if er_col else 0.0
        ship_date   = row.get(date_col) if date_col else None

        if not pn:
            continue

        # ── 無工廠單價或匯率 → 顯示警告 ──
        if not factory_ntd or not ex_rate:
            with st.expander(f"⚠️ **{pn}** — 工廠單價或匯率缺失"):
                st.caption(f"工廠單價：{factory_ntd}  匯率：{ex_rate}")
                old = db.get(pn, {}).get("price")
                if old:
                    st.info(f"使用上次記錄單價 **${old:.4f}**")
                else:
                    st.warning("無法計算，也無歷史記錄")
            continue

        suggested, is_new, reason = decide_customs_price(
            pn, factory_ntd, ex_rate, ship_date, db
        )
        new_calc  = calc_new_price(factory_ntd, ex_rate)
        old_price = db.get(pn, {}).get("price")

        badge = "🆕" if is_new else "📌"
        with st.expander(
            f"{badge} **{pn}**  —  建議：`${suggested:.4f}`  "
            f"{'**(更新)**' if is_new else '(沿用舊)'}",
            expanded=bool(is_new),
        ):
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("工廠單價 (NTD)", f"{factory_ntd:,.2f}")
            c2.metric("海關匯率", f"{ex_rate:,.3f}")
            c3.metric("6% 新單價", f"${new_calc:.4f}" if new_calc else "—")
            c4.metric("上次報關單價", f"${old_price:.4f}" if old_price else "—")
            st.caption(f"📝 {reason}")

            manual = st.number_input(
                "確認 / 覆蓋報關單價 (USD)",
                value=float(suggested or 0),
                step=0.0001, format="%.4f",
                key=f"cp_price_{pn}",
            )
            if st.button(f"✅ 儲存 {pn}", key=f"cp_save_{pn}"):
                confirm_and_save(pn, manual, ship_date, factory_ntd, ex_rate, reason)
                updated = True
                st.success(f"已儲存 **{pn}** = **${manual:.4f}**")

    if updated:
        _fetch_obu_table.clear()
        st.info("✅ 報關單價已更新，OBU 計算頁將自動反映新單價。")

    # ── 現有單價資料庫 ───────────────────────
    st.divider()
    st.markdown("**📚 報關單價資料庫（現有記錄）**")
    db_now = load_price_db()
    if db_now:
        db_df = pd.DataFrame([{"P/N": k, **v} for k, v in db_now.items()])
        db_df = db_df.rename(columns={
            "price":             "報關單價(USD)",
            "last_change_date":  "上次變更日期",
            "factory_price_ntd": "工廠單價(NTD)",
            "exchange_rate_used":"使用匯率",
            "reason":            "決策原因",
            "updated_at":        "更新日",
        })
        st.dataframe(db_df, use_container_width=True, height=260,
                     column_config={
                         "報關單價(USD)": st.column_config.NumberColumn(format="%.4f"),
                     })
    else:
        st.info("尚無記錄，請先確認各料號單價")


# ─────────────────────────────────────────
# Tab 3：匯 HK OBU 金額
# ─────────────────────────────────────────
def render_obu_calc_tab():
    st.markdown("##### 🏦 匯 HK OBU 金額統計")
    st.caption("只計算 ET 帳戶出貨；報關金額 = 出貨數量 × 報關單價（來自 customs_prices.json）。")

    token = _get_token()
    if not token:
        st.error("未設定 TEABLE_TOKEN")
        return

    with st.spinner("載入業績明細資料…"):
        df_raw = _fetch_obu_table(token)

    if df_raw.empty:
        st.warning("業績明細表無資料")
        return

    df = _month_filtered(df_raw, "obu")

    # ── 前期補差 ──────────────────────────
    col_a, col_b = st.columns([1, 2])
    with col_a:
        carryover = st.number_input(
            "前期少匯補差 (USD，正=需補匯，負=可折抵)",
            value=0.0, step=0.5, format="%.2f", key="obu_carryover"
        )
    with col_b:
        carryover_note = st.text_input(
            "補差說明",
            placeholder="例：3月匯$11,050 - 應付$777 - 應付$10,416.5 = -$143.5",
            key="obu_carryover_note"
        )

    st.divider()

    # ── 欄位 ──────────────────────────────
    inv_col    = _pick(df, "invoice")
    pn_col     = _pick(df, "pn")
    qty_col    = _pick(df, "qty")
    date_col   = _pick(df, "ship_date")
    cust_col   = _pick(df, "customer")
    rand_col   = _pick(df, "random_usd")
    er_col     = _pick(df, "exrate")

    # ── Debug：欄位對應 ───────────────────
    with st.expander("🔍 Debug — 欄位對應（無資料時展開查看）", expanded=(inv_col is None or pn_col is None)):
        st.markdown("**Teable 實際欄位清單：**")
        cols_display = [c for c in df.columns if not c.startswith("_")]
        st.code("\n".join(cols_display))
        st.markdown("**目前比對結果：**")
        st.json({
            "invoice (ET/EW/GC)": inv_col,
            "P/N":                pn_col,
            "出貨數量":           qty_col,
            "併貨/出貨日期":      date_col,
            "客戶":               cust_col,
            "隨機金額(USD)":      rand_col,
            "海關匯率":           er_col,
        })
        st.caption("若 invoice 欄為 None，ET 篩選會失效；若 date_col 為 None，月份篩選會失效。")
        if not df.empty:
            st.markdown("**前 3 筆原始資料：**")
            st.dataframe(df.head(3), use_container_width=True)

    db = load_price_db()

    # ── 計算每筆報關金額 ───────────────────
    def _row_customs(row):
        pn    = str(row.get(pn_col, "")).strip() if pn_col else ""
        qty   = _to_float(row.get(qty_col, 0)) if qty_col else 0.0
        price = get_price_for_pn(pn) if pn else None
        if price and qty:
            return round(qty * price, 4)
        return None

    df["報關單價"]  = df.apply(
        lambda r: get_price_for_pn(str(r.get(pn_col, "")).strip()) if pn_col else None,
        axis=1
    )
    df["報關金額"] = df.apply(_row_customs, axis=1)

    # ── ET 篩選 ───────────────────────────
    if inv_col:
        is_et = df[inv_col].astype(str).str.upper().str.startswith("ET")
        et_df = df[is_et].copy()
    else:
        et_df = df.copy()

    # ── ① 各料號出貨明細 ──────────────────
    st.markdown("**① 各料號出貨金額（報關單價 × 出貨數量）**")
    # deduplicate: keep first occurrence of each column name
    show_cols_i = list(dict.fromkeys(
        c for c in [date_col, inv_col, cust_col, pn_col,
                    qty_col, rand_col, "報關單價", "報關金額", er_col]
        if c and c in et_df.columns
    ))
    if not et_df.empty:
        display_i = et_df[show_cols_i].copy()
        # reset any duplicate column names that may exist in the raw Teable data
        display_i.columns = _dedup_cols(list(display_i.columns))
        if date_col and date_col in display_i.columns:
            display_i = display_i.sort_values(date_col, ascending=False)
        st.dataframe(
            display_i, use_container_width=True, height=260,
            column_config={
                "報關金額":  st.column_config.NumberColumn("報關金額 (USD)", format="%.2f"),
                "報關單價":  st.column_config.NumberColumn("報關單價 (USD)", format="%.4f"),
            }
        )
        # 未設定報關單價的警示
        missing_pn = et_df[et_df["報關單價"].isna()][pn_col].dropna().unique().tolist() \
            if pn_col else []
        if missing_pn:
            st.warning(
                f"⚠️ {len(missing_pn)} 個料號尚無報關單價，請至「報關單價」Tab 設定：\n"
                + "、".join(missing_pn[:8]) + ("…" if len(missing_pn) > 8 else "")
            )
    else:
        st.info("本期無 ET 出貨")
        return

    # ── ② 依併貨日期分組 ──────────────────
    st.markdown("**② 依併貨日期分組 — 各批次報關金額合計**")
    if date_col and date_col in et_df.columns:
        by_date = (
            et_df.groupby(date_col, dropna=False)
            .agg(
                筆數=("報關金額", "count"),
                報關金額合計=("報關金額", "sum"),
            )
            .reset_index()
            .sort_values(date_col)
        )
        st.dataframe(
            by_date, use_container_width=True,
            height=min(220, 60 + len(by_date) * 36),
            column_config={
                "報關金額合計": st.column_config.NumberColumn("報關金額合計 (USD)", format="%.2f"),
            }
        )
    else:
        st.caption("找不到併貨日期欄，略過分組")
        by_date = pd.DataFrame()

    # ── ③ 最終應匯金額 ────────────────────
    st.markdown("**③ 本次需匯 HK OBU 總金額計算**")
    et_total  = float(et_df["報關金額"].sum())
    final     = round(et_total + carryover, 2)

    c1, c2, c3 = st.columns(3)
    c1.metric("EUSWAY 報關總金額（ET）", _fmt(et_total))
    sign = "+" if carryover >= 0 else ""
    c2.metric("前期補差", f"{sign}{_fmt(carryover)}", help=carryover_note or "")
    c3.metric("✅ 最終應匯 HK OBU 金額", _fmt(final))

    # ── 附：帳戶分類 ──────────────────────
    if inv_col and rand_col and rand_col in df.columns:
        st.markdown("---")
        st.markdown("**附：4 月出貨累計（依帳戶）**")
        df["_acct"] = df[inv_col].astype(str).str[:2].str.upper()
        acct_sum = (
            df.groupby("_acct")[rand_col]
            .apply(lambda s: s.apply(_to_float).sum())
            .reset_index()
            .rename(columns={rand_col: "隨機合計(USD)"})
        )
        st.dataframe(acct_sum, use_container_width=True, height=160,
                     column_config={
                         "隨機合計(USD)": st.column_config.NumberColumn(format="%.2f")
                     })
