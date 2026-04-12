# -*- coding: utf-8 -*-
"""
obu_page.py  v2
報關單價核算 & 匯 HK OBU 金額 — 依 Teable 實際欄位重新對應

Teable 業績明細表 (tblvy34oHUUw5PDlqd8) 實際欄位：
  日期 / Eusway Invoice / 客戶 / 隨機金額USD / Tooling USD /
  報關金額USD / 匯率 / 差額TWD / 內帳銷貨金額TWD / 外帳銷貨金額TWD /
  月份 / Invoice類型
"""

from __future__ import annotations
from typing import Optional

import pandas as pd
import requests
import streamlit as st

from customs_price import load_price_db   # 僅用於 Tab 2 顯示資料庫

# ─────────────────────────────────────────
# Teable 設定
# ─────────────────────────────────────────
OBU_TABLE_ID = "tblvy34oHUUw5PDlqd8"
OBU_VIEW_ID  = "viwjlnlRnVPxN2P0D1K"
TEABLE_BASE  = "https://app.teable.io/api"

# ─────────────────────────────────────────
# 欄位名稱候選（依 debug 確認的實際欄位）
# ─────────────────────────────────────────
_CAND = {
    "invoice":      ["Eusway Invoice", "西拓訂單編號", "Invoice"],
    "ship_date":    ["日期", "併貨日期", "出貨日期", "Ship date", "Ship Date"],
    "month_col":    ["月份"],
    "random_usd":   ["隨機金額USD", "隨機金額(USD)", "隨機", "美金出貨(USD)"],
    "tooling":      ["Tooling USD", "TOOLING", "Tooling"],
    "customs_amt":  ["報關金額USD", "報關金額(USD)", "報關金額"],
    "exrate":       ["匯率", "海關匯率", "出貨匯率", "接單匯率"],
    "customer":     ["客戶", "Customer"],
    "invoice_type": ["Invoice類型", "Invoice Type"],
    "diff_twd":     ["差額TWD"],
    "inner_twd":    ["內帳銷貨金額TWD"],
    "outer_twd":    ["外帳銷貨金額TWD"],
}


# ─────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────

def _pick(df: pd.DataFrame, key: str) -> Optional[str]:
    if df is None or df.empty:
        return None
    norm = lambda s: str(s).strip().replace("\n", "").replace(" ", "").lower()
    col_map = {norm(c): c for c in df.columns}
    for cand in _CAND.get(key, []):
        if norm(cand) in col_map:
            return col_map[norm(cand)]
    for cand in _CAND.get(key, []):
        nc = norm(cand)
        for on, oc in col_map.items():
            if nc and nc in on:
                return oc
    return None


def _to_float(v) -> float:
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return 0.0


def _dedup_cols(cols: list) -> list:
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


def _fmt(v) -> str:
    try:
        return f"${float(v):,.2f}"
    except Exception:
        return "—"


def _get_token() -> str:
    try:
        return st.secrets.get("TEABLE_TOKEN", "")
    except Exception:
        return ""


# ─────────────────────────────────────────
# Teable fetch
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


# ─────────────────────────────────────────
# 月份篩選
# ─────────────────────────────────────────

def _normalize_ym(v: str) -> str:
    """把各種月份格式統一成 YYYY-MM，失敗回傳空字串。"""
    v = str(v).strip()
    if not v or v in ("nan", "None", "NaT"):
        return ""
    # 已是 YYYY-MM
    import re
    m = re.match(r"^(\d{4})[/-](\d{1,2})$", v)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    # YYYY-MM-DD → take first 7
    m = re.match(r"^(\d{4})[/-](\d{1,2})[/-]\d", v)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    # YYYY年MM月
    m = re.match(r"^(\d{4})年(\d{1,2})月", v)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    # 前7字符 fallback
    if len(v) >= 7 and v[4] in "-/":
        return v[:7]
    return ""


def _month_filtered(df: pd.DataFrame, key_suffix: str = "") -> tuple:
    """Returns (filtered_df, selected_month_str).
    預設永遠選目前年月；若當月無資料則仍顯示當月（空資料）。
    """
    month_col = _pick(df, "month_col")
    date_col  = _pick(df, "ship_date")
    today_ym  = pd.Timestamp.today().strftime("%Y-%m")

    months_set: set = set()

    if month_col:
        for v in df[month_col].dropna().astype(str):
            ym = _normalize_ym(v)
            if ym:
                months_set.add(ym)

    if not months_set and date_col:
        parsed = pd.to_datetime(df[date_col], errors="coerce")
        for p in parsed.dropna().dt.to_period("M").unique():
            months_set.add(str(p))

    # 無論如何，確保當月永遠在選單裡
    months_set.add(today_ym)

    months  = sorted(months_set, reverse=True)
    default = months.index(today_ym)   # 一定存在

    sel = st.selectbox("📅 選擇月份", months, index=default,
                       key=f"obu_month_{key_suffix}")

    if month_col:
        # 先 normalize 再比對
        norm_series = df[month_col].astype(str).apply(_normalize_ym)
        filtered = df[norm_series == sel].copy()
    elif date_col:
        parsed = pd.to_datetime(df[date_col], errors="coerce")
        filtered = df[parsed.dt.to_period("M").astype(str) == sel].copy()
    else:
        filtered = df.copy()

    return filtered, sel


# ─────────────────────────────────────────
# Tab 2：報關單價核算（Invoice 層級）
# ─────────────────────────────────────────

def render_customs_price_tab():
    st.markdown("##### 🏷️ 報關單價 — Invoice 報關比率分析")
    st.caption(
        "業績明細表為 Invoice 層級，顯示各 ET 發票的隨機金額與報關金額對比及報關比率。"
        "歷史 P/N 層級報關單價見下方資料庫。"
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

    df, sel_month = _month_filtered(df_raw, "customs")

    inv_col   = _pick(df, "invoice")
    cust_col  = _pick(df, "customer")
    rand_col  = _pick(df, "random_usd")
    cust_amt  = _pick(df, "customs_amt")
    er_col    = _pick(df, "exrate")
    date_col  = _pick(df, "ship_date")

    # ET 篩選
    if inv_col:
        et_df = df[df[inv_col].astype(str).str.upper().str.startswith("ET")].copy()
    else:
        et_df = df.copy()

    if et_df.empty:
        st.info(f"本期（{sel_month}）無 ET 出貨")
    else:
        rand_s   = et_df[rand_col].apply(_to_float) if rand_col else pd.Series(0.0, index=et_df.index)
        custom_s = et_df[cust_amt].apply(_to_float) if cust_amt else pd.Series(0.0, index=et_df.index)
        et_df["報關比率%"] = (custom_s / rand_s * 100).round(1).where(rand_s > 0, None)

        show_cols = list(dict.fromkeys(c for c in [
            date_col, inv_col, cust_col, rand_col, cust_amt, "報關比率%", er_col
        ] if c and c in et_df.columns))

        st.markdown(f"**{sel_month} ET 出貨 — 報關金額明細**（共 {len(et_df)} 筆）")
        cfg = {}
        if rand_col:
            cfg[rand_col] = st.column_config.NumberColumn("隨機金額(USD)", format="%.2f")
        if cust_amt:
            cfg[cust_amt] = st.column_config.NumberColumn("報關金額(USD)", format="%.2f")
        cfg["報關比率%"] = st.column_config.NumberColumn("報關比率%", format="%.1f")

        st.dataframe(
            et_df[show_cols].reset_index(drop=True),
            use_container_width=True, height=320, hide_index=True,
            column_config=cfg
        )

        c1, c2, c3 = st.columns(3)
        c1.metric("ET 隨機金額合計",  _fmt(rand_s.sum()))
        c2.metric("ET 報關金額合計",  _fmt(custom_s.sum()))
        avg_r = (custom_s.sum() / rand_s.sum() * 100) if rand_s.sum() > 0 else 0
        c3.metric("平均報關比率", f"{avg_r:.1f}%")

    # 歷史 P/N 報關單價資料庫
    st.divider()
    st.markdown("**📚 報關單價資料庫（customs_prices.json）**")
    db = load_price_db()
    if db:
        db_df = pd.DataFrame([{"P/N": k, **v} for k, v in db.items()])
        db_df = db_df.rename(columns={
            "price":             "報關單價(USD)",
            "last_change_date":  "上次變更日期",
            "factory_price_ntd": "工廠單價(NTD)",
            "exchange_rate_used":"使用匯率",
            "reason":            "決策原因",
            "updated_at":        "更新日",
        })
        st.dataframe(db_df, use_container_width=True, height=260,
                     column_config={"報關單價(USD)": st.column_config.NumberColumn(format="%.4f")})
    else:
        st.info("尚無 P/N 層級報關單價記錄")


# ─────────────────────────────────────────
# Tab 3：匯 HK OBU 金額統計
# ─────────────────────────────────────────

def render_obu_calc_tab():
    st.markdown("##### 🏦 匯 HK OBU 金額統計")
    st.caption("只計算 ET 帳戶出貨。報關金額直接使用 Teable「報關金額USD」欄。")

    token = _get_token()
    if not token:
        st.error("未設定 TEABLE_TOKEN")
        return

    with st.spinner("載入業績明細資料…"):
        df_raw = _fetch_obu_table(token)

    if df_raw.empty:
        st.warning("業績明細表無資料")
        return

    # ── 月份選擇（預設當月）─────────────────
    df, sel_month = _month_filtered(df_raw, "obu")

    # ── 欄位偵測 ─────────────────────────
    inv_col   = _pick(df, "invoice")
    date_col  = _pick(df, "ship_date")
    cust_col  = _pick(df, "customer")
    rand_col  = _pick(df, "random_usd")
    cust_amt  = _pick(df, "customs_amt")
    er_col    = _pick(df, "exrate")
    tool_col  = _pick(df, "tooling")
    inner_col = _pick(df, "inner_twd")
    outer_col = _pick(df, "outer_twd")

    if not cust_amt:
        st.error(
            "找不到「報關金額USD」欄位，無法計算。\n"
            f"目前欄位：{[c for c in df.columns if not c.startswith('_')]}"
        )
        return

    # ── ET 篩選 & 計算總額 ───────────────
    if inv_col:
        is_et = df[inv_col].astype(str).str.upper().str.startswith("ET")
        et_df = df[is_et].copy()
    else:
        st.warning("找不到 Eusway Invoice 欄，顯示全部資料")
        et_df = df.copy()

    et_df["_customs"] = et_df[cust_amt].apply(_to_float)
    et_total = float(et_df["_customs"].sum()) if not et_df.empty else 0.0

    # ══════════════════════════════════════
    # 頂部：本月 HK OBU 匯款總覽（置頂顯示）
    # ══════════════════════════════════════
    st.markdown(
        f"""
        <div style="
            background:linear-gradient(135deg,#1a2744,#162035);
            border:1px solid #2e4070;border-radius:14px;
            padding:20px 24px;margin-bottom:16px;">
          <div style="font-size:.85rem;color:#7a9cc8;letter-spacing:.06em;
                      text-transform:uppercase;margin-bottom:10px;">
            🏦 {sel_month} HK OBU 匯款總覽
          </div>
          <div style="display:grid;grid-template-columns:1fr auto 1fr auto 1fr;
                      align-items:center;gap:8px;">
            <div>
              <div style="font-size:.75rem;color:#6b8ab0;margin-bottom:4px;">
                EUSWAY 報關總金額（ET）
              </div>
              <div style="font-size:1.7rem;font-weight:700;color:#e8f0ff;">
                {_fmt(et_total)}
              </div>
            </div>
            <div style="font-size:1.5rem;color:#4a6a9a;padding:0 4px;">+</div>
            <div id="carryover-display">
              <div style="font-size:.75rem;color:#6b8ab0;margin-bottom:4px;">
                前期補差
              </div>
              <div style="font-size:1.7rem;font-weight:700;color:#a8c0e8;">
                ✏️ 見下方輸入
              </div>
            </div>
            <div style="font-size:1.5rem;color:#4a6a9a;padding:0 4px;">=</div>
            <div>
              <div style="font-size:.75rem;color:#6b8ab0;margin-bottom:4px;">
                ✅ 最終應匯 HK OBU
              </div>
              <div style="font-size:1.7rem;font-weight:700;color:#f0c040;">
                ✏️ 見下方
              </div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── 前期補差輸入 ──────────────────────
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

    # ── 確認後的最終總計（大字顯示） ─────
    final = round(et_total + carryover, 2)
    sign  = "+" if carryover >= 0 else ""

    c1, c2, c3 = st.columns(3)
    c1.metric("EUSWAY 報關總金額（ET）", _fmt(et_total),
              delta=f"{len(et_df)} 筆" if not et_df.empty else "0 筆")
    c2.metric("前期補差", f"{sign}{_fmt(carryover)}",
              help=carryover_note or "尚未填寫")
    c3.metric("✅ 最終應匯 HK OBU 金額", _fmt(final))
    if carryover_note:
        st.caption(f"📝 {carryover_note}")

    st.divider()

    # ── ① 各發票明細 ─────────────────────
    st.markdown("**① 各發票出貨金額（ET 帳戶）**")
    if et_df.empty:
        st.info(f"本期（{sel_month}）無 ET 出貨")
    else:
        show_cols = list(dict.fromkeys(c for c in [
            date_col, inv_col, cust_col, rand_col, tool_col, cust_amt, er_col
        ] if c and c in et_df.columns))

        disp = et_df[show_cols].reset_index(drop=True).copy()
        if date_col and date_col in disp.columns:
            disp = disp.sort_values(date_col, ascending=False)

        cfg = {}
        if cust_amt:
            cfg[cust_amt] = st.column_config.NumberColumn("報關金額(USD)", format="%.2f")
        if rand_col:
            cfg[rand_col] = st.column_config.NumberColumn("隨機金額(USD)", format="%.2f")

        st.dataframe(disp, use_container_width=True, height=280,
                     hide_index=True, column_config=cfg)

        zero_inv = et_df[et_df["_customs"] == 0][inv_col].tolist() if inv_col else []
        if zero_inv:
            st.caption(
                f"ℹ️ 報關金額為 0 的發票（{len(zero_inv)} 筆）："
                f"{', '.join(str(x) for x in zero_inv[:6])}"
            )

    # ── ② 依日期分組 ─────────────────────
    st.markdown("**② 依出貨日期分組 — 各批次報關金額合計**")
    if date_col and date_col in et_df.columns and not et_df.empty:
        by_date = (
            et_df.groupby(date_col, dropna=False)
            .agg(筆數=("_customs", "count"), 報關金額合計=("_customs", "sum"))
            .reset_index()
            .sort_values(date_col)
        )
        st.dataframe(
            by_date, use_container_width=True,
            height=min(220, 60 + len(by_date) * 36),
            hide_index=True,
            column_config={
                "報關金額合計": st.column_config.NumberColumn("報關金額合計(USD)", format="%.2f"),
            }
        )
    elif not et_df.empty:
        st.caption("找不到日期欄，略過分組")

    # ── 附：帳戶分類統計 ──────────────────
    if inv_col and rand_col and rand_col in df.columns and not df.empty:
        st.markdown("---")
        st.markdown(f"**附：{sel_month} 出貨累計（依帳戶）**")
        df2 = df.copy()
        df2["_acct"] = df2[inv_col].astype(str).str[:2].str.upper()
        df2["_rand"] = df2[rand_col].apply(_to_float)
        df2["_cust"] = df2[cust_amt].apply(_to_float)
        acct_sum = (
            df2.groupby("_acct")
            .agg(隨機合計=("_rand", "sum"), 報關合計=("_cust", "sum"))
            .reset_index()
            .rename(columns={"_acct": "帳戶"})
        )
        st.dataframe(
            acct_sum, use_container_width=True, height=160, hide_index=True,
            column_config={
                "隨機合計": st.column_config.NumberColumn("隨機金額合計(USD)", format="%.2f"),
                "報關合計": st.column_config.NumberColumn("報關金額合計(USD)", format="%.2f"),
            }
        )

    # ── 完整明細（含所有帳戶）─────────────
    with st.expander("📋 本月完整明細（所有帳戶）", expanded=False):
        if not df.empty:
            all_cols = list(dict.fromkeys(c for c in [
                date_col, inv_col, cust_col, rand_col, tool_col,
                cust_amt, er_col, inner_col, outer_col
            ] if c and c in df.columns))
            cfg2 = {}
            if cust_amt:
                cfg2[cust_amt]  = st.column_config.NumberColumn("報關金額(USD)", format="%.2f")
            if rand_col:
                cfg2[rand_col]  = st.column_config.NumberColumn("隨機金額(USD)", format="%.2f")
            if inner_col:
                cfg2[inner_col] = st.column_config.NumberColumn("內帳銷貨(TWD)", format="%.0f")
            if outer_col:
                cfg2[outer_col] = st.column_config.NumberColumn("外帳銷貨(TWD)", format="%.0f")

            st.dataframe(df[all_cols].reset_index(drop=True),
                         use_container_width=True, height=380,
                         hide_index=True, column_config=cfg2)
            csv = df[all_cols].to_csv(index=False).encode("utf-8-sig")
            st.download_button("📥 下載 CSV", csv,
                               f"obu_{sel_month}.csv", "text/csv")
