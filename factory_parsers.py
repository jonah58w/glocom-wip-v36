# -*- coding: utf-8 -*-
"""
factory_parsers.py  [修正版]

修正內容：
1. 新增 parse_xituo_wip_report()      — 西拓-WIP.xls（料號+製程欄，無 PO#）
2. 新增 parse_203_xituo_report()      — 203-西拓 Wip.xls（雙表頭、交期 0407=>0414 格式）
3. 新增 parse_xituo_simple_report()   — 西拓電子有限公司進度表.xlsx（有「進度」文字欄）
4. 修正 looks_like_xitop_workflow()   — 避免誤判西拓-WIP（無 P/O、無工作流程計劃）
5. 修正 looks_like_203_xituo()        — 正確辨識 203 格式
6. 修正 parse_profit_grand()          — 正確用 'LS' sheet
7. 修正 read_import_dataframe()       — 加入新 parser 的路由
"""

from __future__ import annotations

import re
import pandas as pd
from typing import Any, Dict, List, Optional, Tuple

from excel_reader import (
    read_first_nonempty_sheet_raw,
    read_first_nonempty_sheet_with_header,
)
from utils import safe_text, normalize_columns, compact_text, normalize_due_text
from text_ocr_parsers import parse_email_text_to_rows


# =========================================================
# 標準欄位候選（同原版）
# =========================================================
PO_CANDIDATES = [
    "PO#", "PO", "P/O", "P O", "訂單編號", "訂單號", "訂單號碼",
    "工單", "工單號", "單號", "ORDER NO", "Order No", "Order Number"
]
PART_CANDIDATES = [
    "Part No", "Part No.", "P/N", "PN", "料號", "品號", "客戶料號",
    "Cust. P / N", "LS P/N", "客戶品號", "成品料號", "產品料號",
    "產品編號", "Product No", "Model", "祥竑料號"
]
QTY_CANDIDATES = [
    "Qty", "QTY", "Q'TY", "Order Q'TY (PCS)", "Order Q'TY\n (PCS)",
    "訂購量(PCS)", "訂購量", "訂單量(PCS)", "訂單量", "數量", "數量(PCS)",
    "PCS", "生產數量", "投產數", "訂單數量", "未出貨數量"
]
DUE_CANDIDATES = [
    "Factory Due Date", "工廠交期", "交貨日期", "交期", "出貨日期",
    "Required Ship date", "Required Ship Date", "confirmed DD", "confrimed DD",
    "預交日", "預定交期", "交貨期", "出貨日"
]
SHIP_DATE_CANDIDATES = [
    "Ship Date", "Ship date", "出貨日期", "交貨日期",
    "Required Ship date", "Required Ship Date", "confirmed DD", "confrimed DD"
]
WIP_CANDIDATES = [
    "WIP", "WIP Stage", "進度", "製程", "工序", "目前站別", "生產進度"
]
REMARK_CANDIDATES = [
    "Remark", "備註", "情況", "備註說明", "Note", "說明", "異常備註"
]

PROCESS_ORDER_GENERIC = [
    "發料", "下料", "排版", "內層", "內乾", "內蝕", "黑化", "壓合", "壓板",
    "鑽孔", "鉆孔", "沉銅", "一銅", "電鍍", "乾膜", "外層", "二銅", "二銅蝕刻",
    "AOI", "半測", "防焊", "文字", "噴錫", "化金", "化/鍍金", "OSP", "化銀",
    "成型", "V-CUT", "測試", "成檢", "包裝", "出貨", "庫存",
    "蝕刻", "濕膜", "曝光", "品檢", "印可剝膠", "接收", "ENTEK",
]


# =========================================================
# 共用工具（同原版）
# =========================================================
def _first_existing(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _make_result_row(
    po="", part="", qty="", due="", ship="",
    wip="", remark="", tags=None,
    source_sheet="", source_type="",
) -> Dict[str, Any]:
    return {
        "PO#": safe_text(po),
        "Part No": safe_text(part),
        "Qty": safe_text(qty),
        "Factory Due Date": normalize_due_text(due),
        "Ship Date": normalize_due_text(ship if ship else due),
        "WIP": safe_text(wip),
        "Remark": safe_text(remark),
        "Customer Remark Tags": tags or [],
        "_source_sheet": source_sheet or "",
        "_source_type": source_type or "",
    }


def _detect_process_columns(df: pd.DataFrame) -> List[str]:
    matched = []
    for col in df.columns:
        c = compact_text(col)
        if c in PROCESS_ORDER_GENERIC or any(k in c for k in PROCESS_ORDER_GENERIC):
            matched.append(col)
    return matched


def _guess_wip_from_last_process(row: pd.Series, process_cols: List[str]) -> Tuple[str, str]:
    """回傳 (wip_label, last_step_name)"""
    last_step = ""
    for p in process_cols:
        value = str(row.get(p, "")).strip()
        if value not in {"", "nan", "None", "none", "null", "-", "--", "0", "0.0", "NaN"}:
            last_step = p

    if not last_step:
        return "", ""

    if last_step in {"出貨", "接收"}:
        return "Shipping", last_step
    if last_step == "包裝":
        return "Packing", last_step
    if last_step in {"測試", "成檢", "成檢(1)", "成檢(2)", "FQC", "QA", "品檢"}:
        return "Inspection", last_step
    return "Production", last_step


# =========================================================
# [修正 1] 西拓-WIP parser
# 特徵：欄位有「料號」「交期」「NO」，無 PO# 欄位
# 進度：找最後一個有數值的製程欄
# =========================================================
def looks_like_xituo_wip(raw_df: pd.DataFrame) -> bool:
    """西拓-WIP.xls：有「西拓 WIP 進度表」標題，欄位含料號+交期+製程"""
    if raw_df.empty:
        return False
    text = "".join(
        "".join([compact_text(x) for x in raw_df.iloc[i].tolist()])
        for i in range(min(len(raw_df), 5))
    )
    return (
        "西拓WIP進度表" in text or "西拓WIP" in text
    ) and "料號" in text and "交期" in text


def parse_xituo_wip_report(uploaded_file) -> pd.DataFrame:
    """
    西拓-WIP.xls 格式
    Row 0: 標題 '西拓 WIP 進度表'
    Row 1: 欄位名稱 NO | 料號 | 訂單量(PCS) | 交期 | 備註 | 下料 | 內乾 | ...
    """
    raw_df, sheet_name = read_first_nonempty_sheet_raw(uploaded_file)
    if raw_df.empty:
        raise ValueError("西拓-WIP 讀取失敗：工作表為空")

    # 找 header row（含「料號」且含「交期」的那一行）
    header_row = None
    for i in range(min(len(raw_df), 6)):
        row_text = "".join([compact_text(x) for x in raw_df.iloc[i].tolist()])
        if "料號" in row_text and "交期" in row_text:
            header_row = i
            break

    if header_row is None:
        raise ValueError("西拓-WIP：找不到表頭列")

    df = pd.read_excel(
        uploaded_file if hasattr(uploaded_file, "read") else uploaded_file,
        sheet_name=sheet_name,
        header=header_row,
    )
    df = normalize_columns(df)
    df = df.dropna(how="all").reset_index(drop=True)

    part_col = _first_existing(df, ["料號"] + PART_CANDIDATES)
    qty_col = _first_existing(df, ["訂單量(PCS)", "訂單量"] + QTY_CANDIDATES)
    due_col = _first_existing(df, ["交期"] + DUE_CANDIDATES)
    remark_col = _first_existing(df, ["備註"] + REMARK_CANDIDATES)

    # 製程欄：去掉已知非製程欄
    non_proc = {"NO", "料號", "訂單量(PCS)", "訂單量", "交期", "備註"}
    process_cols = [
        c for c in df.columns
        if c not in non_proc and compact_text(c) in PROCESS_ORDER_GENERIC
    ]
    # 也包含部分匹配
    if not process_cols:
        process_cols = [
            c for c in df.columns
            if c not in non_proc
            and any(k in compact_text(c) for k in PROCESS_ORDER_GENERIC)
        ]

    rows = []
    for _, row in df.iterrows():
        part_val = safe_text(row.get(part_col, "")) if part_col else ""
        if not part_val or part_val.lower() in {"nan", "none"}:
            continue

        wip_label, last_step = _guess_wip_from_last_process(row, process_cols)
        remark_val = safe_text(row.get(remark_col, "")) if remark_col else ""

        extra = []
        if last_step:
            extra.append(f"Last process: {last_step}")
        if remark_val:
            extra.append(remark_val)

        rows.append(_make_result_row(
            po="",           # 西拓-WIP 無 PO# 欄位，留空由人工確認
            part=part_val,
            qty=row.get(qty_col, "") if qty_col else "",
            due=row.get(due_col, "") if due_col else "",
            wip=wip_label,
            remark=" | ".join([x for x in extra if x])[:300],
            tags=["Shipped"] if wip_label == "Shipping" else [],
            source_sheet=sheet_name or "",
            source_type="xituo_wip",
        ))

    return normalize_columns(pd.DataFrame(rows))


# =========================================================
# [修正 2] 203-西拓 Wip parser
# 特徵：雙表頭（欄名橫跨2行），交期格式 "0407=>0414"
# 欄位：訂 單 號 碼 | 交貨日期 | 廠編 | 客戶料號 | 層數 | 訂購量
# =========================================================
def looks_like_203_xituo(raw_df: pd.DataFrame, filename: str = "") -> bool:
    """203-西拓 格式：有工廠名稱 '全興' 或檔名含 '203'"""
    fname = (filename or "").lower()
    if "203" in fname:
        return True

    text = "".join(
        "".join([compact_text(x) for x in raw_df.iloc[i].tolist()])
        for i in range(min(len(raw_df), 8))
    )
    return "全興電子" in text and ("P/O" in text or "訂單號碼" in text) and "工作流程計劃" in text


def _parse_203_due(text: str) -> str:
    """
    解析 203-西拓 的特殊交期格式：
    '0407=>0414' -> 取後面的日期 -> '2026-04-14'
    '0429'       -> '2026-04-29'
    """
    s = safe_text(text).strip()
    if not s or s in {"nan", "None", "日期"}:
        return ""

    # 格式: MMDD=>MMDD 或 MMDD->MMDD
    m = re.search(r"(\d{4})\s*(?:=>|->|~)\s*(\d{4})", s)
    if m:
        target = m.group(2)  # 取後面（修改後）的日期
        try:
            year = pd.Timestamp.today().year
            mm = int(target[:2])
            dd = int(target[2:])
            return pd.Timestamp(year=year, month=mm, day=dd).strftime("%Y-%m-%d")
        except Exception:
            return ""

    # 純 MMDD 格式
    m2 = re.match(r"^(\d{2})(\d{2})$", s)
    if m2:
        try:
            year = pd.Timestamp.today().year
            mm = int(m2.group(1))
            dd = int(m2.group(2))
            return pd.Timestamp(year=year, month=mm, day=dd).strftime("%Y-%m-%d")
        except Exception:
            return ""

    # 一般日期
    try:
        dt = pd.to_datetime(s, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    return ""


def parse_203_xituo_report(uploaded_file) -> pd.DataFrame:
    """
    203-西拓 Wip.xls：全興電子工作流程計劃表
    雙行表頭（行 3 + 行 4，0-indexed），資料從行 5 開始
    """
    raw_df, sheet_name = read_first_nonempty_sheet_raw(uploaded_file)
    if raw_df.empty:
        raise ValueError("203-西拓 讀取失敗")

    # 找表頭（含 P/O 且含 工作流程計劃 的行）
    header_row = None
    for i in range(min(len(raw_df), 12)):
        row_text = "".join([compact_text(x) for x in raw_df.iloc[i].tolist()])
        if ("P/O" in row_text or "訂單號碼" in row_text) and "工作流程計劃" in row_text:
            header_row = i
            break

    if header_row is None:
        raise ValueError("203-西拓：找不到表頭列")

    second_header_row = header_row + 1 if header_row + 1 < len(raw_df) else header_row

    # 合併雙行表頭
    def combine(a, b):
        a1, b1 = compact_text(a), compact_text(b)
        if not a1 and not b1:
            return ""
        if not b1:
            return a1
        if not a1:
            return b1
        merged = a1 + b1
        known = {
            "訂單號碼P/O": "P/O",
            "訂購量(PCS)": "訂購量(PCS)",
            "交貨日期": "交貨日期",
            "客戶料號": "客戶料號",
        }
        return known.get(merged, merged)

    headers = [
        combine(raw_df.iloc[header_row, idx], raw_df.iloc[second_header_row, idx]) or f"COL_{idx}"
        for idx in range(raw_df.shape[1])
    ]

    df = raw_df.iloc[second_header_row + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all").reset_index(drop=True)

    # 去除製程說明列（第一列通常是「程 | 料 | PCS...」這種子標題）
    # 判斷方式：P/O 欄若是純文字說明就跳過
    po_col_name = next((c for c in df.columns if "P/O" in c or "訂單號碼" in c), None)
    due_col_name = next((c for c in df.columns if "交貨日期" in c or c == "交貨"), None)
    part_col_name = next((c for c in df.columns if "客戶料號" in c or (c == "料號")), None)
    qty_col_name = next((c for c in df.columns if "訂購量" in c), None)
    remark_col_name = next((c for c in df.columns if "備註" in c), None)

    if not po_col_name:
        raise ValueError("203-西拓：找不到 P/O 欄位")

    process_order = [
        "下料", "內層", "壓合", "鑽孔", "一銅", "外層", "二銅蝕刻",
        "中檢測", "防焊", "文字", "化金", "無鉛", "有鉛", "OSP", "化錫", "化銀",
        "成型", "測試", "成檢", "包裝", "出貨"
    ]
    existing_proc = [p for p in process_order if p in df.columns]

    rows = []
    for _, row in df.iterrows():
        po_val = safe_text(row.get(po_col_name, ""))
        # 過濾說明列：P/O 值只有1-2字元（如「程」「料」）
        if not po_val or len(po_val) <= 2 or po_val in {"P/O", "訂單號碼", "nan"}:
            continue
        # 過濾備注列
        if "紅色字體" in po_val or "空白格子" in po_val:
            continue

        due_raw = row.get(due_col_name, "") if due_col_name else ""
        due_val = _parse_203_due(safe_text(due_raw))

        part_val = safe_text(row.get(part_col_name, "")) if part_col_name else ""
        qty_val = safe_text(row.get(qty_col_name, "")) if qty_col_name else ""
        remark_val = safe_text(row.get(remark_col_name, "")) if remark_col_name else ""

        wip_label, last_step = _guess_wip_from_last_process(row, existing_proc)

        extra = []
        if last_step:
            extra.append(f"Last process: {last_step}")
        if remark_val:
            extra.append(remark_val)

        rows.append(_make_result_row(
            po=po_val,
            part=part_val,
            qty=qty_val,
            due=due_val,
            wip=wip_label,
            remark=" | ".join([x for x in extra if x])[:300],
            tags=["Shipped"] if wip_label == "Shipping" else [],
            source_sheet=sheet_name or "",
            source_type="203_xituo",
        ))

    return normalize_columns(pd.DataFrame(rows))


# =========================================================
# [修正 3] 西拓電子有限公司進度表 parser
# 特徵：有獨立「進度」文字欄（如「外層鑽孔中」），最簡單
# =========================================================
def looks_like_xituo_simple(raw_df: pd.DataFrame, filename: str = "") -> bool:
    """西拓電子進度表：有公司名稱 '西拓電子有限公司' 且有 '進度' 欄"""
    fname = (filename or "").lower()
    text = "".join(
        "".join([compact_text(x) for x in raw_df.iloc[i].tolist()])
        for i in range(min(len(raw_df), 5))
    )
    return "西拓電子有限公司" in text and ("進度" in text or "出貨日期" in text)


def parse_xituo_simple_report(uploaded_file) -> pd.DataFrame:
    """
    西拓電子有限公司進度表.xlsx
    Row 0: 公司名稱
    Row 1: 欄位：料號 | 數量(PCS) | 下單日期 | 出貨日期 | 進度 | 備註
    """
    raw_df, sheet_name = read_first_nonempty_sheet_raw(uploaded_file)

    # 找 header row
    header_row = None
    for i in range(min(len(raw_df), 6)):
        row_text = "".join([compact_text(x) for x in raw_df.iloc[i].tolist()])
        if "料號" in row_text and ("進度" in row_text or "出貨日期" in row_text):
            header_row = i
            break

    if header_row is None:
        raise ValueError("西拓電子進度表：找不到表頭")

    df = pd.read_excel(
        uploaded_file if hasattr(uploaded_file, "read") else uploaded_file,
        sheet_name=sheet_name,
        header=header_row,
    )
    df = normalize_columns(df)
    df = df.dropna(how="all").reset_index(drop=True)

    part_col = _first_existing(df, ["料         號", "料號"] + PART_CANDIDATES)
    # 找料號欄（可能名稱有多餘空格）
    if not part_col:
        for c in df.columns:
            if "料" in c and "號" in c:
                part_col = c
                break

    qty_col = _first_existing(df, ["數量(PCS)", "數量"] + QTY_CANDIDATES)
    due_col = _first_existing(df, ["出貨日期"] + DUE_CANDIDATES)
    order_date_col = _first_existing(df, ["下單日期"])
    wip_col = _first_existing(df, ["進度"] + WIP_CANDIDATES)
    remark_col = _first_existing(df, ["備註"] + REMARK_CANDIDATES)

    rows = []
    for _, row in df.iterrows():
        part_val = safe_text(row.get(part_col, "")) if part_col else ""
        if not part_val or part_val.lower() in {"nan", "none", ""}:
            continue

        wip_val = safe_text(row.get(wip_col, "")) if wip_col else ""
        remark_val = safe_text(row.get(remark_col, "")) if remark_col else ""

        rows.append(_make_result_row(
            po="",           # 西拓電子進度表無 PO# 欄位
            part=part_val,
            qty=row.get(qty_col, "") if qty_col else "",
            due=row.get(due_col, "") if due_col else "",
            wip=wip_val,    # 直接使用文字進度（如「外層鑽孔中」）
            remark=remark_val,
            tags=[],
            source_sheet=sheet_name or "",
            source_type="xituo_simple",
        ))

    return normalize_columns(pd.DataFrame(rows))


# =========================================================
# 原版 parser（保留，略作修正）
# =========================================================

# --- 西拓工作流程表（203 格式，原名 xitop_workflow）---
def detect_xitop_header_row(raw_df: pd.DataFrame):
    for i in range(min(len(raw_df), 12)):
        row_text = "".join([compact_text(x) for x in raw_df.iloc[i].tolist()])
        if ("P/O" in row_text or "訂單號碼" in row_text) and (
            "工作流程計劃" in row_text or "交貨" in row_text or "成型" in row_text
        ):
            return i
    return None


def combine_header_cells(a, b):
    a1, b1 = compact_text(a), compact_text(b)
    if not a1 and not b1:
        return ""
    if not b1:
        return a1
    if not a1:
        return b1
    merged = a1 + b1
    replacements = {
        "訂單號碼P/O": "P/O", "訂購量(PCS)": "訂購量(PCS)", "交貨日期": "交貨日期",
        "客戶料號": "客戶料號", "成檢": "成檢", "包裝": "包裝", "出貨": "出貨",
        "測試": "測試", "成型": "成型", "防焊": "防焊", "壓合": "壓合", "鑽孔": "鑽孔",
        "內層": "內層", "外層": "外層", "一銅": "一銅", "二銅蝕刻": "二銅蝕刻",
        "化金": "化金", "OSP": "OSP", "化銀": "化銀", "備註": "備註", "文字": "文字",
    }
    return replacements.get(merged, merged)


def looks_like_xitop_workflow(raw_df: pd.DataFrame) -> bool:
    """
    [修正] 更嚴格條件，避免將西拓-WIP（無 P/O、無工作流程計劃）誤判
    必須同時有：工作流程計劃 + P/O/訂單號碼 + 交貨日期
    """
    if raw_df.empty:
        return False
    sample = raw_df.head(8).fillna("").astype(str)
    joined = "".join(sample.apply(lambda col: "".join(col), axis=1).tolist())
    # [修正] 從 >= 2 改為必須同時滿足前兩個必要條件
    has_workflow = "工作流程計劃" in joined
    has_po = "P/O" in joined or "訂單號碼" in joined
    has_due = "交貨" in joined and "日期" in joined
    return has_workflow and has_po and has_due


def parse_xitop_workflow_report(uploaded_file) -> pd.DataFrame:
    """原 203-西拓 workflow parser，現在由 parse_203_xituo_report 取代，保留相容"""
    return parse_203_xituo_report(uploaded_file)


# --- 祥竑（原版，有修正 PO 欄位取得）---
def looks_like_xianghong_two_rows(raw_df: pd.DataFrame) -> bool:
    if raw_df.empty:
        return False
    text = "|".join(
        "".join([compact_text(x) for x in raw_df.iloc[i].tolist()])
        for i in range(min(len(raw_df), 8))
    )
    return all(k in text for k in ["訂單編號", "未出貨", "發料"]) and (
        "料號" in text or "祥竑料號" in text
    )


def parse_xianghong_two_rows(uploaded_file) -> pd.DataFrame:
    raw, sheet = read_first_nonempty_sheet_raw(uploaded_file)
    header_idx = None

    for i in range(min(len(raw), 20)):
        row_text = "|".join([compact_text(x) for x in raw.iloc[i].tolist()])
        conds = [
            "項目" in row_text,
            "訂單編號" in row_text,
            ("料號" in row_text or "祥竑料號" in row_text),
            ("未出貨數量" in row_text or "未出貨" in row_text),
        ]
        if sum(conds) >= 3:
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("無法辨識祥竑表頭")

    headers = [compact_text(x) or f"COL_{i}" for i, x in enumerate(raw.iloc[header_idx].tolist())]
    header_map = {h: i for i, h in enumerate(headers)}

    proc_names = [
        "發料", "內層", "內測", "壓合", "鑽孔", "一銅", "乾膜", "二銅",
        "AOI", "半測", "防焊", "化金", "表面處理", "文字", "成型", "測試",
        "成檢(1)", "OSP", "化銀", "成檢(2)", "包裝", "庫存"
    ]

    rows = []
    i = header_idx + 1
    while i < len(raw):
        row1 = raw.iloc[i].tolist()
        row2 = raw.iloc[i + 1].tolist() if i + 1 < len(raw) else [None] * len(row1)

        po_val = safe_text(row1[header_map["訂單編號"]]) if "訂單編號" in header_map else ""

        # [修正] 料號優先用「料號」欄，其次「祥竑料號」
        part_val = ""
        if "料號" in header_map:
            part_val = safe_text(row1[header_map["料號"]])
        elif "祥竑料號" in header_map:
            part_val = safe_text(row1[header_map["祥竑料號"]])

        if not po_val and not part_val:
            i += 1
            continue

        current_step = ""
        for p in proc_names:
            idx = header_map.get(p)
            if idx is None:
                continue
            v = safe_text(row2[idx]).replace(",", "")
            try:
                if float(v) > 0:
                    current_step = p
            except Exception:
                pass

        wip_label, _ = _guess_wip_from_last_process(
            pd.Series({p: current_step == p for p in proc_names}),
            [current_step] if current_step else []
        )
        # 直接用 current_step 名稱映射
        if current_step in {"庫存"}:
            wip_label = "Shipping"
        elif current_step in {"包裝"}:
            wip_label = "Packing"
        elif current_step in {"測試", "成檢(1)", "成檢(2)"}:
            wip_label = "Inspection"
        elif current_step:
            wip_label = "Production"
        else:
            wip_label = "Production"

        qty_val = ""
        for qc in ["訂單數量", "未出貨數量"]:
            if qc in header_map:
                qty_val = safe_text(row1[header_map[qc]])
                break

        due_val = ""
        for dc in ["交貨日期", "出貨日期"]:
            if dc in header_map:
                due_val = safe_text(row1[header_map[dc]])
                break

        remark_val = safe_text(row1[header_map["備註"]]) if "備註" in header_map else ""

        rows.append(_make_result_row(
            po=po_val,
            part=part_val,
            qty=qty_val,
            due=due_val,
            wip=wip_label,
            remark=remark_val,
            tags=["Shipped"] if wip_label == "Shipping" else [],
            source_sheet=sheet or "",
            source_type="xianghong_two_rows",
        ))
        i += 2

    return normalize_columns(pd.DataFrame(rows))


# --- Profit Grand / Glocom-PG（修正：優先用 LS sheet）---
def looks_like_profit_grand(df: pd.DataFrame, filename: str = "") -> bool:
    if df is None or df.empty:
        return False
    name = (filename or "").lower()
    if "glocom-pg" in name or "profit" in name or " pg" in name or "glocom" in name:
        return True
    cols = [compact_text(c) for c in df.columns]
    joined = "|".join(cols)
    flags = [
        "PO" in cols or "PO" in joined,
        "Cust.P/N" in joined or "Cust.P/N" in joined.replace(" ", ""),
        "LSP/N" in joined or "LSPN" in joined,
        "RequiredShipdate" in joined or "RequiredShipDate" in joined or "confirmedDD" in joined or "confrimedDD" in joined,
        "WIP" in cols or "WIP" in joined,
    ]
    return sum(flags) >= 2


def parse_profit_grand(uploaded_file) -> pd.DataFrame:
    """
    Profit Grand / Glocom-PG
    [修正] 優先使用 'LS' sheet，再 fallback 到第一個可用 sheet
    欄位：PO DATE | PO | Cust. P / N | LS P/N | Q'TY | Required Ship date | confirmed DD | WIP
    """
    from excel_reader import get_excel_file_obj
    xls = get_excel_file_obj(uploaded_file)

    # [修正] 優先選 LS sheet
    target_sheet = None
    for sheet in xls.sheet_names:
        if sheet.strip().upper() == "LS":
            target_sheet = sheet
            break
    if target_sheet is None:
        target_sheet = xls.sheet_names[0]

    df = pd.read_excel(xls, sheet_name=target_sheet, header=0)
    df = normalize_columns(df)
    df = df.dropna(how="all").reset_index(drop=True)

    if df.empty:
        raise ValueError("Profit Grand 報表讀取失敗")

    po_col = _first_existing(df, ["PO"] + PO_CANDIDATES)
    part_col = _first_existing(df, ["Cust. P / N", "LS P/N"] + PART_CANDIDATES)
    qty_col = _first_existing(df, ["Q'TY"] + QTY_CANDIDATES)
    due_col = _first_existing(df, ["Required Ship date", "confirmed DD", "confrimed DD"] + DUE_CANDIDATES)
    ship_col = _first_existing(df, ["confirmed DD", "confrimed DD", "Required Ship date"] + SHIP_DATE_CANDIDATES)
    wip_col = _first_existing(df, ["WIP"] + WIP_CANDIDATES)
    remark_col = _first_existing(df, REMARK_CANDIDATES)

    rows = []
    for _, row in df.iterrows():
        po_val = safe_text(row.get(po_col, "")) if po_col else ""
        part_val = safe_text(row.get(part_col, "")) if part_col else ""
        if not po_val and not part_val:
            continue

        wip_val = safe_text(row.get(wip_col, "")) if wip_col else ""
        if not wip_val:
            wip_val = "Production"

        rows.append(_make_result_row(
            po=po_val,
            part=part_val,
            qty=row.get(qty_col, "") if qty_col else "",
            due=row.get(due_col, "") if due_col else "",
            ship=row.get(ship_col, "") if ship_col else "",
            wip=wip_val,
            remark=row.get(remark_col, "") if remark_col else "",
            tags=["Shipped"] if "ship" in wip_val.lower() or "出貨" in wip_val else [],
            source_sheet=target_sheet or "",
            source_type="profit_grand",
        ))

    return normalize_columns(pd.DataFrame(rows))


# --- 一般 Excel fallback（原版）---
def _score_standard_df(df: pd.DataFrame) -> int:
    if df is None or df.empty:
        return -999
    score = 0
    if _first_existing(df, PO_CANDIDATES):
        score += 20
    if _first_existing(df, PART_CANDIDATES):
        score += 15
    if _first_existing(df, WIP_CANDIDATES):
        score += 12
    if _first_existing(df, QTY_CANDIDATES):
        score += 8
    if _first_existing(df, DUE_CANDIDATES):
        score += 8
    process_cols = _detect_process_columns(df)
    if len(process_cols) >= 3:
        score += 10
    non_empty_cols = len([c for c in df.columns if compact_text(c)])
    if non_empty_cols >= 5:
        score += 5
    if len(df) >= 3:
        score += 5
    unnamed = sum(1 for c in df.columns if compact_text(c).startswith("unnamed"))
    score -= unnamed
    return score


def _standardize_generic_df(df: pd.DataFrame, source_type: str, sheet_name: str = "") -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    df = normalize_columns(df.copy())
    po_col = _first_existing(df, PO_CANDIDATES)
    part_col = _first_existing(df, PART_CANDIDATES)
    qty_col = _first_existing(df, QTY_CANDIDATES)
    due_col = _first_existing(df, DUE_CANDIDATES)
    ship_col = _first_existing(df, SHIP_DATE_CANDIDATES)
    wip_col = _first_existing(df, WIP_CANDIDATES)
    remark_col = _first_existing(df, REMARK_CANDIDATES)
    process_cols = _detect_process_columns(df)

    rows = []
    for _, row in df.iterrows():
        po_val = safe_text(row.get(po_col, "")) if po_col else ""
        part_val = safe_text(row.get(part_col, "")) if part_col else ""
        if not po_val and not part_val:
            continue

        wip_val = safe_text(row.get(wip_col, "")) if wip_col else ""
        if not wip_val and process_cols:
            wip_val, _ = _guess_wip_from_last_process(row, process_cols)

        remark_val = safe_text(row.get(remark_col, "")) if remark_col else ""
        rows.append(_make_result_row(
            po=po_val, part=part_val,
            qty=row.get(qty_col, "") if qty_col else "",
            due=row.get(due_col, "") if due_col else "",
            ship=row.get(ship_col, "") if ship_col else "",
            wip=wip_val, remark=remark_val,
            tags=["Shipped"] if wip_val == "Shipping" else [],
            source_sheet=sheet_name, source_type=source_type,
        ))

    return normalize_columns(pd.DataFrame(rows))


def parse_standard_excel(uploaded_file) -> pd.DataFrame:
    candidates = []
    for header in [0, 1, 2, 3]:
        try:
            df, sheet = read_first_nonempty_sheet_with_header(uploaded_file, header=header)
            if df is not None and not df.empty:
                df = normalize_columns(df)
                score = _score_standard_df(df)
                candidates.append((header, df, sheet, score))
        except Exception:
            continue

    if not candidates:
        raise ValueError("Excel file has no readable non-empty sheet.")

    candidates.sort(key=lambda x: x[3], reverse=True)
    header, best_df, sheet, _score = candidates[0]
    return _standardize_generic_df(
        best_df,
        source_type=f"standard_excel:{sheet}:header{header}",
        sheet_name=sheet or "",
    )


# --- txt / csv（原版）---
def parse_txt_file(uploaded_file) -> pd.DataFrame:
    text = uploaded_file.getvalue().decode("utf-8", errors="ignore")
    rows = parse_email_text_to_rows(text)
    return normalize_columns(pd.DataFrame(rows))


def parse_csv_file(uploaded_file) -> pd.DataFrame:
    raw = pd.read_csv(uploaded_file)
    raw = normalize_columns(raw)
    return _standardize_generic_df(raw, source_type="csv", sheet_name="")


# =========================================================
# [修正] 主入口 — 加入新 parser 路由
# =========================================================
def read_import_dataframe(uploaded_file):
    """
    app.py 匯入主入口
    回傳：(df, parse_mode)

    路由優先順序：
    1. txt  → email_text
    2. csv  → csv
    3. 西拓電子有限公司進度表 → xituo_simple      [新增]
    4. 203-西拓 workflow      → 203_xituo         [新增]
    5. 西拓-WIP               → xituo_wip         [新增]
    6. 祥竑                   → xianghong_two_rows
    7. Profit Grand / Glocom  → profit_grand
    8. fallback               → standard_excel
    """
    name = uploaded_file.name.lower()

    if name.endswith(".txt"):
        return parse_txt_file(uploaded_file), "email_text"

    if name.endswith(".csv"):
        return parse_csv_file(uploaded_file), "csv"

    raw_df, _ = read_first_nonempty_sheet_raw(uploaded_file)
    if raw_df is None or raw_df.empty:
        raise ValueError("Excel 檔案沒有可讀資料")

    # [新增] 西拓電子有限公司進度表（最簡單格式，優先判斷）
    if looks_like_xituo_simple(raw_df, filename=name):
        return parse_xituo_simple_report(uploaded_file), "xituo_simple"

    # [新增] 203-西拓（工作流程計劃表，雙表頭）
    if looks_like_xitop_workflow(raw_df):
        return parse_203_xituo_report(uploaded_file), "203_xituo"

    # [新增] 西拓-WIP（有「西拓 WIP 進度表」標題）
    if looks_like_xituo_wip(raw_df):
        return parse_xituo_wip_report(uploaded_file), "xituo_wip"

    # 祥竑
    if looks_like_xianghong_two_rows(raw_df):
        return parse_xianghong_two_rows(uploaded_file), "xianghong_two_rows"

    # Profit Grand / Glocom-PG
    try:
        df0, _sheet0 = read_first_nonempty_sheet_with_header(uploaded_file, header=0)
        if looks_like_profit_grand(df0, filename=name):
            return parse_profit_grand(uploaded_file), "profit_grand"
    except Exception:
        pass

    return parse_standard_excel(uploaded_file), "standard_excel"
