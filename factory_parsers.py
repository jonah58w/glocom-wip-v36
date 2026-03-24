# -*- coding: utf-8 -*-
"""
factory_parsers.py
工廠進度檔解析器

支援：
- 西拓雙層表頭工作流程表
- 祥竑雙列資料表
- Profit Grand / Glocom-PG
- 一般標準 Excel
- txt / csv

輸出統一欄位：
PO#, Part No, Qty, Factory Due Date, Ship Date, WIP, Remark, Customer Remark Tags
"""

from __future__ import annotations

import pandas as pd
from typing import Any, Dict, List, Optional, Tuple

from excel_reader import (
    read_first_nonempty_sheet_raw,
    read_first_nonempty_sheet_with_header,
)
from utils import safe_text, normalize_columns, compact_text, normalize_due_text
from text_ocr_parsers import parse_email_text_to_rows


# =========================================================
# 標準欄位候選
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
    "預交日", "預定交期", "交貨期"
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
    "鑽孔", "沉銅", "一銅", "電鍍", "乾膜", "外層", "二銅", "二銅蝕刻",
    "AOI", "半測", "防焊", "文字", "噴錫", "化金", "OSP", "化銀",
    "成型", "V-CUT", "測試", "成檢", "包裝", "出貨", "庫存"
]


# =========================================================
# 共用工具
# =========================================================
def _first_existing(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _make_result_row(
    po: Any = "",
    part: Any = "",
    qty: Any = "",
    due: Any = "",
    ship: Any = "",
    wip: Any = "",
    remark: Any = "",
    tags: Optional[List[str]] = None,
    source_sheet: str = "",
    source_type: str = "",
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


def _guess_wip_from_processes(row: pd.Series, process_cols: List[str]) -> str:
    last_step = ""
    for p in process_cols:
        value = compact_text(row.get(p, ""))
        if value not in {"", "nan", "none", "null", "-", "--", "0", "0.0"}:
            last_step = p

    if not last_step:
        return ""

    if last_step == "出貨":
        return "Shipping"
    if last_step == "包裝":
        return "Packing"
    if last_step in {"測試", "成檢", "成檢(1)", "成檢(2)", "FQC", "QA"}:
        return "Inspection"
    if last_step in {"工程", "Gerber", "EQ"}:
        return "Engineering"
    return "Production"


def _standardize_generic_df(df: pd.DataFrame, source_type: str, sheet_name: str = "") -> pd.DataFrame:
    """
    將一般表格映射成標準輸出欄位
    """
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

    rows: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        po_val = safe_text(row.get(po_col, "")) if po_col else ""
        part_val = safe_text(row.get(part_col, "")) if part_col else ""

        if not po_val and not part_val:
            continue

        wip_val = safe_text(row.get(wip_col, "")) if wip_col else ""
        if not wip_val and process_cols:
            wip_val = _guess_wip_from_processes(row, process_cols)

        remark_val = safe_text(row.get(remark_col, "")) if remark_col else ""
        tags: List[str] = []
        if wip_val == "Shipping":
            tags.append("Shipped")

        rows.append(_make_result_row(
            po=po_val,
            part=part_val,
            qty=row.get(qty_col, "") if qty_col else "",
            due=row.get(due_col, "") if due_col else "",
            ship=row.get(ship_col, "") if ship_col else "",
            wip=wip_val,
            remark=remark_val,
            tags=tags,
            source_sheet=sheet_name,
            source_type=source_type,
        ))

    return normalize_columns(pd.DataFrame(rows))


# =========================================================
# 西拓 parser
# =========================================================
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
        "訂單號碼P/O": "P/O",
        "訂購量(PCS)": "訂購量(PCS)",
        "交貨日期": "交貨日期",
        "客戶料號": "客戶料號",
        "成檢": "成檢",
        "包裝": "包裝",
        "出貨": "出貨",
        "測試": "測試",
        "成型": "成型",
        "防焊": "防焊",
        "壓合": "壓合",
        "鑽孔": "鑽孔",
        "內層": "內層",
        "外層": "外層",
        "一銅": "一銅",
        "二銅蝕刻": "二銅蝕刻",
        "化金": "化金",
        "OSP": "OSP",
        "化銀": "化銀",
        "備註": "備註",
        "文字": "文字",
    }
    return replacements.get(merged, merged)


def looks_like_xitop_workflow(raw_df: pd.DataFrame) -> bool:
    if raw_df.empty:
        return False
    sample = raw_df.head(8).fillna("").astype(str)
    joined = "".join(sample.apply(lambda col: "".join(col), axis=1).tolist())
    flags = [
        "工作流程計劃" in joined,
        "P/O" in joined or "訂單號碼" in joined,
        ("交貨" in joined and "日期" in joined),
        any(x in joined for x in ["成型", "測試", "防焊", "包裝", "出貨"]),
    ]
    return sum(flags) >= 2


def parse_xitop_workflow_report(uploaded_file) -> pd.DataFrame:
    raw_df, sheet_name = read_first_nonempty_sheet_raw(uploaded_file)
    if raw_df.empty:
        raise ValueError("西拓報表讀取失敗：工作表為空")

    header_row = detect_xitop_header_row(raw_df)
    if header_row is None:
        raise ValueError("無法辨識西拓報表表頭")

    second_header_row = header_row + 1 if header_row + 1 < len(raw_df) else header_row

    headers = []
    for idx in range(raw_df.shape[1]):
        headers.append(
            combine_header_cells(raw_df.iloc[header_row, idx], raw_df.iloc[second_header_row, idx]) or f"COL_{idx}"
        )

    df = raw_df.iloc[second_header_row + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [compact_text(c) or f"UNNAMED_{i}" for i, c in enumerate(df.columns)]

    po_source = _first_existing(df, ["P/O", "訂單號碼"] + PO_CANDIDATES)
    due_source = _first_existing(df, ["交貨日期", "交貨"] + DUE_CANDIDATES)
    part_source = _first_existing(df, ["客戶料號", "料號"] + PART_CANDIDATES)
    qty_source = _first_existing(df, ["訂購量(PCS)", "訂購量"] + QTY_CANDIDATES)
    remark_source = _first_existing(df, ["備註"] + REMARK_CANDIDATES)

    process_order = [
        "發料", "下料", "內層", "內乾", "內蝕", "黑化", "壓合", "鑽孔",
        "一銅", "乾膜", "外層", "二銅", "二銅蝕刻", "AOI", "半測",
        "防焊", "文字", "化金", "OSP", "化銀", "成型", "測試", "成檢",
        "包裝", "出貨"
    ]
    existing_process_cols = [p for p in process_order if p in df.columns]

    rows = []
    for _, row in df.iterrows():
        po_val = safe_text(row.get(po_source, "")) if po_source else ""
        part_val = safe_text(row.get(part_source, "")) if part_source else ""

        if not po_val and not part_val:
            continue

        last_step = ""
        for p in existing_process_cols:
            if compact_text(row.get(p, "")) not in {"", "nan", "none", "null"}:
                last_step = p

        if last_step == "出貨":
            wip = "Shipping"
        elif last_step == "包裝":
            wip = "Packing"
        elif last_step in ["成檢", "測試"]:
            wip = "Inspection"
        else:
            wip = "Production" if last_step else ""

        remark_val = safe_text(row.get(remark_source, "")) if remark_source else ""
        rows.append(_make_result_row(
            po=po_val,
            part=part_val,
            qty=row.get(qty_source, "") if qty_source else "",
            due=row.get(due_source, "") if due_source else "",
            ship=row.get(due_source, "") if due_source else "",
            wip=wip,
            remark=" | ".join([x for x in [f"Last process: {last_step}" if last_step else "", remark_val] if x])[:300],
            tags=["Shipped"] if wip == "Shipping" else [],
            source_sheet=sheet_name or "",
            source_type="xitop_workflow",
        ))

    return normalize_columns(pd.DataFrame(rows))


# =========================================================
# 祥竑 parser
# =========================================================
def looks_like_xianghong_two_rows(raw_df: pd.DataFrame) -> bool:
    if raw_df.empty:
        return False
    text = "|".join(
        "".join([compact_text(x) for x in raw_df.iloc[i].tolist()])
        for i in range(min(len(raw_df), 8))
    )
    return all(k in text for k in ["訂單編號", "未出貨", "發料"]) and ("料號" in text or "祥竑料號" in text)


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

        po_val = safe_text(row1[header_map.get("訂單編號")]) if "訂單編號" in header_map else ""
        part_val = (
            safe_text(row1[header_map.get("料號")]) if "料號" in header_map else
            safe_text(row1[header_map.get("祥竑料號")]) if "祥竑料號" in header_map else ""
        )

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

        if current_step == "包裝":
            wip = "Packing"
        elif current_step in ["測試", "成檢(1)", "成檢(2)"]:
            wip = "Inspection"
        elif current_step == "庫存":
            wip = "Shipping"
        else:
            wip = current_step or "Production"

        qty_val = ""
        if "訂單數量" in header_map:
            qty_val = safe_text(row1[header_map.get("訂單數量")])
        elif "未出貨數量" in header_map:
            qty_val = safe_text(row1[header_map.get("未出貨數量")])

        due_val = ""
        if "交貨日期" in header_map:
            due_val = safe_text(row1[header_map.get("交貨日期")])
        elif "出貨日期" in header_map:
            due_val = safe_text(row1[header_map.get("出貨日期")])

        remark_val = safe_text(row1[header_map.get("備註")]) if "備註" in header_map else ""

        rows.append(_make_result_row(
            po=po_val,
            part=part_val,
            qty=qty_val,
            due=due_val,
            ship=due_val,
            wip=wip,
            remark=remark_val,
            tags=["Shipped"] if wip == "Shipping" else [],
            source_sheet=sheet or "",
            source_type="xianghong_two_rows",
        ))
        i += 2

    return normalize_columns(pd.DataFrame(rows))


# =========================================================
# Profit Grand / Glocom-PG parser
# =========================================================
def looks_like_profit_grand(df: pd.DataFrame, filename: str = "") -> bool:
    if df is None or df.empty:
        return False

    name = (filename or "").lower()
    if "glocom-pg" in name or "profit" in name or " pg" in name:
        return True

    cols = [compact_text(c) for c in df.columns]
    joined = "|".join(cols)
    flags = [
        "PO" in cols or "PO" in joined,
        ("Cust.P/N" in joined or "Cust.P/N" in cols or "Cust.P/N" in joined.replace(" ", "")),
        ("LSP/N" in joined or "LSPN" in joined or "LSP/N" in cols),
        ("RequiredShipdate" in joined or "RequiredShipDate" in joined),
        ("WIP" in cols or "WIP" in joined),
    ]
    return sum(flags) >= 2


def parse_profit_grand(uploaded_file) -> pd.DataFrame:
    """
    Profit Grand / Glocom-PG 常見欄位：
    PO DATE | PO | Cust. P / N | LS P/N | Q'TY | Required Ship date | confirmed DD | WIP
    """
    df, sheet_name = read_first_nonempty_sheet_with_header(uploaded_file, header=0)
    if df.empty:
        raise ValueError("Profit Grand 報表讀取失敗")

    df = normalize_columns(df)
    df = df[df.notna().any(axis=1)].reset_index(drop=True)

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
            source_sheet=sheet_name or "",
            source_type="profit_grand",
        ))

    return normalize_columns(pd.DataFrame(rows))


# =========================================================
# 一般 Excel fallback
# =========================================================
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


def parse_standard_excel(uploaded_file) -> pd.DataFrame:
    """
    一般標準 Excel fallback
    會嘗試 header=0,1,2,3 找最佳結果
    """
    candidates: List[Tuple[int, pd.DataFrame, Optional[str], int]] = []

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


# =========================================================
# txt / csv
# =========================================================
def parse_txt_file(uploaded_file) -> pd.DataFrame:
    text = uploaded_file.getvalue().decode("utf-8", errors="ignore")
    rows = parse_email_text_to_rows(text)
    return normalize_columns(pd.DataFrame(rows))


def parse_csv_file(uploaded_file) -> pd.DataFrame:
    raw = pd.read_csv(uploaded_file)
    raw = normalize_columns(raw)
    return _standardize_generic_df(raw, source_type="csv", sheet_name="")


# =========================================================
# 主入口
# =========================================================
def read_import_dataframe(uploaded_file):
    """
    app.py 匯入主入口
    回傳：
    (df, parse_mode)
    """
    name = uploaded_file.name.lower()

    if name.endswith(".txt"):
        return parse_txt_file(uploaded_file), "email_text"

    if name.endswith(".csv"):
        return parse_csv_file(uploaded_file), "csv"

    raw_df, _ = read_first_nonempty_sheet_raw(uploaded_file)
    if raw_df is None or raw_df.empty:
        raise ValueError("Excel 檔案沒有可讀資料")

    if looks_like_xitop_workflow(raw_df):
        return parse_xitop_workflow_report(uploaded_file), "xitop_workflow"

    if looks_like_xianghong_two_rows(raw_df):
        return parse_xianghong_two_rows(uploaded_file), "xianghong_two_rows"

    try:
        df0, _sheet0 = read_first_nonempty_sheet_with_header(uploaded_file, header=0)
        if looks_like_profit_grand(df0, filename=name):
            return parse_profit_grand(uploaded_file), "profit_grand"
    except Exception:
        pass

    return parse_standard_excel(uploaded_file), "standard_excel"
