import re
import pandas as pd
from excel_reader import read_first_nonempty_sheet_raw, read_first_nonempty_sheet_with_header
from utils import safe_text, normalize_columns, compact_text, normalize_due_text
from text_ocr_parsers import parse_email_text_to_rows


def detect_xitop_header_row(raw_df: pd.DataFrame):
    for i in range(min(len(raw_df), 12)):
        row_text = "".join([compact_text(x) for x in raw_df.iloc[i].tolist()])
        if ("P/O" in row_text or "訂單號碼" in row_text) and ("工作流程計劃" in row_text or "交貨" in row_text or "成型" in row_text):
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
    if raw_df.empty:
        return False
    sample = raw_df.head(8).fillna("").astype(str)
    joined = "".join(sample.apply(lambda col: "".join(col), axis=1).tolist())
    flags = ["工作流程計劃" in joined, "P/O" in joined or "訂單號碼" in joined, "交貨" in joined and "日期" in joined, any(x in joined for x in ["成型", "測試", "防焊"])]
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
        headers.append(combine_header_cells(raw_df.iloc[header_row, idx], raw_df.iloc[second_header_row, idx]) or f"COL_{idx}")
    df = raw_df.iloc[second_header_row + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [compact_text(c) or f"UNNAMED_{i}" for i, c in enumerate(df.columns)]

    po_source = next((c for c in df.columns if "P/O" in c or "訂單號碼" in c), None)
    due_source = next((c for c in df.columns if "交貨日期" in c or c == "交貨"), None)
    part_source = next((c for c in df.columns if "客戶料號" in c or "料號" in c), None)
    qty_source = next((c for c in df.columns if "訂購量" in c), None)
    remark_source = next((c for c in df.columns if "備註" in c), None)
    if not po_source:
        raise ValueError("西拓報表解析失敗：找不到 P/O 欄位")

    process_order = ["下料", "內層", "壓合", "鑽孔", "一銅", "外層", "二銅蝕刻", "中檢測", "防焊", "文字", "化金", "無鉛", "有鉛", "OSP", "化錫", "化銀", "成型", "測試", "成檢", "包裝", "出貨"]
    existing_process_cols = [p for p in process_order if p in df.columns]

    rows = []
    for _, row in df.iterrows():
        po_val = safe_text(row.get(po_source, ""))
        if not po_val:
            continue
        last_step = ""
        for p in existing_process_cols:
            if compact_text(row.get(p, "")) not in {"", "nan", "None"}:
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
        rows.append({
            "PO#": po_val,
            "Part No": safe_text(row.get(part_source, "")) if part_source else "",
            "Qty": safe_text(row.get(qty_source, "")) if qty_source else "",
            "Factory Due Date": normalize_due_text(row.get(due_source, "")) if due_source else "",
            "Ship Date": normalize_due_text(row.get(due_source, "")) if due_source else "",
            "WIP": wip,
            "Remark": " | ".join([x for x in [f"Last process: {last_step}" if last_step else "", remark_val] if x])[:300],
            "Customer Remark Tags": ["Shipped"] if wip == "Shipping" else [],
            "_source_sheet": sheet_name or "",
            "_source_type": "xitop_workflow",
        })
    return normalize_columns(pd.DataFrame(rows))


def looks_like_xianghong_two_rows(raw_df: pd.DataFrame) -> bool:
    if raw_df.empty:
        return False
    text = "|".join("".join([compact_text(x) for x in raw_df.iloc[i].tolist()]) for i in range(min(len(raw_df), 8)))
    return all(k in text for k in ["訂單編號", "祥竑料號", "未出貨", "發料"])


def parse_xianghong_two_rows(uploaded_file) -> pd.DataFrame:
    raw, sheet = read_first_nonempty_sheet_raw(uploaded_file)
    header_idx = None
    for i in range(min(len(raw), 20)):
        row_text = "|".join([compact_text(x) for x in raw.iloc[i].tolist()])
        conds = ["項目" in row_text, "訂單編號" in row_text, ("料號" in row_text or "祥竑料號" in row_text), ("未出貨數量" in row_text or "未出貨" in row_text)]
        if sum(conds) >= 3:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("無法辨識祥竑表頭")
    headers = [compact_text(x) or f"COL_{i}" for i, x in enumerate(raw.iloc[header_idx].tolist())]
    header_map = {h: i for i, h in enumerate(headers)}
    proc_names = ["發料", "內層", "內測", "壓合", "鑽孔", "一銅", "乾膜", "二銅", "AOI", "半測", "防焊", "化金", "表面處理", "文字", "成型", "測試", "成檢(1)", "OSP", "化銀", "成檢(2)", "包裝", "庫存"]
    rows = []
    i = header_idx + 1
    while i < len(raw):
        row1 = raw.iloc[i].tolist()
        row2 = raw.iloc[i + 1].tolist() if i + 1 < len(raw) else [None] * len(row1)
        po_val = safe_text(row1[header_map.get("訂單編號")]) if "訂單編號" in header_map else ""
        if not po_val:
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
        rows.append({
            "PO#": po_val,
            "Part No": safe_text(row1[header_map.get("料號")]) if "料號" in header_map else "",
            "Qty": safe_text(row1[header_map.get("訂單數量")]) if "訂單數量" in header_map else "",
            "Factory Due Date": normalize_due_text(row1[header_map.get("交貨日期")]) if "交貨日期" in header_map else "",
            "Ship Date": normalize_due_text(row1[header_map.get("交貨日期")]) if "交貨日期" in header_map else "",
            "WIP": current_step or "Production",
            "Remark": safe_text(row1[header_map.get("備註")]) if "備註" in header_map else "",
            "Customer Remark Tags": [],
            "_source_sheet": sheet or "",
            "_source_type": "xianghong_two_rows",
        })
        i += 2
    return normalize_columns(pd.DataFrame(rows))


def read_import_dataframe(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".txt"):
        rows = parse_email_text_to_rows(uploaded_file.getvalue().decode("utf-8", errors="ignore"))
        return normalize_columns(pd.DataFrame(rows)), "email_text"
    if name.endswith(".csv"):
        return normalize_columns(pd.read_csv(uploaded_file)), "csv"
    raw_df, _ = read_first_nonempty_sheet_raw(uploaded_file)
    if looks_like_xitop_workflow(raw_df):
        return parse_xitop_workflow_report(uploaded_file), "xitop_workflow"
    if looks_like_xianghong_two_rows(raw_df):
        return parse_xianghong_two_rows(uploaded_file), "xianghong_two_rows"
    df, sheet = read_first_nonempty_sheet_with_header(uploaded_file, header=0)
    if df.empty:
        raise ValueError("Excel file has no readable non-empty sheet.")
    return df, f"standard_excel:{sheet}"
