import io
import os
import re
import json
import math
import time
from datetime import datetime, date
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="GLOCOM Control Tower",
    page_icon="🏭",
    layout="wide"
)

# =========================================================
# DEFAULT CONFIG
# =========================================================
DEFAULT_TABLE_URL = "https://app.teable.ai/api/table/tbl6c05EPXYtJcZfeir/record"
DEFAULT_TEABLE_WEB_URL = "https://app.teable.ai/base/bsedgLzbHjiK0XoZH01/table/tbl6c05EPXYtJcZfeir"

try:
    TEABLE_TOKEN = st.secrets.get("TEABLE_TOKEN", "")
except Exception:
    TEABLE_TOKEN = ""

try:
    TABLE_URL = st.secrets.get("TEABLE_TABLE_URL", DEFAULT_TABLE_URL)
except Exception:
    TABLE_URL = DEFAULT_TABLE_URL

try:
    TEABLE_WEB_URL = st.secrets.get("TEABLE_WEB_URL", DEFAULT_TEABLE_WEB_URL)
except Exception:
    TEABLE_WEB_URL = DEFAULT_TEABLE_WEB_URL

# =========================================================
# USER-ADJUSTABLE FIELD MAPPING
# 請依你的 Teable 主表欄位名稱微調
# =========================================================
TEABLE_MATCH_FIELDS = [
    "訂單號",
    "PO",
    "PO No",
    "PO NO",
    "Order No",
    "訂單編號",
    "工單號",
]

TEABLE_PART_FIELDS = [
    "料號",
    "客戶料號",
    "Part No",
    "Part Number",
    "P/N",
    "Customer P/N",
    "Cust. P/N",
    "品名料號",
    "客戶品號",
    "產品料號",
]

TEABLE_FACTORY_PART_FIELDS = [
    "工廠料號",
    "廠編",
    "Vendor P/N",
    "Factory P/N",
    "LS P/N",
    "祥竑料號",
]

TEABLE_PROGRESS_TARGET_FIELDS = [
    "進度",
    "WIP",
    "WIP進度",
    "目前進度",
    "工廠進度",
]

TEABLE_REPLY_DATE_FIELDS = [
    "工廠回覆日期",
    "更新日期",
    "WIP更新日期",
    "最新回覆日期",
]

TEABLE_NOTE_FIELDS = [
    "備註",
    "工廠備註",
    "WIP備註",
]

TEABLE_SOURCE_FIELDS = [
    "資料來源",
    "WIP來源",
    "工廠來源",
]

TEABLE_QTY_FIELDS = [
    "數量",
    "QTY",
    "Qty",
    "訂單數量",
]

# =========================================================
# HELPERS
# =========================================================
def safe_strip(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


def is_blank(v: Any) -> bool:
    s = safe_strip(v)
    return s == "" or s.lower() == "nan" or s.lower() == "none"


def normalize_space(s: str) -> str:
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def normalize_part_no(s: Any) -> str:
    """
    料號正規化：
    - 大寫
    - 去括號內容
    - 去 REV / VER / VERSION 之後內容
    - 去空白與符號，只留英數
    """
    s = safe_strip(s).upper()
    if not s:
        return ""

    s = normalize_space(s)

    # 去掉括號內容
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"\[.*?\]", "", s)

    # 去掉版本字尾
    s = re.sub(r"\bREV(?:ISION)?\b[.\s_-]*[A-Z0-9\-_/]*", "", s)
    s = re.sub(r"\bVER(?:SION)?\b[.\s_-]*[A-Z0-9\-_/]*", "", s)

    # 常見尾碼清理
    s = re.sub(r"NEW VERSION.*$", "", s)
    s = re.sub(r"版.*$", "", s)

    # 只留英數
    s = re.sub(r"[^A-Z0-9]", "", s)

    return s.strip()


def normalize_order_no(s: Any) -> str:
    s = safe_strip(s).upper()
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^A-Z0-9\-]", "", s)
    return s


def excel_serial_or_date_to_str(v: Any) -> str:
    if pd.isna(v):
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    s = safe_strip(v)
    return s


def first_existing_field(d: Dict[str, Any], names: List[str]) -> str:
    for name in names:
        if name in d and not is_blank(d.get(name)):
            return safe_strip(d.get(name))
    return ""


def find_first_matching_col(columns: List[str], keywords: List[str]) -> Optional[str]:
    cols_norm = [(c, normalize_space(str(c)).lower()) for c in columns]
    for kw in keywords:
        kw2 = normalize_space(kw).lower()
        for original, c in cols_norm:
            if kw2 == c:
                return original
        for original, c in cols_norm:
            if kw2 in c:
                return original
    return None


def try_parse_numeric(v: Any) -> Optional[float]:
    if pd.isna(v):
        return None
    s = safe_strip(v)
    if not s:
        return None
    s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return None


def normalize_progress_text(s: Any) -> str:
    s = safe_strip(s)
    s = normalize_space(s)
    s = s.replace("待开料", "待開料")
    s = s.replace("备料中", "備料中")
    s = s.replace("备料", "備料")
    s = s.replace("已出货", "已出貨")
    s = s.replace("出货", "出貨")
    return s


def progress_rank_map() -> Dict[str, int]:
    return {
        "待開料": 5,
        "備料中": 8,
        "下料": 10,
        "內層": 20,
        "壓合": 30,
        "鑽孔": 40,
        "一銅": 50,
        "外層": 55,
        "二銅": 60,
        "AOI": 65,
        "半測": 70,
        "防焊": 80,
        "文字": 85,
        "化金": 90,
        "表面處理": 92,
        "成型": 95,
        "測試": 97,
        "成檢": 98,
        "包裝": 99,
        "已出貨": 100,
        "出貨": 100,
    }


def choose_better_progress(old_p: str, new_p: str) -> str:
    rank = progress_rank_map()
    old_r = rank.get(old_p, 0)
    new_r = rank.get(new_p, 0)
    return new_p if new_r >= old_r else old_p


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df.columns = [safe_strip(c) for c in df.columns]
    return df


# =========================================================
# TEABLE API
# =========================================================
def teable_headers() -> Dict[str, str]:
    headers = {"Content-Type": "application/json"}
    if TEABLE_TOKEN:
        headers["Authorization"] = f"Bearer {TEABLE_TOKEN}"
    return headers


def fetch_teable_records(limit: int = 1000) -> List[Dict[str, Any]]:
    if not TEABLE_TOKEN:
        return []

    all_records = []
    page_token = None

    while True:
        params = {"pageSize": 200}
        if page_token:
            params["pageToken"] = page_token

        try:
            resp = requests.get(TABLE_URL, headers=teable_headers(), params=params, timeout=30)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            st.error(f"讀取 Teable 主表失敗：{e}")
            return all_records

        records = data.get("records", []) or data.get("data", {}).get("records", [])
        all_records.extend(records)

        if len(all_records) >= limit:
            break

        page_token = data.get("pageToken") or data.get("nextPageToken") or data.get("data", {}).get("pageToken")
        if not page_token:
            break

    return all_records


def update_teable_record(record_id: str, fields: Dict[str, Any]) -> Tuple[bool, str]:
    if not TEABLE_TOKEN:
        return False, "未設定 TEABLE_TOKEN"

    payload = {"records": [{"id": record_id, "fields": fields}]}

    try:
        resp = requests.patch(TABLE_URL, headers=teable_headers(), json=payload, timeout=30)
        if resp.status_code >= 400:
            return False, f"{resp.status_code}: {resp.text}"
        return True, "OK"
    except Exception as e:
        return False, str(e)


# =========================================================
# MAIN TABLE INDEX
# =========================================================
def build_teable_indexes(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    order_index: Dict[str, List[Dict[str, Any]]] = {}
    part_index: Dict[str, List[Dict[str, Any]]] = {}
    raw_index: List[Dict[str, Any]] = []

    for rec in records:
        fields = rec.get("fields", {}) or {}
        rec_id = rec.get("id", "")

        order_values = []
        for f in TEABLE_MATCH_FIELDS:
            if f in fields:
                ov = normalize_order_no(fields.get(f))
                if ov:
                    order_values.append(ov)

        part_values = []
        for f in TEABLE_PART_FIELDS + TEABLE_FACTORY_PART_FIELDS:
            if f in fields:
                pv = normalize_part_no(fields.get(f))
                if pv:
                    part_values.append(pv)

        item = {
            "id": rec_id,
            "fields": fields,
            "order_values": list(set(order_values)),
            "part_values": list(set(part_values)),
        }
        raw_index.append(item)

        for ov in item["order_values"]:
            order_index.setdefault(ov, []).append(item)

        for pv in item["part_values"]:
            part_index.setdefault(pv, []).append(item)

    return {
        "order_index": order_index,
        "part_index": part_index,
        "raw_index": raw_index,
    }


def match_teable_record(parsed_row: Dict[str, Any], indexes: Dict[str, Any]) -> Tuple[Optional[Dict[str, Any]], str]:
    row_order = normalize_order_no(parsed_row.get("order_no", ""))
    row_part = normalize_part_no(parsed_row.get("part_no", ""))
    row_factory_part = normalize_part_no(parsed_row.get("factory_part_no", ""))

    # 1) 先用訂單號
    if row_order and row_order in indexes["order_index"]:
        candidates = indexes["order_index"][row_order]
        if len(candidates) == 1:
            return candidates[0], f"訂單號命中：{row_order}"
        if row_part:
            for c in candidates:
                if row_part in c["part_values"]:
                    return c, f"訂單號+料號命中：{row_order} / {row_part}"
        if row_factory_part:
            for c in candidates:
                if row_factory_part in c["part_values"]:
                    return c, f"訂單號+工廠料號命中：{row_order} / {row_factory_part}"
        return candidates[0], f"訂單號多筆，先取第一筆：{row_order}"

    # 2) 再用客戶料號
    if row_part and row_part in indexes["part_index"]:
        candidates = indexes["part_index"][row_part]
        if len(candidates) == 1:
            return candidates[0], f"料號命中：{row_part}"
        if row_order:
            for c in candidates:
                if row_order in c["order_values"]:
                    return c, f"料號+訂單號命中：{row_part} / {row_order}"
        return candidates[0], f"料號多筆，先取第一筆：{row_part}"

    # 3) 再用工廠料號
    if row_factory_part and row_factory_part in indexes["part_index"]:
        candidates = indexes["part_index"][row_factory_part]
        if len(candidates) == 1:
            return candidates[0], f"工廠料號命中：{row_factory_part}"
        if row_order:
            for c in candidates:
                if row_order in c["order_values"]:
                    return c, f"工廠料號+訂單號命中：{row_factory_part} / {row_order}"
        return candidates[0], f"工廠料號多筆，先取第一筆：{row_factory_part}"

    # 4) 最後做模糊包含比對
    if row_part:
        for c in indexes["raw_index"]:
            for pv in c["part_values"]:
                if row_part and pv and (row_part in pv or pv in row_part):
                    return c, f"模糊料號命中：{row_part} ~ {pv}"

    if row_factory_part:
        for c in indexes["raw_index"]:
            for pv in c["part_values"]:
                if row_factory_part and pv and (row_factory_part in pv or pv in row_factory_part):
                    return c, f"模糊工廠料號命中：{row_factory_part} ~ {pv}"

    return None, "找不到對應主表資料"


# =========================================================
# PROCESS / PROGRESS LOGIC
# =========================================================
PROCESS_STEPS = [
    ("待開料", ["待開料", "備料", "備料中"]),
    ("下料", ["下料", "工", "下"]),
    ("內層", ["內層", "內"]),
    ("壓合", ["壓合", "壓"]),
    ("鑽孔", ["鑽孔", "鑽"]),
    ("一銅", ["一銅"]),
    ("外層", ["外層"]),
    ("二銅", ["二銅"]),
    ("AOI", ["AOI"]),
    ("半測", ["半測"]),
    ("防焊", ["防焊"]),
    ("文字", ["文字", "文"]),
    ("化金", ["化金"]),
    ("表面處理", ["表面處理", "面處理", "處理", "無鉛", "有鉛", "OSP", "化錫", "化銀"]),
    ("成型", ["成型", "成"]),
    ("測試", ["測試", "測"]),
    ("成檢", ["成檢", "檢"]),
    ("包裝", ["包裝", "包"]),
    ("已出貨", ["出貨"]),
]


def has_real_value(v: Any) -> bool:
    s = safe_strip(v)
    if not s:
        return False
    if s.lower() in ("nan", "none"):
        return False
    if s in ("0", "0.0"):
        return False
    return True


def derive_progress_from_row_values(row: pd.Series, col_map: Dict[str, List[str]]) -> str:
    progress = ""

    for step_name, aliases in PROCESS_STEPS:
        actual_cols = col_map.get(step_name, [])
        for col in actual_cols:
            if col in row.index and has_real_value(row[col]):
                progress = choose_better_progress(progress, step_name)

    return progress


def build_process_col_map(columns: List[str]) -> Dict[str, List[str]]:
    result: Dict[str, List[str]] = {}
    for step_name, aliases in PROCESS_STEPS:
        matched = []
        for col in columns:
            c = normalize_space(str(col)).lower()
            for alias in aliases:
                a = normalize_space(alias).lower()
                if a and a in c:
                    matched.append(col)
                    break
        result[step_name] = list(dict.fromkeys(matched))
    return result


# =========================================================
# FILE PARSERS
# =========================================================
def try_read_excel_all_sheets(uploaded_file) -> List[Tuple[str, pd.DataFrame]]:
    data = uploaded_file.read()
    uploaded_file.seek(0)

    ext = os.path.splitext(uploaded_file.name)[1].lower()
    engine = None
    if ext == ".xlsx":
        engine = "openpyxl"
    elif ext == ".xls":
        engine = "xlrd"

    dfs = []
    try:
        xls = pd.ExcelFile(io.BytesIO(data), engine=engine)
        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(io.BytesIO(data), sheet_name=sheet_name, header=None, engine=engine)
                dfs.append((sheet_name, df))
            except Exception:
                pass
    except Exception:
        try:
            df = pd.read_csv(io.BytesIO(data), header=None, encoding="utf-8")
            dfs.append(("CSV", df))
        except Exception:
            try:
                df = pd.read_csv(io.BytesIO(data), header=None, encoding="big5", errors="ignore")
                dfs.append(("CSV", df))
            except Exception:
                pass

    uploaded_file.seek(0)
    return dfs


def parse_ls_style(uploaded_file) -> List[Dict[str, Any]]:
    sheets = try_read_excel_all_sheets(uploaded_file)
    results = []

    for sheet_name, df_raw in sheets:
        if df_raw.empty:
            continue

        # 針對這種表頭直接在第一列
        df = df_raw.copy()
        if len(df) < 2:
            continue

        header = [safe_strip(x) for x in df.iloc[0].tolist()]
        if not any("Cust. P / N" in h or "LS P/N" in h or "WIP" in h for h in header):
            continue

        body = df.iloc[1:].copy()
        body.columns = header
        body = clean_dataframe(body)

        for _, row in body.iterrows():
            order_no = first_existing_field(row.to_dict(), ["PO", "PO NO", "PO No"])
            part_no = first_existing_field(row.to_dict(), ["Cust. P / N", "Customer P/N", "P/N"])
            factory_part_no = first_existing_field(row.to_dict(), ["LS P/N", "Factory P/N"])
            qty = first_existing_field(row.to_dict(), ["Q'TY", "QTY", "Qty"])
            progress = normalize_progress_text(first_existing_field(row.to_dict(), ["WIP", "進度"]))
            note = first_existing_field(row.to_dict(), ["Note", "備註"])
            due_date = first_existing_field(row.to_dict(), ["Required Ship date", "confrimed DD"])

            if not any([order_no, part_no, factory_part_no]):
                continue

            results.append({
                "source_file": uploaded_file.name,
                "source_sheet": sheet_name,
                "factory_name": "LS/PG",
                "order_no": order_no,
                "part_no": part_no,
                "factory_part_no": factory_part_no,
                "qty": qty,
                "progress": progress,
                "due_date": due_date,
                "note": note,
            })

    return results


def parse_xituo_simple(uploaded_file) -> List[Dict[str, Any]]:
    sheets = try_read_excel_all_sheets(uploaded_file)
    results = []

    for sheet_name, df_raw in sheets:
        if df_raw.empty:
            continue

        # 找到包含料號/進度的 header row
        header_row = None
        for i in range(min(len(df_raw), 10)):
            row_vals = [safe_strip(x) for x in df_raw.iloc[i].tolist()]
            joined = " | ".join(row_vals)
            if "料" in joined and "進度" in joined:
                header_row = i
                break

        if header_row is None:
            continue

        header = [safe_strip(x) for x in df_raw.iloc[header_row].tolist()]
        body = df_raw.iloc[header_row + 1:].copy()
        body.columns = header
        body = clean_dataframe(body)

        for _, row in body.iterrows():
            rowd = row.to_dict()
            part_no = first_existing_field(rowd, ["料         號", "料號", "料 號", "品號"])
            qty = first_existing_field(rowd, ["數量(PCS)", "數量", "QTY", "Qty"])
            order_date = first_existing_field(rowd, ["下單日期"])
            due_date = first_existing_field(rowd, ["出貨日期", "交期"])
            progress = normalize_progress_text(first_existing_field(rowd, ["進度", "WIP"]))
            note = first_existing_field(rowd, ["備註"])

            if not part_no:
                continue

            results.append({
                "source_file": uploaded_file.name,
                "source_sheet": sheet_name,
                "factory_name": "西拓",
                "order_no": "",
                "part_no": part_no,
                "factory_part_no": "",
                "qty": qty,
                "progress": progress,
                "due_date": due_date or order_date,
                "note": note,
            })

    return results


def parse_quanxing_process(uploaded_file) -> List[Dict[str, Any]]:
    sheets = try_read_excel_all_sheets(uploaded_file)
    results = []

    for sheet_name, df_raw in sheets:
        if df_raw.empty or len(df_raw) < 6:
            continue

        # 這類型第 3,4 列合成表頭
        header1 = [safe_strip(x) for x in df_raw.iloc[3].tolist()] if len(df_raw) > 3 else []
        header2 = [safe_strip(x) for x in df_raw.iloc[4].tolist()] if len(df_raw) > 4 else []
        if not header1 or not header2:
            continue

        combined_cols = []
        for a, b in zip(header1, header2):
            col = normalize_space(f"{a} {b}")
            combined_cols.append(col)

        joined_cols = " | ".join(combined_cols)
        if "訂 單 號 碼" not in joined_cols and "客 料 號" not in joined_cols and "出 貨" not in joined_cols:
            continue

        body = df_raw.iloc[5:].copy()
        body.columns = combined_cols
        body = clean_dataframe(body)

        col_map = build_process_col_map(list(body.columns))

        order_col = find_first_matching_col(list(body.columns), ["訂 單 號", "訂單號", "P/O"])
        part_col = find_first_matching_col(list(body.columns), ["客 料 號", "客料號", "料號"])
        qty_col = find_first_matching_col(list(body.columns), ["訂購量", "(PCS)", "數量"])
        due_col = find_first_matching_col(list(body.columns), ["交貨 日期", "交貨", "日期"])
        note_col = find_first_matching_col(list(body.columns), ["備註", "備 註"])

        for _, row in body.iterrows():
            order_no = safe_strip(row[order_col]) if order_col and order_col in row.index else ""
            part_no = safe_strip(row[part_col]) if part_col and part_col in row.index else ""
            qty = safe_strip(row[qty_col]) if qty_col and qty_col in row.index else ""
            due_date = safe_strip(row[due_col]) if due_col and due_col in row.index else ""
            note = safe_strip(row[note_col]) if note_col and note_col in row.index else ""

            if not any([order_no, part_no]):
                continue

            progress = derive_progress_from_row_values(row, col_map)
            progress = normalize_progress_text(progress)

            results.append({
                "source_file": uploaded_file.name,
                "source_sheet": sheet_name,
                "factory_name": "全興/西拓流程表",
                "order_no": order_no,
                "part_no": part_no,
                "factory_part_no": "",
                "qty": qty,
                "progress": progress,
                "due_date": due_date,
                "note": note,
            })

    return results


def parse_xianghong_wip(uploaded_file) -> List[Dict[str, Any]]:
    sheets = try_read_excel_all_sheets(uploaded_file)
    results = []

    for sheet_name, df_raw in sheets:
        if df_raw.empty or len(df_raw) < 4:
            continue

        header_row = 2
        header = [normalize_space(safe_strip(x).replace("\n", "")) for x in df_raw.iloc[header_row].tolist()]
        joined = " | ".join(header)

        if "訂單編號" not in joined and "祥竑料號" not in joined and "未出貨數量" not in joined:
            continue

        body = df_raw.iloc[header_row + 1:].copy()
        body.columns = header
        body = clean_dataframe(body)

        col_map = build_process_col_map(list(body.columns))

        for idx in range(0, len(body), 2):
            row_top = body.iloc[idx]
            row_bottom = body.iloc[idx + 1] if idx + 1 < len(body) else pd.Series(dtype=object)

            rowd = row_top.to_dict()

            order_no = first_existing_field(rowd, ["訂單編號"])
            factory_part_no = first_existing_field(rowd, ["祥竑料號"])
            part_no = first_existing_field(rowd, ["料號"])
            surface = first_existing_field(rowd, ["表面處理"])
            qty = first_existing_field(rowd, ["訂單數量"])
            due_date = first_existing_field(rowd, ["交貨日期"])
            note = first_existing_field(rowd, ["備註"])

            if not any([order_no, part_no, factory_part_no]):
                continue

            progress = derive_progress_from_row_values(row_bottom if not row_bottom.empty else row_top, col_map)
            progress = normalize_progress_text(progress)

            if not progress:
                # 若全部 0，但未出貨數量 > 0，至少標示待開料或備料中
                unshipped = first_existing_field(rowd, ["未出貨數量"])
                uv = try_parse_numeric(unshipped)
                if uv and uv > 0:
                    progress = "待開料"

            note2 = note
            if surface:
                note2 = f"{note2} / 表面:{surface}".strip(" /")

            results.append({
                "source_file": uploaded_file.name,
                "source_sheet": sheet_name,
                "factory_name": "祥竑",
                "order_no": order_no,
                "part_no": part_no,
                "factory_part_no": factory_part_no,
                "qty": qty,
                "progress": progress,
                "due_date": due_date,
                "note": note2,
            })

    return results


def parse_generic_excel(uploaded_file) -> List[Dict[str, Any]]:
    sheets = try_read_excel_all_sheets(uploaded_file)
    results = []

    for sheet_name, df_raw in sheets:
        if df_raw.empty:
            continue

        # 嘗試前 10 列找 header
        best_header_row = None
        best_score = -1

        for i in range(min(len(df_raw), 10)):
            row_vals = [safe_strip(x) for x in df_raw.iloc[i].tolist()]
            joined = " | ".join(row_vals).lower()
            score = 0
            for kw in ["料號", "進度", "wip", "訂單", "po", "qty", "數量"]:
                if kw in joined:
                    score += 1
            if score > best_score:
                best_score = score
                best_header_row = i

        if best_header_row is None or best_score <= 0:
            continue

        header = [normalize_space(safe_strip(x).replace("\n", "")) for x in df_raw.iloc[best_header_row].tolist()]
        body = df_raw.iloc[best_header_row + 1:].copy()
        body.columns = header
        body = clean_dataframe(body)

        part_col = find_first_matching_col(list(body.columns), ["料號", "Part", "P/N"])
        order_col = find_first_matching_col(list(body.columns), ["訂單", "PO"])
        qty_col = find_first_matching_col(list(body.columns), ["數量", "Qty", "QTY"])
        progress_col = find_first_matching_col(list(body.columns), ["進度", "WIP"])
        note_col = find_first_matching_col(list(body.columns), ["備註", "Note"])
        due_col = find_first_matching_col(list(body.columns), ["交期", "出貨", "日期"])

        for _, row in body.iterrows():
            part_no = safe_strip(row[part_col]) if part_col and part_col in row.index else ""
            order_no = safe_strip(row[order_col]) if order_col and order_col in row.index else ""
            qty = safe_strip(row[qty_col]) if qty_col and qty_col in row.index else ""
            progress = normalize_progress_text(safe_strip(row[progress_col])) if progress_col and progress_col in row.index else ""
            note = safe_strip(row[note_col]) if note_col and note_col in row.index else ""
            due_date = safe_strip(row[due_col]) if due_col and due_col in row.index else ""

            if not any([part_no, order_no]):
                continue

            results.append({
                "source_file": uploaded_file.name,
                "source_sheet": sheet_name,
                "factory_name": "Generic",
                "order_no": order_no,
                "part_no": part_no,
                "factory_part_no": "",
                "qty": qty,
                "progress": progress,
                "due_date": due_date,
                "note": note,
            })

    return results


def parse_uploaded_wip_file(uploaded_file) -> List[Dict[str, Any]]:
    filename = uploaded_file.name.lower()

    results = []

    # 依檔名與內容優先套專屬 parser
    if "祥竑" in filename:
        results = parse_xianghong_wip(uploaded_file)
        if results:
            return results

    if "西拓" in filename and ("進度表" in filename or filename.endswith(".xlsx")):
        results = parse_xituo_simple(uploaded_file)
        if results:
            return results

    if "wip" in filename or "203" in filename or "全興" in filename:
        results = parse_quanxing_process(uploaded_file)
        if results:
            return results

    if "glocom-pg" in filename or "pg" in filename or "ls" in filename:
        results = parse_ls_style(uploaded_file)
        if results:
            return results

    # fallback
    for parser in [parse_ls_style, parse_xituo_simple, parse_quanxing_process, parse_xianghong_wip, parse_generic_excel]:
        try:
            uploaded_file.seek(0)
            results = parser(uploaded_file)
            if results:
                return results
        except Exception:
            continue

    return []


# =========================================================
# BUILD UPDATE PAYLOAD
# =========================================================
def pick_first_existing_teable_field(fields: Dict[str, Any], candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in fields:
            return c
    return None


def build_update_fields(teable_fields: Dict[str, Any], parsed_row: Dict[str, Any]) -> Dict[str, Any]:
    updates = {}

    progress_col = pick_first_existing_teable_field(teable_fields, TEABLE_PROGRESS_TARGET_FIELDS)
    reply_date_col = pick_first_existing_teable_field(teable_fields, TEABLE_REPLY_DATE_FIELDS)
    note_col = pick_first_existing_teable_field(teable_fields, TEABLE_NOTE_FIELDS)
    source_col = pick_first_existing_teable_field(teable_fields, TEABLE_SOURCE_FIELDS)
    qty_col = pick_first_existing_teable_field(teable_fields, TEABLE_QTY_FIELDS)

    if progress_col and parsed_row.get("progress"):
        updates[progress_col] = parsed_row["progress"]

    if reply_date_col:
        updates[reply_date_col] = datetime.now().strftime("%Y-%m-%d")

    if note_col:
        note_text = parsed_row.get("note", "")
        src = parsed_row.get("source_file", "")
        match_detail = parsed_row.get("_match_detail", "")
        merged = " / ".join([x for x in [note_text, f"來源:{src}" if src else "", match_detail] if x])
        if merged:
            updates[note_col] = merged[:500]

    if source_col:
        updates[source_col] = parsed_row.get("factory_name", "")

    # 若主表數量為空才補
    if qty_col and parsed_row.get("qty"):
        current_qty = teable_fields.get(qty_col)
        if is_blank(current_qty):
            updates[qty_col] = parsed_row["qty"]

    return updates


# =========================================================
# DATAFRAME DISPLAY
# =========================================================
def parsed_rows_to_df(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()

    show = []
    for r in rows:
        show.append({
            "來源檔案": r.get("source_file", ""),
            "Sheet": r.get("source_sheet", ""),
            "工廠": r.get("factory_name", ""),
            "訂單號": r.get("order_no", ""),
            "客戶料號": r.get("part_no", ""),
            "工廠料號": r.get("factory_part_no", ""),
            "數量": r.get("qty", ""),
            "進度": r.get("progress", ""),
            "交期": r.get("due_date", ""),
            "備註": r.get("note", ""),
            "正規化料號": normalize_part_no(r.get("part_no", "")),
        })
    return pd.DataFrame(show)


# =========================================================
# UI
# =========================================================
st.title("🏭 GLOCOM WIP 進度匯入 / 更新 Teable 主表")
st.caption("修正版：已補上 料號正規化、流程表轉進度、Teable 主表匹配與更新邏輯")

with st.expander("設定與說明", expanded=False):
    st.write(f"**Teable API URL**: {TABLE_URL}")
    st.write(f"**Teable Web URL**: {TEABLE_WEB_URL}")
    if TEABLE_TOKEN:
        st.success("已偵測到 TEABLE_TOKEN")
    else:
        st.warning("尚未設定 TEABLE_TOKEN，無法更新 Teable。可先測試解析結果。")

    st.markdown(
        """
**這版已改善：**
- 西拓簡單進度表：直接抓料號 / 進度
- 全興 / 西拓流程表：由最右側已完成工序推算 WIP 進度
- 祥竑 WIP：由工序數量列推算當前站別
- PG / LS 表：直接抓 Cust. P/N / LS P/N / WIP
- 更新主表時，先比訂單號，再比正規化料號，再做模糊比對
        """
    )

col1, col2 = st.columns([2, 1])

with col1:
    uploaded_files = st.file_uploader(
        "上傳工廠 WIP 檔案（可多檔）",
        type=["xls", "xlsx", "csv"],
        accept_multiple_files=True
    )

with col2:
    st.link_button("Open Teable", TEABLE_WEB_URL)

if "parsed_rows_cache" not in st.session_state:
    st.session_state["parsed_rows_cache"] = []

if st.button("解析上傳檔案", type="primary", use_container_width=True):
    all_rows = []
    parse_logs = []

    if not uploaded_files:
        st.warning("請先上傳檔案。")
    else:
        for f in uploaded_files:
            try:
                rows = parse_uploaded_wip_file(f)
                if rows:
                    all_rows.extend(rows)
                    parse_logs.append(f"✅ {f.name}：解析 {len(rows)} 筆")
                else:
                    parse_logs.append(f"⚠️ {f.name}：沒有抓到有效資料")
            except Exception as e:
                parse_logs.append(f"❌ {f.name}：解析失敗 - {e}")

        st.session_state["parsed_rows_cache"] = all_rows

        for msg in parse_logs:
            st.write(msg)

rows = st.session_state.get("parsed_rows_cache", [])
df_preview = parsed_rows_to_df(rows)

st.subheader("解析結果預覽")
if not df_preview.empty:
    st.dataframe(df_preview, use_container_width=True, height=420)
else:
    st.info("尚無解析結果。請先上傳並按『解析上傳檔案』。")

with st.expander("原始解析 JSON", expanded=False):
    st.json(rows if rows else [])

st.subheader("更新 Teable 主表")

if st.button("讀取主表並比對", use_container_width=True):
    if not TEABLE_TOKEN:
        st.error("未設定 TEABLE_TOKEN，無法讀取主表。")
    elif not rows:
        st.warning("請先解析檔案。")
    else:
        with st.spinner("讀取 Teable 主表中..."):
            teable_records = fetch_teable_records(limit=5000)

        if not teable_records:
            st.error("讀不到 Teable 主表資料。")
        else:
            indexes = build_teable_indexes(teable_records)

            compare_rows = []
            matched_cache = []

            for r in rows:
                matched, reason = match_teable_record(r, indexes)
                rr = dict(r)
                rr["_match_detail"] = reason
                rr["_matched_record"] = matched
                matched_cache.append(rr)

                compare_rows.append({
                    "來源檔案": r.get("source_file", ""),
                    "訂單號": r.get("order_no", ""),
                    "客戶料號": r.get("part_no", ""),
                    "工廠料號": r.get("factory_part_no", ""),
                    "進度": r.get("progress", ""),
                    "比對結果": "命中" if matched else "未命中",
                    "說明": reason,
                })

            st.session_state["matched_rows_cache"] = matched_cache
            st.dataframe(pd.DataFrame(compare_rows), use_container_width=True, height=420)

if "matched_rows_cache" not in st.session_state:
    st.session_state["matched_rows_cache"] = []

matched_rows = st.session_state.get("matched_rows_cache", [])

if st.button("更新到 Teable 主表", type="primary", use_container_width=True):
    if not TEABLE_TOKEN:
        st.error("未設定 TEABLE_TOKEN。")
    elif not matched_rows:
        st.warning("請先按『讀取主表並比對』。")
    else:
        ok_count = 0
        fail_count = 0
        logs = []

        progress_bar = st.progress(0)
        total = len(matched_rows)

        for i, row in enumerate(matched_rows, start=1):
            matched = row.get("_matched_record")
            if not matched:
                fail_count += 1
                logs.append({
                    "來源檔案": row.get("source_file", ""),
                    "訂單號": row.get("order_no", ""),
                    "料號": row.get("part_no", ""),
                    "結果": "未更新",
                    "原因": row.get("_match_detail", "找不到主表"),
                })
                progress_bar.progress(i / total)
                continue

            rec_id = matched["id"]
            teable_fields = matched["fields"]
            update_fields = build_update_fields(teable_fields, row)

            if not update_fields:
                fail_count += 1
                logs.append({
                    "來源檔案": row.get("source_file", ""),
                    "訂單號": row.get("order_no", ""),
                    "料號": row.get("part_no", ""),
                    "結果": "未更新",
                    "原因": "找不到可更新的主表欄位名稱，請檢查 TEABLE_*_FIELDS 設定",
                })
                progress_bar.progress(i / total)
                continue

            ok, msg = update_teable_record(rec_id, update_fields)
            if ok:
                ok_count += 1
                logs.append({
                    "來源檔案": row.get("source_file", ""),
                    "訂單號": row.get("order_no", ""),
                    "料號": row.get("part_no", ""),
                    "結果": "成功",
                    "原因": row.get("_match_detail", ""),
                })
            else:
                fail_count += 1
                logs.append({
                    "來源檔案": row.get("source_file", ""),
                    "訂單號": row.get("order_no", ""),
                    "料號": row.get("part_no", ""),
                    "結果": "失敗",
                    "原因": msg,
                })

            progress_bar.progress(i / total)
            time.sleep(0.03)

        if ok_count:
            st.success(f"更新完成：成功 {ok_count} 筆，失敗 {fail_count} 筆")
        else:
            st.error(f"沒有成功更新。失敗 {fail_count} 筆")

        st.dataframe(pd.DataFrame(logs), use_container_width=True, height=420)

st.divider()

st.subheader("除錯工具")

with st.expander("手動測試料號正規化", expanded=False):
    test_part = st.text_input("輸入料號")
    if test_part:
        st.write("原始：", test_part)
        st.write("正規化：", normalize_part_no(test_part))

with st.expander("建議檢查項目", expanded=False):
    st.markdown(
        """
1. 先確認 **TEABLE_TOKEN** 是否正確  
2. 確認 Teable 主表實際欄位名稱，是否與下列候選欄位相符：  
   - 訂單號 / PO / Order No  
   - 料號 / 客戶料號 / Part No / P/N  
   - 進度 / WIP / WIP進度  
3. 若主表欄位名稱不同，直接修改本檔最上方這幾個清單：  
   - `TEABLE_MATCH_FIELDS`
   - `TEABLE_PART_FIELDS`
   - `TEABLE_FACTORY_PART_FIELDS`
   - `TEABLE_PROGRESS_TARGET_FIELDS`
4. 若某工廠有特殊格式，再補一個專屬 parser 即可
        """
    )
