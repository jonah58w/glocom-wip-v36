# -*- coding: utf-8 -*-
"""
text_ocr_parsers.py
OCR / 貼上文字 / email 文字解析工具

輸出統一欄位：
PO#, Part No, Qty, Factory Due Date, Ship Date, WIP, Remark, Customer Remark Tags
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

import pandas as pd

from utils import safe_text, compact_text, normalize_due_text, normalize_columns


# =========================================================
# 共用規則
# =========================================================
WIP_KEYWORDS = {
    "shipping": "Shipping",
    "shipped": "Shipping",
    "ship": "Shipping",
    "packing": "Packing",
    "pack": "Packing",
    "inspection": "Inspection",
    "inspect": "Inspection",
    "fqc": "Inspection",
    "oqc": "Inspection",
    "iqc": "Inspection",
    "qa": "Inspection",
    "production": "Production",
    "prod": "Production",
    "engineering": "Engineering",
    "engineer": "Engineering",
    "gerber": "Engineering",
    "hold": "On Hold",
    "on hold": "On Hold",
    "pending": "On Hold",
    "waiting": "On Hold",
    "完成": "完成",
    "出貨": "Shipping",
    "已出貨": "Shipping",
    "包裝": "Packing",
    "檢驗": "Inspection",
    "測試": "Inspection",
    "工程": "Engineering",
    "待料": "On Hold",
    "暫停": "On Hold",
    "進行中": "Production",
    "生產中": "Production",
}

TAG_HINTS = {
    "gerber": "Working Gerber for Approval",
    "working gerber": "Working Gerber for Approval",
    "approval": "Waiting Confirmation",
    "eq": "Engineering Question",
    "engineering question": "Engineering Question",
    "payment": "Payment Pending",
    "pending payment": "Payment Pending",
    "remake": "Remake in Process",
    "hold": "On Hold",
    "partial": "Partial Shipment",
    "shipped": "Shipped",
    "waiting confirmation": "Waiting Confirmation",
}

PO_PATTERNS = [
    r"\bPO[#:\s-]*([A-Z0-9][A-Z0-9\-_\/\.]+)\b",
    r"\bP\/O[#:\s-]*([A-Z0-9][A-Z0-9\-_\/\.]+)\b",
    r"訂單編號[:：\s]*([A-Z0-9][A-Z0-9\-_\/\.]+)",
    r"訂單號(?:碼)?[:：\s]*([A-Z0-9][A-Z0-9\-_\/\.]+)",
    r"工單號?[:：\s]*([A-Z0-9][A-Z0-9\-_\/\.]+)",
]

PART_PATTERNS = [
    r"\bPart\s*No\.?[:：\s]*([A-Z0-9][A-Z0-9\-_\/\.]+)\b",
    r"\bP\/N[:：\s]*([A-Z0-9][A-Z0-9\-_\/\.]+)\b",
    r"客戶料號[:：\s]*([A-Z0-9][A-Z0-9\-_\/\.]+)",
    r"料號[:：\s]*([A-Z0-9][A-Z0-9\-_\/\.]+)",
    r"LS\s*P\/N[:：\s]*([A-Z0-9][A-Z0-9\-_\/\.]+)",
    r"Cust\.?\s*P\s*\/\s*N[:：\s]*([A-Z0-9][A-Z0-9\-_\/\.]+)",
]

QTY_PATTERNS = [
    r"\bQty[:：\s]*([0-9][0-9,\.]*)\b",
    r"\bQTY[:：\s]*([0-9][0-9,\.]*)\b",
    r"\bQ'TY[:：\s]*([0-9][0-9,\.]*)\b",
    r"數量[:：\s]*([0-9][0-9,\.]*)",
    r"訂單量(?:\(PCS\))?[:：\s]*([0-9][0-9,\.]*)",
    r"未出貨數量[:：\s]*([0-9][0-9,\.]*)",
]

DATE_PATTERNS = [
    r"\b\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}\b",
    r"\b\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{4}\b",
    r"\b\d{4}年\d{1,2}月\d{1,2}日\b",
]

DUE_LABEL_PATTERNS = [
    r"(?:Required Ship date|Required Ship Date)[:：\s]*([^\n\r]+)",
    r"(?:Ship Date|Ship date)[:：\s]*([^\n\r]+)",
    r"(?:交貨日期|出貨日期|交期|預交日|預定交期)[:：\s]*([^\n\r]+)",
    r"(?:confirmed DD|confrimed DD)[:：\s]*([^\n\r]+)",
]

ROW_SPLIT_HINTS = [
    r"\n\s*\n",
    r"\n(?=PO\b)",
    r"\n(?=P\/O\b)",
    r"\n(?=訂單編號)",
    r"\n(?=訂單號)",
    r"\n(?=Part\s*No)",
    r"\n(?=料號)",
]


# =========================================================
# 基本工具
# =========================================================
def _clean_line(line: Any) -> str:
    s = compact_text(line)
    s = s.replace("：", ":")
    return s.strip()


def _clean_block(text: Any) -> str:
    s = safe_text(text)
    if not s:
        return ""
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("\u3000", " ")
    return s


def _search_first(patterns: List[str], text: str, flags=re.IGNORECASE) -> str:
    for pat in patterns:
        m = re.search(pat, text, flags)
        if m:
            return safe_text(m.group(1))
    return ""


def _search_first_date(text: str) -> str:
    # 先抓帶標籤的日期
    due_text = _search_first(DUE_LABEL_PATTERNS, text)
    if due_text:
        # 從 due_text 中再抽日期
        for pat in DATE_PATTERNS:
            m = re.search(pat, due_text)
            if m:
                return normalize_due_text(m.group(0))
        return normalize_due_text(due_text)

    # 再抓任意日期
    for pat in DATE_PATTERNS:
        m = re.search(pat, text)
        if m:
            return normalize_due_text(m.group(0))
    return ""


def _normalize_qty(value: Any) -> str:
    s = safe_text(value).replace(",", "")
    if not s:
        return ""
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return str(f)
    except Exception:
        return safe_text(value)


def _detect_wip(text: str) -> str:
    low = safe_text(text).lower()
    if not low:
        return ""

    # 先長詞優先
    for k in sorted(WIP_KEYWORDS.keys(), key=len, reverse=True):
        if k.lower() in low:
            return WIP_KEYWORDS[k]
    return ""


def _detect_tags(text: str) -> List[str]:
    low = safe_text(text).lower()
    tags = []

    for k, tag in TAG_HINTS.items():
        if k.lower() in low and tag not in tags:
            tags.append(tag)

    # 由 WIP 補 tag
    wip = _detect_wip(text)
    if wip == "Shipping" and "Shipped" not in tags:
        tags.append("Shipped")
    if wip == "On Hold" and "On Hold" not in tags:
        tags.append("On Hold")

    return tags


def _make_row(
    po: Any = "",
    part: Any = "",
    qty: Any = "",
    due: Any = "",
    ship: Any = "",
    wip: Any = "",
    remark: Any = "",
    tags: Optional[List[str]] = None,
    source_type: str = "text_ocr",
) -> Dict[str, Any]:
    return {
        "PO#": safe_text(po),
        "Part No": safe_text(part),
        "Qty": _normalize_qty(qty),
        "Factory Due Date": normalize_due_text(due),
        "Ship Date": normalize_due_text(ship if ship else due),
        "WIP": safe_text(wip),
        "Remark": safe_text(remark),
        "Customer Remark Tags": tags or [],
        "_source_sheet": "",
        "_source_type": source_type,
    }


# =========================================================
# 單段文字解析
# =========================================================
def parse_text_block_to_row(text: str, source_type: str = "text_block") -> Optional[Dict[str, Any]]:
    """
    將一段文字盡量解析成單筆資料
    """
    block = _clean_block(text)
    if not block:
        return None

    po_val = _search_first(PO_PATTERNS, block)
    part_val = _search_first(PART_PATTERNS, block)
    qty_val = _search_first(QTY_PATTERNS, block)
    due_val = _search_first_date(block)
    wip_val = _detect_wip(block)
    tags = _detect_tags(block)

    if not po_val and not part_val:
        return None

    return _make_row(
        po=po_val,
        part=part_val,
        qty=qty_val,
        due=due_val,
        ship=due_val,
        wip=wip_val or "Production",
        remark=block[:500],
        tags=tags,
        source_type=source_type,
    )


# =========================================================
# 多段文字切塊
# =========================================================
def split_text_into_blocks(text: str) -> List[str]:
    """
    將 email / OCR / 貼上文字切成多筆候選區塊
    """
    raw = _clean_block(text)
    if not raw:
        return []

    blocks = [raw]

    for pat in ROW_SPLIT_HINTS:
        new_blocks = []
        for blk in blocks:
            pieces = re.split(pat, blk, flags=re.IGNORECASE)
            new_blocks.extend(pieces)
        blocks = new_blocks

    cleaned = []
    for blk in blocks:
        b = blk.strip()
        if b:
            cleaned.append(b)

    # 去除過短純噪音
    filtered = []
    for blk in cleaned:
        line_count = len([x for x in blk.split("\n") if x.strip()])
        if len(blk) >= 8 or line_count >= 1:
            filtered.append(blk)

    return filtered


# =========================================================
# 多筆 email / OCR 解析
# =========================================================
def parse_email_text_to_rows(text: str) -> List[Dict[str, Any]]:
    """
    主要給 factory_parsers.parse_txt_file() 使用
    """
    raw = _clean_block(text)
    if not raw:
        return []

    rows: List[Dict[str, Any]] = []

    # 先按區塊拆
    blocks = split_text_into_blocks(raw)

    for blk in blocks:
        row = parse_text_block_to_row(blk, source_type="email_text")
        if row:
            rows.append(row)

    # 如果區塊解析不到，嘗試逐行累積
    if not rows:
        lines = [_clean_line(x) for x in raw.split("\n") if _clean_line(x)]
        buffer = []
        for line in lines:
            buffer.append(line)
            joined = "\n".join(buffer)

            # 遇到新 PO/Part 開頭時，先輸出前一段
            if len(buffer) > 1 and (
                re.search(r"^(PO|P\/O|訂單編號|訂單號|Part\s*No|料號)\b", line, re.IGNORECASE)
            ):
                prev = "\n".join(buffer[:-1])
                row = parse_text_block_to_row(prev, source_type="email_text")
                if row:
                    rows.append(row)
                buffer = [line]

        if buffer:
            row = parse_text_block_to_row("\n".join(buffer), source_type="email_text")
            if row:
                rows.append(row)

    # 最後去重
    deduped = deduplicate_rows(rows)
    return deduped


# =========================================================
# OCR 專用
# =========================================================
def parse_ocr_text_to_rows(text: str) -> List[Dict[str, Any]]:
    """
    OCR 輸出常常有斷行 / 噪音，這裡先做簡單容錯再共用 email parser
    """
    raw = _clean_block(text)
    if not raw:
        return []

    # OCR 常見誤差修正
    replacements = {
        "P O": "PO",
        "P / O": "P/O",
        "ShipDate": "Ship Date",
        "RequiredShipdate": "Required Ship date",
        "PartNo": "Part No",
        "Q TY": "QTY",
    }
    for old, new in replacements.items():
        raw = raw.replace(old, new)

    return parse_email_text_to_rows(raw)


# =========================================================
# DataFrame / 去重
# =========================================================
def deduplicate_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    依 PO# + Part No 去重，保留資訊較完整的一筆
    """
    if not rows:
        return []

    best_map: Dict[str, Dict[str, Any]] = {}

    for row in rows:
        po = safe_text(row.get("PO#", ""))
        part = safe_text(row.get("Part No", ""))
        key = f"{po}||{part}"

        score = 0
        for k in ["PO#", "Part No", "Qty", "Factory Due Date", "Ship Date", "WIP", "Remark"]:
            if safe_text(row.get(k, "")):
                score += 1

        if key not in best_map:
            row["_score"] = score
            best_map[key] = row
        else:
            old_score = best_map[key].get("_score", 0)
            if score > old_score:
                row["_score"] = score
                best_map[key] = row

    out = []
    for row in best_map.values():
        row.pop("_score", None)
        out.append(row)

    return out


def rows_to_dataframe(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=[
            "PO#",
            "Part No",
            "Qty",
            "Factory Due Date",
            "Ship Date",
            "WIP",
            "Remark",
            "Customer Remark Tags",
            "_source_sheet",
            "_source_type",
        ])
    return normalize_columns(pd.DataFrame(rows))


# =========================================================
# 對外統一入口
# =========================================================
def parse_text_to_dataframe(text: str, source_type: str = "text") -> pd.DataFrame:
    if source_type.lower() == "ocr":
        rows = parse_ocr_text_to_rows(text)
    else:
        rows = parse_email_text_to_rows(text)
    return rows_to_dataframe(rows)


def parse_single_text_row(text: str, source_type: str = "text") -> pd.DataFrame:
    row = parse_text_block_to_row(text, source_type=source_type)
    if not row:
        return rows_to_dataframe([])
    return rows_to_dataframe([row])
