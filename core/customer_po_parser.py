# -*- coding: utf-8 -*-
"""
客戶 PO PDF 解析器 v2。

支援的客戶:WESCO / TIETO / GUDE / KCS,其他走通用 fallback。
PDF 文字提取用 pdfplumber(layout-aware,比 pypdf 準很多)。
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


@dataclass
class POItem:
    line: str = ""
    part_number: str = ""
    description: str = ""
    quantity: int = 0
    unit_price: float = 0.0
    amount: float = 0.0
    delivery_date: Optional[str] = None  # ISO YYYY-MM-DD


@dataclass
class ParsedPO:
    customer_name: str = ""
    customer_po_no: str = ""
    po_date: str = ""
    payment_terms: str = ""
    ship_to: str = ""
    ship_via: str = ""
    currency: str = "USD"
    items: list[POItem] = field(default_factory=list)
    total_amount: float = 0.0
    issuing_company_detected: str = "GLOCOM"
    parser_used: str = ""
    raw_text: str = ""
    parse_warnings: list[str] = field(default_factory=list)


def detect_customer(text: str) -> str:
    t = text.upper()
    if "WESCO" in t:
        return "WESCO"
    if "TIETO-OSKARI" in t or ("TIETO" in t and "KAJAANI" in t):
        return "TIETO"
    if "GUDE SYSTEMS" in t or ("GUDE" in t and "KÖLN" in t):
        return "GUDE"
    if "KCS BV" in t or "TRACE.ME" in t or ("KCS" in t and "DORDRECHT" in t):
        return "KCS"
    if "ECDATA" in t:
        return "ECDATA"
    if "VORNE" in t:
        return "VORNE"
    return ""


def detect_issuing_company(text: str) -> str:
    t = text.upper()
    if "EUSWAY" in t:
        return "EUSWAY"
    if "GLOCOM" in t:
        return "GLOCOM"
    return "GLOCOM"


def _parse_iso_date(s: str) -> str:
    s = s.strip()
    fmts = [
        "%m/%d/%Y", "%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d",
        "%d/%m/%Y", "%d.%m.%y", "%d-%m-%y",
    ]
    for f in fmts:
        try:
            dt = datetime.strptime(s, f)
            if dt.year < 2000:
                dt = dt.replace(year=dt.year + 2000)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s


def _eu_num(s: str) -> float:
    s = s.strip().replace(" ", "").replace("\u00a0", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _us_num(s: str) -> float:
    s = s.replace("$", "").replace(",", "").strip()
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


# ─── WESCO ────────────────────────────────────
WESCO_ITEM_RE = re.compile(
    r"^(\d+)\s+(\d{1,2}/\d{1,2}/\d{4})\s+(.+?)\s+([\d,]+)\s+\w+\.?\s+\$([\d,.]+)\s+\$([\d,.]+)\s*$",
    re.M
)


def parse_wesco(text: str) -> ParsedPO:
    p = ParsedPO(parser_used="WESCO", raw_text=text, customer_name="WESCO", currency="USD")

    m = re.search(r"PO Number:\s*(PO\d+)", text)
    if not m:
        m = re.search(r"\b(PO\d{4,})\b", text)
    if m:
        p.customer_po_no = m.group(1)

    m = re.search(r"PO Date:\s*(\d{1,2}/\d{1,2}/\d{4})", text)
    if m:
        p.po_date = _parse_iso_date(m.group(1))

    m = re.search(r"(Net\s+\d+)", text, re.I)
    if m:
        p.payment_terms = m.group(1).strip()

    m = re.search(r"\bTotal:\s+\$([\d,.]+)", text)
    if m:
        p.total_amount = _us_num(m.group(1))

    m = re.search(r"Please ship\s+([^\n.]+?)(?:\.|\s*Thank)", text, re.I)
    if m:
        p.ship_via = m.group(1).strip()

    lines = text.split("\n")
    desc_lookup = {}
    for i, line in enumerate(lines):
        if WESCO_ITEM_RE.match(line):
            for j in range(i + 1, min(i + 3, len(lines))):
                if lines[j].strip().startswith("PCB"):
                    desc_lookup[i] = lines[j].strip()
                    break

    for i, line in enumerate(lines):
        m = WESCO_ITEM_RE.match(line)
        if m:
            p.items.append(POItem(
                line=m.group(1),
                part_number=m.group(3).strip(),
                description=desc_lookup.get(i, ""),
                delivery_date=_parse_iso_date(m.group(2)),
                quantity=int(m.group(4).replace(",", "")),
                unit_price=_us_num(m.group(5)),
                amount=_us_num(m.group(6)),
            ))

    if not p.items:
        p.parse_warnings.append("WESCO: 沒解出品項,請手動補")
    return p


# ─── TIETO ────────────────────────────────────
TIETO_ITEM_RE = re.compile(
    r"^(\d+)\s+(\S+)\s+(.+?)\s+([\d, ]+)\s+pcs\s+([\d,]+)\s+([\d, ]+,\d{2})\s+(\d{1,2}\.\d{1,2}\.\d{4})\s*$",
    re.M
)


def parse_tieto(text: str) -> ParsedPO:
    p = ParsedPO(parser_used="TIETO", raw_text=text, customer_name="TIETO", currency="USD")

    m = re.search(r"\b(\d{4})\s+(\d{1,2}\.\d{1,2}\.\d{4})\s+\d+\s*\(\d", text)
    if m:
        p.customer_po_no = m.group(1)
        p.po_date = _parse_iso_date(m.group(2))

    m = re.search(r"Terms of payment\s+([^\n]+)", text)
    if m:
        p.payment_terms = m.group(1).strip()

    m = re.search(r"Method of delivery\s+([^\n]+)", text)
    if m:
        p.ship_via = m.group(1).strip()

    m = re.search(r"Total amount\s+\w+\s+([\d,. ]+)", text)
    if m:
        p.total_amount = _eu_num(m.group(1))

    for m in TIETO_ITEM_RE.finditer(text):
        try:
            qty = int(_eu_num(m.group(4)))
        except ValueError:
            qty = 0
        p.items.append(POItem(
            line=m.group(1),
            part_number=m.group(2),
            description=m.group(3).strip(),
            quantity=qty,
            unit_price=_eu_num(m.group(5)),
            amount=_eu_num(m.group(6)),
            delivery_date=_parse_iso_date(m.group(7)),
        ))

    if not p.items:
        p.parse_warnings.append("TIETO: 沒解出品項,請手動補")
    return p


# ─── GUDE ────────────────────────────────────
GUDE_ITEM_RE = re.compile(
    r"^(\d+)\s+(\S+)\s+(.+?)\s+(\d{1,2}\.\d{1,2}\.\d{4})\s+(\d+)\s*pcs\.?\s+([\d,]+)\s+([\d,.]+)\s*$",
    re.M
)


def parse_gude(text: str) -> ParsedPO:
    p = ParsedPO(parser_used="GUDE", raw_text=text, customer_name="GUDE", currency="USD")

    m = re.search(r"Document no\.\s+(\d{4}-\d+)", text)
    if not m:
        m = re.search(r"\b(\d{4}-\d{4,5})\b", text)
    if m:
        p.customer_po_no = m.group(1)

    # GUDE 的 Date 在 Document no 那行右邊
    m = re.search(r"Date\s+(\d{1,2}\.\d{1,2}\.\d{4})", text)
    if m:
        p.po_date = _parse_iso_date(m.group(1))

    m = re.search(r"(\d+\s*days[^\n]*?)\s+without", text, re.I)
    if not m:
        m = re.search(r"^(\d+\s*days)\b", text, re.I | re.M)
    if m:
        p.payment_terms = m.group(1).strip()

    m = re.search(r"Total US\$\s+([\d,.]+)", text)
    if m:
        p.total_amount = _eu_num(m.group(1))

    for m in GUDE_ITEM_RE.finditer(text):
        p.items.append(POItem(
            line=m.group(1),
            part_number=m.group(2),
            description=m.group(3).strip(),
            delivery_date=_parse_iso_date(m.group(4)),
            quantity=int(m.group(5)),
            unit_price=_eu_num(m.group(6)),
            amount=_eu_num(m.group(7)),
        ))

    if not p.items:
        p.parse_warnings.append("GUDE: 沒解出品項,請手動補")
    return p


# ─── KCS ────────────────────────────────────
KCS_ITEM_RE = re.compile(
    r"^(\d+)\s+(\S+)\s+(.+?)\s+(\S+)\s+(\d+)\s+(\d{1,2}-\d{1,2}-\d{2,4})\s+\$\s*([\d.,]+)\s+\$\s*([\d.,]+)\s*$",
    re.M
)


def parse_kcs(text: str) -> ParsedPO:
    p = ParsedPO(parser_used="KCS", raw_text=text, customer_name="KCS", currency="USD")

    # KCS 的 PO 號:Purchase order 那段下面的 6xxxxx
    m = re.search(r"\b(6\d{5})\b", text)
    if m:
        p.customer_po_no = m.group(1)

    m = re.search(r"Order date:\s*(\d{1,2}-\d{1,2}-\d{2,4})", text)
    if m:
        p.po_date = _parse_iso_date(m.group(1))

    m = re.search(r"Total excl\.?\s+VAT\s+\$\s*([\d.,]+)", text, re.I)
    if m:
        p.total_amount = _eu_num(m.group(1))

    for m in KCS_ITEM_RE.finditer(text):
        # KCS 單價是 5,75000(歐式 5 位小數),只取前 2 位小數
        unit_raw = m.group(7).replace(" ", "")
        if "," in unit_raw and "." not in unit_raw:
            int_part, dec_part = unit_raw.split(",", 1)
            dec_part = dec_part[:2].ljust(2, "0")
            try:
                unit_price = float(f"{int_part}.{dec_part}")
            except ValueError:
                unit_price = _eu_num(unit_raw)
        else:
            unit_price = _eu_num(unit_raw)

        p.items.append(POItem(
            line=m.group(1),
            part_number=m.group(2),
            description=m.group(3).strip(),
            quantity=int(m.group(5)),
            delivery_date=_parse_iso_date(m.group(6)),
            unit_price=unit_price,
            amount=_eu_num(m.group(8)),
        ))

    if not p.items:
        p.parse_warnings.append("KCS: 沒解出品項,請手動補")
    return p


def parse_generic(text: str) -> ParsedPO:
    p = ParsedPO(parser_used="GENERIC", raw_text=text)
    p.parse_warnings.append("無法辨識客戶格式,請手動輸入所有欄位")
    return p


def parse_customer_po(text: str) -> ParsedPO:
    customer = detect_customer(text)
    parsers = {
        "WESCO": parse_wesco, "TIETO": parse_tieto,
        "GUDE": parse_gude, "KCS": parse_kcs,
    }
    parser_fn = parsers.get(customer, parse_generic)
    result = parser_fn(text)
    result.issuing_company_detected = detect_issuing_company(text)
    return result


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """用 pdfplumber 提取文字。"""
    import pdfplumber
    pages = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            pages.append(t)
    return "\n".join(pages)
