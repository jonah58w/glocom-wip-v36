# -*- coding: utf-8 -*-
"""
客戶 PO PDF 解析器 v6 (2026-04-27)。

支援的客戶:WESCO / TIETO / GUDE / KCS / VORNE,其他走通用 fallback。

v6 變更:
- TIETO 支援新版欄序 (6053+):
  Pos | Code | Item name | Amount | Unit (pcs/kpl) | Unit price | [Disc.%] | Total | Time
- TIETO 仍相容舊版 (5976/5988/6043):
  Pos | Code | Item name | Amount | pcs | Unit price | Total | Time
- 單位識別擴展:pcs / kpl / 任何 2-4 個英文字母

WESCO/TIETO/GUDE/KCS 用 pdfplumber.extract_text() + regex。
VORNE 是表格型 PDF,用 pdfplumber.extract_tables() 直接抓表格。

主入口:
- parse_customer_po_from_pdf(pdf_bytes)  ← 推薦,從 PDF bytes 直接解析
- parse_customer_po(text)                  ← 舊介面,只接受文字
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


# ─── 偵測 ────────────────────────────────────────
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
    if "VORNE INDUSTRIES" in t or "VORNE.COM" in t:
        return "VORNE"
    if "ECDATA" in t:
        return "ECDATA"
    return ""


def detect_issuing_company(text: str) -> str:
    t = text.upper()
    if "EUSWAY" in t:
        return "EUSWAY"
    if "GLOCOM" in t:
        return "GLOCOM"
    return "GLOCOM"


# ─── 共用工具 ─────────────────────────────────────
def _parse_iso_date(s: str) -> str:
    s = (s or "").strip()
    fmts = [
        "%m/%d/%Y", "%m/%d/%y",
        "%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d",
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
    """處理歐洲數字格式 '1 005,00' / '300,00' / '3,350'"""
    s = (s or "").strip().replace(" ", "").replace("\u00a0", "")
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
    s = (s or "").replace("$", "").replace(",", "").strip()
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


# ─── WESCO ───────────────────────────────────────
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


# ─── TIETO (v6: 同時支援新舊欄序) ──────────────
# 新版欄序 (6053+): Pos | Code | Item name | Amount | Unit | Unit price | Disc.% | Total | Time
# 範例: "1 00006713v1.1 REV 2 FR4/2L/1,6/1pp/LF HAL/Bronto 300,00 kpl 3,350 1 005,00 27.4.2026"
TIETO_ITEM_NEW_RE = re.compile(
    r"^(\d+)\s+"                              # 1: pos
    r"(\S+)\s+"                                # 2: code
    r"(.+?)\s+"                                # 3: item name
    r"([\d, ]+,\d{2})\s+"                      # 4: amount (歐式: 300,00)
    r"(pcs|kpl|kpl\.|pcs\.|[a-zA-Z]{2,4})\s+"  # 5: unit (pcs/kpl/...)
    r"([\d,]+)\s+"                             # 6: unit price (3,350)
    r"([\d, ]+,\d{2})\s+"                      # 7: total (1 005,00)
    r"(\d{1,2}\.\d{1,2}\.\d{4})\s*$",          # 8: time
    re.M
)

# 舊版欄序 (5976/5988/6043): Pos | Code | Item name | Amount | pcs | Unit price | Total | Time
# 用同一個 regex 也能 match,因為 unit 改寬鬆了
# 但保留作為文件參考
TIETO_ITEM_OLD_RE = re.compile(
    r"^(\d+)\s+(\S+)\s+(.+?)\s+([\d, ]+)\s+pcs\s+([\d,]+)\s+([\d, ]+,\d{2})\s+(\d{1,2}\.\d{1,2}\.\d{4})\s*$",
    re.M
)


def parse_tieto(text: str) -> ParsedPO:
    p = ParsedPO(parser_used="TIETO", raw_text=text, customer_name="TIETO", currency="USD")

    # PO 號 + 日期
    m = re.search(r"\b(\d{4})\s+(\d{1,2}\.\d{1,2}\.\d{4})\s+\d+\s*\(\d", text)
    if m:
        p.customer_po_no = m.group(1)
        p.po_date = _parse_iso_date(m.group(2))

    # 付款條件
    m = re.search(r"Terms of payment\s+([^\n]+)", text)
    if m:
        p.payment_terms = m.group(1).strip()

    # 出貨方式
    m = re.search(r"Method of delivery\s+([^\n]+)", text)
    if m:
        p.ship_via = m.group(1).strip()

    # Total amount (新舊都同樣的 pattern: "Total amount USD 1005,00" 或 "Total amount EUR 5,143.00")
    m = re.search(r"Total amount\s+\w+\s+([\d,. ]+)", text)
    if m:
        p.total_amount = _eu_num(m.group(1))

    # 品項 - 先試新版欄序(更寬鬆,涵蓋 pcs / kpl 等)
    matches_new = list(TIETO_ITEM_NEW_RE.finditer(text))
    if matches_new:
        for m in matches_new:
            try:
                qty = int(_eu_num(m.group(4)))
            except (ValueError, TypeError):
                qty = 0
            p.items.append(POItem(
                line=m.group(1),
                part_number=m.group(2),
                description=m.group(3).strip(),
                quantity=qty,
                unit_price=_eu_num(m.group(6)),
                amount=_eu_num(m.group(7)),
                delivery_date=_parse_iso_date(m.group(8)),
            ))
    else:
        # Fallback: 試舊版欄序
        for m in TIETO_ITEM_OLD_RE.finditer(text):
            try:
                qty = int(_eu_num(m.group(4)))
            except (ValueError, TypeError):
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


# ─── GUDE ───────────────────────────────────────
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


# ─── KCS ────────────────────────────────────────
KCS_ITEM_RE = re.compile(
    r"^(\d+)\s+(\S+)\s+(.+?)\s+(\S+)\s+(\d+)\s+(\d{1,2}-\d{1,2}-\d{2,4})\s+\$\s*([\d.,]+)\s+\$\s*([\d.,]+)\s*$",
    re.M
)


def parse_kcs(text: str) -> ParsedPO:
    p = ParsedPO(parser_used="KCS", raw_text=text, customer_name="KCS", currency="USD")

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


# ─── VORNE(用 extract_tables)─────────────────
def parse_vorne(pdf_bytes: bytes, text: str = "") -> ParsedPO:
    """VORNE 是表格型 PDF,用 pdfplumber 的 extract_tables() 抓。"""
    import pdfplumber

    p = ParsedPO(parser_used="VORNE", raw_text=text, customer_name="VORNE", currency="USD")

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page = pdf.pages[0]
            tables = page.extract_tables()
    except Exception as e:
        p.parse_warnings.append(f"VORNE: PDF 表格讀取失敗 ({e}),請手動補")
        return p

    current_item = None
    for tbl in tables:
        if not tbl:
            continue

        # Table 1: ['Date', '2/23/2026', 'P.O. #', '62887']
        if len(tbl) >= 1 and tbl[0]:
            header_row = [str(c) if c else "" for c in tbl[0]]
            if "P.O. #" in header_row:
                for i, cell in enumerate(header_row):
                    if cell == "P.O. #" and i + 1 < len(header_row):
                        p.customer_po_no = header_row[i + 1].strip()
                    if cell == "Date" and i + 1 < len(header_row):
                        p.po_date = _parse_iso_date(header_row[i + 1])

        # Table 2: 品項表(header 含 "Vorne P/N")
        if len(tbl) >= 2 and tbl[0] and any("Vorne P/N" in str(c) for c in tbl[0] if c):
            for row in tbl[1:]:
                if not row:
                    continue
                row_s = [(c or "").strip() for c in row]
                while len(row_s) < 10:
                    row_s.append("")

                item_id = row_s[0]
                qty_s = row_s[1]
                pn = row_s[2]
                desc = row_s[3]
                due = row_s[5]
                price = row_s[8]
                amt = row_s[9]

                # 主品項列(Item 是 A/B/C/D 字母,且有料號)
                if item_id and re.match(r"^[A-Z]$", item_id) and pn:
                    qty = int(qty_s) if qty_s.isdigit() else 0
                    current_item = POItem(
                        line=item_id,
                        part_number=pn,
                        description=desc,
                        quantity=qty,
                        delivery_date=_parse_iso_date(due) if due else "",
                        unit_price=_us_num(price),
                        amount=_us_num(amt),
                    )
                    p.items.append(current_item)
                # 描述補充列(上一個品項的延伸)
                elif current_item and not item_id and desc and not pn:
                    if not desc.startswith("90-") and "($" not in desc:
                        current_item.description += " " + desc
                # 最後一列 Total
                if amt and not item_id and not pn and not desc:
                    p.total_amount = _us_num(amt)

    if not p.items:
        p.parse_warnings.append("VORNE: 沒解出品項,請手動補")
    return p


def parse_generic(text: str) -> ParsedPO:
    p = ParsedPO(parser_used="GENERIC", raw_text=text)
    p.parse_warnings.append("無法辨識客戶格式,請手動輸入所有欄位")
    return p


# ─── 主入口 ───────────────────────────────────────
def parse_customer_po_from_pdf(pdf_bytes: bytes) -> ParsedPO:
    """從 PDF bytes 直接解析(推薦使用,VORNE 必須走這個)。"""
    text = extract_text_from_pdf(pdf_bytes)
    customer = detect_customer(text)

    if customer == "VORNE":
        result = parse_vorne(pdf_bytes, text)
    elif customer == "WESCO":
        result = parse_wesco(text)
    elif customer == "TIETO":
        result = parse_tieto(text)
    elif customer == "GUDE":
        result = parse_gude(text)
    elif customer == "KCS":
        result = parse_kcs(text)
    else:
        result = parse_generic(text)

    result.issuing_company_detected = detect_issuing_company(text)
    return result


def parse_customer_po(text: str) -> ParsedPO:
    """舊介面:只用文字解析(VORNE 會走 generic fallback)。"""
    customer = detect_customer(text)
    parsers = {
        "WESCO": parse_wesco,
        "TIETO": parse_tieto,
        "GUDE": parse_gude,
        "KCS": parse_kcs,
    }
    if customer == "VORNE":
        result = parse_generic(text)
        result.parse_warnings = [
            "VORNE 必須用 parse_customer_po_from_pdf(pdf_bytes) 才能解析"
        ]
    else:
        result = parsers.get(customer, parse_generic)(text)
    result.issuing_company_detected = detect_issuing_company(text)
    return result


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    import pdfplumber
    pages = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            pages.append(t)
    return "\n".join(pages)
