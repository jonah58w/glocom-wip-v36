import re
from PIL import Image
import pytesseract
from utils import parse_tags_from_text, safe_text


def parse_quick_text_line(line: str):
    parts = [x.strip() for x in line.split("|")]
    if not parts or not parts[0]:
        return None
    return {
        "po": parts[0],
        "wip": parts[1] if len(parts) > 1 else "",
        "ship_date": parts[2] if len(parts) > 2 else "",
        "tags": parse_tags_from_text(parts[3]) if len(parts) > 3 else [],
        "remark": parts[4] if len(parts) > 4 else "",
    }


def parse_email_text_to_rows(text: str):
    rows = []
    for raw in str(text).splitlines():
        line = " ".join(raw.split())
        if not line:
            continue
        po = extract_po_from_text(line)
        if not po:
            continue
        qty = ""
        m_qty = re.search(r"(\d[\d,]*)\s*(PCS|PCS\b)", line, flags=re.I)
        if m_qty:
            qty = m_qty.group(1).replace(",", "")
        ship_date = extract_date_from_text(line)
        wip = infer_wip_from_text(line) or "Production"
        part_no = ""
        m_part = re.search(rf"{re.escape(po)}\s+(.+?)\s*[-–—]+\s*\d[\d,]*\s*PCS", line, flags=re.I)
        if m_part:
            part_no = m_part.group(1).strip()
        remark = infer_remark_from_text(line)
        rows.append({
            "PO#": po,
            "Part No": part_no,
            "Qty": qty,
            "Factory Due Date": ship_date,
            "Ship Date": ship_date,
            "WIP": wip,
            "Remark": remark,
            "Customer Remark Tags": infer_customer_tags_from_text(line),
        })
    return rows


def ocr_image_to_text(image: Image.Image) -> str:
    try:
        return pytesseract.image_to_string(image, lang="eng")
    except Exception as e:
        return f"OCR_ERROR: {e}"


def extract_po_from_text(text: str) -> str:
    patterns = [
        r"\bPO[-\s]?\d+\b",
        r"\bPO\d+\b",
        r"\bEW[-\s]?\d+\b",
        r"\bEW\d+\b",
        r"\b[A-Z]{1,4}\d{5,}(?:-\d+)?[A-Z]?\b",
    ]
    for p in patterns:
        m = re.search(p, text, flags=re.IGNORECASE)
        if m:
            return m.group(0).replace(" ", "")
    return ""


def extract_date_from_text(text: str) -> str:
    patterns = [
        r"\b20\d{2}[-/]\d{1,2}[-/]\d{1,2}\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b",
        r"\b\d{1,2}/\d{1,2}\b",
        r"\b\d{4}=>\d{4}\b",
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return m.group(0)
    return ""


def infer_wip_from_text(text: str) -> str:
    t = text.lower()
    if any(k in t for k in ["complete", "completed", "finish", "finished"]) or "完成" in text:
        return "完成"
    if any(k in t for k in ["shipping", "ship out", "shipped", "待出貨"]):
        return "Shipping"
    if any(k in t for k in ["packing", "packed", "pack"]) or "包裝" in text:
        return "Packing"
    if any(k in t for k in ["hold", "on hold"]):
        return "On Hold"
    if any(k in t for k in ["remake", "rework"]):
        return "Remake in Process"
    if any(k in t for k in ["gerber", "eq", "engineering question"]):
        return "Engineering"
    if any(k in t for k in ["fqc", "qa", "inspection"]) or any(k in text for k in ["測試", "成檢"]):
        return "Inspection"
    if any(k in t for k in ["aoi", "drill", "drilling", "plating", "routing", "route", "inner layer", "inner"]):
        return "Production"
    return ""


def infer_customer_tags_from_text(text: str):
    t = text.lower()
    tags = []
    if "working gerber" in t or "gerber for approval" in t:
        tags.append("Working Gerber for Approval")
    if "eq" in t or "engineering question" in t:
        tags.append("Engineering Question")
    if "payment" in t:
        tags.append("Payment Pending")
    if "remake" in t or "rework" in t:
        tags.append("Remake in Process")
    if "hold" in t:
        tags.append("On Hold")
    if "partial shipment" in t:
        tags.append("Partial Shipment")
    if "shipped" in t or "ship out" in t:
        tags.append("Shipped")
    if "waiting confirmation" in t or "await confirmation" in t:
        tags.append("Waiting Confirmation")
    return list(dict.fromkeys(tags))


def infer_remark_from_text(text: str) -> str:
    cleaned = " ".join(str(text).split())
    return cleaned[:300] if cleaned else ""
