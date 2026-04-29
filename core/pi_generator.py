# -*- coding: utf-8 -*-
"""
Proforma Invoice 產生器 v2.

功能:接 PI context → 渲染 PI_GLOCOM.docx → 轉 PDF。

新增:
- pcb_spec_to_english():中文工廠 PO 規格 → 英文 PI description
"""

import re
import shutil
import subprocess
import sys
from copy import deepcopy
from datetime import date
from pathlib import Path

from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm

HERE = Path(__file__).parent.parent
TEMPLATE_DIR = HERE / "templates"
OUTPUT_DIR = HERE / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# logo 檔位置(優先順序):
#   1. templates/glocom_logo.png  (PNG 檔)
#   2. templates/glocom_logo.jpg  (JPG)
#   3. 從 PO_GLOCOM.docx 解出的 image1.png
LOGO_CANDIDATES = [
    TEMPLATE_DIR / "glocom_logo.png",
    TEMPLATE_DIR / "glocom_logo.jpg",
]


def _find_logo_path():
    for p in LOGO_CANDIDATES:
        if p.exists():
            return p
    # fallback: 從 PO_GLOCOM.docx 拿
    po_template = TEMPLATE_DIR / "PO_GLOCOM.docx"
    if po_template.exists():
        try:
            import zipfile
            extracted = TEMPLATE_DIR / "_glocom_logo_extracted.png"
            if not extracted.exists():
                with zipfile.ZipFile(po_template) as z:
                    for name in z.namelist():
                        if name.startswith("word/media/image1."):
                            extracted = TEMPLATE_DIR / f"_glocom_logo_extracted{Path(name).suffix}"
                            with z.open(name) as src, open(extracted, "wb") as dst:
                                dst.write(src.read())
                            break
            if extracted.exists():
                return extracted
        except Exception:
            pass
    return None


# ============================================================
# 中文 PCB 規格 → 英文 PI description
# ============================================================
def pcb_spec_to_english(spec_text: str, brand_ul: bool = True) -> str:
    """
    把工廠 PO 中文規格轉成 PI 給客戶看的英文描述。
    
    輸入範例:
      Working Gerber承認後,才可生產!請於下午14:00前傳working gerber
      Material: 4L; Tg170; Board thickness: 1.6mm; Copper: 1oz/1oz;
      Surface Finish: ENIG 2u"; S/M: Green; S/L: White
      須添加西拓UL logo & date code (YYWW);
      樣板: 除試錫板外,須另外提供樣板供備份.
    
    輸出範例:
      Bare Printed Circuit Board; Brand UL
      4L, FR4(Tg170), 1.60mm, 4up, Cu: 1oz finish all layers,
      S/M: Green; S/L: White, Surface Finish: ENIG 2u
    """
    if not spec_text:
        return ""
    
    text = str(spec_text)
    
    # 抽出主要規格元素
    parts = {}
    
    # Material (層數)
    m = re.search(r'Material:\s*(\d+L)', text, re.IGNORECASE)
    if m:
        parts['layers'] = m.group(1)
    
    # Tg
    m = re.search(r'\bTg\s*(\d+)', text, re.IGNORECASE)
    if m:
        parts['tg'] = f"FR4(Tg{m.group(1)})"
    else:
        parts['tg'] = "FR4"
    
    # 厚度
    m = re.search(r'(?:Board\s+)?thickness:\s*([\d.]+)mm', text, re.IGNORECASE)
    if m:
        thick = float(m.group(1))
        parts['thickness'] = f"{thick:.2f}mm"
    
    # Copper
    m = re.search(r'Copper:\s*([\w/]+(?:oz)?)', text, re.IGNORECASE)
    if m:
        copper_raw = m.group(1).strip()
        if 'oz' in copper_raw.lower():
            parts['copper'] = f"Cu: {copper_raw} finish all layers"
        else:
            parts['copper'] = f"Cu: {copper_raw}oz finish all layers"
    
    # Up panel
    m = re.search(r'(\d+)\s*up\s*panel', text, re.IGNORECASE)
    if m:
        parts['panel'] = f"{m.group(1)}up"
    else:
        m2 = re.search(r'(\d+)\s*up', text, re.IGNORECASE)
        if m2:
            parts['panel'] = f"{m2.group(1)}up"
    
    # Surface Finish
    m = re.search(r'Surface\s+Finish:\s*([^;,\n]+)', text, re.IGNORECASE)
    if m:
        parts['finish'] = m.group(1).strip().rstrip('.')
    
    # S/M
    m = re.search(r'S/M:\s*([^;,\n]+)', text, re.IGNORECASE)
    if m:
        parts['sm'] = m.group(1).strip()
    
    # S/L
    m = re.search(r'S/L:\s*([^;,\n]+)', text, re.IGNORECASE)
    if m:
        parts['sl'] = m.group(1).strip()
    
    # 組第 1 行
    line1 = "Bare Printed Circuit Board"
    if brand_ul:
        line1 += "; Brand UL"
    
    # 組第 2 行 (主規格): layers, FR4(Tg), thickness, panel, copper
    line2_parts = []
    if 'layers' in parts:
        line2_parts.append(parts['layers'])
    if 'tg' in parts:
        line2_parts.append(parts['tg'])
    if 'thickness' in parts:
        line2_parts.append(parts['thickness'])
    if 'panel' in parts:
        line2_parts.append(parts['panel'])
    if 'copper' in parts:
        line2_parts.append(parts['copper'])
    line2 = ", ".join(line2_parts)
    
    # 組第 3 行 (S/M, S/L, Surface Finish)
    line3_parts = []
    if 'sm' in parts:
        line3_parts.append(f"S/M: {parts['sm']}")
    if 'sl' in parts:
        line3_parts.append(f"S/L: {parts['sl']}")
    if 'finish' in parts:
        line3_parts.append(f"Surface Finish: {parts['finish']}")
    line3 = "; ".join(line3_parts) if line3_parts else ""
    
    # 組合
    lines = [line1, line2]
    if line3:
        lines.append(line3)
    
    return "\n".join(lines)


# ============================================================
# 工具函式
# ============================================================
def _format_date_us_clean(d) -> str:
    if not d:
        return ""
    if isinstance(d, str):
        return d
    return d.strftime("%b. %d, %Y")


def _money(amount: float, symbol: str = "$") -> str:
    try:
        return f"{symbol}{float(amount):,.2f}"
    except (ValueError, TypeError):
        return f"{symbol}0.00"


def _qty_display(qty, unit: str) -> str:
    try:
        n = int(qty)
        return f"{n:,}{unit}"
    except (ValueError, TypeError):
        return f"{qty}{unit}"


def _escape_for_docx(text: str) -> str:
    """XML escape & < > 給 |safe filter 用"""
    if not text:
        return ""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _iter_runs_in_tr(tr_element):
    class RunWrapper:
        def __init__(self, t_element):
            self._t = t_element
        @property
        def text(self):
            return self._t.text or ""
        @text.setter
        def text(self, value):
            self._t.text = value
    
    for t in tr_element.iter(qn("w:t")):
        yield RunWrapper(t)


# ============================================================
# 動態複製 row
# ============================================================
def _duplicate_item_row_for_multi(docx_path: Path, item_count: int):
    """
    跟 PO 模板一致 — 把 {{ it.xxx }} 改成 {{ items[0].xxx }},
    複製 N-1 個 row,改成 items[1..N-1].xxx。
    """
    if item_count <= 1:
        return

    doc = Document(str(docx_path))
    
    # 找含 {{ it. 的 row
    items_table = None
    item_row_idx = None
    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            for cell in row.cells:
                if "{{ it." in cell.text or "{{it." in cell.text:
                    items_table = table
                    item_row_idx = ri
                    break
            if items_table:
                break
        if items_table:
            break
    
    if items_table is None or item_row_idx is None:
        return
    
    item_row = items_table.rows[item_row_idx]
    item_tr = item_row._tr
    
    fields = ["item_no", "description", "quantity_display", "unit_price_display", "amount_display"]
    
    def _replace_tags(text, target_idx):
        new = text
        for f in fields:
            for suffix in ("|safe", ""):
                old = "{{ it." + f + suffix + " }}"
                rep = "{{ items[" + str(target_idx) + "]." + f + suffix + " }}"
                if old in new:
                    new = new.replace(old, rep)
        return new
    
    # 第一個 row: it → items[0]
    for run in _iter_runs_in_tr(item_tr):
        run.text = _replace_tags(run.text, 0)
    
    # 複製 N-1 個 row: items[i]
    current_tr = item_tr
    for idx in range(1, item_count):
        new_tr = deepcopy(item_tr)
        for run in _iter_runs_in_tr(new_tr):
            t = run.text
            for f in fields:
                for suffix in ("|safe", ""):
                    old = "{{ items[0]." + f + suffix + " }}"
                    rep = "{{ items[" + str(idx) + "]." + f + suffix + " }}"
                    if old in t:
                        t = t.replace(old, rep)
            run.text = t
        current_tr.addnext(new_tr)
        current_tr = new_tr
    
    doc.save(str(docx_path))


# ============================================================
# 渲染 PI docx
# ============================================================
def render_pi_docx(pi_ctx: dict, output_path: Path = None) -> Path:
    template_path = TEMPLATE_DIR / "PI_GLOCOM.docx"
    if not template_path.exists():
        raise FileNotFoundError(f"PI 模板不存在: {template_path}")

    if output_path is None:
        invoice_no = pi_ctx.get("invoice_no", "UNKNOWN")
        customer_short = pi_ctx.get("customer_short", "")
        safe_name = f"PI_{invoice_no}_{customer_short}".replace("/", "_").replace(" ", "_")
        output_path = OUTPUT_DIR / f"{safe_name}.docx"

    items_data = pi_ctx.get("items", [])
    currency_symbol = pi_ctx.get("currency_symbol", "$")
    
    # 處理 items
    items_ctx = []
    total_amount = 0.0
    for it in items_data:
        unit_price = float(it.get("unit_price", 0))
        amount = float(it.get("amount", 0))
        total_amount += amount
        items_ctx.append({
            "item_no": _escape_for_docx(it.get("item_no", "")),
            "description": _escape_for_docx(it.get("description", "")),
            "quantity_display": _qty_display(it.get("quantity", 0), it.get("quantity_unit", "pcs")),
            "unit_price_display": _money(unit_price, currency_symbol),
            "amount_display": _money(amount, currency_symbol),
        })
    
    # Bank fee
    bank_fee = float(pi_ctx.get("bank_fee", 0))
    if bank_fee > 0:
        items_ctx.append({
            "item_no": "",
            "description": "Bank fee",
            "quantity_display": "1set",
            "unit_price_display": _money(bank_fee, currency_symbol),
            "amount_display": _money(bank_fee, currency_symbol),
        })
        total_amount += bank_fee

    # 複製模板 → 取得 N 列 → render
    shutil.copy(str(template_path), str(output_path))
    if len(items_ctx) > 1:
        _duplicate_item_row_for_multi(output_path, len(items_ctx))
    
    doc = DocxTemplate(str(output_path))

    pi_date = pi_ctx.get("date") or date.today()
    date_str = _format_date_us_clean(pi_date) if not isinstance(pi_date, str) else pi_date

    # logo 動態載入(若沒檔則用空白佔位)
    logo_path = _find_logo_path()
    if logo_path:
        logo_img = InlineImage(doc, str(logo_path), width=Cm(3.6))
    else:
        logo_img = ""

    context = {
        "logo_img": logo_img,
        "invoice_no": pi_ctx.get("invoice_no", ""),
        "po_no": pi_ctx.get("po_no", ""),
        "customer_po_no": pi_ctx.get("customer_po_no", pi_ctx.get("po_no", "")),
        "customer_code": pi_ctx.get("customer_code", ""),
        "date_str": date_str,
        "contact_person": pi_ctx.get("contact_person", ""),
        "customer_name": pi_ctx.get("customer_name", ""),
        "customer_short": pi_ctx.get("customer_short", ""),
        "customer_address1": pi_ctx.get("customer_address1", ""),
        "customer_address2": pi_ctx.get("customer_address2", ""),
        "customer_tel": pi_ctx.get("customer_tel", ""),
        "customer_fax": pi_ctx.get("customer_fax", ""),
        "shipment_text": pi_ctx.get("shipment_text", ""),
        "from_country": pi_ctx.get("from_country", "Taiwan"),
        "to_country": pi_ctx.get("to_country", ""),
        "terms_text": pi_ctx.get("terms_text", ""),
        "item": items_ctx[0] if items_ctx else {},
        "items": items_ctx,
        "total_display": _money(total_amount, currency_symbol),
    }

    doc.render(context)
    doc.save(str(output_path))
    return output_path


def docx_to_pdf(docx_path: Path) -> Path:
    pdf_path = docx_path.with_suffix(".pdf")
    if sys.platform == "win32":
        try:
            from docx2pdf import convert
            convert(str(docx_path), str(pdf_path))
            return pdf_path
        except ImportError:
            raise RuntimeError("Windows 環境請: pip install docx2pdf")
    else:
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf",
             "--outdir", str(docx_path.parent), str(docx_path)],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice 轉換失敗: {result.stderr}")
        return pdf_path


def generate_pi_files(pi_ctx: dict) -> dict:
    docx_path = render_pi_docx(pi_ctx)
    result = {"docx_path": docx_path, "pdf_path": None, "error": None}
    try:
        pdf_path = docx_to_pdf(docx_path)
        result["pdf_path"] = pdf_path
    except Exception as e:
        result["error"] = f"PDF 轉換失敗(docx 已產生): {e}"
    return result
