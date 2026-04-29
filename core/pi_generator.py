# -*- coding: utf-8 -*-
"""
Proforma Invoice 產生器 v2。

主要更新(v2):
- ★ 模板重做:含 GLOCOM Logo + 緊湊 3 區塊 + 全部一頁
- ★ 多品項用動態 row 複製(類似 pdf_generator)
- ★ description 用 |safe filter,& 自動 escape
"""

import subprocess
import sys
from copy import deepcopy
from datetime import date
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docxtpl import DocxTemplate

HERE = Path(__file__).parent.parent
TEMPLATE_DIR = HERE / "templates"
OUTPUT_DIR = HERE / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


def escape_for_docx(text: str) -> str:
    if not text:
        return ""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _format_date(d) -> str:
    if not d:
        return ""
    if hasattr(d, 'strftime'):
        return d.strftime("%b. %d, %Y")
    return str(d)


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


def _duplicate_item_row_for_multi(docx_path: Path, item_count: int):
    if item_count <= 1:
        return

    doc = Document(str(docx_path))
    item_table = None
    item_row_idx = None
    for ti, tbl in enumerate(doc.tables):
        for ri, row in enumerate(tbl.rows):
            row_text = "".join(cell.text for cell in row.cells)
            if "item.item_no" in row_text:
                item_table = tbl
                item_row_idx = ri
                break
        if item_table is not None:
            break

    if item_table is None or item_row_idx is None:
        return

    item_row = item_table.rows[item_row_idx]
    item_tr = item_row._tr

    fields = ["item_no", "description", "quantity_display",
              "unit_price_display", "amount_display"]

    # 第 0 個改成 items[0].xxx
    for run in _iter_runs_in_tr(item_tr):
        for f in fields:
            for tmpl in [
                ("{{ item." + f + " }}", "{{ items[0]." + f + " }}"),
                ("{{ item." + f + "|safe }}", "{{ items[0]." + f + "|safe }}"),
            ]:
                if tmpl[0] in run.text:
                    run.text = run.text.replace(tmpl[0], tmpl[1])

    # 複製 row 給其他 idx
    current_tr = item_tr
    for idx in range(1, item_count):
        new_tr = deepcopy(item_tr)
        for run in _iter_runs_in_tr(new_tr):
            for f in fields:
                for tmpl in [
                    ("{{ items[0]." + f + " }}", "{{ items[" + str(idx) + "]." + f + " }}"),
                    ("{{ items[0]." + f + "|safe }}", "{{ items[" + str(idx) + "]." + f + "|safe }}"),
                ]:
                    if tmpl[0] in run.text:
                        run.text = run.text.replace(tmpl[0], tmpl[1])
        current_tr.addnext(new_tr)
        current_tr = new_tr

    doc.save(str(docx_path))


def render_pi_docx(pi_ctx: dict, output_path=None) -> Path:
    template_path = TEMPLATE_DIR / "PI_GLOCOM.docx"
    if not template_path.exists():
        raise FileNotFoundError(
            f"Template missing: {template_path}\n"
            f"請確認 templates/PI_GLOCOM.docx 已上傳到 GitHub。"
        )

    if output_path is None:
        invoice_no = pi_ctx.get("invoice_no", "PI_UNKNOWN")
        cust_short = pi_ctx.get("customer_name", "")[:5].replace(" ", "_")
        safe_name = f"PI_{invoice_no}_{cust_short}".replace("/", "_")
        output_path = OUTPUT_DIR / f"{safe_name}.docx"

    items_data = pi_ctx.get("items", [])

    import shutil
    shutil.copy(str(template_path), str(output_path))

    if len(items_data) > 1:
        _duplicate_item_row_for_multi(output_path, len(items_data))

    cur = pi_ctx.get("currency_symbol", "$")

    items_ctx = []
    for it in items_data:
        unit = it.get("quantity_unit", "pcs")
        up = float(it["unit_price"])
        # Setup/Tooling 用整數,PCB 用兩位小數
        if it.get("quantity_unit") == "set" or up == int(up):
            up_disp = f"{cur}{up:,.0f}"
        else:
            up_disp = f"{cur}{up:,.2f}"
        if unit == "pcs":
            qty_disp = f"{it['quantity']}{unit}"
        else:
            qty_disp = f"{it['quantity']} {unit}"
        items_ctx.append({
            "item_no": it.get("item_no", ""),
            "description": escape_for_docx(it.get("description", "")),
            "quantity_display": qty_disp,
            "unit_price_display": up_disp,
            "amount_display": f"{cur}{it['amount']:,.2f}",
        })

    # Total = items 總和 + bank_fee
    bank_fee = float(pi_ctx.get("bank_fee", 45))
    total = sum(float(it["amount"]) for it in items_data) + bank_fee
    bank_fee_disp = f"{cur}{bank_fee:,.0f}" if bank_fee == int(bank_fee) else f"{cur}{bank_fee:,.2f}"

    doc = DocxTemplate(str(output_path))
    context = {
        "invoice_no": pi_ctx.get("invoice_no", ""),
        "po_no": pi_ctx.get("po_no", ""),
        "customer_po_no": pi_ctx.get("customer_po_no", ""),
        "customer_code": pi_ctx.get("customer_code", ""),
        "date_display": _format_date(pi_ctx.get("date")) or pi_ctx.get("date_display", ""),
        "contact_person": pi_ctx.get("contact_person", ""),
        "customer_name": pi_ctx.get("customer_name", ""),
        "customer_address1": pi_ctx.get("customer_address1", ""),
        "customer_address2": pi_ctx.get("customer_address2", ""),
        "customer_tel": pi_ctx.get("customer_tel", ""),
        "customer_fax": pi_ctx.get("customer_fax", ""),
        "shipment_text": pi_ctx.get("shipment_text", ""),
        "from_country": pi_ctx.get("from_country", "Taiwan"),
        "to_country": pi_ctx.get("to_country", "USA"),
        "terms_text": pi_ctx.get("terms_text", "Exwork Taiwan (US$ Exwork Taiwan)"),
        "items": items_ctx,
        "bank_fee_display": bank_fee_disp,
        "total_display": f"{cur}{total:,.2f}",
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
