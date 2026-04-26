# -*- coding: utf-8 -*-
"""
PDF 產生器 v2:接 PO 上下文 → 渲染 docx → 轉 PDF。

v2 更新:
- 支援 REVISED 印章:當 po_ctx['is_revised']=True 時,
  在 Logo 下方插入紅色斜印「REVISED」(模板用 jinja {% if order.is_revised %}{{ revised_stamp }}{% endif %})
- 印章圖檔位置:templates/revised_stamp.png

跟 v1 的差別 (除上述):多品項處理邏輯不變。
"""

import subprocess
import sys
from datetime import date
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm

HERE = Path(__file__).parent.parent
TEMPLATE_DIR = HERE / "templates"
OUTPUT_DIR = HERE / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

REVISED_STAMP_PATH = TEMPLATE_DIR / "revised_stamp.png"


def _format_qty_display(qty: int, panel_qty: int | None) -> str:
    if panel_qty:
        return f"{qty}pcs ({panel_qty}panel)"
    return f"{qty}pcs"


def _format_delivery_display(d: date | None, note: str) -> str:
    if not d:
        return note or ""
    base = d.strftime("%b %d, %Y").upper()
    if note:
        return f"{base}\n{note}"
    return base


def _format_date_display(d: date | None) -> str:
    if not d:
        return ""
    return d.strftime("%b. %d, %Y").upper().replace("..", ".")


def _duplicate_item_row_for_multi(docx_path: Path, item_count: int):
    """多品項處理:複製品項列(主表格 row 6)成 N 列。"""
    if item_count <= 1:
        return

    from copy import deepcopy
    from docx import Document

    doc = Document(str(docx_path))
    main_table = doc.tables[1]

    item_row_idx = None
    for ri, row in enumerate(main_table.rows):
        first_cell_text = row.cells[0].text
        if "item.part_number" in first_cell_text:
            item_row_idx = ri
            break

    if item_row_idx is None:
        return

    item_row = main_table.rows[item_row_idx]
    item_tr = item_row._tr

    fields = ["part_number", "spec_text", "delivery_display",
              "quantity_display", "unit_price", "amount"]

    for run in _iter_runs_in_tr(item_tr):
        for f in fields:
            old = "{{ item." + f + " }}"
            new = "{{ items[0]." + f + " }}"
            if old in run.text:
                run.text = run.text.replace(old, new)

    current_tr = item_tr
    for idx in range(1, item_count):
        new_tr = deepcopy(item_tr)
        for run in _iter_runs_in_tr(new_tr):
            for f in fields:
                old = "{{ items[0]." + f + " }}"
                new = "{{ items[" + str(idx) + "]." + f + " }}"
                if old in run.text:
                    run.text = run.text.replace(old, new)
        current_tr.addnext(new_tr)
        current_tr = new_tr

    doc.save(str(docx_path))


def _iter_runs_in_tr(tr_element):
    from docx.oxml.ns import qn

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


def render_docx_from_po_ctx(po_ctx: dict, output_path: Path | None = None) -> Path:
    """渲染 docx。is_revised=True 會在 Logo 下方蓋紅色印章。"""

    issuing = po_ctx.get("issuing_company", "GLOCOM")
    template_name = (
        "PO_EUSWAY.docx" if issuing == "EUSWAY" else "PO_GLOCOM.docx"
    )
    template_path = TEMPLATE_DIR / template_name
    if not template_path.exists():
        raise FileNotFoundError(
            f"Template missing: {template_path}\n"
            f"請確認 templates/{template_name} 已上傳到 GitHub。"
        )

    if output_path is None:
        po_no = po_ctx.get("po_no", "UNKNOWN")
        factory_short = po_ctx.get("factory_short", "")
        safe_name = f"{po_no}_{factory_short}".replace("/", "_").replace(" ", "_")
        output_path = OUTPUT_DIR / f"{safe_name}.docx"

    factory = po_ctx.get("factory") or {}
    items_data = po_ctx.get("items", [])

    import shutil
    shutil.copy(str(template_path), str(output_path))

    if len(items_data) > 1:
        _duplicate_item_row_for_multi(output_path, len(items_data))

    items_ctx = [
        {
            "part_number": it["part_number"],
            "spec_text": it["spec_text"],
            "quantity_display": _format_qty_display(it["quantity"], it.get("panel_qty")),
            "delivery_display": _format_delivery_display(
                it.get("delivery_date"), it.get("delivery_note", "")
            ),
            "unit_price": f"{it['unit_price']:,.3f}",
            "amount": f"{it['amount']:,.2f}",
        }
        for it in items_data
    ]

    is_revised = bool(po_ctx.get("is_revised", False))

    # docxtpl 必須先 instantiate doc 才能建 InlineImage
    doc = DocxTemplate(str(output_path))

    # REVISED 印章:有勾才放 InlineImage,沒勾就放空字串(讓 jinja {% if %} 不顯示)
    if is_revised and REVISED_STAMP_PATH.exists():
        revised_stamp = InlineImage(doc, str(REVISED_STAMP_PATH), width=Cm(2.2))
    else:
        revised_stamp = ""

    context = {
        "order": {
            "no": po_ctx.get("po_no", ""),
            "date_display": _format_date_display(po_ctx.get("order_date")),
            "is_revised": is_revised,
            "purchase_responsible": po_ctx.get("purchase_responsible", "Amy"),
            "currency": po_ctx.get("currency", "NT$"),
            "payment_terms": po_ctx.get("payment_terms", ""),
            "shipment_method": po_ctx.get("shipment_method", "待通知"),
            "ship_to": po_ctx.get("ship_to") or po_ctx.get("ship_to_default", "待通知"),
            "customer_po_no": po_ctx.get("customer_po_no", ""),
            "total_amount": f"{po_ctx.get('total_amount', 0):,.2f}",
        },
        "factory": {
            "name": factory.get("factory_name", po_ctx.get("factory_short", "")),
            "address": factory.get("address", ""),
            "contact": factory.get("contact_person", ""),
            "phone": factory.get("phone", ""),
            "fax": factory.get("fax", ""),
        },
        "item": items_ctx[0] if items_ctx else {},
        "items": items_ctx,
        "revised_stamp": revised_stamp,
    }

    doc.render(context)
    doc.save(str(output_path))
    return output_path


def docx_to_pdf(docx_path: Path) -> Path:
    """docx → PDF。Linux/Mac 用 LibreOffice;Windows 用 docx2pdf。"""
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


def generate_po_files(po_ctx: dict) -> dict:
    """完整流程: PO context → docx + PDF."""
    docx_path = render_docx_from_po_ctx(po_ctx)
    result = {"docx_path": docx_path, "pdf_path": None, "error": None}
    try:
        pdf_path = docx_to_pdf(docx_path)
        result["pdf_path"] = pdf_path
    except Exception as e:
        result["error"] = f"PDF 轉換失敗(docx 已產生): {e}"
    return result
