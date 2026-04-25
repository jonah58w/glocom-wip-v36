# -*- coding: utf-8 -*-
"""
PDF 產生器:接 PO 上下文 → 渲染 docx → 轉 PDF。

跟 v2 的差別:支援多品項(主表一個西拓編號可能對應多 P/N)。
"""

import subprocess
import sys
from datetime import date
from pathlib import Path
from docxtpl import DocxTemplate

HERE = Path(__file__).parent.parent
TEMPLATE_DIR = HERE / "templates"
OUTPUT_DIR = HERE / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


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
    """
    多品項處理:複製品項列(主表格 row 6)成 N 列。

    為什麼這樣做:docxtpl 的 {%tr for %} 對 cell merge + 多 cell 的 row
    有相容性問題,改用「程式端先複製 row」的策略,template 維持單品項版,
    但渲染前依品項數複製 row,然後 docxtpl 用 list index 填值。
    """
    if item_count <= 1:
        return

    from copy import deepcopy
    from docx import Document

    doc = Document(str(docx_path))
    # 主表格 = tables[1](tables[0] 是抬頭)
    main_table = doc.tables[1]

    # 找出品項列(第一個 cell 含 {{ item.part_number }})
    item_row_idx = None
    for ri, row in enumerate(main_table.rows):
        first_cell_text = row.cells[0].text
        if "item.part_number" in first_cell_text:
            item_row_idx = ri
            break

    if item_row_idx is None:
        return

    item_row = main_table.rows[item_row_idx]
    item_tr = item_row._tr  # lxml element

    # 變數對應:從 {{ item.xxx }} 改成 {{ items[N].xxx }}
    fields = ["part_number", "spec_text", "delivery_display",
              "quantity_display", "unit_price", "amount"]

    # 先把原本 row 改成 items[0] 索引
    for run in _iter_runs_in_tr(item_tr):
        for f in fields:
            old = "{{ item." + f + " }}"
            new = "{{ items[0]." + f + " }}"
            if old in run.text:
                run.text = run.text.replace(old, new)

    # 為每個額外品項複製一份 row,並改寫變數索引
    current_tr = item_tr
    for idx in range(1, item_count):
        new_tr = deepcopy(item_tr)
        # 改寫 new_tr 內所有 {{ items[0].xxx }} → {{ items[idx].xxx }}
        for run in _iter_runs_in_tr(new_tr):
            for f in fields:
                old = "{{ items[0]." + f + " }}"
                new = "{{ items[" + str(idx) + "]." + f + " }}"
                if old in run.text:
                    run.text = run.text.replace(old, new)
        # 插到當前 row 之後
        current_tr.addnext(new_tr)
        current_tr = new_tr

    doc.save(str(docx_path))


def _iter_runs_in_tr(tr_element):
    """遍歷 tr 元素內所有 <w:r> run,回傳可改 .text 的物件。

    用 python-docx 的方式:把 tr 當成虛擬 table 處理。
    """
    from docx.oxml.ns import qn

    # 直接走 XML,找 w:r 並用 w:t 操作
    # 為了能回傳「可改 text 的物件」,我們用 wrapper class
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
    """把 build_po_context() 的輸出渲染成 docx。

    多品項處理:
      1. 先用 docxtpl 把 template 複製到 output_path(此時還是單品項版)
      2. 在 docx 上手動把品項列複製 N 次,並改寫變數為 items[0]、items[1]...
      3. 用 docxtpl 第二次渲染,把實際資料填進去
    """

    issuing = po_ctx.get("issuing_company", "GLOCOM")
    template_name = (
        "PO_EUSWAY.docx" if issuing == "EUSWAY" else "PO_GLOCOM.docx"
    )
    template_path = TEMPLATE_DIR / template_name
    if not template_path.exists():
        raise FileNotFoundError(
            f"Template missing: {template_path}\n"
            f"請先執行 templates/build_po_glocom_template.py"
        )

    if output_path is None:
        po_no = po_ctx.get("po_no", "UNKNOWN")
        factory_short = po_ctx.get("factory_short", "")
        safe_name = f"{po_no}_{factory_short}".replace("/", "_").replace(" ", "_")
        output_path = OUTPUT_DIR / f"{safe_name}.docx"

    factory = po_ctx.get("factory") or {}
    items_data = po_ctx.get("items", [])

    # Step 1: 複製 template 到 output(實體檔)
    import shutil
    shutil.copy(str(template_path), str(output_path))

    # Step 2: 多品項時,在 docx 上預先複製 row + 改寫變數
    if len(items_data) > 1:
        _duplicate_item_row_for_multi(output_path, len(items_data))

    # Step 3: 組 docxtpl context(items 列表 + item 第一筆方便相容)
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

    context = {
        "order": {
            "no": po_ctx.get("po_no", ""),
            "date_display": _format_date_display(po_ctx.get("order_date")),
            "is_revised": po_ctx.get("is_revised", False),
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
    }

    # Step 4: docxtpl 渲染
    doc = DocxTemplate(str(output_path))
    doc.render(context)
    doc.save(str(output_path))
    return output_path


def docx_to_pdf(docx_path: Path) -> Path:
    """docx → PDF。Linux/Mac 用 LibreOffice;Windows 改用 docx2pdf。"""
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
