# -*- coding: utf-8 -*-
"""
Proforma Invoice 產生器 v1。

主要功能:
- 渲染 PI docx → 轉 PDF
- 模板 templates/PI_GLOCOM.docx 用 docxtpl loop 處理多品項
- description 用 |safe filter,& 自動 escape

API:
    generate_pi_files(pi_ctx) → {"docx_path": ..., "pdf_path": ..., "error": ...}

pi_ctx 結構:
{
    "invoice_no": "G1150014",                # PI 主號(去掉 -01)
    "po_no": "G1150014-01",                  # 西拓訂單號(完整)
    "customer_po_no": "62887",               # 客戶 PO#
    "customer_code": "GC290",                # Cust# (從客戶檔抓)
    "date": date(2026, 4, 29),               # PI 開立日期
    "contact_person": "Ana Castaneda",       # Messrs.
    "customer_name": "Vorne Industries Incorporated",
    "customer_address1": "1445 Industrial Drive Itasca,",
    "customer_address2": "IL 60143-1849",
    "customer_tel": "(630) 875-3600",
    "customer_fax": "(630) 875-3609",
    "shipment_text": "Since the T/T payment is received and WG approved, we will confirm the ship date then",
    "from_country": "Taiwan",
    "to_country": "USA",
    "terms_text": "Exwork Taiwan (US$ Exwork Taiwan)",
    "items": [
        {
            "item_no": "90-0341-00",
            "description": "Bare Printed Circuit Board; Brand UL\\n4L, FR4(Tg170), ...",
            "quantity": 500,
            "quantity_unit": "pcs",       # 'pcs' / 'set'
            "unit_price": 2.95,
            "amount": 1475.0,
        },
        ...
    ],
    "bank_fee": 45.0,                        # 預設 45,Sandy 可改
    "currency_symbol": "$",                  # 預設 USD
}
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


def escape_for_docx(text: str) -> str:
    """XML escape 給 docxtpl |safe filter 用"""
    if not text:
        return ""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _format_date_pi(d: date | None) -> str:
    """PI 日期格式: Feb. 26, 2026"""
    if not d:
        return ""
    return d.strftime("%b. %d, %Y").replace("..", ".")


def _format_qty_display(qty, unit="pcs"):
    """數量顯示: 500pcs / 1 set / 1,200pcs"""
    try:
        qty_num = int(qty)
        if unit == "set":
            return f"{qty_num} set"
        return f"{qty_num:,}{unit}"
    except Exception:
        return f"{qty}{unit}"


def _format_unit_price(price, currency="$"):
    """單價顯示: $2.95 / $400"""
    try:
        p = float(price)
        # 整數就不顯示小數
        if p == int(p):
            return f"{currency}{int(p)}"
        return f"{currency}{p:.2f}"
    except Exception:
        return f"{currency}{price}"


def _format_amount(amount, currency="$"):
    """金額顯示: $1,475.00"""
    try:
        a = float(amount)
        return f"{currency}{a:,.2f}"
    except Exception:
        return f"{currency}{amount}"


def render_pi_docx(pi_ctx: dict, output_path: Path | None = None) -> Path:
    """渲染 PI docx 模板。"""
    template_path = TEMPLATE_DIR / "PI_GLOCOM.docx"
    if not template_path.exists():
        raise FileNotFoundError(
            f"PI 模板不存在: {template_path}\n"
            f"請確認 templates/PI_GLOCOM.docx 已上傳到 GitHub。"
        )
    
    if output_path is None:
        invoice_no = pi_ctx.get("invoice_no", "UNKNOWN")
        customer_short = pi_ctx.get("customer_short", "")
        safe_name = f"PI_{invoice_no}_{customer_short}".replace("/", "_").replace(" ", "_")
        output_path = OUTPUT_DIR / f"{safe_name}.docx"
    
    items_data = pi_ctx.get("items", []) or []
    bank_fee = float(pi_ctx.get("bank_fee", 0) or 0)
    currency = pi_ctx.get("currency_symbol", "$")
    
    # 組品項列(line_items)— 按範本順序:先 PCB 主品項,再 Setup,最後 Bank fee
    line_items = []
    
    # PCB 主品項
    for it in items_data:
        if it.get("is_setup", False):
            continue  # Setup 後面處理
        desc = it.get("description", "")
        line_items.append({
            "item_no": it.get("item_no", ""),
            "description": escape_for_docx(desc),
            "quantity_display": _format_qty_display(
                it.get("quantity", 0),
                it.get("quantity_unit", "pcs"),
            ),
            "unit_price_display": _format_unit_price(it.get("unit_price", 0), currency),
            "amount_display": _format_amount(it.get("amount", 0), currency),
        })
    
    # Setup 品項
    for it in items_data:
        if not it.get("is_setup", False):
            continue
        line_items.append({
            "item_no": it.get("item_no", ""),
            "description": escape_for_docx(it.get("description", "Setup & Tooling Charge")),
            "quantity_display": _format_qty_display(
                it.get("quantity", 1),
                it.get("quantity_unit", "set"),
            ),
            "unit_price_display": _format_unit_price(it.get("unit_price", 0), currency),
            "amount_display": _format_amount(it.get("amount", 0), currency),
        })
    
    # Bank fee 一行(如果 > 0)
    bank_fee_amount = 0.0
    if bank_fee > 0:
        bank_fee_amount = bank_fee
        line_items.append({
            "item_no": "",
            "description": "Bank fee",
            "quantity_display": "1 set",
            "unit_price_display": _format_unit_price(bank_fee, currency),
            "amount_display": _format_amount(bank_fee, currency),
        })
    
    # 算 Total
    total_amount = sum(float(it.get("amount", 0) or 0) for it in items_data) + bank_fee_amount
    
    # 組 context
    context = {
        "pi": {
            "contact_person": pi_ctx.get("contact_person", ""),
            "customer_name": pi_ctx.get("customer_name", ""),
            "customer_address1": pi_ctx.get("customer_address1", ""),
            "customer_address2": pi_ctx.get("customer_address2", ""),
            "customer_tel": pi_ctx.get("customer_tel", ""),
            "customer_fax": pi_ctx.get("customer_fax", ""),
            "customer_code": pi_ctx.get("customer_code", ""),
            "date_display": _format_date_pi(pi_ctx.get("date")),
            "invoice_no": pi_ctx.get("invoice_no", ""),
            "customer_po_no": pi_ctx.get("customer_po_no", ""),
            "shipment_text": pi_ctx.get("shipment_text", ""),
            "from_country": pi_ctx.get("from_country", "Taiwan"),
            "to_country": pi_ctx.get("to_country", "USA"),
            "terms_text": pi_ctx.get("terms_text", ""),
            "line_items": line_items,
            "total_display": _format_amount(total_amount, currency),
        }
    }
    
    # 渲染
    import shutil
    shutil.copy(str(template_path), str(output_path))
    doc = DocxTemplate(str(output_path))
    doc.render(context)
    doc.save(str(output_path))
    return output_path


def docx_to_pdf(docx_path: Path) -> Path:
    """docx → PDF"""
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
    """完整流程: PI context → docx + PDF"""
    docx_path = render_pi_docx(pi_ctx)
    result = {"docx_path": docx_path, "pdf_path": None, "error": None}
    try:
        pdf_path = docx_to_pdf(docx_path)
        result["pdf_path"] = pdf_path
    except Exception as e:
        result["error"] = f"PDF 轉換失敗(docx 已產生): {e}"
    return result
