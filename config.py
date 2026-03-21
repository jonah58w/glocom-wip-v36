
# -*- coding: utf-8 -*-
import streamlit as st
import pytesseract

st.set_page_config(
    page_title="GLOCOM Control Tower",
    page_icon="🏭",
    layout="wide",
)

DEFAULT_TABLE_URL = "https://app.teable.ai/api/table/tbl6c05EPXYtJcZfeir/record"
TEABLE_WEB_URL = "https://app.teable.ai/base/bsedgLzbHjiK0XoZH01/table/tbl6c05EPXYtJcZfeir"

try:
    TEABLE_TOKEN = st.secrets.get("TEABLE_TOKEN", "")
except Exception:
    TEABLE_TOKEN = ""

try:
    TABLE_URL = st.secrets.get("TEABLE_TABLE_URL", DEFAULT_TABLE_URL)
except Exception:
    TABLE_URL = DEFAULT_TABLE_URL

try:
    TESSERACT_CMD = st.secrets.get("TESSERACT_CMD", "")
except Exception:
    TESSERACT_CMD = ""

if TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

HEADERS = {
    "Authorization": f"Bearer {TEABLE_TOKEN}",
    "Content-Type": "application/json",
}

GLOBAL_STYLE = """
<style>
.portal-box {
    padding: 18px 20px;
    border: 1px solid rgba(120,120,120,.22);
    border-radius: 16px;
    background: rgba(255,255,255,.03);
    margin-bottom: 14px;
}
.portal-title {
    font-size: 1.2rem;
    font-weight: 700;
    margin-bottom: 4px;
}
.tag-chip {
    display: inline-block;
    padding: 4px 10px;
    margin: 2px 6px 2px 0;
    border-radius: 999px;
    font-size: 0.82rem;
    border: 1px solid rgba(120,120,120,.25);
    background: rgba(255,255,255,.05);
}
.wip-chip {
    display: inline-block;
    padding: 4px 10px;
    border-radius: 999px;
    font-size: 0.82rem;
    font-weight: 600;
}
</style>
"""

PO_CANDIDATES = ["PO#", "PO", "P/O", "訂單編號", "訂單號", "訂單號碼", "工單", "工單號", "單號"]
CUSTOMER_CANDIDATES = ["Customer", "客戶", "客戶名稱"]
PART_CANDIDATES = ["Part No", "Part No.", "P/N", "客戶料號", "Cust. P / N", "LS P/N", "料號", "品號", "成品料號", "產品料號"]
QTY_CANDIDATES = ["Qty", "Order Q'TY (PCS)", "Order Q'TY\n (PCS)", "訂購量(PCS)", "訂購量", "Q'TY", "數量", "PCS", "訂單量", "生產數量", "投產數"]
FACTORY_CANDIDATES = ["Factory", "工廠", "廠編"]
WIP_CANDIDATES = ["WIP", "WIP Stage", "進度", "製程", "工序", "目前站別", "生產進度"]
FACTORY_DUE_CANDIDATES = ["Factory Due Date", "工廠交期", "交貨日期", "Required Ship date", "confrimed DD", "交期", "預交日", "預定交期", "交貨期"]
SHIP_DATE_CANDIDATES = ["Ship Date", "Ship date", "出貨日期", "交貨日期", "Required Ship date", "confrimed DD"]
REMARK_CANDIDATES = ["Remark", "備註", "情況", "備註說明", "Note", "說明", "異常備註"]
CUSTOMER_TAG_CANDIDATES = ["Customer Remark Tags", "Customer Tags", "客戶備註標籤"]

MERGE_DATE_CANDIDATES = ["併貨日期", "併貨日", "併貨日期(限內部使用)"]
ORDER_DATE_CANDIDATES = ["客戶下單日期", "下單日期", "Order Date"]
FACTORY_ORDER_DATE_CANDIDATES = ["工廠下單日期"]
CHANGED_DUE_DATE_CANDIDATES = ["交期(更改)", "更改交期", "工廠交期變更"]
SHIPMENT_CANDIDATES = ["SHIPMENT", "Shipment", "出貨"]

MULTI_SELECT_MODE = True

TAG_OPTIONS = [
    "Working Gerber for Approval",
    "Engineering Question",
    "Payment Pending",
    "Remake in Process",
    "On Hold",
    "Partial Shipment",
    "Shipped",
    "Waiting Confirmation",
]

DONE_WIP_VALUES = {"完成", "DONE", "COMPLETE", "COMPLETED", "FINISHED", "FINISH"}
