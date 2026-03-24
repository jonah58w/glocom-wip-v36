# -*- coding: utf-8 -*-
"""
config.py
GLOCOM Control Tower 共用設定
"""

from __future__ import annotations

import streamlit as st
import pytesseract


# =========================================================
# Teable 設定
# =========================================================
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


# =========================================================
# OCR / Tesseract 設定
# =========================================================
try:
    TESSERACT_CMD = st.secrets.get("TESSERACT_CMD", "")
except Exception:
    TESSERACT_CMD = ""

if TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD


# =========================================================
# Sales / 外部報表路徑
# =========================================================
try:
    SALES_BASE_PATH = st.secrets.get("SALES_BASE_PATH", "")
except Exception:
    SALES_BASE_PATH = ""


# =========================================================
# API Headers
# =========================================================
HEADERS = {
    "Authorization": f"Bearer {TEABLE_TOKEN}",
    "Content-Type": "application/json",
}


# =========================================================
# 主鍵 / 欄位候選字
# =========================================================
PO_CANDIDATES = [
    "PO#",
    "PO",
    "P/O",
    "P O",
    "訂單編號",
    "訂單號",
    "訂單號碼",
    "工單",
    "工單號",
    "單號",
    "單據號",
    "ORDER NO",
    "ORDER#",
    "Order No",
    "Order Number",
    "工令",
    "工令號",
    "製令單號",
]

CUSTOMER_CANDIDATES = [
    "Customer",
    "客戶",
    "客戶名稱",
    "客戶名",
    "Company",
    "Customer Name",
]

PART_CANDIDATES = [
    "Part No",
    "Part No.",
    "P/N",
    "PN",
    "客戶料號",
    "Cust. P / N",
    "LS P/N",
    "料號",
    "品號",
    "成品料號",
    "產品料號",
    "客戶品號",
    "產品編號",
    "Product No",
    "Model",
    "料品編號",
    "品名規格",
    "祥竑料號",
]

QTY_CANDIDATES = [
    "Qty",
    "QTY",
    "Q'TY",
    "Order Q'TY (PCS)",
    "Order Q'TY\n (PCS)",
    "訂購量(PCS)",
    "訂購量",
    "Q'TY",
    "數量",
    "數量(PCS)",
    "PCS",
    "訂單量",
    "訂單量(PCS)",
    "生產數量",
    "投產數",
    "訂單數量",
    "未出貨數量",
]

FACTORY_CANDIDATES = [
    "Factory",
    "工廠",
    "廠編",
    "供應商",
    "Vendor",
]

WIP_CANDIDATES = [
    "WIP",
    "WIP Stage",
    "進度",
    "製程",
    "工序",
    "目前站別",
    "生產進度",
    "站別",
    "狀態",
]

FACTORY_DUE_CANDIDATES = [
    "Factory Due Date",
    "工廠交期",
    "交貨日期",
    "Required Ship date",
    "Required Ship Date",
    "confrimed DD",
    "confirmed DD",
    "交期",
    "預交日",
    "預定交期",
    "交貨期",
]

SHIP_DATE_CANDIDATES = [
    "Ship Date",
    "Ship date",
    "出貨日期",
    "交貨日期",
    "Required Ship date",
    "Required Ship Date",
    "confrimed DD",
    "confirmed DD",
]

REMARK_CANDIDATES = [
    "Remark",
    "備註",
    "情況",
    "備註說明",
    "Note",
    "說明",
    "異常備註",
]

CUSTOMER_TAG_CANDIDATES = [
    "Customer Remark Tags",
    "Customer Tags",
    "客戶備註標籤",
    "客戶標籤",
]


# =========================================================
# 日期欄位候選字
# =========================================================
MERGE_DATE_CANDIDATES = [
    "Merge Date",
    "合併日期",
    "併單日期",
]

ORDER_DATE_CANDIDATES = [
    "Order Date",
    "PO DATE",
    "下單日期",
    "訂單日期",
]

FACTORY_ORDER_DATE_CANDIDATES = [
    "Factory Order Date",
    "工廠下單日期",
    "工廠下單日",
]

CHANGED_DUE_DATE_CANDIDATES = [
    "Changed Due Date",
    "更改交期",
    "新交期",
    "改交期",
]


# =========================================================
# 客戶備註標籤
# =========================================================
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


# =========================================================
# WIP 完成值
# =========================================================
DONE_WIP_VALUES = {
    "完成",
    "DONE",
    "COMPLETE",
    "COMPLETED",
    "FINISHED",
    "FINISH",
    "CLOSED",
    "結案",
}


# =========================================================
# UI / 行為設定
# =========================================================
MULTI_SELECT_MODE = True


# =========================================================
# 製程關鍵字（供 parser / reader 參考）
# =========================================================
PROCESS_KEYWORDS = [
    "發料",
    "下料",
    "排版",
    "內層",
    "內乾",
    "內蝕",
    "黑化",
    "壓合",
    "壓板",
    "鑽孔",
    "沉銅",
    "一銅",
    "電鍍",
    "乾膜",
    "外層",
    "二銅",
    "二銅蝕刻",
    "AOI",
    "半測",
    "防焊",
    "文字",
    "噴錫",
    "化金",
    "OSP",
    "化銀",
    "成型",
    "V-CUT",
    "測試",
    "成檢",
    "包裝",
    "出貨",
    "庫存",
]
