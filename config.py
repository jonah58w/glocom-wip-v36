# -*- coding: utf-8 -*-
"""
GLOCOM Control Tower - 配置檔案
"""

# Teable API 配置
TABLE_URL = "https://app.teable.ai/api/table/tbl6c05EPXYtJcZfeir/record"
TEABLE_WEB_URL = "https://app.teable.ai/base/bsedgLzbHjiK0XoZH01/table/tbl6c05EPXYtJcZfeir"

# 請從 Streamlit Secrets 或環境變數載入 Token
try:
    import streamlit as st
    TEABLE_TOKEN = st.secrets.get("TEABLE_TOKEN", "")
except Exception:
    TEABLE_TOKEN = ""

# 欄位候補名稱（用於自動匹配不同工廠的 Excel 格式）
PO_CANDIDATES = [
    "PO#", "PO", "P/O", "訂單編號", "訂單號", "訂單號碼", "工單", "工單號", "單號"
]

CUSTOMER_CANDIDATES = ["Customer", "客戶", "客戶名稱"]

PART_CANDIDATES = [
    "Part No", "Part No.", "P/N", "客戶料號", "Cust. P / N", "LS P/N",
    "料號", "品號", "成品料號", "產品料號"
]

QTY_CANDIDATES = [
    "Qty", "Order Q'TY (PCS)", "Order Q'TY\n(PCS)", "Order Q'TY\n (PCS)",
    "訂購量(PCS)", "訂購量", "Q'TY", "數量", "PCS", "訂單量", "生產數量", "投產數",
]

FACTORY_CANDIDATES = ["Factory", "工廠", "廠編"]

WIP_CANDIDATES = ["WIP", "WIP Stage", "進度", "製程", "工序", "目前站別", "生產進度"]

FACTORY_DUE_CANDIDATES = [
    "Factory Due Date", "工廠交期", "交貨日期", "Required Ship date",
    "confrimed DD", "交期", "預交日", "預定交期", "交貨期",
]

SHIP_DATE_CANDIDATES = [
    "Ship Date", "Ship date", "出貨日期", "交貨日期",
    "Required Ship date", "confrimed DD",
]

REMARK_CANDIDATES = [
    "Remark", "備註", "情況", "備註說明", "Note", "說明", "異常備註", "異常"
]

CUSTOMER_TAG_CANDIDATES = ["Customer Remark Tags", "Customer Tags", "客戶備註標籤"]

MERGE_DATE_CANDIDATES = [
    "併貨日期", "併貨日", "併貨日期(限內部使用)", "併貨日期\n(限內部使用)",
]

ORDER_DATE_CANDIDATES = ["客戶下單日期", "客戶下\n單日期", "下單日期", "Order Date"]

FACTORY_ORDER_DATE_CANDIDATES = ["工廠下單日期", "工廠下\n單日期"]

CHANGED_DUE_DATE_CANDIDATES = ["交期(更改)", "更改交期", "工廠交期變更"]

SHIPMENT_CANDIDATES = ["SHIPMENT", "Shipment", "出貨"]

# 標籤選項（用於客戶備註）
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

# WIP 完成狀態值
DONE_WIP_VALUES = {"完成", "DONE", "COMPLETE", "COMPLETED", "FINISHED", "FINISH"}

# 其他設定
MULTI_SELECT_MODE = True
