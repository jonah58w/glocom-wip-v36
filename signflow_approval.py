# -*- coding: utf-8 -*-
"""
signflow_approval.py  ──  GLOCOM 內部簽核平台 (SignFlow)
版本：v1.2  針對 glocom-wip-v36 的欄位結構與 Teable API 整合

安裝依賴：
    pip install streamlit-drawable-canvas
    （加入 requirements.txt）

在 app.py 使用：
    from signflow_approval import render_approval_page
    ...
    elif menu == "✍ 簽核平台":
        render_approval_page()
"""

from __future__ import annotations

import os
from datetime import date, datetime

import pandas as pd
import requests
import streamlit as st

try:
    from streamlit_drawable_canvas import st_canvas
    HAS_CANVAS = True
except ImportError:
    HAS_CANVAS = False


# ════════════════════════════════════════════════════════════════
#  CONFIG  ── 讀取 app.py 已有的 Teable 設定
# ════════════════════════════════════════════════════════════════
try:
    _TEABLE_TOKEN = st.secrets.get("TEABLE_TOKEN", "")
except Exception:
    _TEABLE_TOKEN = os.environ.get("TEABLE_TOKEN", "")

try:
    _TABLE_URL = st.secrets.get("TEABLE_TABLE_URL",
        "https://app.teable.ai/api/table/tbl6c05EPXYtJcZfeir/record")
except Exception:
    _TABLE_URL = "https://app.teable.ai/api/table/tbl6c05EPXYtJcZfeir/record"

_HEADERS = {
    "Authorization": f"Bearer {_TEABLE_TOKEN}",
    "Content-Type": "application/json",
}

# Teable 欄位名（與 app.py 一致）
_PO_FIELD       = "PO#"
_CUSTOMER_FIELD = "Customer"
_PART_FIELD     = "Part No"
_QTY_FIELD      = "Order Q'TY (PCS)"
_WIP_FIELD      = "WIP"
_SHIP_FIELD     = "Ship date"
_REMARK_FIELD   = "Remark"


# ════════════════════════════════════════════════════════════════
#  CONSTANTS
# ════════════════════════════════════════════════════════════════
STATUS_EMOJI  = {"approved": "✅", "current": "🔄", "pending": "⬜", "rejected": "❌"}
STATUS_LABEL  = {"approved": "已核准", "current": "審核中", "pending": "待審核", "rejected": "已退回"}

DOC_BADGE     = {"inv": "🧾 INV", "pck": "📦 PCK", "wip": "🏭 WIP"}

TEMPLATES: dict[str, dict] = {
    "inv": {
        "label": "🧾 Invoice 簽核",
        "desc": "發票核准：業務確認 → 財務核對 → 稅務審查 → 財務長 → 總經理",
        "fields": [
            ("Invoice 編號 *", "text", "INV-2026-XXX"),
            ("客戶名稱 *",     "text", ""),
            ("金額",           "text", "USD 0.00"),
            ("幣別",           "select", ["USD","EUR","TWD","JPY","CNY"]),
            ("Invoice 日期",   "date",  None),
            ("付款條件",       "select", ["Net 30","Net 60","Net 90","T/T in advance","L/C"]),
            ("負責業務",       "text", ""),
            ("Email",          "email", ""),
            ("備註",           "textarea", ""),
        ],
        "approvers": [
            {"name": "王業務",   "email": "wang@glocom.com",  "role": "業務確認"},
            {"name": "謝財務",   "email": "hsieh@glocom.com", "role": "財務核對"},
            {"name": "林稅務",   "email": "lin@glocom.com",   "role": "稅務審查"},
            {"name": "陳財務長", "email": "chen@glocom.com",  "role": "財務長核准"},
            {"name": "吳總經理", "email": "wu@glocom.com",    "role": "總經理核定"},
        ],
    },
    "pck": {
        "label": "📦 Packing List 簽核",
        "desc": "出貨清單：業務確認 → 倉儲核對 → 品管放行 → 出貨主管 → 總監",
        "fields": [
            ("PL 編號 *",       "text", "PL-2026-XXX"),
            ("對應訂單 PO#",    "text", ""),
            ("客戶名稱 *",      "text", ""),
            ("目的地",          "text", ""),
            ("總箱數 (CTN)",    "text", ""),
            ("總重量 (KG)",     "text", ""),
            ("出貨日期",        "date", None),
            ("運輸方式",        "select", ["Sea Freight","Air Freight","Express","Land Transport"]),
            ("負責業務",        "text", ""),
            ("Email",           "email", ""),
            ("品項說明",        "textarea", ""),
        ],
        "approvers": [
            {"name": "李出貨",   "email": "lee@glocom.com",   "role": "業務確認"},
            {"name": "張倉管",   "email": "chang@glocom.com", "role": "倉儲核對"},
            {"name": "劉品管",   "email": "liu@glocom.com",   "role": "品管放行"},
            {"name": "黃出貨長", "email": "huang@glocom.com", "role": "出貨主管"},
            {"name": "蔡總監",   "email": "tsai@glocom.com",  "role": "總監核定"},
        ],
    },
    "wip": {
        "label": "🏭 新訂單 WIP 簽核",
        "desc": "投產全程：業務接單 → 工程評估 → 生管排程 → 採購備料 → 廠長 → 總經理",
        "fields": [
            ("WIP 單號 *",    "text", "WIP-2026-XXX"),
            ("客戶 PO#",      "text", ""),
            ("客戶名稱 *",    "text", ""),
            ("產品型號",      "text", ""),
            ("訂單數量",      "text", "0 PCS"),
            ("訂單金額",      "text", "USD 0.00"),
            ("交貨期 (ETD)", "date", None),
            ("生產類別",      "select", ["標準品","客製化","樣品","試產"]),
            ("負責業務",      "text", ""),
            ("Email",         "email", ""),
            ("訂單說明",      "textarea", ""),
        ],
        "approvers": [
            {"name": "陳業務",   "email": "chen@glocom.com",  "role": "業務接單"},
            {"name": "方工程師", "email": "fang@glocom.com",  "role": "工程評估"},
            {"name": "鄭生管",   "email": "cheng@glocom.com", "role": "生管排程"},
            {"name": "林採購",   "email": "lin2@glocom.com",  "role": "採購備料"},
            {"name": "蘇廠長",   "email": "su@glocom.com",    "role": "廠長核准"},
            {"name": "吳總經理", "email": "wu@glocom.com",    "role": "總經理核定"},
        ],
    },
}

ALL_ROLES = [
    "業務確認","業務接單","財務核對","稅務審查","倉儲核對","品管放行",
    "出貨主管","工程評估","生管排程","採購備料","廠長核准","財務長核准",
    "總監核定","副總核准","總經理核定",
]


# ════════════════════════════════════════════════════════════════
#  SESSION STATE
# ════════════════════════════════════════════════════════════════
def _init() -> None:
    defaults = {
        "sf_docs":        _demo_docs(),
        "sf_view":        "list",          # list | detail | create
        "sf_current":     0,
        "sf_tpl":         "inv",
        "sf_email_sent":  False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


# ════════════════════════════════════════════════════════════════
#  DEMO DATA  (pre-populated sample documents)
# ════════════════════════════════════════════════════════════════
def _demo_docs() -> list[dict]:
    def _s(role, person, email, status, time=None, signed=False):
        return {"name": role, "person": person, "role": role,
                "email": email, "status": status, "time": time, "signed": signed}
    def _l(person, action, comment, time, typ):
        return {"person": person, "action": action, "comment": comment, "time": time, "type": typ}

    return [
        {
            "id": "INV-2026-031", "title": "Invoice #INV-2026-031",
            "doc_type": "inv", "customer": "ABC Electronics Co.",
            "applicant": "王業務", "email": "wang@glocom.com",
            "fields": {"金額": "USD 48,500", "幣別": "USD",
                       "付款條件": "Net 60", "日期": "2026-03-29"},
            "date": "2026-03-29", "status": "pending",
            "stations": [
                _s("業務確認",   "王業務",   "wang@glocom.com",  "approved", "2026-03-29 14:30", True),
                _s("財務核對",   "謝財務",   "hsieh@glocom.com", "current"),
                _s("稅務審查",   "林稅務",   "lin@glocom.com",   "pending"),
                _s("財務長核准", "陳財務長", "chen@glocom.com",  "pending"),
                _s("總經理核定", "吳總經理", "wu@glocom.com",    "pending"),
            ],
            "logs": [
                _l("王業務", "提交 Invoice 簽核", "ABC Electronics 訂單，請協助核對金額與稅務。", "2026-03-29 13:50", "submitted"),
                _l("王業務", "業務確認完成",       "金額與訂單一致，Invoice 無誤。",               "2026-03-29 14:30", "approved"),
            ],
        },
        {
            "id": "PL-2026-018", "title": "Packing List #PL-2026-018",
            "doc_type": "pck", "customer": "XYZ Trading GmbH",
            "applicant": "李出貨", "email": "lee@glocom.com",
            "fields": {"對應訂單": "PO-2026-045", "總箱數": "128 CTN",
                       "總重量": "3,240 KG", "目的地": "Hamburg, Germany",
                       "運輸方式": "Sea Freight"},
            "date": "2026-03-30", "status": "pending",
            "stations": [
                _s("業務確認",  "李出貨",   "lee@glocom.com",   "approved", "2026-03-30 09:00", True),
                _s("倉儲核對",  "張倉管",   "chang@glocom.com", "approved", "2026-03-30 11:20", True),
                _s("品管放行",  "劉品管",   "liu@glocom.com",   "current"),
                _s("出貨主管",  "黃出貨長", "huang@glocom.com", "pending"),
                _s("總監核定",  "蔡總監",   "tsai@glocom.com",  "pending"),
            ],
            "logs": [
                _l("李出貨", "提交 Packing List 簽核", "XYZ GmbH 出貨，128箱，請倉儲核對。",    "2026-03-30 08:30", "submitted"),
                _l("李出貨", "業務確認完成",            "品項與數量確認一致。",                  "2026-03-30 09:00", "approved"),
                _l("張倉管", "倉儲核對完成",            "箱數與重量核對完畢，符合出貨清單。",    "2026-03-30 11:20", "approved"),
            ],
        },
        {
            "id": "WIP-2026-009", "title": "新訂單 WIP #WIP-2026-009",
            "doc_type": "wip", "customer": "TechCorp USA Inc.",
            "applicant": "陳業務", "email": "chen@glocom.com",
            "fields": {"客戶PO": "PO-TC-20260331", "產品型號": "TC-X500-BLK",
                       "訂單數量": "5,000 PCS", "金額": "USD 125,000",
                       "交貨期": "2026-05-30"},
            "date": "2026-03-31", "status": "pending",
            "stations": [
                _s("業務接單",   "陳業務",   "chen@glocom.com",  "current"),
                _s("工程評估",   "方工程師", "fang@glocom.com",  "pending"),
                _s("生管排程",   "鄭生管",   "cheng@glocom.com", "pending"),
                _s("採購備料",   "林採購",   "lin2@glocom.com",  "pending"),
                _s("廠長核准",   "蘇廠長",   "su@glocom.com",    "pending"),
                _s("總經理核定", "吳總經理", "wu@glocom.com",    "pending"),
            ],
            "logs": [
                _l("陳業務", "提交新訂單 WIP 簽核", "TechCorp 美國新訂單，客製化產品，請依序審核。", "2026-03-31 09:15", "submitted"),
            ],
        },
        {
            "id": "INV-2026-028", "title": "Invoice #INV-2026-028",
            "doc_type": "inv", "customer": "Pacific Trade Co.",
            "applicant": "吳業務", "email": "wu2@glocom.com",
            "fields": {"金額": "USD 31,200", "幣別": "USD",
                       "付款條件": "T/T in advance", "日期": "2026-03-20"},
            "date": "2026-03-20", "status": "approved",
            "stations": [
                _s("業務確認",   "吳業務",   "wu2@glocom.com",   "approved", "2026-03-20 10:00", True),
                _s("財務核對",   "謝財務",   "hsieh@glocom.com", "approved", "2026-03-20 14:00", True),
                _s("稅務審查",   "林稅務",   "lin@glocom.com",   "approved", "2026-03-21 09:30", True),
                _s("財務長核准", "陳財務長", "chen@glocom.com",  "approved", "2026-03-21 16:00", True),
                _s("總經理核定", "吳總經理", "wu@glocom.com",    "approved", "2026-03-22 09:00", True),
            ],
            "logs": [
                _l("吳業務",   "提交 Invoice 簽核", "Pacific Trade T/T 預付款訂單。",  "2026-03-20 09:30", "submitted"),
                _l("吳業務",   "業務確認完成",       "金額確認無誤。",                  "2026-03-20 10:00", "approved"),
                _l("謝財務",   "財務核對完成",       "帳目核對正確，付款條件無異議。",  "2026-03-20 14:00", "approved"),
                _l("林稅務",   "稅務審查通過",       "稅務計算正確，可出具發票。",       "2026-03-21 09:30", "approved"),
                _l("陳財務長", "財務長核准",         "審核通過，同意付款。",             "2026-03-21 16:00", "approved"),
                _l("吳總經理", "總經理核定",         "Invoice 核定通過。",               "2026-03-22 09:00", "approved"),
            ],
        },
        {
            "id": "PL-2026-015", "title": "Packing List #PL-2026-015",
            "doc_type": "pck", "customer": "Euro Parts BV",
            "applicant": "李出貨", "email": "lee@glocom.com",
            "fields": {"對應訂單": "PO-2026-038", "總箱數": "64 CTN",
                       "總重量": "1,820 KG", "目的地": "Rotterdam",
                       "運輸方式": "Sea Freight"},
            "date": "2026-03-26", "status": "rejected",
            "stations": [
                _s("業務確認",  "李出貨",  "lee@glocom.com",   "approved", "2026-03-26 10:00", True),
                _s("倉儲核對",  "張倉管",  "chang@glocom.com", "approved", "2026-03-26 14:00", True),
                _s("品管放行",  "劉品管",  "liu@glocom.com",   "rejected", "2026-03-27 09:00", False),
                _s("出貨主管",  "黃出貨長","huang@glocom.com", "pending"),
                _s("總監核定",  "蔡總監",  "tsai@glocom.com",  "pending"),
            ],
            "logs": [
                _l("李出貨", "提交 Packing List", "Euro Parts BV 出貨清單。",              "2026-03-26 09:30", "submitted"),
                _l("李出貨", "業務確認完成",       "訂單數量與品項確認。",                 "2026-03-26 10:00", "approved"),
                _l("張倉管", "倉儲核對完成",       "實物點數完畢。",                       "2026-03-26 14:00", "approved"),
                _l("劉品管", "品管退回",           "Lot #PL-064 有外觀瑕疵，需重工後重新提交。", "2026-03-27 09:00", "rejected"),
            ],
        },
    ]


# ════════════════════════════════════════════════════════════════
#  CSS
# ════════════════════════════════════════════════════════════════
_CSS = """
<style>
/* ── pipeline ── */
.sf-pipeline{display:flex;gap:0;overflow-x:auto;padding:4px 0 14px;}
.sf-station{min-width:142px;flex:1;}
.sf-station-card{
  background:white;border:1.5px solid #d5d8de;
  border-radius:8px;padding:12px 11px;margin-right:15px;
}
.sf-station-card.approved{border-color:#2c7a4b;background:#f0faf3;}
.sf-station-card.current {border-color:#c8973a;background:#fffcf2;
  box-shadow:0 0 0 3px rgba(200,151,58,.18);}
.sf-station-card.rejected{border-color:#c0392b;background:#fff5f5;}
.sf-station-card.pending {opacity:.52;}
.sf-step  {font-size:9px;letter-spacing:2px;color:#999;font-family:monospace;margin-bottom:4px;}
.sf-sname {font-size:12px;font-weight:700;margin-bottom:2px;}
.sf-sperson{font-size:11px;color:#666;margin-bottom:7px;}
.sf-badge {display:inline-block;font-size:10px;padding:2px 7px;border-radius:10px;font-weight:600;}
.sf-badge.approved{background:rgba(44,122,75,.12);color:#2c7a4b;}
.sf-badge.current {background:rgba(200,151,58,.15);color:#9a6820;}
.sf-badge.pending {background:rgba(0,0,0,.07);color:#888;}
.sf-badge.rejected{background:rgba(192,57,43,.12);color:#c0392b;}
.sf-stime {font-size:9px;color:#bbb;margin-top:4px;font-family:monospace;}
.sf-ssig  {font-size:9px;color:#2c7a4b;margin-top:2px;}
/* ── sig board ── */
.sf-sigboard{display:flex;border:1px solid #d5d8de;border-radius:8px;overflow:hidden;margin-bottom:16px;}
.sf-sigcell{flex:1;padding:10px 12px;border-right:1px solid #d5d8de;min-width:100px;}
.sf-sigcell:last-child{border-right:none;}
.sf-sigrole{font-size:8px;letter-spacing:2px;color:#999;font-family:monospace;margin-bottom:2px;}
.sf-signame{font-size:11px;font-weight:700;margin-bottom:6px;}
.sf-sigline{height:40px;border-bottom:1px solid #999;margin-bottom:3px;
  display:flex;align-items:center;justify-content:center;font-size:18px;}
.sf-sigdate{font-size:9px;color:#aaa;font-family:monospace;}
/* ── log ── */
.sf-log{display:flex;gap:10px;margin-bottom:14px;}
.sf-av{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;
  justify-content:center;font-size:13px;font-weight:700;color:white;flex-shrink:0;}
.sf-av.submitted{background:#2c4a7a;}.sf-av.approved{background:#2c7a4b;}
.sf-av.rejected{background:#c0392b;}
.sf-lbody{flex:1;}
.sf-lmeta{font-size:11px;color:#888;margin-bottom:2px;}
.sf-lcomment{font-size:12px;background:#f5f0e8;border-left:3px solid #c8bfaa;
  padding:6px 10px;border-radius:0 4px 4px 0;margin-top:4px;}
.sf-lcomment.ac{border-left-color:#2c7a4b;}.sf-lcomment.rc{border-left-color:#c0392b;}
/* ── arrow ── */
.sf-arrow{display:flex;align-items:center;padding-top:26px;color:#ccc;font-size:20px;
  margin:0 -7px;z-index:3;flex-shrink:0;}
</style>
"""


# ════════════════════════════════════════════════════════════════
#  COMPONENT RENDERERS
# ════════════════════════════════════════════════════════════════
def _pipeline(stations: list[dict]) -> None:
    html = '<div class="sf-pipeline">'
    for i, s in enumerate(stations):
        sc = s["status"]
        t  = f'<div class="sf-stime">{s["time"]}</div>' if s.get("time") else ""
        sg = '<div class="sf-ssig">✍ 已簽名</div>' if s.get("signed") else ""
        arrow = ('' if i == len(stations) - 1
                 else '<div class="sf-arrow">›</div>')
        html += f"""
<div class="sf-station">
  <div class="sf-station-card {sc}">
    <div class="sf-step">STEP {i+1}</div>
    <div class="sf-sname">{s['name']}</div>
    <div class="sf-sperson">{s['person']}</div>
    <span class="sf-badge {sc}">{STATUS_EMOJI[sc]} {STATUS_LABEL[sc]}</span>
    {t}{sg}
  </div>
</div>{arrow}"""
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def _sig_board(stations: list[dict]) -> None:
    html = '<div class="sf-sigboard">'
    for s in stations:
        content = "✍" if s.get("signed") else ""
        color   = "#1a1410" if s.get("signed") else "#e8e8e8"
        dt      = s["time"].split(" ")[0] if s.get("time") else "—"
        html += f"""
<div class="sf-sigcell">
  <div class="sf-sigrole">{s['role']}</div>
  <div class="sf-signame">{s['person']}</div>
  <div class="sf-sigline" style="color:{color}">{content}</div>
  <div class="sf-sigdate">{dt}</div>
</div>"""
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def _log(logs: list[dict]) -> None:
    html = ""
    for lg in logs:
        cc  = "ac" if lg["type"] == "approved" else ("rc" if lg["type"] == "rejected" else "")
        av  = lg["person"][0]
        html += f"""
<div class="sf-log">
  <div class="sf-av {lg['type']}">{av}</div>
  <div class="sf-lbody">
    <div class="sf-lmeta"><strong>{lg['person']}</strong> &nbsp;·&nbsp; {lg['action']} &nbsp;·&nbsp; {lg['time']}</div>
    <div class="sf-lcomment {cc}">{lg['comment']}</div>
  </div>
</div>"""
    st.markdown(html, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════
#  ACTIONS
# ════════════════════════════════════════════════════════════════
def _do_approve(doc: dict, station: dict, comment: str) -> None:
    station.update(status="approved", time=_now(), signed=True)
    nxt = next((s for s in doc["stations"] if s["status"] == "pending"), None)
    if nxt:
        nxt["status"] = "current"
    else:
        doc["status"] = "approved"
    doc["logs"].append({"person": station["person"], "action": f"{station['role']}完成",
                        "comment": comment, "time": _now(), "type": "approved"})


def _do_reject(doc: dict, station: dict, comment: str) -> None:
    station.update(status="rejected", time=_now())
    doc["status"] = "rejected"
    doc["logs"].append({"person": station["person"], "action": "退回",
                        "comment": comment, "time": _now(), "type": "rejected"})


def _send_teable_remark(po: str, doc_id: str, doc_type: str, status: str) -> None:
    """把最新核准狀態寫回 Teable Remark 欄位（選用，不影響主流程）"""
    if not _TEABLE_TOKEN or not po:
        return
    try:
        label = {"inv": "Invoice", "pck": "Packing List", "wip": "WIP"}[doc_type]
        remark = f"[SignFlow] {label} {doc_id} — {status} {_now()}"
        # 先查 record_id
        resp = requests.get(
            _TABLE_URL,
            headers=_HEADERS,
            params={"fieldKeyType": "name", "cellFormat": "text",
                    "filterByFormula": f'{{PO#}} = "{po}"'},
            timeout=10,
        )
        if resp.status_code != 200:
            return
        records = resp.json().get("records", [])
        if not records:
            return
        rec_id = records[0].get("id", "")
        if not rec_id:
            return
        requests.patch(
            f"{_TABLE_URL}/{rec_id}",
            headers=_HEADERS,
            json={"record": {"fields": {_REMARK_FIELD: remark}}},
            timeout=10,
        )
    except Exception:
        pass   # 不影響主流程


# ════════════════════════════════════════════════════════════════
#  VIEWS
# ════════════════════════════════════════════════════════════════
def _view_list() -> None:
    docs = st.session_state.sf_docs
    total = len(docs)
    pend  = sum(1 for d in docs if d["status"] == "pending")
    appr  = sum(1 for d in docs if d["status"] == "approved")
    rej   = sum(1 for d in docs if d["status"] == "rejected")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("本月總件數", total)
    c2.metric("🟡 審核中", pend)
    c3.metric("✅ 已核准", appr)
    c4.metric("❌ 已退回", rej)
    st.divider()

    rows = []
    for i, d in enumerate(docs):
        sig_done  = sum(1 for s in d["stations"] if s.get("signed"))
        sig_total = len(d["stations"])
        done      = sum(1 for s in d["stations"] if s["status"] == "approved")
        cur       = next((s for s in d["stations"] if s["status"] == "current"), None)
        amount    = (d["fields"].get("金額") or d["fields"].get("訂單數量")
                     or d["fields"].get("總箱數") or "—")
        s_map     = {"pending": "🟡 審核中", "approved": "✅ 已核准", "rejected": "❌ 已退回"}
        rows.append({
            "類型":     DOC_BADGE[d["doc_type"]],
            "文件編號": d["id"],
            "客戶":     d["customer"],
            "金額/數量": amount,
            "目前關卡": (f"{cur['name']} ({done+1}/{sig_total})" if cur
                         else f"完成 {sig_total}/{sig_total}"),
            "✍ 簽名":   f"{sig_done}/{sig_total}",
            "狀態":     s_map[d["status"]],
        })

    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.caption("點擊下方按鈕進入詳情")

    # clickable buttons
    n_cols = min(len(docs), 4)
    cols = st.columns(n_cols)
    for i, d in enumerate(docs):
        with cols[i % n_cols]:
            badge = {"inv": "🧾", "pck": "📦", "wip": "🏭"}[d["doc_type"]]
            st_map = {"pending": "🟡", "approved": "✅", "rejected": "❌"}
            if st.button(f"{badge} {d['id']}\n{st_map[d['status']]}",
                         key=f"sf_open_{i}", use_container_width=True):
                st.session_state.sf_current = i
                st.session_state.sf_view    = "detail"
                st.rerun()


def _view_detail() -> None:
    idx = st.session_state.sf_current
    doc = st.session_state.sf_docs[idx]

    tl = {"inv": "🧾 Invoice", "pck": "📦 Packing List", "wip": "🏭 新訂單 WIP"}[doc["doc_type"]]
    st.subheader(f"{tl}  ·  {doc['title']}")
    st.caption(f"**{doc['customer']}** · 提交：{doc['applicant']} · {doc['date']}")

    ca, cb, _ = st.columns([1.2, 1.2, 7])
    with ca:
        if st.button("← 返回列表", use_container_width=True):
            st.session_state.sf_view = "list"; st.rerun()
    with cb:
        if st.button("📧 寄送通知", use_container_width=True):
            st.session_state.sf_email_sent = False
            _email_popup(doc)

    st.divider()

    # progress
    done  = sum(1 for s in doc["stations"] if s["status"] == "approved")
    total = len(doc["stations"])
    st.progress(done / total, text=f"簽核進度：{done}/{total} 關卡（{int(done/total*100)}%）")

    # pipeline
    st.markdown("**簽核流程**")
    _pipeline(doc["stations"])

    # sig board
    st.markdown("**✍ 簽核人員簽名欄**")
    _sig_board(doc["stations"])

    # info
    st.markdown("**文件資訊**")
    all_fields = {**doc["fields"],
                  "申請人": doc["applicant"],
                  "提交日期": doc["date"],
                  "✍ 簽名進度": f"{sum(1 for s in doc['stations'] if s.get('signed'))}/{total}"}
    info_cols = st.columns(3)
    for i, (k, v) in enumerate(all_fields.items()):
        with info_cols[i % 3]:
            st.metric(k, v)

    st.divider()

    # action panel
    cur = next((s for s in doc["stations"] if s["status"] == "current"), None)
    if cur and doc["status"] == "pending":
        with st.container(border=True):
            st.markdown(f"### ✍ 我的審核動作 — **{cur['person']}**（{cur['role']}）")

            if HAS_CANVAS:
                st.markdown("**手寫簽名** ← 請在白色區域簽名")
                canvas_result = st_canvas(
                    fill_color="rgba(255,255,255,0)",
                    stroke_width=2,
                    stroke_color="#1a1410",
                    background_color="#ffffff",
                    height=130,
                    width=480,
                    drawing_mode="freedraw",
                    key=f"sf_canvas_{idx}_{cur['role']}",
                )
                has_sig = (canvas_result.image_data is not None
                           and canvas_result.image_data.sum() > 0)
            else:
                st.info(
                    "💡 安裝 **streamlit-drawable-canvas** 可啟用手寫簽名。\n\n"
                    "```\npip install streamlit-drawable-canvas\n```\n\n"
                    "加入 `requirements.txt` 後重新部署即可。"
                )
                has_sig = st.checkbox(
                    "☑ 我確認以數位方式簽署（暫代手寫簽名）",
                    key=f"sf_chk_{idx}_{cur['role']}"
                )

            comment = st.text_area(
                "審核意見（選填）", placeholder="填寫審核意見...",
                key=f"sf_comment_{idx}_{cur['role']}"
            )

            cok, crej, _ = st.columns([2, 2, 6])
            with cok:
                if st.button("✅ 核准並完成簽名", type="primary", use_container_width=True,
                             key=f"sf_approve_{idx}"):
                    if not has_sig:
                        st.error("⚠️ 請先完成手寫簽名後再核准！")
                    else:
                        _do_approve(doc, cur, comment or "已審核，同意核准。")
                        if doc["status"] == "approved":
                            _send_teable_remark(
                                doc["fields"].get("客戶PO") or doc["id"],
                                doc["id"], doc["doc_type"], "已全部核准"
                            )
                        st.rerun()
            with crej:
                if st.button("❌ 退回申請", use_container_width=True,
                             key=f"sf_reject_{idx}"):
                    _do_reject(doc, cur, comment or "審核不通過，請修改後重新提交。")
                    st.rerun()

    elif doc["status"] == "approved":
        st.success("🎉 此文件已全部簽核通過！")
    elif doc["status"] == "rejected":
        st.error("❌ 此文件已被退回，請申請人修改後重新提交。")

    st.divider()
    st.markdown("**簽核歷程記錄**")
    _log(doc["logs"])


def _email_popup(doc: dict) -> None:
    cur       = next((s for s in doc["stations"] if s["status"] == "current"), None)
    tl        = {"inv": "Invoice", "pck": "Packing List", "wip": "新訂單 WIP"}[doc["doc_type"]]
    sig_done  = sum(1 for s in doc["stations"] if s.get("signed"))
    sig_total = len(doc["stations"])
    field_txt = "  \n".join(f"**{k}：** {v}" for k, v in doc["fields"].items())

    with st.expander("📧 Email 通知預覽（點擊展開）", expanded=True):
        st.markdown(f"""
**收件人：** `{cur['email'] if cur else '—'}`
**副本：** `{doc['email']}`
**主旨：** `【GLOCOM 待簽核】{tl} {doc['id']} — {doc['customer']}`

---
您好 **{cur['person'] if cur else ''}**，

以下 **{tl}** 正等待您的簽核（**{cur['role'] if cur else '—'}**），
請於 **2 個工作天**內完成審核並完成手寫電子簽名。

**文件編號：** {doc['id']}
**客戶：** {doc['customer']}
{field_txt}

**✍ 簽名進度：** {sig_done} / {sig_total} 人已完成

審核連結：`https://glocom-wip-v36-xko7x2byuccqjtpr9eaivj.streamlit.app/`
（請開啟後選擇左側 ✍ 簽核平台）

GLOCOM SignFlow 簽核系統
        """)
        if st.button("✉️ 確認寄出", type="primary", key="sf_send_email"):
            st.success("📧 通知信已寄出！")


def _view_create() -> None:
    st.subheader("新增簽核文件")

    # ── Step 1: Template ──
    st.markdown("**STEP 1 · 選擇簽核模板**")
    tpl_opts = {k: v["label"] for k, v in TEMPLATES.items()}
    tpl_key = st.radio(
        "模板", list(tpl_opts.keys()),
        format_func=lambda x: tpl_opts[x],
        horizontal=True,
        label_visibility="collapsed",
        index=list(tpl_opts.keys()).index(st.session_state.sf_tpl),
        key="sf_tpl_radio",
    )
    st.session_state.sf_tpl = tpl_key
    st.caption(TEMPLATES[tpl_key]["desc"])
    st.divider()

    # ── Step 2: Fields ──
    st.markdown("**STEP 2 · 填寫文件資訊**")
    field_values: dict[str, str] = {}
    grid_fields = TEMPLATES[tpl_key]["fields"]
    col_a, col_b = st.columns(2)
    toggle = True
    for fname, ftype, fdefault in grid_fields:
        target = col_a if toggle else col_b
        toggle = not toggle
        if ftype == "textarea":
            with st.container():
                field_values[fname] = st.text_area(fname, value=fdefault or "",
                                                   height=80, key=f"sf_f_{tpl_key}_{fname}")
            toggle = True   # reset after full-width
        elif ftype == "select":
            with target:
                field_values[fname] = st.selectbox(fname, fdefault,
                                                   key=f"sf_f_{tpl_key}_{fname}")
        elif ftype == "date":
            with target:
                dv = st.date_input(fname, value=date.today(),
                                   key=f"sf_f_{tpl_key}_{fname}")
                field_values[fname] = str(dv)
        else:
            with target:
                field_values[fname] = st.text_input(fname, value=fdefault or "",
                                                    key=f"sf_f_{tpl_key}_{fname}")

    st.divider()

    # ── Step 3: Approvers ──
    st.markdown("**STEP 3 · 簽核人員設定（5～6 位）**")
    tpl_approvers = TEMPLATES[tpl_key]["approvers"]
    num_ap = st.slider("簽核人數", 5, 6, len(tpl_approvers), key="sf_num_ap")
    new_stations: list[dict] = []
    for i in range(num_ap):
        default = tpl_approvers[i] if i < len(tpl_approvers) else \
                  {"name": "", "email": "", "role": ALL_ROLES[0]}
        ca, cb, cc = st.columns([2, 3, 2])
        name   = ca.text_input(f"姓名 {i+1}", value=default["name"], key=f"sf_ap_n_{tpl_key}_{i}")
        aemail = cb.text_input(f"Email {i+1}", value=default["email"], key=f"sf_ap_e_{tpl_key}_{i}")
        ridx   = ALL_ROLES.index(default["role"]) if default["role"] in ALL_ROLES else 0
        role   = cc.selectbox(f"職責 {i+1}", ALL_ROLES, index=ridx,
                              key=f"sf_ap_r_{tpl_key}_{i}")
        new_stations.append({
            "name": role, "person": name, "role": role, "email": aemail,
            "status": "current" if i == 0 else "pending",
            "time": None, "signed": False,
        })

    st.divider()

    # figure out doc_id and customer from field_values
    # fix: add parens to avoid operator precedence bug
    id_field   = next((v for k, v in field_values.items() if ("*" in k and ("編號" in k or "單號" in k))), "")
    cust_field = next((v for k, v in field_values.items() if "客戶名稱" in k), "")
    applicant  = next((v for k, v in field_values.items() if "負責業務" in k), "—")
    app_email  = field_values.get("Email", "")

    cok, ccancel, _ = st.columns([2, 1, 5])
    with cok:
        if st.button("🚀 提交並啟動流程", type="primary", use_container_width=True,
                     key=f"sf_submit_{tpl_key}"):
            if not id_field.strip():
                st.error("請填寫文件編號（標 * 欄位）！")
            elif not cust_field.strip():
                st.error("請填寫客戶名稱！")
            else:
                tpl_label_clean = TEMPLATES[tpl_key]["label"].split(" ", 1)[1].replace(" 簽核", "")
                new_doc = {
                    "id":        id_field.strip(),
                    "title":     f"{tpl_label_clean} #{id_field.strip()}",
                    "doc_type":  tpl_key,
                    "customer":  cust_field.strip(),
                    "applicant": applicant or "—",
                    "email":     app_email or "—",
                    "fields":    {k.rstrip(" *"): v for k, v in field_values.items()},
                    "date":      str(date.today()),
                    "status":    "pending",
                    "stations":  new_stations,
                    "logs": [{
                        "person":  applicant or "申請人",
                        "action":  "提交簽核申請",
                        "comment": field_values.get("備註", field_values.get("訂單說明", field_values.get("品項說明", "已提交，請依序審核。"))),
                        "time":    _now(),
                        "type":    "submitted",
                    }],
                }
                st.session_state.sf_docs.insert(0, new_doc)
                st.session_state.sf_current = 0
                st.session_state.sf_view    = "list"
                st.success(f"✅ {id_field} 已提交，簽核流程已啟動！")
                st.rerun()
    with ccancel:
        if st.button("取消", use_container_width=True, key="sf_cancel"):
            st.session_state.sf_view = "list"; st.rerun()


# ════════════════════════════════════════════════════════════════
#  MAIN ENTRY POINT  ← app.py 呼叫這裡
# ════════════════════════════════════════════════════════════════
def render_approval_page() -> None:
    """
    在 app.py 的 elif menu == "✍ 簽核平台": 呼叫此函式。
    """
    _init()
    st.markdown(_CSS, unsafe_allow_html=True)

    # ── 頁首 ──
    st.markdown(
        "## ✍ 簽核平台 "
        "<span style='font-size:14px;color:#c8973a;font-weight:400'>"
        "SignFlow &nbsp;·&nbsp; GLOCOM Internal</span>",
        unsafe_allow_html=True,
    )
    st.caption("Invoice  /  Packing List  /  新訂單 WIP  ——  5～6 人簽核 + 手寫電子簽名")

    # ── sub-nav ──
    c1, c2, c3, _ = st.columns([1.4, 1.4, 1.2, 6])
    with c1:
        if st.button("📋 簽核總覽", use_container_width=True,
                     type="primary" if st.session_state.sf_view == "list" else "secondary",
                     key="sf_nav_list"):
            st.session_state.sf_view = "list"; st.rerun()
    with c2:
        if st.button("🔍 簽核詳情", use_container_width=True,
                     type="primary" if st.session_state.sf_view == "detail" else "secondary",
                     key="sf_nav_detail"):
            st.session_state.sf_view = "detail"; st.rerun()
    with c3:
        if st.button("＋ 新增", use_container_width=True,
                     type="primary" if st.session_state.sf_view == "create" else "secondary",
                     key="sf_nav_create"):
            st.session_state.sf_view = "create"; st.rerun()

    st.divider()

    # ── route ──
    v = st.session_state.sf_view
    if v == "list":
        _view_list()
    elif v == "detail":
        if st.session_state.sf_docs:
            _view_detail()
        else:
            st.info("尚無文件，請先新增簽核。")
    elif v == "create":
        _view_create()
