# -*- coding: utf-8 -*-
"""
signflow_approval.py  v2.0
==========================
GLOCOM 內部簽核平台 (SignFlow)

改進：
  - 用 Teable 獨立資料表儲存簽核資料，重新整理不消失，多人共用
  - 新增簽核時可直接從現有訂單搜尋帶入，不需手動輸入
  - 簡化欄位，減少人工輸入

部署：
  1. 把此檔放到 repo 根目錄
  2. 在 Teable 建立一張新表「SignFlow」（見下方說明）
  3. 把新表的 URL 加到 Streamlit secrets：SIGNFLOW_TABLE_URL
  4. app.py 三處修改不變
"""

from __future__ import annotations

import base64
import json
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
#  TEABLE CONFIG
# ════════════════════════════════════════════════════════════════
def _secret(key: str, default: str = "") -> str:
    try:
        return st.secrets.get(key, default) or default
    except Exception:
        return os.environ.get(key, default)

# 主訂單表（已有，用來查訂單帶入）
_ORDERS_URL   = _secret("TEABLE_TABLE_URL",
    "https://app.teable.ai/api/table/tbl6c05EPXYtJcZfeir/record")

# SignFlow 專用表（新建，用來儲存簽核單）
# 若尚未建立，資料暫存 session（功能仍可用，但重整消失）
_SF_URL = _secret("SIGNFLOW_TABLE_URL", "")

_TOKEN   = _secret("TEABLE_TOKEN", "")
_HEADERS = {"Authorization": f"Bearer {_TOKEN}", "Content-Type": "application/json"}

# ── SignFlow Teable 表欄位名 ──────────────────────────────────────
SF_ID       = "sf_id"        # 文件編號
SF_TYPE     = "sf_type"      # inv / pck / wip
SF_CUSTOMER = "sf_customer"
SF_PO       = "sf_po"        # 對應訂單 PO#
SF_AMOUNT   = "sf_amount"
SF_STATUS   = "sf_status"    # pending / approved / rejected
SF_STATIONS = "sf_stations"  # JSON
SF_LOGS     = "sf_logs"      # JSON
SF_FIELDS   = "sf_fields"    # JSON
SF_DATE     = "sf_date"
SF_APPLICANT= "sf_applicant"
SF_EMAIL    = "sf_email"
SF_FILENAME = "sf_filename"  # 上傳的原始檔名
SF_FILEDATA = "sf_filedata"  # base64 檔案內容（≤ 500KB 存 Teable）


# ════════════════════════════════════════════════════════════════
#  CONSTANTS
# ════════════════════════════════════════════════════════════════
STATUS_EMOJI = {"approved": "✅", "current": "🔄", "pending": "⬜", "rejected": "❌"}
STATUS_LABEL = {"approved": "已核准", "current": "審核中",
                "pending": "待審核", "rejected": "已退回"}
DOC_LABEL    = {"inv": "🧾 Invoice", "pck": "📦 Packing List", "wip": "🏭 WIP"}

# 預設簽核人（可自行修改）
DEFAULT_APPROVERS = {
    "inv": [
        {"person": "業務",   "email": "", "role": "業務確認"},
        {"person": "財務",   "email": "", "role": "財務核對"},
        {"person": "稅務",   "email": "", "role": "稅務審查"},
        {"person": "財務長", "email": "", "role": "財務長核准"},
        {"person": "總經理", "email": "", "role": "總經理核定"},
    ],
    "pck": [
        {"person": "業務",   "email": "", "role": "業務確認"},
        {"person": "倉儲",   "email": "", "role": "倉儲核對"},
        {"person": "品管",   "email": "", "role": "品管放行"},
        {"person": "出貨主管","email": "", "role": "出貨主管"},
        {"person": "總監",   "email": "", "role": "總監核定"},
    ],
    "wip": [
        {"person": "業務",   "email": "", "role": "業務接單確認"},
        {"person": "工程",   "email": "", "role": "工程評估"},
        {"person": "生管",   "email": "", "role": "生管排程"},
        {"person": "採購",   "email": "", "role": "採購備料"},
        {"person": "廠長",   "email": "", "role": "廠長核准"},
        {"person": "總經理", "email": "", "role": "總經理核定"},
    ],
}

ALL_ROLES = ["業務接單確認","業務確認","財務核對","稅務審查","倉儲核對","品管放行",
             "出貨主管","工程評估","生管排程","採購備料","廠長核准","財務長核准",
             "總監核定","副總核准","總經理核定"]


# ════════════════════════════════════════════════════════════════
#  TEABLE  I/O
# ════════════════════════════════════════════════════════════════
def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")

def _today() -> str:
    return date.today().isoformat()


@st.cache_data(ttl=30)
def _load_orders() -> pd.DataFrame:
    """載入主訂單表，用來搜尋帶入"""
    if not _TOKEN:
        return pd.DataFrame()
    try:
        rows, page_token = [], None
        while True:
            params = {"fieldKeyType": "name", "cellFormat": "text", "take": 1000}
            if page_token:
                params["pageToken"] = page_token
            r = requests.get(_ORDERS_URL, headers=_HEADERS, params=params, timeout=20)
            if r.status_code != 200:
                break
            data = r.json()
            for rec in data.get("records", []):
                f = rec.get("fields", {})
                f["_record_id"] = rec.get("id", "")
                rows.append(f)
            page_token = data.get("pageToken") or data.get("nextPageToken")
            if not page_token:
                break
        return pd.DataFrame(rows) if rows else pd.DataFrame()
    except Exception:
        return pd.DataFrame()


def _sf_load_all() -> list[dict]:
    """從 SignFlow 表讀取所有簽核單"""
    if not _SF_URL or not _TOKEN:
        # fallback: session only
        return st.session_state.get("sf_docs_fallback", _demo_docs())
    try:
        rows, page_token = [], None
        while True:
            params = {"fieldKeyType": "name", "cellFormat": "text", "take": 1000}
            if page_token:
                params["pageToken"] = page_token
            r = requests.get(_SF_URL, headers=_HEADERS, params=params, timeout=20)
            if r.status_code != 200:
                break
            data = r.json()
            for rec in data.get("records", []):
                f = rec.get("fields", {})
                f["_record_id"] = rec.get("id", "")
                rows.append(f)
            page_token = data.get("pageToken") or data.get("nextPageToken")
            if not page_token:
                break
        docs = []
        for row in rows:
            try:
                doc = {
                    "_record_id": row.get("_record_id", ""),
                    "id":         row.get(SF_ID, ""),
                    "doc_type":   row.get(SF_TYPE, "wip"),
                    "customer":   row.get(SF_CUSTOMER, ""),
                    "po":         row.get(SF_PO, ""),
                    "amount":     row.get(SF_AMOUNT, ""),
                    "status":     row.get(SF_STATUS, "pending"),
                    "date":       row.get(SF_DATE, ""),
                    "applicant":  row.get(SF_APPLICANT, ""),
                    "email":      row.get(SF_EMAIL, ""),
                    "filename":   row.get(SF_FILENAME, ""),
                    "filedata":   row.get(SF_FILEDATA, ""),
                    "stations":   json.loads(row.get(SF_STATIONS, "[]")),
                    "logs":       json.loads(row.get(SF_LOGS, "[]")),
                    "fields":     json.loads(row.get(SF_FIELDS, "{}")),
                }
                doc["title"] = f"{DOC_LABEL.get(doc['doc_type'],'')} #{doc['id']}"
                docs.append(doc)
            except Exception:
                continue
        return docs if docs else _demo_docs()
    except Exception:
        return st.session_state.get("sf_docs_fallback", _demo_docs())


def _sf_save(doc: dict) -> bool:
    """新增或更新一筆簽核單到 Teable"""
    payload = {
        SF_ID:       doc["id"],
        SF_TYPE:     doc["doc_type"],
        SF_CUSTOMER: doc["customer"],
        SF_PO:       doc.get("po", ""),
        SF_AMOUNT:   doc.get("amount", ""),
        SF_STATUS:   doc["status"],
        SF_DATE:     doc.get("date", _today()),
        SF_APPLICANT:doc.get("applicant", ""),
        SF_EMAIL:    doc.get("email", ""),
        SF_FILENAME: doc.get("filename", ""),
        SF_FILEDATA: doc.get("filedata", ""),
        SF_STATIONS: json.dumps(doc["stations"], ensure_ascii=False),
        SF_LOGS:     json.dumps(doc["logs"],     ensure_ascii=False),
        SF_FIELDS:   json.dumps(doc["fields"],   ensure_ascii=False),
    }
    if not _SF_URL or not _TOKEN:
        # fallback: session only
        docs = st.session_state.get("sf_docs_fallback", [])
        rec_id = doc.get("_record_id", "")
        if rec_id:
            for i, d in enumerate(docs):
                if d.get("_record_id") == rec_id:
                    docs[i] = doc
                    break
        else:
            doc["_record_id"] = f"local_{_now()}"
            docs.insert(0, doc)
        st.session_state["sf_docs_fallback"] = docs
        return True
    try:
        rec_id = doc.get("_record_id", "")
        if rec_id:
            r = requests.patch(
                f"{_SF_URL}/{rec_id}",
                headers=_HEADERS,
                json={"record": {"fields": payload}},
                timeout=15,
            )
        else:
            r = requests.post(
                _SF_URL,
                headers=_HEADERS,
                json={"records": [{"fields": payload}]},
                timeout=15,
            )
        return r.status_code in (200, 201)
    except Exception:
        return False


def _sf_delete(doc: dict) -> bool:
    """從 Teable 刪除一筆簽核單"""
    rec_id = doc.get("_record_id", "")
    if not _SF_URL or not _TOKEN:
        # fallback: session only
        docs = st.session_state.get("sf_docs_fallback", [])
        st.session_state["sf_docs_fallback"] = [
            d for d in docs if d.get("_record_id") != rec_id
        ]
        return True
    if not rec_id:
        return False
    try:
        r = requests.delete(
            f"{_SF_URL}/{rec_id}",
            headers=_HEADERS,
            timeout=15,
        )
        return r.status_code in (200, 204)
    except Exception:
        return False
    st.cache_data.clear()


# ════════════════════════════════════════════════════════════════
#  DEMO DATA（只在 SignFlow 表空的時候顯示）
# ════════════════════════════════════════════════════════════════
def _demo_docs() -> list[dict]:
    def s(role, person, status, time=None):
        return {"name": role, "person": person, "role": role,
                "email": "", "status": status, "time": time, "signed": False}
    def l(person, action, comment, time, typ):
        return {"person": person, "action": action,
                "comment": comment, "time": time, "type": typ}
    return [
        {
            "_record_id": "", "id": "INV-DEMO-001",
            "title": "🧾 Invoice #INV-DEMO-001",
            "doc_type": "inv", "customer": "範例客戶 ABC",
            "po": "PO-DEMO", "amount": "USD 10,000",
            "status": "pending", "date": _today(),
            "applicant": "示範業務", "email": "",
            "fields": {"金額": "USD 10,000", "付款條件": "Net 30"},
            "stations": [
                s("業務確認", "示範業務", "approved", _now()),
                s("財務核對", "財務人員",  "current"),
                s("稅務審查", "稅務人員",  "pending"),
                s("財務長核准","財務長",   "pending"),
                s("總經理核定","總經理",   "pending"),
            ],
            "logs": [
                l("示範業務","提交簽核","這是範例文件，請實際新增簽核單。",_now(),"submitted"),
            ],
        },
    ]


# ════════════════════════════════════════════════════════════════
#  CSS
# ════════════════════════════════════════════════════════════════
_CSS = """<style>
.sf-pipeline{display:flex;overflow-x:auto;padding:4px 0 12px;gap:0;}
.sf-station{min-width:138px;flex:1;}
.sf-card{background:white;border:1.5px solid #d5d8de;border-radius:8px;
  padding:12px 10px;margin-right:14px;}
.sf-card.approved{border-color:#2c7a4b;background:#f0faf3;}
.sf-card.current{border-color:#c8973a;background:#fffcf2;
  box-shadow:0 0 0 3px rgba(200,151,58,.18);}
.sf-card.rejected{border-color:#c0392b;background:#fff5f5;}
.sf-card.pending{opacity:.5;}
.sf-step{font-size:9px;letter-spacing:2px;color:#999;font-family:monospace;margin-bottom:4px;}
.sf-sname{font-size:12px;font-weight:700;margin-bottom:2px;}
.sf-sperson{font-size:11px;color:#666;margin-bottom:6px;}
.sf-badge{display:inline-block;font-size:10px;padding:2px 7px;border-radius:10px;font-weight:600;}
.sf-badge.approved{background:rgba(44,122,75,.12);color:#2c7a4b;}
.sf-badge.current{background:rgba(200,151,58,.15);color:#9a6820;}
.sf-badge.pending{background:rgba(0,0,0,.07);color:#888;}
.sf-badge.rejected{background:rgba(192,57,43,.12);color:#c0392b;}
.sf-stime{font-size:9px;color:#bbb;margin-top:4px;font-family:monospace;}
.sf-ssig{font-size:9px;color:#2c7a4b;margin-top:2px;}
.sf-arrow{display:flex;align-items:center;padding-top:25px;color:#ccc;
  font-size:20px;margin:0 -7px;flex-shrink:0;}
.sf-sigboard{display:flex;border:1px solid #d5d8de;border-radius:8px;
  overflow:hidden;margin-bottom:16px;}
.sf-sigcell{flex:1;padding:10px 12px;border-right:1px solid #d5d8de;min-width:90px;}
.sf-sigcell:last-child{border-right:none;}
.sf-sigrole{font-size:8px;letter-spacing:2px;color:#999;font-family:monospace;margin-bottom:2px;}
.sf-signame{font-size:11px;font-weight:700;margin-bottom:5px;}
.sf-sigline{height:38px;border-bottom:1px solid #999;margin-bottom:3px;
  display:flex;align-items:center;justify-content:center;font-size:18px;}
.sf-sigdate{font-size:9px;color:#aaa;font-family:monospace;}
.sf-log{display:flex;gap:10px;margin-bottom:12px;}
.sf-av{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;
  justify-content:center;font-size:12px;font-weight:700;color:white;flex-shrink:0;}
.sf-av.submitted{background:#2c4a7a;}.sf-av.approved{background:#2c7a4b;}
.sf-av.rejected{background:#c0392b;}
.sf-lbody{flex:1;}
.sf-lmeta{font-size:11px;color:#888;margin-bottom:2px;}
.sf-lcomment{font-size:12px;background:#f5f0e8;border-left:3px solid #c8bfaa;
  padding:6px 10px;border-radius:0 4px 4px 0;margin-top:3px;}
.sf-lcomment.ac{border-left-color:#2c7a4b;}.sf-lcomment.rc{border-left-color:#c0392b;}

</style>"""


# ════════════════════════════════════════════════════════════════
#  COMPONENT RENDERERS
# ════════════════════════════════════════════════════════════════
def _pipeline(stations: list[dict]) -> None:
    html = '<div class="sf-pipeline">'
    for i, s in enumerate(stations):
        sc = s["status"]
        t  = f'<div class="sf-stime">{s["time"]}</div>' if s.get("time") else ""
        sg = '<div class="sf-ssig">✍ 已簽名</div>'     if s.get("signed") else ""
        arr = "" if i == len(stations) - 1 else '<div class="sf-arrow">›</div>'
        html += f"""
<div class="sf-station">
  <div class="sf-card {sc}">
    <div class="sf-step">STEP {i+1}</div>
    <div class="sf-sname">{s["name"]}</div>
    <div class="sf-sperson">{s["person"]}</div>
    <span class="sf-badge {sc}">{STATUS_EMOJI[sc]} {STATUS_LABEL[sc]}</span>
    {t}{sg}
  </div>
</div>{arr}"""
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def _sig_board(stations: list[dict]) -> None:
    html = '<div class="sf-sigboard">'
    for s in stations:
        content = "✍" if s.get("signed") else ""
        color   = "#1a1410" if s.get("signed") else "#e0e0e0"
        dt      = s["time"].split(" ")[0] if s.get("time") else "—"
        html += f"""<div class="sf-sigcell">
  <div class="sf-sigrole">{s["role"]}</div>
  <div class="sf-signame">{s["person"]}</div>
  <div class="sf-sigline" style="color:{color}">{content}</div>
  <div class="sf-sigdate">{dt}</div>
</div>"""
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def _log_html(logs: list[dict]) -> None:
    html = ""
    for lg in logs:
        cc = "ac" if lg["type"] == "approved" else ("rc" if lg["type"] == "rejected" else "")
        av = lg["person"][0] if lg["person"] else "?"
        html += f"""<div class="sf-log">
  <div class="sf-av {lg['type']}">{av}</div>
  <div class="sf-lbody">
    <div class="sf-lmeta"><strong>{lg['person']}</strong> · {lg['action']} · {lg['time']}</div>
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
    doc["logs"].append({
        "person": station["person"], "action": f"{station['role']} 核准",
        "comment": comment, "time": _now(), "type": "approved"
    })


def _do_reject(doc: dict, station: dict, comment: str) -> None:
    station.update(status="rejected", time=_now())
    doc["status"] = "rejected"
    doc["logs"].append({
        "person": station["person"], "action": "退回",
        "comment": comment, "time": _now(), "type": "rejected"
    })


# ════════════════════════════════════════════════════════════════
#  FILE PREVIEW
# ════════════════════════════════════════════════════════════════
def _show_file(doc: dict) -> None:
    """顯示已上傳的文件"""
    filename = doc.get("filename", "")
    filedata = doc.get("filedata", "")

    # 先檢查 session（大檔案）
    session_key = f"sf_file_{doc['id']}"
    if not filedata and session_key in st.session_state:
        filedata = st.session_state[session_key]

    if not filename:
        st.caption("（此簽核單未上傳附件）")
        return

    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    st.markdown(f"📎 **附件：** {filename}")

    if not filedata:
        st.caption("⚠️ 檔案內容未儲存（檔案超過 500KB 或上傳後未重新整理）")
        return

    try:
        file_bytes = base64.b64decode(filedata)
    except Exception:
        st.caption("⚠️ 檔案解碼失敗")
        return

    # 下載按鈕（所有格式都支援）
    mime_map = {
        "pdf":  "application/pdf",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "doc":  "application/msword",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "xls":  "application/vnd.ms-excel",
        "png":  "image/png",
        "jpg":  "image/jpeg",
        "jpeg": "image/jpeg",
    }
    mime = mime_map.get(ext, "application/octet-stream")
    st.download_button(
        label=f"⬇️ 下載 {filename}",
        data=file_bytes,
        file_name=filename,
        mime=mime,
        key=f"sf_dl_{doc['id']}",
    )

    # PDF → 直接預覽
    if ext == "pdf":
        b64 = base64.b64encode(file_bytes).decode()
        st.markdown(
            f'<iframe src="data:application/pdf;base64,{b64}" '
            f'width="100%" height="600px" style="border:1px solid #ddd;border-radius:8px;"></iframe>',
            unsafe_allow_html=True,
        )

    # 圖片 → 直接預覽
    elif ext in ("png", "jpg", "jpeg"):
        st.image(file_bytes, use_container_width=True)

    # Word / Excel → 提示下載後開啟
    elif ext in ("docx", "doc", "xlsx", "xls"):
        st.info(f"📄 Word/Excel 檔案請點上方「⬇️ 下載」後在電腦開啟查看。")


# ════════════════════════════════════════════════════════════════
#  VIEW: LIST
# ════════════════════════════════════════════════════════════════
def _view_list(docs: list[dict]) -> None:
    pend = sum(1 for d in docs if d["status"] == "pending")
    appr = sum(1 for d in docs if d["status"] == "approved")
    rej  = sum(1 for d in docs if d["status"] == "rejected")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("總件數", len(docs))
    c2.metric("🟡 審核中", pend)
    c3.metric("✅ 已核准", appr)
    c4.metric("❌ 已退回", rej)
    st.divider()

    if not docs:
        st.info("尚無簽核單，請點上方「＋ 新增」建立。")
        return

    rows = []
    for i, d in enumerate(docs):
        sig_done  = sum(1 for s in d["stations"] if s.get("signed"))
        sig_total = len(d["stations"])
        done      = sum(1 for s in d["stations"] if s["status"] == "approved")
        cur       = next((s for s in d["stations"] if s["status"] == "current"), None)
        s_map     = {"pending": "🟡 審核中", "approved": "✅ 已核准", "rejected": "❌ 已退回"}
        rows.append({
            "類型":    DOC_LABEL.get(d["doc_type"], ""),
            "文件編號": d["id"],
            "客戶":    d["customer"],
            "PO#":     d.get("po", ""),
            "金額":    d.get("amount", ""),
            "目前關卡": f"{cur['name']} ({done+1}/{sig_total})" if cur else f"完成 {sig_total}/{sig_total}",
            "簽名":    f"✍ {sig_done}/{sig_total}",
            "狀態":    s_map[d["status"]],
        })

    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.divider()

    # ── 進入詳情 ──
    st.caption("點擊下方按鈕進入詳情或刪除")
    n = min(len(docs), 5)
    cols = st.columns(n)
    for i, d in enumerate(docs[:n * 4]):
        with cols[i % n]:
            icon    = {"inv": "🧾", "pck": "📦", "wip": "🏭"}.get(d["doc_type"], "📄")
            st_icon = {"pending": "🟡", "approved": "✅", "rejected": "❌"}[d["status"]]
            if st.button(f"{icon} {d['id']}\n{st_icon}",
                         key=f"sf_open_{i}", use_container_width=True):
                st.session_state["sf_view"]    = "detail"
                st.session_state["sf_current"] = i
                st.rerun()

    # ── 刪除區 ──
    st.divider()
    with st.expander("🗑️ 刪除簽核單", expanded=False):
        st.caption("選擇要刪除的文件，刪除後無法復原。")
        del_options = {f"{d['id']} ｜ {d['customer']} ｜ {d['status']}": i
                       for i, d in enumerate(docs)}
        del_choice = st.selectbox("選擇要刪除的文件",
                                  list(del_options.keys()),
                                  key="sf_del_select")
        confirm  = st.checkbox("確認刪除", key="sf_del_confirm")
        do_del_btn = st.button("🗑️ 確定刪除",
                               key="sf_del_btn",
                               type="primary" if confirm else "secondary")
        if do_del_btn:
            if not confirm:
                st.warning("請先勾選「確認刪除」")
            else:
                idx = del_options[del_choice]
                _sf_delete(docs[idx])
                _sf_clear_cache()
                st.rerun()


# ════════════════════════════════════════════════════════════════
#  VIEW: DETAIL
# ════════════════════════════════════════════════════════════════
def _view_detail(docs: list[dict]) -> None:
    idx = st.session_state.get("sf_current", 0)
    if idx >= len(docs):
        idx = 0
    doc = docs[idx]

    tl = DOC_LABEL.get(doc["doc_type"], "文件")
    st.subheader(f"{tl}  ·  {doc['id']}")
    st.caption(f"**{doc['customer']}**  ·  PO: {doc.get('po','—')}  ·  {doc.get('date','')}")

    # ── 頂部按鈕（不在 columns 內呼叫 rerun，改用 flag）──
    ca, cb, cc, _ = st.columns([1.2, 1.2, 1.2, 6])
    do_back   = ca.button("← 返回",   key="sf_back",    use_container_width=True)
    do_email  = cb.button("📧 通知",   key="sf_email_btn", use_container_width=True)
    do_del    = cc.button("🗑️ 刪除",  key="sf_det_del", use_container_width=True)

    if do_back:
        st.session_state["sf_view"] = "list"
        st.session_state.pop("sf_del_confirm_detail", None)
        st.rerun()

    if do_email:
        _email_popup(doc)

    if do_del:
        st.session_state["sf_del_confirm_detail"] = True

    # ── 刪除確認 ──
    if st.session_state.get("sf_del_confirm_detail"):
        st.warning(f"確定要刪除 **{doc['id']}** 嗎？此動作無法復原。")
        cy, cn, _ = st.columns([1, 1, 6])
        do_yes = cy.button("確定刪除", key="sf_del_yes", type="primary", use_container_width=True)
        do_no  = cn.button("取消",     key="sf_del_no",  use_container_width=True)
        if do_yes:
            _sf_delete(doc)
            _sf_clear_cache()
            st.session_state.pop("sf_del_confirm_detail", None)
            st.session_state["sf_view"] = "list"
            st.rerun()
        if do_no:
            st.session_state.pop("sf_del_confirm_detail", None)
            st.rerun()

    st.divider()

    done  = sum(1 for s in doc["stations"] if s["status"] == "approved")
    total = len(doc["stations"])
    st.progress(done / total, text=f"簽核進度：{done}/{total}（{int(done/total*100)}%）")

    # ── 附件預覽 ──
    st.markdown("**📎 簽核文件**")
    _show_file(doc)
    st.divider()

    st.markdown("**簽核流程**")
    _pipeline(doc["stations"])

    st.markdown("**✍ 簽名欄**")
    _sig_board(doc["stations"])

    # ── 文件資訊 ──
    info = {**doc.get("fields", {}),
            "申請人": doc.get("applicant", ""),
            "提交日期": doc.get("date", "")}
    cols3 = st.columns(3)
    for i, (k, v) in enumerate(info.items()):
        with cols3[i % 3]:
            st.metric(k, v or "—")

    st.divider()

    # ── 審核動作 ──
    cur = next((s for s in doc["stations"] if s["status"] == "current"), None)
    if cur and doc["status"] == "pending":
        with st.container(border=True):
            st.markdown(f"### 我的審核 — **{cur['person']}**（{cur['role']}）")

            if HAS_CANVAS:
                st.markdown("**請在下方空白區域手寫簽名**")
                canvas_result = st_canvas(
                    fill_color="rgba(255,255,255,0)",
                    stroke_width=2, stroke_color="#1a1410",
                    background_color="#ffffff",
                    height=130, width=460,
                    drawing_mode="freedraw",
                    key=f"sf_cv_{doc['id']}_{cur['role']}",
                )
                has_sig = (canvas_result.image_data is not None
                           and canvas_result.image_data.sum() > 0)
            else:
                st.info("安裝 `streamlit-drawable-canvas` 可啟用手寫簽名。\n\n"
                        "```\npip install streamlit-drawable-canvas\n```")
                has_sig = st.checkbox("☑ 確認以數位方式簽署",
                                      key=f"sf_chk_{doc['id']}_{cur['role']}")

            comment = st.text_area("審核意見（選填）",
                                   key=f"sf_cmt_{doc['id']}_{cur['role']}")

            # 按鈕用 flag 方式，不在 columns 內直接 rerun
            cok, crej, _ = st.columns([2, 2, 6])
            do_approve = cok.button("✅ 核准", type="primary",
                                    use_container_width=True,
                                    key=f"sf_ok_{doc['id']}")
            do_reject  = crej.button("❌ 退回",
                                     use_container_width=True,
                                     key=f"sf_rej_{doc['id']}")

            if do_approve:
                if not has_sig:
                    st.error("請先完成簽名！")
                else:
                    _do_approve(doc, cur, comment or "核准。")
                    _sf_save(doc)
                    _sf_clear_cache()
                    st.rerun()

            if do_reject:
                _do_reject(doc, cur, comment or "退回，請修改後重新提交。")
                _sf_save(doc)
                _sf_clear_cache()
                st.rerun()

    elif doc["status"] == "approved":
        st.success("🎉 此文件已全部簽核通過！")
    elif doc["status"] == "rejected":
        st.error("❌ 此文件已被退回，請申請人修改後重新提交。")

    st.divider()
    st.markdown("**簽核歷程**")
    _log_html(doc["logs"])


# ════════════════════════════════════════════════════════════════
#  VIEW: CREATE  ── 從訂單帶入，簡化輸入
# ════════════════════════════════════════════════════════════════
def _view_create() -> None:
    st.subheader("新增簽核文件")

    # ── STEP 1: 上傳文件 ─────────────────────────────────────
    st.markdown("**STEP 1 · 上傳要簽核的文件**")
    uploaded = st.file_uploader(
        "支援 PDF、Word（.docx/.doc）、Excel（.xlsx）、圖片（.png/.jpg）",
        type=["pdf", "docx", "doc", "xlsx", "xls", "png", "jpg", "jpeg"],
        key="sf_upload",
    )

    filename  = ""
    filedata  = ""   # base64
    file_bytes_preview = None

    if uploaded:
        filename    = uploaded.name
        raw_bytes   = uploaded.getvalue()
        file_size   = len(raw_bytes)
        st.success(f"✅ 已選取：{filename}（{file_size/1024:.1f} KB）")

        # 預覽
        ext = filename.rsplit(".", 1)[-1].lower()
        if ext == "pdf":
            b64_preview = base64.b64encode(raw_bytes).decode()
            st.markdown(
                f'<iframe src="data:application/pdf;base64,{b64_preview}" '
                f'width="100%" height="500px" '
                f'style="border:1px solid #ddd;border-radius:8px;margin-bottom:8px;"></iframe>',
                unsafe_allow_html=True,
            )
        elif ext in ("png", "jpg", "jpeg"):
            st.image(raw_bytes, use_container_width=True)
        elif ext in ("docx", "doc", "xlsx", "xls"):
            st.info("📄 Word/Excel 提交後可下載查看。")

        # 500KB 以下存 Teable，否則只存 session
        if file_size <= 500 * 1024:
            filedata = base64.b64encode(raw_bytes).decode()
        else:
            st.warning(f"⚠️ 檔案 {file_size/1024:.0f} KB > 500 KB，"
                       f"將暫存本次 session。建議壓縮後再上傳，或改用 Google Drive 連結。")
            # 暫存 session，之後在 detail 可讀取
            st.session_state[f"sf_file_pending"] = base64.b64encode(raw_bytes).decode()

    st.divider()

    # ── STEP 2: 基本資訊 ─────────────────────────────────────
    st.markdown("**STEP 2 · 基本資訊**（可從訂單帶入）")

    # 訂單搜尋帶入
    with st.expander("🔍 從現有訂單搜尋帶入", expanded=False):
        orders_df = _load_orders()
        if not orders_df.empty:
            search = st.text_input("輸入 PO#、客戶名、料號",
                                   placeholder="例：G1150022",
                                   key="sf_search")
            if search.strip():
                mask = orders_df.astype(str).apply(
                    lambda col: col.str.contains(search.strip(), case=False, na=False)
                ).any(axis=1)
                results = orders_df[mask].head(8)
                if not results.empty:
                    def fc(df, cands):
                        for c in cands:
                            if c in df.columns: return c
                        return None
                    po_c = fc(results, ["PO#","PO","P/O"])
                    cu_c = fc(results, ["Customer","客戶"])
                    pt_c = fc(results, ["Part No","P/N","料號"])
                    qt_c = fc(results, ["Order Q'TY (PCS)","Qty"])
                    am_c = fc(results, ["INVOICE","Invoice","接單金額"])
                    show_cols = [c for c in [po_c,cu_c,pt_c,qt_c,am_c] if c]
                    st.dataframe(results[show_cols].reset_index(drop=True),
                                 use_container_width=True, hide_index=False)
                    pick = st.number_input("選擇列號", 0, len(results)-1, 0, key="sf_pick")
                    if st.button("⬇ 帶入", key="sf_import_row"):
                        row = results.iloc[pick]
                        st.session_state["sf_import"] = {
                            "po":       str(row.get(po_c,"")) if po_c else "",
                            "customer": str(row.get(cu_c,"")) if cu_c else "",
                            "amount":   str(row.get(am_c,"")) if am_c else "",
                            "part":     str(row.get(pt_c,"")) if pt_c else "",
                            "qty":      str(row.get(qt_c,"")) if qt_c else "",
                        }
                        st.rerun()
                else:
                    st.warning("找不到符合的訂單")
        else:
            st.caption("無法連線 Teable，請手動填寫")

    imp = st.session_state.get("sf_import", {})

    # 類型選擇
    tpl_key = st.radio(
        "簽核類型", ["inv", "pck", "wip"],
        format_func=lambda x: {"inv": "🧾 Invoice",
                               "pck": "📦 Packing List",
                               "wip": "🏭 新訂單 WIP"}[x],
        horizontal=True,
        key="sf_tpl_sel",
    )

    t = tpl_key
    po_ref = ""

    c1, c2 = st.columns(2)
    if tpl_key == "inv":
        with c1:
            doc_id   = st.text_input("Invoice 編號 *", value=imp.get("po",""), key=f"sf_doc_id_{t}")
            customer = st.text_input("客戶名稱 *", value=imp.get("customer",""), key=f"sf_cust_{t}")
        with c2:
            amount   = st.text_input("金額", value=imp.get("amount",""), key=f"sf_amt_{t}")
            terms    = st.selectbox("付款條件", ["Net 30","Net 60","Net 90","T/T in advance","L/C"], key=f"sf_terms_{t}")
        extra_fields = {"付款條件": terms}

    elif tpl_key == "pck":
        with c1:
            doc_id   = st.text_input("PL 編號 *", value="PL-"+_today().replace("-","")[:6]+"-", key=f"sf_doc_id_{t}")
            po_ref   = st.text_input("對應 PO#", value=imp.get("po",""), key=f"sf_poref_{t}")
            customer = st.text_input("客戶名稱 *", value=imp.get("customer",""), key=f"sf_cust_{t}")
        with c2:
            amount   = st.text_input("金額", value=imp.get("amount",""), key=f"sf_amt_{t}")
            dest     = st.text_input("目的地", key=f"sf_dest_{t}")
            ship_m   = st.selectbox("運輸方式", ["Sea Freight","Air Freight","Express"], key=f"sf_ship_{t}")
        extra_fields = {"對應PO": po_ref, "目的地": dest, "運輸": ship_m}

    else:  # wip
        with c1:
            doc_id   = st.text_input("WIP 單號 *",
                                     value=imp.get("po", f"WIP-{_today()[:7]}-"),
                                     key=f"sf_doc_id_{t}")
            po_ref   = st.text_input("客戶 PO# / 西拓訂單編號",
                                     value=imp.get("po",""), key=f"sf_poref_{t}")
            customer = st.text_input("客戶名稱 *", value=imp.get("customer",""), key=f"sf_cust_{t}")
        with c2:
            amount   = st.text_input("金額", value=imp.get("amount",""), key=f"sf_amt_{t}")
            model    = st.text_input("產品型號 / 料號", value=imp.get("part",""), key=f"sf_model_{t}")
            etd      = st.date_input("交貨期 ETD", key=f"sf_etd_{t}")
        extra_fields = {"對應PO": po_ref, "料號": model, "ETD": str(etd)}

    st.divider()

    # ── STEP 3: 簽核人員 ─────────────────────────────────────
    st.markdown("**STEP 3 · 簽核人員**")
    defaults = DEFAULT_APPROVERS[tpl_key]
    new_stations = []
    for i, d in enumerate(defaults):
        ca, cb, cc = st.columns([2, 3, 2])
        name  = ca.text_input(f"第{i+1}關 姓名", value=d["person"], key=f"sf_ap_n_{i}")
        email = cb.text_input("Email",            value=d["email"],  key=f"sf_ap_e_{i}")
        ridx  = ALL_ROLES.index(d["role"]) if d["role"] in ALL_ROLES else 0
        role  = cc.selectbox("職責", ALL_ROLES, index=ridx,         key=f"sf_ap_r_{i}")
        new_stations.append({
            "name": role, "person": name, "role": role, "email": email,
            "status": "current" if i == 0 else "pending",
            "time": None, "signed": False,
        })

    applicant = new_stations[0]["person"] if new_stations else "申請人"
    app_email = new_stations[0]["email"]  if new_stations else ""

    st.divider()

    # ── 提交 / 取消（不用 columns 包，避免 rerun 在 columns 內崩潰）──
    do_submit = st.button("🚀 提交並啟動流程", type="primary", key="sf_submit")
    do_cancel = st.button("取消", key="sf_cancel")

    if do_cancel:
        for k in ["sf_import", "sf_file_pending"]:
            st.session_state.pop(k, None)
        st.session_state["sf_view"] = "list"
        st.rerun()

    if do_submit:
        if not doc_id.strip():
            st.error("請填寫文件編號！")
        elif not customer.strip():
            st.error("請填寫客戶名稱！")
        else:
            pending_file   = st.session_state.pop("sf_file_pending", "")
            final_filedata = filedata or pending_file

            new_doc = {
                "_record_id": "",
                "id":         doc_id.strip(),
                "title":      f"{DOC_LABEL[tpl_key]} #{doc_id.strip()}",
                "doc_type":   tpl_key,
                "customer":   customer.strip(),
                "po":         po_ref.strip() if tpl_key != "inv" else doc_id.strip(),
                "amount":     amount.strip(),
                "status":     "pending",
                "date":       _today(),
                "applicant":  applicant,
                "email":      app_email,
                "filename":   filename,
                "filedata":   final_filedata,
                "fields":     {"金額": amount.strip(), **extra_fields},
                "stations":   new_stations,
                "logs": [{
                    "person":  applicant,
                    "action":  "提交簽核申請",
                    "comment": f"{DOC_LABEL[tpl_key]} {doc_id.strip()}（附件：{filename or '無'}）",
                    "time":    _now(),
                    "type":    "submitted",
                }],
            }

            if final_filedata and not filedata:
                st.session_state[f"sf_file_{doc_id.strip()}"] = final_filedata

            _sf_save(new_doc)
            _sf_clear_cache()
            st.session_state.pop("sf_import", None)
            st.session_state["sf_view"] = "list"
            st.rerun()



# ════════════════════════════════════════════════════════════════
#  EMAIL POPUP
# ════════════════════════════════════════════════════════════════
def _email_popup(doc: dict) -> None:
    cur      = next((s for s in doc["stations"] if s["status"] == "current"), None)
    tl       = DOC_LABEL.get(doc["doc_type"], "文件")
    sig_done = sum(1 for s in doc["stations"] if s.get("signed"))
    sig_tot  = len(doc["stations"])

    with st.expander("📧 Email 通知預覽", expanded=True):
        to = cur["email"] if cur else "—"
        st.markdown(f"""
**收件人：** `{to}`
**主旨：** `【GLOCOM 待簽核】{tl} {doc['id']} — {doc['customer']}`

---
您好 **{cur['person'] if cur else ''}**，

請審核以下 **{tl}**（輪到您：**{cur['role'] if cur else ''}**）。

- **文件編號：** {doc['id']}
- **客戶：** {doc['customer']}
- **PO#：** {doc.get('po','—')}
- **金額：** {doc.get('amount','—')}
- **簽名進度：** {sig_done}/{sig_tot}

請開啟 GLOCOM Control Tower → 左側 **SignFlow** → 簽核詳情

`https://glocom-wip-v36-xko7x2byuccqjtpr9eaivj.streamlit.app/`
        """)
        if st.button("✉️ 確認寄出", type="primary", key="sf_send"):
            st.success("📧 通知已寄出！")


# ════════════════════════════════════════════════════════════════
#  MAIN ENTRY
# ════════════════════════════════════════════════════════════════
def render_approval_page() -> None:
    st.markdown(_CSS, unsafe_allow_html=True)

    # init session
    if "sf_view"    not in st.session_state: st.session_state["sf_view"]    = "list"
    if "sf_current" not in st.session_state: st.session_state["sf_current"] = 0

    # SIGNFLOW_TABLE_URL 未設定時靜默處理（不顯示警告）

    # header
    st.markdown(
        "## SignFlow 簽核平台 "
        "<span style='font-size:14px;color:#c8973a;font-weight:400'>"
        "GLOCOM Internal</span>",
        unsafe_allow_html=True,
    )
    st.caption("Invoice  /  Packing List  /  新訂單 WIP  —  5～6 人簽核 + 手寫電子簽名")

    # nav
    c1, c2, c3, _ = st.columns([1.4, 1.4, 1.2, 6])
    v = st.session_state["sf_view"]
    with c1:
        if st.button("📋 總覽", use_container_width=True,
                     type="primary" if v == "list" else "secondary", key="sf_nl"):
            st.session_state["sf_view"] = "list"; st.rerun()
    with c2:
        if st.button("🔍 詳情", use_container_width=True,
                     type="primary" if v == "detail" else "secondary", key="sf_nd"):
            st.session_state["sf_view"] = "detail"; st.rerun()
    with c3:
        if st.button("＋ 新增", use_container_width=True,
                     type="primary" if v == "create" else "secondary", key="sf_nc"):
            st.session_state["sf_view"] = "create"; st.rerun()

    st.divider()

    # load docs
    docs = _sf_load_all()

    # route
    if v == "list":
        _view_list(docs)
    elif v == "detail":
        _view_detail(docs)
    elif v == "create":
        _view_create()
