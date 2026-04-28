# -*- coding: utf-8 -*-
"""
spec_history_writer.py - 從 Streamlit 即時更新 spec_history.json 到 GitHub

【何時用】
Sandy 在 Streamlit 寫一張新的工廠 PO,寫進 Teable 主表後,
順便把這張單的料號 + 規格寫進 spec_history.json,直接 commit + push 到 GitHub。

【原理】
透過 GitHub REST API 的 Contents endpoint:
  GET  /repos/{owner}/{repo}/contents/{path}        (拿目前的 sha + 內容)
  PUT  /repos/{owner}/{repo}/contents/{path}        (用新內容 + sha 寫回)

【認證】
從 Streamlit Secrets 讀 GitHub PAT。
secrets.toml 範例:
    [github]
    token = "ghp_xxxxxxxxxxxxxxxxxxxxxx"
    owner = "jonah58w"
    repo = "glocom-wip-v36"
    branch = "main"
    spec_history_path = "data/spec_history.json"

【併發處理】
GitHub API 用 sha 作 optimistic locking:
- 拿檔時 GitHub 回 sha
- 寫回時帶這個 sha,如果 sha 變了(別人剛寫過)→ 409 衝突 → 重 retry
- 我們重試 3 次,過了還是衝突就放棄(讓 Sandy 知道,通常等 1 分鐘就好)
"""

from __future__ import annotations

import base64
import json
import time
from datetime import datetime
from typing import Optional

import requests
import streamlit as st


# ─── Secrets 載入 ────────────────────────────
def _get_github_config() -> Optional[dict]:
    """從 st.secrets 讀 GitHub 設定"""
    try:
        gh = st.secrets["github"]
        return {
            "token": gh["token"],
            "owner": gh.get("owner", "jonah58w"),
            "repo": gh.get("repo", "glocom-wip-v36"),
            "branch": gh.get("branch", "main"),
            "spec_history_path": gh.get("spec_history_path", "data/spec_history.json"),
        }
    except (KeyError, AttributeError, FileNotFoundError):
        return None


def is_github_writer_available() -> bool:
    """檢查 GitHub Writer 是否可用(secrets 有設好)"""
    return _get_github_config() is not None


# ─── GitHub API 操作 ──────────────────────────
def _api_url(cfg: dict) -> str:
    return f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}/contents/{cfg['spec_history_path']}"


def _get_current_file(cfg: dict) -> Optional[tuple[dict, str]]:
    """
    從 GitHub 抓現在的 spec_history.json
    回傳 (data_dict, sha) 或 None
    """
    headers = {
        "Authorization": f"Bearer {cfg['token']}",
        "Accept": "application/vnd.github+json",
    }
    params = {"ref": cfg["branch"]}
    
    try:
        r = requests.get(_api_url(cfg), headers=headers, params=params, timeout=15)
        if r.status_code != 200:
            return None
        body = r.json()
        sha = body.get("sha", "")
        content_b64 = body.get("content", "")
        if not content_b64:
            return None
        content = base64.b64decode(content_b64).decode("utf-8")
        data = json.loads(content)
        return data, sha
    except Exception:
        return None


def _put_file(cfg: dict, data: dict, sha: str, message: str) -> tuple[bool, str]:
    """
    把新內容寫回 GitHub
    回傳 (success, message)
    """
    headers = {
        "Authorization": f"Bearer {cfg['token']}",
        "Accept": "application/vnd.github+json",
    }
    new_content = json.dumps(data, ensure_ascii=False, indent=2)
    new_content_b64 = base64.b64encode(new_content.encode("utf-8")).decode("ascii")
    
    payload = {
        "message": message,
        "content": new_content_b64,
        "sha": sha,
        "branch": cfg["branch"],
    }
    
    try:
        r = requests.put(_api_url(cfg), headers=headers, json=payload, timeout=20)
        if r.status_code in (200, 201):
            return True, "OK"
        elif r.status_code == 409:
            return False, "CONFLICT"  # sha 不一致(有人剛寫過),retry
        else:
            return False, f"HTTP {r.status_code}: {r.text[:300]}"
    except Exception as e:
        return False, f"Exception: {e}"


# ─── 主入口:append 一筆規格 ──────────────────
def append_spec_history(
    part_number: str,
    spec_text: str,
    po_no: str,
    factory: str,
    date_str: str,
    max_retries: int = 3,
) -> tuple[bool, str]:
    """
    把一筆 (part_number, spec_text, po_no, factory, date) append 到 GitHub 上的 spec_history.json
    
    Returns:
        (success: bool, message: str)
    """
    cfg = _get_github_config()
    if cfg is None:
        return False, "GitHub Writer 未設定(.streamlit/secrets.toml 缺 [github] 區段)"
    
    pn = (part_number or "").strip()
    spec = (spec_text or "").strip()
    po = (po_no or "").strip()
    
    if not pn or not spec or not po:
        return False, "料號 / 規格 / PO 號不能為空"
    
    new_entry = {
        "po": po,
        "date": date_str or datetime.now().strftime("%Y-%m-%d"),
        "factory": factory or "",
        "spec_text": spec,
    }
    
    for attempt in range(max_retries):
        # 1. 拿現有的
        result = _get_current_file(cfg)
        if result is None:
            return False, "無法從 GitHub 抓 spec_history.json"
        data, sha = result
        
        spec_history = data.get("spec_history", {})
        
        # 2. 合併
        if pn not in spec_history:
            spec_history[pn] = {
                "latest": new_entry.copy(),
                "history": [new_entry.copy()],
            }
            action = "新增料號"
        else:
            existing = spec_history[pn]
            # 不重複加同 PO
            if any(h.get("po") == po for h in existing.get("history", [])):
                return True, f"料號 {pn} / PO {po} 已存在,不重複加"
            existing["history"].append(new_entry.copy())
            # 更新 latest 如果這筆比較新
            if new_entry["date"] > existing.get("latest", {}).get("date", ""):
                existing["latest"] = new_entry.copy()
            action = "加歷史"
        
        # 3. 更新 meta
        if "_meta" not in data:
            data["_meta"] = {}
        data["_meta"]["unique_part_numbers"] = len(spec_history)
        data["_meta"]["last_streamlit_update"] = datetime.now().isoformat()
        
        # 4. 寫回
        commit_msg = f"Streamlit auto-update: {action} {pn} from {po} ({factory})"
        ok, msg = _put_file(cfg, data, sha, commit_msg)
        
        if ok:
            return True, f"{action}: {pn} / PO {po}"
        
        if msg == "CONFLICT":
            # sha 不一致(別人剛寫了),等一下再重試
            time.sleep(1 + attempt * 0.5)
            continue
        
        # 其他錯誤直接放棄
        return False, msg
    
    return False, f"重試 {max_retries} 次都失敗(GitHub 競爭寫入)"


def append_multiple_spec_history(
    items: list[dict],
) -> tuple[int, int, list[str]]:
    """
    批次寫多筆 (整張訂單會有多個品項)
    
    Args:
        items: list of {
            "part_number": str,
            "spec_text": str,
            "po_no": str,
            "factory": str,
            "date_str": str,
        }
    
    Returns:
        (success_count, fail_count, [messages])
    """
    cfg = _get_github_config()
    if cfg is None:
        return 0, len(items), ["GitHub Writer 未設定"]
    
    success = 0
    fail = 0
    messages = []
    
    # 為了減少 API 呼叫(每次寫都要先 get 再 put),
    # 一次撈 spec_history,本機合併完所有品項,再一次寫回
    
    for attempt in range(3):
        result = _get_current_file(cfg)
        if result is None:
            return 0, len(items), ["無法從 GitHub 抓 spec_history.json"]
        data, sha = result
        
        spec_history = data.get("spec_history", {})
        
        # 合併所有 items
        per_item_msgs = []
        any_change = False
        
        for item in items:
            pn = (item.get("part_number") or "").strip()
            spec = (item.get("spec_text") or "").strip()
            po = (item.get("po_no") or "").strip()
            factory = item.get("factory") or ""
            date_str = item.get("date_str") or datetime.now().strftime("%Y-%m-%d")
            
            if not pn or not spec or not po:
                fail += 1
                per_item_msgs.append(f"❌ 跳過 {pn or '(空)'}: 缺少欄位")
                continue
            
            new_entry = {
                "po": po, "date": date_str,
                "factory": factory, "spec_text": spec,
            }
            
            if pn not in spec_history:
                spec_history[pn] = {
                    "latest": new_entry.copy(),
                    "history": [new_entry.copy()],
                }
                per_item_msgs.append(f"✅ 新增料號: {pn}")
                any_change = True
                success += 1
            else:
                existing = spec_history[pn]
                if any(h.get("po") == po for h in existing.get("history", [])):
                    per_item_msgs.append(f"⏭️ {pn} / PO {po} 已存在")
                    success += 1  # 視為已處理(不算失敗)
                    continue
                existing["history"].append(new_entry.copy())
                if new_entry["date"] > existing.get("latest", {}).get("date", ""):
                    existing["latest"] = new_entry.copy()
                per_item_msgs.append(f"✅ 加歷史: {pn}")
                any_change = True
                success += 1
        
        if not any_change:
            return success, fail, per_item_msgs
        
        # 更新 meta
        if "_meta" not in data:
            data["_meta"] = {}
        data["_meta"]["unique_part_numbers"] = len(spec_history)
        data["_meta"]["last_streamlit_update"] = datetime.now().isoformat()
        
        # 寫回
        po_list = list({(item.get("po_no") or "").strip()
                       for item in items if item.get("po_no")})
        commit_msg = (
            f"Streamlit auto-update: {len(items)} items from "
            f"{', '.join(po_list[:3])}{'...' if len(po_list) > 3 else ''}"
        )
        ok, msg = _put_file(cfg, data, sha, commit_msg)
        
        if ok:
            messages.extend(per_item_msgs)
            messages.append(f"✅ GitHub commit OK: {commit_msg}")
            return success, fail, messages
        
        if msg == "CONFLICT":
            time.sleep(1 + attempt * 0.5)
            continue
        
        # 其他錯誤
        messages.append(f"❌ GitHub 寫入失敗: {msg}")
        return 0, len(items), messages
    
    messages.append(f"❌ 重試 3 次都失敗(GitHub 競爭寫入)")
    return 0, len(items), messages
