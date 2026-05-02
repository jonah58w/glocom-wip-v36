# -*- coding: utf-8 -*-
"""
legacy_history.py — 把 113~114 年舊訂單 (RTF/DOCX) 解析後的歷史
                    無縫合併進現有的 spec_history。

設計目標:
- legacy_orders.json 的結構跟 spec_history.json 完全一致 (_meta + P/N 平放)
- 兩邊紀錄的欄位 (po_no/factory/date/spec_text/...) 也對齊
- merge 時用 (po_no, factory) 當 key 去重,避免同一筆被算兩次

用法 (在 factory_po_create_page.py 裡):

    from legacy_history import merge_legacy_into_spec_history

    # 原本的 spec_history 載入
    spec_history_data = load_spec_history()

    # 加這一行就能合併歷史 (失敗時自動 fallback,不影響原邏輯)
    spec_history_data = merge_legacy_into_spec_history(spec_history_data)

    # 後續所有 _compute_old_pn_factory_suggestions / fetch_previous_spec
    # 都不用改,自動讀到完整歷史
"""
from __future__ import annotations
import json
from pathlib import Path

DEFAULT_LEGACY_PATH = Path(__file__).parent / "data" / "legacy_orders.json"


def load_legacy_history(path: Path | str | None = None) -> dict:
    """讀 legacy_orders.json。失敗回傳 {}."""
    p = Path(path) if path else DEFAULT_LEGACY_PATH
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def merge_legacy_into_spec_history(
    spec_history: dict,
    legacy_path: Path | str | None = None,
) -> dict:
    """把 legacy_orders.json 的紀錄合併進 spec_history dict。

    - spec_history 的紀錄保持優先 (新的覆蓋舊的)
    - (po_no, factory) 為 dedup key
    - 失敗時回傳原 spec_history (永不破壞既有資料)
    """
    if not isinstance(spec_history, dict):
        return spec_history

    legacy = load_legacy_history(legacy_path)
    if not legacy:
        return spec_history

    # 複製 spec_history 避免改到原 dict
    merged = dict(spec_history)
    legacy_count = 0

    for pn, recs in legacy.items():
        if pn.startswith("_"):  # _meta 等元資料跳過
            continue
        if not isinstance(recs, list):
            continue

        existing = merged.get(pn, [])
        if not isinstance(existing, list):
            existing = [existing] if isinstance(existing, dict) else []

        # 用 (po_no, factory) 去重
        existing_keys = set()
        for r in existing:
            if isinstance(r, dict):
                existing_keys.add((r.get("po_no", ""), r.get("factory", "")))

        # 只 append legacy 中不存在的
        for r in recs:
            if not isinstance(r, dict):
                continue
            key = (r.get("po_no", ""), r.get("factory", ""))
            if key not in existing_keys:
                existing.append(r)
                existing_keys.add(key)
                legacy_count += 1

        # 按日期降冪重排
        existing.sort(key=lambda x: x.get("date", "") if isinstance(x, dict) else "", reverse=True)
        merged[pn] = existing

    # 加個 _meta 標記合併了多少筆
    if "_meta" not in merged or not isinstance(merged.get("_meta"), dict):
        merged["_meta"] = {}
    merged["_meta"]["legacy_records_merged"] = legacy_count

    return merged


def get_legacy_stats() -> dict:
    """讀 legacy_orders.json _meta,給診斷面板顯示用。"""
    legacy = load_legacy_history()
    return legacy.get("_meta", {}) if isinstance(legacy, dict) else {}


if __name__ == "__main__":
    # CLI 自我測試
    legacy = load_legacy_history()
    if not legacy:
        print("⚠️  data/legacy_orders.json 不存在或無法解析")
    else:
        meta = legacy.get("_meta", {})
        pns = [k for k in legacy if not k.startswith("_")]
        print(f"✅ legacy_orders.json 載入成功")
        print(f"   料號數: {len(pns)}")
        print(f"   工廠分布: {meta.get('factory_distribution', {})}")
        print(f"   字首分布: {meta.get('prefix_distribution', {})}")

        # 測 merge
        fake_spec_history = {
            "TEST-PN-001": [{"po_no": "ET1150001-01", "factory": "全興", "date": "2026-01-01"}],
        }
        merged = merge_legacy_into_spec_history(fake_spec_history)
        merged_pns = [k for k in merged if not k.startswith("_")]
        print(f"\n   merge 測試:")
        print(f"   原 spec_history: 1 個 P/N")
        print(f"   merge 後: {len(merged_pns)} 個 P/N")
        print(f"   合併紀錄數 (_meta): {merged['_meta'].get('legacy_records_merged')}")
