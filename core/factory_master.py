# -*- coding: utf-8 -*-
"""
工廠主檔載入模組。

設計原則:
- 跟 GLOCOM Control Tower 的 safe_text / normalize 風格一致
- 短名 (short) 是 Teable 主表「工廠」欄位的值,當查表 key
- _TODO_* 條目和 is_active=False 的工廠不出現在下拉選單
"""

import json
from pathlib import Path
from typing import Optional

# 預設讀同層 data/factories.json,可由呼叫者覆寫
_DEFAULT_PATH = Path(__file__).parent.parent / "data" / "factories.json"


def load_factories(path: Optional[Path] = None, active_only: bool = True) -> dict:
    """載入工廠主檔。

    Returns:
        {short_name: factory_dict}
        例如 {"宏棋": {"factory_name": "宏棋科技有限公司", ...}, ...}
    """
    target = Path(path) if path else _DEFAULT_PATH
    with open(target, encoding="utf-8") as f:
        raw = json.load(f)
    factories = raw.get("factories", {})

    if not active_only:
        return factories

    return {
        k: v for k, v in factories.items()
        if not k.startswith("_") and v.get("is_active", True)
    }


def get_factory(short_name: str, path: Optional[Path] = None) -> Optional[dict]:
    """從短名取單一工廠資料。找不到回 None。

    Args:
        short_name: Teable 主表「工廠」欄位的值,例如「宏棋」
    """
    if not short_name:
        return None
    return load_factories(path=path, active_only=False).get(str(short_name).strip())


def list_factory_options(path: Optional[Path] = None) -> list[tuple[str, str]]:
    """給 Streamlit 下拉選單用的清單。

    Returns:
        [(short_name, display_label), ...]
        例如 [("宏棋", "宏棋 - 宏棋科技有限公司 (Taiwan)"), ...]
    """
    factories = load_factories(path=path, active_only=True)
    options = []
    for short, f in factories.items():
        full_name = f.get("factory_name", "")
        region = f.get("region", "")
        # 已補完整資料的工廠 vs 空殼
        is_stub = "[請補]" in full_name or "[請補]" in f.get("address", "")
        marker = " ⚠️" if is_stub else ""
        label = f"{short} - {full_name} ({region}){marker}"
        options.append((short, label))
    return sorted(options, key=lambda x: x[0])


def has_complete_data(factory: dict) -> tuple[bool, list[str]]:
    """檢查工廠資料是否完整(用於 UI 警告)。

    Returns:
        (is_complete, missing_fields)
    """
    if not factory:
        return False, ["entire factory record"]

    required = ["factory_name", "address", "phone", "contact_person"]
    missing = []
    for field in required:
        v = factory.get(field, "")
        if not v or "[請補]" in str(v):
            missing.append(field)
    return len(missing) == 0, missing
