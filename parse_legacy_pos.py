# -*- coding: utf-8 -*-
"""
parse_legacy_pos.py — 解析 113~114 年舊訂單 RTF / DOCX 為 legacy_orders.json

支援格式:
- RTF (.rtf) — 用 striprtf 轉純文字
- DOCX (.docx) — 用 python-docx 取出每個 row 的 unique cell,組成跟 RTF 一致的格式

用法:
    把所有舊 PO 檔放進 legacy_pos/ 目錄,然後:
    python parse_legacy_pos.py

    輸出: data/legacy_orders.json

擴充: 之後有新檔案丟進 legacy_pos/ 重跑即可。
"""
from __future__ import annotations
import json
import re
import sys
from pathlib import Path

from striprtf.striprtf import rtf_to_text  # type: ignore
from docx import Document  # type: ignore


# 廠商全名 → 短名 (跟 spec_history.json 一致)
FACTORY_NAME_MAP = {
    "全興電子有限公司": "全興",
    "優技電子股份有限公司": "優技",
    "優技電子工業股份有限公司": "優技",
    "宏棋科技有限公司": "宏棋",
    "宏棋電子有限公司": "宏棋",
    "星晨電路股份有限公司": "星晨",
    "柏承科技股份有限公司": "柏承",
    "百為實業有限公司": "百為",
    "龍偉電子股份有限公司": "龍偉",
    "雙銘電子股份有限公司": "雙銘",
    "雙銘": "雙銘",
}


def normalize_factory_name(full_name: str) -> str:
    full_name = (full_name or "").strip()
    if full_name in FACTORY_NAME_MAP:
        return FACTORY_NAME_MAP[full_name]
    # 嘗試前綴比對
    for full, short in FACTORY_NAME_MAP.items():
        if full_name.startswith(short):
            return short
    return full_name


def parse_date(date_str: str) -> str:
    """'AUG. 19, 2024' / 'OCT.23, 2024' → '2024-08-19'。失敗回傳原字串。"""
    s = (date_str or "").strip().upper()
    # 把 . 和 , 換成空格,再壓多重空白
    s = re.sub(r"[.,]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    parts = s.split()
    if len(parts) != 3:
        return date_str.strip()
    mon_map = {
        "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04",
        "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
        "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12",
    }
    try:
        mon = mon_map[parts[0][:3]]
        day = parts[1].zfill(2)
        year = parts[2]
        return f"{year}-{mon}-{day}"
    except Exception:
        return date_str.strip()


def get_po_prefix(po_no: str) -> str:
    """ET1130091-01 → ET, G1130045-01 → G, EW1130117-01 → EW。"""
    m = re.match(r"^([A-Z]+)", po_no.strip())
    return m.group(1) if m else ""


def normalize_spec_text(spec: str) -> str:
    """\\t → 空格、合併多重空白、保留換行。"""
    if not spec:
        return ""
    spec = spec.replace("\t", " ")
    lines = []
    for line in spec.split("\n"):
        line = re.sub(r"[ \u3000]+", " ", line).strip()
        if line:
            lines.append(line)
    return "\n".join(lines)


def docx_to_pipe_text(path: Path) -> str:
    """DOCX → '|val1|val2|val3|' 行的文字 (跟 RTF 同形式,方便共用 regex)。"""
    doc = Document(str(path))
    out_lines: list[str] = []
    for tbl in doc.tables:
        for row in tbl.rows:
            seen: set[str] = set()
            uniques: list[str] = []
            for cell in row.cells:
                t = cell.text.strip()
                if t and t not in seen:
                    seen.add(t)
                    uniques.append(t)
            if uniques:
                out_lines.append("|" + "|".join(uniques) + "|")
    return "\n".join(out_lines)


def load_text(path: Path) -> str:
    if path.suffix.lower() == ".rtf":
        raw = path.read_text(encoding="utf-8", errors="ignore")
        return rtf_to_text(raw, errors="ignore")
    elif path.suffix.lower() == ".docx":
        return docx_to_pipe_text(path)
    else:
        raise ValueError(f"Unsupported: {path.suffix}")


# ─── 解析品項 ────────────────────────────────────────────────────────────
# 表頭: |產品編號|產品規格|交期|數量|單價|小計|
# 品項行: |<P/N>|<spec...>|<交期>|<數量>|<單價>|<小計>|
# 有些 spec 含換行,所以要逐行掃,尾段是「合計|NT$xxx」

ITEM_HEADER_RE = re.compile(r"產品編號\s*\|\s*產品規格\s*\|\s*交期[\s\n]*\|?\s*數量[\s\n]*\|?\s*單價\s*\|\s*小計", re.MULTILINE)
TOTAL_RE = re.compile(r"^\s*\|?\s*合\s*計\s*\|")
QTY_LINE_RE = re.compile(r"\d[\d,]*\s*(?:pcs|PCS|片|panels?|PNL)", re.IGNORECASE)
PRICE_RE = re.compile(r"\$?[\d,]+\.\d+")


def parse_items_section(section_text: str) -> list[dict]:
    """Stream-based: 整個 section 按 '|' 切,跟行界無關。

    結構: [P/N | spec | delivery | qty | price | subtotal] 重複。
    delivery 認 ', 20XX' (英文月份+年) 當 anchor。
    """
    raw = [c.strip() for c in section_text.split("|")]
    cells = [c for c in raw if c]

    items: list[dict] = []
    i = 0
    while i + 5 < len(cells):
        # 期待 cells[i+2] = delivery 含 ', 20XX'
        if re.search(r",\s*20\d\d", cells[i + 2]):
            items.append({
                "part_number": cells[i],
                "spec": cells[i + 1],
                "delivery": cells[i + 2],
                "quantity": cells[i + 3],
                "unit_price": cells[i + 4],
                "subtotal": cells[i + 5],
            })
            i += 6
        else:
            i += 1

    return items


def clean_pn(pn: str) -> str:
    """有時 P/N 跨兩行 (e.g. '9004 PCB Rev D\\n(PCBA 9004 Rev D)'),只取第一行。"""
    pn = pn.replace("\r", "").strip()
    if "\n" in pn:
        pn = pn.split("\n")[0].strip()
    return pn


def parse_one_file(path: Path) -> dict | None:
    """回傳 {po_no, factory, factory_full, prefix, date, items: [...]}。"""
    text = load_text(path)

    # 廠商
    m_factory = re.search(r"廠商名稱\s*:\s*\|([^|]+)\|", text)
    factory_full = m_factory.group(1).strip() if m_factory else ""
    factory_short = normalize_factory_name(factory_full)

    # 採購單號
    m_po = re.search(r"採購單號\s*:\s*\|([^|]+)\|", text)
    po_no = m_po.group(1).strip() if m_po else ""

    # 採購日期
    m_date = re.search(r"採購日期\s*:\s*\|([^|]+)\|", text)
    date = parse_date(m_date.group(1).strip() if m_date else "")

    if not po_no:
        return None

    # 抓品項區塊
    m_header = ITEM_HEADER_RE.search(text)
    if not m_header:
        return None
    after = text[m_header.end():]
    # 切到「合計」之前
    m_total = re.search(r"\n[^\n]*合\s*計\s*\|", after)
    section = after[:m_total.start()] if m_total else after

    items_raw = parse_items_section(section)
    items = []
    for it in items_raw:
        items.append({
            "part_number": clean_pn(it["part_number"]),
            "spec": normalize_spec_text(it["spec"]),
            "delivery": it["delivery"],
            "quantity": it["quantity"],
            "unit_price": it["unit_price"].replace("$", "").replace(",", ""),
            "subtotal": it["subtotal"].replace("$", "").replace(",", ""),
        })

    return {
        "po_no": po_no,
        "prefix": get_po_prefix(po_no),
        "factory": factory_short,
        "factory_full": factory_full,
        "date": date,
        "items": items,
        "source_file": path.name,
    }


# ─── Main ─────────────────────────────────────────────────────────────────


def main():
    legacy_dir = Path(__file__).parent / "legacy_pos"
    out_path = Path(__file__).parent / "data" / "legacy_orders.json"
    out_path.parent.mkdir(exist_ok=True)

    files = sorted(legacy_dir.glob("*.rtf")) + sorted(legacy_dir.glob("*.docx"))
    print(f"📂 在 {legacy_dir} 找到 {len(files)} 個檔案")

    pn_history: dict[str, list[dict]] = {}  # P/N → [records...]
    errors: list[tuple[str, str]] = []
    parsed_count = 0

    skipped_ew: list[str] = []
    for f in files:
        # EW 字首是 EUSWAY → 海外客戶的銷貨單,英文表頭,不是工廠 PO,跳過
        if f.name.startswith("EW"):
            skipped_ew.append(f.name)
            continue
        try:
            result = parse_one_file(f)
            if result is None:
                errors.append((f.name, "parse_one_file returned None"))
                continue
            if not result["items"]:
                errors.append((f.name, "no items parsed"))
                continue
            parsed_count += 1
            for it in result["items"]:
                pn = it["part_number"]
                if not pn:
                    continue
                rec = {
                    "po_no": result["po_no"],
                    "prefix": result["prefix"],
                    "factory": result["factory"],
                    "factory_full": result["factory_full"],
                    "date": result["date"],
                    "spec_text": it["spec"],
                    "quantity": it["quantity"],
                    "unit_price": it["unit_price"],
                    "subtotal": it["subtotal"],
                    "delivery": it["delivery"],
                    "source_file": result["source_file"],
                }
                pn_history.setdefault(pn, []).append(rec)
        except Exception as e:
            errors.append((f.name, f"{type(e).__name__}: {e}"))

    # 每個 P/N 的 records 按日期降冪排序
    for pn, recs in pn_history.items():
        recs.sort(key=lambda r: r["date"], reverse=True)

    # 統計
    total_records = sum(len(v) for v in pn_history.values())
    factory_counts: dict[str, int] = {}
    prefix_counts: dict[str, int] = {}
    for recs in pn_history.values():
        for r in recs:
            factory_counts[r["factory"]] = factory_counts.get(r["factory"], 0) + 1
            prefix_counts[r["prefix"]] = prefix_counts.get(r["prefix"], 0) + 1

    output = {
        "_meta": {
            "source": "113~114 年舊訂單 RTF/DOCX",
            "total_files": len(files),
            "parsed_files": parsed_count,
            "unique_part_numbers": len(pn_history),
            "total_records": total_records,
            "factory_distribution": factory_counts,
            "prefix_distribution": prefix_counts,
            "errors": errors,
            "skipped_ew_files": len(skipped_ew),
        },
    }
    # 將每個 P/N 平放在 root (跟 spec_history.json 一致),方便直接 merge
    output.update(pn_history)

    out_path.write_text(
        json.dumps(output, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"✅ 已輸出 {out_path}")
    print(f"   檔案總數: {len(files)}")
    print(f"   成功解析: {parsed_count}")
    print(f"   失敗: {len(errors)}")
    print(f"   不同料號: {len(pn_history)}")
    print(f"   總紀錄數: {total_records}")
    print(f"   工廠分布: {factory_counts}")
    print(f"   字首分布: {prefix_counts}")
    if skipped_ew:
        print(f"   跳過 EW 字首: {len(skipped_ew)} 筆 (英文版客戶銷貨單,非工廠 PO)")
    if errors:
        print(f"\n⚠️  解析失敗 ({len(errors)}):")
        for name, msg in errors[:20]:
            print(f"   - {name}: {msg}")
        if len(errors) > 20:
            print(f"   ... 還有 {len(errors)-20} 個")


if __name__ == "__main__":
    main()
