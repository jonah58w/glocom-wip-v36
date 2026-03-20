import re
import pandas as pd
from utils import safe_text, part_no_core, normalize_due_text


def normalize_match_text(v):
    s = safe_text(v).upper()
    s = s.replace("－", "-").replace("—", "-")
    s = re.sub(r"\s+", "", s)
    return s


def normalize_match_qty(v):
    s = safe_text(v).replace(",", "")
    if not s:
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def normalize_match_date(v):
    text = normalize_due_text(v)
    if not text:
        return None
    dt = pd.to_datetime(text, errors="coerce")
    if pd.isna(dt):
        return None
    return dt.strftime("%Y-%m-%d")


def build_record_for_match(po_value="", part_value="", qty_value="", due_value="", record_id="", source_row=None):
    return {
        "_record_id": record_id,
        "po": normalize_match_text(po_value),
        "part_no": normalize_match_text(part_no_core(part_value)),
        "qty": normalize_match_qty(qty_value),
        "factory_due_date": normalize_match_date(due_value),
        "_source_row": source_row,
    }


def match_score(imported_rec, existing_rec):
    score = 0
    matched_fields = []
    if imported_rec["po"] and existing_rec["po"] and imported_rec["po"] == existing_rec["po"]:
        score += 1
        matched_fields.append("PO#")
    if imported_rec["part_no"] and existing_rec["part_no"] and imported_rec["part_no"] == existing_rec["part_no"]:
        score += 1
        matched_fields.append("Part No")
    if imported_rec["qty"] is not None and existing_rec["qty"] is not None and imported_rec["qty"] == existing_rec["qty"]:
        score += 1
        matched_fields.append("Qty")
    if imported_rec["factory_due_date"] and existing_rec["factory_due_date"] and imported_rec["factory_due_date"] == existing_rec["factory_due_date"]:
        score += 1
        matched_fields.append("Factory Due Date")
    return score, matched_fields


def build_teable_match_records(current_df, teable_po_col, teable_part_col, teable_qty_col, teable_factory_due_col):
    results = []
    if current_df.empty:
        return results
    for _, row in current_df.iterrows():
        rec = build_record_for_match(
            po_value=row.get(teable_po_col, "") if teable_po_col else "",
            part_value=row.get(teable_part_col, "") if teable_part_col else "",
            qty_value=row.get(teable_qty_col, "") if teable_qty_col else "",
            due_value=row.get(teable_factory_due_col, "") if teable_factory_due_col else "",
            record_id=row.get("_record_id", ""),
            source_row=row,
        )
        results.append(rec)
    return results


def find_best_match_by_4fields(imported_rec, teable_match_records):
    scored = []
    for rec in teable_match_records:
        score, matched_fields = match_score(imported_rec, rec)
        if score > 0:
            scored.append({"record": rec, "score": score, "matched_fields": matched_fields})
    if not scored:
        return {"status": "manual_review", "target": None, "score": 0, "matched_fields": [], "candidates": []}
    scored.sort(key=lambda x: x["score"], reverse=True)
    best_score = scored[0]["score"]
    best_list = [x for x in scored if x["score"] == best_score]
    if best_score >= 3 and len(best_list) == 1:
        return {
            "status": "matched",
            "target": best_list[0]["record"],
            "score": best_score,
            "matched_fields": best_list[0]["matched_fields"],
            "candidates": best_list,
        }
    return {
        "status": "manual_review",
        "target": None,
        "score": best_score,
        "matched_fields": best_list[0]["matched_fields"] if best_list else [],
        "candidates": best_list,
    }


def dedupe_import_df_by_key(import_df, import_po_col, import_part_col, import_qty_col, import_factory_due_col):
    if import_df.empty:
        return import_df.copy(), []
    deduped_rows = []
    duplicate_keys = []
    seen = set()
    for _, row in import_df.iterrows():
        key = (
            normalize_match_text(row.get(import_po_col, "") if import_po_col else ""),
            normalize_match_text(part_no_core(row.get(import_part_col, "") if import_part_col else "")),
            normalize_match_qty(row.get(import_qty_col, "") if import_qty_col else ""),
            normalize_match_date(row.get(import_factory_due_col, "") if import_factory_due_col else ""),
        )
        usable_fields = sum([bool(key[0]), bool(key[1]), key[2] is not None, bool(key[3])])
        if usable_fields < 3:
            deduped_rows.append(row)
            continue
        if key in seen:
            duplicate_keys.append(key)
            continue
        seen.add(key)
        deduped_rows.append(row)
    deduped_df = pd.DataFrame(deduped_rows).reset_index(drop=True) if deduped_rows else import_df.iloc[0:0].copy()
    return deduped_df, duplicate_keys
