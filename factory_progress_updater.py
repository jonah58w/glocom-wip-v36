# -*- coding: utf-8 -*-
import re
import pandas as pd
import requests


def safe_text(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def parse_tags_from_text(text):
    if not text:
        return []
    return [x.strip() for x in str(text).split(",") if x.strip()]


def build_tags_value(tags, multi_select_mode=True):
    tags = [str(x).strip() for x in tags if str(x).strip()]
    if multi_select_mode:
        return tags
    return ", ".join(tags)


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


def parse_mmdd_range_to_date(text: str):
    s = safe_text(text)
    m = re.search(r"(\d{2})(\d{2})\s*(?:=>|->|~|-|至)\s*(\d{2})(\d{2})", s)
    if not m:
        return None
    try:
        year = pd.Timestamp.today().year
        mm = int(m.group(3))
        dd = int(m.group(4))
        return pd.Timestamp(year=year, month=mm, day=dd).strftime("%Y-%m-%d")
    except Exception:
        return None


def normalize_match_date(v):
    text = safe_text(v)
    if not text:
        return None

    special = parse_mmdd_range_to_date(text)
    if special:
        return special

    for candidate in [text.replace(".", "/"), text]:
        try:
            dt = pd.to_datetime(candidate, errors="coerce")
            if not pd.isna(dt):
                return dt.strftime("%Y-%m-%d")
        except Exception:
            pass

    m = re.search(r"\b(\d{1,2})/(\d{1,2})\b", text)
    if m:
        try:
            year = pd.Timestamp.today().year
            mm = int(m.group(1))
            dd = int(m.group(2))
            return pd.Timestamp(year=year, month=mm, day=dd).strftime("%Y-%m-%d")
        except Exception:
            pass

    return None


def normalize_part_no(v):
    s = safe_text(v).upper()
    if not s:
        return ""
    s = s.replace("－", "-").replace("—", "-")
    s = s.replace("\n", " ")
    s = re.sub(r"\([^)]*\)", " ", s)
    s = re.sub(r"\bREV(?:ISION)?\b[\s:_-]*[A-Z0-9._-]*", " ", s)
    s = re.sub(r"\bISS(?:UE)?\b[\s:_-]*[A-Z0-9._-]*", " ", s)
    s = re.sub(r"\bNEW\s+VERSION\b", " ", s)
    s = re.sub(r"\bNEW\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.replace(" ", "")


def normalize_wip_value(value: str, done_wip_values=None) -> str:
    if done_wip_values is None:
        done_wip_values = {"完成", "DONE", "COMPLETE", "COMPLETED", "FINISHED", "FINISH"}
    text = safe_text(value)
    if not text:
        return ""
    if text.upper() in done_wip_values or text in done_wip_values:
        return "完成"
    return text


def build_record_for_match(po_value="", part_value="", qty_value="", due_value="", record_id="", source_row=None):
    return {
        "_record_id": record_id,
        "po": normalize_match_text(po_value),
        "part_no": normalize_part_no(part_value),
        "qty": normalize_match_qty(qty_value),
        "factory_due_date": normalize_match_date(due_value),
        "_source_row": source_row,
    }


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
        score = 0
        matched_fields = []

        po_match = imported_rec["po"] and rec["po"] and imported_rec["po"] == rec["po"]
        part_match = imported_rec["part_no"] and rec["part_no"] and imported_rec["part_no"] == rec["part_no"]
        qty_match = imported_rec["qty"] is not None and rec["qty"] is not None and imported_rec["qty"] == rec["qty"]
        due_match = imported_rec["factory_due_date"] and rec["factory_due_date"] and imported_rec["factory_due_date"] == rec["factory_due_date"]

        if po_match:
            score += 100
            matched_fields.append("PO#")
        if part_match:
            score += 10
            matched_fields.append("Part No")
        if qty_match:
            score += 5
            matched_fields.append("Qty")
        if due_match:
            score += 2
            matched_fields.append("Factory Due Date")

        if score > 0:
            scored.append({"record": rec, "score": score, "matched_fields": matched_fields})

    if not scored:
        return {"status": "manual_review", "target": None, "score": 0, "matched_fields": [], "candidates": []}

    scored.sort(key=lambda x: x["score"], reverse=True)
    best_score = scored[0]["score"]
    best_list = [x for x in scored if x["score"] == best_score]

    if len(best_list) != 1:
        return {
            "status": "manual_review",
            "target": None,
            "score": best_score,
            "matched_fields": best_list[0]["matched_fields"] if best_list else [],
            "candidates": best_list,
        }

    best = best_list[0]
    mf = set(best["matched_fields"])

    if "PO#" in mf and ("Part No" in mf or "Qty" in mf):
        return {
            "status": "matched",
            "target": best["record"],
            "score": best["score"],
            "matched_fields": best["matched_fields"],
            "candidates": best_list,
        }

    if {"Part No", "Qty", "Factory Due Date"}.issubset(mf):
        return {
            "status": "matched",
            "target": best["record"],
            "score": best["score"],
            "matched_fields": best["matched_fields"],
            "candidates": best_list,
        }

    return {
        "status": "manual_review",
        "target": None,
        "score": best_score,
        "matched_fields": best["matched_fields"],
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
            normalize_part_no(row.get(import_part_col, "") if import_part_col else ""),
            normalize_match_qty(row.get(import_qty_col, "") if import_qty_col else ""),
            normalize_match_date(row.get(import_factory_due_col, "") if import_factory_due_col else ""),
        )

        usable_fields = sum([bool(key[0]), bool(key[1]), key[2] is not None, bool(key[3])])
        if usable_fields < 2:
            deduped_rows.append(row)
            continue

        if key in seen:
            duplicate_keys.append(key)
            continue

        seen.add(key)
        deduped_rows.append(row)

    if deduped_rows:
        deduped_df = pd.DataFrame(deduped_rows).reset_index(drop=True)
    else:
        deduped_df = import_df.iloc[0:0].copy()

    return deduped_df, duplicate_keys


def patch_record_by_id(record_id: str, payload_fields: dict, table_url: str, headers: dict):
    try:
        r = requests.patch(
            f"{table_url}/{record_id}",
            headers=headers,
            json={"record": {"fields": payload_fields}},
            timeout=30,
        )
        if r.status_code in (200, 201):
            return True, r.text
        return False, f"{r.status_code} | {r.text}"
    except Exception as e:
        return False, str(e)


def update_working_orders_local(working_orders: pd.DataFrame, record_id: str, payload_fields: dict):
    if working_orders.empty or "_record_id" not in working_orders.columns:
        return working_orders

    mask = working_orders["_record_id"].astype(str) == str(record_id)
    if not mask.any():
        return working_orders

    for field_name, field_value in payload_fields.items():
        if field_name not in working_orders.columns:
            working_orders[field_name] = ""
        working_orders.loc[mask, field_name] = field_value

    return working_orders


def classify_and_update_factory_row(
    current_df: pd.DataFrame,
    teable_po_col: str | None,
    teable_part_col: str | None,
    teable_qty_col: str | None,
    teable_wip_col: str | None,
    teable_customer_col: str | None,
    teable_ship_date_col: str | None,
    teable_factory_due_col: str | None,
    teable_remark_col: str | None,
    teable_tag_col: str | None,
    import_row,
    import_po_col: str | None,
    import_part_col: str | None,
    import_qty_col: str | None,
    import_wip_col: str | None,
    import_customer_col: str | None,
    import_ship_col: str | None,
    import_factory_due_col: str | None,
    import_remark_col: str | None,
    import_tag_col: str | None,
    table_url: str,
    headers: dict,
    done_wip_values=None,
    multi_select_mode=True,
):
    po_value = safe_text(import_row.get(import_po_col, "")) if import_po_col else ""
    part_value = safe_text(import_row.get(import_part_col, "")) if import_part_col else ""
    qty_value = safe_text(import_row.get(import_qty_col, "")) if import_qty_col else ""
    due_value = safe_text(import_row.get(import_factory_due_col, "")) if import_factory_due_col else ""
    wip_value = normalize_wip_value(
        safe_text(import_row.get(import_wip_col, "")) if import_wip_col else "",
        done_wip_values=done_wip_values,
    )

    ship_value = safe_text(import_row.get(import_ship_col, "")) if import_ship_col else ""
    remark_value = safe_text(import_row.get(import_remark_col, "")) if import_remark_col else ""
    raw_tags = import_row.get(import_tag_col, "") if import_tag_col else ""
    tags_value = raw_tags if isinstance(raw_tags, list) else parse_tags_from_text(raw_tags)

    imported_rec = build_record_for_match(
        po_value=po_value,
        part_value=part_value,
        qty_value=qty_value,
        due_value=due_value,
    )

    if not imported_rec["po"] and not imported_rec["part_no"]:
        return {
            "success": False,
            "action": "MANUAL_REVIEW",
            "message": "缺少 PO# / Part No，無法安全比對",
            "payload_fields": {},
            "match_info": {"score": 0, "matched_fields": [], "candidates": []},
        }

    teable_match_records = build_teable_match_records(
        current_df=current_df,
        teable_po_col=teable_po_col,
        teable_part_col=teable_part_col,
        teable_qty_col=teable_qty_col,
        teable_factory_due_col=teable_factory_due_col,
    )

    match_result = find_best_match_by_4fields(imported_rec, teable_match_records)

    if match_result["status"] != "matched":
        return {
            "success": False,
            "action": "MANUAL_REVIEW",
            "message": "找不到唯一對應主表資料",
            "payload_fields": {},
            "match_info": match_result,
        }

    target = match_result["target"]
    record_id = target.get("_record_id", "")
    if not record_id:
        return {
            "success": False,
            "action": "MANUAL_REVIEW",
            "message": "找到候選但 record_id 缺失",
            "payload_fields": {},
            "match_info": match_result,
        }

    payload_fields = {}
    if teable_wip_col and wip_value:
        payload_fields[teable_wip_col] = wip_value
    if teable_factory_due_col and due_value:
        payload_fields[teable_factory_due_col] = due_value
    if teable_ship_date_col and ship_value:
        payload_fields[teable_ship_date_col] = ship_value
    if teable_remark_col and remark_value:
        payload_fields[teable_remark_col] = remark_value[:300]
    if teable_tag_col and tags_value:
        payload_fields[teable_tag_col] = build_tags_value(tags_value, multi_select_mode=multi_select_mode)

    if not payload_fields:
        return {
            "success": False,
            "action": "SKIP",
            "message": "沒有可更新欄位",
            "payload_fields": {},
            "match_info": match_result,
        }

    success, msg = patch_record_by_id(record_id, payload_fields, table_url=table_url, headers=headers)

    return {
        "success": success,
        "action": "UPDATED" if success else "FAILED",
        "message": msg if success else f"更新失敗: {msg}",
        "payload_fields": payload_fields,
        "match_info": match_result,
        "record_id": record_id,
    }


def build_manual_review_item(
    import_row,
    import_po_col,
    import_part_col,
    import_qty_col,
    import_factory_due_col,
    import_wip_col,
    import_remark_col,
    match_info,
    reason,
    teable_po_col,
    teable_part_col,
    teable_qty_col,
    teable_factory_due_col,
    teable_wip_col,
):
    po_value = safe_text(import_row.get(import_po_col, "")) if import_po_col else ""
    part_value = safe_text(import_row.get(import_part_col, "")) if import_part_col else ""
    qty_value = safe_text(import_row.get(import_qty_col, "")) if import_qty_col else ""
    due_value = safe_text(import_row.get(import_factory_due_col, "")) if import_factory_due_col else ""
    wip_value = safe_text(import_row.get(import_wip_col, "")) if import_wip_col else ""
    remark_value = safe_text(import_row.get(import_remark_col, "")) if import_remark_col else ""

    candidate_desc = []
    for c in match_info.get("candidates", [])[:5]:
        rec = c.get("record", {})
        src_row = rec.get("_source_row")
        candidate_desc.append({
            "record_id": rec.get("_record_id", ""),
            "score": c.get("score", 0),
            "matched_fields": ", ".join(c.get("matched_fields", [])),
            "PO#": safe_text(src_row.get(teable_po_col, "")) if src_row is not None and teable_po_col else "",
            "Part No": safe_text(src_row.get(teable_part_col, "")) if src_row is not None and teable_part_col else "",
            "Qty": safe_text(src_row.get(teable_qty_col, "")) if src_row is not None and teable_qty_col else "",
            "Factory Due Date": safe_text(src_row.get(teable_factory_due_col, "")) if src_row is not None and teable_factory_due_col else "",
            "WIP": safe_text(src_row.get(teable_wip_col, "")) if src_row is not None and teable_wip_col else "",
        })

    return {
        "PO#": po_value,
        "Part No": part_value,
        "Qty": qty_value,
        "Factory Due Date": due_value,
        "New WIP": wip_value,
        "Remark": remark_value,
        "Reason": reason,
        "Best Score": match_info.get("score", 0),
        "Matched Fields": ", ".join(match_info.get("matched_fields", [])),
        "Candidates": candidate_desc,
    }
