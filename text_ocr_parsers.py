import pandas as pd
import requests
from config import TEABLE_TOKEN, TABLE_URL, HEADERS
from utils import normalize_columns, get_series_by_col


def load_orders():
    if not TEABLE_TOKEN:
        return pd.DataFrame(), "NO_TOKEN", "TEABLE_TOKEN is empty"
    try:
        response = requests.get(
            TABLE_URL,
            headers=HEADERS,
            params={"fieldKeyType": "name", "cellFormat": "text", "take": 1000},
            timeout=30,
        )
        status = response.status_code
        text = response.text
        if status != 200:
            return pd.DataFrame(), status, text
        data = response.json()
        rows = []
        for rec in data.get("records", []):
            fields = rec.get("fields", {})
            fields["_record_id"] = rec.get("id", "")
            rows.append(fields)
        df = normalize_columns(pd.DataFrame(rows))
        return df, status, text
    except Exception as e:
        return pd.DataFrame(), "EXCEPTION", str(e)


def find_record_id_by_po(df: pd.DataFrame, po_value: str, po_col: str | None):
    if df.empty or not po_col or po_col not in df.columns:
        return None
    po_series = get_series_by_col(df, po_col)
    if po_series is None:
        return None
    matched = df[po_series.astype(str).str.strip().str.lower() == str(po_value).strip().lower()]
    if matched.empty:
        return None
    return matched.iloc[0].get("_record_id") if "_record_id" in matched.columns else None


def upsert_to_teable(current_df: pd.DataFrame, po_col_name: str, po_value: str, updates: dict):
    if not po_value:
        return False, "PO is empty"
    record_id = find_record_id_by_po(current_df, po_value=po_value, po_col=po_col_name)
    payload_fields = dict(updates)
    payload_fields[po_col_name] = po_value
    try:
        if record_id:
            r = requests.patch(f"{TABLE_URL}/{record_id}", headers=HEADERS, json={"record": {"fields": payload_fields}}, timeout=30)
        else:
            r = requests.post(TABLE_URL, headers=HEADERS, json={"records": [{"fields": payload_fields}]}, timeout=30)
        if r.status_code in (200, 201):
            return True, r.text
        return False, f"{r.status_code} | {r.text}"
    except Exception as e:
        return False, str(e)


def patch_record_by_id(record_id: str, payload_fields: dict):
    try:
        r = requests.patch(f"{TABLE_URL}/{record_id}", headers=HEADERS, json={"record": {"fields": payload_fields}}, timeout=30)
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
