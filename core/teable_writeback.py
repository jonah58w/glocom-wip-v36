# -*- coding: utf-8 -*-
"""
回寫 Teable 主表的 Factory PO PDF 欄位。

重要:這個模組沿用 Control Tower app.py 既有的 Teable 連線(TABLE_URL + HEADERS),
不重新建立連線,確保跟現有 patch_record_by_id 行為一致。

呼叫方式:
    from core.teable_writeback import write_pdf_url_to_records
    write_pdf_url_to_records(record_ids, pdf_url, table_url, headers)
"""

from typing import Iterable
import requests

# 必須跟 Teable 上實際加的欄位名一致
PDF_FIELD_NAME = "Factory PO PDF"


def write_pdf_url_to_records(
    record_ids: Iterable[str],
    pdf_url: str,
    table_url: str,
    headers: dict,
    field_name: str = PDF_FIELD_NAME,
) -> dict:
    """把 PDF URL 寫到指定的多個 record 上。

    Args:
        record_ids: Teable record id 清單(同一張 PO 可能有多筆 row)
        pdf_url: 要寫入的 URL 字串
        table_url: 例如 https://app.teable.ai/api/table/tbl.../record
        headers: 含 Authorization 的 dict
        field_name: Teable 上的欄位名,預設 "Factory PO PDF"

    Returns:
        {"success": int, "failed": int, "errors": [...]}
    """
    success_count = 0
    failed_count = 0
    errors = []

    for rid in record_ids:
        rid = str(rid).strip()
        if not rid:
            continue

        try:
            r = requests.patch(
                f"{table_url}/{rid}",
                headers=headers,
                json={"record": {"fields": {field_name: pdf_url}}},
                timeout=30,
            )
            if r.status_code in (200, 201):
                success_count += 1
            else:
                failed_count += 1
                errors.append(f"{rid}: {r.status_code} | {r.text[:200]}")
        except Exception as e:
            failed_count += 1
            errors.append(f"{rid}: {e}")

    return {
        "success": success_count,
        "failed": failed_count,
        "errors": errors,
    }
