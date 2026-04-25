"""端到端測試:從主表 CSV 撈 G1150014-01(2 個 P/N)產 PDF。

這個測試模擬整個流程:
1. 載入 Teable 主表(從 CSV 模擬)
2. 列出可選的西拓訂單編號
3. 撈某個編號的所有 row
4. 帶入工廠主檔
5. 產出 docx + PDF
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd
from core.teable_query import (
    list_glocom_po_options,
    get_po_rows,
    build_po_context,
    parse_glocom_po_no,
    COL_GLOCOM_PO,
)
from core.factory_master import get_factory, list_factory_options
from core.pdf_generator import generate_po_files


def load_test_orders() -> pd.DataFrame:
    """模擬 Control Tower 載入 Teable 後的 DataFrame。"""
    csv_path = "/mnt/user-data/uploads/主表.csv"
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    # 加 fake _record_id(實際 Teable 載入時會有)
    df["_record_id"] = [f"recTEST{i:04d}" for i in range(len(df))]
    return df


def test_list_options():
    """測試列出主表的西拓訂單編號。"""
    df = load_test_orders()
    options = list_glocom_po_options(df)
    print(f"✓ 主表共 {len(df)} row,共有 {len(options)} 個西拓訂單編號")
    print("\n  前 5 個選項:")
    for po, label in options[:5]:
        print(f"  {label}")


def test_factory_options():
    """測試工廠下拉選單。"""
    options = list_factory_options()
    print(f"\n✓ 工廠主檔共 {len(options)} 家")
    for short, label in options:
        print(f"  {label}")


def test_parse_po_no():
    """測試訂單號解析。"""
    cases = [
        ("ET1150029-01", "ET", "GLOCOM"),
        ("EW1150018-01", "EW", "EUSWAY"),
        ("G1150030-01", "G", "GLOCOM"),
        ("G1150014-01", "G", "GLOCOM"),
    ]
    for po, expected_type, expected_issue in cases:
        result = parse_glocom_po_no(po)
        assert result["order_type"] == expected_type, f"{po} → {result}"
        assert result["issuing_company"] == expected_issue, f"{po} → {result}"
        print(f"✓ {po} → {result['order_type']} / {result['issuing_company']}")


def test_end_to_end_g1150014():
    """完整測試:G1150014-01(VORNE 2 P/N → 祥竑)."""
    df = load_test_orders()
    target_po = "G1150014-01"

    rows = get_po_rows(df, target_po)
    print(f"\n=== {target_po} 撈到 {len(rows)} 筆 ===")
    if rows.empty:
        print(f"⚠️ 主表沒有 {target_po},改試其他 PO")
        # 找一個有資料的
        options = list_glocom_po_options(df)
        if not options:
            print("無任何 PO 可測")
            return
        target_po = options[0][0]
        rows = get_po_rows(df, target_po)
        print(f"  改測 {target_po}: {len(rows)} 筆")

    # 印出每一筆的關鍵欄位
    qty_col = "Order Q'TY\n (PCS)"
    for i, r in rows.iterrows():
        print(f"  品項 {i+1}: P/N={r['P/N']}, Qty={r[qty_col]}, 工廠={r['工廠']}")

    # 找工廠
    first_factory_short = str(rows.iloc[0]["工廠"]).strip()
    factory = get_factory(first_factory_short)
    if not factory:
        print(f"⚠️ 找不到工廠 {first_factory_short},用空殼")
        factory = {"factory_name": first_factory_short, "address": "[請補]"}

    # 建 PO context
    po_ctx = build_po_context(target_po, rows, factory)
    print(f"\n  訂單類型: {po_ctx['order_type']} → {po_ctx['issuing_company']}")
    print(f"  客戶: {po_ctx['customer_name']}")
    print(f"  工廠: {po_ctx['factory']['factory_name']}")
    print(f"  品項數: {len(po_ctx['items'])}")
    print(f"  合計: {po_ctx['currency']} {po_ctx['total_amount']:,.2f}")

    # 產 PDF
    result = generate_po_files(po_ctx)
    print(f"\n  docx → {result['docx_path']}")
    print(f"  pdf  → {result['pdf_path']}")
    if result.get("error"):
        print(f"  ⚠️ {result['error']}")


if __name__ == "__main__":
    print("=" * 60)
    print("Test 1: 列出主表西拓訂單編號")
    print("=" * 60)
    test_list_options()

    print("\n" + "=" * 60)
    print("Test 2: 工廠主檔")
    print("=" * 60)
    test_factory_options()

    print("\n" + "=" * 60)
    print("Test 3: 訂單號解析")
    print("=" * 60)
    test_parse_po_no()

    print("\n" + "=" * 60)
    print("Test 4: 端到端 G1150014-01")
    print("=" * 60)
    test_end_to_end_g1150014()

    print("\n" + "=" * 60)
    print("✓ 全部通過")
    print("=" * 60)
