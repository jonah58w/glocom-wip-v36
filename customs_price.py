"""
customs_price.py
報關單價計算引擎

儲存格式 customs_prices.json：
{
  "PCBWOO005000": {
    "price": 0.47,
    "last_change_date": "2026-03-26",
    "factory_price_ntd": 11.0,
    "exchange_rate_used": 31.265,
    "reason": "...",
    "updated_at": "2026-04-15"
  },
  ...
}
"""

import json, os
from datetime import date, datetime
from typing import Optional, Tuple

PRICE_DB_FILE = "customs_prices.json"
DEFAULT_MARGIN = 0.06   # 6%
CHANGE_THRESHOLD = 0.03  # 3% 差異門檻


# ─────────────────────────────────────────
# I/O helpers
# ─────────────────────────────────────────

def load_price_db() -> dict:
    if os.path.exists(PRICE_DB_FILE):
        with open(PRICE_DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_price_db(db: dict) -> None:
    with open(PRICE_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2, default=str)


# ─────────────────────────────────────────
# Core calculation
# ─────────────────────────────────────────

def calc_new_price(factory_ntd: float, exchange_rate: float,
                   margin: float = DEFAULT_MARGIN) -> Optional[float]:
    """工廠單價(NTD) × (1 + margin%) ÷ 匯率 → 建議報關單價(USD)"""
    if factory_ntd and exchange_rate and exchange_rate > 0:
        return round(factory_ntd * (1 + margin) / exchange_rate, 4)
    return None


def _parse_date(d) -> Optional[date]:
    if d is None:
        return None
    if isinstance(d, date):
        return d if not isinstance(d, datetime) else d.date()
    if isinstance(d, str):
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(d[:10], fmt).date()
            except ValueError:
                continue
    return None


def decide_customs_price(
    pn: str,
    factory_price_ntd: float,
    exchange_rate: float,
    shipment_date,
    price_db: dict
) -> Tuple[float, bool, str]:
    """
    決定本次使用舊單價還是新單價。

    Returns
    -------
    (price, is_new_price, reason_str)
    """
    pn_key = pn.strip()
    new_price = calc_new_price(factory_price_ntd, exchange_rate)
    history = price_db.get(pn_key, {})
    old_price: Optional[float] = history.get("price")
    last_change_str = history.get("last_change_date")

    # ── 首次出貨：直接用新單價 ──
    if old_price is None or old_price == 0:
        return (new_price or 0), True, "首次設定報關單價，使用 6% 新單價"

    # ── 無法計算新單價（資料不足）──
    if new_price is None:
        return old_price, False, "工廠單價或匯率缺失，沿用舊單價"

    # ── 舊單價 < 工廠成本（倒掛）→ 強制更新 ──
    factory_cost_usd = factory_price_ntd / exchange_rate if exchange_rate else 0
    if old_price < factory_cost_usd:
        return new_price, True, "舊單價小於工廠單價，須使用 6% 新單價"

    # ── 計算距上次變更月數 ──
    sd = _parse_date(shipment_date) or date.today()
    lcd = _parse_date(last_change_str)

    if lcd:
        months = (sd.year - lcd.year) * 12 + (sd.month - lcd.month)
    else:
        months = 999  # 視為超過一年

    period_label = ("未過半年" if months < 6
                    else "已過半年" if months < 12
                    else "已過一年")

    # ── 未過半年：用舊 ──
    if months < 6:
        return old_price, False, f"距上次單價變更{period_label}，使用 6% 舊單價"

    # ── 超過半年 ──
    diff_pct = abs(old_price - new_price) / old_price if old_price else 0

    if new_price > old_price:
        return old_price, False, f"距上次單價變更{period_label}，新單價>舊單價，使用 6% 舊單價"

    # old > new
    if diff_pct <= CHANGE_THRESHOLD:
        return old_price, False, (
            f"距上次單價變更{period_label}，舊單價>新單價，"
            f"差異 {diff_pct*100:.1f}%≤3%，使用 6% 舊單價"
        )
    else:
        return new_price, True, (
            f"距上次單價變更{period_label}，舊單價>新單價，"
            f"差異 {diff_pct*100:.1f}%>3%，使用 6% 新單價"
        )


def confirm_and_save(pn: str, price: float, shipment_date,
                     factory_price_ntd: float, exchange_rate: float,
                     reason: str) -> dict:
    """確認後將報關單價寫入 JSON 資料庫"""
    db = load_price_db()
    pn_key = pn.strip()
    db[pn_key] = {
        "price": price,
        "last_change_date": str(_parse_date(shipment_date) or date.today()),
        "factory_price_ntd": factory_price_ntd,
        "exchange_rate_used": exchange_rate,
        "reason": reason,
        "updated_at": str(date.today()),
    }
    save_price_db(db)
    return db


def get_price_for_pn(pn: str) -> Optional[float]:
    """快速查詢某 P/N 目前的報關單價"""
    db = load_price_db()
    return db.get(pn.strip(), {}).get("price")
