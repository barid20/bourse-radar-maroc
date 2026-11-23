import json
import os
from datetime import datetime, timezone
import requests
from bs4 import BeautifulSoup

DATA_FILE = "data.json"


def load_existing():
    if not os.path.exists(DATA_FILE):
        return []
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        content = json.load(f)
    if isinstance(content, list):
        return content
    if isinstance(content, dict) and isinstance(content.get("records"), list):
        return content["records"]
    return []


def fetch_source_records():
    """
    ضع هنا منطق الجلب من المصدر المستهدف وإرجاع قائمة قواميس متوافقة مع normalizeStored:
    {
      "ticker": "ADH",
      "name": "DOUJA PROM ADDOHA",
      "date": "2026-11-23",
      "open": 38.0,
      "close": 38.5,
      "high": 39.0,
      "low": 37.8,
      "volume": 6672769,
      "shares": 173000,
      "trades": 120
    }
    """
    # مثال تحقّق بسيط من إمكانية الوصول للشبكة
    url = "https://www.casablanca-bourse.com/"
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(url, headers=headers, timeout=20)
    resp.raise_for_status()
    _ = BeautifulSoup(resp.text, "html.parser")

    # أرجِع السجلات المستخرجة هنا عند ربط مصدر/جدول محدد
    return []


def main():
    existing = load_existing()
    merged = {f"{r.get('ticker')}_{r.get('date')}": r for r in existing if r.get("ticker") and r.get("date")}

    for item in fetch_source_records():
        ticker = str(item.get("ticker", "")).strip().upper()
        date = str(item.get("date", "")).strip()
        close = float(item.get("close") or 0)
        if not ticker or not date or close <= 0:
            continue
        merged[f"{ticker}_{date}"] = {
            "ticker": ticker,
            "name": str(item.get("name") or ticker).strip(),
            "date": date,
            "open": float(item.get("open") or close),
            "close": close,
            "high": float(item.get("high") or close),
            "low": float(item.get("low") or close),
            "volume": float(item.get("volume") or 0),
            "shares": float(item.get("shares") or 0),
            "trades": float(item.get("trades") or 0),
        }

    output = {
        "updatedAt": datetime.now(timezone.utc).isoformat(),
        "records": sorted(merged.values(), key=lambda x: (x["date"], x["ticker"])),
    }

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"Saved {len(output['records'])} records to {DATA_FILE}")


if __name__ == "__main__":
    main()
