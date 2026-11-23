import csv
import glob
import io
import json
import os
import re
from datetime import datetime, timezone
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

LIVE_MARKET_URL = os.getenv(
    "LIVE_MARKET_URL", "https://www.casablanca-bourse.com/live-market/overview"
).strip()
DATA_DIR = os.getenv("DATA_DIR", "imports").strip()
OUT_FILE = "data.json"


def to_num(val):
    if val is None or isinstance(val, bool):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val) if val == val else 0.0
    s = (
        str(val)
        .strip()
        .replace("\u00a0", "")
        .replace(" ", "")
        .replace("٬", "")
        .replace("٫", ".")
        .replace("%", "")
    )
    if not s:
        return 0.0
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def norm_key(k):
    return re.sub(r"[^a-z0-9]", "", str(k or "").lower())


def parse_date_iso(val):
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if mo > 12 and d <= 12:
            d, mo = mo, d
        return f"{y:04d}-{mo:02d}-{d:02d}"
    m = re.match(r"^(\d{4})[./-](\d{1,2})[./-](\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return ""


def infer_ticker(name_or_path):
    base = os.path.splitext(os.path.basename(name_or_path))[0].upper()
    parts = [p for p in re.split(r"[\s_-]+", base) if p]
    if not parts:
        return ""
    if parts[0] in ("COURS", "HISTO", "DATA", "CSE", "BVC") and len(parts) > 1:
        cand = re.sub(r"[^A-Z0-9.-]", "", parts[1])
        if cand:
            return cand
    return re.sub(r"[^A-Z0-9.-]", "", parts[0])


def pick(row_norm, aliases):
    for a in aliases:
        nk = norm_key(a)
        if nk in row_norm and row_norm[nk] not in ("", None):
            return row_norm[nk]
    return ""


def normalize_row(raw, fallback_ticker="", fallback_date=""):
    rn = {norm_key(k): v for k, v in raw.items()}
    ticker = re.sub(
        r"[^A-Z0-9.-]",
        "",
        str(pick(rn, ["ticker", "code", "symbole", "mnemonic", "instrument"]) or fallback_ticker)
        .strip()
        .upper(),
    )
    date = parse_date_iso(pick(rn, ["date", "session", "seance", "séance", "heure", "time"])) or fallback_date
    close = to_num(
        pick(rn, ["close", "lastprice", "derniercours", "coursdecloture", "cours", "last", "price"])
    )
    if not ticker or not date or close <= 0:
        return None

    open_p = to_num(pick(rn, ["open", "opening", "ouverture", "premiercours"])) or close
    high_p = max(close, open_p, to_num(pick(rn, ["high", "plushaut", "max"])))
    low_raw = to_num(pick(rn, ["low", "plusbas", "min"]))
    low_p = min(close, open_p, low_raw) if low_raw > 0 else min(close, open_p)

    return {
        "ticker": ticker,
        "name": str(pick(rn, ["name", "instrument", "valeur", "libelle"]) or ticker).strip(),
        "date": date,
        "open": open_p,
        "close": close,
        "high": high_p,
        "low": low_p,
        "volume": to_num(pick(rn, ["volume", "volumemad", "capitaux", "montant"])),
        "shares": to_num(pick(rn, ["shares", "tradedshares", "titresechanges", "quantite"])),
        "trades": to_num(pick(rn, ["trades", "numberoftrades", "nbtransactions", "transactions"])),
    }


def parse_csv_bytes(content_bytes, fallback_ticker):
    text = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = content_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if not text:
        return []
    lines = [ln for ln in text.splitlines() if ln.strip()]
    first_line = lines[0] if lines else ""
    delimiter = ";" if first_line.count(";") > first_line.count(",") else ","
    reader = csv.DictReader(io.StringIO("\n".join(lines)), delimiter=delimiter)
    return [rec for row in reader if (rec := normalize_row(row, fallback_ticker))]


def parse_xlsx_bytes(content_bytes, fallback_ticker):
    wb = load_workbook(filename=io.BytesIO(content_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h or "").strip() for h in rows[0]]
    out = []
    for r in rows[1:]:
        if not any(r):
            continue
        obj = {headers[i]: (r[i] if i < len(headers) else "") for i in range(len(headers))}
        rec = normalize_row(obj, fallback_ticker)
        if rec:
            out.append(rec)
    return out


def parse_bytes_by_name(content_bytes, name_hint):
    ext = os.path.splitext(name_hint.lower())[1]
    fb = infer_ticker(name_hint)
    if ext in (".xlsx", ".xlsm"):
        return parse_xlsx_bytes(content_bytes, fb)
    return parse_csv_bytes(content_bytes, fb)


def fetch_live_market_rows():
    if not LIVE_MARKET_URL:
        return []
    today_iso = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    try:
        resp = requests.get(
            LIVE_MARKET_URL,
            headers={
                "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
                "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
            },
            timeout=25,
        )
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        extracted = []

        for table in soup.find_all("table"):
            trs = table.find_all("tr")
            if len(trs) < 2:
                continue
            headers = [th.get_text(" ", strip=True) for th in trs[0].find_all(["th", "td"])]
            if not headers:
                continue
            for tr in trs[1:]:
                cells = [td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"])]
                if len(cells) < 2:
                    continue
                obj = {headers[i]: cells[i] for i in range(min(len(headers), len(cells)))}
                rec = normalize_row(obj, fallback_date=today_iso)
                if rec:
                    extracted.append(rec)

        return extracted
    except Exception as e:
        print(f"Live market fetch note: {e}")
        return []


def load_existing_records():
    if not os.path.exists(OUT_FILE):
        return []
    try:
        with open(OUT_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return []
    items = data if isinstance(data, list) else data.get("records", [])
    if not isinstance(items, list):
        return []
    return [rec for item in items if isinstance(item, dict) and (rec := normalize_row(item))]


def scan_folder():
    if not os.path.isdir(DATA_DIR):
        return [], []
    files = sorted(
        glob.glob(os.path.join(DATA_DIR, "*.csv"))
        + glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
        + glob.glob(os.path.join(DATA_DIR, "*.xlsm"))
    )
    rows = []
    for fp in files:
        with open(fp, "rb") as f:
            rows.extend(parse_bytes_by_name(f.read(), fp))
    return files, rows


def main():
    merged = {f"{r['ticker']}_{r['date']}": r for r in load_existing_records()}
    files, folder_rows = scan_folder()
    live_rows = fetch_live_market_rows()

    for rec in folder_rows + live_rows:
        merged[f"{rec['ticker']}_{rec['date']}"] = rec

    sorted_records = sorted(merged.values(), key=lambda x: (x["date"], x["ticker"]))
    payload = {
        "updatedAt": datetime.now(timezone.utc).isoformat(),
        "sourceUrl": LIVE_MARKET_URL,
        "filesProcessed": [os.path.basename(p) for p in files],
        "liveRowsExtracted": len(live_rows),
        "count": len(sorted_records),
        "records": sorted_records,
    }

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(
        f"Folder files: {len(files)}, folder rows: {len(folder_rows)}, "
        f"live rows: {len(live_rows)}, total: {len(sorted_records)}"
    )


if __name__ == "__main__":
    main()
