import csv
import glob
import io
import json
import os
import re
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook

DATA_DIR = os.getenv("DATA_DIR", "imports").strip()
OUT_FILE = "data.json"
# اختياري: أضف رموز أسهم محددة مثل "ADH,ATW,IAM,BCP,LHM" أو اتركه فارغاً لجلب السوق المباشر بالكامل
TARGET_TICKERS = [t.strip().upper() for t in os.getenv("TARGET_TICKERS", "").split(",") if t.strip()]


def to_num(val):
    if val is None or isinstance(val, bool):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val) if val == val else 0.0
    s = str(val).strip().replace("\u00a0", "").replace(" ", "").replace("٬", "").replace("٫", ".")
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
        str(pick(rn, ["ticker", "code", "symbole", "mnemonic", "instrument", "symbol"]) or fallback_ticker)
        .strip()
        .upper(),
    )
    date = parse_date_iso(pick(rn, ["date", "session", "seance", "séance", "datetime"])) or fallback_date
    close = to_num(
        pick(rn, ["close", "lastprice", "derniercours", "coursdecloture", "cours", "last", "price", "value"])
    )
    if not ticker or not date or close <= 0:
        return None

    open_p = to_num(pick(rn, ["open", "opening", "ouverture", "premiercours"])) or close
    high_p = max(close, open_p, to_num(pick(rn, ["high", "plushaut", "max"])))
    low_raw = to_num(pick(rn, ["low", "plusbas", "min"]))
    low_p = min(close, open_p, low_raw) if low_raw > 0 else min(close, open_p)

    return {
        "ticker": ticker,
        "name": str(pick(rn, ["name", "instrument", "valeur", "libelle", "label"]) or ticker).strip(),
        "date": date,
        "open": open_p,
        "close": close,
        "high": high_p,
        "low": low_p,
        "volume": to_num(pick(rn, ["volume", "volumemad", "capitaux", "montant", "turnover"])),
        "shares": to_num(pick(rn, ["shares", "tradedshares", "titresechanges", "quantite", "qty"])),
        "trades": to_num(pick(rn, ["trades", "numberoftrades", "nbtransactions", "transactions"])),
    }


def fetch_from_casabourse():
    rows = []
    try:
        import casabourse as cb
    except ImportError as e:
        print(f"casabourse import error: {e}")
        return rows

    today_iso = datetime.now(timezone.utc).strftime("%Y-%m-%d")

    for fn_name in ("get_live_market_data", "get_market_data"):
        fn = getattr(cb, fn_name, None)
        if callable(fn):
            try:
                df_live = fn()
                if df_live is not None and not getattr(df_live, "empty", True):
                    print(f"{fn_name} columns:", list(df_live.columns))
                    for rec in df_live.to_dict(orient="records"):
                        normed = normalize_row(rec, fallback_date=today_iso)
                        if normed:
                            rows.append(normed)
                if rows:
                    break
            except Exception as e:
                print(f"{fn_name} error: {e}")

    if TARGET_TICKERS and hasattr(cb, "get_historical_data_auto"):
        start_str = (datetime.now(timezone.utc) - timedelta(days=30)).strftime("%Y-%m-%d")
        for tkr in TARGET_TICKERS:
            try:
                df_hist = cb.get_historical_data_auto(tkr, start_str, today_iso)
                if df_hist is not None and not getattr(df_hist, "empty", True):
                    for rec in df_hist.to_dict(orient="records"):
                        normed = normalize_row(rec, fallback_ticker=tkr, fallback_date=today_iso)
                        if normed:
                            rows.append(normed)
            except Exception as e:
                print(f"get_historical_data_auto({tkr}) error: {e}")

    return rows


def infer_ticker(filepath):
    base = os.path.splitext(os.path.basename(filepath))[0].upper()
    parts = [p for p in re.split(r"[\s_-]+", base) if p]
    if not parts:
        return ""
    if parts[0] in ("COURS", "HISTO", "DATA", "CSE", "BVC") and len(parts) > 1:
        cand = re.sub(r"[^A-Z0-9.-]", "", parts[1])
        if cand:
            return cand
    return re.sub(r"[^A-Z0-9.-]", "", parts[0])


def parse_csv_file(filepath, fallback_ticker):
    with open(filepath, "rb") as f:
        raw = f.read()
    text = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if not text:
        return []
    lines = [ln for ln in text.splitlines() if ln.strip()]
    delimiter = ";" if lines and lines[0].count(";") > lines[0].count(",") else ","
    reader = csv.DictReader(io.StringIO("\n".join(lines)), delimiter=delimiter)
    return [rec for r in reader if (rec := normalize_row(r, fallback_ticker))]


def parse_xlsx_file(filepath, fallback_ticker):
    wb = load_workbook(filename=filepath, data_only=True)
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


def scan_folder():
    if not os.path.isdir(DATA_DIR):
        return [], []
    files = sorted(
        glob.glob(os.path.join(DATA_DIR, "*.csv"))
        + glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
        + glob.glob(os.path.join(DATA_DIR, "*.xlsm"))
    )
    out = []
    for fp in files:
        fb = infer_ticker(fp)
        if fp.lower().endswith((".xlsx", ".xlsm")):
            out.extend(parse_xlsx_file(fp, fb))
        else:
            out.extend(parse_csv_file(fp, fb))
    return files, out


def load_existing():
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


def main():
    merged = {f"{r['ticker']}_{r['date']}": r for r in load_existing()}
    files, folder_rows = scan_folder()
    cb_rows = fetch_from_casabourse()

    for rec in folder_rows + cb_rows:
        merged[f"{rec['ticker']}_{rec['date']}"] = rec

    sorted_records = sorted(merged.values(), key=lambda x: (x["date"], x["ticker"]))
    payload = {
        "updatedAt": datetime.now(timezone.utc).isoformat(),
        "filesProcessed": [os.path.basename(p) for p in files],
        "casabourseRows": len(cb_rows),
        "count": len(sorted_records),
        "records": sorted_records,
    }

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"Folder rows: {len(folder_rows)}, Casabourse rows: {len(cb_rows)}, total: {len(sorted_records)}")


if __name__ == "__main__":
    main()
