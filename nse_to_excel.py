# -*- coding: utf-8 -*-
"""
NSE option chain -> Excel ("NIFTY DATA" sheet, same 37-column layout as your Power Query)

Ye script aapki Excel ki Power Query ('NIFTY DATA' + 'Cookie' + expiry logic) ka Python version hai.
Tracker (nifty_oi_tracker.py) is file ko bina kisi badlav ke padh leta hai.

Usage:
  python nse_to_excel.py --test                 # ek baar fetch, NSE access check (kuch save nahi karta)
  python nse_to_excel.py --once                 # ek baar fetch + data/NIFTY_DATA.xlsx save
  python nse_to_excel.py --loop                 # har 3 min fetch + save (Ctrl+C se band)
  python nse_to_excel.py --test --expiry 03-Nov-2026

Optional: agar auto-cookie na chale to browser wali cookie string env var NSE_COOKIE me daal do
(wahi cookie jo aap Excel ki Cookie sheet me paste karte ho).
"""
import argparse
import datetime as dt
import os
import sys
import time

import pandas as pd
import requests

BASE = "https://www.nseindia.com"
API = BASE + "/api/option-chain-v3?type=Indices&symbol=NIFTY"
OUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "NIFTY_DATA.xlsx")
POLL_SECONDS = 180
SESSION_MAX_AGE = 480  # cookie har 8 min me refresh

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "*/*",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "en-US,en;q=0.9,en-IN;q=0.8",
    "Referer": BASE + "/option-chain",
}

# --- Excel layout (aapki Power Query ke columns, wahi order) -------------------------------
CE_KEYS = ["PChange", "change", "changeinOpenInterest", "expiryDate", "impliedVolatility",
           "lastPrice", "openInterest", "pChange", "pchangeinOpenInterest", "strikePrice",
           "totalTradedVolume", "underlying", "underlyingValue",
           "buyQuantity1", "buyPrice1", "sellQuantity1", "sellPrice1"]
CE_NAMES = ["PChange.1", "change", "changeinOpenInterest", "expiryDate", "impliedVolatility",
            "lastPrice", "openInterest", "pChange", "pchangeinOpenInterest", "strikePrice.1",
            "totalTradedVolume", "underlying", "underlyingValue",
            "BidQty", "Bid", "AskQty", "Ask"]
PE_KEYS = ["PChange", "change", "changeinOpenInterest", "expiryDate", "impliedVolatility",
           "lastPrice", "openInterest", "pchangeinOpenInterest", "strikePrice",
           "totalTradedVolume", "underlying", "underlyingValue",
           "buyQuantity1", "buyPrice1", "sellQuantity1", "sellPrice1"]
PE_NAMES = ["PChange.2", "change.1", "changeinOpenInterest.1", "expiryDate.1", "impliedVolatility.1",
            "lastPrice.1", "openInterest.1", "pchangeinOpenInterest.1", "strikePrice.2",
            "totalTradedVolume.1", "underlying.1", "underlyingValue.1",
            "BidQty.1", "Bid.1", "AskQty.1", "Ask.1"]
COLUMNS = (["expiryDates"] + CE_NAMES + PE_NAMES + ["strikePrice", "LastUpdated", "NSE_Timestamp"])
assert len(COLUMNS) == 37
DATE_COLS = ["expiryDates", "expiryDate", "expiryDate.1"]


# --- helpers -------------------------------------------------------------------------------
def parse_nse_date(text):
    return dt.datetime.strptime(text, "%d-%b-%Y").date()


# [FIX M8] din ki expiry ek baar mil gayi to dobara "bina expiry" wali API call nahi (har fetch par 2 ki jagah 1 call)
_EXP_CACHE = {}   # {aaj ki date: expiry date}


def parse_nse_ts(text):
    """NSE 'records.timestamp' (jaise '31-Oct-2026 15:30:00') -> datetime, warna None."""
    if not text:
        return None
    t = str(text).strip()
    for f in ("%d-%b-%Y %H:%M:%S", "%d-%b-%Y %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(t, f)
        except ValueError:
            continue
    return None


def new_session():
    s = requests.Session()
    s.headers.update(HEADERS)
    cookie = os.environ.get("NSE_COOKIE", "").strip()
    if cookie:
        s.headers["Cookie"] = cookie  # manual cookie (browser se)
    else:
        s.get(BASE, timeout=20)                    # homepage se cookies
        s.get(BASE + "/option-chain", timeout=20)  # option-chain page se aur cookies
    return s


def get_json(session, url):
    r = session.get(url, timeout=25)
    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code} (NSE ne mana kiya / cookie invalid) url=...{url[-40:]}")
    try:
        return r.json()
    except ValueError:
        raise RuntimeError(f"JSON nahi mila (content-type={r.headers.get('content-type')}, "
                           f"starts: {r.text[:80]!r})")


def candidate_expiries(session, today, forced=None):
    """Selector -> NSE ki nearest future expiry -> Tuesday/Monday/Friday/Wednesday fallback (aapki M logic)."""
    out = []
    if forced:
        out.append(parse_nse_date(forced))
    cached = _EXP_CACHE.get(today)
    if cached and cached not in out:
        out.append(cached)          # [FIX M8] cached expiry sabse pehle, API call bachi
    try:
        if cached or forced:
            raise LookupError("expiry already known")   # list wali extra call skip
        j = get_json(session, API)
        future = [d for d in (parse_nse_date(x) for x in j["records"]["expiryDates"]) if d >= today]
        if future:
            out.append(min(future))
    except Exception:
        pass
    dow = today.weekday()  # Monday=0
    fb = today + dt.timedelta(days=(1 - dow) if dow <= 1 else (8 - dow))
    for d in (fb, fb - dt.timedelta(days=1), fb - dt.timedelta(days=4), fb + dt.timedelta(days=1)):
        if d not in out:
            out.append(d)
    return out


def fetch_chain(session, forced_expiry=None):
    today = dt.datetime.now().date()
    last_err = None
    for d in candidate_expiries(session, today, forced_expiry):
        txt = d.strftime("%d-%b-%Y")
        try:
            j = get_json(session, API + "&expiry=" + txt)
            if len(j["filtered"]["data"]) > 0:
                _EXP_CACHE.clear()
                _EXP_CACHE[today] = d
                return j, txt
        except Exception as e:  # noqa: BLE001
            last_err = e
            if _EXP_CACHE.get(today) == d:
                _EXP_CACHE.clear()   # cached expiry ne data nahi diya -> agli baar list se dobara dhundo
    raise RuntimeError(f"Koi valid expiry ka data nahi mila. Last error: {last_err}")


def build_dataframe(j):
    records = j.get("records", {})
    nse_ts = str(records.get("timestamp", "N/A"))
    now = dt.datetime.now()
    rows = []
    for d in j["filtered"]["data"]:
        ce, pe = d.get("CE") or {}, d.get("PE") or {}
        row = [d.get("expiryDates")]
        row += [ce.get(k) for k in CE_KEYS]
        row += [pe.get(k) for k in PE_KEYS]
        row += [d.get("strikePrice"), now, nse_ts]
        rows.append(row)
    df = pd.DataFrame(rows, columns=COLUMNS)
    for c in DATE_COLS:
        df[c] = pd.to_datetime(df[c], format="%d-%b-%Y", errors="coerce")
    return df


def write_excel(df, path=OUT_FILE):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path[:-5] + ".tmp.xlsx" if path.endswith(".xlsx") else path + ".tmp.xlsx"
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="NIFTY DATA", index=False)
    os.replace(tmp, path)  # half-saved file kabhi nahi dikhegi


class Fetcher:
    def __init__(self, forced_expiry=None):
        self.session, self.born, self.forced = None, 0.0, forced_expiry

    def snapshot(self, attempts=3):
        err = None
        for i in range(attempts):
            try:
                if self.session is None or time.time() - self.born > SESSION_MAX_AGE:
                    self.session, self.born = new_session(), time.time()
                j, expiry = fetch_chain(self.session, self.forced)
                return build_dataframe(j), expiry
            except Exception as e:  # noqa: BLE001
                err = e
                self.session = None
                time.sleep(3 * (i + 1))
        raise err


def describe(df, expiry):
    ce, pe = df["openInterest"].sum(), df["openInterest.1"].sum()
    spot = df["underlyingValue"].dropna().iloc[0] if df["underlyingValue"].notna().any() else None
    return (f"expiry={expiry} strikes={len(df)} spot={spot} CE_OI={int(ce):,} PE_OI={int(pe):,} "
            f"PCR={pe / ce:.3f}" if ce else f"expiry={expiry} strikes={len(df)} spot={spot}")


# --- CLI -----------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--test", action="store_true")
    g.add_argument("--once", action="store_true")
    g.add_argument("--loop", action="store_true")
    ap.add_argument("--expiry", default=os.environ.get("NSE_EXPIRY"), help="dd-Mon-yyyy (optional)")
    ap.add_argument("--out", default=OUT_FILE)
    a = ap.parse_args()

    f = Fetcher(a.expiry)
    if a.test or a.once:
        try:
            df, expiry = f.snapshot()
        except Exception as e:  # noqa: BLE001
            print(f"FAIL: NSE se data nahi aaya -> {e}")
            return 1
        print("OK:", describe(df, expiry))
        if a.once:
            write_excel(df, a.out)
            print("saved:", a.out)
        return 0

    while True:
        try:
            df, expiry = f.snapshot()
            write_excel(df, a.out)
            print(dt.datetime.now().strftime("%H:%M:%S"), "saved |", describe(df, expiry), flush=True)
        except Exception as e:  # noqa: BLE001
            print(dt.datetime.now().strftime("%H:%M:%S"), f"fetch error: {e}", flush=True)
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    sys.exit(main())
