# -*- coding: utf-8 -*-
"""
GitHub Actions runner: nse_to_excel (data fetch) + nifty_oi_tracker.py (chart + Telegram) ek saath.
Aapki nifty_oi_tracker.py me koi badlav nahi hota (token bhi waisa hi rehta hai) - sirf file ke paths yahin override hote hain.

Usage:  python run_actions.py --until 14:50

Is version me:
  * watchdog: tracker thread mar jaye to Telegram alert + auto-restart (max 10 baar)
  * fetch lagataar fail ho (3 baar ~ 9 min) to Telegram alert, theek hone par "recovered" msg
  * holiday: 10:15 tak bhi NSE ka data purani date ka ho to run turant band (Actions minutes bachte hain)
  * deadline har 5s par check (handover 12:30 ke aas-paas tight na rahe)
  * tracker_ran.flag: sirf asli run (backup/late/holiday nahi) cache save kare
"""
import argparse
import datetime as dt
import os
import threading
import time

os.environ.setdefault("TZ", "Asia/Kolkata")  # runner UTC hota hai, tracker IST maanta hai
if hasattr(time, "tzset"):
    time.tzset()

import nse_to_excel as nx  # noqa: E402
import nifty_oi_tracker as t  # noqa: E402
import logging  # noqa: E402
logging.getLogger("matplotlib").setLevel(logging.WARNING)  # findfont DEBUG spam band
logging.getLogger("PIL").setLevel(logging.WARNING)

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(HERE, "data", "NIFTY_DATA.xlsx")
RAN_FLAG = os.path.join(HERE, "tracker_ran.flag")

FAIL_ALERT_AFTER = 3        # itne lagataar fetch fail (3 x 3min ~ 9 min) par alert
MAX_TRACKER_RESTARTS = 10
HOLIDAY_CHECK_FROM = (10, 15)   # is time ke baad bhi NSE data purani date ka = holiday

# tracker ke Windows paths ki jagah runner ke paths
t.EXCEL_FILE = DATA_FILE
t.OUTPUT_FILE = os.path.join(HERE, "nifty_oi_tracker.jpg")

STATE = {"nse_date": None, "fails": 0, "alerted": False}


def tg(msg):
    try:
        t.send_telegram_text(msg)
    except Exception as e:  # noqa: BLE001
        print("telegram error:", e, flush=True)


def note_data(df):
    """Fetch hue data ki NSE-timestamp date yaad rakho (holiday pehchanne ke liye)."""
    try:
        ts = nx.parse_nse_ts(df["NSE_Timestamp"].iloc[0])
        if ts:
            STATE["nse_date"] = ts.date()
    except Exception:  # noqa: BLE001
        pass


def is_holiday_now():
    n = dt.datetime.now()
    if (n.hour, n.minute) < HOLIDAY_CHECK_FROM:
        return False
    d = STATE["nse_date"]
    return d is not None and d < n.date()


def fetch_ok():
    STATE["fails"] = 0
    if STATE["alerted"]:
        STATE["alerted"] = False
        tg("NSE data fetch wapas theek ho gaya (recovered).")


def fetch_failed(err):
    STATE["fails"] += 1
    if STATE["fails"] >= FAIL_ALERT_AFTER and not STATE["alerted"] and t.is_market_open():
        STATE["alerted"] = True
        tg(f"WARNING: NSE fetch {STATE['fails']} baar lagataar fail: {str(err)[:150]}")


def fetch_loop(stop):
    fetcher = nx.Fetcher()
    stop.wait(nx.POLL_SECONDS - 15)  # pehla data main() me ho chuka, doosra point tab (09:27 do baar na aaye)
    while not stop.is_set():
        try:
            df, expiry = fetcher.snapshot()
            nx.write_excel(df, DATA_FILE)
            note_data(df)
            fetch_ok()
            print(dt.datetime.now().strftime("%H:%M:%S"), "data saved |", nx.describe(df, expiry), flush=True)
        except Exception as e:  # noqa: BLE001
            print(dt.datetime.now().strftime("%H:%M:%S"), f"fetch error: {e}", flush=True)
            fetch_failed(e)
        stop.wait(nx.POLL_SECONDS)


def tracker_thread(stop):
    """Watchdog: t.main() crash kare to alert + restart."""
    restarts = 0
    while not stop.is_set():
        try:
            t.main()
            return
        except Exception as e:  # noqa: BLE001
            restarts += 1
            print(f"tracker stopped ({restarts}/{MAX_TRACKER_RESTARTS}):", repr(e), flush=True)
            if restarts >= MAX_TRACKER_RESTARTS:
                tg(f"ERROR: tracker {restarts} baar crash hua, ab band: {str(e)[:150]}")
                return
            tg(f"WARNING: tracker crash hua ({restarts}/{MAX_TRACKER_RESTARTS}): {str(e)[:150]} - 15s me restart")
            stop.wait(15)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--until", default="15:42", help="HH:MM IST - is time par band (cache save ho jaye)")
    a = ap.parse_args()
    h, m = map(int, a.until.split(":"))
    deadline = dt.datetime.now().replace(hour=h, minute=m, second=0, microsecond=0)

    print(f"[RUN] now={dt.datetime.now():%Y-%m-%d %H:%M:%S} until={a.until} deadline={deadline:%H:%M:%S}", flush=True)
    if dt.datetime.now() >= deadline:  # backup/late run: is session ka time nikal chuka
        print("deadline already passed, nothing to do", flush=True)
        return

    stop = threading.Event()
    # pehla data tracker start hone se pehle hi taiyar (handover ke baad alert me deri nahi)
    try:
        df, expiry = nx.Fetcher().snapshot()
        nx.write_excel(df, DATA_FILE)
        note_data(df)
        print("first data ready |", nx.describe(df, expiry), flush=True)
    except Exception as e:  # noqa: BLE001
        print("first fetch error:", e, flush=True)
        fetch_failed(e)

    if is_holiday_now():   # tracker start hi nahi - holiday par Actions minutes + duplicate msg dono bachte hain
        print(f"NSE data purani date ({STATE['nse_date']}) ka hai - market holiday, exiting", flush=True)
        return

    open(RAN_FLAG, "w").write(dt.datetime.now().isoformat())   # workflow: sirf asli run cache save kare
    threading.Thread(target=fetch_loop, args=(stop,), daemon=True).start()
    threading.Thread(target=tracker_thread, args=(stop,), daemon=True).start()

    while dt.datetime.now() < deadline:
        time.sleep(5)
        if is_holiday_now():
            print(f"holiday detected (NSE data date {STATE['nse_date']}), exiting", flush=True)
            break
    stop.set()
    print(f"deadline reached ({deadline:%H:%M}), now={dt.datetime.now():%H:%M:%S}, exiting cleanly", flush=True)


if __name__ == "__main__":
    main()
