# -*- coding: utf-8 -*-
"""
GitHub Actions runner: nse_to_excel (data fetch) + nifty_oi_tracker.py (chart + Telegram) ek saath.
Aapki nifty_oi_tracker.py me koi badlav nahi hota (token bhi waisa hi rehta hai) - sirf file ke paths yahin override hote hain.

Usage:  python run_actions.py --until 14:50
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

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(HERE, "data", "NIFTY_DATA.xlsx")

# tracker ke Windows paths ki jagah runner ke paths
t.EXCEL_FILE = DATA_FILE
t.OUTPUT_FILE = os.path.join(HERE, "nifty_oi_tracker.jpg")


def fetch_loop(stop):
    fetcher = nx.Fetcher()
    while not stop.is_set():
        try:
            df, expiry = fetcher.snapshot()
            nx.write_excel(df, DATA_FILE)
            print(dt.datetime.now().strftime("%H:%M:%S"), "data saved |", nx.describe(df, expiry), flush=True)
        except Exception as e:  # noqa: BLE001
            print(dt.datetime.now().strftime("%H:%M:%S"), f"fetch error: {e}", flush=True)
        stop.wait(nx.POLL_SECONDS)


def tracker_thread():
    try:
        t.main()
    except Exception as e:  # noqa: BLE001
        print("tracker stopped:", repr(e), flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--until", default="15:42", help="HH:MM IST - is time par band (cache save ho jaye)")
    a = ap.parse_args()
    h, m = map(int, a.until.split(":"))
    deadline = dt.datetime.now().replace(hour=h, minute=m, second=0, microsecond=0)

    stop = threading.Event()
    # pehla data tracker start hone se pehle hi taiyar (handover ke baad alert me deri nahi)
    try:
        df, expiry = nx.Fetcher().snapshot()
        nx.write_excel(df, DATA_FILE)
        print("first data ready |", nx.describe(df, expiry), flush=True)
    except Exception as e:  # noqa: BLE001
        print("first fetch error:", e, flush=True)
    threading.Thread(target=fetch_loop, args=(stop,), daemon=True).start()
    threading.Thread(target=tracker_thread, daemon=True).start()

    while dt.datetime.now() < deadline:
        time.sleep(15)
    stop.set()
    print("deadline reached, exiting cleanly", flush=True)


if __name__ == "__main__":
    main()
