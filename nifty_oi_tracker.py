# -*- coding: utf-8 -*-
"""
+==================================================================+
|      NIFTY OPTIONS OI TRACKER  -  Excel + Telegram              |
|  * Reads data from NIFTY_OPTION_CHAIN Excel file (auto-updated) |
|  * 3-panel chart: Total OI | COI | PCR + MaxPain + Underlying   |
|  * Sends chart to Telegram every 3 minutes                      |
+==================================================================+

SETUP:
    pip install openpyxl matplotlib pandas numpy requests

USAGE:
    python nifty_oi_tracker.py              <- live mode
    python nifty_oi_tracker.py --demo       <- demo/test mode
"""

import os, sys, time, datetime, requests, logging

# -- FILE LOGGING - sab kuch tracker_log.txt mein likhega ---------
_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tracker_log.txt")
logging.basicConfig(
    level    = logging.DEBUG,
    format   = "[%(asctime)s] %(message)s",
    datefmt  = "%d-%m-%Y %H:%M:%S",
    handlers = [
        logging.FileHandler(_log_path, encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ]
)
log = logging.getLogger()
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot   as plt
import matplotlib.gridspec as gridspec
import matplotlib.ticker   as mticker
import matplotlib.dates    as mdates
import numpy  as np
import pandas as pd
import openpyxl
import warnings
warnings.filterwarnings('ignore')

# -- ***  CONFIG  *** ------------------------------------------------------

# Path to your Excel file - update this to your actual path
EXCEL_FILE    = r"C:\Users\faird\Desktop\mbot\NIFTY_OPTION_CHAIN_dynamic_expiry1.xlsm"

TELEGRAM = {
    # GitHub Secrets / environment variables se aate hain (code me token nahi)
    "bot_token" : os.environ.get("TELEGRAM_BOT_TOKEN", ""),
    "chat_id"   : os.environ.get("TELEGRAM_CHAT_ID", ""),
}

POLL_INTERVAL = 180           # 3 minutes (same as Excel refresh)
OUTPUT_FILE   = r"C:\Users\faird\Desktop\mbot\nifty_oi_tracker.jpg"
SHEET_NAME    = "NIFTY DATA"

# -- THEME - WHITE ----------------------------------------------------------
DARK_BG  = '#ffffff'
PANEL_BG = '#ffffff'
GRID_CLR = '#dddddd'
TEXT_CLR = '#1a1a2e'
CE_CLR   = '#1565c0'
PE_CLR   = '#c62828'
DIFF_CLR = '#546e7a'
PCR_CLR  = '#2e7d32'
MAX_CLR  = '#c62828'
UND_CLR  = '#1565c0'

# -- READ EXCEL -------------------------------------------------------------
def _parse_nse_ts(v):
    """[FIX H1] NSE_Timestamp (col AK=37) -> datetime. 'N/A' / bekar value par None."""
    if isinstance(v, datetime.datetime):
        return v
    t = str(v or "").strip()
    for f in ("%d-%b-%Y %H:%M:%S", "%d-%b-%Y %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(t, f)
        except ValueError:
            continue
    return None

def read_excel_snapshot(excel_path: str) -> dict | None:
    """
    Read current OI data from Excel sheet.
    Returns dict with: ce_oi_by_strike, pe_oi_by_strike,
                       underlying, last_updated, expiry
    """
    if not os.path.exists(excel_path):
        log.info(f"[EXCEL] X File not found: {excel_path}")
        return None

    try:
        # PATCH_V2: temp copy se padho (lock / half-saved file issue nahi aayega)
        import shutil, tempfile
        _tmp = os.path.join(tempfile.gettempdir(), 'nifty_oi_read_copy.xlsm')
        shutil.copy2(excel_path, _tmp)
        wb = openpyxl.load_workbook(_tmp, keep_vba=False, data_only=True)
        ws = wb[SHEET_NAME]

        ce_oi_by_strike = {}
        pe_oi_by_strike = {}
        underlying      = None
        last_updated    = None
        nse_ts          = None
        expiry          = None
        tot_ce_coi = 0   # PATCH_CHART_V3
        tot_pe_coi = 0

        # Columns (1-based):
        # A=1  expiryDates
        # H=8  CE openInterest
        # K=11 CE strikePrice
        # Y=25 PE openInterest
        # AI=35 PE strikePrice
        # N=14 underlyingValue
        # AJ=36 LastUpdated

        for row in ws.iter_rows(min_row=2, values_only=True):
            strike_ce = row[10]   # col K(11) = CE strikePrice
            ce_oi     = row[7]    # col H(8)  = CE openInterest
            strike_pe = row[26]   # col AA(27)= PE strikePrice
            pe_oi     = row[24]   # col Y(25) = PE openInterest
            und_val   = row[13]   # col N(14) = underlyingValue
            updated   = row[35]   # col AJ(36)= LastUpdated (fetch ka waqt)
            nse_ts_v  = row[36] if len(row) > 36 else None   # col AK(37)= NSE_Timestamp (NSE ka asli data-time)
            exp_date  = row[0]    # col A(1)  = expiryDates
            try:
                tot_ce_coi += int(row[3] or 0)    # col D  = CE changeinOpenInterest
                tot_pe_coi += int(row[20] or 0)   # col U  = PE changeinOpenInterest
            except Exception:
                pass

            if strike_ce and ce_oi is not None:
                ce_oi_by_strike[int(strike_ce)] = int(ce_oi) if ce_oi else 0

            if strike_pe and pe_oi is not None:
                pe_oi_by_strike[int(strike_pe)] = int(pe_oi) if pe_oi else 0

            if und_val and underlying is None:
                underlying = float(und_val)

            if updated and last_updated is None:
                if isinstance(updated, datetime.datetime):
                    last_updated = updated
                else:
                    try:
                        last_updated = datetime.datetime.strptime(str(updated), "%Y-%m-%d %H:%M:%S.%f")
                    except:
                        last_updated = datetime.datetime.now()

            if nse_ts is None and nse_ts_v:
                nse_ts = _parse_nse_ts(nse_ts_v)

            if exp_date and expiry is None:
                if isinstance(exp_date, datetime.datetime):
                    expiry = exp_date.strftime("%d-%b-%Y")
                else:
                    expiry = str(exp_date)

        wb.close()

        if not ce_oi_by_strike:
            log.info("[EXCEL] X No data found in sheet")
            return None

        # [FIX H1] freshness / dedupe / chart-time sab NSE ke timestamp par. Parse na ho (N/A) to purana tareeka (fetch time).
        if nse_ts is not None:
            last_updated = nse_ts
        else:
            log.info("[EXCEL] NSE_Timestamp parse nahi hua - LastUpdated (fetch time) use ho raha hai")

        total_ce = sum(ce_oi_by_strike.values())
        total_pe = sum(pe_oi_by_strike.values())
        pcr      = total_pe / max(total_ce, 1)

        log.info(f"[EXCEL] OK Read {len(ce_oi_by_strike)} strikes | "
              f"CE={total_ce:,} | PE={total_pe:,} | PCR={pcr:.3f} | "
              f"Spot={underlying:,.2f} | Updated={last_updated}")

        # MONTHLY sheet (optional - Power Query wali purani Excel me ye sheet nahi hoti to monthly text bas nahi aata)
        monthly = None
        try:
            if "MONTHLY" in wb.sheetnames:
                _r = list(wb["MONTHLY"].iter_rows(min_row=2, max_row=2, values_only=True))
                if _r and _r[0][1] is not None and _r[0][2] is not None:
                    monthly = {"expiry": str(_r[0][0]), "ce_oi": int(_r[0][1]), "pe_oi": int(_r[0][2])}
        except Exception as _e:
            log.info(f"[EXCEL] monthly sheet read error: {_e}")

        return {
            "monthly"         : monthly,
            "timestamp"       : datetime.datetime.now(),
            "excel_updated"   : last_updated,
            "expiry"          : expiry,
            "underlying"      : underlying,
            "total_ce_oi"     : total_ce,
            "total_pe_oi"     : total_pe,
            "total_ce_coi"    : tot_ce_coi,
            "total_pe_coi"    : tot_pe_coi,
            "ce_oi_by_strike" : ce_oi_by_strike,
            "pe_oi_by_strike" : pe_oi_by_strike,
        }

    except Exception as e:
        log.info(f"[EXCEL] X Error reading file: {e}")
        return None

# -- CALCULATIONS -----------------------------------------------------------
def compute_coi(history: list):
    coi_ce, coi_pe = [], []
    prev_ce = prev_pe = None
    for s in history:
        ce, pe = s["total_ce_oi"], s["total_pe_oi"]
        if prev_ce is None:
            coi_ce.append(0); coi_pe.append(0)
        else:
            coi_ce.append(ce - prev_ce); coi_pe.append(pe - prev_pe)
        prev_ce, prev_pe = ce, pe
    return list(pd.Series(coi_ce).cumsum()), list(pd.Series(coi_pe).cumsum())

def compute_maxpain(ce_oi: dict, pe_oi: dict) -> int:
    all_strikes = sorted(set(list(ce_oi) + list(pe_oi)))
    if not all_strikes:
        return 0
    pain = {}
    for s in all_strikes:
        loss  = sum((s - k) * v for k, v in ce_oi.items() if s > k)
        loss += sum((k - s) * v for k, v in pe_oi.items() if s < k)
        pain[s] = loss
    return min(pain, key=pain.get)

# -- TELEGRAM ---------------------------------------------------------------
# ===== PATCH_BIAS_V13 : PCR + OI-flow + price + MaxPain BIAS engine =====
LOT_SIZE = 65          # Nifty lot size (sirf Cr/L conversion ke liye)
# -- OI ko Cr / L me dikhane ka formatter (sirf display) ----------------------
def fmt_oi(v, dec=2, trim=False):
    """Sensibull jaisa: 14.42Cr / 23.04Cr / -72.22L / -1.24L  (v = NSE ke lots; andar lot size se multiply)."""
    q = float(v) * LOT_SIZE      # aapka LOT_SIZE = 65 (line ~225)
    a = abs(q)
    if a >= 1e7:   num, unit = a / 1e7, "Cr"
    elif a >= 1e5: num, unit = a / 1e5, "L"
    elif a >= 1e3: num, unit = a / 1e3, "K"
    else:          num, unit = a, ""
    txt = f"{num:.{dec}f}" if unit else f"{num:.0f}"
    if trim and "." in txt:
        txt = txt.rstrip("0").rstrip(".")
    return ("-" if q < 0 and float(txt or 0) != 0 else "") + txt + unit

BIAS_CFG = {
    "look_min"    : 10.0,   # OI flow window (minute): 5 / 10 / 15 - bars bhi isi window ke bante hain   PATCH_BIAS_WINDOW
    "max_gap_min" : 25.0,   # isse bada gap (script band tha) to flow skip
    "w"           : {"flow": 0.25, "pcr_tr": 0.20, "pcr_rel": 0.10, "price": 0.15, "mp": 0.10, "xover": 0.20},   # PATCH_XOVER: PCR vs Underlying (z-score) crossover
    "ema"         : 0.5,    # smoothing (chhota = zyada smooth)
    "on"          : 0.25,   # |score| isse upar => BULLISH / BEARISH
    "off"         : 0.10,   # |score| isse neeche => wapas NEUTRAL (hysteresis)
    "strong"      : 0.60,   # |score| isse upar => STRONG
    "min_hist"    : 8,      # itne points ke baad hi z-score
    "flat_oi_pct" : 0.2,    # OI change avg OI ka itna % se kam => Flat
    "flat_px"     : 5.0,    # price move itne points se kam => sideways
    "xo_eps"      : 0.08,   # PATCH_XOVER: PCR-UND farq (z-score/2) isse kam => "flat", jhootha crossover nahi
}
_WARN = {"holiday_day": None, "stale": False}   # [FIX H1] ek-baar-wali Telegram warnings ka state
LAST_BIAS = {}   # render_chart yahan latest bias rakhta hai, build_caption yahin se padhta hai

def _snap_min(s):
    t = s.get("excel_updated") or s["timestamp"]
    return t.timestamp() / 60.0

def _flow_tags(dce, dpe, dpx, oi_avg):
    """OI change + price direction se buildup tag (premium ka data nahi, isliye ye andaza hai)."""
    c   = BIAS_CFG
    thr = oi_avg * c["flat_oi_pct"] / 100.0
    up, dn = dpx > c["flat_px"], dpx < -c["flat_px"]
    if dce < -thr:
        ce = "Short Covering (Bull)" if up else ("Long Unwinding (Bear)" if dn else "Unwinding (Neutral)")
    elif dce > thr:
        ce = "Build-up (Mixed)" if up else "Writing (Bear)"
    else:
        ce = "Flat"
    if dpe > thr:
        pe = "Writing (Bull)" if up else ("Buying (Bear)" if dn else "Writing (Support)")
    elif dpe < -thr:
        pe = "Unwinding (Bull)" if up else ("Unwinding (Bear)" if dn else "Unwinding (Neutral)")
    else:
        pe = "Flat"
    return ce, pe

def compute_bias(hist, mp=None):
    """hist = sirf aaj ke snapshots. Return: ema(list), labels, bars[(i,dCE,dPE)], last{}"""
    c = BIAS_CFG
    n = len(hist)
    out = {"ema": [], "raw": [], "labels": [], "bars": [], "last": {}}
    if n < 3:
        return out
    t   = np.array([_snap_min(s) for s in hist], float)
    ce  = np.array([s["total_ce_oi"] for s in hist], float)
    pe  = np.array([s["total_pe_oi"] for s in hist], float)
    und = np.array([s["underlying"] for s in hist], float)
    pcr = pe / np.maximum(ce, 1.0)

    def back(i):   # look_min ke sabse nazdeek snapshot (data 3-min ke gap par aata hai)
        best = None
        for k in range(i - 1, -1, -1):
            d = t[i] - t[k]
            if d > max(c["look_min"] * 1.5, c["look_min"] + 4.0):
                break
            if best is None or abs(d - c["look_min"]) < abs(t[i] - t[best] - c["look_min"]):
                best = k
        if best is None or t[i] - t[best] < c["look_min"] * 0.5:
            return None
        return best if t[i] - t[best] <= c["max_gap_min"] else None

    feats = {k: [None] * n for k in c["w"]}
    for i in range(n):
        j = back(i)
        if j is not None:
            feats["flow"][i]   = ((pe[i] - pe[j]) - (ce[i] - ce[j])) / ((ce[i] + pe[i]) / 2.0)
            feats["pcr_tr"][i] = pcr[i] - pcr[j]
            feats["price"][i]  = und[i] - und[j]
        feats["pcr_rel"][i] = pcr[i] - pcr[:i + 1].mean()
        if mp is not None and mp[i] == mp[i]:
            feats["mp"][i] = float(np.clip((mp[i] - und[i]) / 50.0, -1, 1))
        # PATCH_XOVER: PCR aur Underlying ko aaj ke din ki apni mean/std se same scale par laao;
        # PCR-line Underlying-line ke upar => bullish (+), neeche => bearish (-). (sirf ab tak ka data, future nahi)
        if i + 1 >= c["min_hist"]:
            sp, su = float(np.std(pcr[:i + 1])), float(np.std(und[:i + 1]))
            if sp > 1e-12 and su > 1e-12:
                zp = (pcr[i] - float(np.mean(pcr[:i + 1]))) / sp
                zu = (und[i] - float(np.mean(und[:i + 1]))) / su
                feats["xover"][i] = float(np.clip((zp - zu) / 2.0, -1, 1))

    raw = np.zeros(n)
    for i in range(n):
        num = den = 0.0
        for k, w in c["w"].items():
            v = feats[k][i]
            if v is None:
                continue
            if k in ("mp", "xover"):   # PATCH_XOVER
                z = v
            else:
                hh = [x for x in feats[k][:i + 1] if x is not None]
                if len(hh) < c["min_hist"]:
                    continue
                sd = float(np.std(hh))
                if sd <= 1e-12:
                    continue
                z = float(np.clip(v / sd / 2.0, -1, 1))
            num += w * z
            den += w
        raw[i] = num / den if den else 0.0

    ema = np.zeros(n)
    lab = [""] * n
    st  = "NEUTRAL"
    for i in range(n):
        ema[i] = raw[i] if i == 0 else c["ema"] * raw[i] + (1 - c["ema"]) * ema[i - 1]
        e = ema[i]
        if st == "NEUTRAL":
            st = "BULLISH" if e > c["on"] else ("BEARISH" if e < -c["on"] else "NEUTRAL")
        elif st == "BULLISH" and e < c["off"]:
            st = "NEUTRAL"
        elif st == "BEARISH" and e > -c["off"]:
            st = "NEUTRAL"
        lab[i] = st if st == "NEUTRAL" else (("STRONG " if abs(e) > c["strong"] else "") + st)

    bars, i = [], n - 1          # peeche se non-overlapping look_min-minute buckets (abhi 10 min)
    while i > 0:
        j = back(i)
        if j is None:
            i -= 1
            continue
        bars.append((i, int(ce[i] - ce[j]), int(pe[i] - pe[j])))
        i = j
    bars.reverse()

    j = back(n - 1)
    if j is None:
        j = max(0, n - 6)
    dce, dpe, dpx = int(ce[-1] - ce[j]), int(pe[-1] - pe[j]), float(und[-1] - und[j])
    ce_tag, pe_tag = _flow_tags(dce, dpe, dpx, (ce[-1] + pe[-1]) / 2.0)
    xo_now = feats["xover"][-1]
    xo_tag = "n/a"
    if xo_now is not None:
        eps = c["xo_eps"]
        sides = [0 if v is None or abs(v) <= eps else (1 if v > 0 else -1) for v in feats["xover"]]
        cur = sides[-1]
        if cur == 0:
            xo_tag = "PCR ~ UND (Flat)"
        else:
            xo_tag = "PCR > UND (Bull)" if cur > 0 else "PCR < UND (Bear)"
            earlier = [x for x in sides[:-1] if x != 0]
            # asli crossover: pichhla stable side ulta tha aur abhi (is snapshot par) naya side bana
            if earlier and earlier[-1] != cur and (n < 2 or sides[-2] != cur):
                xo_tag = "CROSSOVER! " + xo_tag
    out.update(ema=list(ema), raw=list(raw), labels=lab, bars=bars)
    out["last"] = {"label": lab[-1], "score": float(ema[-1]), "dce": dce, "dpe": dpe,
                   "dpx": dpx, "ce_tag": ce_tag, "pe_tag": pe_tag, "xo_tag": xo_tag}
    return out

def _safe_bias(hist, mp):
    try:
        return compute_bias(hist, mp)
    except Exception as e:
        log.info(f"[BIAS] error: {e}")
        return {"ema": [], "raw": [], "labels": [], "bars": [], "last": {}}

def _bias_text(pcr):
    b = LAST_BIAS
    if b.get("label"):
        return f' {b["label"]} ({b["score"]:+.2f})'
    return " BULLISH" if pcr > 1.0 else " BEARISH"      # fallback (history kam ho)

def _fmt_tag(tag, side):
    """Caption me plain 'Unwinding (...)' ko saaf bhasha me: 'Short Unwinding, PUT BUY KARO' (CE side par CALL BUY KARO).
    Short Covering bhi: 'Short Covering, CALL BUY KARO'. Baaki tags (Long Unwinding, Writing, Build-up...) jaise hain waise. Chart ke andar ka text nahi badlata."""
    act = 'CALL' if side == 'CE' else 'PUT'
    if str(tag).startswith("Unwinding"):
        return f"Short Unwinding, {act} BUY KARO"
    if str(tag).startswith("Short Covering"):      # short covering = buying; CE side par aata hai => CALL BUY KARO
        return f"Short Covering, {act} BUY KARO"
    return tag

def _bias_extra():
    b = LAST_BIAS
    if not b:
        return ""
    # Flow line hata di. Build-up line: BULLISH => sirf CE side, BEARISH => sirf PE side (NEUTRAL => dono)
    label = str(b.get("label", "")).upper()
    if "BULL" in label:
        build = f"CE: {_fmt_tag(b['ce_tag'], 'CE')}"
    elif "BEAR" in label:
        build = f"PE: {_fmt_tag(b['pe_tag'], 'PE')}"
    else:
        build = f"CE: {_fmt_tag(b['ce_tag'], 'CE')} | PE: {_fmt_tag(b['pe_tag'], 'PE')}"
    return (f"{build}\n"
            f"X-over: {b.get('xo_tag', 'n/a')}\n")

def build_caption(snap: dict, pcr: float, maxpain: int) -> str:
    diff = snap["total_pe_oi"] - snap["total_ce_oi"]
    bias = _bias_text(pcr)   # PATCH_BIAS_V13   # PATCH_BIAS_V8
    # PATCH_FIXSCALE_V10: caption - har item apni line mein, koi gap nahi (MP = MaxPain)
    return (
        f"DIFF: `{diff:,}`\n"
        f"Bias: {bias.strip()}\n"
        f"{_bias_extra()}"
        f"PCR: `{pcr:.3f}`\n"
        f"MP: `{maxpain:,}`\n"
        f"Spot: `{snap['underlying']:,.2f}`"
    )

def build_alert_caption(snap: dict, pcr: float, maxpain: int):
    """Caption SIRF tab jab PCR aur Underlying ka CROSSOVER ho; warna None (khali image jaati hai)."""
    if str(LAST_BIAS.get("xo_tag", "")).startswith("CROSSOVER!"):
        return build_caption(snap, pcr, maxpain)
    return None

def send_telegram_photo(image_path: str, caption: str) -> bool:
    url = f"https://api.telegram.org/bot{TELEGRAM['bot_token']}/sendPhoto"
    try:
        with open(image_path, 'rb') as photo:
            resp = requests.post(url, data={
                "chat_id"    : TELEGRAM["chat_id"],
                "caption"    : caption,
                "parse_mode" : "Markdown",
            }, files={"photo": photo}, timeout=30)
        if resp.status_code == 200:
            log.info("[TG] OK Chart sent!")
            return True
        else:
            log.info(f"[TG] X {resp.status_code}: {resp.text[:200]}")
            return False
    except Exception as e:
        log.info(f"[TG] X Error: {e}")
        return False

def send_telegram_text(message: str) -> bool:
    url = f"https://api.telegram.org/bot{TELEGRAM['bot_token']}/sendMessage"
    try:
        resp = requests.post(url, data={
            "chat_id"    : TELEGRAM["chat_id"],
            "text"       : message,
            "parse_mode" : "Markdown",
        }, timeout=15)
        return resp.status_code == 200
    except Exception as e:
        log.info(f"[TG] X {e}")
        return False

# -- CHART ------------------------------------------------------------------
def _style(ax, ax2=None):
    ax.set_facecolor(PANEL_BG)
    ax.tick_params(colors=TEXT_CLR, labelsize=6.5, length=3)
    for sp in ax.spines.values(): sp.set_color(GRID_CLR)
    ax.grid(True, color=GRID_CLR, linewidth=0.4, alpha=0.8)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
    ax.xaxis.set_major_locator(mdates.MinuteLocator(byminute=range(0, 60, 3)))
    plt.setp(ax.get_xticklabels(), rotation=90, ha='center', fontsize=6)
    if ax2:
        ax2.set_facecolor(PANEL_BG)
        ax2.tick_params(colors=DIFF_CLR, labelsize=6.5, length=3)
        for sp in ax2.spines.values(): sp.set_color(GRID_CLR)

def _infobox(ax, lines, x=0.68, y=0.88):
    ax.text(x, y, '\n'.join(lines),
            transform=ax.transAxes,
            fontsize=12, fontfamily='monospace', color=TEXT_CLR,
            va='top', ha='left', fontweight='bold',
            linespacing=2.0, zorder=5)

def _lleg(ax, handles):
    ax.legend(handles=handles, loc='upper left',
              framealpha=0.9, facecolor='#f5f5f5', edgecolor=GRID_CLR,
              labelcolor=TEXT_CLR, fontsize=7.5, handlelength=1.8)

def _rleg(ax, handles):
    ax.legend(handles=handles, loc='upper right',
              framealpha=0.9, facecolor='#f5f5f5', edgecolor=GRID_CLR,
              labelcolor=DIFF_CLR, fontsize=7.5, handlelength=1.8)

# -- HISTORY SAVE/LOAD (restart pe chart reset nahi hoga) ---------------------
import pickle, math

def _hist_path(day):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        f"oi_history_{day.strftime('%Y%m%d')}.pkl")

def load_history(day):
    path = _hist_path(day)
    try:
        with open(path, 'rb') as f:
            h = pickle.load(f)
        return h if isinstance(h, list) else []
    except FileNotFoundError:
        return []
    except Exception as e:   # [FIX H2] corrupt file ab chup-chap nahi - log + side me rakh do
        log.info(f"[HIST] load error ({os.path.basename(path)}): {e}")
        try:
            os.replace(path, path + ".corrupt")
        except Exception:
            pass
        return []

def save_history(day, history):
    """[FIX H2] ATOMIC save: pehle tmp file, fir os.replace - beech me kat bhi jaye to purani pkl salamat."""
    path = _hist_path(day)
    tmp  = path + ".tmp"
    try:
        with open(tmp, 'wb') as f:
            pickle.dump(history, f)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    except Exception as e:
        log.info(f"[HIST] save error: {e}")
        try:
            os.remove(tmp)
        except Exception:
            pass

_XT = [0.705]

def _endtext(ax, lines, y0=0.60, gap=0.13):
    """Line ke end ke baad (right side) coloured values."""
    for i, (txt, clr) in enumerate(lines):
        ax.text(_XT[0], y0 - i * gap, txt, transform=ax.transAxes,
                fontsize=14, fontfamily='monospace', fontweight='bold',
                color=clr, va='center', ha='left', zorder=6)

# -- CHART (v3: mBOT style) ---------------------------------------------------
# PATCH_ABC_V7 : (A) zero line DIFF axis pe  (B) MaxPain din bhar ka history  (C) pichhle din ka tail + vertical line
PREV_DAY_TOTAL_POINTS = 80   # subah chart mein kul itne points (pichhla din + aaj); aaj ke points badhne par tail ghat'ta jata hai

def _maxpain_fast(ce_oi: dict, pe_oi: dict):
    """compute_maxpain jaisa hi result, bas tez (numpy)."""
    strikes = sorted(set(list(ce_oi) + list(pe_oi)))
    if not strikes:
        return float('nan')
    S  = np.array(strikes, dtype=float)
    kc = np.array(list(ce_oi.keys()), dtype=float);  vc = np.array(list(ce_oi.values()), dtype=float)
    kp = np.array(list(pe_oi.keys()), dtype=float);  vp = np.array(list(pe_oi.values()), dtype=float)
    loss = (np.clip(S[:, None] - kc[None, :], 0, None) * vc[None, :]).sum(axis=1) \
         + (np.clip(kp[None, :] - S[:, None], 0, None) * vp[None, :]).sum(axis=1)
    return int(S[int(np.argmin(loss))])

def _prev_day_tail(history: list, want: int) -> list:
    """Pichhle trading din ki saved history (oi_history_YYYYMMDD.pkl) ke aakhri `want` points.
    Sirf wahi jinka expiry aaj jaisa ho (alag expiry ka OI mix nahi karte)."""
    if want <= 0 or not history:
        return []
    try:
        day0 = history[0]["timestamp"].date()
        exp  = history[-1].get("expiry")
        for k in range(1, 8):
            ph = load_history(day0 - datetime.timedelta(days=k))
            if not ph:
                continue
            ph = [s for s in ph if s.get("expiry") == exp and s.get("underlying")
                  and s.get("total_ce_oi") and s.get("total_pe_oi") and _is_fresh(s)]
            return ph[-want:]
    except Exception as e:
        log.info(f"[CHART] prev-day tail error: {e}")
    return []

def _is_fresh(s: dict) -> bool:
    """False agar snapshot ka Excel-time us din ka nahi jis din ye read hua (yaani purana/stale data)."""
    try:
        e, t = s.get("excel_updated"), s.get("timestamp")
        if e is not None and t is not None and hasattr(e, "date") and hasattr(t, "date"):
            return e.date() == t.date()
    except Exception:
        pass
    return True   # PATCH_STALE_V9

# PATCH_FIXSCALE_V10: panel 3 ka fixed scale (yahan badal sakte ho)
PCR_FIXED_AXIS = (0.90, 1.30)   # PCR axis (neeche, upar)
UND_FIXED_HALF = 200            # Underlying axis = din ke pehle point ke aas-paas +- itne points

def _spike_free(vals, rel=0.15):
    """PATCH_AXIS_V14: sirf AKELE bad points ko scale se hatao, genuine trend/move ko nahi
    (purana +-20%-of-median filter trend wale din line ko chart se bahar kar deta tha).
    Point tab spike hai jab apne aas-paas ke 5 points ke median se rel% se zyada door ho. Aakhri point hamesha rakha jata hai."""
    v = np.array(vals, dtype=float)
    if len(v) < 5:
        return v
    pad = np.pad(v, 2, mode='edge')
    med = np.array([np.median(pad[i:i + 5]) for i in range(len(v))])
    keep = np.abs(v - med) <= rel * np.maximum(np.abs(med), 1e-9)
    keep[-1] = True
    return v[keep]

PANEL3_ZSCALE = True   # True = PCR aur Underlying same sigma scale (alert ke crossover jaisa) | False = purana alag-alag auto axis

BIAS_BAR_SCALE = 1.0   # PATCH_BIAS_V15: background bars ki oonchai ka multiplier (chhote lagein to 1.5 / 2 karo)

def render_chart(history: list, expiry_label: str, outfile: str) -> tuple:
    history = [s for s in history if _is_fresh(s)]   # PATCH_STALE_V9: purana-date data chart se hatao
    if len(history) < 2:
        log.info("[CHART] Need >=2 snapshots - skipping.")
        return None, None

    # PATCH_ABC_V7 (C): subah pichhle din ka tail jodo (aaj ke points badhne par tail ghat'ta hai)
    _tail  = _prev_day_tail(history, PREV_DAY_TOTAL_POINTS - len(history))
    ntail  = len(_tail)
    history = _tail + history

    n      = len(history)
    x      = list(range(n))
    labels = [((s.get("excel_updated") or s["timestamp"]).strftime("%H:%M")) for s in history]

    ce_oi = [s["total_ce_oi"] for s in history]
    pe_oi = [s["total_pe_oi"] for s in history]
    und   = [s["underlying"]  for s in history]
    pcr   = [p / max(c, 1) for c, p in zip(ce_oi, pe_oi)]

    if all("total_ce_coi" in s for s in history):
        coi_ce = [s["total_ce_coi"] for s in history]
        coi_pe = [s["total_pe_coi"] for s in history]
    else:
        coi_ce, coi_pe = compute_coi(history)

    diff_oi  = [p - c for c, p in zip(ce_oi,  pe_oi)]
    diff_coi = [p - c for c, p in zip(coi_ce, coi_pe)]

    # PATCH_ABC_V7 (B): har snapshot ka apna MaxPain (din bhar badalta hua step line)
    mp     = [_maxpain_fast(s["ce_oi_by_strike"], s["pe_oi_by_strike"]) for s in history]
    mp_val = mp[-1]
    _bias = _safe_bias(history[ntail:], mp[ntail:])   # PATCH_BIAS_V13
    LAST_BIAS.clear(); LAST_BIAS.update(_bias["last"])

    C_CE, C_PE, C_DF = '#1f4fd8', '#e0201b', '#111111'
    fmt_k = mticker.FuncFormatter(lambda v, _: fmt_oi(v, 2, True) if abs(v) > 1e-9 else '0')   # Sensibull style Cr/L

    # lines ka end axes ke 68% pe - baaki 32% mein values (mBOT jaisa)
    x0 = -0.5
    x1 = x0 + (max(n - 1, 40) - x0) / 0.68   # PATCH_CHART_V31
    _XT[0] = min((n - 1 - x0) / (x1 - x0) + 0.03, 0.705)
    step = 1   # har snapshot (3 min) pe label
    _lfs = min(9, max(5.5, 400 / n))   # zyada labels ho to font chhota

    with plt.style.context('ggplot'):
        fig = plt.figure(figsize=(15.65, 9.48), dpi=125)   # PATCH_BIAS_V14: original 1956x1185, 3 panel
        gs  = fig.add_gridspec(3, 1, hspace=0.16, top=0.935, bottom=0.085,
                               left=0.062, right=0.925)
        ax1 = fig.add_subplot(gs[0])
        ax2 = fig.add_subplot(gs[1], sharex=ax1)
        ax3 = fig.add_subplot(gs[2], sharex=ax1)

        sync = datetime.datetime.now().strftime('%d-%b-%Y  %H:%M:%S')
        fig.suptitle(f"NIFTY EXPIRY  {expiry_label}      @ {sync}  (data {labels[-1]})   |   sync: 3min",
                     fontsize=15, fontweight='bold', fontfamily='monospace', y=0.985)

        # Panel 1 : Total OI
        r1 = ax1.twinx(); r1.grid(False)
        a, = ax1.plot(x, ce_oi, color=C_CE, lw=2.0, label='TOTAL CE')
        b, = ax1.plot(x, pe_oi, color=C_PE, lw=2.0, label='TOTAL PE')
        d, = r1.plot(x, diff_oi, color=C_DF, lw=1.5, ls='--', label='DIFF PE-CE')
        r1.axhline(0, color='#6b7a8a', lw=1.0)   # PATCH_ABC_V7 (A): DIFF axis ki zero line
        ax1.yaxis.set_major_formatter(fmt_k); r1.yaxis.set_major_formatter(fmt_k)
        ax1.legend(handles=[a, b], loc='upper left', fontsize=10)
        r1.legend(handles=[d], loc='upper right', fontsize=10)
        _endtext(ax1, [(f"CE OI  {fmt_oi(ce_oi[-1])}", C_CE),
                       (f"PE OI  {fmt_oi(pe_oi[-1])}", C_PE),
                       (f"PE-CE  {fmt_oi(diff_oi[-1])}", C_DF)])
        # MONTHLY expiry ka total CE OI / PE OI / DIFF - panel 1 ke right-top (blue mark wali jagah), chhote letters
        _mo = history[-1].get("monthly")
        if _mo:
            _mce, _mpe = _mo["ce_oi"], _mo["pe_oi"]
            for _y, _txt, _clr, _fs in [(0.965, f"MONTHLY {_mo['expiry']}", '#555555', 7.5),
                                        (0.885, f"CE OI  {fmt_oi(_mce)}", C_CE, 10),
                                        (0.815, f"PE OI  {fmt_oi(_mpe)}", C_PE, 10),
                                        (0.745, f"DIFF   {fmt_oi(_mpe - _mce)}", C_DF, 10)]:
                ax1.text(_XT[0], _y, _txt, transform=ax1.transAxes, fontsize=_fs, fontfamily='monospace',
                         fontweight='bold', color=_clr, va='center', ha='left', zorder=6)

        # Panel 2 : COI
        r2 = ax2.twinx(); r2.grid(False)
        a, = ax2.plot(x, coi_ce, color=C_CE, lw=2.0, label='TOTAL COI CE')
        b, = ax2.plot(x, coi_pe, color=C_PE, lw=2.0, label='TOTAL COI PE')
        _leg2 = [a, b]   # PATCH_BIAS_V15
        d, = r2.plot(x, diff_coi, color=C_DF, lw=1.5, ls='--', label='DIFF PE-CE')
        r2.axhline(0, color='#6b7a8a', lw=1.0)   # PATCH_ABC_V7 (A): DIFF axis ki zero line
        ax2.yaxis.set_major_formatter(fmt_k); r2.yaxis.set_major_formatter(fmt_k)
        ax2.legend(handles=[a, b], loc='upper left', fontsize=10)
        r2.legend(handles=[d], loc='upper right', fontsize=10)
        _endtext(ax2, [(f"COI CE  {fmt_oi(coi_ce[-1])}", C_CE),
                       (f"COI PE  {fmt_oi(coi_pe[-1])}", C_PE),
                       (f"PE-CE   {fmt_oi(diff_coi[-1])}", C_DF)])

        # Panel 3 : PCR + MaxPain + Underlying
        r3 = ax3.twinx(); r3.grid(False)
        a, = ax3.plot(x, pcr, color=C_CE, lw=2.0, label='PCR')
        b, = r3.plot(x, mp,  color=C_PE, lw=2.0, label='MaxPain')
        d, = r3.plot(x, und, color=C_DF, lw=1.5, ls='--', label='Underlying')
        # PCR aur Underlying apni apni range mein fit (Sensibull jaisa); PCR ka bad point scale kharab nahi karega
        ax3.yaxis.set_major_locator(mticker.MaxNLocator(nbins=8))
        _pk  = _spike_free(pcr)   # PATCH_AXIS_V14: sirf akela bad point scale se bahar
        _plo, _phi = float(_pk.min()), float(_pk.max())
        _ppad = max((_phi - _plo) * 0.15, 0.005)
        ax3.set_ylim(_plo - _ppad, _phi + _ppad)
        _rng = (_phi - _plo) or 0.01
        _dec = 4 if _rng < 0.01 else (3 if _rng < 0.1 else 2)
        ax3.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.' + str(_dec) + 'f'))
        # right axis: Underlying + wo MaxPain jo paas ho (door wala MaxPain scale kharab na kare)
        _ulo, _uhi = min(und), max(und)
        _ur  = max(_uhi - _ulo, 5.0)
        _mv  = [m for m in mp if m == m and _ulo - 0.75 * _ur <= m <= _uhi + 0.75 * _ur]
        _lo, _hi = min([_ulo] + _mv), max([_uhi] + _mv)
        _upad = max((_hi - _lo) * 0.15, 5.0)
        r3.set_ylim(_lo - _upad, _hi + _upad)
        # PATCH_FIXSCALE_V10: FIXED scale - upar wala auto-fit override. Tail (pichhla din) scale ko nahi chhedta.
        # PATCH_PCR_AUTO_V11: PCR axis AUTO-FIT (sirf aaj ke points, bad outlier ignore)
        _pt = _spike_free(pcr)   # PATCH_AXIS_V14: trend wala din line ko chart se bahar nahi karega
        if len(_pt):
            _lo_p, _hi_p = float(_pt.min()), float(_pt.max())
            _rng_p = max(_hi_p - _lo_p, 0.04)
            _mid_p = (_hi_p + _lo_p) / 2.0
            _pad_p = _rng_p * 0.12
            ax3.set_ylim(_mid_p - _rng_p / 2 - _pad_p, _mid_p + _rng_p / 2 + _pad_p)
        ax3.yaxis.set_major_locator(mticker.MaxNLocator(nbins=8))
        ax3.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.2f'))
        # PATCH_ALL_AUTO_V12: Underlying + MaxPain axis AUTO-FIT (sirf aaj ke points)
        _vals = [v for v in (list(und) + list(mp)) if v == v]
        if _vals:
            _vlo, _vhi = float(min(_vals)), float(max(_vals))
            _vr  = max(_vhi - _vlo, 40.0)
            _vm  = (_vhi + _vlo) / 2.0
            _vp  = _vr * 0.12
            r3.set_ylim(_vm - _vr / 2 - _vp, _vm + _vr / 2 + _vp)
        # [Z-AXIS] PCR aur Underlying dono axes SAME sigma scale par (+-K sigma, aaj ke data ki mean/std) -
        # yehi z-score alert ke CROSSOVER me use hota hai, isliye lines chart par wahin cross karengi jahan alert aata hai.
        # Auto min-max se har din/waqt relative position badal jaati thi.
        try:
            if not PANEL3_ZSCALE:
                raise RuntimeError('PANEL3_ZSCALE=False - auto axis')
            _sl = slice(None)   # hamesha jo chart me dikh raha hai (tail + aaj) usi se scale; sirf-aaj par zoom karne se kal ka hissa gayab ho jata tha
            _pa = np.array(pcr[_sl], dtype=float)
            _ua = np.array(und[_sl], dtype=float)
            _sdp, _sdu = float(np.std(_pa)), float(np.std(_ua))
            if _sdp > 1e-9 and _sdu > 1e-9:
                _mup, _muu = float(np.mean(_pa)), float(np.mean(_ua))
                _zp = (_spike_free(_pa) - _mup) / _sdp
                _zu = (_ua - _muu) / _sdu
                _zlo = min(float(_zp.min()), float(_zu.min()))
                _zhi = max(float(_zp.max()), float(_zu.max()))
                _sp0 = max(_zhi - _zlo, 0.04 / _sdp, 40.0 / _sdu)      # bahut chhoti range na bane
                _mid = (_zhi + _zlo) / 2.0
                _zlo, _zhi = _mid - _sp0 / 2.0, _mid + _sp0 / 2.0
                # MaxPain sirf tab scale badhaye jab paas ho (door wala MaxPain lines ko dabaye nahi)
                _zm = [(m - _muu) / _sdu for m in mp[_sl] if m == m]
                _zm = [z for z in _zm if _zlo - 0.5 * _sp0 <= z <= _zhi + 0.5 * _sp0]
                if _zm:
                    _zlo, _zhi = min(_zlo, min(_zm)), max(_zhi, max(_zm))
                _pd = (_zhi - _zlo) * 0.10
                _zlo, _zhi = _zlo - _pd, _zhi + _pd
                ax3.set_ylim(_mup + _zlo * _sdp, _mup + _zhi * _sdp)
                r3.set_ylim(_muu + _zlo * _sdu, _muu + _zhi * _sdu)
        except Exception as e:
            if PANEL3_ZSCALE:
                log.info(f"[CHART] z-axis error (auto axis par wapas): {e}")
        r3.yaxis.set_major_locator(mticker.MaxNLocator(nbins=8))
        r3.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'{int(v):,}'))   # MaxPain/Underlying = index points, Cr/L nahi
        ax3.legend(handles=[a], loc='upper left', fontsize=10)
        r3.legend(handles=[b, d], loc='upper right', fontsize=10)
        _endtext(ax3, [(f"PCR         {pcr[-1]:.2f}", C_CE),
                       (f"MAXPAIN     {mp_val:,}", C_PE),
                       (f"UNDERLYING  {und[-1]:,.2f}", C_DF)])

        # PATCH_BIAS_V15: Sensibull jaisa - CE/PE OI change bars panel 2 (COI) ke BACKGROUND mein,
        # baseline wahi zero line (DIFF axis ki) jo pehle se hai. Bias line bhi usi zero ke around.
        _e4 = list(_bias["ema"])
        if _e4 and _bias["last"]:
            _bl = _bias["last"]
            _wm = int(BIAS_CFG["look_min"])
            _l2, _r2l = ax2.get_ylim(), r2.get_ylim()
            ax2.set_ylim(*_l2); r2.set_ylim(*_r2l)          # limits freeze (bars scale na badlein)
            _k2 = (_l2[1] - _l2[0]) / (_r2l[1] - _r2l[0])    # DIFF-axis units -> left-axis units
            _z0 = _l2[0] + (0 - _r2l[0]) * _k2               # DIFF axis ka zero, left axis ke coordinates mein
            _bars = _bias["bars"]
            _mx = max([abs(b_[1]) for b_ in _bars] + [abs(b_[2]) for b_ in _bars] + [1]) if _bars else 1
            if _bars:
                _bx = [ntail + b_[0] for b_ in _bars]
                _sp = min([q[0] - p_[0] for p_, q in zip(_bars, _bars[1:])] or [5])
                _bw = max(0.5, min(1.4, _sp * 0.45))
                ax2.bar([v - _bw / 2 for v in _bx], [b_[1] * BIAS_BAR_SCALE * _k2 for b_ in _bars], bottom=_z0,
                        width=_bw, color='#1e8e3e', alpha=0.55, lw=0, zorder=1.5)   # CE flow bars ab GREEN (line/text blue hi hain)
                ax2.bar([v + _bw / 2 for v in _bx], [b_[2] * BIAS_BAR_SCALE * _k2 for b_ in _bars], bottom=_z0,
                        width=_bw, color=C_PE, alpha=0.42, lw=0, zorder=1.5)
            # BIAS score wali purple dotted line panel 2 se hata di (bars aur BIAS text waise hi hain)
            ax2.legend(handles=_leg2, loc='upper left', fontsize=10)
            _bc = '#1e8e3e' if 'BULL' in _bl["label"] else ('#e65100' if 'BEAR' in _bl["label"] else '#555555')
            def _t(y, txt, clr, fs):
                ax2.text(_XT[0], y, txt, transform=ax2.transAxes, fontsize=fs, fontfamily='monospace',
                         fontweight='bold', color=clr, va='center', ha='left', zorder=6)
            _t(0.935, f"BIAS  {_bl['label']}  {_bl['score']:+.2f}", _bc, 12)
            _t(0.845, f"CE {_wm}m {('+' if _bl['dce'] > 0 else '') + fmt_oi(_bl['dce'])}  {_bl['ce_tag']}", C_CE, 8.5)
            _t(0.765, f"PE {_wm}m {('+' if _bl['dpe'] > 0 else '') + fmt_oi(_bl['dpe'])}  {_bl['pe_tag']}", C_PE, 8.5)
            _t(0.685, f"X-over  {_bl.get('xo_tag', 'n/a')}", '#444444', 8.5)

        ax1.set_xlim(x0, x1)
        if ntail > 0:   # PATCH_ABC_V7 (C): naya din shuru hone ki vertical line
            for _ax in (ax1, ax2, ax3):
                _ax.axvline(ntail, color='#6b7a8a', lw=1.0, zorder=1)
        for ax in (ax1, ax2):
            plt.setp(ax.get_xticklabels(), visible=False)
        ax3.set_xticks(x[::step])
        ax3.set_xticklabels(labels[::step], rotation=90, fontsize=_lfs)

        fig.savefig(outfile, dpi=125, format='jpeg', facecolor=fig.get_facecolor(),
                    pil_kwargs={'quality': 92, 'optimize': True})
        plt.close(fig)

    log.info(f"[CHART] OK Saved -> {outfile}")
    return pcr[-1], mp_val


# -- MARKET HOURS -----------------------------------------------------------
def is_market_open() -> bool:
    now = datetime.datetime.now()
    if now.weekday() >= 5: return False
    open_t  = now.replace(hour=9,  minute=15, second=0, microsecond=0)
    close_t = now.replace(hour=15, minute=40, second=0, microsecond=0)   # alerts 09:15 se 15:40 tak
    return open_t <= now <= close_t

# -- MAIN LOOP --------------------------------------------------------------
def main():
    print(f"\n[MAIN] NIFTY OI Tracker - Excel Mode")
    log.info(f"[MAIN] Excel file : {EXCEL_FILE}")
    log.info(f"[MAIN] Poll every : {POLL_INTERVAL}s\n")

    history              = []
    current_session_date = None
    last_excel_time      = None   # track if Excel actually updated

    send_telegram_text(
        " *NIFTY OI Tracker STARTED*\n"
        "Mode: `Excel sync`\n"
        f"Refresh: every `{POLL_INTERVAL // 60} min`"
    )

    while True:
        now = datetime.datetime.now()

        if not is_market_open():
            # 09:10-09:15 ke beech tez check, taaki pehla alert 09:15 par hi nikle
            _nap = 10 if (now.weekday() < 5 and now.hour == 9 and 10 <= now.minute < 15) else 60
            log.info(f"[MAIN] Market closed ({now.strftime('%H:%M')}) - sleeping {_nap}s...")
            time.sleep(_nap)
            continue

        # Fresh graph every new trading session
        today = now.date()
        if current_session_date != today:
            if current_session_date is not None:
                log.info(f"[MAIN]  New session ({today}) - clearing history!")
                send_telegram_text(f" *New trading session*\n`{today}` - Fresh chart started")
            current_session_date = today
            history = load_history(today)
            _WARN["holiday_day"] = None; _WARN["stale"] = False   # naya din - warnings reset
            last_excel_time = history[-1].get("excel_updated") if history else None
            log.info(f"[MAIN] Session: {today} | History loaded: {len(history)} snapshots")

        # Read Excel
        snap = read_excel_snapshot(EXCEL_FILE)
        if snap is None:
            log.info(f"[MAIN] Could not read Excel - retry in 30s")
            time.sleep(30)
            continue

        # Skip if Excel hasn't updated since last read
        excel_t = snap.get("excel_updated")
        # PATCH_STALE_V9: Excel ka data aaj ka nahi (kal ka purana) to skip, fresh data ka wait
        if excel_t is not None and hasattr(excel_t, "date") and excel_t.date() != today:
            log.info(f"[MAIN] Excel data purana ({excel_t}) - aaj ka data aane tak wait...")
            # [FIX H1] 09:30 tak bhi aaj ka NSE data nahi -> holiday / NSE issue; din me EK hi baar batao
            if _WARN["holiday_day"] != today and now.hour * 60 + now.minute >= 9 * 60 + 30:
                _WARN["holiday_day"] = today
                send_telegram_text(f"NSE data aaj ({today}) ka nahi aa raha (last data: `{excel_t:%d-%b %H:%M}`).\n"
                                   "Market holiday ho sakta hai ya NSE access issue. Aaj ka data aate hi chart chalu ho jayega.")
            time.sleep(30)
            continue
        # [FIX H1] ab NSE ke timestamp par: wahi/purana timestamp = naya data nahi aaya (duplicate point nahi banega)
        if excel_t and last_excel_time and excel_t <= last_excel_time:
            log.info(f"[MAIN] NSE data abhi naya nahi (last: {last_excel_time}) - waiting...")
            try:
                _age = (datetime.datetime.now() - excel_t).total_seconds() / 60
                _live = now.hour * 60 + now.minute <= 15 * 60 + 30     # 15:30 ke baad NSE data band hota hai, warning bekaar
                if _age > 10 and _live and not _WARN["stale"]:
                    _WARN["stale"] = True
                    send_telegram_text(f'WARNING: NSE data {_age:.0f} min se update nahi hua (NSE/fetch issue?)')
                elif _age <= 10:
                    _WARN["stale"] = False
            except Exception:
                pass
            time.sleep(30)
            continue
        last_excel_time = excel_t

        history.append(snap)
        save_history(today, history)
        expiry_label = snap.get("expiry", "N/A")

        # [FIX H3] ek kharab chart/caption poore tracker ko crash na kare; 3 baar lagataar fail ho to Telegram alert
        try:
            pcr_val, mp_val = render_chart(history, expiry_label, OUTPUT_FILE)
            if pcr_val is not None:
                caption = build_alert_caption(snap, pcr_val, mp_val)   # None => sirf image, crossover par caption
                send_telegram_photo(OUTPUT_FILE, caption)
            _WARN["render_fail"] = 0
        except Exception as e:
            import traceback
            _WARN["render_fail"] = _WARN.get("render_fail", 0) + 1
            log.info(f"[MAIN] chart/send error #{_WARN['render_fail']}: {e}\n{traceback.format_exc()}")
            if _WARN["render_fail"] == 3:
                send_telegram_text(f"WARNING: chart 3 baar lagataar fail hua: {e}")

        log.info(f"[MAIN] Next read in {POLL_INTERVAL}s...\n")
        time.sleep(POLL_INTERVAL)

# -- DEMO MODE --------------------------------------------------------------
def demo_mode():
    """Simulate data from Excel structure - no file needed."""
    log.info("[DEMO] Running demo - simulating Excel data...\n")
    np.random.seed(7)
    n = 125

    ce_list  = [6_500_000]; pe_list = [5_800_000]; und_list = [22100.0]
    for _ in range(n - 1):
        ce_list.append(max(6_000_000, ce_list[-1]  + int(np.random.normal(3000, 18000))))
        pe_list.append(max(5_000_000, pe_list[-1]  + int(np.random.normal(2000, 14000))))
        und_list.append(und_list[-1] + np.random.normal(0.5, 5.0))

    start   = datetime.datetime(2026, 10, 22, 9, 15)
    history = []
    for i in range(n):
        ts  = start + datetime.timedelta(minutes=i * 3)
        atm = round(und_list[i] / 50) * 50
        ce_by = {atm + j * 50: max(0, ce_list[i] // 30 + int(np.random.normal(0, 500)))
                 for j in range(-10, 11)}
        pe_by = {atm + j * 50: max(0, pe_list[i] // 30 + int(np.random.normal(0, 500)))
                 for j in range(-10, 11)}
        history.append({
            "timestamp"       : ts,
            "excel_updated"   : ts,
            "expiry"          : "27-Oct-2026",
            "underlying"      : und_list[i],
            "total_ce_oi"     : ce_list[i],
            "total_pe_oi"     : pe_list[i],
            "ce_oi_by_strike" : ce_by,
            "pe_oi_by_strike" : pe_by,
        })

    pcr_val, mp_val = render_chart(history, "27-Oct-2026", OUTPUT_FILE)
    if pcr_val:
        cap = build_caption(history[-1], pcr_val, mp_val)
        print("\n-- Caption Preview ---------------------------------")
        print(cap.replace('*', '').replace('`', ''))
        print("----------------------------------------------------")
        send_telegram_photo(OUTPUT_FILE, cap)
    print(f"\n[DEMO] OK Chart saved -> {OUTPUT_FILE}")

# -- ENTRY ------------------------------------------------------------------
if __name__ == "__main__":
    # -- Full error logging to file --------------------------------
    import traceback
    _log = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tracker_log.txt")

    def _write_log(msg):
        ts = datetime.datetime.now().strftime("[%d-%m-%Y %H:%M:%S]")
        line = f"{ts} {msg}\n"
        with open(_log, 'a', encoding='utf-8') as f:
            f.write(line)
        print(line, end='')

    try:
        _write_log("Python script STARTED")
        if "--demo" in sys.argv:
            _write_log("Mode: DEMO")
            demo_mode()
        else:
            _write_log("Mode: LIVE")
            main()
    except Exception as e:
        err = traceback.format_exc()
        _write_log(f"FATAL ERROR: {e}")
        _write_log(f"TRACEBACK:\n{err}")
        # Try to send error to Telegram
        try:
            url = f"https://api.telegram.org/bot{TELEGRAM['bot_token']}/sendMessage"
            requests.post(url, data={
                "chat_id": TELEGRAM["chat_id"],
                "text": f"ERROR OI Tracker CRASHED:\n`{e}`",
                "parse_mode": "Markdown"
            }, timeout=10)
        except:
            pass