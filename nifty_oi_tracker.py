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
            updated   = row[35]   # col AJ(36)= LastUpdated
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

            if exp_date and expiry is None:
                if isinstance(exp_date, datetime.datetime):
                    expiry = exp_date.strftime("%d-%b-%Y")
                else:
                    expiry = str(exp_date)

        wb.close()

        if not ce_oi_by_strike:
            log.info("[EXCEL] X No data found in sheet")
            return None

        total_ce = sum(ce_oi_by_strike.values())
        total_pe = sum(pe_oi_by_strike.values())
        pcr      = total_pe / max(total_ce, 1)

        log.info(f"[EXCEL] OK Read {len(ce_oi_by_strike)} strikes | "
              f"CE={total_ce:,} | PE={total_pe:,} | PCR={pcr:.3f} | "
              f"Spot={underlying:,.2f} | Updated={last_updated}")

        return {
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
def build_caption(snap: dict, pcr: float, maxpain: int) -> str:
    diff = snap["total_pe_oi"] - snap["total_ce_oi"]
    bias = " SUPER BULLISH / CAUTION" if pcr > 1.2 else (" BULLISH" if pcr > 0.8 else " BEARISH")   # PATCH_BIAS_V8
    # PATCH_FIXSCALE_V10: caption - har item apni line mein, koi gap nahi (MP = MaxPain)
    return (
        f"CE OI: `{snap['total_ce_oi']:,}`\n"
        f"PE OI: `{snap['total_pe_oi']:,}`\n"
        f"PE-CE: `{diff:,}`\n"
        f"Bias: {bias.strip()}\n"
        f"PCR: `{pcr:.3f}`\n"
        f"MP: `{maxpain:,}`\n"
        f"Spot: `{snap['underlying']:,.2f}`"
    )

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
    try:
        with open(_hist_path(day), 'rb') as f:
            h = pickle.load(f)
        return h if isinstance(h, list) else []
    except Exception:
        return []

def save_history(day, history):
    try:
        with open(_hist_path(day), 'wb') as f:
            pickle.dump(history, f)
    except Exception as e:
        log.info(f"[HIST] save error: {e}")

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

    C_CE, C_PE, C_DF = '#1f4fd8', '#e0201b', '#111111'
    fmt_k = mticker.FuncFormatter(lambda v, _: f'{int(v):,}')

    # lines ka end axes ke 68% pe - baaki 32% mein values (mBOT jaisa)
    x0 = -0.5
    x1 = x0 + (max(n - 1, 40) - x0) / 0.68   # PATCH_CHART_V31
    _XT[0] = min((n - 1 - x0) / (x1 - x0) + 0.03, 0.705)
    step = 1   # har snapshot (3 min) pe label
    _lfs = min(9, max(5.5, 400 / n))   # zyada labels ho to font chhota

    with plt.style.context('ggplot'):
        fig = plt.figure(figsize=(15.65, 9.48), dpi=125)
        gs  = fig.add_gridspec(3, 1, hspace=0.16, top=0.935, bottom=0.085,
                               left=0.062, right=0.925)
        ax1 = fig.add_subplot(gs[0])
        ax2 = fig.add_subplot(gs[1], sharex=ax1)
        ax3 = fig.add_subplot(gs[2], sharex=ax1)

        sync = datetime.datetime.now().strftime('%d-%b-%Y  %H:%M:%S')
        fig.suptitle(f"NIFTY EXPIRY  {expiry_label}      @ {sync}  (data {labels[-1]})   |   sync: 3min",
                     fontsize=15, fontweight='bold', fontfamily='monospace', y=0.985)
        fig.text(0.99, 0.99, '(c) mBOT', fontsize=8, color='#888888', ha='right', va='top')

        # Panel 1 : Total OI
        r1 = ax1.twinx(); r1.grid(False)
        a, = ax1.plot(x, ce_oi, color=C_CE, lw=2.0, label='TOTAL CE')
        b, = ax1.plot(x, pe_oi, color=C_PE, lw=2.0, label='TOTAL PE')
        d, = r1.plot(x, diff_oi, color=C_DF, lw=1.5, ls='--', label='DIFF PE-CE')
        r1.axhline(0, color='#6b7a8a', lw=1.0)   # PATCH_ABC_V7 (A): DIFF axis ki zero line
        ax1.yaxis.set_major_formatter(fmt_k); r1.yaxis.set_major_formatter(fmt_k)
        ax1.legend(handles=[a, b], loc='upper left', fontsize=10)
        r1.legend(handles=[d], loc='upper right', fontsize=10)
        _endtext(ax1, [(f"CE OI  {ce_oi[-1]:,}", C_CE),
                       (f"PE OI  {pe_oi[-1]:,}", C_PE),
                       (f"PE-CE  {diff_oi[-1]:,}", C_DF)])

        # Panel 2 : COI
        r2 = ax2.twinx(); r2.grid(False)
        a, = ax2.plot(x, coi_ce, color=C_CE, lw=2.0, label='TOTAL COI CE')
        b, = ax2.plot(x, coi_pe, color=C_PE, lw=2.0, label='TOTAL COI PE')
        d, = r2.plot(x, diff_coi, color=C_DF, lw=1.5, ls='--', label='DIFF PE-CE')
        r2.axhline(0, color='#6b7a8a', lw=1.0)   # PATCH_ABC_V7 (A): DIFF axis ki zero line
        ax2.yaxis.set_major_formatter(fmt_k); r2.yaxis.set_major_formatter(fmt_k)
        ax2.legend(handles=[a, b], loc='upper left', fontsize=10)
        r2.legend(handles=[d], loc='upper right', fontsize=10)
        _endtext(ax2, [(f"COI CE  {coi_ce[-1]:,}", C_CE),
                       (f"COI PE  {coi_pe[-1]:,}", C_PE),
                       (f"PE-CE   {diff_coi[-1]:,}", C_DF)])

        # Panel 3 : PCR + MaxPain + Underlying
        r3 = ax3.twinx(); r3.grid(False)
        a, = ax3.plot(x, pcr, color=C_CE, lw=2.0, label='PCR')
        b, = r3.plot(x, mp,  color=C_PE, lw=2.0, label='MaxPain')
        d, = r3.plot(x, und, color=C_DF, lw=1.5, ls='--', label='Underlying')
        # PCR aur Underlying apni apni range mein fit (Sensibull jaisa); PCR ka bad point scale kharab nahi karega
        ax3.yaxis.set_major_locator(mticker.MaxNLocator(nbins=8))
        _p = np.array(pcr, dtype=float)
        _pm  = float(np.median(_p))
        _pk  = _p[np.abs(_p - _pm) <= 0.20 * _pm] if len(_p) >= 8 else _p   # sirf bad outlier point scale se bahar
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
        _pt = np.array(pcr[ntail:], dtype=float)
        if len(_pt) >= 8:
            _pt = _pt[np.abs(_pt - np.median(_pt)) <= 0.20 * np.median(_pt)]
        _plo_f, _phi_f = PCR_FIXED_AXIS
        if len(_pt):
            _plo_f = min(_plo_f, math.floor(_pt.min() / 0.05) * 0.05)
            _phi_f = max(_phi_f, math.ceil(_pt.max() / 0.05) * 0.05)
        ax3.set_ylim(_plo_f, _phi_f)
        ax3.yaxis.set_major_locator(mticker.MultipleLocator(0.05))
        ax3.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.2f'))
        _uc  = round(history[ntail]["underlying"] / 50.0) * 50.0
        _ut  = und[ntail:]
        _ulo_f = min(_uc - UND_FIXED_HALF, math.floor((min(_ut) - 10) / 50.0) * 50.0)
        _uhi_f = max(_uc + UND_FIXED_HALF, math.ceil((max(_ut) + 10) / 50.0) * 50.0)
        r3.set_ylim(_ulo_f, _uhi_f)
        r3.yaxis.set_major_locator(mticker.MultipleLocator(50))
        r3.yaxis.set_major_formatter(fmt_k)
        ax3.legend(handles=[a], loc='upper left', fontsize=10)
        r3.legend(handles=[b, d], loc='upper right', fontsize=10)
        _endtext(ax3, [(f"PCR         {pcr[-1]:.2f}", C_CE),
                       (f"MAXPAIN     {mp_val:,}", C_PE),
                       (f"UNDERLYING  {und[-1]:,.2f}", C_DF)])

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
    close_t = now.replace(hour=15, minute=40, second=0, microsecond=0)   # PATCH_BIAS_V8
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
            log.info(f"[MAIN] Market closed ({now.strftime('%H:%M')}) - sleeping 60s...")
            time.sleep(60)
            continue

        # Fresh graph every new trading session
        today = now.date()
        if current_session_date != today:
            if current_session_date is not None:
                log.info(f"[MAIN]  New session ({today}) - clearing history!")
                send_telegram_text(f" *New trading session*\n`{today}` - Fresh chart started")
            current_session_date = today
            history = load_history(today)
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
            time.sleep(30)
            continue
        if excel_t and excel_t == last_excel_time:
            log.info(f"[MAIN] Excel not yet updated (last: {last_excel_time}) - waiting...")
            try:
                _age = (datetime.datetime.now() - excel_t).total_seconds() / 60
                if _age > 10 and not globals().get('_stale_warned'):
                    globals()['_stale_warned'] = True
                    send_telegram_text(f'WARNING: Excel data {_age:.0f} min se update nahi hua (file save nahi ho rahi?)')
                elif _age <= 10:
                    globals()['_stale_warned'] = False
            except Exception:
                pass
            time.sleep(30)
            continue
        last_excel_time = excel_t

        history.append(snap)
        save_history(today, history)
        expiry_label = snap.get("expiry", "N/A")

        pcr_val, mp_val = render_chart(history, expiry_label, OUTPUT_FILE)
        if pcr_val is not None:
            caption = build_caption(snap, pcr_val, mp_val)
            send_telegram_photo(OUTPUT_FILE, caption)

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