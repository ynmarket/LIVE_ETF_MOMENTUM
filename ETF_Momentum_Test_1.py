"""
ETF_Momentum_Test_1.py — ETF Momentum Rotation Backtest (C54 universe, NSE)

Historical backtest from START_DATE to END_DATE, plus a separate, freshly
initialised LIVE portfolio that starts on LIVE_START_DATE with INITIAL_CAPITAL
(no positions, entry prices or peaks inherited from the historical backtest).

Run:  python ETF_Momentum_Test_1.py
"""
import subprocess, sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')


def _ensure(pkg, import_as=None):
    try:
        __import__(import_as or pkg)
    except ImportError:
        print(f"Installing {pkg}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet"])


for _p, _i in [("yfinance", "yfinance"), ("pandas", "pandas"), ("numpy", "numpy"),
               ("openpyxl", "openpyxl"), ("plotly", "plotly"), ("google-auth", "google.auth"),
               ("requests", "requests")]:
    _ensure(_p, _i)

import base64
import copy
import hashlib
import html as htmlmod
import json
import math
import os
import warnings
import webbrowser
from datetime import datetime, timedelta, timezone
from pathlib import Path

warnings.filterwarnings("ignore")

# ── Runtime environment (deployment only — no effect on the strategy) ─────────
# Indian market: wall-clock "today" and report timestamps always use IST, so the
# result is the same on a Windows laptop and on a UTC GitHub Actions runner.
IST = timezone(timedelta(hours=5, minutes=30))


def now_ist():
    """Current India time as a naive datetime (IST wall clock)."""
    return datetime.now(IST).replace(tzinfo=None)


IN_CI = os.environ.get("CI", "").lower() == "true"   # set automatically by GitHub Actions


def open_report(path):
    """Open the HTML report locally; never try to launch a browser on a CI server."""
    if not IN_CI:
        webbrowser.open(str(path.resolve()))

# ═════════════════════════════════════════════════════════════════════════════
# USER INPUTS
# ═════════════════════════════════════════════════════════════════════════════
# START_DATE is NOT the live start — it is only the historical data start used for
# momentum / 200-DMA history (and the collapsed historical backtest reference).
START_DATE      = "2021-01-01"
END_DATE        = "2026-09-30"    # end of the historical backtest reference; live data always runs to the latest available day

INITIAL_CAPITAL = 300_000         # ₹3 lakh — starting cash of the fresh live portfolio

# DEMO DATE FOR TESTING
# I am intentionally using 2026-09-01 for demo purposes.
# Later I will change this to 2026-10-01 for my actual live start.
LIVE_START_DATE = "2026-10-01"

# ═════════════════════════════════════════════════════════════════════════════
# OUTPUT DIRECTORY + SKIP LOGIC
# ═════════════════════════════════════════════════════════════════════════════
BASE_DIR = Path(__file__).resolve().parent          # works on Windows and on Linux runners
OUTPUT_DIR = BASE_DIR / "reports"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

EXCEL_PATH = OUTPUT_DIR / "ETF_Momentum_Test_1_Report.xlsx"
HTML_PATH  = OUTPUT_DIR / "ETF_Momentum_Test_1_Report.html"

# Daily-run behaviour: skip only if both outputs were already generated TODAY.
# Older outputs are rebuilt with fresh prices. Pass --force to always rebuild.
FORCE_RERUN = "--force" in sys.argv


def _made_today(p):
    return datetime.fromtimestamp(p.stat().st_mtime, IST).date() == now_ist().date()


if (not FORCE_RERUN and EXCEL_PATH.exists() and HTML_PATH.exists()
        and _made_today(EXCEL_PATH) and _made_today(HTML_PATH)):
    print("Outputs already exist (generated today). Skipping backtest. Use --force to rebuild.")
    print(f"  Excel: {EXCEL_PATH.resolve()}")
    print(f"  HTML : {HTML_PATH.resolve()}")
    open_report(HTML_PATH)
    sys.exit(0)

import numpy as np
import pandas as pd
import yfinance as yf
import plotly.graph_objects as go
import plotly.offline as pyo
from plotly.subplots import make_subplots
from openpyxl.styles import PatternFill, Font

# ═════════════════════════════════════════════════════════════════════════════
# STRATEGY PARAMETERS (hardcoded)
# ═════════════════════════════════════════════════════════════════════════════
N_HOLD      = 6
LB          = [21, 63, 126, 252]
WEIGHTS     = [0.15, 0.40, 0.30, 0.15]
DMA_PERIOD  = 200
MIN_HISTORY = 262
SL_PCT      = 0.15
TRAIL_PCT   = 0.20
LIQUID_RATE = 0.05
RF_ANNUAL   = 0.06
BM_SYMBOL   = "^CRSLDX"

# ═════════════════════════════════════════════════════════════════════════════
# ETF UNIVERSE — C54
# ═════════════════════════════════════════════════════════════════════════════
UNIVERSE = [
    ("NIFTYBEES", "Broad_Equity"), ("JUNIORBEES", "Broad_Equity"), ("MID150BEES", "Broad_Equity"),
    ("MIDSELIETF", "Broad_Equity"), ("MONIFTY500", "Broad_Equity"), ("HDFCSML250", "Broad_Equity"),
    ("MOM100", "Broad_Equity"),
    ("BANKBEES", "Banking_Finance"), ("PSUBNKBEES", "Banking_Finance"),
    ("PVTBANIETF", "Banking_Finance"), ("FINIETF", "Banking_Finance"),
    ("ITBEES", "Sector"), ("PHARMABEES", "Sector"), ("AUTOBEES", "Sector"), ("CONSUMBEES", "Sector"),
    ("CONSUMER", "Sector"), ("FMCGIETF", "Sector"), ("CHEMICAL", "Sector"), ("METALIETF", "Sector"),
    ("OILIETF", "Sector"), ("COMMOIETF", "Sector"), ("INFRAIETF", "Sector"), ("CPSEETF", "Sector"),
    ("ICICIB22", "Sector"), ("MOCAPITAL", "Sector"), ("MODEFENCE", "Sector"), ("MOREALTY", "Sector"),
    ("MOTOUR", "Sector"), ("GROWWPOWER", "Sector"), ("GROWWRAIL", "Sector"), ("GROWWHOSPI", "Sector"),
    ("DIVOPPBEES", "Thematic"), ("EVINDIA", "Thematic"), ("INTERNET", "Thematic"), ("MNC", "Thematic"),
    ("SELECTIPO", "Thematic"), ("TOP10ADD", "Thematic"),
    ("LOWVOLIETF", "Factor"), ("NV20IETF", "Factor"), ("QUAL30IETF", "Factor"), ("NIFTYQLITY", "Factor"),
    ("HDFCGROWTH", "Factor"), ("MOM50", "Factor"), ("MOMENTUM50", "Factor"),
    ("MOVALUE", "Factor"),
    ("MON100", "International"), ("MONQ50", "International"),    # Motilal Oswal Nasdaq Q 50 ETF
    ("MAFANG", "International"), ("HNGSNGBEES", "International"),
    ("GOLDBEES", "Gold_Silver"), ("SILVERBEES", "Gold_Silver"),
    ("LTGILTBEES", "Bonds"), ("EBBETF0430", "Bonds"), ("GILT5YBEES", "Bonds"),
]
CATEGORY = dict(UNIVERSE)


# ═════════════════════════════════════════════════════════════════════════════
# DATA DOWNLOAD
# ═════════════════════════════════════════════════════════════════════════════
def fetch_close(yf_symbol, start, end, volume_out=None):
    """Download daily Close for one symbol; returns a clean Series (may be empty).
    volume_out (dict): if given, the symbol's daily Volume series is stored in it as well."""
    try:
        df = yf.download(yf_symbol, start=start, end=end, progress=False, threads=False)
    except Exception:
        return pd.Series(dtype=float)
    if df is None or df.empty:
        return pd.Series(dtype=float)
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    if "Close" not in df.columns:
        return pd.Series(dtype=float)
    s = df["Close"]
    if isinstance(s, pd.DataFrame):
        s = s.iloc[:, 0]
    s = pd.to_numeric(s, errors="coerce").dropna()
    s = s[s > 0]
    s.index = pd.to_datetime(s.index).tz_localize(None).normalize()
    s = s[~s.index.duplicated(keep="last")].sort_index()
    if volume_out is not None and "Volume" in df.columns:
        v = df["Volume"]
        if isinstance(v, pd.DataFrame):
            v = v.iloc[:, 0]
        v = pd.to_numeric(v, errors="coerce").fillna(0)
        v.index = pd.to_datetime(v.index).tz_localize(None).normalize()
        volume_out[yf_symbol] = v[~v.index.duplicated(keep="last")].sort_index()
    return s


start_ts = pd.Timestamp(START_DATE)
end_ts = pd.Timestamp(END_DATE)
live_ts = pd.Timestamp(LIVE_START_DATE)
dl_start = (start_ts - timedelta(days=400)).strftime("%Y-%m-%d")
# Live monitoring needs the latest prices, so download up to max(END_DATE, today)
data_end_ts = max(end_ts, pd.Timestamp(now_ist().date()))
dl_end = (data_end_ts + timedelta(days=1)).strftime("%Y-%m-%d")   # yfinance end is exclusive

# ── Market-hours guard (data handling only) ──────────────────────────────────
# During NSE hours Yahoo returns TODAY's partial bar (the live price, not a close). It must never
# be used as a closing price: it is removed from the price history and kept separately as a
# live quote for the pre-close ORDER PLAN. From 16:00 IST the day's bar is treated as the close.
MARKET_OPEN_IST = (9, 0)
BAR_FINAL_IST = (16, 0)
_now = now_ist()
MARKET_HOURS = _now.weekday() < 5 and MARKET_OPEN_IST <= (_now.hour, _now.minute) < BAR_FINAL_IST
TODAY_TS = pd.Timestamp(_now.date())
live_quotes = {}                        # symbol -> today's live price (market hours only)

print(f"Downloading data for {len(UNIVERSE)} ETFs..."
      + (f" (market open — today's {TODAY_TS.date()} prices are live, not closes)" if MARKET_HOURS else ""))
raw = {}
dropped = []
dropped_info = {}                       # symbol -> rows received (for the data-status report)
_fetched, _volumes = {}, {}
for nse, _cat in UNIVERSE:
    s = fetch_close(f"{nse}.NS", dl_start, dl_end, volume_out=_volumes)
    if MARKET_HOURS and not s.empty and s.index[-1] >= TODAY_TS:
        live_quotes[nse] = float(s.iloc[-1])
        s = s[s.index < TODAY_TS]
    _fetched[nse] = s

# ── Market session calendar (data handling only) ─────────────────────────────
# On exchange holidays Yahoo inserts placeholder bars (volume 0, prices = previous close, e.g.
# 2026-05-01). They are not trading sessions and must not become entry days. A date counts as a
# real session if at least one ETF in the universe traded on it (volume > 0); a few illiquid ETFs
# report volume 0 every day, so volume is judged across the whole universe, never per ETF.
SESSION_DATES = set()
for _v in _volumes.values():
    SESSION_DATES |= set(_v.index[_v > 0])
_all_bar_dates = set().union(*[s.index for s in _fetched.values() if not s.empty])
HOLIDAY_PLACEHOLDERS = sorted(_all_bar_dates - SESSION_DATES) if SESSION_DATES else []
if HOLIDAY_PLACEHOLDERS:
    print(f"Market calendar: removed {len(HOLIDAY_PLACEHOLDERS)} holiday placeholder date(s) with no trading "
          f"in any ETF: " + ", ".join(str(d.date()) for d in HOLIDAY_PLACEHOLDERS))
elif not SESSION_DATES:
    print("  [warn] no volume data from Yahoo — holiday placeholder bars cannot be detected")

for nse, _cat in UNIVERSE:
    s = _fetched[nse]
    if HOLIDAY_PLACEHOLDERS and not s.empty:
        s = s[~s.index.isin(HOLIDAY_PLACEHOLDERS)]
    if s.empty:
        print(f"  [warn] {nse}: no yfinance data — skipped")
        dropped.append(nse)
        dropped_info[nse] = 0
        continue
    if len(s) < MIN_HISTORY:
        print(f"  [warn] {nse}: only {len(s)} rows (< {MIN_HISTORY}) — excluded")
        dropped.append(nse)
        dropped_info[nse] = len(s)
        continue
    raw[nse] = s
print(f"{len(raw)} ETFs loaded, {len(dropped)} dropped (insufficient history)")

bm_raw = fetch_close(BM_SYMBOL, dl_start, dl_end)
bm_live = None                          # Nifty 500 live value (market hours only)
if MARKET_HOURS and not bm_raw.empty:
    if bm_raw.index[-1] >= TODAY_TS:
        bm_live = float(bm_raw.iloc[-1])
    bm_raw = bm_raw[bm_raw.index < TODAY_TS]
if HOLIDAY_PLACEHOLDERS and not bm_raw.empty:
    bm_raw = bm_raw[~bm_raw.index.isin(HOLIDAY_PLACEHOLDERS)]
if bm_raw.empty:
    print(f"  [warn] benchmark {BM_SYMBOL}: no data — benchmark comparison disabled")

SYMBOLS = list(raw.keys())
# Deployment guard: never build (or publish) a report from a mostly failed download
MIN_ETFS_LOADED = 40
if len(SYMBOLS) < MIN_ETFS_LOADED:
    print(f"[DATA ERROR] Only {len(SYMBOLS)} of {len(UNIVERSE)} ETFs downloaded from Yahoo Finance "
          f"(minimum {MIN_ETFS_LOADED}). Yahoo may be down or rate-limiting this machine. Aborting.")
    sys.exit(2)

# Union trading calendar; forward-fill gaps (never look ahead)
all_dates = sorted(set().union(*[s.index for s in raw.values()]))
cal = pd.DatetimeIndex(all_dates)
etf_close = pd.DataFrame({k: v.reindex(cal) for k, v in raw.items()}).ffill()
hist_vals = {k: v.values for k, v in raw.items()}
hist_idx = {k: v.index for k, v in raw.items()}

cal_bt = cal[(cal >= start_ts) & (cal <= end_ts)]
if len(cal_bt) == 0:
    print("No trading dates inside [START_DATE, END_DATE]. Aborting.")
    sys.exit(1)
LAST_DATE = cal_bt[-1]          # last date of the historical backtest reference
DATA_LAST = cal[-1]             # latest available market date (live "as of" date)


def month_first_days(dates):
    """First trading day of each calendar month present in the price data."""
    s = pd.Series(dates, index=dates)
    return list(s.groupby([dates.year, dates.month]).first())


def rebalance_schedule(execution_dates):
    """[(signal_date, execution_date)]: the FIRST TRADING DAY of each month (the entry day) is both
    the signal day and the execution day — signal and fills use that day's close."""
    return [(e, e) for e in execution_dates]


REBAL_DATES = month_first_days(cal_bt)          # entry days (first trading day of each month)
REBAL_SCHEDULE = rebalance_schedule(REBAL_DATES)


def missing_on(d, symbols=None):
    """ETFs (with enough history) that have NO actual price row on date d (Yahoo gaps)."""
    return [s for s in (SYMBOLS if symbols is None else symbols) if d not in hist_idx[s]]


# ── Saved entry-day decisions (live_state/entry_decisions.json) ─────────────
# On the entry day the user runs the workflow during market hours (~14:00 IST). The ETFs chosen in
# that run (from that moment's live prices) are saved here — the last market-hours run of the day
# wins — and the evening run records them at the entry day's close. If no entry-day run happened,
# the first run with complete entry-day closes decides from the close and saves that decision.
# Only the decision (which ETFs to buy / sell / hold / skip) is saved; prices always come from Yahoo.
ENTRY_PATH = BASE_DIR / "live_state" / "entry_decisions.json"


def load_entries():
    try:
        doc = json.loads(ENTRY_PATH.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except ValueError as e:
        print(f"  [warn] {ENTRY_PATH.name} is not valid JSON ({e}) — ignoring saved entry decisions")
        return {}
    if doc.get("live_start") != LIVE_START_DATE or doc.get("initial_capital") != INITIAL_CAPITAL:
        print(f"  [warn] {ENTRY_PATH.name} belongs to live start {doc.get('live_start')} / capital "
              f"{doc.get('initial_capital')} — ignored for LIVE_START_DATE {LIVE_START_DATE}")
        return {}
    return doc.get("entries", {})


def entry_decisions(rec):
    """Saved entry-day decision → the `decisions` argument of rebalance()."""
    if not rec:
        return None
    return {"top6": rec["top6"], "buy": rec["buy"], "skip": rec["skip"],
            "sell": [(x["symbol"], x["reason"]) for x in rec["sell"]], "scores": rec.get("scores", {})}


ENTRIES = load_entries()
if ENTRIES:
    print("Saved entry-day decisions: " + ", ".join(
        f"{m} ({r['decided_from']})" for m, r in sorted(ENTRIES.items())))


# ═════════════════════════════════════════════════════════════════════════════
# SIGNAL HELPERS
# ═════════════════════════════════════════════════════════════════════════════
def compute_scores(d, extra=None):
    """Momentum score + 200-DMA for every ETF with >= MIN_HISTORY rows up to date d.
    extra = {symbol: live price}: appended as the entry day's price (entry-day order plan only)."""
    out = {}
    for sym in SYMBOLS:
        n = hist_idx[sym].searchsorted(d, side="right")
        v = hist_vals[sym][:n]
        if extra and sym in extra:
            v = np.append(v, extra[sym])
        if len(v) < MIN_HISTORY:
            continue
        today = v[-1]
        rets = [(today - v[-1 - lb]) / v[-1 - lb] * 100 for lb in LB]
        score = sum(r * w for r, w in zip(rets, WEIGHTS))
        dma = float(np.mean(v[-DMA_PERIOD:]))
        out[sym] = dict(score=score, r1=rets[0], r3=rets[1], r6=rets[2], r12=rets[3],
                        price=float(today), dma=dma)
    return out


def px(sym, d):
    return float(etf_close.at[d, sym])


# ═════════════════════════════════════════════════════════════════════════════
# STRATEGY ENGINE (shared by historical backtest and live portfolio)
# ═════════════════════════════════════════════════════════════════════════════
def new_state(capital):
    """Empty portfolio: all cash, no positions, no history."""
    return dict(cash=float(capital), positions={}, trades=[], monthly=[], scores=[], log=[],
                prev_d=None, prev_signal=None, prev_val=float(capital), tid=0)


def rebalance(st, signal_date, execution_date, tag="", verbose=True, audit=False, preview=False,
              decisions=None, live=None):
    """One monthly rebalance on the ENTRY DAY (first trading day of the month), in two phases:

    PHASE A — SIGNAL (data up to and including the entry day's price):
        momentum, 200-DMA, ranking, top 6, stop/rotation exits, entry gate → decisions
    PHASE B — EXECUTION (the same entry day's price):
        sells, buys, entry/exit prices, cash, portfolio value

    preview=True is only for dashboard previews; the state passed in must be a copy.
      decisions — a SAVED entry-day decision (the ETFs chosen in the user's entry-day run) replaces
                  the Phase A recomputation, so the record matches what was decided that day.
      live      — {symbol: live price} for the entry-day ORDER PLAN during market hours: used as the
                  entry day's price for the signal and for the quantities (preview only).
    """
    sd, ed = signal_date, execution_date
    if not preview:
        assert sd == ed, "signal and execution happen on the same entry day"
    positions = st["positions"]   # sym -> dict(qty, entry_date, entry_price, peak, cur, score_entry, ...)
    source = "saved-decision" if decisions else "computed"

    def xpx(sym):
        """Entry-day price: live quote (order plan) → entry-day close."""
        if live and sym in live:
            return float(live[sym])
        return px(sym, ed)

    # ═══ PHASE A — SIGNAL: data up to the entry day's price ═════════════════════
    # Mark held positions at the entry-day price and ratchet the peak with it
    for sym, p in positions.items():
        p["sig"] = xpx(sym)
        if p["sig"] > p["peak"]:
            p["peak"] = p["sig"]
    sc = compute_scores(sd, extra=live)         # momentum + 200-DMA through the entry day
    ranked = sorted(sc.keys(), key=lambda k: sc[k]["score"], reverse=True)
    rank_of = {s: i for i, s in enumerate(ranked, 1)}
    top = ranked[:N_HOLD]

    sell_dec, hold_dec = [], []                 # exits in priority order: Hard SL → Trail → Rotation
    for sym, p in positions.items():
        if p["sig"] <= p["entry_price"] * (1 - SL_PCT):
            sell_dec.append((sym, "SL"))
        elif p["sig"] <= p["peak"] * (1 - TRAIL_PCT):
            sell_dec.append((sym, "Trail"))
        elif sym not in top:
            sell_dec.append((sym, "Rotation"))
        else:
            hold_dec.append(sym)
    sold = {s for s, _ in sell_dec}
    buy_dec, skip_dec = [], []                  # entry gate on signal-date close vs signal-date 200-DMA
    for sym in top:
        if sym in positions and sym not in sold:
            continue
        if sc[sym]["price"] > sc[sym]["dma"]:
            buy_dec.append(sym)
        else:
            skip_dec.append(sym)

    if decisions:                               # saved entry-day decision replaces the recomputation
        top = list(decisions["top6"])
        sell_dec = [(s, r) for s, r in decisions["sell"] if s in positions]
        sold = {s for s, _ in sell_dec}
        hold_dec = [s for s in positions if s not in sold]
        buy_dec = list(decisions["buy"])
        skip_dec = list(decisions["skip"])
    saved_scores = (decisions or {}).get("scores", {})
    if saved_scores:                            # show the saved rank/score, not a recomputed one
        saved_rank = {s: i for i, s in enumerate(sorted(saved_scores, key=lambda k: -saved_scores[k]), 1)}
        rank_of = {**rank_of, **saved_rank}

    # ═══ PHASE B — EXECUTION: prices on execution_date only ═══════════════════
    # Step 1: liquid interest on idle cash since the previous execution
    if st["prev_d"] is not None:
        st["cash"] += st["cash"] * LIQUID_RATE / 365 * (ed - st["prev_d"]).days

    def log(action, sym, exec_price, **kw):
        i = sc.get(sym, {})
        score = saved_scores.get(sym, i["score"] if i else None)
        st["log"].append(dict(
            date=ed.date(), signal_date=sd.date(), action=action, symbol=sym, category=CATEGORY[sym],
            rank=rank_of.get(sym), score=round(score, 3) if score is not None else None,
            signal_price=round(xpx(sym), 4), dma200=round(i["dma"], 4) if i else None,
            price=round(exec_price, 4), source=source, **kw))

    for sym, reason in sell_dec:
        p = positions.pop(sym)
        xp = xpx(sym)
        st["tid"] += 1
        cost = p["entry_price"] * p["qty"]
        proceeds = xp * p["qty"]
        st["cash"] += proceeds
        st["trades"].append(dict(
            trade_id=st["tid"], symbol=sym, category=CATEGORY[sym],
            entry_date=p["entry_date"].date(), exit_date=ed.date(),
            entry_price=round(p["entry_price"], 4), exit_price=round(xp, 4),
            qty=p["qty"], cost=round(cost, 2), proceeds=round(proceeds, 2),
            pnl_inr=round(proceeds - cost, 2),
            pnl_pct=round((xp / p["entry_price"] - 1) * 100, 2),
            hold_days=(ed - p["entry_date"]).days,
            peak_price=round(p["peak"], 4), exit_reason=reason,
            momentum_score_at_entry=round(p["score_entry"], 3),
            entry_signal_date=p["entry_signal_date"].date(), exit_signal_date=sd.date(),
            exit_signal_price=round(p["sig"], 4)))
        log("SELL", sym, xp, qty=p["qty"], value=round(proceeds, 2), target=None, reason=reason,
            pnl_pct=round((xp / p["entry_price"] - 1) * 100, 2))
    for sym in hold_dec:
        p = positions[sym]
        p["cur"] = xpx(sym)
        if p["cur"] > p["peak"]:                # bookkeeping only — decision above already made
            p["peak"] = p["cur"]
        log("HOLD", sym, p["cur"], qty=p["qty"], value=round(p["qty"] * p["cur"], 2), target=None,
            reason="In top 6, no stop hit", pnl_pct=round((p["cur"] / p["entry_price"] - 1) * 100, 2))
    total_val = st["cash"] + sum(p["qty"] * p["cur"] for p in positions.values())
    alloc = total_val / N_HOLD
    entries = []
    for sym in top:                             # keep rank order for buys and skips
        if sym in skip_dec:
            log("SKIP", sym, xpx(sym), qty=0, value=0.0, target=round(alloc, 2),
                reason="Below 200 DMA (slot left empty)", pnl_pct=None)
        elif sym in buy_dec:
            xp = xpx(sym)
            qty = math.floor(min(alloc, st["cash"]) / xp)   # never let cash go negative
            if qty <= 0:
                log("SKIP", sym, xp, qty=0, value=0.0, target=round(alloc, 2),
                    reason="Insufficient cash", pnl_pct=None)
                continue
            st["cash"] -= qty * xp
            score_entry = saved_scores.get(sym, sc.get(sym, {}).get("score", float("nan")))
            positions[sym] = dict(qty=qty, entry_date=ed, entry_price=xp, peak=xp, cur=xp,
                                  score_entry=score_entry, entry_signal_date=sd)
            log("BUY", sym, xp, qty=qty, value=round(qty * xp, 2), target=round(alloc, 2),
                reason="Top 6 and above 200 DMA", pnl_pct=None)
            entries.append(sym)

    # Step 8: snapshot at execution-date prices
    held_val = sum(p["qty"] * p["cur"] for p in positions.values())
    pv = st["cash"] + held_val
    st["monthly"].append(dict(
        rebal_date=ed, signal_date=sd, month_label=ed.strftime("%b-%Y"), portfolio_value=round(pv, 2),
        invested=round(held_val, 2), cash=round(st["cash"], 2), n_held=len(positions),
        monthly_return_pct=round((pv / st["prev_val"] - 1) * 100, 3),
        holdings=", ".join(positions.keys()), exits=", ".join(s for s, _ in sell_dec),
        entries=", ".join(entries), skipped=", ".join(skip_dec),
        port_state={s: p["qty"] for s, p in positions.items()}))
    for rk, sym in enumerate(ranked, 1):
        i = sc[sym]
        st["scores"].append(dict(rebal_date=ed.date(), signal_date=sd.date(), symbol=sym,
                                 category=CATEGORY[sym], rank=rk, score=round(i["score"], 3),
                                 ret_1m=round(i["r1"], 2), ret_3m=round(i["r3"], 2),
                                 ret_6m=round(i["r6"], 2), ret_12m=round(i["r12"], 2),
                                 price=round(i["price"], 4), dma200=round(i["dma"], 4),
                                 above_dma=int(i["price"] > i["dma"]),
                                 in_top6=int(sym in top), held=int(sym in positions)))
    if verbose and audit:
        mine = [a for a in st["log"] if a["date"] == ed.date() and a["signal_date"] == sd.date()]
        print("=" * 60)
        print(f"{tag}MONTHLY REBALANCE")
        print(f"Entry Day      : {ed.date()}  (signal + execution)")
        print("=" * 60)
        print(f"Signal from data up to the {ed.date()} close"
              + ("  [saved entry-day decision replayed]" if source == "saved-decision" else ""))
        for act in ("BUY", "SELL", "HOLD", "SKIP"):
            items = [a for a in mine if a["action"] == act]
            print(f"{act}:" + ("" if items else " none"))
            for a in items:
                extra = f" ({a['reason']})" if act in ("SELL", "SKIP") else ""
                print(f"  {a['symbol']}{extra}")
        print("Execution prices:")
        for a in mine:
            if a["action"] in ("BUY", "SELL"):
                print(f"  {a['symbol']} {a['action']} → ₹{a['price']:,.2f} on {ed.date()}")
        print(f"After execution: Held {len(positions)} | Cash ₹{st['cash']:,.0f} | Portfolio ₹{pv:,.0f}\n")
    elif verbose:
        print(f"{tag}entry {ed.date()} | Held: {len(positions)} | "
              f"Cash: ₹{st['cash']:,.0f} | Portfolio: ₹{pv:,.0f}")
    st["prev_d"], st["prev_signal"], st["prev_val"] = ed, sd, pv


def run_strategy(schedule, capital, tag, verbose=True, audit=False, entries=None):
    """Deterministic simulation: fresh capital, then every entry day in order.
    `entries` (live portfolio only): {"YYYY-MM": saved entry-day decision} — replayed when present."""
    st = new_state(capital)
    for sd, ed in schedule:
        rec = (entries or {}).get(ed.strftime("%Y-%m"))
        rebalance(st, sd, ed, tag, verbose, audit, decisions=entry_decisions(rec))
    return st


def daily_equity(res, end_date):
    """Daily portfolio value: holdings fixed between rebalances, cash accrues liquid interest."""
    mon = res["monthly"]
    rows = []
    for i, m in enumerate(mon):
        seg_start = m["rebal_date"]
        seg_end = mon[i + 1]["rebal_date"] if i + 1 < len(mon) else end_date + timedelta(days=1)
        seg_dates = cal[(cal >= seg_start) & (cal < seg_end) & (cal <= end_date)]
        if len(seg_dates) == 0:
            continue
        held_val = pd.Series(0.0, index=seg_dates)
        for nse, qty in m["port_state"].items():
            prices = etf_close[nse].reindex(seg_dates, method="ffill").fillna(0)
            held_val += qty * prices
        days = np.array([(x - seg_start).days for x in seg_dates])
        cash_s = pd.Series(m["cash"] * (1 + LIQUID_RATE / 365 * days), index=seg_dates)
        rows.append(pd.DataFrame({"held": held_val, "cash": cash_s}))
    df = pd.concat(rows)
    df["value"] = df["held"] + df["cash"]
    df["peak"] = df["value"].cummax()
    df["dd_pct"] = (df["value"] / df["peak"] - 1) * 100
    return df


def open_positions_df(res, as_of):
    rows = []
    for sym, p in res["positions"].items():
        cur = px(sym, as_of)
        cost = p["qty"] * p["entry_price"]
        val = p["qty"] * cur
        rows.append(dict(symbol=sym, category=CATEGORY[sym], entry_date=p["entry_date"].date(),
                         entry_price=round(p["entry_price"], 4), current_price=round(cur, 4),
                         qty=p["qty"], cost=round(cost, 2), current_value=round(val, 2),
                         unrealised_pnl_inr=round(val - cost, 2),
                         unrealised_pnl_pct=round((cur / p["entry_price"] - 1) * 100, 2),
                         hold_days=(as_of - p["entry_date"]).days,
                         peak_price=round(p["peak"], 4),       # stored peak = the level the engine checks
                         sl_level=round(p["entry_price"] * (1 - SL_PCT), 4),
                         trail_level=round(p["peak"] * (1 - TRAIL_PCT), 4)))
    return pd.DataFrame(rows)


# ═════════════════════════════════════════════════════════════════════════════
# RUN HISTORICAL BACKTEST (internal reference only — never feeds the live portfolio)
# ═════════════════════════════════════════════════════════════════════════════
print(f"\nHistorical backtest: {REBAL_DATES[0].date()} → {LAST_DATE.date()} ({len(REBAL_DATES)} rebalances, "
      f"signal and execution on the first trading day's close)")
bt = run_strategy(REBAL_SCHEDULE, INITIAL_CAPITAL, tag="")
daily_df = daily_equity(bt, LAST_DATE)

# ═════════════════════════════════════════════════════════════════════════════
# LIVE PORTFOLIO — fresh start on LIVE_START_DATE (nothing inherited)
# ═════════════════════════════════════════════════════════════════════════════
# The live portfolio is rebuilt deterministically on every run:
#   new_state(INITIAL_CAPITAL)
#   → initialisation on the first trading day on/after LIVE_START_DATE (signal + execution)
#   → every later month: first trading day = entry day (signal + execution)
#   → today's state. Saved entry-day decisions are replayed (live_state/entry_decisions.json).
# Nothing is taken from the historical backtest `bt`.
live = None
live_daily = None
live_init_date = None
live_schedule = []
_cal_live = cal[cal >= live_ts]
if len(_cal_live) == 0:
    print(f"\n[info] No market data yet on/after LIVE_START_DATE {LIVE_START_DATE} — "
          f"live portfolio not started; dashboard shows the initialisation preview")
else:
    live_schedule = rebalance_schedule(month_first_days(_cal_live))   # first = initialisation
    live_init_date = live_schedule[0][1]
    print(f"\nLive portfolio: initialised {live_init_date.date()} (entry day: signal + execution) "
          f"with ₹{INITIAL_CAPITAL:,.0f} ({len(live_schedule)} live rebalance(s))\n")
    live = run_strategy(live_schedule, INITIAL_CAPITAL, tag="[LIVE] ", audit=True, entries=ENTRIES)
    live_daily = daily_equity(live, DATA_LAST)


# ═════════════════════════════════════════════════════════════════════════════
# LOOK-AHEAD VALIDATION (runs every time; raises if any check fails)
# ═════════════════════════════════════════════════════════════════════════════
def _independent_score(sym, upto):
    """Recompute score/DMA from the raw series TRUNCATED at `upto` (independent of compute_scores)."""
    v = raw[sym][raw[sym].index <= upto].values
    if len(v) < MIN_HISTORY:
        return None
    r = [(v[-1] / v[-1 - lb] - 1) * 100 for lb in LB]
    return sum(a * w for a, w in zip(r, WEIGHTS)), float(np.mean(v[-DMA_PERIOD:])), float(v[-1])


def validate_no_lookahead(st, label, show):
    """Check every rebalance of `st`: signal and execution on the same entry day, the signal uses no
    data after that day's close, and every fill is at that day's close. Prints one example."""
    n_checked = 0
    for m in st["monthly"]:
        sd, ed = m["signal_date"], m["rebal_date"]
        assert sd == ed and ed in cal, f"{label}: signal {sd} and execution {ed} must be the same entry day"
        rows = [r for r in st["scores"] if r["rebal_date"] == ed.date()]
        for r in rows:                                       # every score/DMA/rank input <= entry day
            s, dma, sp = _independent_score(r["symbol"], ed)
            assert abs(s - r["score"]) < 1e-3 and abs(dma - r["dma200"]) < 1e-3 and abs(sp - r["price"]) < 1e-3, \
                f"{label}: {r['symbol']} signal on {ed.date()} does not match data <= {ed.date()}"
        for a in st["log"]:                                   # every fill at the entry-day close
            if a["date"] == ed.date() and a["action"] in ("BUY", "SELL"):
                assert abs(a["price"] - px(a["symbol"], ed)) < 1e-3, f"{label}: {a['symbol']} fill not at {ed.date()} close"
        n_checked += 1
    for t in st["trades"]:
        assert t["entry_signal_date"] == t["entry_date"] and t["exit_signal_date"] == t["exit_date"]

    m = st["monthly"][show]
    ed = m["rebal_date"]
    rows = sorted((r for r in st["scores"] if r["rebal_date"] == ed.date()), key=lambda r: r["rank"])
    print("\n" + "=" * 72)
    print(f"LOOK-AHEAD VALIDATION — {label}: {n_checked} entry days checked, all passed")
    print("=" * 72)
    print(f"Example entry day : {ed.date()} ({ed:%a}) — signal and execution on this day's close")
    print(f"Top 6 from data <= {ed.date()} : {', '.join(r['symbol'] for r in rows[:N_HOLD])}")
    print(f"{'Symbol':<11}{'Action':<7}{'Score':>10}{'Close':>12}{'Fill':>12}")
    for a in (x for x in st["log"] if x["date"] == ed.date() and x["action"] in ("BUY", "SELL", "HOLD")):
        print(f"{a['symbol']:<11}{a['action']:<7}{(a['score'] or float('nan')):>10.3f}"
              f"{px(a['symbol'], ed):>12.2f}{a['price']:>12.2f}")
    print("✓ momentum, 200-DMA, ranking and stop/entry decisions use data up to the entry-day close only")
    print("✓ every BUY/SELL price equals the entry-day close\n")


validate_no_lookahead(bt, "HISTORICAL BACKTEST", show=-1)
if live is not None:
    validate_no_lookahead(live, "LIVE PORTFOLIO", show=0)

# ═════════════════════════════════════════════════════════════════════════════
# HISTORICAL REFERENCE METRICS (shown only in the collapsed section of the report)
# ═════════════════════════════════════════════════════════════════════════════
mdf = pd.DataFrame(bt["monthly"])
mdf["cumulative_return_pct"] = (mdf["portfolio_value"] / INITIAL_CAPITAL - 1) * 100
mdf["drawdown_pct"] = (mdf["portfolio_value"] / mdf["portfolio_value"].cummax() - 1) * 100
trades_df = pd.DataFrame(bt["trades"])
if not trades_df.empty:
    trades_df = trades_df.sort_values(["exit_date", "trade_id"]).reset_index(drop=True)
open_df = open_positions_df(bt, LAST_DATE)

final_value = float(daily_df["value"].iloc[-1])
first_d = REBAL_DATES[0]
years = (LAST_DATE - first_d).days / 365.25
cagr = ((final_value / INITIAL_CAPITAL) ** (1 / years) - 1) * 100 if years > 0 else 0.0

# Monthly return series: rebalance-to-rebalance, plus final partial period to LAST_DATE
mvals = list(mdf["portfolio_value"])
if LAST_DATE > REBAL_DATES[-1]:
    mvals.append(final_value)
mret = pd.Series(mvals).pct_change().dropna()
mret = pd.concat([pd.Series([mvals[0] / INITIAL_CAPITAL - 1]), mret], ignore_index=True)
rf_m = RF_ANNUAL / 12
std_m = mret.std()
sharpe = (mret.mean() - rf_m) / std_m * math.sqrt(12) if std_m > 0 else 0.0
neg = mret[mret < 0]
dstd = neg.std() if len(neg) > 1 else 0.0
sortino = (mret.mean() - rf_m) / dstd * math.sqrt(12) if dstd > 0 else 0.0
max_dd_pct = float(mdf["drawdown_pct"].min())
max_dd_daily = float(daily_df["dd_pct"].min())
calmar = cagr / abs(max_dd_daily) if max_dd_daily != 0 else 0.0
vol = std_m * math.sqrt(12) * 100
winrate_m = (mret > 0).mean() * 100

bm_daily = None
bm_cagr = float("nan")
if not bm_raw.empty:
    bm_daily = bm_raw.reindex(daily_df.index, method="ffill").ffill().bfill()
    if bm_daily.notna().all():
        bm_cagr = ((bm_daily.iloc[-1] / bm_daily.iloc[0]) ** (1 / years) - 1) * 100
    else:
        bm_daily = None
alpha = cagr - bm_cagr if not math.isnan(bm_cagr) else float("nan")

n_tr = len(trades_df)
if n_tr:
    trade_wr = (trades_df["pnl_inr"] > 0).mean() * 100
    avg_hold = trades_df["hold_days"].mean()
    avg_pnl = trades_df["pnl_pct"].mean()
    best_tr = trades_df["pnl_pct"].max()
    worst_tr = trades_df["pnl_pct"].min()
    ex_counts = trades_df["exit_reason"].value_counts()
else:
    trade_wr = avg_hold = avg_pnl = best_tr = worst_tr = 0.0
    ex_counts = pd.Series(dtype=int)

# Calendar-month and calendar-year returns (from daily equity)
_me = daily_df["value"].resample("ME").last()
cal_mret = _me.pct_change() * 100
cal_mret.iloc[0] = (_me.iloc[0] / INITIAL_CAPITAL - 1) * 100
_ye = daily_df["value"].resample("YE").last()
annual = _ye.pct_change() * 100
annual.iloc[0] = (_ye.iloc[0] / INITIAL_CAPITAL - 1) * 100
annual.index = annual.index.year
bm_annual = None
if bm_daily is not None:
    _bye = bm_daily.resample("YE").last()
    bm_annual = _bye.pct_change() * 100
    bm_annual.iloc[0] = (_bye.iloc[0] / bm_daily.iloc[0] - 1) * 100
    bm_annual.index = bm_annual.index.year

metrics = [
    ("CAGR%", round(cagr, 2)),
    ("MaxDD% (monthly)", round(max_dd_pct, 2)),
    ("MaxDD% (daily)", round(max_dd_daily, 2)),
    ("Sharpe", round(sharpe, 2)),
    ("Sortino", round(sortino, 2)),
    ("Calmar", round(calmar, 2)),
    ("Volatility%", round(vol, 2)),
    ("WinRate_M%", round(winrate_m, 2)),
    ("BM_CAGR% (Nifty 500)", round(bm_cagr, 2) if not math.isnan(bm_cagr) else "n/a"),
    ("Alpha%", round(alpha, 2) if not math.isnan(alpha) else "n/a"),
    ("Final_Value ₹", round(final_value, 2)),
    ("N_Trades", n_tr),
    ("Trade_WR%", round(trade_wr, 2)),
    ("Avg_Hold_Days", round(avg_hold, 1)),
    ("Avg_PnL%", round(avg_pnl, 2)),
    ("Best_Trade%", round(best_tr, 2)),
    ("Worst_Trade%", round(worst_tr, 2)),
    ("Exit_SL", int(ex_counts.get("SL", 0))),
    ("Exit_Trail", int(ex_counts.get("Trail", 0))),
    ("Exit_Rotation", int(ex_counts.get("Rotation", 0))),
    ("Period", f"{first_d.date()} → {LAST_DATE.date()} ({years:.2f} yrs)"),
    ("Rebalances", len(REBAL_DATES)),
    ("ETFs loaded / dropped", f"{len(SYMBOLS)} / {len(dropped)}"),
]

# ═════════════════════════════════════════════════════════════════════════════
# LIVE DASHBOARD STATE (everything below the historical block is live-only)
# ═════════════════════════════════════════════════════════════════════════════
AS_OF = DATA_LAST                      # latest close in the data (today's live bar is never a close)
TODAY = pd.Timestamp(now_ist().date())
NEAR_STOP_PCT = 5.0                    # reporting only: flag positions within 5% of a stop

live_started = live is not None
live_state = live if live_started else new_state(INITIAL_CAPITAL)


# ── Entry-day timing: the 1st trading day of each month is signal AND execution day ──
# Future sessions are not in the price data, so a future entry day is an estimate that skips
# weekends only (exchange holidays are unknown in advance). Past/present dates come from the data.
def _first_weekday_on_or_after(d):
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d


def _last_weekday_before(d):
    d -= timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


cur_rebal = live["prev_d"] if live_started else None            # last recorded entry day
cur_signal = cur_rebal                                           # signal day = entry day
executed_today = live_started and cur_rebal == AS_OF

next_anchor = live_ts if not live_started else (AS_OF + pd.offsets.MonthBegin(1)).normalize()
_later = cal[cal >= next_anchor]
if len(_later):                                   # (only possible if data runs past the anchor)
    next_rebal, next_confirmed = _later[0], True
else:
    next_rebal, next_confirmed = _first_weekday_on_or_after(next_anchor), False
days_to_next = (next_rebal - TODAY).days
next_month = next_rebal.strftime("%Y-%m")
entry_today = bool(next_rebal == TODAY and not executed_today)
if entry_today and MARKET_HOURS:
    next_confirmed = True                         # the market is open today, so today IS a trading day
next_signal, next_signal_known = next_rebal, next_confirmed     # signal and execution: same entry day

# Entry-day ORDER PLAN: during market hours on the entry day, today's live Yahoo price is used as the
# entry day's price for the signal (momentum, 200-DMA, ranking, stops) and for the quantities.
order_plan = bool(MARKET_HOURS and entry_today)
live_quote_time = _now.strftime("%H:%M") if MARKET_HOURS else None
# LIVE VIEW: during market hours every CURRENT value (positions, P&L, stops, signals/ranking "now",
# My Account, Telegram) uses today's live price. It is never recorded as a close: the live portfolio
# record, saved entry decisions, backtest and history charts stay on closing prices.
LIVE_VIEW = bool(MARKET_HOURS and live_quotes)
live_sc = compute_scores(AS_OF, extra=live_quotes) if LIVE_VIEW else None


def cur_px(sym):
    """Current price: live quote during market hours (if available), else the latest close."""
    return float(live_quotes[sym]) if LIVE_VIEW and sym in live_quotes else px(sym, AS_OF)


def cur_px_label(sym):
    """'live HH:MM' for a live quote, else the date of the latest actual close."""
    if LIVE_VIEW and sym in live_quotes:
        return f"live {live_quote_time}"
    return last_px_date[sym].date() if sym in last_px_date else AS_OF.date()

# Live-price completeness: ETFs without a live quote would be ranked on yesterday's close. Up to 2
# gaps are tolerated only for ETFs that are neither held nor in the live top 10.
QUOTE_GAP_TOLERANCE = 2
signal_missing = [s for s in SYMBOLS if s not in live_quotes] if order_plan else []
_live_top10 = sorted(live_sc, key=lambda k: live_sc[k]["score"], reverse=True)[:10] if live_sc else []
_key_syms = set(live_state["positions"]) | set(_live_top10)
signal_incomplete = bool(order_plan and signal_missing
                         and (len(signal_missing) > QUOTE_GAP_TOLERANCE or _key_syms & set(signal_missing)))

# Ranking shown on the dashboard: the live ranking during market hours, else the latest close
cur_sc = live_sc if LIVE_VIEW else compute_scores(AS_OF)
cur_ranked = sorted(cur_sc, key=lambda k: cur_sc[k]["score"], reverse=True)
cur_rank = {s: i for i, s in enumerate(cur_ranked, 1)}
cur_top = set(cur_ranked[:N_HOLD])

live_cash = float(live_daily["cash"].iloc[-1]) if live_started else float(INITIAL_CAPITAL)

# Current live positions
pos_rows = []
for sym, p in live_state["positions"].items():
    price = cur_px(sym)
    peak = p["peak"]          # stored peak = highest entry-day close since entry: the level the engine checks
    hard = p["entry_price"] * (1 - SL_PCT)
    trail = peak * (1 - TRAIL_PCT)
    d_sl = (price / hard - 1) * 100
    d_tr = (price / trail - 1) * 100
    i = cur_sc.get(sym, {})
    if price <= hard:
        status = "EXIT RISK — Hard SL hit"
    elif price <= trail:
        status = "EXIT RISK — Trail hit"
    elif sym not in cur_top:
        status = "EXIT RISK — outside top 6"
    elif min(d_sl, d_tr) <= NEAR_STOP_PCT:
        status = "EXIT RISK — near stop"
    else:
        status = "HOLD"
    cost = p["qty"] * p["entry_price"]
    val = p["qty"] * price
    pos_rows.append(dict(
        rank=cur_rank.get(sym), symbol=sym, category=CATEGORY[sym],
        score=i.get("score"), r1=i.get("r1"), r3=i.get("r3"), r6=i.get("r6"), r12=i.get("r12"),
        dma=i.get("dma"), price=price, above_dma=bool(i) and price > i["dma"],
        entry_date=p["entry_date"].date(), entry_price=p["entry_price"], qty=p["qty"],
        cost=cost, value=val, pnl=val - cost, pnl_pct=(price / p["entry_price"] - 1) * 100,
        hold_days=(AS_OF - p["entry_date"]).days, peak=peak, hard_sl=hard, trail_sl=trail,
        dist_sl=d_sl, dist_trail=d_tr, weight=0.0, status=status))
invested = sum(r["value"] for r in pos_rows)
live_value = invested + live_cash
for r in pos_rows:
    r["weight"] = r["value"] / live_value * 100
pos_rows.sort(key=lambda r: (r["rank"] is None, r["rank"] or 0))
live_pnl = live_value - INITIAL_CAPITAL
live_ret = live_pnl / INITIAL_CAPITAL * 100

if executed_today:
    signal_status = "EXECUTED"
elif signal_incomplete:
    signal_status = "INCOMPLETE"
elif order_plan:
    signal_status = "ORDER PLAN"
elif entry_today:
    signal_status = "ENTRY DAY"
else:
    signal_status = "PENDING"

# Scheduled midday run (GitHub, ~14:30 IST): an entry-day choice already saved TODAY (e.g. by the
# user's own order-plan run) is kept, never replaced — the order plan then shows that saved choice
# with quantities from live prices. Manual runs behave as before (the last one of the day wins).
MIDDAY_RUN = os.environ.get("RUN_KIND") == "midday"
_saved_today = ENTRIES.get(next_month) if order_plan else None
kept_decision = bool(MIDDAY_RUN and _saved_today and _saved_today.get("entry_date") == str(TODAY.date()))
kept_decision_info = (f"{_saved_today['decided_from']} (saved {_saved_today.get('saved', {}).get('at_ist', '?')} IST)"
                      if kept_decision else None)

# Actions: the recorded entry-day rebalance (today's close), or the SAME engine run on a copy —
# with today's live prices during the entry-day order plan, otherwise a preview from the latest close.
if executed_today:
    act_mode = "executed"
    act_log = [a for a in live["log"] if a["date"] == AS_OF.date()]
else:
    act_mode = "incomplete" if signal_incomplete else "orderplan" if order_plan else "preview"
    _pv = copy.deepcopy(live_state)
    rebalance(_pv, AS_OF, AS_OF, verbose=False, preview=True, live=live_quotes if LIVE_VIEW else None,
              decisions=entry_decisions(_saved_today) if kept_decision else None)
    act_log = [a for a in _pv["log"] if a["date"] == AS_OF.date() and a["signal_date"] == AS_OF.date()]
acts = {k: [a for a in act_log if a["action"] == k] for k in ("BUY", "HOLD", "SELL", "SKIP")}
_rec_now = ENTRIES.get(AS_OF.strftime("%Y-%m")) if executed_today else None
signal_source = ("live" if act_mode in ("orderplan", "incomplete")
                 else (_rec_now or {}).get("decided_from", "entry-day close") if executed_today else "latest close")
order_plan_missing_quotes = [a["symbol"] for a in act_log
                             if order_plan and a["action"] in ("BUY", "SELL", "HOLD") and a["symbol"] not in live_quotes]
# Market hours by the clock, but Yahoo has no bar for today at all (NSE holiday, or not updated yet)
NO_LIVE_TODAY = bool(MARKET_HOURS and not live_quotes)
if LIVE_VIEW:
    print(f"\n  [MARKET OPEN] Today's ({TODAY.date()}) prices are live quotes, not closes — excluded from "
          f"the price history; data shown up to {AS_OF.date()}.")
elif NO_LIVE_TODAY:
    print(f"\n  [NO LIVE PRICES] Market hours, but Yahoo has no prices for today ({TODAY.date()}) — NSE holiday, "
          f"or Yahoo not updated yet; values use the {AS_OF.date()} close.")
if order_plan and not signal_incomplete:
    print(f"  [ORDER PLAN] Entry day {TODAY.date()}: signal and quantities from live prices at "
          f"{live_quote_time} IST — NOT recorded until today's close")
if signal_incomplete:
    print(f"  [DATA INCOMPLETE] {len(signal_missing)} ETFs have no live price: "
          f"{', '.join(signal_missing)} — do not trade on this order plan")

# ═════════════════════════════════════════════════════════════════════════════
# SAVE ENTRY-DAY DECISIONS (GitHub Actions only; the workflow commits the file)
# ═════════════════════════════════════════════════════════════════════════════
def _decisions_from_log(entries, top6, score_map):
    return {
        "top6": list(top6),
        "buy": [a["symbol"] for a in entries
                if a["action"] == "BUY" or (a["action"] == "SKIP" and a["reason"] == "Insufficient cash")],
        "skip": [a["symbol"] for a in entries if a["action"] == "SKIP" and a["reason"] != "Insufficient cash"],
        "sell": [{"symbol": a["symbol"], "reason": a["reason"]} for a in entries if a["action"] == "SELL"],
        "hold": [a["symbol"] for a in entries if a["action"] == "HOLD"],
        "scores": {s: round(float(v), 4) for s, v in score_map.items()},
    }


entry_notes = []
_entries_new = copy.deepcopy(ENTRIES)
_stamp = {"at_ist": now_ist().strftime("%Y-%m-%d %H:%M"), "run_id": os.environ.get("GITHUB_RUN_ID")}
# 1) Entry-day run during market hours: save the ETFs chosen now (the last such run of the day wins;
#    the scheduled midday run never replaces a choice already saved today)
if kept_decision:
    print(f"  [KEPT] {next_month} entry-day choice already saved today — {kept_decision_info}; "
          f"scheduled midday run does not replace it")
elif order_plan and not signal_incomplete:
    _entries_new[next_month] = {
        "entry_date": str(TODAY.date()),
        "decided_from": f"entry-day run, live prices {live_quote_time} IST",
        **_decisions_from_log(act_log, cur_ranked[:N_HOLD], {s: cur_sc[s]["score"] for s in cur_ranked[:10]}),
        "saved": _stamp,
    }
    entry_notes.append(f"{next_month} decision from live prices at {live_quote_time} IST")
# 2) Entry day recorded at the close without an entry-day run: save the close-based decision
#    (only when every ETF has that day's close, so a Yahoo gap is never saved)
if live_started:
    for m in live["monthly"]:
        ed = m["rebal_date"]
        month = ed.strftime("%Y-%m")
        if month in _entries_new or missing_on(ed):
            continue
        entries = [a for a in live["log"] if a["date"] == ed.date()]
        rows = sorted((r for r in live["scores"] if r["rebal_date"] == ed.date()), key=lambda r: r["rank"])
        _entries_new[month] = {
            "entry_date": str(ed.date()),
            "decided_from": "entry-day close (no entry-day run)",
            **_decisions_from_log(entries, [r["symbol"] for r in rows[:N_HOLD]],
                                  {r["symbol"]: r["score"] for r in rows[:10]}),
            "saved": _stamp,
        }
        entry_notes.append(f"{month} decision from the {ed.date()} close")
if entry_notes:
    if IN_CI:
        ENTRY_PATH.parent.mkdir(parents=True, exist_ok=True)
        ENTRY_PATH.write_text(json.dumps({"live_start": LIVE_START_DATE, "initial_capital": INITIAL_CAPITAL,
                                          "entries": dict(sorted(_entries_new.items()))},
                                         indent=2, ensure_ascii=False), encoding="utf-8")
        print("  [SAVED] " + "; ".join(entry_notes) + f" → {ENTRY_PATH.relative_to(BASE_DIR)}")
    else:
        print("  [SAVE] Would save " + "; ".join(entry_notes) + " (only GitHub Actions writes the file)")

# ── Data freshness (reporting only — no calculation is changed) ─────────────
# An ETF is "delayed" when Yahoo has no row for it on the latest market date (AS_OF);
# the strategy then uses its last available close (forward-fill), which the dashboard flags.
last_px_date = {s: raw[s].index[-1] for s in SYMBOLS}
sessions_behind = {s: int(((cal > d) & (cal <= AS_OF)).sum()) for s, d in last_px_date.items()}
stale = {s: last_px_date[s] for s in SYMBOLS if last_px_date[s] < AS_OF}
stale_groups = {}
for s, d in sorted(stale.items(), key=lambda x: (x[1], x[0])):
    stale_groups.setdefault(d, []).append(s)
expected_latest = _last_weekday_before(TODAY)      # most recent completed weekday session
market_delayed = AS_OF < expected_latest
stale_held = [s for s in live_state["positions"] if s in stale]
stale_signal = [s for s in cur_ranked[:N_HOLD] if s in stale]
stale_actions = sorted({a["symbol"] for a in act_log if a["symbol"] in stale})
near_miss = {s: n for s, n in dropped_info.items() if MIN_HISTORY - 10 <= n < MIN_HISTORY}
data_ok = not stale and not market_delayed
data_status = "CURRENT" if data_ok else "DELAYED"
if not data_ok:
    print(f"\n  [DATA WARNING] {len(stale)} of {len(SYMBOLS)} ETFs have no price for {AS_OF.date()}"
          + (f"; latest market data {AS_OF.date()} is older than expected {expected_latest.date()}"
             if market_delayed else ""))
    for d, syms in stale_groups.items():
        print(f"    last price {d.date()}: {', '.join(syms)}")

# Full ranked universe with signal
target_now = live_value / N_HOLD
sig_rows = []
for sym in cur_ranked:
    i = cur_sc[sym]
    held = sym in live_state["positions"]
    above = i["price"] > i["dma"]
    if sym in cur_top:
        sig = "HOLD" if held else ("NEW ENTRY" if above else "SKIP — BELOW 200 DMA")
    else:
        sig = "ROTATION RISK" if held else "OUTSIDE TOP 6"
    sig_rows.append(dict(
        rank=cur_rank[sym], symbol=sym, category=CATEGORY[sym], score=i["score"],
        r1=i["r1"], r3=i["r3"], r6=i["r6"], r12=i["r12"], dma=i["dma"], price=i["price"],
        above_dma=above, held=held, signal=sig,
        target=target_now if sig in ("HOLD", "NEW ENTRY") else None,
        est_qty=math.floor(target_now / i["price"]) if sig == "NEW ENTRY" else None,
        price_date=cur_px_label(sym),
        behind=0 if (LIVE_VIEW and sym in live_quotes) else sessions_behind[sym]))
for r in pos_rows:
    r["price_date"] = cur_px_label(r["symbol"])
    r["behind"] = 0 if (LIVE_VIEW and r["symbol"] in live_quotes) else sessions_behind[r["symbol"]]

# Live performance (only from LIVE_START_DATE)
live_trades_df = pd.DataFrame(live_state["trades"])
if live_started and not live_trades_df.empty:
    live_trades_df = live_trades_df[live_trades_df["entry_date"] >= live_init_date.date()]
    live_trades_df = live_trades_df.sort_values(["exit_date", "trade_id"]).reset_index(drop=True)
lt = live_trades_df
n_lt = len(lt)
lt_stats = dict(
    total=n_lt,
    wins=int((lt["pnl_inr"] > 0).sum()) if n_lt else 0,
    losses=int((lt["pnl_inr"] <= 0).sum()) if n_lt else 0,
    wr=(lt["pnl_inr"] > 0).mean() * 100 if n_lt else None,
    realised=float(lt["pnl_inr"].sum()) if n_lt else 0.0,
    avg_pnl=float(lt["pnl_pct"].mean()) if n_lt else None,
    avg_hold=float(lt["hold_days"].mean()) if n_lt else None,
    best=float(lt["pnl_pct"].max()) if n_lt else None,
    worst=float(lt["pnl_pct"].min()) if n_lt else None,
    sl=int((lt["exit_reason"] == "SL").sum()) if n_lt else 0,
    trail=int((lt["exit_reason"] == "Trail").sum()) if n_lt else 0,
    rot=int((lt["exit_reason"] == "Rotation").sum()) if n_lt else 0,
)

live_max_dd = None
live_mret = pd.Series(dtype=float)
if live_started:
    live_daily["daily_ret_pct"] = live_daily["value"].pct_change() * 100
    live_daily.iloc[0, live_daily.columns.get_loc("daily_ret_pct")] = (
        live_daily["value"].iloc[0] / INITIAL_CAPITAL - 1) * 100
    live_daily["cum_ret_pct"] = (live_daily["value"] / INITIAL_CAPITAL - 1) * 100
    live_max_dd = float(live_daily["dd_pct"].min())
    _lme = live_daily["value"].resample("ME").last()
    live_mret = _lme.pct_change() * 100
    live_mret.iloc[0] = (_lme.iloc[0] / INITIAL_CAPITAL - 1) * 100
win_months = int((live_mret > 0).sum())
lose_months = int((live_mret <= 0).sum())

print("\n" + "═" * 60)
print("  LIVE PORTFOLIO — as of", AS_OF.date())
print("═" * 60)
if live_started:
    print(f"  Live start      {live_init_date.date()}   Initial ₹{INITIAL_CAPITAL:,.0f}")
    print(f"  Value           ₹{live_value:,.0f}  ({live_ret:+.2f}%, P&L ₹{live_pnl:,.0f})")
    print(f"  Invested / Cash ₹{invested:,.0f} / ₹{live_cash:,.0f}")
    print(f"  Positions       {len(pos_rows)}   Max DD {live_max_dd:.2f}%   Closed trades {n_lt}")
    for r in pos_rows:
        print(f"    #{str(r['rank']):<3} {r['symbol']:<11} {r['pnl_pct']:+6.2f}%  {r['status']}")
else:
    print(f"  Not started — LIVE_START_DATE {LIVE_START_DATE} has no market data yet.")
    print(f"  Initialisation preview (as of {AS_OF.date()}): BUY {', '.join(a['symbol'] for a in acts['BUY']) or 'none'}")
print(f"  Actions ({act_mode}): BUY {len(acts['BUY'])} | HOLD {len(acts['HOLD'])} | "
      f"SELL {len(acts['SELL'])} | SKIP {len(acts['SKIP'])}")
if live_started:
    print(f"  Last entry day  {cur_rebal.date()} (signal + execution)")
print(f"  Latest data     {AS_OF.date()}")
print(f"  Next entry day  {next_rebal.date()}{'' if next_confirmed else ' (expected)'}   Status {signal_status}")
print(f"  [Historical reference {first_d.date()}→{LAST_DATE.date()}: CAGR {cagr:.2f}% | "
      f"MaxDD {max_dd_daily:.2f}% | Sharpe {sharpe:.2f}]")
print("═" * 60)

# ── Ranking trend (reporting only): each ETF's rank on recent entry days + now ──
# Ranks are the strategy's own momentum ranks on each entry day: the historical backtest for past
# months and the live record for live months (same rules); "now" = latest close, or the live
# ranking whenever live prices are used (market hours).
RANK_TREND_DAYS = 6
_rank_by_day = {}
for _r in bt["scores"]:
    _rank_by_day.setdefault(pd.Timestamp(_r["rebal_date"]), {})[_r["symbol"]] = _r["rank"]
if live_started:
    for _d in {pd.Timestamp(_r["rebal_date"]) for _r in live["scores"]}:
        _rank_by_day[_d] = {_r["symbol"]: _r["rank"] for _r in live["scores"] if pd.Timestamp(_r["rebal_date"]) == _d}
trend_days = sorted(_rank_by_day)[-RANK_TREND_DAYS:]
rank_trend = []
for _sym in cur_ranked:
    _ranks = [_rank_by_day[_d].get(_sym) for _d in trend_days]
    _last = next((x for x in reversed(_ranks) if x is not None), None)
    _now = cur_rank.get(_sym)
    rank_trend.append(dict(symbol=_sym, category=CATEGORY[_sym], now=_now, ranks=_ranks,
                           change=(_last - _now) if _last and _now else None,
                           held=_sym in live_state["positions"]))

# ═════════════════════════════════════════════════════════════════════════════
# MY ACCOUNT — real trades from the Google Sheet (read-only; never changes the strategy)
# ═════════════════════════════════════════════════════════════════════════════
# The user records every real fill in the Google Sheet "Trades" tab (and optionally deposits etc. in
# "Cash"). A Google service account with VIEWER access reads it. Credentials:
#   GitHub Actions: secrets GSHEET_ID + GOOGLE_SERVICE_ACCOUNT_JSON (passed as environment variables)
#   Local runs:     local_config.json (git-ignored) {"gsheet_id": ...} — the key file is found without
#                   any machine-specific path (see _find_key_file), so the project works from any
#                   drive / folder / laptop.
LOCAL_CONFIG_PATH = BASE_DIR / "local_config.json"
KEY_DIRS = [BASE_DIR / "keys", BASE_DIR.parent / "keys"]   # <project>/keys (git-ignored) or a sibling keys/
XIRR_MIN_DAYS = 30                                # annualising a few days' return is meaningless
MATCH_TOL_PCT = 5.0                               # reporting only: size difference shown as a match
SHEET_EPOCH = datetime(1899, 12, 30)              # Google Sheets serial-date day 0
TRADE_COLS = {"date": "Trade Date", "symbol": "Symbol", "side": "Side", "qty": "Quantity",
              "price": "Price", "charges": "Charges ₹", "month": "Entry Month", "order_id": "Order ID",
              "notes": "Notes"}
CASH_COLS = {"date": "Date", "type": "Type", "amount": "Amount ₹", "notes": "Notes"}
UNIVERSE_SYMBOLS = {s for s, _ in UNIVERSE}


def _load_key(path):
    """Service-account key dict from a JSON file, or None if it is missing / not a service-account key."""
    try:
        info = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return info if isinstance(info, dict) and info.get("type") == "service_account" else None


def _find_key_file(cfg):
    """Locate the service-account key, independent of drive / folder / laptop. First match wins:
      1. env GOOGLE_SERVICE_ACCOUNT_FILE or GOOGLE_APPLICATION_CREDENTIALS (a file path)
      2. local_config.json "service_account_file" — absolute, or relative to the project folder
         (a path that does not exist on this machine is skipped with a warning, not an error)
      3. any service-account key *.json in <project>/keys/ or in a keys/ folder next to the project
    Returns (info, source, places_tried)."""
    tried = []
    for env in ("GOOGLE_SERVICE_ACCOUNT_FILE", "GOOGLE_APPLICATION_CREDENTIALS"):
        if os.environ.get(env):
            tried.append(f"${env}")
            info = _load_key(os.environ[env])
            if info:
                return info, f"${env}", tried
    if cfg.get("service_account_file"):
        p = Path(cfg["service_account_file"])
        p = p if p.is_absolute() else BASE_DIR / p
        tried.append(f"local_config.json → {p}")
        info = _load_key(p)
        if info:
            return info, str(p), tried
        print(f"  [warn] key file from local_config.json not found on this machine: {p} — searching keys/ folders")
    for d in KEY_DIRS:
        tried.append(f"{d}{os.sep}*.json")
        for f in sorted(d.glob("*.json")) if d.is_dir() else []:
            info = _load_key(f)
            if info:
                return info, str(f), tried
    return None, None, tried


def _account_credentials():
    """(sheet_id, service-account info dict, problem). problem is None when both are found,
    "not configured" when nothing is set up, else a message explaining what is missing."""
    cfg = {}
    try:
        cfg = json.loads(LOCAL_CONFIG_PATH.read_text(encoding="utf-8"))
    except FileNotFoundError:
        pass
    except ValueError as e:
        print(f"  [warn] {LOCAL_CONFIG_PATH.name} is not valid JSON ({e})")
    sheet_id = (os.environ.get("GSHEET_ID") or cfg.get("gsheet_id") or "").strip() or None
    if os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON"):                 # GitHub Actions secret
        info, tried = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]), []
    else:
        info, _src, tried = _find_key_file(cfg)
    if not sheet_id and not info:
        return None, None, "not configured"
    if not sheet_id:
        return None, None, (f"Google Sheet ID missing — add \"gsheet_id\" to {LOCAL_CONFIG_PATH.name} "
                            f"(or set GSHEET_ID)")
    if not info:
        return None, None, ("service-account key not found. Put the key .json in "
                            f"{BASE_DIR / 'keys'} (git-ignored) or a keys folder next to the project, or "
                            f"set GOOGLE_SERVICE_ACCOUNT_FILE. Looked in: " + "; ".join(tried))
    return sheet_id, info, None


def read_trade_sheet():
    """Read the Trades and Cash tabs. Returns (trade_rows, cash_rows, error) — rows are lists of dicts
    keyed by column header; error is None on success, a message otherwise. Never raises."""
    try:
        sheet_id, info, problem = _account_credentials()
        if problem:
            return None, None, problem
        import requests
        from google.oauth2 import service_account
        from google.auth.transport.requests import Request
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
        creds.refresh(Request())
        r = requests.get(
            f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values:batchGet",
            params=[("ranges", "Trades!A1:I"), ("ranges", "Cash!A1:D"),
                    ("valueRenderOption", "UNFORMATTED_VALUE"), ("dateTimeRenderOption", "SERIAL_NUMBER")],
            headers={"Authorization": f"Bearer {creds.token}"}, timeout=30)
        if r.status_code != 200:
            return None, None, f"Google Sheets API HTTP {r.status_code}: {r.text[:200]}"
        tabs = []
        for vr in r.json().get("valueRanges", []):
            vals = vr.get("values", [])
            header = [str(h).strip() for h in vals[0]] if vals else []
            rows = []
            for n, row in enumerate(vals[1:], start=2):              # n = sheet row number
                if not any(str(c).strip() for c in row):
                    continue
                d = {h: (row[i] if i < len(row) else "") for i, h in enumerate(header)}
                d["_row"] = n
                rows.append(d)
            tabs.append((header, rows))
        (t_head, t_rows), (c_head, c_rows) = tabs
        missing = [h for h in ("Trade Date", "Symbol", "Side", "Quantity", "Price") if h not in t_head]
        if missing:
            return None, None, f"Trades tab is missing header(s): {', '.join(missing)}"
        return t_rows, c_rows, None
    except Exception as e:                                           # network, key, permission …
        return None, None, f"{type(e).__name__}: {e}"


def _sheet_date(v):
    """Sheets serial number or text → Timestamp (None if empty/invalid)."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return pd.Timestamp(SHEET_EPOCH + timedelta(days=float(v))).normalize()
    t = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d %b %Y", "%d-%b-%Y"):
        try:
            return pd.Timestamp(datetime.strptime(t, fmt))
        except ValueError:
            pass
    return None


def _num(v):
    if v in ("", None):
        return None
    try:
        return float(str(v).replace(",", "").replace("₹", "").strip())
    except ValueError:
        return None


def parse_trades(t_rows):
    """Validate sheet rows → (good trades, issues). Bad rows are listed, never silently used."""
    good, issues = [], []
    for row in t_rows:
        n, errs = row["_row"], []
        d = _sheet_date(row.get(TRADE_COLS["date"], ""))
        sym = str(row.get(TRADE_COLS["symbol"], "")).strip().upper()
        side = str(row.get(TRADE_COLS["side"], "")).strip().upper()
        qty, price = _num(row.get(TRADE_COLS["qty"])), _num(row.get(TRADE_COLS["price"]))
        ch = _num(row.get(TRADE_COLS["charges"]))
        if d is None:
            errs.append("Trade Date missing or not a date")
        if sym not in UNIVERSE_SYMBOLS:
            errs.append(f"unknown Symbol '{sym}'")
        if side not in ("BUY", "SELL"):
            errs.append(f"Side must be BUY or SELL (got '{side}')")
        if qty is None or qty <= 0 or qty != int(qty):
            errs.append("Quantity must be a positive whole number")
        if price is None or price <= 0:
            errs.append("Price must be a positive number")
        if row.get(TRADE_COLS["charges"], "") not in ("", None) and (ch is None or ch < 0):
            errs.append("Charges must be a number ≥ 0")
        m_raw = row.get(TRADE_COLS["month"], "")
        if isinstance(m_raw, (int, float)) and not isinstance(m_raw, bool):
            month = _sheet_date(m_raw).strftime("%Y-%m")         # Sheets turned "2026-10" into a date
        elif str(m_raw).strip():
            month = str(m_raw).strip()[:7]
        else:
            month = d.strftime("%Y-%m") if d is not None else None
        if errs:
            issues.append({"row": n, "problem": "; ".join(errs)})
            continue
        good.append(dict(row=n, date=d, symbol=sym, side=side, qty=int(qty), price=price,
                         charges=ch or 0.0, month=month,
                         order_id=str(row.get(TRADE_COLS["order_id"], "")).strip(),
                         notes=str(row.get(TRADE_COLS["notes"], "")).strip()))
    good.sort(key=lambda t: (t["date"], t["row"]))
    return good, issues


def parse_cash(c_rows):
    good, issues = [], []
    for row in c_rows or []:
        d = _sheet_date(row.get(CASH_COLS["date"], ""))
        typ = str(row.get(CASH_COLS["type"], "")).strip().title()
        amt = _num(row.get(CASH_COLS["amount"]))
        errs = []
        if d is None:
            errs.append("Date missing or not a date")
        if typ not in ("Deposit", "Withdrawal", "Dividend", "Interest"):
            errs.append(f"Type must be Deposit/Withdrawal/Dividend/Interest (got '{typ}')")
        if amt is None or amt < 0:
            errs.append("Amount must be a number ≥ 0")
        if errs:
            issues.append({"row": row["_row"], "problem": "Cash tab: " + "; ".join(errs)})
            continue
        good.append(dict(date=d, type=typ, amount=amt, notes=str(row.get(CASH_COLS["notes"], "")).strip()))
    return good, issues


def _xirr(flows):
    """Annualised money-weighted return (%) of dated cash flows [(date, amount)], investor's view:
    money in = negative, money out / today's value = positive. None if it cannot be solved."""
    flows = [(d, a) for d, a in flows if a]
    if len(flows) < 2 or not any(a > 0 for _, a in flows) or not any(a < 0 for _, a in flows):
        return None
    t0 = min(d for d, _ in flows)

    def npv(r):
        return sum(a / (1 + r) ** ((d - t0).days / 365.0) for d, a in flows)

    lo, hi = -0.9999, 10.0
    f_lo = npv(lo)
    if f_lo * npv(hi) > 0:
        return None
    mid = lo
    for _ in range(200):                         # bisection: robust, no extra packages
        mid = (lo + hi) / 2
        f = npv(mid)
        if abs(f) < 1e-7:
            break
        if f_lo * f < 0:
            hi = mid
        else:
            lo, f_lo = mid, f
    return mid * 100


def _account_daily(accepted, cash_moves, capital_assumed, start):
    """Daily account value from START (cash + holdings at each close) and a TIME-WEIGHTED index
    (deposits/withdrawals do not count as performance). Returns (rows, symbols without prices)."""
    ext = ([(start, float(INITIAL_CAPITAL))] if capital_assumed else
           [(c["date"], c["amount"] if c["type"] == "Deposit" else -c["amount"])
            for c in cash_moves if c["type"] in ("Deposit", "Withdrawal")])
    income = [(c["date"], c["amount"]) for c in cash_moves if c["type"] in ("Dividend", "Interest")]
    rows, missing = [], set()
    prev_cum, prev_v, idx = 0.0, None, 100.0
    for d in cal[(cal >= start) & (cal <= AS_OF)]:
        held, tcash = {}, 0.0
        for t in accepted:
            if t["date"] <= d:
                held[t["symbol"]] = held.get(t["symbol"], 0) + (t["qty"] if t["side"] == "BUY" else -t["qty"])
                tcash += (-(t["qty"] * t["price"] + t["charges"]) if t["side"] == "BUY"
                          else t["qty"] * t["price"] - t["charges"])
        hv = 0.0
        for sym, q in held.items():
            if q:
                if sym in SYMBOLS:
                    hv += q * px(sym, d)
                else:
                    missing.add(sym)
        cum = sum(a for dd, a in ext if dd <= d)
        v = cum + sum(a for dd, a in income if dd <= d) + tcash + hv
        f = cum - prev_cum                           # money added (−withdrawn) since the previous close
        if prev_v is None:
            r = v / f - 1 if f else 0.0              # first day: return on the money put in
        else:
            r = (v - f) / prev_v - 1 if prev_v else 0.0
        idx *= 1 + r
        rows.append(dict(date=d, value=v, flow=f, idx=idx))
        prev_cum, prev_v = cum, v
    return rows, missing


def build_account(trades, cash_moves, issues):
    """Average-cost holdings, realised/unrealised P&L (after charges), cash and plan vs actual."""
    held = {}                                   # sym -> {"qty", "cost"} (cost includes buy charges)
    realised = charges = 0.0
    ignored_rows = set()
    for t in trades:
        h = held.setdefault(t["symbol"], {"qty": 0, "cost": 0.0, "pcost": 0.0, "since": None})
        charges += t["charges"]
        if t["side"] == "BUY":
            if h["qty"] == 0:
                h["since"] = t["date"]          # first buy of the current holding
            h["qty"] += t["qty"]
            h["cost"] += t["qty"] * t["price"] + t["charges"]
            h["pcost"] += t["qty"] * t["price"]  # fill prices only (no charges) — reference stops
        else:
            if t["qty"] > h["qty"]:
                issues.append({"row": t["row"], "problem": f"SELL {t['qty']} {t['symbol']} is more than the "
                                                          f"{h['qty']} held at that point — row ignored"})
                charges -= t["charges"]
                ignored_rows.add(t["row"])
                continue
            avg = h["cost"] / h["qty"]
            realised += t["qty"] * t["price"] - t["charges"] - avg * t["qty"]
            h["pcost"] -= h["pcost"] / h["qty"] * t["qty"]
            h["cost"] -= avg * t["qty"]
            h["qty"] -= t["qty"]
    strat_pos = live_state["positions"]
    entry_days = [m["rebal_date"] for m in live["monthly"]] if live_started else []
    holdings = []
    for sym in sorted(set(held) | set(strat_pos), key=lambda s: (cur_rank.get(s, 999), s)):
        h = held.get(sym, {"qty": 0, "cost": 0.0, "pcost": 0.0, "since": None})
        price = cur_px(sym) if sym in SYMBOLS else None
        value = h["qty"] * price if price is not None else None
        avg = h["cost"] / h["qty"] if h["qty"] else None
        sq = strat_pos[sym]["qty"] if sym in strat_pos else 0
        if h["qty"] == 0 and sq == 0:
            continue
        # Reference stops from YOUR fills (information only — exits follow the strategy levels):
        # hard SL = avg fill (no charges) × 0.85; peak starts at the avg fill and is raised only by
        # entry-day closes AFTER the buy date (rule A); trailing SL = peak × 0.80.
        ref = dict(avg_fill=None, my_hard_sl=None, my_peak=None, my_trail_sl=None,
                   my_dist_sl=None, my_dist_trail=None, my_stop_status=None, my_price_stale=False)
        if h["qty"] and price is not None:
            fill = h["pcost"] / h["qty"]
            peak = max([fill] + [px(sym, ed) for ed in entry_days if ed.date() > h["since"].date()])
            hard, trail = fill * (1 - SL_PCT), peak * (1 - TRAIL_PCT)
            d_sl, d_tr = (price / hard - 1) * 100, (price / trail - 1) * 100
            ref = dict(avg_fill=fill, my_hard_sl=hard, my_peak=peak, my_trail_sl=trail,
                       my_dist_sl=d_sl, my_dist_trail=d_tr,
                       my_stop_status=("Hard SL hit" if price <= hard else "Trail hit" if price <= trail
                                       else "Near stop" if min(d_sl, d_tr) <= NEAR_STOP_PCT else "OK"),
                       my_price_stale=not (LIVE_VIEW and sym in live_quotes) and sym in stale)
        holdings.append(dict(**ref,
            symbol=sym, category=CATEGORY.get(sym, ""), qty=h["qty"], avg_cost=avg, invested=h["cost"],
            price=price, value=value,
            upnl=(value - h["cost"]) if value is not None and h["qty"] else None,
            upnl_pct=((value / h["cost"] - 1) * 100) if value is not None and h["cost"] else None,
            strategy_qty=sq, qty_diff=h["qty"] - sq,
            status=("MATCH" if h["qty"] == sq else "NOT IN STRATEGY" if sq == 0
                    else "NOT HELD" if h["qty"] == 0
                    else "MATCH (size diff)" if abs(h["qty"] / sq - 1) * 100 <= MATCH_TOL_PCT else "QTY DIFF")))
    deposits = sum(c["amount"] for c in cash_moves if c["type"] == "Deposit")
    withdrawals = sum(c["amount"] for c in cash_moves if c["type"] == "Withdrawal")
    income = sum(c["amount"] for c in cash_moves if c["type"] in ("Dividend", "Interest"))
    capital_assumed = deposits == 0
    capital = (deposits - withdrawals) if not capital_assumed else float(INITIAL_CAPITAL)
    trade_cash = sum((-(t["qty"] * t["price"] + t["charges"]) if t["side"] == "BUY"
                      else t["qty"] * t["price"] - t["charges"]) for t in trades
                     if not any(i["row"] == t["row"] for i in issues))
    cash_bal = capital + income + trade_cash
    mv = sum(h["value"] for h in holdings if h["value"] is not None and h["qty"])
    invested_open = sum(h["invested"] for h in holdings if h["qty"])
    upnl = sum(h["upnl"] for h in holdings if h["upnl"] is not None)
    acct_value = mv + cash_bal
    acct_value_now = acct_value
    pnl = acct_value - capital
    # Plan vs actual for every recorded live entry day (strategy plan = recorded at that day's close)
    pva = []
    if live_started:
        for m in live["monthly"]:
            ed, month = m["rebal_date"], m["rebal_date"].strftime("%Y-%m")
            plan = {(a["symbol"], a["action"]): a for a in live["log"]
                    if a["date"] == ed.date() and a["action"] in ("BUY", "SELL")}
            actual = {}
            for t in trades:
                if t["month"] == month:
                    k = (t["symbol"], t["side"])
                    x = actual.setdefault(k, {"qty": 0, "value": 0.0, "charges": 0.0})
                    x["qty"] += t["qty"]
                    x["value"] += t["qty"] * t["price"]
                    x["charges"] += t["charges"]
            for (sym, side) in sorted(set(plan) | set(actual), key=lambda k: (k[1], cur_rank.get(k[0], 999))):
                p, a = plan.get((sym, side)), actual.get((sym, side))
                a_avg = a["value"] / a["qty"] if a else None
                # Reporting only: the order plan sizes from live prices, the plan record from the close, so
                # quantities rarely match exactly — a value within ±MATCH_TOL_PCT % counts as a match.
                _vdiff = (a["value"] / (p["qty"] * p["price"]) - 1) * 100 if p and a and p["qty"] else None
                status = ("NOT DONE" if a is None else "NOT IN PLAN" if p is None
                          else "MATCH" if a["qty"] == p["qty"]
                          else f"MATCH (size {a['qty'] - p['qty']:+d}, value {_vdiff:+.1f}%)"
                          if _vdiff is not None and abs(_vdiff) <= MATCH_TOL_PCT
                          else f"QTY DIFF ({a['qty'] - p['qty']:+d})")
                # Execution vs the strategy close, in ₹ (+ = cost: bought higher / sold lower; − = saved)
                slip = ((a_avg - p["price"]) * a["qty"] * (1 if side == "BUY" else -1)) if p and a else None
                pva.append(dict(month=month, entry_day=ed.date(), symbol=sym, side=side,
                                plan_qty=p["qty"] if p else None, plan_price=p["price"] if p else None,
                                actual_qty=a["qty"] if a else None, actual_avg=a_avg,
                                charges=a["charges"] if a else None,
                                price_diff_pct=((a_avg / p["price"] - 1) * 100) if p and a else None,
                                slippage=slip, traded_value=a["value"] if a else None,
                                status=status))
    latest_month = pva[-1]["month"] if pva else None
    latest = [r for r in pva if r["month"] == latest_month]

    # H — execution cost by entry month (only trades that match a planned action)
    exec_months, cum = [], 0.0
    for month in sorted({r["month"] for r in pva}):
        mr = [r for r in pva if r["month"] == month and r["slippage"] is not None]
        if not mr:
            continue
        val = sum(r["traded_value"] for r in mr)
        slip = sum(r["slippage"] for r in mr)
        chg = sum(r["charges"] or 0.0 for r in mr)
        cum += slip + chg
        exec_months.append(dict(month=month, trades=len(mr), traded_value=val, slippage=slip, charges=chg,
                                total=slip + chg, total_pct=(slip + chg) / val * 100 if val else None,
                                cumulative=cum,
                                unmatched=sum(1 for r in pva if r["month"] == month and r["slippage"] is None)))

    # E — XIRR (money-weighted, annualised) and F — time-weighted comparison since the start
    accepted = [t for t in trades if t["row"] not in ignored_rows]
    deposits_ = sum(c["amount"] for c in cash_moves if c["type"] == "Deposit")
    starts = [t["date"] for t in accepted] + [c["date"] for c in cash_moves if c["type"] == "Deposit"]
    start = min(starts) if starts else None
    xirr = xirr_days = None
    daily, perf, unpriced = [], None, set()
    if start is not None:
        xirr_flows = ([(start, -float(INITIAL_CAPITAL))] if deposits_ == 0 else
                      [(c["date"], -c["amount"] if c["type"] == "Deposit" else c["amount"])
                       for c in cash_moves if c["type"] in ("Deposit", "Withdrawal")])
        xirr_days = (AS_OF - start).days
        xirr = _xirr(xirr_flows + [(AS_OF, acct_value_now)]) if xirr_days >= XIRR_MIN_DAYS else None
        daily, unpriced = _account_daily(accepted, cash_moves, deposits_ == 0, start)
        if daily:
            dts = pd.DatetimeIndex([r["date"] for r in daily])
            strat = (live_daily["value"].reindex(dts, method="ffill") / INITIAL_CAPITAL * 100).fillna(100.0) \
                if live_started else pd.Series(100.0, index=dts)
            nifty = None
            if not bm_raw.empty:
                b = bm_raw.reindex(dts, method="ffill")
                if b.notna().all():
                    nifty = b / b.iloc[0] * 100
            for i, r in enumerate(daily):
                r["strategy_idx"] = float(strat.iloc[i])
                r["nifty_idx"] = float(nifty.iloc[i]) if nifty is not None else None
            last = daily[-1]
            perf = dict(start=start.date(), account_pct=last["idx"] - 100, strategy_pct=last["strategy_idx"] - 100,
                        nifty_pct=(last["nifty_idx"] - 100) if last["nifty_idx"] is not None else None)

    return dict(
        xirr=xirr, xirr_days=xirr_days, perf=perf, daily=daily, unpriced=sorted(unpriced),
        exec_months=exec_months, exec_total=cum if exec_months else None,
        holdings=holdings, pva=pva, latest_month=latest_month,
        latest_mismatches=sum(not r["status"].startswith("MATCH") for r in latest),
        latest_entered=any(r["actual_qty"] for r in latest),
        capital=capital, capital_assumed=capital_assumed, deposits=deposits, withdrawals=withdrawals,
        income=income, cash=cash_bal, market_value=mv, invested=invested_open, upnl=upnl,
        realised=realised, charges=charges, value=acct_value, pnl=pnl,
        ret_pct=pnl / capital * 100 if capital else None)


_t_rows, _c_rows, account_error = read_trade_sheet()
account = None
account_trades, account_cash, account_issues = [], [], []
if account_error is None:
    account_trades, account_issues = parse_trades(_t_rows)
    account_cash, _cash_issues = parse_cash(_c_rows)
    account_issues += _cash_issues
    account = build_account(account_trades, account_cash, account_issues)
    # Live performance since the start (market hours): the closing-price index extended by today's
    # live values — the same numbers as the ★ live point on the chart. Never stored in the history.
    account["perf_live"] = None
    if LIVE_VIEW and account["daily"] and account["perf"]:
        _ld = account["daily"][-1]
        _ai = _ld["idx"] * (account["value"] / _ld["value"]) if _ld["value"] else None
        _ni = (_ld["nifty_idx"] * bm_live / float(bm_raw.iloc[-1])
               if bm_live and _ld["nifty_idx"] is not None and not bm_raw.empty else None)
        _si = live_value / INITIAL_CAPITAL * 100 if live_started else None
        account["perf_live"] = dict(start=account["perf"]["start"], time=live_quote_time,
                                    account_idx=_ai, strategy_idx=_si, nifty_idx=_ni,
                                    account_pct=_ai - 100 if _ai is not None else None,
                                    strategy_pct=_si - 100 if _si is not None else None,
                                    nifty_pct=_ni - 100 if _ni is not None else None)
    print(f"  [MY ACCOUNT] Google Sheet: {len(account_trades)} trade row(s), {len(account_cash)} cash row(s), "
          f"{len(account_issues)} problem row(s)"
          + ("" if IN_CI else f" · value ₹{account['value']:,.0f} ({account['ret_pct']:+.2f}%)"))
elif account_error == "not configured":
    print("  [MY ACCOUNT] Google Sheet not configured — skipped")
else:
    print(f"  [MY ACCOUNT] Could not read the Google Sheet: {account_error}")

# ═════════════════════════════════════════════════════════════════════════════
# EXCEL OUTPUT — Live_* sheets first, historical reference as Hist_* sheets
# ═════════════════════════════════════════════════════════════════════════════
live_summary_df = pd.DataFrame([
    ("Live Start Date", str(live_init_date.date()) if live_started else f"{LIVE_START_DATE} (not started)"),
    ("As Of", str(AS_OF.date())),
    ("Initial Capital", INITIAL_CAPITAL),
    ("Current Portfolio Value", round(live_value, 2)),
    ("Invested Value", round(invested, 2)),
    ("Cash", round(live_cash, 2)),
    ("Cash %", round(live_cash / live_value * 100, 2)),
    ("Total P&L", round(live_pnl, 2)),
    ("Return %", round(live_ret, 2)),
    ("Positions", len(pos_rows)),
    ("Live Max DD %", round(live_max_dd, 2) if live_max_dd is not None else "n/a"),
    ("Winning Months", win_months),
    ("Losing Months", lose_months),
    ("Closed Trades", n_lt),
    ("Win Rate %", round(lt_stats["wr"], 2) if lt_stats["wr"] is not None else "n/a"),
    ("Realised P&L", round(lt_stats["realised"], 2)),
    ("Last Entry Day", str(cur_rebal.date()) if cur_rebal is not None else "n/a"),
    ("Latest Market Data", str(AS_OF.date())),
    ("Next Entry Day", f"{next_rebal.date()}{'' if next_confirmed else ' (expected)'}"),
    ("Status", signal_status),
    ("Decision source", signal_source),
    ("Methodology", "Entry day = first trading day of the month: signal and execution at that day's close "
                    "(order plan from live prices during market hours)"),
    ("Data Status", data_status + ("" if data_ok else
                                   f" — {len(stale)} of {len(SYMBOLS)} ETFs without a {AS_OF.date()} price"
                                   + (f"; market data older than expected {expected_latest.date()}"
                                      if market_delayed else ""))),
    ("Delayed ETFs (held)", ", ".join(stale_held) or "none"),
    ("Delayed ETFs (current top 6)", ", ".join(stale_signal) or "none"),
], columns=["Metric", "Value"])

live_data_xl = pd.DataFrame(
    [{"Symbol": s, "Category": CATEGORY[s], "Last Price Date": last_px_date[s].date(),
      "Sessions Behind": sessions_behind[s], "Status": "DELAYED" if s in stale else "CURRENT",
      "Held": "YES" if s in live_state["positions"] else "NO", "Current Rank": cur_rank.get(s)}
     for s in sorted(SYMBOLS, key=lambda s: (-sessions_behind[s], s))]
    + [{"Symbol": s, "Category": CATEGORY[s], "Last Price Date": None, "Sessions Behind": None,
        "Status": f"EXCLUDED — {n} of {MIN_HISTORY} rows on Yahoo", "Held": "NO", "Current Rank": None}
       for s, n in dropped_info.items()])

live_pos_xl = pd.DataFrame([{
    "Rank": r["rank"], "Symbol": r["symbol"], "Category": r["category"],
    "Momentum Score": r["score"], "1M %": r["r1"], "3M %": r["r3"], "6M %": r["r6"], "12M %": r["r12"],
    "200 DMA": r["dma"], "Current Price": r["price"], "Above DMA": "YES" if r["above_dma"] else "NO",
    "Entry Date": r["entry_date"], "Entry Price": r["entry_price"], "Qty": r["qty"],
    "Cost": r["cost"], "Current Value": r["value"], "Unrealised P&L": r["pnl"],
    "Unrealised P&L %": r["pnl_pct"], "Hold Days": r["hold_days"], "Peak Price": r["peak"],
    "Hard SL": r["hard_sl"], "Trailing SL": r["trail_sl"], "Distance to SL %": r["dist_sl"],
    "Distance to Trail %": r["dist_trail"], "Position Weight %": r["weight"], "Status": r["status"],
    "Price Date": r["price_date"], "Data": "DELAYED" if r["behind"] else "CURRENT",
} for r in pos_rows]).round(4)

live_sig_xl = pd.DataFrame([{
    "As Of": AS_OF.date(), "Rank": r["rank"], "Symbol": r["symbol"], "Category": r["category"],
    "Momentum Score": r["score"], "1M %": r["r1"], "3M %": r["r3"], "6M %": r["r6"], "12M %": r["r12"],
    "200 DMA": r["dma"], "Current Price": r["price"], "Above 200 DMA": "YES" if r["above_dma"] else "NO",
    "Currently Held": "YES" if r["held"] else "NO", "Signal": r["signal"],
    "Target Allocation": r["target"], "Est. Qty": r["est_qty"],
    "Price Date": r["price_date"], "Data": "DELAYED" if r["behind"] else "CURRENT",
} for r in sig_rows]).round(4)

if live_started:
    live_eq_xl = pd.DataFrame({
        "Date": live_daily.index.date, "Portfolio Value": live_daily["value"].round(2),
        "Invested Value": live_daily["held"].round(2), "Cash": live_daily["cash"].round(2),
        "Daily Return %": live_daily["daily_ret_pct"].round(3),
        "Cumulative Return %": live_daily["cum_ret_pct"].round(3),
        "Drawdown %": live_daily["dd_pct"].round(3)})
    live_log_xl = pd.DataFrame(live["log"])
    live_monthly_xl = pd.DataFrame(live["monthly"]).drop(columns=["port_state"])
    live_monthly_xl["rebal_date"] = live_monthly_xl["rebal_date"].dt.date
    live_monthly_xl["signal_date"] = live_monthly_xl["signal_date"].dt.date
else:
    live_eq_xl = pd.DataFrame(columns=["Date", "Portfolio Value", "Invested Value", "Cash",
                                       "Daily Return %", "Cumulative Return %", "Drawdown %"])
    live_log_xl = pd.DataFrame(columns=["date", "action", "symbol"])
    live_monthly_xl = pd.DataFrame(columns=["rebal_date"])
live_act_xl = pd.DataFrame(act_log)
if not live_act_xl.empty:
    live_act_xl.insert(0, "mode", act_mode)

# Historical reference sheets (unchanged content, renamed Hist_*)
summary_rows = [dict(metric=k, value=v) for k, v in metrics]
summary_rows.append(dict(metric="", value=""))
summary_rows.append(dict(metric="YEAR-BY-YEAR RETURNS", value=""))
for y, r in annual.items():
    bmv = f" | BM {bm_annual.get(y):+.2f}%" if bm_annual is not None and y in bm_annual.index else ""
    summary_rows.append(dict(metric=str(y), value=f"{r:+.2f}%{bmv}"))
summary_df = pd.DataFrame(summary_rows)

mp_cols = ["signal_date", "rebal_date", "month_label", "portfolio_value", "cash", "n_held", "monthly_return_pct",
           "cumulative_return_pct", "drawdown_pct", "holdings", "exits", "entries"]
mp_out = mdf[mp_cols].copy()
mp_out["rebal_date"] = mp_out["rebal_date"].dt.date
mp_out["signal_date"] = mp_out["signal_date"].dt.date
mp_out = mp_out.rename(columns={"monthly_return_pct": "monthly_return%",
                                "cumulative_return_pct": "cumulative_return%",
                                "drawdown_pct": "drawdown%"}).round(3)
scores_df = pd.DataFrame(bt["scores"])

GREEN = PatternFill("solid", fgColor="C6EFCE")
RED = PatternFill("solid", fgColor="FFC7CE")
AMBERF = PatternFill("solid", fgColor="FFE699")


def color_rows(ws, df, fill_fn):
    for r_i, row in enumerate(ws.iter_rows(min_row=2, max_row=len(df) + 1)):
        fill = fill_fn(df.iloc[r_i])
        if fill is not None:
            for c in row:
                c.fill = fill


def _or_empty(df, col):
    return df if not df.empty else pd.DataFrame(columns=[col])


try:                                   # the workbook is locked while it is open in Excel
    with open(EXCEL_PATH, "ab"):
        pass
except PermissionError:
    _locked = EXCEL_PATH
    EXCEL_PATH = EXCEL_PATH.with_name(f"{EXCEL_PATH.stem}_{now_ist():%Y%m%d_%H%M%S}{EXCEL_PATH.suffix}")
    print(f"  [warn] {_locked.name} is open in another program — saving to {EXCEL_PATH.name} instead")

# My Account sheets (only when the Google Sheet was read successfully)
my_xl = {}
if account is not None:
    my_xl["My_Account"] = pd.DataFrame([
        ("Source", "Google Sheet 'ETF Momentum – My Trades' (read-only)"),
        ("Read at", now_ist().strftime("%Y-%m-%d %H:%M IST")),
        ("Prices as of", str(AS_OF.date())),
        ("Capital", round(account["capital"], 2)),
        ("Capital basis", "assumed INITIAL_CAPITAL (no Deposit rows in Cash tab)" if account["capital_assumed"]
         else "Cash tab deposits − withdrawals"),
        ("Account Value", round(account["value"], 2)),
        ("Total P&L", round(account["pnl"], 2)),
        ("Return %", round(account["ret_pct"], 2)),
        ("Market Value (holdings)", round(account["market_value"], 2)),
        ("Invested (open cost incl. charges)", round(account["invested"], 2)),
        ("Unrealised P&L", round(account["upnl"], 2)),
        ("Realised P&L", round(account["realised"], 2)),
        ("Total Charges", round(account["charges"], 2)),
        ("Cash", round(account["cash"], 2)),
        ("Dividends + Interest", round(account["income"], 2)),
        ("XIRR % (annualised)", round(account["xirr"], 2) if account["xirr"] is not None
         else f"shown after {XIRR_MIN_DAYS} days ({account['xirr_days'] or 0} so far)"),
        ("Performance since", str(account["perf"]["start"]) if account["perf"] else "n/a"),
        ("My account % (time-weighted)", round(account["perf"]["account_pct"], 2) if account["perf"] else "n/a"),
        ("Strategy record %", round(account["perf"]["strategy_pct"], 2) if account["perf"] else "n/a"),
        ("Nifty 500 %", round(account["perf"]["nifty_pct"], 2)
         if account["perf"] and account["perf"]["nifty_pct"] is not None else "n/a"),
        ("Execution vs close ₹ (+cost / −saved)", round(account["exec_total"], 2)
         if account["exec_total"] is not None else "n/a"),
        ("Trade rows read", len(account_trades)),
        ("Problem rows", len(account_issues)),
    ], columns=["Metric", "Value"])
    my_xl["My_Holdings"] = pd.DataFrame([{
        "Symbol": h["symbol"], "Category": h["category"], "My Qty": h["qty"],
        "Avg Cost (incl. charges)": h["avg_cost"], "Invested": h["invested"], "Price": h["price"],
        "Value": h["value"], "Unrealised P&L": h["upnl"], "Unrealised P&L %": h["upnl_pct"],
        "Strategy Qty": h["strategy_qty"], "Qty Diff": h["qty_diff"], "Status": h["status"],
        "Avg Fill (excl. charges)": h["avg_fill"], "My Peak (ref)": h["my_peak"],
        "My Hard SL (ref)": h["my_hard_sl"], "My Trailing SL (ref)": h["my_trail_sl"],
        "My Distance to Hard SL %": h["my_dist_sl"], "My Distance to Trail %": h["my_dist_trail"],
        "My Stop Status (ref)": h["my_stop_status"],
    } for h in account["holdings"]]).round(4)
    my_xl["My_Plan_vs_Actual"] = pd.DataFrame([{
        "Month": r["month"], "Entry Day": r["entry_day"], "Symbol": r["symbol"], "Side": r["side"],
        "Plan Qty": r["plan_qty"], "Plan Price (close)": r["plan_price"], "Actual Qty": r["actual_qty"],
        "Actual Avg Price": r["actual_avg"], "Charges": r["charges"], "Price Diff %": r["price_diff_pct"],
        "₹ vs Close (+cost/−saved)": r["slippage"], "Status": r["status"],
    } for r in account["pva"]]).round(4)
    my_xl["My_Trades"] = pd.DataFrame([{
        "Sheet Row": t["row"], "Trade Date": t["date"].date(), "Symbol": t["symbol"], "Side": t["side"],
        "Quantity": t["qty"], "Price": t["price"], "Charges": t["charges"], "Entry Month": t["month"],
        "Order ID": t["order_id"], "Notes": t["notes"], "Value": t["qty"] * t["price"],
    } for t in account_trades] + [{"Sheet Row": i["row"], "Notes": "PROBLEM: " + i["problem"]}
                                   for i in account_issues])
    my_xl["My_Performance"] = pd.DataFrame([{
        "Date": r["date"].date(), "Account Value": r["value"], "Money In (−Out)": r["flow"],
        "My Account Index": r["idx"], "Strategy Index": r["strategy_idx"], "Nifty 500 Index": r["nifty_idx"],
    } for r in account["daily"]]).round(4)
    my_xl["My_Execution_Cost"] = pd.DataFrame([{
        "Month": r["month"], "Trades": r["trades"], "Traded Value": r["traded_value"],
        "Price vs Close ₹": r["slippage"], "Charges ₹": r["charges"], "Total ₹ (+cost/−saved)": r["total"],
        "Total % of Traded": r["total_pct"], "Cumulative ₹": r["cumulative"], "Unmatched rows": r["unmatched"],
    } for r in account["exec_months"]]).round(4)
    my_xl["My_Cash"] = pd.DataFrame([{"Date": c["date"].date(), "Type": c["type"], "Amount": c["amount"],
                                      "Notes": c["notes"]} for c in account_cash])


def _write_excel():
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as xw:
        live_summary_df.to_excel(xw, sheet_name="Live_Summary", index=False)
        _or_empty(live_pos_xl, "Symbol").to_excel(xw, sheet_name="Live_Positions", index=False)
        _or_empty(live_act_xl, "action").to_excel(xw, sheet_name="Live_Actions", index=False)
        live_sig_xl.to_excel(xw, sheet_name="Live_Signals", index=False)
        live_data_xl.to_excel(xw, sheet_name="Live_Data_Status", index=False)
        pd.DataFrame([{"Symbol": r["symbol"], "Category": r["category"],
                       **{str(d.date()): rk for d, rk in zip(trend_days, r["ranks"])},
                       ("Now (live)" if order_plan else f"Now ({AS_OF.date()})"): r["now"],
                       "Change vs last entry day (+ = up)": r["change"], "Held": "YES" if r["held"] else "NO"}
                      for r in rank_trend]).apply(
            lambda c: c.astype("Int64") if c.dtype.kind == "f" else c      # whole-number ranks, blanks kept
        ).to_excel(xw, sheet_name="Rank_Trend", index=False)
        for _name, _df in my_xl.items():
            _or_empty(_df, "Notes").to_excel(xw, sheet_name=_name, index=False)
        _or_empty(live_trades_df, "trade_id").to_excel(xw, sheet_name="Live_Closed_Trades", index=False)
        live_eq_xl.to_excel(xw, sheet_name="Live_Equity", index=False)
        _or_empty(live_monthly_xl, "rebal_date").to_excel(xw, sheet_name="Live_Rebalances", index=False)
        _or_empty(live_log_xl, "action").to_excel(xw, sheet_name="Live_Rebalance_Log", index=False)
        summary_df.to_excel(xw, sheet_name="Hist_Summary", index=False)
        mp_out.to_excel(xw, sheet_name="Hist_Monthly_Portfolio", index=False)
        _or_empty(trades_df, "trade_id").to_excel(xw, sheet_name="Hist_Closed_Trades", index=False)
        _or_empty(open_df, "symbol").to_excel(xw, sheet_name="Hist_Open_Positions", index=False)
        scores_df.to_excel(xw, sheet_name="Hist_Momentum_Scores", index=False)

        if not live_pos_xl.empty:
            color_rows(xw.sheets["Live_Positions"], live_pos_xl,
                       lambda r: AMBERF if r["Status"] != "HOLD" else (GREEN if r["Unrealised P&L"] > 0 else RED))
        color_rows(xw.sheets["Live_Signals"], live_sig_xl,
                   lambda r: GREEN if r["Signal"] in ("HOLD", "NEW ENTRY") else
                   (AMBERF if r["Signal"] in ("ROTATION RISK", "SKIP — BELOW 200 DMA") else None))
        color_rows(xw.sheets["Live_Data_Status"], live_data_xl,
                   lambda r: None if r["Status"] == "CURRENT" else AMBERF if r["Status"] == "DELAYED" else RED)
        if not live_trades_df.empty:
            color_rows(xw.sheets["Live_Closed_Trades"], live_trades_df,
                       lambda r: GREEN if r["pnl_inr"] > 0 else RED)
        if not trades_df.empty:
            color_rows(xw.sheets["Hist_Closed_Trades"], trades_df,
                       lambda r: GREEN if r["pnl_inr"] > 0 else RED)
        for ws in xw.sheets.values():
            for c in ws[1]:
                c.font = Font(bold=True)
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 60)
            ws.freeze_panes = "A2"


try:
    _write_excel()
except Exception as e:
    print(f"[EXCEL ERROR] Could not write {EXCEL_PATH}: {e}")
    sys.exit(3)
print(f"Saved: {EXCEL_PATH.name}")
print(f"  → {EXCEL_PATH}")

# ═════════════════════════════════════════════════════════════════════════════
# HTML — LIVE PORTFOLIO DASHBOARD
# ═════════════════════════════════════════════════════════════════════════════
BG, PANEL, TXT, MUTED = "#0D1B2A", "#1B2A3A", "#F0F0F0", "#9AA8B6"
AMBER, GRAY, GRN, RD, BLUE = "#F5A623", "#9E9E9E", "#2ECC71", "#E74C3C", "#5DADE2"

# Plotly.js is loaded once in <head> from a CDN (same version as the Python package, pinned with an
# integrity hash; second CDN as fallback). Embedding it made the page ~5 MB, and the encrypted page
# (not compressible) ~6.7 MB on every visit; from a CDN it is downloaded once and cached.
PLOTLY_JS_VERSION = pyo.get_plotlyjs_version()
# The bundled plotly.min.js is byte-identical to the CDN file → its hash is the integrity check
PLOTLY_SRI = "sha384-" + base64.b64encode(hashlib.sha384(pyo.get_plotlyjs().encode("utf-8")).digest()).decode()
PLOTLY_CDNS = [f"https://cdnjs.cloudflare.com/ajax/libs/plotly.js/{PLOTLY_JS_VERSION}/plotly.min.js",
               f"https://cdn.jsdelivr.net/npm/plotly.js-dist-min@{PLOTLY_JS_VERSION}/plotly.min.js"]


def fig_to_div(fig):
    cfg = {"responsive": True, "displaylogo": False, "scrollZoom": False}
    return pyo.plot(fig, include_plotlyjs=False, output_type="div", config=cfg)


def style(fig, title, h=400):
    fig.update_layout(title=dict(text=title, x=0.01, xanchor="left", font=dict(size=15)),
                      template="plotly_dark", paper_bgcolor=PANEL, plot_bgcolor=PANEL,
                      font=dict(color=TXT), height=h, margin=dict(l=55, r=20, t=50, b=40),
                      legend=dict(orientation="h", yanchor="top", y=-0.14, x=0))   # below the chart
    return fig


def _isnum(v):
    return isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool)


def inr(v, dec=0):
    """₹ with Indian digit grouping (₹3,00,000)."""
    if v is None or (_isnum(v) and math.isnan(v)):
        return "—"
    neg, s = v < 0, f"{abs(v):.{dec}f}"
    whole, _, frac = s.partition(".")
    if len(whole) > 3:
        head, tail, parts = whole[:-3], whole[-3:], []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        whole = ",".join(parts) + "," + tail
    return ("-" if neg else "") + "₹" + whole + ("." + frac if frac else "")


def pct(v, sign=True):
    if v is None or (_isnum(v) and math.isnan(v)):
        return "—"
    return f"{v:+.2f}%" if sign else f"{v:.2f}%"


def num(v, d=2):
    return "—" if v is None or (_isnum(v) and math.isnan(v)) else f"{v:,.{d}f}"


def yn(v):
    return "YES" if v else "NO"


def table(rows, cols, tid, row_cls=None, cell_cls=None):
    """Sortable table. cols = [(key, header, formatter)]."""
    if not rows:
        return "<p class='muted'>None.</p>"
    th = "".join(f"<th onclick=\"sortTable('{tid}',{j})\">{h}</th>" for j, (_, h, _) in enumerate(cols))
    body = []
    for r in rows:
        tds = []
        for k, _, f in cols:
            v = r.get(k)
            txt = f(v) if f else ("—" if v is None else str(v))
            sv = f"{float(v):.6f}" if _isnum(v) and not math.isnan(v) else ("" if v is None else str(v))
            cc = cell_cls(k, r) if cell_cls else ""
            if k == "price_date" and r.get("behind"):
                cc += " stale"                           # delayed price → amber cell with ⚠
            tds.append(f"<td class='{cc}' data-v='{htmlmod.escape(sv, quote=True)}'>{txt}</td>")
        rc = row_cls(r) if row_cls else ""
        body.append(f"<tr class='{rc}'>{''.join(tds)}</tr>")
    return (f"<div class='tw'><table id='{tid}'><thead><tr>{th}</tr></thead>"
            f"<tbody>{''.join(body)}</tbody></table></div>")


def pn(v):
    return "g" if _isnum(v) and v > 0 else ("r" if _isnum(v) and v < 0 else "")


def kpi(label, value, cls=""):
    return f"<div class='kpi'><div class='kl'>{label}</div><div class='kv {cls}'>{value}</div></div>"


def badge(text):
    t = text.upper()
    cls = ("b-buy" if t in ("NEW ENTRY", "BUY") else
           "b-hold" if t == "HOLD" else
           "b-sell" if t.startswith("EXIT RISK —") and ("HIT" in t) or t == "SELL" else
           "b-risk" if t.startswith("EXIT RISK") or t == "ROTATION RISK" else
           "b-skip" if t.startswith("SKIP") else "b-out")
    return f"<span class='badge {cls}'>{text}</span>"


def pxdate(v):
    """Price source: 🟢 live quote, ✓ latest close, or ⚠ an older close (price carried forward)."""
    if isinstance(v, str):
        return f"🟢 {v}"
    if v < AS_OF.date():
        return f"⚠ {v} ({sessions_behind_by_date(v)} behind)"
    return f"✓ {v}" + (" (no live quote)" if LIVE_VIEW else "")


def sessions_behind_by_date(v):
    n = int(((cal > pd.Timestamp(v)) & (cal <= AS_OF)).sum())
    return f"{n} session" + ("" if n == 1 else "s")


# ── Data freshness banner ────────────────────────────────────────────────────
if data_ok:
    data_html = (f"<div class='dataok'>✓ Data current — all {len(SYMBOLS)} ETFs have prices for "
                 f"{AS_OF.date()}.</div>")
else:
    _lines = []
    if stale:
        _lines.append(f"<b>{len(stale)} of {len(SYMBOLS)} ETFs have no price for {AS_OF.date()}</b> "
                      f"(latest market date). Their last available close is used instead:")
        _lines.append("<ul>" + "".join(
            f"<li>Last price <b>{d.date()}</b> ({sessions_behind_by_date(d.date())} behind): "
            f"{', '.join(syms)}</li>" for d, syms in stale_groups.items()) + "</ul>")
    if market_delayed:
        _lines.append(f"Latest market data is <b>{AS_OF.date()}</b>, but the most recent weekday session "
                      f"was <b>{expected_latest.date()}</b>: Yahoo data may be delayed (or it was an "
                      f"exchange holiday).")
    if stale_held:
        _lines.append(f"⚠ <b>Your positions affected:</b> {', '.join(stale_held)} — values, P&L and stop "
                      f"distances use an older price.")
    if stale_signal or stale_actions:
        _lines.append(f"⚠ <b>Ranking affected:</b> {', '.join(sorted(set(stale_signal) | set(stale_actions)))} "
                      f"— ranking and actions may change once the missing prices arrive.")
    _lines.append("Rerun later with <code>--force</code> to refresh once Yahoo has the missing prices.")
    data_html = ("<div class='datawarn'><div class='dw-title'>⚠ DATA DELAYED</div>"
                 + "".join(f"<p>{l}</p>" if not l.startswith("<ul>") else l for l in _lines) + "</div>")
if near_miss:
    data_html += ("<p class='muted'>Excluded for short Yahoo history, only a few rows below the "
                  f"{MIN_HISTORY}-row minimum (a delayed day can cause this): "
                  + ", ".join(f"{s} ({n}/{MIN_HISTORY})" for s, n in near_miss.items()) + "</p>")
act_data_warn = ("" if not (stale_actions or stale_signal or market_delayed) else
                 "<div class='datawarn slim'>⚠ Some prices behind these actions are delayed — "
                 "see the data warning at the top. Actions may change after the data catches up.</div>")

# ── My Account (real trades from the Google Sheet) ──────────────────────────
if account is None:
    account_html = ("<p class='muted'>" + ("Google Sheet not connected (add GSHEET_ID and "
                    "GOOGLE_SERVICE_ACCOUNT_JSON secrets, or local_config.json locally)." if account_error == "not configured"
                    else f"⚠ Could not read the Google Sheet: {htmlmod.escape(str(account_error))}") + "</p>")
else:
    _a = account
    _acct_kpis = "".join([
        kpi("Account Value", inr(_a["value"])),
        kpi("Total P&L", inr(_a["pnl"]), pn(_a["pnl"])),
        kpi("Return %", pct(_a["ret_pct"]), pn(_a["ret_pct"])),
        kpi("Market Value", inr(_a["market_value"])),
        kpi("Invested (open)", inr(_a["invested"])),
        kpi("Unrealised P&L", inr(_a["upnl"]), pn(_a["upnl"])),
        kpi("Realised P&L", inr(_a["realised"]), pn(_a["realised"])),
        kpi("Charges Paid", inr(_a["charges"], 2)),
        kpi("Cash", inr(_a["cash"])),
        kpi("XIRR (annualised)", pct(_a["xirr"]) if _a["xirr"] is not None else
            (f"after {XIRR_MIN_DAYS} days" if _a["xirr_days"] is not None else "—"), pn(_a["xirr"])),
        kpi("vs Nifty 500" + (f" since {_pf['start']}" if (_pf := _a.get("perf_live") or _a["perf"]) else ""),
            f"{_pf['account_pct'] - _pf['nifty_pct']:+.2f} pp"
            if _pf and _pf["account_pct"] is not None and _pf["nifty_pct"] is not None else "—",
            pn(_pf["account_pct"] - _pf["nifty_pct"])
            if _pf and _pf["account_pct"] is not None and _pf["nifty_pct"] is not None else ""),
        kpi("Execution vs close", (("saved " if _a["exec_total"] < 0 else "cost ") + inr(abs(_a["exec_total"])))
            if _a["exec_total"] is not None else "—", pn(-_a["exec_total"]) if _a["exec_total"] is not None else ""),
    ])
    _hold_cols = [
        ("symbol", "Symbol", None), ("category", "Category", None), ("qty", "My Qty", str),
        ("avg_cost", "Avg Cost (incl. charges)", num), ("invested", "Invested", inr),
        ("price", f"Price ({'live ' + live_quote_time if LIVE_VIEW else AS_OF.date()})", num), ("value", "Value", inr), ("upnl", "Unrealised P&L ₹", inr),
        ("upnl_pct", "Unrealised P&L %", pct), ("strategy_qty", "Strategy Qty", str),
        ("qty_diff", "Diff", lambda v: f"{v:+d}" if v else "0"), ("status", "Status", None),
    ]
    _pva_cols = [
        ("symbol", "Symbol", None), ("month", "Month", None), ("side", "Side", None),
        ("plan_qty", "Plan Qty", lambda v: "—" if v is None else str(v)),
        ("plan_price", "Plan Price (close)", num),
        ("actual_qty", "Actual Qty", lambda v: "—" if v is None else str(v)),
        ("actual_avg", "Actual Avg Price", num), ("charges", "Charges ₹", lambda v: inr(v, 2)),
        ("price_diff_pct", "Price vs Close", pct), ("slippage", "₹ vs Close", inr),
        ("status", "Status", None),
    ]
    _ref_cols = [
        ("symbol", "Symbol", None), ("qty", "My Qty", str), ("avg_fill", "Avg Fill (excl. charges)", num),
        ("price", f"Price ({'live ' + live_quote_time if LIVE_VIEW else AS_OF.date()})", num),
        ("my_peak", "My Peak", num), ("my_hard_sl", "My Hard SL", num), ("my_trail_sl", "My Trailing SL", num),
        ("my_dist_sl", "Distance to Hard SL %", lambda v: pct(v, False)),
        ("my_dist_trail", "Distance to Trail %", lambda v: pct(v, False)),
        ("my_stop_status", "Status (reference)", None),
    ]
    _ref_rows = [dict(h, my_stop_status=h["my_stop_status"] + (" · price stale" if h["my_price_stale"] else ""))
                 for h in _a["holdings"] if h["qty"] and h["avg_fill"] is not None]
    _ok = lambda r: "pos" if r["status"].startswith("MATCH") else "warn"
    _issues_html = ("" if not account_issues else
                    "<div class='datawarn slim'>⚠ <b>Sheet rows with problems</b> (not used until fixed): "
                    + "; ".join(f"row {i['row']}: {htmlmod.escape(i['problem'])}" for i in account_issues) + "</div>")
    _cap_note = (" Capital assumed " + inr(_a["capital"]) + " — add a Deposit row in the Cash tab to set it."
                 if _a["capital_assumed"] else "")
    _latest_note = ("" if not _a["latest_month"] else
                    f" Latest entry month {_a['latest_month']}: "
                    + ("no trades entered yet." if not _a["latest_entered"] else
                       "✓ matches the strategy plan." if _a["latest_mismatches"] == 0 else
                       f"⚠ {_a['latest_mismatches']} difference(s) from the strategy plan."))
    account_html = f"""
<p class='muted'>From your Google Sheet: {len(account_trades)} trade row(s), {len(account_cash)} cash row(s), read
{now_ist():%Y-%m-%d %H:%M} IST. Prices = {f"live {live_quote_time} IST" if LIVE_VIEW else f"{AS_OF.date()} close"}; P&L after charges (average-cost method).{_cap_note}{_latest_note}</p>
{_issues_html}
<div class="kpis">{_acct_kpis}</div>
<h3>My holdings vs strategy</h3>
{table(_a["holdings"], _hold_cols, "tMyHold", _ok, lambda k, r: pn(r[k]) if k in ("upnl", "upnl_pct") else ("sym" if k == "symbol" else ""))}
<h3>My stop levels from my fills (reference only)</h3>
<p class='muted'><b>Exits follow the strategy levels</b> (Current Live Positions / Risk Monitor) — these are for
reference only. My Hard SL = my average fill (excl. charges) × {1 - SL_PCT:.2f}; My Peak starts at my average fill and
is raised only by entry-day closes after my buy date; My Trailing SL = My Peak × {1 - TRAIL_PCT:.2f}.</p>
{table(_ref_rows, _ref_cols, "tMyStops", lambda r: "warn" if r["my_stop_status"] != "OK" else "",
       lambda k, r: ("r" if r[k] <= 0 else "amber" if r[k] <= NEAR_STOP_PCT else "")
       if k in ("my_dist_sl", "my_dist_trail") else ("sym" if k == "symbol" else ""))}
<h3>Plan vs actual (each entry day)</h3>
<p class='muted'>Plan = the strategy record at the entry-day close; actual = your fills for that month in the sheet.
Your order-plan quantities come from live prices, so they rarely equal the close-based plan exactly: the same ETF and
side with a value within ±{MATCH_TOL_PCT:.0f}% of the plan shows as <b>MATCH (size …)</b>.</p>
{table(_a["pva"], _pva_cols, "tPva", _ok, lambda k, r: pn(-r[k]) if k == "price_diff_pct" and r.get("side") == "BUY" and r[k] is not None else ("sym" if k == "symbol" else ""))}"""

if account is not None:
    _a = account
    if _a["daily"]:
        _dd = [r["date"] for r in _a["daily"]]
        f_acct = go.Figure()
        f_acct.add_trace(go.Scatter(x=_dd, y=[r["idx"] for r in _a["daily"]], name="My account",
                                    mode="lines+markers", line=dict(color=GRN, width=2.5)))
        f_acct.add_trace(go.Scatter(x=_dd, y=[r["strategy_idx"] for r in _a["daily"]], name="Strategy record",
                                    mode="lines+markers", line=dict(color=AMBER, width=2)))
        if _a["daily"][0]["nifty_idx"] is not None:
            f_acct.add_trace(go.Scatter(x=_dd, y=[r["nifty_idx"] for r in _a["daily"]], name="Nifty 500",
                                        mode="lines+markers", line=dict(color=GRAY, dash="dash")))
        _pl = _a.get("perf_live")
        if _pl:                                    # separate, clearly marked live point (not history)
            _lx = [pd.Timestamp(now_ist())]
            for _y, _n, _c in ((_pl["account_idx"], "My account (live)", GRN),
                               (_pl["strategy_idx"], "Strategy (live)", AMBER),
                               (_pl["nifty_idx"], "Nifty 500 (live)", GRAY)):
                if _y is not None:
                    f_acct.add_trace(go.Scatter(x=_lx, y=[_y], name=f"{_n} {live_quote_time}", mode="markers",
                                                marker=dict(color=_c, size=13, symbol="star",
                                                            line=dict(color="#fff", width=1))))
        f_acct.add_hline(y=100, line=dict(color=MUTED, dash="dot", width=1))
        style(f_acct, f"Growth of 100 since {_a['perf']['start']}", 380)
        f_acct.update_layout(yaxis_title="Index (start = 100)")
        _perf = _a.get("perf_live") or _a["perf"]
        account_html += f"""
<h3>My account vs strategy vs Nifty 500</h3>
<p class='muted'>Since {_perf['start']}{" (live " + _perf["time"] + " IST)" if _a.get("perf_live") else ""}: my account {pct(_perf['account_pct'])} · strategy record {pct(_perf['strategy_pct'])}
· Nifty 500 {pct(_perf['nifty_pct'])}. Account return is time-weighted, so deposits and withdrawals do not count as
performance. XIRR (money-weighted, annualised) is shown once {XIRR_MIN_DAYS}+ days have passed.
{"⚠ No price for: " + ", ".join(_a["unpriced"]) + " — valued at 0 in this chart." if _a["unpriced"] else ""}</p>
<div class='card'>{fig_to_div(f_acct)}</div>"""
    if _a["exec_months"]:
        _ex_cols = [("month", "Month", None), ("trades", "Trades", str), ("traded_value", "Traded ₹", inr),
                    ("slippage", "Price vs close ₹", inr), ("charges", "Charges ₹", lambda v: inr(v, 2)),
                    ("total", "Total ₹", inr), ("total_pct", "Total % of traded", pct),
                    ("cumulative", "Cumulative ₹", inr)]
        account_html += f"""
<h3>Execution cost vs the strategy close</h3>
<p class='muted'>For every trade that matches the plan: (your price − strategy close) × quantity for buys, the reverse
for sells, plus charges. Positive = cost, negative (green) = you saved compared with buying/selling at the close.</p>
{table(_a["exec_months"], _ex_cols, "tExec", None, lambda k, r: pn(-r[k]) if k in ("slippage", "total", "total_pct", "cumulative") and r[k] is not None else "")}"""

# ── Header / KPI cards ───────────────────────────────────────────────────────
live_start_label = str(live_init_date.date()) if live_started else LIVE_START_DATE
state_note = ("" if live_started else
              f"<div class='note'>Live portfolio has not started yet — no market data on/after "
              f"{LIVE_START_DATE}. Figures below show a fresh ₹{INITIAL_CAPITAL:,} portfolio and the "
              f"initialisation preview using the latest data ({AS_OF.date()}). The first session on/after "
              f"{LIVE_START_DATE} is the first entry day: signal and purchases at that day's close (order "
              f"plan from live prices during market hours).</div>")

kpis_port = "".join([
    kpi("Initial Capital", inr(INITIAL_CAPITAL)),
    kpi("Current Value", inr(live_value)),
    kpi("Total Return ₹", inr(live_pnl), pn(live_pnl)),
    kpi("Total Return %", pct(live_ret), pn(live_ret)),
    kpi("Invested Value", inr(invested)),
    kpi("Cash", inr(live_cash)),
    kpi("Cash %", pct(live_cash / live_value * 100, False)),
    kpi("Positions", f"{len(pos_rows)} / {N_HOLD}"),
])

kpis_perf = "".join([                     # return / P&L are in the top cards (not repeated here)
    kpi("Live Max Drawdown", pct(live_max_dd, False) if live_max_dd is not None else "—",
        "r" if live_max_dd else ""),
    kpi("Winning Months", str(win_months)),
    kpi("Losing Months", str(lose_months)),
    kpi("Closed Trades", str(n_lt)),
    kpi("Trade Win Rate", pct(lt_stats["wr"], False) if lt_stats["wr"] is not None else "—"),
    kpi("Avg Trade P&L", pct(lt_stats["avg_pnl"]) if lt_stats["avg_pnl"] is not None else "—",
        pn(lt_stats["avg_pnl"])),
    kpi("Avg Holding Days", num(lt_stats["avg_hold"], 0) if lt_stats["avg_hold"] is not None else "—"),
])

# ── Actions ──────────────────────────────────────────────────────────────────
def act_lines(items, fmt_fn):
    if not items:
        return "<div class='muted'>None</div>"
    return "".join(f"<div class='al'>{fmt_fn(a)}</div>" for a in items)


_px_word = {"executed": "exec", "orderplan": f"live {live_quote_time}"}.get(act_mode, "est.")
buy_html = act_lines(acts["BUY"], lambda a:
                     f"<b>{a['symbol']}</b><span>#{a['rank']} · Target {inr(a['target'])} · "
                     f"{a['qty']} × {num(a['price'])} ({_px_word}) = {inr(a['value'])}</span>")
hold_html = act_lines(acts["HOLD"], lambda a:
                      f"<b>{a['symbol']}</b><span>#{a['rank']} · {inr(a['value'])} · {pct(a['pnl_pct'])}</span>")
sell_html = act_lines(acts["SELL"], lambda a:
                      f"<b>{a['symbol']}</b><span>{a['reason']} · #{a['rank'] or '—'} · {inr(a['value'])} · "
                      f"{pct(a['pnl_pct'])}</span>")
skip_html = act_lines(acts["SKIP"], lambda a:
                      f"<b>{a['symbol']}</b><span>#{a['rank']} · {a['reason']} · price {num(a['price'])} "
                      f"vs DMA {num(a['dma200'])}</span>")
out_html = act_lines([dict(symbol=s, rank=cur_rank[s], score=cur_sc[s]["score"],
                           above=cur_sc[s]["price"] > cur_sc[s]["dma"]) for s in cur_ranked[N_HOLD:N_HOLD + 4]],
                     lambda a: f"<b>{a['symbol']}</b><span>#{a['rank']} · score {num(a['score'])} · "
                               f"{'above' if a['above'] else 'below'} DMA</span>")
_exp = lambda known: "" if known else " (expected)"
market_html = (f"<div class='marketopen'>🟢 <b>Market open — live prices {live_quote_time} IST</b>. Current values "
               f"(positions, P&amp;L, stops, signals, ranking, My Account) use today's live Yahoo prices; ETFs "
               f"without a live quote use the {AS_OF.date()} close. Live prices are never recorded as closes: "
               f"history charts use closing prices plus a ★ live point.</div>" if LIVE_VIEW else
               f"<div class='marketopen nolive'>⚪ <b>No live prices today</b> — it is market hours, but Yahoo has no "
               f"prices for {TODAY.date()} (NSE holiday, or Yahoo not updated yet). All values use the "
               f"{AS_OF.date()} close.</div>" if NO_LIVE_TODAY else "")
if act_mode == "orderplan":
    act_banner = (f"🛒 ORDER PLAN — entry day {TODAY.date()} · signal and quantities from live prices at "
                  f"{live_quote_time} IST · NOT recorded yet")
    act_caption = (f"Today is the entry day: momentum, 200-DMA, ranking and stops use today's live Yahoo price "
                   f"at {live_quote_time} IST as today's price. Quantities = target ÷ live price, whole units "
                   f"(Yahoo may lag — check your broker's price). This run's choice is saved; the trade is recorded "
                   f"at today's close by the next run after 16:00 IST (the 20:15 evening run)."
                   + (f" No live quote for: {', '.join(order_plan_missing_quotes)} — last close used."
                      if order_plan_missing_quotes else ""))
    act_prefix = "ORDER"
elif act_mode == "incomplete" and NO_LIVE_TODAY:
    act_banner = "⛔ NO LIVE PRICES TODAY — do not trade on this order plan"
    act_caption = (f"Yahoo has no prices at all for {TODAY.date()}: most likely an NSE holiday (the entry day then "
                   f"moves to the next trading session), or Yahoo is not updated yet — re-run the workflow in a few "
                   f"minutes to check.")
    act_prefix = "UNRELIABLE"
elif act_mode == "incomplete":
    act_banner = "⛔ DATA INCOMPLETE — do not trade on this order plan"
    act_caption = (f"{len(signal_missing)} ETFs have no live price on Yahoo ({', '.join(signal_missing)}), so the "
                   f"ranking below uses yesterday's close for them and may be wrong. Re-run the workflow in a few "
                   f"minutes.")
    act_prefix = "UNRELIABLE"
elif act_mode == "executed":
    act_banner = f"EXECUTED — {AS_OF:%B} entry-day rebalance recorded at the {AS_OF.date()} close"
    act_caption = (f"Entry day {AS_OF.date()}: signal and fills at that day's close. "
                   f"Decision: {signal_source}.")
    act_prefix = "EXECUTED"
else:
    _basis = f"live prices {live_quote_time} IST" if LIVE_VIEW else f"the latest close ({AS_OF.date()})"
    act_banner = (f"PREVIEW — from {_basis}; final decision on the entry day "
                  f"{next_rebal.date()}{_exp(next_confirmed)}")
    act_caption = (f"Uses {_basis}, which is not the final decision. On the entry day "
                   f"{next_rebal.date()}{_exp(next_confirmed)} run the workflow during market hours (~14:00 IST) "
                   f"for the order plan from live prices. Prices are estimates.")
    if signal_status == "ENTRY DAY":
        act_caption = (f"Today ({TODAY.date()}) is the entry day, but the market is closed and today's close is "
                       f"not in the data yet. Run the workflow during market hours for the order plan, or after "
                       f"Yahoo publishes today's close to record it.")
    act_prefix = "PREVIEW"
    if not live_started:
        act_caption = f"Initialisation preview for a fresh {inr(INITIAL_CAPITAL)}. " + act_caption

actions_html = f"""
<div class='banner {act_mode}'>{act_banner}</div>
{act_data_warn}
<p class='muted'>{act_caption}</p>
<div class='acts'>
 <div class='act a-buy'><h4>{act_prefix} · BUY ({len(acts['BUY'])})</h4>{buy_html}</div>
 <div class='act a-hold'><h4>{act_prefix} · HOLD ({len(acts['HOLD'])})</h4>{hold_html}</div>
 <div class='act a-sell'><h4>{act_prefix} · SELL ({len(acts['SELL'])})</h4>{sell_html}</div>
 <div class='act a-skip'><h4>{act_prefix} · SKIP — top 6 below 200 DMA ({len(acts['SKIP'])})</h4>{skip_html}</div>
 <div class='act a-out'><h4>OUTSIDE TOP 6 (next ranked)</h4>{out_html}</div>
</div>"""

# ── Tables ───────────────────────────────────────────────────────────────────
# Column order: the most important columns first (Symbol stays fixed when a table scrolls sideways)
pos_cols = [
    ("symbol", "Symbol", None), ("status", "Status", badge), ("price", "Current Price", num),
    ("pnl_pct", "Unrealised P&L %", pct), ("pnl", "Unrealised P&L ₹", inr),
    ("hard_sl", "Hard SL", num), ("trail_sl", "Trailing SL", num),
    ("dist_sl", "Distance to SL %", lambda v: pct(v, False)),
    ("dist_trail", "Distance to Trail %", lambda v: pct(v, False)),
    ("rank", "Rank", lambda v: "—" if v is None else str(v)), ("value", "Current Value", inr),
    ("weight", "Position Weight %", lambda v: pct(v, False)), ("qty", "Qty", str),
    ("entry_price", "Entry Price", num), ("entry_date", "Entry Date", None), ("cost", "Cost", inr),
    ("hold_days", "Hold Days", str), ("score", "Momentum Score", num),
    ("r1", "1M", pct), ("r3", "3M", pct), ("r6", "6M", pct), ("r12", "12M", pct),
    ("dma", "200 DMA", num), ("above_dma", "Above 200 DMA", yn), ("category", "Category", None),
    ("price_date", "Price Date", lambda v: pxdate(v)),
]


def pos_cell(k, r):
    if k in ("pnl", "pnl_pct"):
        return pn(r[k])
    if k == "above_dma":
        return "g" if r[k] else "r"
    if k in ("r1", "r3", "r6", "r12"):
        return pn(r[k])
    if k in ("dist_sl", "dist_trail"):
        return "r" if r[k] <= 0 else ("amber" if r[k] <= NEAR_STOP_PCT else "")
    return "sym" if k == "symbol" else ""


def pos_row(r):
    if r["status"] != "HOLD":
        return "warn"
    return "pos" if r["pnl"] > 0 else "neg"


positions_html = table(pos_rows, pos_cols, "tPos", pos_row, pos_cell)

risk_rows = sorted(pos_rows, key=lambda r: min(r["dist_sl"], r["dist_trail"]))
for r in risk_rows:
    r["nearest"] = "Hard SL" if r["dist_sl"] <= r["dist_trail"] else "Trailing SL"
    r["cushion"] = min(r["dist_sl"], r["dist_trail"])
risk_cols = [
    ("symbol", "Symbol", None), ("nearest", "Nearest Stop", None),
    ("cushion", "Cushion %", lambda v: pct(v, False)), ("status", "Status", badge),
    ("price", "Current Price", num), ("hard_sl", "Hard SL", num), ("trail_sl", "Trailing SL", num),
    ("dist_sl", "Distance to Hard SL %", lambda v: pct(v, False)),
    ("dist_trail", "Distance to Trailing SL %", lambda v: pct(v, False)),
    ("entry_price", "Entry Price", num), ("peak", "Peak Price", num),
    ("price_date", "Price Date", lambda v: pxdate(v)),
]
risk_html = table(risk_rows, risk_cols, "tRisk",
                  lambda r: "warn" if r["cushion"] <= NEAR_STOP_PCT else "",
                  lambda k, r: ("r" if r[k] <= 0 else "amber" if r[k] <= NEAR_STOP_PCT else "")
                  if k in ("dist_sl", "dist_trail", "cushion") else ("sym" if k == "symbol" else ""))

sig_cols = [
    ("symbol", "Symbol", None), ("rank", "Rank", str), ("signal", "Signal", badge),
    ("score", "Momentum Score", num), ("price", "Current Price", num), ("dma", "200 DMA", num),
    ("above_dma", "Above 200 DMA", yn), ("held", "Currently Held", yn),
    ("target", "Target Allocation", inr), ("est_qty", "Est. Qty", lambda v: "—" if v is None else str(v)),
    ("r1", "1M %", pct), ("r3", "3M %", pct), ("r6", "6M %", pct), ("r12", "12M %", pct),
    ("category", "Category", None), ("price_date", "Price Date", lambda v: pxdate(v)),
]
SIG_SHOW = 15                                 # rows shown before "Show all"
signals_html = table(
    sig_rows, sig_cols, "tSig",
    lambda r: ("top" if r["rank"] <= N_HOLD else "") + (" more" if r["rank"] > SIG_SHOW else ""),
    lambda k, r: ("g" if r[k] else "r") if k == "above_dma" else
    (pn(r[k]) if k in ("r1", "r3", "r6", "r12") else ("sym" if k == "symbol" else "")))

# ── Ranking trend table ──
_trend_rows = [dict(now=r["now"], symbol=r["symbol"], category=r["category"], change=r["change"],
                    held=r["held"], **{f"r{i}": rk for i, rk in enumerate(r["ranks"])})
               for r in rank_trend if (r["now"] or 999) <= 15 or r["held"]]
_trend_cols = ([("symbol", "Symbol", None), ("now", "Now" + (f" (live {live_quote_time})" if LIVE_VIEW else ""), str),
                ("change", "Change", lambda v: "—" if v is None else ("↑" if v > 0 else "↓" if v < 0 else "=")
                 + (str(abs(v)) if v else ""))]
               + [(f"r{i}", d.strftime("%d %b %y"), lambda v: "—" if v is None else str(v))
                  for i, d in reversed(list(enumerate(trend_days)))]           # newest entry day first
               + [("held", "Held", yn), ("category", "Category", None)])


def _trend_cell(k, r):
    v = r.get(k)
    if k == "change":
        return pn(v)
    if (k == "now" or k.startswith("r")) and isinstance(v, int) and v <= N_HOLD:
        return "t6"
    return "sym" if k == "symbol" else ""


rank_trend_html = table(_trend_rows, _trend_cols, "tTrend", lambda r: "top" if r["held"] else "", _trend_cell)

trade_cols = [
    ("symbol", "Symbol", None), ("trade_id", "Trade #", str), ("category", "Category", None),
    ("entry_date", "Entry Date", None), ("exit_date", "Exit Date", None),
    ("entry_price", "Entry Price", num), ("exit_price", "Exit Price", num), ("qty", "Qty", str),
    ("cost", "Cost ₹", inr), ("proceeds", "Proceeds ₹", inr), ("pnl_inr", "P&L ₹", inr),
    ("pnl_pct", "P&L %", pct), ("hold_days", "Hold Days", str), ("peak_price", "Peak Price", num),
    ("exit_reason", "Exit Reason", None), ("momentum_score_at_entry", "Momentum Score at Entry", num),
]
live_trades_html = table(live_trades_df.to_dict("records") if n_lt else [], trade_cols, "tLTr",
                         lambda r: "pos" if r["pnl_inr"] > 0 else "neg",
                         lambda k, r: pn(r[k]) if k in ("pnl_inr", "pnl_pct") else "")
trade_stats_html = "".join([
    kpi("Total Closed Trades", str(lt_stats["total"])),
    kpi("Winning Trades", str(lt_stats["wins"]), "g" if lt_stats["wins"] else ""),
    kpi("Losing Trades", str(lt_stats["losses"]), "r" if lt_stats["losses"] else ""),
    kpi("Win Rate", pct(lt_stats["wr"], False) if lt_stats["wr"] is not None else "—"),
    kpi("Total Realised P&L", inr(lt_stats["realised"]), pn(lt_stats["realised"])),
    kpi("Average P&L %", pct(lt_stats["avg_pnl"]), pn(lt_stats["avg_pnl"])),
    kpi("Average Hold Days", num(lt_stats["avg_hold"], 0) if lt_stats["avg_hold"] is not None else "—"),
    kpi("Best Trade", pct(lt_stats["best"]), pn(lt_stats["best"])),
    kpi("Worst Trade", pct(lt_stats["worst"]), pn(lt_stats["worst"])),
    kpi("Exits: Hard SL / Trail / Rotation", f"{lt_stats['sl']} / {lt_stats['trail']} / {lt_stats['rot']}"),
])

rebal_rows = []
if live_started:
    for m in live["monthly"]:
        rebal_rows.append(dict(sdate=m["signal_date"].date(), date=m["rebal_date"].date(),
                               value=m["portfolio_value"],
                               invested=m["invested"], cash=m["cash"], n_held=m["n_held"],
                               ret=m["monthly_return_pct"],
                               entries=m["entries"] or "—", exits=m["exits"] or "—",
                               skipped=m["skipped"] or "—", holdings=m["holdings"] or "—"))
rebal_cols = [
    ("date", "Entry Day", None),
    ("value", "Portfolio Value", inr), ("invested", "Invested", inr),
    ("cash", "Cash", inr), ("n_held", "Positions", str), ("ret", "Return since prev. rebalance", pct),
    ("entries", "Bought", None), ("exits", "Sold", None), ("skipped", "Skipped (below DMA)", None),
    ("holdings", "Holdings after rebalance", None),
]
rebal_html = table(rebal_rows, rebal_cols, "tReb", None, lambda k, r: pn(r[k]) if k == "ret" else "")

# ── Charts (live only) ───────────────────────────────────────────────────────
charts = {}
if pos_rows or live_cash > 0:
    labels = [r["symbol"] for r in pos_rows] + ["Cash"]
    values = [r["value"] for r in pos_rows] + [live_cash]
    f_alloc = go.Figure(go.Pie(labels=labels, values=values, hole=0.55, sort=False,
                               texttemplate="%{label}<br>%{percent:.1%}",
                               hovertemplate="%{label}<br>₹%{value:,.0f}<br>%{percent:.2%}<extra></extra>"))
    f_alloc.update_traces(marker=dict(line=dict(color=PANEL, width=2)))
    style(f_alloc, "Current Portfolio Allocation", 420)
    f_alloc.update_layout(showlegend=False)
    charts["alloc"] = fig_to_div(f_alloc)
if pos_rows:
    f_pnl = go.Figure(go.Bar(                     # horizontal: symbol names stay readable on phones
        y=[f"{r['symbol']}  {r['pnl_pct']:+.1f}%" for r in pos_rows], x=[r["pnl"] for r in pos_rows],
        orientation="h", marker_color=[GRN if r["pnl"] >= 0 else RD for r in pos_rows],
        customdata=[[r["symbol"], r["pnl_pct"], r["value"]] for r in pos_rows],
        hovertemplate="%{customdata[0]}<br>P&L ₹%{x:,.0f}<br>P&L %{customdata[1]:+.2f}%<br>"
                      "Value ₹%{customdata[2]:,.0f}<extra></extra>"))
    style(f_pnl, "Position P&L (unrealised ₹)", 420)
    f_pnl.update_layout(xaxis_title="Unrealised P&L ₹", showlegend=False,
                        yaxis=dict(autorange="reversed", automargin=True))
    charts["pnl"] = fig_to_div(f_pnl)
if live_started:
    f_val = go.Figure()
    # Cash is shown in the hover (not as a line): a ~₹500 cash line forced the axis down to 0 and
    # flattened the portfolio line
    f_val.add_trace(go.Scatter(x=live_daily.index, y=live_daily["value"], name="Portfolio Value",
                               line=dict(color=AMBER, width=2.5), customdata=live_daily["cash"],
                               hovertemplate="%{x|%d %b %Y}<br>Value ₹%{y:,.0f}<br>Cash ₹%{customdata:,.0f}"
                                             "<extra></extra>"))
    f_val.add_trace(go.Scatter(x=live_daily.index, y=live_daily["held"], name="Invested Value",
                               line=dict(color=BLUE, width=1.5)))
    if LIVE_VIEW:                                  # separate, clearly marked live point (not history)
        f_val.add_trace(go.Scatter(x=[pd.Timestamp(now_ist())], y=[live_value],
                                   name=f"Live {live_quote_time} IST", mode="markers",
                                   marker=dict(color=AMBER, size=14, symbol="star", line=dict(color="#fff", width=1))))
    f_val.add_hline(y=INITIAL_CAPITAL, line=dict(color=MUTED, dash="dash", width=1),
                    annotation_text=f"Initial {inr(INITIAL_CAPITAL)}", annotation_font_color=MUTED)
    style(f_val, "Live Portfolio Value", 430)
    f_val.update_layout(yaxis_title="Portfolio Value ₹", yaxis_tickformat=",.0f")
    charts["val"] = fig_to_div(f_val)

    f_dd = go.Figure(go.Scatter(x=live_daily.index, y=live_daily["dd_pct"], fill="tozeroy",
                                name="Drawdown %", line=dict(color=RD, width=1),
                                fillcolor="rgba(231,76,60,0.35)"))
    f_dd.add_hline(y=live_max_dd, line=dict(color=AMBER, dash="dash"),
                   annotation_text=f"Max live DD {live_max_dd:.2f}%", annotation_font_color=AMBER)
    style(f_dd, "Live Portfolio Drawdown", 330)
    f_dd.update_layout(yaxis_title="Drawdown %", showlegend=False)
    charts["dd"] = fig_to_div(f_dd)

    mlabels = [ts.strftime("%b %Y") + (" (MTD)" if ts.to_period("M") == AS_OF.to_period("M")
                                       and AS_OF < ts else "") for ts in live_mret.index]
    f_m = go.Figure(go.Bar(x=mlabels, y=live_mret.values,
                           marker_color=[GRN if v >= 0 else RD for v in live_mret.values],
                           text=[f"{v:+.2f}%" for v in live_mret.values], textposition="outside"))
    style(f_m, "Live Monthly Returns %", 330)
    f_m.update_layout(yaxis_title="Return %", showlegend=False)
    charts["mret"] = fig_to_div(f_m)


def chart(key):
    return f"<div class='card'>{charts[key]}</div>" if key in charts else ""


charts_html = (f"<div class='grid2'>{chart('alloc')}{chart('pnl')}</div>{chart('val')}"
               f"<div class='grid2'>{chart('dd')}{chart('mret')}</div>")
if not charts_html.replace("<div class='grid2'></div>", ""):
    charts_html = "<p class='muted'>No live history yet.</p>"

# ── Next rebalance / rules ───────────────────────────────────────────────────
_status_txt = {
    "EXECUTED": f"{AS_OF:%B} entry-day rebalance recorded at the {AS_OF.date()} close.",
    "ORDER PLAN": f"Entry day — ORDER PLAN from live prices at {live_quote_time} IST.",
    "ENTRY DAY": "Today is the entry day — run the workflow during market hours for the order plan.",
    "PENDING": f"Entry day not reached — latest data is {AS_OF.date()}. Current actions are a PREVIEW only.",
    "INCOMPLETE": f"Live prices missing for {len(signal_missing)} ETFs — order plan not reliable yet.",
}[signal_status]
next_html = f"""<div class='kpis'>
{kpi("Latest Market Data", str(AS_OF.date()))}
{kpi("Last Entry Day", str(cur_rebal.date()) if cur_rebal is not None else "Not started")}
{kpi("Next Entry Day" + _exp(next_confirmed), next_rebal.strftime("%a %d %b %Y"))}
{kpi("Status", signal_status, {"EXECUTED": "g", "ORDER PLAN": "amber", "ENTRY DAY": "amber", "PENDING": "", "INCOMPLETE": "r"}[signal_status])}
{kpi("Days to Entry Day", str(max(days_to_next, 0)))}
</div>
<p class='muted'>{_status_txt} The entry day (first trading day of the month) is both the signal day and the
execution day. {"" if next_confirmed else
 "Future sessions are not in the price data yet, so expected dates skip weekends only; an NSE holiday "
 "moves them to the adjacent trading session."}</p>"""

rules = [
    ("Live Start Date", live_start_label), ("Initial Capital", inr(INITIAL_CAPITAL)),
    ("N_HOLD", str(N_HOLD)),
    ("Momentum Lookbacks", " / ".join(str(x) for x in LB) + " trading days"),
    ("Weights", " / ".join(f"{w * 100:.0f}%" for w in WEIGHTS) + "  (1M / 3M / 6M / 12M)"),
    ("200 DMA", f"{DMA_PERIOD} days — entry gate only; no substitution with rank #{N_HOLD + 1}"),
    ("Hard SL", f"{SL_PCT * 100:.0f}% below entry"),
    ("Trailing SL", f"{TRAIL_PCT * 100:.0f}% below peak (peak ratchets on rebalance dates)"),
    ("Idle Cash Rate", f"{LIQUID_RATE * 100:.0f}% p.a."),
    ("Rebalance timing", "Entry day = first trading day of the month: signal (momentum, 200-DMA, ranking, "
                         "stops) and execution both at that day's close. Live: the order plan uses the entry "
                         "day's live price; the chosen ETFs are recorded at the close"),
    ("Position sizing", "New entries = total value at execution / N_HOLD, whole units; no resizing"),
    ("Universe", f"C54 — {len(SYMBOLS)} with enough history"
                 + (f"; dropped: {', '.join(dropped)}" if dropped else "")),
]
rules_html = "".join(f"<tr><td>{a}</td><td>{b}</td></tr>" for a, b in rules)

# ── Historical reference (collapsed, secondary) ─────────────────────────────
f_hist = go.Figure()
f_hist.add_trace(go.Scatter(x=daily_df.index, y=daily_df["value"] / INITIAL_CAPITAL * 100,
                            name="Strategy (backtest)", line=dict(color=AMBER, width=1.5)))
if bm_daily is not None:
    f_hist.add_trace(go.Scatter(x=bm_daily.index, y=bm_daily / bm_daily.iloc[0] * 100,
                                name="Nifty 500 B&H", line=dict(color=GRAY, dash="dash")))
style(f_hist, "Historical backtest (normalised to 100) — reference only", 380)
hist_div = fig_to_div(f_hist)
hist_kpis = "".join(kpi(k, (f"{v:,.2f}" if isinstance(v, float) else str(v))) for k, v in metrics
                    if k in ("CAGR%", "MaxDD% (daily)", "Sharpe", "Sortino", "Calmar", "Volatility%",
                             "BM_CAGR% (Nifty 500)", "Alpha%", "N_Trades", "Trade_WR%", "Avg_PnL%"))
hist_annual = " · ".join(f"{y}: {r:+.1f}%" for y, r in annual.items())
hist_html = f"""
<details class='hist'>
 <summary>Historical Backtest Reference ({first_d.date()} → {LAST_DATE.date()}) — not your live portfolio</summary>
 <p class='muted'>Simulated strategy history used only as a reference. Its positions, prices and P&L are
 <b>not</b> part of the live portfolio above.</p>
 <div class='kpis small'>{hist_kpis}</div>
 <p class='muted'>Annual returns: {hist_annual}</p>
 <div class='card'>{hist_div}</div>
</details>"""

_status_cls = {"EXECUTED": "g", "ORDER PLAN": "amber", "ENTRY DAY": "amber", "PENDING": "", "INCOMPLETE": "r"}[signal_status]
_nav = [("actions", "Actions"), ("positions", "Positions"), ("account", "My Account"), ("risk", "Risk"),
        ("signals", "Signals"), ("trend", "Ranking"), ("perf", "Performance"), ("next", "Next Entry"),
        ("charts", "Charts"), ("history", "History"), ("trades", "Closed Trades"), ("rules", "Rules")]
nav_html = "".join(f"<a href='#{k}'>{t}</a>" for k, t in _nav)
_plotly_tags = (f'<script src="{PLOTLY_CDNS[0]}" integrity="{PLOTLY_SRI}" crossorigin="anonymous"></script>\n'
                f'<script>window.Plotly || document.write(\'<script src="{PLOTLY_CDNS[1]}" integrity="{PLOTLY_SRI}" '
                f'crossorigin="anonymous"><\\/script>\')</script>')


def _sec(sid, title, body, note=""):
    """Collapsible section (closed by default) — the nav bar opens it when clicked."""
    return (f"<details class='sec' id='{sid}'><summary><h2>{title}</h2>"
            f"<span class='muted'>{note}</span></summary>{body}</details>")


page_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ETF Momentum Live Dashboard</title>
<style>
 * {{ box-sizing:border-box; }}
 html {{ scroll-behavior:smooth; -webkit-text-size-adjust:100%; }}
 body {{ background:{BG}; color:{TXT}; font-family:Segoe UI,Roboto,Arial,sans-serif; margin:0; padding:0 24px 24px; }}
 .wrap {{ max-width:1500px; margin:0 auto; }}
 h1 {{ margin:18px 0 8px; font-size:24px; line-height:1.25; }}
 h2 {{ font-size:18px; letter-spacing:.3px; border-bottom:1px solid #2E4053; padding-bottom:6px; margin:30px 0 12px; scroll-margin-top:60px; }}
 h4 {{ margin:0 0 8px; font-size:13px; letter-spacing:.5px; }}
 /* sticky section bar */
 nav.toc {{ position:sticky; top:0; z-index:20; background:{BG}; display:flex; gap:6px; overflow-x:auto;
   padding:10px 0; margin:0 0 4px; border-bottom:1px solid #22364A; scrollbar-width:none; }}
 nav.toc::-webkit-scrollbar {{ display:none; }}
 nav.toc a {{ flex:0 0 auto; color:{TXT}; text-decoration:none; font-size:13px; padding:6px 12px; border-radius:16px;
   background:{PANEL}; border:1px solid #2E4053; }}
 nav.toc a:hover {{ border-color:{AMBER}; color:{AMBER}; }}
 /* compact status strip */
 .strip {{ display:flex; flex-wrap:wrap; gap:8px; margin:4px 0 6px; }}
 .chip {{ background:{PANEL}; border-radius:8px; padding:6px 10px; font-size:13px; color:{MUTED}; }}
 .chip b {{ color:{TXT}; font-weight:600; }} .chip b.g {{ color:{GRN}; }} .chip b.amber {{ color:{AMBER}; }} .chip b.r {{ color:{RD}; }}
 .sub {{ color:{MUTED}; font-size:12px; margin:2px 0 0; }}
 .muted {{ color:{MUTED}; font-size:13px; line-height:1.45; }}
 .banner {{ display:inline-block; padding:6px 12px; border-radius:6px; font-weight:700; font-size:13px; margin-bottom:6px; }}
 .banner.preview {{ background:#34495E; color:#fff; }} .banner.signal {{ background:{AMBER}; color:#2B1D02; }}
 .banner.executed {{ background:{GRN}; color:#0B2716; }}
 .banner.orderplan {{ background:{BLUE}; color:#0B1F2E; }} .banner.incomplete {{ background:{RD}; color:#fff; }}
 .marketopen {{ background:rgba(93,173,226,0.12); border:1px solid {BLUE}; border-radius:8px; padding:8px 14px; margin:12px 0; font-size:14px; }}
 .marketopen.nolive {{ background:rgba(158,158,158,0.12); border-color:{GRAY}; }}
 .datawarn {{ background:rgba(231,76,60,0.12); border:1px solid {RD}; border-left:5px solid {RD};
   border-radius:8px; padding:10px 16px; margin:14px 0; font-size:14px; }}
 .datawarn p {{ margin:6px 0; }} .datawarn ul {{ margin:4px 0 6px; padding-left:22px; }}
 .datawarn.slim {{ padding:6px 12px; margin:6px 0; font-size:13px; }}
 .dw-title {{ color:{RD}; font-weight:700; letter-spacing:.5px; }}
 .dataok {{ color:{GRN}; font-size:13px; margin:8px 0 0; }}
 td.stale {{ color:{AMBER}; font-weight:600; }}
 td.t6 {{ color:{BLUE}; font-weight:700; }}
 .note {{ background:#3B2F10; border:1px solid {AMBER}; border-radius:8px; padding:10px 14px; margin:14px 0; }}
 .kpis {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:12px; margin:14px 0; }}
 .kpi {{ background:{PANEL}; border-radius:10px; padding:12px 16px; min-width:0; }}
 .kl {{ color:{MUTED}; font-size:12px; text-transform:uppercase; letter-spacing:.4px; }}
 .kv {{ font-size:24px; font-weight:600; margin-top:4px; font-variant-numeric:tabular-nums; overflow-wrap:anywhere; }}
 .small .kv {{ font-size:18px; }}
 .k4 {{ grid-template-columns:repeat(4,1fr); }}
 @media (max-width:800px) {{ .k4 {{ grid-template-columns:repeat(2,1fr); }} }}
 .g {{ color:{GRN}; }} .r {{ color:{RD}; }} .amber {{ color:{AMBER}; font-weight:600; }}
 .sym {{ font-weight:600; }}
 .acts {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(260px,1fr)); gap:12px; }}
 .act {{ background:{PANEL}; border-radius:10px; padding:12px 14px; border-top:3px solid {GRAY}; }}
 .a-buy {{ border-top-color:{GRN}; }} .a-hold {{ border-top-color:{BLUE}; }} .a-sell {{ border-top-color:{RD}; }}
 .a-skip {{ border-top-color:{AMBER}; }} .a-out {{ border-top-color:#566573; }}
 .al {{ display:flex; flex-direction:column; padding:6px 0; border-bottom:1px solid #22364A; font-size:14px; }}
 .al:last-child {{ border-bottom:none; }} .al span {{ color:{MUTED}; font-size:12px; }}
 .card {{ background:{PANEL}; border-radius:10px; padding:6px; margin-bottom:12px; min-width:0; }}
 /* Fixed 2 columns: with auto-fit, a chart drawn before its row-mate exists is sized full-width and then clipped */
 .grid2 {{ display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; }}
 @media (max-width:900px) {{ .grid2 {{ grid-template-columns:1fr; }} }}
 .tw {{ overflow:auto; max-height:640px; border-radius:10px; background:{PANEL}; -webkit-overflow-scrolling:touch; }}
 table {{ border-collapse:separate; border-spacing:0; width:100%; font-size:13px; font-variant-numeric:tabular-nums; }}
 th {{ background:#22364A; position:sticky; top:0; cursor:pointer; padding:8px; text-align:left; white-space:nowrap; z-index:2; }}
 th:hover {{ color:{AMBER}; }}
 td {{ padding:6px 8px; border-bottom:1px solid #22364A; white-space:nowrap; }}
 /* first column (Symbol) stays visible while a wide table scrolls sideways */
 .tw td:first-child {{ position:sticky; left:0; z-index:1; background-color:{PANEL}; box-shadow:1px 0 0 #2E4053; }}
 .tw th:first-child {{ position:sticky; left:0; z-index:3; box-shadow:1px 0 0 #2E4053; }}
 /* row tints painted over an opaque panel colour, so the sticky first cell keeps the tint */
 tr.pos td {{ background-image:linear-gradient(rgba(46,204,113,0.10),rgba(46,204,113,0.10)); background-color:{PANEL}; }}
 tr.neg td {{ background-image:linear-gradient(rgba(231,76,60,0.12),rgba(231,76,60,0.12)); background-color:{PANEL}; }}
 tr.warn td {{ background-image:linear-gradient(rgba(245,166,35,0.14),rgba(245,166,35,0.14)); background-color:{PANEL}; }}
 tr.top td {{ background-image:linear-gradient(rgba(93,173,226,0.08),rgba(93,173,226,0.08)); background-color:{PANEL}; }}
 .collapsed tr.more {{ display:none; }}
 .morebtn {{ margin:8px 0 0; background:{PANEL}; color:{TXT}; border:1px solid #2E4053; border-radius:8px;
   padding:8px 14px; font-size:13px; cursor:pointer; }}
 .morebtn:hover {{ border-color:{AMBER}; color:{AMBER}; }}
 .badge {{ padding:2px 8px; border-radius:10px; font-size:11px; font-weight:700; letter-spacing:.3px; }}
 .b-buy {{ background:{GRN}; color:#0B2716; }} .b-hold {{ background:{BLUE}; color:#0B1F2E; }}
 .b-sell {{ background:{RD}; color:#fff; }} .b-risk {{ background:{AMBER}; color:#2B1D02; }}
 .b-skip {{ background:#7D6608; color:#fff; }} .b-out {{ background:#34495E; color:#D5DBDB; }}
 table.rules td {{ white-space:normal; overflow-wrap:anywhere; vertical-align:top; }}   /* long rule text wraps on phones */
 table.rules td:first-child {{ color:{MUTED}; width:220px; }}
 /* collapsible sections */
 details.sec > summary {{ list-style:none; cursor:pointer; scroll-margin-top:60px; }}
 details.sec > summary::-webkit-details-marker {{ display:none; }}
 details.sec > summary h2 {{ display:flex; justify-content:space-between; align-items:center; margin-bottom:4px; }}
 details.sec > summary h2::after {{ content:"▸ show"; font-size:12px; font-weight:400; color:{MUTED}; }}
 details.sec[open] > summary h2::after {{ content:"▾ hide"; }}
 details.sec > summary .muted {{ display:block; margin-bottom:10px; }}
 details.hist {{ margin-top:40px; background:#122130; border-radius:10px; padding:10px 16px; opacity:.9; }}
 details.hist summary {{ cursor:pointer; color:{MUTED}; font-weight:600; }}
 .nojs {{ display:none; }}
 @media (max-width:600px) {{
   body {{ padding:0 12px 16px; }} h1 {{ font-size:19px; }} h2 {{ font-size:16px; margin-top:24px; }}
   .kpis {{ gap:8px; grid-template-columns:repeat(2,minmax(0,1fr)); }} .kpi {{ padding:10px 12px; }}
   .kv {{ font-size:19px; }} .small .kv {{ font-size:16px; }} .kl {{ font-size:11px; }}
   .acts {{ grid-template-columns:1fr; gap:8px; }}
   table {{ font-size:12px; }} td, th {{ padding:6px; }}
   table.rules td:first-child {{ width:40%; }}
 }}
</style>
{_plotly_tags}
<script>
function sortTable(id, col) {{
  var t = document.getElementById(id), tb = t.tBodies[0], rows = Array.from(tb.rows);
  var asc = !(t.dataset.sc == col && t.dataset.sd === 'asc');
  rows.sort(function(a, b) {{
    var x = a.cells[col].dataset.v, y = b.cells[col].dataset.v;
    var r = (x !== '' && y !== '' && isFinite(x) && isFinite(y)) ? parseFloat(x) - parseFloat(y)
            : String(x).localeCompare(String(y));
    return asc ? r : -r;
  }});
  rows.forEach(function(r) {{ tb.appendChild(r); }});
  t.dataset.sc = col; t.dataset.sd = asc ? 'asc' : 'desc';
}}
function resizeCharts() {{
  if (!window.Plotly) return;
  document.querySelectorAll('.js-plotly-plot').forEach(function(p) {{
    if (p.offsetWidth) Plotly.Plots.resize(p);
  }});
}}
// Phones / touch screens: charts must not capture the finger (page scrolls instead of zooming the
// chart), and tall charts are made a little shorter. Tap still shows the values.
function fitChartsForDevice() {{
  if (!window.Plotly) return;
  var small = window.matchMedia('(max-width:700px)').matches;
  var touch = window.matchMedia('(pointer:coarse)').matches;
  document.querySelectorAll('.js-plotly-plot').forEach(function(p) {{
    var u = {{}}, isPie = p.data && p.data[0] && p.data[0].type === 'pie';
    if (small || touch) {{ u.dragmode = false; if (!isPie) {{ u['xaxis.fixedrange'] = true; u['yaxis.fixedrange'] = true; }} }}
    if (small && p.layout && p.layout.height > 320) {{
      u.height = Math.max(300, Math.round(p.layout.height * 0.8));
      p.style.height = u.height + 'px';            // the chart's own box keeps its old height otherwise
    }}
    if (Object.keys(u).length) Plotly.relayout(p, u);
  }});
}}
function toggleMore(btn) {{
  var w = btn.previousElementSibling, c = w.classList.toggle('collapsed');
  btn.textContent = c ? btn.dataset.more : btn.dataset.less;
}}
window.addEventListener('load', function() {{
  fitChartsForDevice(); resizeCharts();
  if (!window.Plotly) document.querySelectorAll('.nojs').forEach(function(e) {{ e.style.display = 'block'; }});
  // nav links open a collapsed section before jumping to it
  document.querySelectorAll('nav.toc a').forEach(function(a) {{
    a.addEventListener('click', function() {{
      var t = document.getElementById(a.getAttribute('href').slice(1));
      if (t && t.tagName === 'DETAILS') t.open = true;
    }});
  }});
}});
// Charts are drawn while the page is still loading; re-fit each one to its final box
document.addEventListener('toggle', resizeCharts, true);
</script></head><body><div class="wrap">

<h1>ETF Momentum Rotation — Live Portfolio Dashboard</h1>
<div class="strip">
 <span class="chip">Status <b class="{_status_cls}">{signal_status}</b></span>
 <span class="chip">Next entry <b>{next_rebal.strftime("%a %d %b")}{_exp(next_confirmed)}</b></span>
 <span class="chip">Data <b class="{'g' if data_ok else 'amber'}">{'✓' if data_ok else '⚠'} {AS_OF.strftime("%d %b %Y")}</b></span>
 <span class="chip">Updated <b>{now_ist():%d %b %H:%M} IST</b></span>
</div>
<div class="sub">Live since {live_start_label} · Initial capital {inr(INITIAL_CAPITAL)} · Universe C54 ({len(SYMBOLS)} ETFs with data)</div>
<nav class="toc">{nav_html}</nav>
{market_html}
{data_html}
{state_note}
<div class="kpis k4">{kpis_port}</div>

<h2 id="actions">Live Portfolio Actions</h2>
{actions_html}

<h2 id="positions">Current Live Positions</h2>
<p class='muted'>Positions opened on or after {live_start_label} only. Rows: green = profit, red = loss,
amber = exit risk (stop hit, outside top 6, or within {NEAR_STOP_PCT:.0f}% of a stop). Click a header to sort;
swipe sideways for more columns.</p>
{positions_html}

<h2 id="account">My Account — real trades (Google Sheet)</h2>
{account_html}

<h2 id="risk">Risk Monitor</h2>
<p class='muted'>Sorted by the position closest to an exit. Peak = highest entry-day close since entry (daily
closes and today's price are not used) — the trailing level the strategy checks on the next entry day.</p>
{risk_html}

<h2 id="signals">Current Strategy Signals</h2>
<p class='muted'><b>{"PREVIEW — " if signal_status == "PENDING" else ""}</b>Full ranked universe from
{f"live prices {live_quote_time} IST (the {AS_OF.date()} close for ETFs without a live quote)" if LIVE_VIEW else f"the {AS_OF.date()} close"}{" (not the final signal)" if signal_status == "PENDING" else ""}.
Top {N_HOLD} are highlighted. Target allocation = current portfolio value / {N_HOLD} = {inr(target_now)}.</p>
<div class="collapsed">{signals_html}</div>
<button class="morebtn" onclick="toggleMore(this)" data-more="Show all {len(sig_rows)} ETFs" data-less="Show top {SIG_SHOW} only">Show all {len(sig_rows)} ETFs</button>

{_sec("trend", "Ranking Trend", rank_trend_html,
      f"Momentum rank now and on the last {len(trend_days)} entry days, newest first (top {N_HOLD} in blue). Change = "
      f"rank movement since the last entry day (↑ = moved up). Current top 15 plus anything you hold.")}

<h2 id="perf">Live Performance</h2>
<div class="kpis">{kpis_perf}</div>

<h2 id="next">Next Rebalance</h2>
{next_html}

<h2 id="charts">Live Charts</h2>
<p class="muted nojs">⚠ Charts could not be loaded (no internet connection to the chart library). All tables and
figures above are complete.</p>
{charts_html}

<h2 id="history">Live Rebalance History</h2>
{rebal_html}

{_sec("trades", f"Live Closed Trades ({n_lt})", live_trades_html + f'<div class="kpis small">{trade_stats_html}</div>')}

{_sec("rules", "Strategy Rules", f'<table class="rules">{rules_html}</table>')}

{hist_html}
</div></body></html>"""

try:
    HTML_PATH.write_text(page_html, encoding="utf-8")
except Exception as e:
    print(f"[HTML ERROR] Could not write {HTML_PATH}: {e}")
    sys.exit(4)
print(f"Saved: {HTML_PATH.name}")
print(f"  → {HTML_PATH}")

# Public variant (GitHub Pages without PAGE_PASSWORD): the same page and summary WITHOUT the
# My Account section — real account data is never published unencrypted.
PUBLIC_HTML_PATH = OUTPUT_DIR / "ETF_Momentum_Test_1_Report_public.html"
_acct_block = f"<h2 id=\"account\">My Account — real trades (Google Sheet)</h2>\n{account_html}"
if page_html.count(_acct_block) != 1:
    print("[HTML ERROR] could not isolate the My Account section for the public page")
    sys.exit(4)
_public_html = page_html.replace(_acct_block, "<h2 id=\"account\">My Account</h2>\n<p class='muted'>🔒 Hidden on the public "
                                 "page. Add the PAGE_PASSWORD secret to publish the full, password-protected "
                                 "dashboard.</p>")
try:
    PUBLIC_HTML_PATH.write_text(_public_html, encoding="utf-8")
except Exception as e:
    print(f"[HTML ERROR] Could not write {PUBLIC_HTML_PATH}: {e}")
    sys.exit(4)


# ═════════════════════════════════════════════════════════════════════════════
# SUMMARY JSON — machine-readable snapshot for notifications (reporting only)
# ═════════════════════════════════════════════════════════════════════════════
SUMMARY_PATH = OUTPUT_DIR / "summary.json"


def _j(v):
    """JSON-safe value: numpy → python, NaN → None, dates → ISO strings."""
    if isinstance(v, dict):
        return {str(k): _j(x) for k, x in v.items()}
    if isinstance(v, (list, tuple)):
        return [_j(x) for x in v]
    if isinstance(v, (np.bool_, bool)):
        return bool(v)
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (float, np.floating)):
        return None if math.isnan(v) else round(float(v), 4)
    if hasattr(v, "isoformat"):
        return v.isoformat()[:10]
    return v


summary = {
    "generated_ist": now_ist().strftime("%Y-%m-%d %H:%M"),
    "today_ist": str(TODAY.date()),
    "as_of": str(AS_OF.date()),
    "live_start": live_start_label,
    "live_started": live_started,
    "initial_capital": INITIAL_CAPITAL,
    "value": live_value, "pnl": live_pnl, "ret_pct": live_ret,
    "invested": invested, "cash": live_cash,
    "n_positions": len(pos_rows), "n_hold": N_HOLD,
    "max_dd_pct": live_max_dd, "closed_trades": n_lt,
    "signal_status": signal_status,               # PENDING / ENTRY DAY / ORDER PLAN / INCOMPLETE / EXECUTED
    "act_mode": act_mode,                         # preview / orderplan / incomplete / executed
    "market_hours": MARKET_HOURS,
    "live_quote_time": live_quote_time,
    "live_view": LIVE_VIEW,
    "no_live_today": NO_LIVE_TODAY,               # market hours but Yahoo has no prices for today
    "order_plan_missing_quotes": order_plan_missing_quotes,
    "signal_source": signal_source,               # live / decision source of a recorded entry / latest close
    "signal_missing": signal_missing,             # ETFs without a live price (order plan)
    "saved_entries": sorted(_entries_new),
    "kept_decision": kept_decision_info,          # midday run kept today's saved entry-day choice
    "entry_notes": entry_notes,
    "last_signal_date": cur_signal.date() if cur_signal is not None else None,
    "last_exec_date": cur_rebal.date() if cur_rebal is not None else None,
    "next_signal_date": next_signal.date(), "next_signal_known": next_signal_known,
    "next_exec_date": next_rebal.date(), "next_exec_confirmed": next_confirmed,
    "near_stop_pct": NEAR_STOP_PCT,
    "target_allocation": target_now,
    "top6": cur_ranked[:N_HOLD],
    "positions": [{k: r[k] for k in ("rank", "symbol", "qty", "entry_date", "entry_price", "price", "value",
                                     "pnl", "pnl_pct", "hard_sl", "trail_sl", "dist_sl", "dist_trail",
                                     "status", "price_date")} for r in pos_rows],
    "actions": {k: [{f: a.get(f) for f in ("symbol", "rank", "qty", "price", "signal_price", "value",
                                           "target", "reason", "pnl_pct")} for a in v]
                for k, v in acts.items()},
    "account": None if account is None else {
        "value": account["value"], "pnl": account["pnl"], "ret_pct": account["ret_pct"],
        "cash": account["cash"], "charges": account["charges"], "realised": account["realised"],
        "upnl": account["upnl"], "capital": account["capital"], "capital_assumed": account["capital_assumed"],
        "trades": len(account_trades), "issues": account_issues,
        "latest_month": account["latest_month"], "latest_entered": account["latest_entered"],
        "latest_mismatches": account["latest_mismatches"],
        "latest_pva": [r for r in account["pva"] if r["month"] == account["latest_month"]],
        "xirr": account["xirr"], "xirr_days": account["xirr_days"], "xirr_min_days": XIRR_MIN_DAYS,
        "perf": account["perf"], "perf_live": account["perf_live"], "exec_total": account["exec_total"],
        "exec_latest": account["exec_months"][-1] if account["exec_months"] else None,
        "holdings": [{k: h[k] for k in ("symbol", "qty", "avg_cost", "value", "upnl", "upnl_pct",
                                         "strategy_qty", "status", "price", "avg_fill", "my_peak",
                                         "my_hard_sl", "my_trail_sl", "my_dist_sl", "my_dist_trail",
                                         "my_stop_status", "my_price_stale")} for h in account["holdings"]],
    },
    "account_error": account_error,
    "data": {
        "status": data_status,
        "etfs_loaded": len(SYMBOLS), "universe": len(UNIVERSE),
        "stale": {str(d.date()): syms for d, syms in stale_groups.items()},
        "stale_held": stale_held, "stale_signal": stale_signal, "stale_actions": stale_actions,
        "market_delayed": market_delayed, "expected_latest": expected_latest.date(),
        "excluded": dropped_info,
    },
}
try:
    SUMMARY_PATH.write_text(json.dumps(_j(summary), indent=2, ensure_ascii=False), encoding="utf-8")
    (OUTPUT_DIR / "summary_public.json").write_text(            # public variant: no account data
        json.dumps(_j({k: v for k, v in summary.items() if k not in ("account", "account_error")}),
                   indent=2, ensure_ascii=False), encoding="utf-8")
except Exception as e:
    print(f"[SUMMARY ERROR] Could not write {SUMMARY_PATH}: {e}")
    sys.exit(5)
print(f"Saved: {SUMMARY_PATH.name}")

open_report(HTML_PATH)
