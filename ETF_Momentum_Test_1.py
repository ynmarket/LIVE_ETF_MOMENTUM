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
               ("openpyxl", "openpyxl"), ("plotly", "plotly")]:
    _ensure(_p, _i)

import copy
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
    ("MONQ50", "Factor"), ("HDFCGROWTH", "Factor"), ("MOM50", "Factor"), ("MOMENTUM50", "Factor"),
    ("MOVALUE", "Factor"),
    ("MON100", "International"), ("MAFANG", "International"), ("HNGSNGBEES", "International"),
    ("GOLDBEES", "Gold_Silver"), ("SILVERBEES", "Gold_Silver"),
    ("LTGILTBEES", "Bonds"), ("EBBETF0430", "Bonds"), ("GILT5YBEES", "Bonds"),
]
CATEGORY = dict(UNIVERSE)


# ═════════════════════════════════════════════════════════════════════════════
# DATA DOWNLOAD
# ═════════════════════════════════════════════════════════════════════════════
def fetch_close(yf_symbol, start, end):
    """Download daily Close for one symbol; returns a clean Series (may be empty)."""
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
for nse, _cat in UNIVERSE:
    s = fetch_close(f"{nse}.NS", dl_start, dl_end)
    if MARKET_HOURS and not s.empty and s.index[-1] >= TODAY_TS:
        live_quotes[nse] = float(s.iloc[-1])
        s = s[s.index < TODAY_TS]
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
if MARKET_HOURS and not bm_raw.empty:
    bm_raw = bm_raw[bm_raw.index < TODAY_TS]
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


def prev_trading_day(d):
    """Immediately preceding date in the actual market calendar (never calendar-day arithmetic)."""
    i = cal.get_loc(d)
    if i == 0:
        raise ValueError(f"No trading day before {d.date()} in the price data")
    return cal[i - 1]


def rebalance_schedule(execution_dates):
    """[(signal_date, execution_date)]: signal = previous trading day's close,
    execution = first trading day of the month."""
    return [(prev_trading_day(e), e) for e in execution_dates]


REBAL_DATES = month_first_days(cal_bt)          # execution dates (first trading day of each month)
REBAL_SCHEDULE = rebalance_schedule(REBAL_DATES)


def missing_on(d, symbols=None):
    """ETFs (with enough history) that have NO actual price row on date d (Yahoo gaps)."""
    return [s for s in (SYMBOLS if symbols is None else symbols) if d not in hist_idx[s]]


# ── Frozen live rebalances (live_state/rebalances.json) ──────────────────────
# Yahoo sometimes withdraws a day's prices for many ETFs (seen for Sep 28 and Sep 30, 2026).
# Because the live portfolio is rebuilt from history on every run, such a gap could silently
# rewrite a past signal. Once a signal (and later its fills) is computed from COMPLETE data, it
# is frozen here by the GitHub workflow and replayed by every later run. Rules are unchanged.
FROZEN_PATH = BASE_DIR / "live_state" / "rebalances.json"


def load_frozen():
    try:
        doc = json.loads(FROZEN_PATH.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {"live_start": LIVE_START_DATE, "initial_capital": INITIAL_CAPITAL, "rebalances": {}}, {}
    except ValueError as e:
        print(f"  [warn] {FROZEN_PATH.name} is not valid JSON ({e}) — ignoring frozen rebalances")
        return {"live_start": LIVE_START_DATE, "initial_capital": INITIAL_CAPITAL, "rebalances": {}}, {}
    if doc.get("live_start") != LIVE_START_DATE or doc.get("initial_capital") != INITIAL_CAPITAL:
        print(f"  [warn] {FROZEN_PATH.name} belongs to live start {doc.get('live_start')} / capital "
              f"{doc.get('initial_capital')} — ignored for LIVE_START_DATE {LIVE_START_DATE}")
        return {"live_start": LIVE_START_DATE, "initial_capital": INITIAL_CAPITAL, "rebalances": {}}, {}
    return doc, doc.get("rebalances", {})


def frozen_decisions(rec):
    if not rec or not rec.get("signal"):
        return None
    sig = rec["signal"]
    return {"top6": sig["top6"], "buy": sig["buy"], "skip": sig["skip"],
            "sell": [(x["symbol"], x["reason"]) for x in sig["sell"]], "scores": sig.get("scores", {})}


def frozen_fills(rec):
    if not rec or not rec.get("fills"):
        return None
    f = rec["fills"]
    return {"BUY": {x["symbol"]: (x["qty"], x["price"]) for x in f["buy"]},
            "SELL": {x["symbol"]: x["price"] for x in f["sell"]}}


frozen_doc, FROZEN = load_frozen()
if FROZEN:
    print(f"Frozen live rebalances: " + ", ".join(
        f"{m} ({'signal+fills' if r.get('fills') else 'signal'})" for m, r in sorted(FROZEN.items())))


# ═════════════════════════════════════════════════════════════════════════════
# SIGNAL HELPERS
# ═════════════════════════════════════════════════════════════════════════════
def compute_scores(d):
    """Momentum score + 200-DMA for every ETF with >= MIN_HISTORY rows up to date d."""
    out = {}
    for sym in SYMBOLS:
        n = hist_idx[sym].searchsorted(d, side="right")
        if n < MIN_HISTORY:
            continue
        v = hist_vals[sym][:n]
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
              decisions=None, fills=None, exec_prices=None):
    """One monthly rebalance, in two strictly separated phases:

    PHASE A — SIGNAL (uses ONLY data on or before signal_date = previous trading day):
        momentum, 200-DMA, ranking, top 6, stop/rotation exits, entry gate → decisions
    PHASE B — EXECUTION (uses ONLY prices on execution_date = first trading day of month):
        sells, buys, entry/exit prices, cash, portfolio value

    preview=True is only for the dashboard's "what if the signal were taken on the latest close"
    view: execution prices are then estimates, and the state passed in must be a copy.

    Live-portfolio data integrity (no rule changes):
      decisions   — a FROZEN signal (computed earlier from complete signal-day data) replaces the
                    Phase A recomputation, so later Yahoo data gaps cannot rewrite a past signal.
      fills       — FROZEN fills {"BUY": {sym: (qty, price)}, "SELL": {sym: price}} replay the
                    recorded execution exactly.
      exec_prices — live prices for the pre-close ORDER PLAN (preview only).
    """
    sd, ed = signal_date, execution_date
    if not preview:
        assert sd < ed and sd == prev_trading_day(ed), "signal date must be the previous trading day"
    positions = st["positions"]   # sym -> dict(qty, entry_date, entry_price, peak, cur, score_entry, ...)
    source = "frozen-fills" if fills else ("frozen-signal" if decisions else "computed")

    def xpx(sym):
        """Execution price: frozen fill → live order-plan price → execution-date close."""
        if fills:
            if sym in fills.get("BUY", {}):
                return float(fills["BUY"][sym][1])
            if sym in fills.get("SELL", {}):
                return float(fills["SELL"][sym])
        if exec_prices and sym in exec_prices:
            return float(exec_prices[sym])
        return px(sym, ed)

    # ═══ PHASE A — SIGNAL: data <= signal_date only ═══════════════════════════
    # Mark held positions at the signal-date close and ratchet the peak with it
    for sym, p in positions.items():
        p["sig"] = px(sym, sd)
        if p["sig"] > p["peak"]:
            p["peak"] = p["sig"]
    sc = compute_scores(sd)                     # momentum + 200-DMA through sd
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

    if decisions:                               # frozen signal replaces the recomputed decisions
        top = list(decisions["top6"])
        sell_dec = [(s, r) for s, r in decisions["sell"] if s in positions]
        sold = {s for s, _ in sell_dec}
        hold_dec = [s for s in positions if s not in sold]
        buy_dec = list(decisions["buy"])
        skip_dec = list(decisions["skip"])
    frozen_scores = (decisions or {}).get("scores", {})

    # ═══ PHASE B — EXECUTION: prices on execution_date only ═══════════════════
    # Step 1: liquid interest on idle cash since the previous execution
    if st["prev_d"] is not None:
        st["cash"] += st["cash"] * LIQUID_RATE / 365 * (ed - st["prev_d"]).days

    def log(action, sym, exec_price, **kw):
        i = sc.get(sym, {})
        st["log"].append(dict(
            date=ed.date(), signal_date=sd.date(), action=action, symbol=sym, category=CATEGORY[sym],
            rank=rank_of.get(sym), score=round(i["score"], 3) if i else None,
            signal_price=round(px(sym, sd), 4), dma200=round(i["dma"], 4) if i else None,
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
            if fills and sym in fills.get("BUY", {}):
                qty = int(fills["BUY"][sym][0])               # replay the recorded fill exactly
            else:
                qty = math.floor(min(alloc, st["cash"]) / xp)   # never let cash go negative
            if qty <= 0:
                log("SKIP", sym, xp, qty=0, value=0.0, target=round(alloc, 2),
                    reason="Insufficient cash", pnl_pct=None)
                continue
            st["cash"] -= qty * xp
            score_entry = frozen_scores.get(sym, sc.get(sym, {}).get("score", float("nan")))
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
        print(f"Signal Date    : {sd.date()}")
        print(f"Execution Date : {ed.date()}")
        print("=" * 60)
        print(f"Signal generated from data available through: {sd.date()}"
              + {"frozen-fills": "  [FROZEN signal + fills replayed]",
                 "frozen-signal": "  [FROZEN signal replayed]"}.get(source, ""))
        for act in ("BUY", "SELL", "HOLD", "SKIP"):
            items = [a for a in mine if a["action"] == act]
            print(f"{act}:" + ("" if items else " none"))
            for a in items:
                extra = f" ({a['reason']})" if act in ("SELL", "SKIP") else ""
                print(f"  {a['symbol']}{extra}")
        print("Execution prices:")
        for a in mine:
            if a["action"] in ("BUY", "SELL"):
                print(f"  {a['symbol']} {a['action']} → ₹{a['price']:,.2f} on {ed.date()} "
                      f"(signal close ₹{a['signal_price']:,.2f} on {sd.date()})")
        print(f"After execution: Held {len(positions)} | Cash ₹{st['cash']:,.0f} | Portfolio ₹{pv:,.0f}\n")
    elif verbose:
        print(f"{tag}signal {sd.date()} → exec {ed.date()} | Held: {len(positions)} | "
              f"Cash: ₹{st['cash']:,.0f} | Portfolio: ₹{pv:,.0f}")
    st["prev_d"], st["prev_signal"], st["prev_val"] = ed, sd, pv


def run_strategy(schedule, capital, tag, verbose=True, audit=False, frozen=None):
    """Deterministic simulation: fresh capital, then every (signal_date, execution_date) in order.
    `frozen` (live portfolio only): {"YYYY-MM": record} — frozen signals/fills are replayed."""
    st = new_state(capital)
    for sd, ed in schedule:
        rec = (frozen or {}).get(ed.strftime("%Y-%m"))
        rebalance(st, sd, ed, tag, verbose, audit,
                  decisions=frozen_decisions(rec), fills=frozen_fills(rec))
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
                         peak_price=round(max(p["peak"], cur), 4),
                         sl_level=round(p["entry_price"] * (1 - SL_PCT), 4),
                         trail_level=round(max(p["peak"], cur) * (1 - TRAIL_PCT), 4)))
    return pd.DataFrame(rows)


# ═════════════════════════════════════════════════════════════════════════════
# RUN HISTORICAL BACKTEST (internal reference only — never feeds the live portfolio)
# ═════════════════════════════════════════════════════════════════════════════
print(f"\nHistorical backtest: {REBAL_DATES[0].date()} → {LAST_DATE.date()} ({len(REBAL_DATES)} rebalances, "
      f"signal = previous trading day close, execution = first trading day)")
bt = run_strategy(REBAL_SCHEDULE, INITIAL_CAPITAL, tag="")
daily_df = daily_equity(bt, LAST_DATE)

# ═════════════════════════════════════════════════════════════════════════════
# LIVE PORTFOLIO — fresh start on LIVE_START_DATE (nothing inherited)
# ═════════════════════════════════════════════════════════════════════════════
# The live portfolio is rebuilt deterministically on every run:
#   new_state(INITIAL_CAPITAL)
#   → initialisation: signal on the trading day before LIVE_START_DATE's first session,
#     executed on the first trading day on/after LIVE_START_DATE
#   → every later month: signal on previous trading day, execute on first trading day
#   → today's state.
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
    print(f"\nLive portfolio: initialised {live_init_date.date()} (signal {live_schedule[0][0].date()}) "
          f"with ₹{INITIAL_CAPITAL:,.0f} ({len(live_schedule)} live rebalance(s))\n")
    live = run_strategy(live_schedule, INITIAL_CAPITAL, tag="[LIVE] ", audit=True, frozen=FROZEN)
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
    """Check every rebalance of `st`; print a detailed proof for the rebalance at index `show`."""
    n_checked = 0
    for m in st["monthly"]:
        sd, ed = m["signal_date"], m["rebal_date"]
        assert sd < ed and sd == prev_trading_day(ed), f"{label}: bad signal date {sd} for {ed}"
        rows = [r for r in st["scores"] if r["rebal_date"] == ed.date()]
        for r in rows:                                       # every score/DMA/rank input <= signal date
            s, dma, sp = _independent_score(r["symbol"], sd)
            assert abs(s - r["score"]) < 1e-3 and abs(dma - r["dma200"]) < 1e-3 and abs(sp - r["price"]) < 1e-3, \
                f"{label}: {r['symbol']} signal on {ed.date()} does not match data <= {sd.date()}"
        for a in st["log"]:                                   # every fill at the execution-day price
            # (frozen fills replay the price recorded at the time, even if Yahoo revised it since)
            if a["date"] == ed.date() and a["action"] in ("BUY", "SELL") and a.get("source") != "frozen-fills":
                assert abs(a["price"] - px(a["symbol"], ed)) < 1e-3, f"{label}: {a['symbol']} fill not at {ed.date()} price"
                assert abs(a["signal_price"] - px(a["symbol"], sd)) < 1e-3
        n_checked += 1
    for t in st["trades"]:
        assert t["entry_signal_date"] < t["entry_date"] and t["exit_signal_date"] < t["exit_date"]

    m = st["monthly"][show]
    sd, ed = m["signal_date"], m["rebal_date"]
    rows = sorted((r for r in st["scores"] if r["rebal_date"] == ed.date()), key=lambda r: r["rank"])
    top_sig = [r["symbol"] for r in rows[:N_HOLD]]
    alt = {s: _independent_score(s, ed) for s in SYMBOLS}     # what a same-day (look-ahead) signal would say
    top_alt = [s for s, _ in sorted(((s, v[0]) for s, v in alt.items() if v), key=lambda x: -x[1])[:N_HOLD]]
    print("\n" + "=" * 72)
    print(f"LOOK-AHEAD VALIDATION — {label}: {n_checked} rebalances checked, all passed")
    print("=" * 72)
    print(f"Example rebalance    : signal {sd.date()} ({sd:%a}) → execution {ed.date()} ({ed:%a})")
    print(f"Previous trading day : {prev_trading_day(ed).date()}  (from market calendar)")
    print(f"Top 6 from data <= {sd.date()} : {', '.join(top_sig)}")
    print(f"Top 6 IF {ed.date()} close were used (NOT used): {', '.join(top_alt)}")
    print(f"{'Symbol':<11}{'Action':<7}{'Score@signal':>13}{'Score@exec(unused)':>20}"
          f"{'Signal close':>14}{'Exec price':>12}")
    for a in (x for x in st["log"] if x["date"] == ed.date() and x["action"] in ("BUY", "SELL", "HOLD")):
        s_sig = _independent_score(a["symbol"], sd)[0]
        s_ex = alt[a["symbol"]][0] if alt.get(a["symbol"]) else float("nan")
        print(f"{a['symbol']:<11}{a['action']:<7}{s_sig:>13.3f}{s_ex:>20.3f}"
              f"{a['signal_price']:>14.2f}{a['price']:>12.2f}")
    print("✓ momentum, 200-DMA, ranking and stop/entry decisions use data <= signal date")
    print("✓ every BUY/SELL price equals the execution-date close\n")


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
AS_OF = DATA_LAST                      # latest market date; no data after this is used anywhere
TODAY = pd.Timestamp(now_ist().date())
NEAR_STOP_PCT = 5.0                    # reporting only: flag positions within 5% of a stop

cur_sc = compute_scores(AS_OF)
cur_ranked = sorted(cur_sc, key=lambda k: cur_sc[k]["score"], reverse=True)
cur_rank = {s: i for i, s in enumerate(cur_ranked, 1)}
cur_top = set(cur_ranked[:N_HOLD])

live_started = live is not None
live_state = live if live_started else new_state(INITIAL_CAPITAL)
live_cash = float(live_daily["cash"].iloc[-1]) if live_started else float(INITIAL_CAPITAL)

# Current live positions
pos_rows = []
for sym, p in live_state["positions"].items():
    price = px(sym, AS_OF)
    peak = max(p["peak"], price)
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

# ── Rebalance timing: signal = previous trading day close, execution = first trading day ──
# Future sessions are not in the price data, so future dates are estimates that skip weekends
# only (exchange holidays are unknown in advance). Past/present dates come from the calendar.
def _first_weekday_on_or_after(d):
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d


def _last_weekday_before(d):
    d -= timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


cur_rebal = live["prev_d"] if live_started else None            # last executed execution date
cur_signal = live["prev_signal"] if live_started else None      # ... and its signal date
executed_today = live_started and cur_rebal == AS_OF

next_anchor = live_ts if not live_started else (AS_OF + pd.offsets.MonthBegin(1)).normalize()
_later = cal[cal >= next_anchor]
if len(_later):                                   # (only possible if data runs past the anchor)
    next_rebal, next_confirmed = _later[0], True
else:
    next_rebal, next_confirmed = _first_weekday_on_or_after(next_anchor), False
_est_signal = _last_weekday_before(next_rebal)
if next_confirmed:
    next_signal, next_signal_known = prev_trading_day(next_rebal), True
elif AS_OF >= _est_signal:                        # latest close IS the last session before execution
    next_signal, next_signal_known = AS_OF, True
else:
    next_signal, next_signal_known = _est_signal, False
days_to_next = (next_rebal - TODAY).days

# A signal already frozen (from complete data) for the next execution month is final and is used
# as-is, even if Yahoo has since withdrawn some signal-day prices.
next_month = next_rebal.strftime("%Y-%m")
frozen_next = FROZEN.get(next_month) if not executed_today else None
frozen_next_dec = frozen_decisions(frozen_next)
if frozen_next_dec:
    next_signal, next_signal_known = pd.Timestamp(frozen_next["signal_date"]), True

# Signal-day completeness: a signal computed while ETFs are missing their signal-day price is
# NOT trusted (it would silently use older prices). Up to 2 gaps are tolerated only for ETFs that
# are neither held nor ranked in the top 10.
SIGNAL_GAP_TOLERANCE = 2
signal_missing = missing_on(AS_OF) if (next_signal_known and not executed_today) else []
_key_syms = set(live_state["positions"]) | set(cur_ranked[:10])
signal_incomplete = bool(
    next_signal_known and not executed_today and frozen_next_dec is None and signal_missing
    and (len(signal_missing) > SIGNAL_GAP_TOLERANCE or _key_syms & set(signal_missing)))

# Pre-close ORDER PLAN: market is open on the execution day → yesterday's (final) signal, sized
# with today's live prices. Never recorded as executed; the close-based record is made after 16:00.
order_plan = bool(MARKET_HOURS and next_rebal == TODAY and next_signal_known and not signal_incomplete)
live_quote_time = _now.strftime("%H:%M") if MARKET_HOURS else None

if executed_today:
    signal_status = "EXECUTED"
elif signal_incomplete:
    signal_status = "INCOMPLETE"
elif next_signal_known:
    signal_status = "READY"
else:
    signal_status = "PENDING"

# Actions: the executed rebalance (today), or the SAME engine run on a copy with the latest close
# as signal. For READY/PENDING the execution prices are estimates (latest close); for the order
# plan they are today's live prices.
if executed_today:
    act_mode = "executed"
    act_log = [a for a in live["log"] if a["date"] == AS_OF.date()]
else:
    act_mode = ("orderplan" if order_plan else "incomplete" if signal_status == "INCOMPLETE"
                else "signal" if signal_status == "READY" else "preview")
    _pv = copy.deepcopy(live_state)
    rebalance(_pv, AS_OF, AS_OF, verbose=False, preview=True, decisions=frozen_next_dec,
              exec_prices=live_quotes if order_plan else None)
    act_log = [a for a in _pv["log"] if a["date"] == AS_OF.date() and a["signal_date"] == AS_OF.date()]
acts = {k: [a for a in act_log if a["action"] == k] for k in ("BUY", "HOLD", "SELL", "SKIP")}
signal_source = ("frozen" if frozen_next_dec and act_mode in ("signal", "orderplan") else
                 "frozen" if act_mode == "executed" and FROZEN.get(AS_OF.strftime("%Y-%m")) else "computed")
order_plan_missing_quotes = [a["symbol"] for a in act_log
                             if order_plan and a["action"] in ("BUY", "SELL", "HOLD") and a["symbol"] not in live_quotes]
if MARKET_HOURS:
    print(f"\n  [MARKET OPEN] Today's ({TODAY.date()}) prices are live quotes, not closes — excluded from "
          f"the price history; data shown up to {AS_OF.date()}.")
if order_plan:
    print(f"  [ORDER PLAN] Execution day: signal {next_signal.date()} ({signal_source}), "
          f"sized with live prices at {live_quote_time} IST — NOT executed")
if signal_incomplete:
    print(f"  [SIGNAL INCOMPLETE] {len(signal_missing)} ETFs have no {AS_OF.date()} price: "
          f"{', '.join(signal_missing)} — do not trade on this signal")

# ═════════════════════════════════════════════════════════════════════════════
# FREEZE completed live rebalances (GitHub Actions only; the workflow commits the file)
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


freeze_notes = []
_frozen_new = copy.deepcopy(FROZEN)
_stamp = {"at_ist": now_ist().strftime("%Y-%m-%d %H:%M"), "run_id": os.environ.get("GITHUB_RUN_ID")}
# 1) Signal for the upcoming execution month: only from a COMPLETE signal day (zero gaps)
if (signal_status == "READY" and frozen_next_dec is None and not missing_on(AS_OF)
        and next_month not in _frozen_new):
    _frozen_new[next_month] = {
        "signal_date": str(AS_OF.date()), "expected_execution_date": str(next_rebal.date()),
        "signal": _decisions_from_log(act_log, cur_ranked[:N_HOLD],
                                      {s: cur_sc[s]["score"] for s in cur_ranked[:10]}),
        "signal_frozen": {**_stamp, "data": f"{len(SYMBOLS)}/{len(SYMBOLS)} ETFs priced on {AS_OF.date()}"},
    }
    freeze_notes.append(f"signal for {next_month} (signal date {AS_OF.date()})")
# 2) Fills of executed live rebalances: execution-day closes present for every traded/held ETF
if live_started:
    for m in live["monthly"]:
        sd, ed = m["signal_date"], m["rebal_date"]
        month = ed.strftime("%Y-%m")
        rec = _frozen_new.get(month, {})
        if rec.get("fills"):
            continue
        entries = [a for a in live["log"] if a["date"] == ed.date()]
        traded = {a["symbol"] for a in entries if a["action"] in ("BUY", "SELL", "HOLD")}
        if missing_on(ed, list(traded & set(SYMBOLS))):
            continue                                  # execution-day closes not all available yet
        if not rec.get("signal") and missing_on(sd):
            continue                                  # signal day was incomplete — never freeze it
        if not rec.get("signal"):
            rows = sorted((r for r in live["scores"] if r["rebal_date"] == ed.date()), key=lambda r: r["rank"])
            rec = {"signal_date": str(sd.date()), "expected_execution_date": str(ed.date()),
                   "signal": _decisions_from_log(entries, [r["symbol"] for r in rows[:N_HOLD]],
                                                 {r["symbol"]: r["score"] for r in rows[:10]}),
                   "signal_frozen": {**_stamp, "data": f"complete on {sd.date()}"}}
        rec["execution_date"] = str(ed.date())
        rec["fills"] = {
            "buy": [{"symbol": a["symbol"], "qty": int(a["qty"]), "price": a["price"]}
                    for a in entries if a["action"] == "BUY"],
            "sell": [{"symbol": a["symbol"], "qty": int(a["qty"]), "price": a["price"], "reason": a["reason"]}
                     for a in entries if a["action"] == "SELL"],
            "hold": [a["symbol"] for a in entries if a["action"] == "HOLD"],
            "cash_after": m["cash"], "value_after": m["portfolio_value"],
            "basis": "execution-day close (strategy record)",
        }
        rec["fills_frozen"] = {**_stamp, "data": f"closes on {ed.date()} for all traded ETFs"}
        _frozen_new[month] = rec
        freeze_notes.append(f"fills for {month} (executed {ed.date()})")
if freeze_notes:
    if IN_CI:
        FROZEN_PATH.parent.mkdir(parents=True, exist_ok=True)
        FROZEN_PATH.write_text(json.dumps({"live_start": LIVE_START_DATE, "initial_capital": INITIAL_CAPITAL,
                                           "rebalances": dict(sorted(_frozen_new.items()))},
                                          indent=2, ensure_ascii=False), encoding="utf-8")
        print("  [FREEZE] Saved " + "; ".join(freeze_notes) + f" → {FROZEN_PATH.relative_to(BASE_DIR)}")
    else:
        print("  [FREEZE] Would freeze " + "; ".join(freeze_notes) + " (only GitHub Actions writes the file)")

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
        price_date=last_px_date[sym].date(), behind=sessions_behind[sym]))
for r in pos_rows:
    r["price_date"] = last_px_date[r["symbol"]].date()
    r["behind"] = sessions_behind[r["symbol"]]

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
    print(f"  Last rebalance  signal {cur_signal.date()} → executed {cur_rebal.date()}")
print(f"  Latest data     {AS_OF.date()}")
print(f"  Next signal     {next_signal.date()}{'' if next_signal_known else ' (expected)'}   "
      f"Next execution {next_rebal.date()}{'' if next_confirmed else ' (expected)'}   Status {signal_status}")
print(f"  [Historical reference {first_d.date()}→{LAST_DATE.date()}: CAGR {cagr:.2f}% | "
      f"MaxDD {max_dd_daily:.2f}% | Sharpe {sharpe:.2f}]")
print("═" * 60)

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
    ("Last Rebalance Signal Date", str(cur_signal.date()) if cur_signal is not None else "n/a"),
    ("Last Rebalance Execution Date", str(cur_rebal.date()) if cur_rebal is not None else "n/a"),
    ("Latest Market Data", str(AS_OF.date())),
    ("Next Signal Date", f"{next_signal.date()}{'' if next_signal_known else ' (expected)'}"),
    ("Next Execution Date", f"{next_rebal.date()}{'' if next_confirmed else ' (expected)'}"),
    ("Signal Status", signal_status),
    ("Methodology", "Signal = previous trading day close; execution = first trading day close"),
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

def _write_excel():
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as xw:
        live_summary_df.to_excel(xw, sheet_name="Live_Summary", index=False)
        _or_empty(live_pos_xl, "Symbol").to_excel(xw, sheet_name="Live_Positions", index=False)
        _or_empty(live_act_xl, "action").to_excel(xw, sheet_name="Live_Actions", index=False)
        live_sig_xl.to_excel(xw, sheet_name="Live_Signals", index=False)
        live_data_xl.to_excel(xw, sheet_name="Live_Data_Status", index=False)
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

_first_fig = [True]


def fig_to_div(fig):
    cfg = {"responsive": True, "displaylogo": False}
    if _first_fig[0]:
        _first_fig[0] = False
        return pyo.plot(fig, include_plotlyjs=True, output_type="div", config=cfg)
    return pyo.plot(fig, include_plotlyjs=False, output_type="div", config=cfg)


def style(fig, title, h=400):
    fig.update_layout(title=title, template="plotly_dark", paper_bgcolor=PANEL, plot_bgcolor=PANEL,
                      font=dict(color=TXT), height=h, margin=dict(l=60, r=25, t=55, b=45),
                      legend=dict(orientation="h", y=1.1, x=0))
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
    """Last actual price date; ⚠ when older than the latest market date (price carried forward)."""
    return f"⚠ {v} ({sessions_behind_by_date(v)} behind)" if v < AS_OF.date() else f"✓ {v}"


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
    if signal_source == "frozen" and act_mode in ("signal", "orderplan"):
        _lines.append("✓ <b>Signal not affected:</b> the rebalance signal is frozen from complete "
                      f"{next_signal.date()} data, so these gaps cannot change it.")
    elif stale_signal or stale_actions:
        _lines.append(f"⚠ <b>Signal affected:</b> {', '.join(sorted(set(stale_signal) | set(stale_actions)))} "
                      f"— ranking and actions may change once the missing prices arrive.")
    if signal_source == "frozen" and act_mode in ("signal", "orderplan"):
        pass
    elif signal_status == "READY" and (stale_signal or stale_actions or market_delayed):
        _lines.append("<b>Do not act on this signal yet.</b> Rerun later with <code>--force</code> "
                      "once all prices are available.")
    else:
        _lines.append("Rerun later with <code>--force</code> to refresh once Yahoo has the missing prices.")
    data_html = ("<div class='datawarn'><div class='dw-title'>⚠ DATA DELAYED</div>"
                 + "".join(f"<p>{l}</p>" if not l.startswith("<ul>") else l for l in _lines) + "</div>")
if near_miss:
    data_html += ("<p class='muted'>Excluded for short Yahoo history, only a few rows below the "
                  f"{MIN_HISTORY}-row minimum (a delayed day can cause this): "
                  + ", ".join(f"{s} ({n}/{MIN_HISTORY})" for s, n in near_miss.items()) + "</p>")
act_data_warn = ("" if not (stale_actions or stale_signal or market_delayed)
                 or (signal_source == "frozen" and act_mode in ("signal", "orderplan")) else
                 "<div class='datawarn slim'>⚠ Some prices behind these actions are delayed — "
                 "see the data warning at the top. Actions may change after the data catches up.</div>")

# ── Header / KPI cards ───────────────────────────────────────────────────────
live_start_label = str(live_init_date.date()) if live_started else LIVE_START_DATE
state_note = ("" if live_started else
              f"<div class='note'>Live portfolio has not started yet — no market data on/after "
              f"{LIVE_START_DATE}. Figures below show a fresh ₹{INITIAL_CAPITAL:,} portfolio and the "
              f"initialisation preview using the latest data ({AS_OF.date()}). The real initial signal "
              f"uses the close of the trading day before the first session on/after {LIVE_START_DATE}, "
              f"and the purchases execute on that first session.</div>")

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

kpis_perf = "".join([
    kpi("Live Return %", pct(live_ret), pn(live_ret)),
    kpi("Live P&L", inr(live_pnl), pn(live_pnl)),
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
market_html = ("" if not MARKET_HOURS else
               f"<div class='marketopen'>🕒 <b>Market open</b> — today's ({TODAY.date()}) prices are live quotes, "
               f"not closing prices, so they are not used in the portfolio history. Values below are as of the "
               f"{AS_OF.date()} close" + (" (the order plan uses live prices)." if order_plan else ".") + "</div>")
_src_txt =(f" Signal frozen from complete {next_signal.date()} data "
            f"({(frozen_next or {}).get('signal_frozen', {}).get('at_ist', '')} IST)." if signal_source == "frozen" else "")
if act_mode == "orderplan":
    act_banner = (f"🛒 ORDER PLAN — execute TODAY ({TODAY.date()}) before the close · "
                  f"live prices at {live_quote_time} IST · NOT executed yet")
    act_caption = (f"Final signal from the {next_signal.date()} close.{_src_txt} Quantities = target ÷ the live "
                   f"price at {live_quote_time} IST (Yahoo may lag — check your broker's price; quantity = target ÷ "
                   f"price, whole units). The strategy record is made from today's close after 16:00 IST."
                   + (f" No live quote for: {', '.join(order_plan_missing_quotes)} — last close used."
                      if order_plan_missing_quotes else ""))
    act_prefix = "ORDER"
elif act_mode == "incomplete":
    act_banner = "⛔ SIGNAL DATA INCOMPLETE — do not trade on this signal"
    act_caption = (f"{len(signal_missing)} ETFs have no {AS_OF.date()} price on Yahoo "
                   f"({', '.join(signal_missing)}), so the ranking below uses older prices for them and may be "
                   f"wrong. Wait for the data to complete and re-run (tonight's 20:15 IST run re-checks).")
    act_prefix = "UNRELIABLE"
elif act_mode == "executed":
    act_banner = f"EXECUTED — {AS_OF:%B} rebalance executed on {AS_OF.date()}"
    act_caption = (f"Signal from the {cur_signal.date()} close; orders filled at the {AS_OF.date()} "
                   f"execution-day price shown.")
    act_prefix = "EXECUTED"
elif act_mode == "signal":
    act_banner = (f"SIGNAL FINAL — generated from the {AS_OF.date()} close · execution scheduled for "
                  f"{next_rebal.date()}{_exp(next_confirmed)}")
    act_caption = ("Planned orders. Not executed yet: quantities and values use the signal-day close as an "
                   "estimate; actual fills use the execution-day price." + _src_txt)
    act_prefix = "SIGNAL"
else:
    act_banner = "PREVIEW — signal may change before signal date"
    act_caption = (f"Uses the latest close ({AS_OF.date()}), which is not the final signal. Final signal: "
                   f"{next_signal.date()}{_exp(next_signal_known)} close · execution: "
                   f"{next_rebal.date()}{_exp(next_confirmed)}. Prices are estimates.")
    act_prefix = "PREVIEW"
    if not live_started:
        act_caption = (f"Initialisation preview for a fresh {inr(INITIAL_CAPITAL)} using the latest close "
                       f"({AS_OF.date()}). " + act_caption.split(". ", 1)[1])

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
pos_cols = [
    ("rank", "Rank", lambda v: "—" if v is None else str(v)), ("symbol", "Symbol", None),
    ("category", "Category", None), ("score", "Momentum Score", num),
    ("r1", "1M", pct), ("r3", "3M", pct), ("r6", "6M", pct), ("r12", "12M", pct),
    ("dma", "200 DMA", num), ("price", "Current Price", num), ("above_dma", "Above 200 DMA", yn),
    ("entry_date", "Entry Date", None), ("entry_price", "Entry Price", num), ("qty", "Qty", str),
    ("cost", "Cost", inr), ("value", "Current Value", inr), ("pnl", "Unrealised P&L ₹", inr),
    ("pnl_pct", "Unrealised P&L %", pct), ("hold_days", "Hold Days", str),
    ("hard_sl", "Hard SL", num), ("trail_sl", "Trailing SL", num),
    ("dist_sl", "Distance to SL %", lambda v: pct(v, False)),
    ("dist_trail", "Distance to Trail %", lambda v: pct(v, False)),
    ("weight", "Position Weight %", lambda v: pct(v, False)), ("status", "Status", badge),
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
    ("symbol", "Symbol", None), ("price", "Current Price", num), ("entry_price", "Entry Price", num),
    ("peak", "Peak Price", num), ("hard_sl", "Hard SL", num), ("trail_sl", "Trailing SL", num),
    ("dist_sl", "Distance to Hard SL %", lambda v: pct(v, False)),
    ("dist_trail", "Distance to Trailing SL %", lambda v: pct(v, False)),
    ("nearest", "Nearest Stop", None), ("cushion", "Cushion %", lambda v: pct(v, False)),
    ("status", "Status", badge), ("price_date", "Price Date", lambda v: pxdate(v)),
]
risk_html = table(risk_rows, risk_cols, "tRisk",
                  lambda r: "warn" if r["cushion"] <= NEAR_STOP_PCT else "",
                  lambda k, r: ("r" if r[k] <= 0 else "amber" if r[k] <= NEAR_STOP_PCT else "")
                  if k in ("dist_sl", "dist_trail", "cushion") else ("sym" if k == "symbol" else ""))

sig_cols = [
    ("rank", "Rank", str), ("symbol", "Symbol", None), ("category", "Category", None),
    ("score", "Momentum Score", num), ("r1", "1M %", pct), ("r3", "3M %", pct), ("r6", "6M %", pct),
    ("r12", "12M %", pct), ("dma", "200 DMA", num), ("price", "Current Price", num),
    ("above_dma", "Above 200 DMA", yn), ("held", "Currently Held", yn), ("signal", "Signal", badge),
    ("target", "Target Allocation", inr), ("est_qty", "Est. Qty", lambda v: "—" if v is None else str(v)),
    ("price_date", "Price Date", lambda v: pxdate(v)),
]
signals_html = table(
    sig_rows, sig_cols, "tSig",
    lambda r: "top" if r["rank"] <= N_HOLD else "",
    lambda k, r: ("g" if r[k] else "r") if k == "above_dma" else
    (pn(r[k]) if k in ("r1", "r3", "r6", "r12") else ("sym" if k == "symbol" else "")))

trade_cols = [
    ("trade_id", "Trade #", str), ("symbol", "Symbol", None), ("category", "Category", None),
    ("entry_signal_date", "Entry Signal Date", None), ("entry_date", "Entry Date (exec)", None),
    ("exit_signal_date", "Exit Signal Date", None), ("exit_date", "Exit Date (exec)", None),
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
    ("sdate", "Signal Date", None), ("date", "Execution Date", None),
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
    f_pnl = go.Figure(go.Bar(
        x=[r["symbol"] for r in pos_rows], y=[r["pnl"] for r in pos_rows],
        marker_color=[GRN if r["pnl"] >= 0 else RD for r in pos_rows],
        customdata=[[r["pnl_pct"], r["value"]] for r in pos_rows],
        text=[f"{r['pnl_pct']:+.1f}%" for r in pos_rows], textposition="outside",
        hovertemplate="%{x}<br>P&L ₹%{y:,.0f}<br>P&L %{customdata[0]:+.2f}%<br>Value ₹%{customdata[1]:,.0f}<extra></extra>"))
    style(f_pnl, "Position P&L (unrealised ₹)", 420)
    f_pnl.update_layout(yaxis_title="Unrealised P&L ₹", showlegend=False)
    charts["pnl"] = fig_to_div(f_pnl)
if live_started:
    f_val = go.Figure()
    f_val.add_trace(go.Scatter(x=live_daily.index, y=live_daily["value"], name="Portfolio Value",
                               line=dict(color=AMBER, width=2.5)))
    f_val.add_trace(go.Scatter(x=live_daily.index, y=live_daily["held"], name="Invested Value",
                               line=dict(color=BLUE, width=1.5)))
    f_val.add_trace(go.Scatter(x=live_daily.index, y=live_daily["cash"], name="Cash",
                               line=dict(color=GRAY, width=1, dash="dot")))
    f_val.add_hline(y=INITIAL_CAPITAL, line=dict(color=MUTED, dash="dash", width=1),
                    annotation_text=f"Initial {inr(INITIAL_CAPITAL)}", annotation_font_color=MUTED)
    style(f_val, "Live Portfolio Value", 430)
    f_val.update_layout(xaxis_title="Date", yaxis_title="Portfolio Value ₹", yaxis_tickformat=",.0f")
    charts["val"] = fig_to_div(f_val)

    f_dd = go.Figure(go.Scatter(x=live_daily.index, y=live_daily["dd_pct"], fill="tozeroy",
                                name="Drawdown %", line=dict(color=RD, width=1),
                                fillcolor="rgba(231,76,60,0.35)"))
    f_dd.add_hline(y=live_max_dd, line=dict(color=AMBER, dash="dash"),
                   annotation_text=f"Max live DD {live_max_dd:.2f}%", annotation_font_color=AMBER)
    style(f_dd, "Live Portfolio Drawdown", 330)
    f_dd.update_layout(xaxis_title="Date", yaxis_title="Drawdown %")
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
    "EXECUTED": f"{AS_OF:%B} rebalance executed on {AS_OF.date()} (signal {cur_signal.date() if cur_signal is not None else '—'}).",
    "READY": f"Signal finalized using {next_signal.date()} close · execution scheduled for {next_rebal.date()}{_exp(next_confirmed)}.",
    "PENDING": f"Signal date not reached — latest data is {AS_OF.date()}. Current signals are a PREVIEW only.",
    "INCOMPLETE": f"Signal-day prices missing for {len(signal_missing)} ETFs — signal not reliable yet.",
}[signal_status]
if order_plan:
    _status_txt += f" Execution day — ORDER PLAN uses live prices at {live_quote_time} IST."
next_html = f"""<div class='kpis'>
{kpi("Latest Market Data", str(AS_OF.date()))}
{kpi("Last Rebalance", (f"{cur_signal.date()} → {cur_rebal.date()}" if cur_rebal is not None else "Not started"))}
{kpi("Next Signal Date" + _exp(next_signal_known), next_signal.strftime("%a %d %b %Y"))}
{kpi("Next Execution Date" + _exp(next_confirmed), next_rebal.strftime("%a %d %b %Y"))}
{kpi("Signal Status", signal_status, {"EXECUTED": "g", "READY": "amber", "PENDING": "", "INCOMPLETE": "r"}[signal_status])}
{kpi("Days to Execution", str(max(days_to_next, 0)))}
</div>
<p class='muted'>{_status_txt} Signal = previous trading day's close; execution = first trading day of the
month. {"" if next_confirmed and next_signal_known else
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
    ("Rebalance timing", "Signal (momentum, 200-DMA, ranking, stops) on the previous trading day's close; "
                         "execution on the first trading day of the month at that day's close"),
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

page_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ETF Momentum Live Dashboard</title>
<style>
 * {{ box-sizing:border-box; }}
 body {{ background:{BG}; color:{TXT}; font-family:Segoe UI,Roboto,Arial,sans-serif; margin:0; padding:24px; }}
 .wrap {{ max-width:1500px; margin:0 auto; }}
 h1 {{ margin:0 0 6px; font-size:26px; }}
 h2 {{ font-size:18px; letter-spacing:.3px; border-bottom:1px solid #2E4053; padding-bottom:6px; margin:34px 0 12px; }}
 h4 {{ margin:0 0 8px; font-size:13px; letter-spacing:.5px; }}
 .sub {{ display:flex; flex-wrap:wrap; gap:8px 22px; color:{MUTED}; font-size:14px; }}
 .sub b {{ color:{TXT}; font-weight:600; }}
 .muted {{ color:{MUTED}; font-size:13px; }}
 .banner {{ display:inline-block; padding:6px 12px; border-radius:6px; font-weight:700; font-size:13px; margin-bottom:6px; }}
 .banner.preview {{ background:#34495E; color:#fff; }} .banner.signal {{ background:{AMBER}; color:#2B1D02; }}
 .banner.executed {{ background:{GRN}; color:#0B2716; }}
 .banner.orderplan {{ background:{BLUE}; color:#0B1F2E; }} .banner.incomplete {{ background:{RD}; color:#fff; }}
 .marketopen {{ background:rgba(93,173,226,0.12); border:1px solid {BLUE}; border-radius:8px; padding:8px 14px; margin:12px 0; font-size:14px; }}
 .datawarn {{ background:rgba(231,76,60,0.12); border:1px solid {RD}; border-left:5px solid {RD};
   border-radius:8px; padding:10px 16px; margin:14px 0; font-size:14px; }}
 .datawarn p {{ margin:6px 0; }} .datawarn ul {{ margin:4px 0 6px; padding-left:22px; }}
 .datawarn.slim {{ padding:6px 12px; margin:6px 0; font-size:13px; }}
 .dw-title {{ color:{RD}; font-weight:700; letter-spacing:.5px; }}
 .dataok {{ color:{GRN}; font-size:13px; margin:12px 0 0; }}
 td.stale {{ color:{AMBER}; font-weight:600; }}
 .note {{ background:#3B2F10; border:1px solid {AMBER}; border-radius:8px; padding:10px 14px; margin:14px 0; }}
 .kpis {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:12px; margin:14px 0; }}
 .kpi {{ background:{PANEL}; border-radius:10px; padding:12px 16px; }}
 .kl {{ color:{MUTED}; font-size:12px; text-transform:uppercase; letter-spacing:.4px; }}
 .kv {{ font-size:24px; font-weight:600; margin-top:4px; font-variant-numeric:tabular-nums; }}
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
 .tw {{ overflow:auto; max-height:640px; border-radius:10px; background:{PANEL}; }}
 table {{ border-collapse:collapse; width:100%; font-size:13px; font-variant-numeric:tabular-nums; }}
 th {{ background:#22364A; position:sticky; top:0; cursor:pointer; padding:8px; text-align:left; white-space:nowrap; z-index:1; }}
 th:hover {{ color:{AMBER}; }}
 td {{ padding:6px 8px; border-bottom:1px solid #22364A; white-space:nowrap; }}
 tr.pos td {{ background:rgba(46,204,113,0.10); }} tr.neg td {{ background:rgba(231,76,60,0.12); }}
 tr.warn td {{ background:rgba(245,166,35,0.14); }} tr.top td {{ background:rgba(93,173,226,0.08); }}
 .badge {{ padding:2px 8px; border-radius:10px; font-size:11px; font-weight:700; letter-spacing:.3px; }}
 .b-buy {{ background:{GRN}; color:#0B2716; }} .b-hold {{ background:{BLUE}; color:#0B1F2E; }}
 .b-sell {{ background:{RD}; color:#fff; }} .b-risk {{ background:{AMBER}; color:#2B1D02; }}
 .b-skip {{ background:#7D6608; color:#fff; }} .b-out {{ background:#34495E; color:#D5DBDB; }}
 table.rules td {{ white-space:normal; overflow-wrap:anywhere; }}   /* long rule text wraps on phones */
 table.rules td:first-child {{ color:{MUTED}; width:220px; }}
 details.hist {{ margin-top:40px; background:#122130; border-radius:10px; padding:10px 16px; opacity:.9; }}
 details.hist summary {{ cursor:pointer; color:{MUTED}; font-weight:600; }}
 @media (max-width:600px) {{ body {{ padding:16px; }} .grid2 {{ grid-template-columns:1fr; }} .kv {{ font-size:20px; }} }}
</style>
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
// Charts are drawn while the page is still loading; re-fit each one to its final box
window.addEventListener('load', resizeCharts);
document.addEventListener('toggle', resizeCharts, true);
</script></head><body><div class="wrap">

<h1>ETF Momentum Rotation — Live Portfolio Dashboard</h1>
<div class="sub">
 <span>Live Start: <b>{live_start_label}</b></span>
 <span>Initial Capital: <b>{inr(INITIAL_CAPITAL)}</b></span>
 <span>Strategy: <b>ETF Momentum Rotation</b></span>
 <span>Universe: <b>C54</b></span>
 <span>Latest Market Data: <b>{AS_OF.date()}</b></span>
 <span>Next Signal Date: <b>{next_signal.date()}{_exp(next_signal_known)}</b></span>
 <span>Next Execution Date: <b>{next_rebal.date()}{_exp(next_confirmed)}</b></span>
 <span>Status: <b>{signal_status}</b></span>
 <span>Data: <b class="{'g' if data_ok else 'amber'}">{'✓ Current' if data_ok else '⚠ Delayed'}</b></span>
 <span>Generated: <b>{now_ist():%Y-%m-%d %H:%M} IST</b></span>
</div>
{market_html}
{data_html}
{state_note}
<div class="kpis k4">{kpis_port}</div>

<h2>Live Portfolio Actions</h2>
{actions_html}

<h2>Current Live Positions</h2>
<p class='muted'>Positions opened on or after {live_start_label} only. Rows: green = profit, red = loss,
amber = exit risk (stop hit, outside top 6, or within {NEAR_STOP_PCT:.0f}% of a stop). Click a header to sort.</p>
{positions_html}

<h2>Risk Monitor</h2>
<p class='muted'>Sorted by the position closest to an exit. Peak includes the latest close.</p>
{risk_html}

<h2>Current Strategy Signals</h2>
<p class='muted'><b>{"PREVIEW — " if signal_status == "PENDING" else ""}</b>Full ranked universe from the
{AS_OF.date()} close{" (not the final signal)" if signal_status == "PENDING" else ""}. Top {N_HOLD} are highlighted. Target allocation
= current portfolio value / {N_HOLD} = {inr(target_now)}.</p>
{signals_html}

<h2>Live Performance</h2>
<div class="kpis">{kpis_perf}</div>

<h2>Next Rebalance</h2>
{next_html}

<h2>Live Charts</h2>
{charts_html}

<h2>Live Rebalance History</h2>
{rebal_html}

<h2>Live Closed Trades</h2>
{live_trades_html}
<div class="kpis small">{trade_stats_html}</div>

<h2>Strategy Rules</h2>
<table class="rules">{rules_html}</table>

{hist_html}
</div></body></html>"""

try:
    HTML_PATH.write_text(page_html, encoding="utf-8")
except Exception as e:
    print(f"[HTML ERROR] Could not write {HTML_PATH}: {e}")
    sys.exit(4)
print(f"Saved: {HTML_PATH.name}")
print(f"  → {HTML_PATH}")

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
    "signal_status": signal_status,               # PENDING / READY / INCOMPLETE / EXECUTED
    "act_mode": act_mode,                         # preview / signal / orderplan / incomplete / executed
    "market_hours": MARKET_HOURS,
    "live_quote_time": live_quote_time,
    "order_plan_missing_quotes": order_plan_missing_quotes,
    "signal_source": signal_source,               # frozen / computed
    "signal_missing": signal_missing,
    "frozen_months": sorted(FROZEN),
    "freeze_notes": freeze_notes,
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
except Exception as e:
    print(f"[SUMMARY ERROR] Could not write {SUMMARY_PATH}: {e}")
    sys.exit(5)
print(f"Saved: {SUMMARY_PATH.name}")

open_report(HTML_PATH)
