"""
notify_telegram.py — Telegram alerts for the ETF momentum live dashboard.

Reads reports/summary.json (written by ETF_Momentum_Test_1.py) and sends:
  🛒 order plan                             (entry day, run during market hours: live prices)
  ⛔ data incomplete                        (live prices missing on the entry day — do not trade)
  ✅ rebalance recorded                      (entry day's close, status EXECUTED)
  ⚠️ exit risk                               (stop hit, near a stop, or outside the top 6)
  👤 your account (reference)                (stop levels from your own fills hit / near — info only)
  ⚠️ data delayed                            (held / top-6 / action ETFs without the latest price)
  📊 daily summary                           (evening and manual runs)
  ❌ run failed                              (--failure)

Alert messages are de-duplicated against the previously published summary, so the
same alert is not repeated on every run (manual runs always send everything).
Standard library only. Never prints the bot token.

Usage:
  python notify_telegram.py --summary reports/summary.json --previous reports/previous_summary.json --mode evening
  (--mode morning is kept for compatibility; no morning schedule is used any more)
  python notify_telegram.py --failure "Data download problem"
  python notify_telegram.py --summary reports/summary.json --mode manual --dry-run
Env: TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID (GitHub secrets); GITHUB_* set by Actions.
"""
import argparse
import html
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MAX_LEN = 4000   # Telegram limit is 4096 characters per message


# ── formatting ────────────────────────────────────────────────────────────────
def esc(s):
    return html.escape(str(s), quote=False)


def inr(v):
    """₹ with Indian digit grouping (₹3,00,000)."""
    if v is None:
        return "—"
    neg, whole = v < 0, f"{abs(v):.0f}"
    if len(whole) > 3:
        head, tail, parts = whole[:-3], whole[-3:], []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        whole = ",".join(parts) + "," + tail
    return ("-" if neg else "") + "₹" + whole


def pct(v, sign=True):
    if v is None:
        return "—"
    return f"{v:+.2f}%" if sign else f"{v:.2f}%"


def num(v):
    return "—" if v is None else f"{v:,.2f}"


def d(s):
    """'2026-10-01' → 'Thu 01 Oct 2026'."""
    return date.fromisoformat(s).strftime("%a %d %b %Y") if s else "—"


def month_name(s):
    return date.fromisoformat(s).strftime("%B") if s else ""


def links():
    repo = os.environ.get("GITHUB_REPOSITORY", "")
    page = ""
    if "/" in repo:
        owner, name = repo.split("/", 1)
        page = f"https://{owner.lower()}.github.io/{name}/"
    run_id = os.environ.get("GITHUB_RUN_ID")
    run = f"{os.environ.get('GITHUB_SERVER_URL', 'https://github.com')}/{repo}/actions/runs/{run_id}" if run_id else ""
    return page, run


def footer():
    page, _ = links()
    return f'\n🔗 <a href="{page}">Open report</a>' if page else ""


# ── message builders ─────────────────────────────────────────────────────────
def msg_summary(s, prev):
    lines = [f"📊 <b>ETF Momentum — Daily Summary</b>",
             (f"🟢 Live prices {esc(s['live_quote_time'])} IST (market open) · last close {d(s['as_of'])}"
              if s.get("live_view") else f"Data as of {d(s['as_of'])}") + f" · generated {esc(s['generated_ist'])} IST", ""]
    if s["live_started"]:
        lines += [f"💼 Value <b>{inr(s['value'])}</b> ({pct(s['ret_pct'])})",
                  f"P&amp;L {inr(s['pnl'])} · Cash {inr(s['cash'])} "
                  f"({pct(s['cash'] / s['value'] * 100, False)}) · Max DD {pct(s['max_dd_pct'], False)}",
                  "", f"<b>Positions ({s['n_positions']}/{s['n_hold']})</b>"]
        for p in s["positions"]:
            flag = "" if p["status"] == "HOLD" else f"  ⚠️ {esc(p['status'].replace('EXIT RISK — ', ''))}"
            lines.append(f"• {esc(p['symbol'])}  {pct(p['pnl_pct'])}  {inr(p['value'])}{flag}")
    else:
        buys = ", ".join(esc(a["symbol"]) for a in s["actions"]["BUY"]) or "none"
        basis = ("order plan, live prices" if s.get("act_mode") in ("orderplan", "incomplete")
                 else "preview from the latest close")
        lines += [f"⏳ Live portfolio starts <b>{esc(s['live_start'])}</b> with {inr(s['initial_capital'])}.",
                  f"Initial buys ({basis}): {buys}"]
    exp = lambda known: "" if known else " (expected)"
    lines += ["", f"🗓 Next entry day (signal + execution): {d(s['next_exec_date'])}"
                  f"{exp(s['next_exec_confirmed'])} · <b>{s['signal_status']}</b>"]
    if s["signal_status"] == "ENTRY DAY":
        lines.append("ℹ️ Today is the entry day — run the workflow during market hours (~14:00 IST) "
                     "for the order plan from live prices.")
    data = s["data"]
    if data["status"] == "CURRENT":
        lines.append(f"📡 Data: ✓ current ({data['etfs_loaded']}/{data['universe']} ETFs)")
    else:
        n = sum(len(v) for v in data["stale"].values())
        lines.append(f"📡 Data: ⚠️ delayed ({n} ETFs without a {s['as_of']} price)")
    if prev and prev.get("as_of") == s["as_of"] and prev.get("generated_ist", "")[:10] != s["generated_ist"][:10]:
        lines.append("ℹ️ No new market data since the last run (holiday, or Yahoo not updated yet).")
    lines += account_lines(s)
    return "\n".join(lines) + footer()


def account_lines(s):
    """'My account' block from the Google Sheet (real trades)."""
    a, err = s.get("account"), s.get("account_error")
    if a is None:
        return [] if err in (None, "not configured") else ["", f"👤 My account: ⚠️ could not read the Google Sheet ({esc(err)[:120]})"]
    out = ["", "👤 <b>My account</b> (Google Sheet)",
           f"Value <b>{inr(a['value'])}</b> ({pct(a['ret_pct'])}) · P&amp;L {inr(a['pnl'])}",
           f"Unrealised {inr(a['upnl'])} · Realised {inr(a['realised'])} · Charges {inr(a['charges'])} · "
           f"Cash {inr(a['cash'])}"]
    perf = a.get("perf_live") or a.get("perf")
    if perf:
        live_tag = f" (live {esc(perf['time'])})" if a.get("perf_live") else ""
        out.append(f"📈 Since {perf['start']}{live_tag}: account {pct(perf['account_pct'])} · strategy "
                   f"{pct(perf['strategy_pct'])} · Nifty 500 {pct(perf['nifty_pct'])}")
    if a.get("xirr") is not None:
        out.append(f"XIRR (annualised): {pct(a['xirr'])}")
    if a.get("exec_total") is not None:
        verb = "saved" if a["exec_total"] < 0 else "cost"
        out.append(f"🧾 Execution vs close: {verb} {inr(abs(a['exec_total']))} (incl. charges)")
    if a.get("latest_month"):
        if not a.get("latest_entered"):
            out.append(f"📝 No trades entered for {a['latest_month']} yet")
        elif a["latest_mismatches"] == 0:
            out.append(f"✓ {a['latest_month']} trades match the strategy plan")
        else:
            out.append(f"⚠️ {a['latest_mismatches']} difference(s) from the {a['latest_month']} strategy plan")
    if a.get("issues"):
        out.append(f"⚠️ {len(a['issues'])} sheet row(s) have problems — see the report")
    return out


def msg_order_plan(s):
    a = s["actions"]
    month = month_name(s["next_exec_date"])
    lines = [f"🛒 <b>ORDER PLAN — {month} entry day, execute TODAY before close</b>",
             f"{d(s['next_exec_date'])} · signal and quantities from live prices at "
             f"{esc(s['live_quote_time'])} IST · <b>not recorded yet</b>", ""]
    if a["BUY"]:
        lines.append("🟢 <b>BUY</b>")
        lines += [f"• {esc(x['symbol'])} — <b>{x['qty']}</b> units × {num(x['price'])} = {inr(x['value'])} "
                  f"(target {inr(x['target'])})" for x in a["BUY"]]
    if a["SELL"]:
        lines.append("🔴 <b>SELL</b>")
        lines += [f"• {esc(x['symbol'])} — all {x['qty']} units ({esc(x['reason'])}) ≈ {inr(x['value'])}"
                  for x in a["SELL"]]
    if a["HOLD"]:
        lines.append("⚪ <b>HOLD</b> (no action): " + ", ".join(esc(x["symbol"]) for x in a["HOLD"]))
    if a["SKIP"]:
        lines.append("🟡 <b>SKIP</b>: " + ", ".join(f"{esc(x['symbol'])} ({esc(x['reason'])})" for x in a["SKIP"]))
    if not any(a.values()):
        lines.append("No trades today.")
    if s.get("order_plan_missing_quotes"):
        lines.append("⚠️ No live quote for " + ", ".join(esc(x) for x in s["order_plan_missing_quotes"])
                     + " — last close used.")
    lines += ["", "Yahoo prices can lag — use your broker's live price: quantity = target ÷ price, rounded "
                  "down. This choice is saved and recorded at today's close by the 20:15 evening run."]
    return "\n".join(lines) + footer()


def msg_incomplete(s):
    miss = s.get("signal_missing", [])
    lines = ["⛔ <b>DATA INCOMPLETE — do not trade on this order plan</b>",
             f"{len(miss)} ETFs have no live price on Yahoo: " + ", ".join(esc(x) for x in miss),
             "The ranking would use yesterday's close for them and may be wrong.",
             "Re-run the workflow in a few minutes."]
    return "\n".join(lines) + footer()


def key_incomplete(s):
    return (s["as_of"], s["signal_status"], sorted(s.get("signal_missing", [])))


def msg_executed(s):
    a = s["actions"]
    month = month_name(s["last_exec_date"])
    lines = [f"✅ <b>{month} entry-day rebalance recorded</b>",
             f"Entry day {d(s['last_exec_date'])} · fills at that day's close · "
             f"decision: {esc(s.get('signal_source', ''))}", ""]
    for x in a["BUY"]:
        lines.append(f"🟢 BUY {esc(x['symbol'])} {x['qty']} × {num(x['price'])} = {inr(x['value'])}")
    for x in a["SELL"]:
        lines.append(f"🔴 SELL {esc(x['symbol'])} {x['qty']} × {num(x['price'])} = {inr(x['value'])} "
                     f"({esc(x['reason'])}, {pct(x['pnl_pct'])})")
    if a["HOLD"]:
        lines.append("⚪ HOLD: " + ", ".join(esc(x["symbol"]) for x in a["HOLD"]))
    if a["SKIP"]:
        lines.append("🟡 SKIP: " + ", ".join(f"{esc(x['symbol'])} ({esc(x['reason'])})" for x in a["SKIP"]))
    lines += ["", f"Portfolio after: {inr(s['value'])} · cash {inr(s['cash'])} · "
                  f"{s['n_positions']} positions"]
    return "\n".join(lines) + footer()


def risky_positions(s):
    return [p for p in s["positions"] if p["status"] != "HOLD"]


def msg_exit_risk(s):
    lines = ["⚠️ <b>Exit risk</b>", f"Prices as of {d(s['as_of'])}", ""]
    for p in risky_positions(s):
        why = p["status"].replace("EXIT RISK — ", "")
        lines.append(f"• <b>{esc(p['symbol'])}</b> {num(p['price'])} — {esc(why)}")
        lines.append(f"   Hard SL {num(p['hard_sl'])} ({pct(p['dist_sl'], False)} away) · "
                     f"Trail {num(p['trail_sl'])} ({pct(p['dist_trail'], False)} away)")
    lines += ["", f"Exits happen only on the entry day: {d(s['next_exec_date'])}."]
    return "\n".join(lines) + footer()


def my_stop_positions(s):
    """Your holdings whose reference stop (from your own fills) is hit or within the near-stop band."""
    a = s.get("account") or {}
    return [h for h in a.get("holdings", []) if h.get("my_stop_status") not in (None, "OK")]


def msg_my_stops(s):
    lines = ["👤 <b>Your account (reference) — stop levels from your fills</b>",
             (f"Live prices {esc(s['live_quote_time'])} IST" if s.get("live_view")
              else f"Prices as of {d(s['as_of'])}"), ""]
    for h in my_stop_positions(s):
        stale = " · ⚠️ price stale" if h.get("my_price_stale") else ""
        lines.append(f"• <b>{esc(h['symbol'])}</b> {num(h['price'])} — {esc(h['my_stop_status'])}{stale}")
        lines.append(f"   Avg fill {num(h['avg_fill'])} · My Hard SL {num(h['my_hard_sl'])} "
                     f"({pct(h['my_dist_sl'], False)} away) · My Trail {num(h['my_trail_sl'])} "
                     f"({pct(h['my_dist_trail'], False)} away)")
    lines += ["", f"Reference only — exits follow the strategy levels on the entry day "
                  f"({d(s['next_exec_date'])})."]
    return "\n".join(lines) + footer()


def data_affected(s):
    data = s["data"]
    signal_hit = data["stale_signal"] or data["stale_actions"]
    return bool(data["stale_held"] or signal_hit or data["market_delayed"])


def msg_data(s):
    data = s["data"]
    lines = ["⚠️ <b>Data delayed</b>"]
    n = sum(len(v) for v in data["stale"].values())
    if n:
        lines.append(f"{n} of {data['etfs_loaded']} ETFs have no price for {d(s['as_of'])}; "
                     f"their last close is used.")
    if data["market_delayed"]:
        lines.append(f"Latest market data is {d(s['as_of'])}, but {d(data['expected_latest'])} was expected "
                     f"(Yahoo delay or exchange holiday).")
    if data["stale_held"]:
        lines.append("Held: " + ", ".join(esc(x) for x in data["stale_held"]))
    sig = sorted(set(data["stale_signal"]) | set(data["stale_actions"]))
    if sig:
        lines.append("Ranking / actions: " + ", ".join(esc(x) for x in sig))
    if s["signal_status"] in ("ORDER PLAN", "INCOMPLETE"):
        lines.append("<b>Check before trading</b> — re-run the workflow once Yahoo has updated.")
    return "\n".join(lines) + footer()


def msg_failure(problem):
    _, run = links()
    lines = ["❌ <b>ETF strategy run failed</b>", f"Problem: <b>{esc(problem)}</b>",
             "Nothing new was published — the previous report stays online."]
    if run:
        lines.append(f'🔗 <a href="{run}">Open the run log</a>')
    return "\n".join(lines)


# ── de-duplication keys (what makes an alert "new") ──────────────────────────
def key_executed(s):
    return (s["signal_status"], s["last_exec_date"])


def key_risk(s):
    return sorted((p["symbol"], p["status"]) for p in risky_positions(s))


def key_my_stops(s):
    return sorted((h["symbol"], h["my_stop_status"]) for h in my_stop_positions(s))


def key_data(s):
    data = s["data"]
    return (s["as_of"], data_affected(s), sorted(data["stale_held"]), sorted(data["stale_signal"]),
            sorted(data["stale_actions"]), data["market_delayed"])


def build_messages(s, prev, mode):
    dedupe = mode != "manual" and prev is not None

    def new(keyfn):
        if not dedupe:
            return True
        try:
            return keyfn(s) != keyfn(prev)
        except (KeyError, TypeError):   # previous summary from an older format
            return True

    msgs = []
    act_mode = s.get("act_mode")
    if act_mode == "orderplan":
        msgs.append(msg_order_plan(s))              # manual pre-close run: always send
    elif act_mode == "incomplete" and new(key_incomplete):
        msgs.append(msg_incomplete(s))
    if s["signal_status"] == "EXECUTED" and new(key_executed):
        msgs.append(msg_executed(s))
    if risky_positions(s) and new(key_risk):
        msgs.append(msg_exit_risk(s))
    if my_stop_positions(s) and new(key_my_stops):
        msgs.append(msg_my_stops(s))
    if data_affected(s) and new(key_data):
        msgs.append(msg_data(s))
    if mode in ("evening", "manual"):
        msgs.append(msg_summary(s, prev))
    return msgs


# ── Telegram API ─────────────────────────────────────────────────────────────
def send(token, chat_id, text):
    chunks = [text[i:i + MAX_LEN] for i in range(0, len(text), MAX_LEN)]
    for chunk in chunks:
        body = urllib.parse.urlencode({"chat_id": chat_id, "text": chunk, "parse_mode": "HTML",
                                       "disable_web_page_preview": "true"}).encode()
        for attempt in range(3):
            try:
                req = urllib.request.Request(f"https://api.telegram.org/bot{token}/sendMessage", data=body)
                with urllib.request.urlopen(req, timeout=20) as r:
                    if json.load(r).get("ok"):
                        break
            except urllib.error.HTTPError as e:
                detail = e.read().decode(errors="replace")[:300]
                if e.code in (400, 401, 403, 404):          # bad token / chat id / message — no retry
                    raise RuntimeError(f"Telegram rejected the message (HTTP {e.code}): {detail}") from None
            except urllib.error.URLError as e:
                detail = str(e.reason)
            if attempt == 2:
                raise RuntimeError(f"Telegram send failed after 3 attempts: {detail}")
            time.sleep(3)


def load_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--summary", default="reports/summary.json")
    ap.add_argument("--previous", default=None, help="previous published summary.json (for de-duplication)")
    ap.add_argument("--mode", choices=["evening", "morning", "manual"], default="manual")
    ap.add_argument("--failure", metavar="PROBLEM", help="send a failure alert instead of the report messages")
    ap.add_argument("--dry-run", action="store_true", help="print messages instead of sending them")
    args = ap.parse_args()

    if args.failure:
        msgs = [msg_failure(args.failure)]
    else:
        s = load_json(args.summary)
        if s is None:
            print(f"::error::Summary file not found or invalid: {args.summary}")
            return 1
        prev = load_json(args.previous) if args.previous else None
        print(f"Mode: {args.mode} | previous summary: {'found' if prev else 'none'} | "
              f"status {s['signal_status']} | data {s['data']['status']}")
        msgs = build_messages(s, prev, args.mode)

    if not msgs:
        print("Nothing new to report — no Telegram message sent.")
        return 0
    if args.dry_run:
        for m in msgs:
            print("─" * 60)
            print(m)
        print("─" * 60)
        print(f"[dry run] {len(msgs)} message(s) not sent.")
        return 0

    token, chat_id = os.environ.get("TELEGRAM_BOT_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("::warning title=Telegram not configured::Add the TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID "
              "repository secrets to receive alerts. Skipping.")
        return 0
    for m in msgs:
        send(token, chat_id, m)
    print(f"Sent {len(msgs)} Telegram message(s).")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except RuntimeError as e:
        print(f"::error title=Telegram problem::{e}")
        sys.exit(1)
