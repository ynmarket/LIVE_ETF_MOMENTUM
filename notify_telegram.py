"""
notify_telegram.py — Telegram alerts for the ETF momentum live dashboard.

Reads reports/summary.json (written by ETF_Momentum_Test_1.py) and sends:
  🔔 final / provisional rebalance signal   (signal day, status READY)
  ✅ rebalance executed                      (execution day, status EXECUTED)
  ⚠️ exit risk                               (stop hit, near a stop, or outside the top 6)
  ⚠️ data delayed                            (held / top-6 / action ETFs without the latest price)
  📊 daily summary                           (evening and manual runs)
  ❌ run failed                              (--failure)

Alert messages are de-duplicated against the previously published summary, so the
same alert is not repeated on every run (manual runs always send everything).
Standard library only. Never prints the bot token.

Usage:
  python notify_telegram.py --summary reports/summary.json --previous reports/previous_summary.json --mode evening
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
             f"Data as of {d(s['as_of'])} · generated {esc(s['generated_ist'])} IST", ""]
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
        lines += [f"⏳ Live portfolio starts <b>{esc(s['live_start'])}</b> with {inr(s['initial_capital'])}.",
                  f"Preview of initial buys (latest close): {buys}"]
    exp = lambda known: "" if known else " (expected)"
    lines += ["", f"🗓 Next: signal {d(s['next_signal_date'])}{exp(s['next_signal_known'])} → "
                  f"execute {d(s['next_exec_date'])}{exp(s['next_exec_confirmed'])} · <b>{s['signal_status']}</b>"]
    if s["signal_status"] == "PENDING" and s["next_signal_date"] == s["today_ist"]:
        lines.append("ℹ️ Today is signal day, but today's close is not in the data yet — "
                     "the final signal will be checked again at 08:15 IST.")
    data = s["data"]
    if data["status"] == "CURRENT":
        lines.append(f"📡 Data: ✓ current ({data['etfs_loaded']}/{data['universe']} ETFs)")
    else:
        n = sum(len(v) for v in data["stale"].values())
        lines.append(f"📡 Data: ⚠️ delayed ({n} ETFs without a {s['as_of']} price)")
    if prev and prev.get("as_of") == s["as_of"] and prev.get("generated_ist", "")[:10] != s["generated_ist"][:10]:
        lines.append("ℹ️ No new market data since the last run (holiday, or Yahoo not updated yet).")
    return "\n".join(lines) + footer()


def _signal_provisional(s):
    data = s["data"]
    return bool(data["stale_signal"] or data["stale_actions"] or data["market_delayed"])


def msg_signal(s):
    a = s["actions"]
    prov = _signal_provisional(s)
    month = month_name(s["next_exec_date"])
    head = (f"⏳ <b>PROVISIONAL {month} signal — data incomplete</b>" if prov
            else f"🔔 <b>FINAL {month} signal</b>")
    lines = [head,
             f"Signal close: {d(s['as_of'])} → execute on <b>{d(s['next_exec_date'])}</b> at close", ""]
    if a["BUY"]:
        lines.append("🟢 <b>BUY</b>")
        lines += [f"• {esc(x['symbol'])} #{x['rank']} — ~{x['qty']} units, target {inr(x['target'])} "
                  f"(est. @ {num(x['price'])})" for x in a["BUY"]]
    if a["SELL"]:
        lines.append("🔴 <b>SELL</b>")
        lines += [f"• {esc(x['symbol'])} — {esc(x['reason'])} · {x['qty']} units · P&amp;L {pct(x['pnl_pct'])}"
                  for x in a["SELL"]]
    if a["HOLD"]:
        lines.append("⚪ <b>HOLD</b>: " + ", ".join(esc(x["symbol"]) for x in a["HOLD"]))
    if a["SKIP"]:
        lines.append("🟡 <b>SKIP</b>: " + ", ".join(f"{esc(x['symbol'])} #{x['rank']} ({esc(x['reason'])})"
                                                   for x in a["SKIP"]))
    if not any(a.values()):
        lines.append("No trades.")
    lines += ["", "Quantities use the signal-day close as an estimate; the strategy fills at the "
                  "execution-day close."]
    if prov:
        lines.append("⚠️ Some prices are delayed — this signal may still change. It will be re-checked at "
                     "08:15 IST, or re-run the workflow from GitHub Actions.")
    return "\n".join(lines) + footer()


def msg_executed(s):
    a = s["actions"]
    month = month_name(s["last_exec_date"])
    lines = [f"✅ <b>{month} rebalance executed</b>",
             f"Executed {d(s['last_exec_date'])} · signal {d(s['last_signal_date'])}", ""]
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
    lines += ["", f"Exits happen only at the monthly rebalance: signal {d(s['next_signal_date'])} → "
                  f"execute {d(s['next_exec_date'])}."]
    return "\n".join(lines) + footer()


def data_affected(s):
    data = s["data"]
    return bool(data["stale_held"] or data["stale_signal"] or data["stale_actions"] or data["market_delayed"])


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
        lines.append("Signal / actions: " + ", ".join(esc(x) for x in sig))
    if s["signal_status"] == "READY":
        lines.append("<b>Do not act on the signal yet</b> — it will be re-checked at 08:15 IST.")
    return "\n".join(lines) + footer()


def msg_failure(problem):
    _, run = links()
    lines = ["❌ <b>ETF strategy run failed</b>", f"Problem: <b>{esc(problem)}</b>",
             "Nothing new was published — the previous report stays online."]
    if run:
        lines.append(f'🔗 <a href="{run}">Open the run log</a>')
    return "\n".join(lines)


# ── de-duplication keys (what makes an alert "new") ──────────────────────────
def key_signal(s):
    a = s["actions"]
    return (s["signal_status"], s["next_exec_date"], _signal_provisional(s),
            sorted((x["symbol"], x["qty"]) for x in a["BUY"]),
            sorted((x["symbol"], x["reason"]) for x in a["SELL"]),
            sorted(x["symbol"] for x in a["SKIP"]))


def key_executed(s):
    return (s["signal_status"], s["last_exec_date"])


def key_risk(s):
    return sorted((p["symbol"], p["status"]) for p in risky_positions(s))


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
    if s["signal_status"] == "READY" and new(key_signal):
        msgs.append(msg_signal(s))
    if s["signal_status"] == "EXECUTED" and new(key_executed):
        msgs.append(msg_executed(s))
    if risky_positions(s) and new(key_risk):
        msgs.append(msg_exit_risk(s))
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
