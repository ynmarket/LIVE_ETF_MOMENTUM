"""
entry_day_reminder.py — Telegram reminder on the ENTRY DAY (first trading day of the month).

Run by .github/workflows/entry_day_reminder.yml at 14:00 IST on days 1–7 of every month. It only
sends a reminder to run the "Run ETF Strategy" workflow for the order plan — it never runs the
strategy, saves anything or changes the report.

Entry day = first trading day of the month, taken from real NSE trading data (NIFTYBEES on Yahoo),
so weekends and exchange holidays are handled: today must have a trading bar with volume and no
earlier day this month may have one (Yahoo's zero-volume holiday placeholder bars are ignored).

Usage:
  python entry_day_reminder.py                      # check today (IST), send if entry day
  python entry_day_reminder.py --dry-run            # print instead of sending
  python entry_day_reminder.py --today 2026-11-02   # check another date (testing)
  python entry_day_reminder.py --force              # send a test reminder regardless of the date
Env: TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID (GitHub secrets).
"""
import argparse
import os
import sys
from datetime import date, datetime, timedelta, timezone

from notify_telegram import esc, send

IST = timezone(timedelta(hours=5, minutes=30))
CALENDAR_SYMBOL = "NIFTYBEES.NS"           # liquid ETF: a bar on every NSE trading day


def trading_days_this_month(day):
    """NSE trading dates from the 1st of `day`'s month up to and including `day` (Yahoo daily bars;
    during market hours Yahoo already returns today's bar)."""
    import yfinance as yf
    start = day.replace(day=1)
    df = yf.download(CALENDAR_SYMBOL, start=start.isoformat(), end=(day + timedelta(days=1)).isoformat(),
                     progress=False, threads=False)
    if df is None or df.empty:
        return []
    vol = df["Volume"]
    if hasattr(vol, "columns"):                  # yfinance may return a one-column frame
        vol = vol.iloc[:, 0]
    # Yahoo inserts zero-volume placeholder bars on exchange holidays (e.g. 2026-05-01):
    # NIFTYBEES trades millions of units on every real session, so volume 0 = market closed.
    return sorted({ts.date() for ts, v in vol.items() if start <= ts.date() <= day and v > 0})


def check_entry_day(day):
    """(is_entry_day, reason)."""
    if day.weekday() >= 5:
        return False, "weekend"
    days = trading_days_this_month(day)
    earlier = [d for d in days if d < day]
    if earlier:
        return False, f"first trading day of the month was {earlier[0]}"
    if day in days:
        return True, "today is the first trading day of the month"
    return False, "no trading data for today (exchange holiday, or Yahoo not updated)"


def reminder_text(day, test=False):
    repo = os.environ.get("GITHUB_REPOSITORY", "ynmarket/LIVE_ETF_MOMENTUM")
    owner, name = repo.split("/", 1)
    actions = f"https://github.com/{repo}/actions/workflows/run_strategy.yml"
    page = f"https://{owner.lower()}.github.io/{name}/"
    return "\n".join([
        ("🧪 <b>TEST — </b>" if test else "") + f"🔔 <b>Entry day today — {day:%a %d %b %Y}</b>",
        "",
        "Run the workflow now for the 🛒 order plan (signal and quantities from live prices):",
        "GitHub → Actions → <b>Run ETF Strategy</b> → <b>Run workflow</b> (keep Telegram ticked).",
        "If you don't, the automatic ~14:30 run sends the order plan; it never replaces a choice "
        "already saved by your own run today.",
        "",
        "After placing the orders, add your fills (price, quantity, charges) to the Google Sheet.",
        f'🔗 <a href="{actions}">Open the workflow</a> · <a href="{page}">Open report</a>',
    ])


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--today", help="check this date (YYYY-MM-DD) instead of today in IST")
    ap.add_argument("--force", action="store_true", help="send a test reminder regardless of the date")
    ap.add_argument("--dry-run", action="store_true", help="print the message instead of sending it")
    args = ap.parse_args()

    day = date.fromisoformat(args.today) if args.today else datetime.now(IST).date()
    if args.force:
        is_entry, reason = True, "forced test reminder"
    else:
        is_entry, reason = check_entry_day(day)
    print(f"{day} ({day:%a}): entry day = {is_entry} — {reason}")
    if not is_entry:
        return 0

    text = reminder_text(day, test=args.force)
    if args.dry_run:
        print("─" * 60)
        print(text)
        print("─" * 60)
        print("[dry run] reminder not sent.")
        return 0
    token, chat_id = os.environ.get("TELEGRAM_BOT_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("::warning title=Telegram not configured::TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID missing — "
              "reminder not sent.")
        return 0
    send(token, chat_id, text)
    print("Entry-day reminder sent.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except RuntimeError as e:
        print(f"::error title=Telegram problem::{esc(str(e))}")
        sys.exit(1)
