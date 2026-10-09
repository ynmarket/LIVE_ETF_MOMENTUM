# CLAUDE.md — ETF Momentum Rotation (live portfolio system)

This file is loaded automatically in every Claude Code session opened in this folder, on any laptop.
The repository is **public**: never put secrets or the user's real account figures in this file.

## First: new session? Read the handover

**If you do not already know this project (new session, no prior context), read `handover.md` in this
folder completely before doing anything else.** It explains the whole system: goal, strategy, every
file, timing model, Google Sheet "My Account", GitHub workflows, privacy design, current state, change
history, what did not work, and next steps. Then `git pull --rebase` to get the latest state.

## What this project is (one paragraph)

A monthly **ETF momentum rotation strategy on NSE** traded with **real money** (₹3,00,000 from
2026-10-01). `ETF_Momentum_Test_1.py` downloads Yahoo Finance data for a fixed 54-ETF universe, runs a
historical backtest (reference) and a live strategy record, gives an **entry-day order plan from live
prices**, compares the user's **real trades from a private Google Sheet** ("My Account"), and writes an
HTML dashboard, Excel and `summary.json`. GitHub Actions ("Run ETF Strategy", 20:15 IST and ~14:30 IST midday Mon–Fri +
manual) publishes a **password-protected** dashboard on GitHub Pages and sends Telegram alerts; a
separate "Entry Day Reminder" workflow pings the user at 14:00 IST on the first trading day of the month.
Repo: https://github.com/ynmarket/LIVE_ETF_MOMENTUM · Page: https://ynmarket.github.io/LIVE_ETF_MOMENTUM/

## Working rules — MUST follow (set by the user)

1. **Never change the strategy rules** unless the user explicitly asks: universe (54 ETFs), momentum
   weights 15/40/30/15 % on 21/63/126/252 trading days, top 6, 200-DMA entry gate (no substitution),
   15 % hard stop, 20 % trailing stop, rotation, sizing (total value / 6, whole units, no resizing),
   5 % cash interest, and the **entry-day timing** (first trading day of the month = signal AND execution
   day at that day's close; order plan from live prices). The user trades this with real money.
   For any other work (deployment, reports, data handling, styling), change only I/O, paths, styling and
   error handling — and **prove** nothing in the strategy changed by running the old and new code on the
   same download and comparing the Excel sheets (e.g. all `Hist_*` sheets identical).
2. **Never `git push`, and never trigger a GitHub workflow run that changes the live page, without the
   user's explicit approval.** Committing locally is fine. Flow: implement → test locally → commit locally
   → show the results → ask "OK to push?" → push only after "yes".
3. **If a request can be read in more than one way, restate your understanding and ask before
   implementing.** (Once "remove this schedule day scenario" was misread as "remove the 20:15 schedule";
   it actually meant "remove the previous-day signal". The wrong change was pushed and had to be reverted.)
   When asked to "inspect first" or "suggest", present the plan and wait for confirmation.
4. **Ask the user for personal details** such as the Google Sheet ID. **Do not** search their Google
   Drive or other connected accounts to find them.
5. **Never ask for, print or commit secrets** (Telegram token, page password, Google key contents, Sheet
   ID in tracked files). Read key files only for non-secret fields (e.g. `client_email`).
6. **Keep the user's real account data private.** It must never appear unencrypted on GitHub: not in
   tracked files (incl. `CLAUDE.md`, `handover.md`), not in workflow artifacts, not in Actions logs (CI
   prints only row counts), not on the page without `PAGE_PASSWORD` (public variant has no My Account).
7. **Keep the 20:15 IST evening schedule** (and its keep-alive) and the ~14:30 IST midday update (added at the
   user's request 2026-10-08; it never replaces an entry-day choice already saved that day). The 08:15 morning run was removed on
   purpose (Yahoo withdraws the previous day's prices overnight) — do not re-add it.
8. **Universe changes only exactly as instructed** (a swap was once reverted at the user's request).
9. Investment decisions are the user's; the system only reports what the strategy rules produce.
10. When the system changes, **update `handover.md`** (and this file if a rule changes) in the same commit.
11. **Do not use Claude's local memory** for this project. Anything important (decisions, preferences,
    research results) goes into `CLAUDE.md` or `handover.md`, so every laptop gets it with `git clone`.
    (Never put secrets or the user's real account figures there — the repository is public.)

## Run & test

```bash
python ETF_Momentum_Test_1.py --force
```
Outputs go to `reports/` (git-ignored): full HTML, public HTML (no My Account), Excel (22 sheets),
`summary.json`, `summary_public.json`. Without `--force` a run is skipped if both outputs were already
generated today.

* Local Google Sheet access: `local_config.json` (git-ignored) with `{"gsheet_id": "..."}` (ask the user
  for the ID) and the service-account key `.json` in `<project>/keys/` (git-ignored) or a `keys/` folder
  next to the project. No machine-specific paths.
* Telegram dry run: `python notify_telegram.py --summary reports/summary.json --mode manual --dry-run`
* Reminder dry run: `python entry_day_reminder.py --today 2026-11-02 --dry-run`
* Scenario tests: copy the script to a scratch folder and patch lines (e.g. force `MARKET_HOURS`, fix
  `TODAY` / `_now` / `data_end_ts`, or stub `read_trade_sheet()`); run with `CI=true` for CI-only paths.
* Market hours (09:00–16:00 IST): today's Yahoo bar is a **live quote, never a close**; current values use
  live prices (labelled "live HH:MM"), history stays on closes. Do not break this.

## Git / GitHub

* **Always `git pull --rebase` before pushing** — the workflow commits `live_state/entry_decisions.json`
  to `main`.
* Commit messages end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
* On the user's Windows laptop: `gh` CLI may be at `C:\Program Files\GitHub CLI\gh.exe` (account
  `ynmarket`); push without changing global git config:
  `git -c credential.helper= -c "credential.https://github.com.helper=!'C:/Program Files/GitHub CLI/gh.exe' auth git-credential" push origin main`
* GitHub secrets (names only): `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID`, `GSHEET_ID`,
  `GOOGLE_SERVICE_ACCOUNT_JSON`, `PAGE_PASSWORD`.
* Python 3.14; packages pinned in `requirements.txt`.

## Files

| File | Role |
|---|---|
| `ETF_Momentum_Test_1.py` | Everything: data, backtest, live record, order plan, My Account, Excel, HTML, summary |
| `notify_telegram.py` | Telegram messages from `summary.json` |
| `entry_day_reminder.py` | 14:00 IST entry-day reminder |
| `page_crypto.py` | Encrypts the published page/summary (AES-256-GCM, PBKDF2) |
| `.github/workflows/run_strategy.yml` | Main workflow (20:15 IST + ~14:30 IST midday + manual) |
| `.github/workflows/entry_day_reminder.yml` | Reminder workflow |
| `live_state/entry_decisions.json` | Saved entry-day ETF choices (written by CI, committed by the workflow) |
| `handover.md` | Full system documentation — read it first in a new session |
