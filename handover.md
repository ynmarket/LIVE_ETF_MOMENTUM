# HANDOVER — ETF Momentum Rotation: Live Portfolio System

> Written 2026-10-06 (IST) for a brand-new session with **zero prior context**. Read §1–§3 first, then use
> the rest as reference. Everything described here is implemented, tested and **pushed** (GitHub
> `main` = local `main` = commit `1be7362`) unless explicitly marked otherwise.

---

## 1. What this is (core goal)

The user (GitHub account **`ynmarket`**, Windows laptop, trades in India/NSE) runs a **monthly ETF
momentum rotation strategy with real money (₹3,00,000 starting 2026-10-01)**. This project:

1. **Computes the strategy** (ranking of a fixed universe of 54 NSE ETFs) and a **historical backtest**.
2. **Maintains a "live portfolio" strategy record** from `LIVE_START_DATE` (2026-10-01).
3. On the **entry day** (first trading day of each month) gives an **order plan from live prices** so the
   user can place orders (around **14:00 IST**) before the close.
4. Tracks the user's **real account** (fills typed into a private **Google Sheet**) against the strategy.
5. Publishes a **password-protected mobile dashboard** on GitHub Pages and sends **Telegram** alerts.
6. Runs automatically on **GitHub Actions** (laptop can be off), and can be run manually from a phone.

Live dashboard: **https://ynmarket.github.io/LIVE_ETF_MOMENTUM/** (password-protected).
Repository: **https://github.com/ynmarket/LIVE_ETF_MOMENTUM** (public).
Local folder: **`F:\1.AI_Market\TEST_LIVE_ETF_MOMENTUM`** (repo name on GitHub is LIVE_ETF_MOMENTUM; the
local folder name was intentionally left as-is).

---

## 2. Working rules with this user — FOLLOW STRICTLY

These came from explicit user feedback (some after mistakes). They override convenience.

1. **Never change the strategy rules** (universe, weights, lookbacks, top 6, 200-DMA gate, stops,
   rotation, sizing, cash, entry-day timing) unless the user explicitly asks. Deployment/report/data-
   handling changes only. When in doubt, prove "no strategy change" by running old vs new code on the
   **same download** and comparing Excel sheets (see §12).
2. **Never `git push` (or trigger a GitHub run that changes the live page) without explicit approval.**
   Workflow: implement → test **locally** → commit **locally** → show results → ask "OK to push?" → push.
3. **If a request has more than one reading, restate it and ask before implementing.** (A past mistake:
   "remove this schedule day scenario" meant *remove the previous-day signal*, not *remove the 20:15
   schedule*. The wrong change was pushed and then reverted.)
4. **Ask the user for IDs/details** (e.g. Google Sheet ID) — do **not** search their Google Drive or other
   connectors.
5. **Never ask for or print secrets** (bot token, page password, key contents). Read key files only for
   non-secret fields (e.g. `client_email`).
6. Communicate in plain language; the user reviews local results before pushing.
7. Investment decisions are the user's; the system reports what the strategy rules produce.

---

## 3. System at a glance

```
                        Yahoo Finance (yfinance)
                                 │  daily OHLCV for 54 ETFs + ^CRSLDX (Nifty 500)
                                 ▼
  ┌───────────────────── ETF_Momentum_Test_1.py (single script) ─────────────────────┐
  │ data cleaning (holiday placeholders, market-hours live bar)                        │
  │ historical backtest 2021-01 → END_DATE  (reference only)                           │
  │ live portfolio record from LIVE_START_DATE (replays saved entry decisions)         │
  │ entry-day order plan (live prices) · live view (live prices for current values)    │
  │ My Account (Google Sheet, read-only) · XIRR · vs Nifty · execution cost            │
  │ outputs: reports/*.html, reports/*.xlsx, reports/summary*.json                     │
  └────────────────────────────────────────────────────────────────────────────────────┘
          │ GitHub Actions "Run ETF Strategy"                         ▲ live_state/entry_decisions.json
          ▼                                                           │ (committed by the workflow)
   page_crypto.py → encrypted page + summary → GitHub Pages ───────────┘
   notify_telegram.py → Telegram (summary, order plan, alerts, failures)
   entry_day_reminder.py ("Entry Day Reminder" workflow, 14:00 IST on the entry day) → Telegram
```

Monthly user routine:
1. **Entry day (1st trading day), ~14:00 IST**: Telegram reminder arrives → user runs
   *Actions → Run ETF Strategy → Run workflow* (Telegram ticked) → receives **🛒 ORDER PLAN** (live
   prices) → places orders before 15:30.
2. Same evening **20:15 IST** scheduled run records the trades at the **close** using the saved 2 PM choice.
3. User types real fills (date, symbol, side, qty, price, charges) into the Google Sheet → next run shows
   **My Account** and plan-vs-actual.
4. Every weekday **~14:30 IST** (scheduled 14:25, GitHub often starts late): 🕑 **Midday Update** on Telegram
   with live prices — strategy positions with Hard SL / Trail SL + distances, My Account holdings with the
   reference levels from the user's fills; alerts (exit risk, 👤 your account, data) only when new.
   On the entry day this run also sends the 🛒 order plan; it **never replaces an entry-day choice already
   saved that day** (env `RUN_KIND=midday` → `kept_decision`; the order plan then shows the saved choice
   with live quantities). If no choice was saved yet, it saves its own (as a manual run would).
5. Every weekday 20:15 IST: daily summary on Telegram (same per-position SL detail, closing prices).

---

## 4. The strategy (DO NOT CHANGE)

| Item | Value |
|---|---|
| Universe | 54 NSE ETFs (list `UNIVERSE` in the script; `C54`) — ETFs with < 262 rows are excluded automatically until they have enough history |
| Holdings | Top **6** (`N_HOLD = 6`) |
| Momentum score | `ret_1m*0.15 + ret_3m*0.40 + ret_6m*0.30 + ret_12m*0.15`; lookbacks **21/63/126/252 trading days** (positional) |
| Min history | 262 rows |
| Entry gate | price > **200-DMA** (mean of last 200 closes); top-6 ETF below DMA → **slot stays empty, no substitution** |
| Exits (priority) | 1) Hard SL: price ≤ entry×0.85 · 2) Trailing: price ≤ peak×0.80 (peak ratchets on rebalance dates) · 3) Rotation: not in top 6 |
| Sizing | new entry = **total portfolio value / 6**, whole units (`floor`), capped by available cash; existing positions are **not** resized |
| Cash | earns 5% p.a. (`LIQUID_RATE`), accrued between rebalances |
| Timing | **Entry day = first trading day of the month = signal day AND execution day** (signal and fills at that day's close; live: order plan from live prices) |
| Benchmark | Nifty 500 (`^CRSLDX`), comparison only |
| Other constants | `START_DATE 2021-01-01`, `END_DATE 2026-09-30` (end of the historical *reference* backtest; live data always runs to today), `INITIAL_CAPITAL 300000`, `LIVE_START_DATE 2026-10-01`, `RF_ANNUAL 0.06` (Sharpe) |

Categories are display labels only (MONQ50 = International, it is the Motilal Oswal **Nasdaq** Q 50 ETF).

---

## 5. Files

### Tracked in git (pushed)
| File | Lines | Role |
|---|---|---|
| `ETF_Momentum_Test_1.py` | ~2400 | **Everything**: download, cleaning, backtest, live engine, dashboard state, My Account, Excel, HTML, summary.json. Run: `python ETF_Momentum_Test_1.py [--force]` |
| `notify_telegram.py` | ~330 | Builds/sends Telegram messages from `reports/summary.json` (stdlib only). `--mode evening|midday|manual`, `--previous` for de-duplication, `--failure "<problem>"`, `--dry-run` |
| `entry_day_reminder.py` | ~100 | Detects entry day from NIFTYBEES daily bars **with volume > 0**; sends reminder. `--today YYYY-MM-DD`, `--force`, `--dry-run` |
| `page_crypto.py` | ~125 | AES-256-GCM + PBKDF2-SHA256 (600k) encryption. `page <html> <out>` builds the unlock page; `enc`/`dec` for JSON. Password from `CRYPTO_PASSWORD` or `PAGE_PASSWORD` env |
| `.github/workflows/run_strategy.yml` | ~260 | Main workflow "Run ETF Strategy" (§9) |
| `.github/workflows/entry_day_reminder.yml` | ~57 | "Entry Day Reminder" workflow |
| `live_state/entry_decisions.json` | — | **Saved entry-day decisions** (which ETFs to buy/sell/hold/skip per month). Written by the script **only in CI** and committed by the workflow. Currently: `2026-10` → MONQ50, MAFANG, MON100, PHARMABEES, METALIETF, MODEFENCE (decided from the Oct 1 close) |
| `handover.md` | — | This document. Public: keep secrets and the user's real account figures out of it |
| `requirements.txt` | — | Pinned: yfinance 1.7.0, pandas 3.0.6, numpy 2.5.3, openpyxl 3.1.5, plotly 7.1.0, google-auth 2.59.1, requests 2.34.2, cryptography 50.0.2 |
| `.gitignore` | — | Ignores `reports/`, `public/`, `ETF_Momentum_Output/`, `local_config.json`, `*service_account*.json`, caches |

### Local only (NOT in git)
| Path | Role |
|---|---|
| `local_config.json` | `{"gsheet_id": "..."}` (optionally `"service_account_file"`, absolute or relative to the project) — lets local runs read the Google Sheet. Per machine, **never commit** (git-ignored). |
| `keys/*.json` or `../keys/*.json` | Google service-account key (a password). Found automatically in `<project>/keys/` (git-ignored) or a `keys/` folder next to the project (this laptop: `F:\1.AI_Market\keys\`). Never print/commit. |
| `reports/` | Local outputs: `ETF_Momentum_Test_1_Report.html` (full), `..._Report_public.html` (no My Account), `ETF_Momentum_Test_1_Report.xlsx`, `summary.json`, `summary_public.json` |
| `ETF_Momentum_Output/` | Old output folder from before the move to `reports/` (ignored; can be deleted by the user) |

---

## 6. ETF_Momentum_Test_1.py — internal map (top to bottom)

1. **Auto-install + settings** (`_ensure`, `START_DATE`… `LIVE_START_DATE`), `IST` timezone helpers
   (`now_ist()`), `IN_CI` (env `CI=true` on GitHub), `open_report()` (no browser in CI).
2. **Output paths / skip logic**: `reports/`; if both outputs were generated **today** and no `--force`
   → skip and open the report.
3. **Universe + parameters** (`UNIVERSE`, `CATEGORY`, `N_HOLD`, `LB`, `WEIGHTS`, …).
4. **Download** `fetch_close()` (Close + optional Volume). **Market-hours guard**: 09:00–16:00 IST on
   weekdays, today's Yahoo bar is a *live quote* → stored in `live_quotes` and **removed from history**
   (same for the benchmark → `bm_live`).
5. **Holiday placeholder removal**: Yahoo inserts zero-volume bars on NSE holidays (e.g. 2026-01-15,
   05-01, 05-28, 06-26, 09-14, 10-02). A date counts as a session only if **any** ETF had volume > 0;
   other dates are removed from every series. Then MIN_HISTORY check; `MIN_ETFS_LOADED = 40` guard
   (exit code 2).
6. **Calendar**: `cal` (union of dates, forward-filled `etf_close`), `month_first_days()`,
   `rebalance_schedule()` → `[(entry_day, entry_day)]`, `missing_on()`.
7. **Saved entry decisions**: `load_entries()`, `entry_decisions()` (ignored if `live_start` /
   `initial_capital` in the file don't match the settings).
8. **Signals**: `compute_scores(d, extra=None)` (`extra` = live prices appended as today's value).
9. **Engine**: `new_state()`, `rebalance(st, signal_date, execution_date, …, decisions=None, live=None)`
   — Phase A signal (data ≤ entry day), Phase B execution; `assert sd == ed` for real rebalances;
   `run_strategy(schedule, capital, tag, entries=None)`; `daily_equity()`; `open_positions_df()`.
10. **Historical backtest** `bt` and **live record** `live` (replays `ENTRIES`).
11. **Look-ahead validation** `validate_no_lookahead()` — runs every time, raises on failure (signal =
    execution day; scores recomputed from truncated data; fills = entry-day close).
12. **Historical reference metrics** (CAGR, Sharpe, …) — only shown in a collapsed section / Hist_ sheets.
13. **Live dashboard state**: `AS_OF` (latest close), timing (`next_rebal`, `entry_today`), **order plan**
    (`order_plan = MARKET_HOURS and entry_today`), **LIVE_VIEW** (`MARKET_HOURS and live_quotes`),
    `cur_px()` / `cur_px_label()`, `live_sc`, `cur_sc/cur_ranked/cur_rank`, positions (`pos_rows`),
    status (`EXECUTED / INCOMPLETE / ORDER PLAN / ENTRY DAY / PENDING`), actions (`act_mode`
    `executed / incomplete / orderplan / preview`), data-incomplete check (`QUOTE_GAP_TOLERANCE`).
14. **Save entry-day decisions** (CI only): during an entry-day order plan (last market-hours run wins),
    or from the entry-day close if no entry-day run happened (only when that day's closes are complete).
15. **Data freshness** (`stale`, `market_delayed`, …), signals table rows, live performance, console summary.
16. **Ranking trend** (`rank_trend`, `RANK_TREND_DAYS = 6`).
17. **My Account** (§8): `read_trade_sheet()`, `parse_trades()`, `parse_cash()`, `_xirr()`,
    `_account_daily()`, `build_account()`, `perf_live`.
18. **Excel** (`_write_excel()`, exit 3 on failure).
19. **HTML** (Plotly JS embedded once in `<head>` as `PLOTLY_JS`; `fig_to_div()` never includes the
    library), sections, then writes full HTML (exit 4) and the **public variant** without My Account.
20. **summary.json** + **summary_public.json** (exit 5).

Exit codes: 2 data, 3 Excel, 4 HTML, 5 summary (mapped to messages by the workflow).

---

## 7. Timing model (important, was changed several times)

* **Entry day** = first trading day of the month (from real data). Signal and execution are the **same
  day**. History (backtest + live record) uses that day's **close**.
* **During market hours (09:00–16:00 IST)**: today's Yahoo bar is never treated as a close.
  * **Entry day** → 🛒 **ORDER PLAN**: ranking/DMA/stops/quantities from **live prices**; the chosen ETFs
    are saved to `live_state/entry_decisions.json` (CI); the evening run records them at the close.
    If live quotes are missing for held/top-10 ETFs or > 2 ETFs → **⛔ DATA INCOMPLETE — do not trade**.
  * **Any other day** → **LIVE VIEW**: all *current* values (positions, P&L, stops, signals, ranking
    "now", actions preview, My Account, Telegram) use live prices, labelled `live HH:MM`; history charts
    stay on closes plus a ★ live point.
* **After 16:00 IST**: today's bar is treated as the close (e.g. the 20:15 run records the entry day).
* Outside market hours the output is exactly as before the live-view feature (verified).

---

## 8. My Account (real trades from Google Sheet)

* Sheet **"ETF Momentum – My Trades"** (owned by the user), tabs **`Trades`**, **`Cash`**, **`Read me`**.
  Shared as **Viewer** with service account `etf-sheet-reader@etf-momentum.iam.gserviceaccount.com`
  (Google Cloud project `etf-momentum`, Sheets API enabled).
* `Trades` row 1 headers (exact): `Trade Date | Symbol | Side | Quantity | Price | Charges ₹ | Entry Month
  | Order ID | Notes | Value ₹ (formula) | Net Cash ₹ (formula)`. One row per fill; Side BUY/SELL;
  Entry Month optional (defaults to the trade date's month).
* `Cash` headers: `Date | Type | Amount ₹ | Notes`; Type = Deposit / Withdrawal / Dividend / Interest.
  Initial capital is recorded there as a Deposit.
* The user's real fills and account figures are **private**: read them from the Google Sheet (local run
  with `local_config.json`), the local report/Excel, or the password-protected page — never write them
  into tracked files (this repository is public).
* Reading: Sheets API v4 `values:batchGet`, `UNFORMATTED_VALUE` + serial dates. A read failure never
  stops the run. Credentials (`_account_credentials` / `_find_key_file`, no machine-specific paths):
  Sheet ID = env `GSHEET_ID` or `local_config.json` `gsheet_id`; key = env `GOOGLE_SERVICE_ACCOUNT_JSON`
  (CI) → env `GOOGLE_SERVICE_ACCOUNT_FILE` / `GOOGLE_APPLICATION_CREDENTIALS` → `local_config.json`
  `service_account_file` (absolute or project-relative; skipped with a warning if missing) → first
  service-account `*.json` in `<project>/keys/` or `<project>/../keys/`.
* **New laptop setup**: clone the repo, copy the key `.json` into `<repo>/keys/`, create
  `local_config.json` with `{"gsheet_id": "<ask the user>"}`, `pip install -r requirements.txt`, run.
* Validation: bad rows (unknown symbol, bad side, non-positive qty/price, bad date, selling more than held)
  are listed and **not used**.
* Calculations: average-cost holdings incl. charges; realised/unrealised P&L after charges; cash; account
  value; return vs capital (deposits − withdrawals, or assumed 300000 if no deposits); **XIRR** (shown
  after 30 days); **time-weighted index** vs strategy vs Nifty 500 (deposits are not performance);
  **plan vs actual** per entry month; **execution cost** = (your price − strategy close) × qty (+charges),
  negative = saved.
* **Reference stop levels from the user's fills** (table "My stop levels from my fills", Excel
  `My_Holdings` "(ref)" columns, summary `account.holdings[].my_*`): **information only — the user exits on
  the strategy levels.** My Hard SL = average fill price **excl. charges** × 0.85; My Peak starts at that
  average fill and is raised **only by entry-day closes after the first buy date** of the holding (rule A,
  no daily closes, no today's price, the buy-day close itself does not count); My Trailing SL = My Peak ×
  0.80; current price = live during market hours, else latest close (flagged "price stale" if delayed).
  Telegram sends **👤 Your account (reference)** when a reference level is hit or within 5 % (de-duplicated
  in evening runs).
* **Strategy stop display** (Current Live Positions, Risk Monitor, Telegram ⚠️, `Live_Positions`): Peak =
  the **stored engine peak** (highest entry-day close since entry) — exactly the level the engine checks on
  the next entry day. Until 2026-10-08 the page used max(stored peak, today's price), which showed a
  misleading trail (today × 0.80) for winners; verified on real backtest trades (all 101 recorded peaks =
  highest entry-day close; SILVERBEES Feb 2026 example). Statuses were never affected. `Hist_Open_Positions`
  (historical reference) still uses the old display formula — left unchanged on purpose.
* Decided with the user (2026-10-08): a research run (scratch only, not committed) compared peak/stop
  methods on the backtest — A current 31.1 % CAGR / −24.8 % DD; B daily peak 31.6 / −24.6; C daily stop
  30.8 / −20.7; C2 daily stop next-day 29.5 / −21.0. User keeps rule A; strategy unchanged.

---

## 9. GitHub setup

* **Auth/push from this laptop**: `gh` CLI at `C:\Program Files\GitHub CLI\gh.exe` (may not be on PATH
  in old shells), logged in as `ynmarket`. Repo-local git identity `ynmarket` /
  `335328897+ynmarket@users.noreply.github.com`. Push **without** changing global config:
  ```powershell
  $cred = @("-c","credential.helper=","-c","credential.https://github.com.helper=!'C:/Program Files/GitHub CLI/gh.exe' auth git-credential")
  git @cred pull --rebase -q origin main      # ALWAYS first: the workflow commits live_state/ to main
  git @cred push origin main
  ```
  Commit messages end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`. In PowerShell,
  messages containing quotes break `git commit -m` → write the message to a file and use `-F`.
* **Secrets** (Settings → Secrets and variables → Actions): `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID`,
  `GSHEET_ID`, `GOOGLE_SERVICE_ACCOUNT_JSON`, `PAGE_PASSWORD`. Never ask for their values.
* **Workflow "Run ETF Strategy"** (`run_strategy.yml`):
  * Triggers: manual (`workflow_dispatch`, input `telegram` default true) + **schedule
    `45 14 * * 1-5` = 20:15 IST Mon–Fri** (user wants this kept) + **`55 8 * * 1-5` = ~14:30 IST Mon–Fri
    midday update** (added 2026-10-08 at the user's request; sets `RUN_KIND=midday`, Telegram `--mode midday`). Concurrency group `run-etf-strategy`.
  * Job `run-strategy` (contents: write): checkout → Python 3.14 → `pip install -r requirements.txt` →
    run script `--force` (exit-code → error titles) → **Verify reports** (incl. **privacy check**: public
    page/summary must not contain My Account) → **Save entry-day decision** (commits `live_state/`) →
    fetch previous published summary (decrypts `summary.json.enc` with `PAGE_PASSWORD`) → **encrypt
    summary for the Telegram job** with `CRYPTO_PASSWORD = TELEGRAM_BOT_TOKEN` → upload artifact
    `notify-data` (1 day) → Identify problem (on failure) → **Prepare Pages site** (with
    `PAGE_PASSWORD`: encrypted full page + `summary.json.enc`; without: public variant) → upload Pages.
  * Job `deploy-pages` (official deploy-pages action).
  * Job `notify`: Python + cryptography → download `notify-data` → decrypt → `notify_telegram.py`
    (failure message if run/deploy failed; `--mode evening` for the cron, else manual).
  * Job `keepalive` (schedule only): `gh api -X PUT …/workflows/run_strategy.yml/enable` (prevents the
    60-day auto-disable).
* **Workflow "Entry Day Reminder"** (`entry_day_reminder.yml`): cron `30 8 1-7 * *` (14:00 IST, days 1–7),
  silent unless today is the entry day; manual run (input `force`, default true) sends a 🧪 TEST
  reminder; own keepalive job.
* **Privacy design (user requirement)**: real account data must never be public. → Full page/summary only
  **encrypted**; without `PAGE_PASSWORD` the Pages site gets the public variant (no My Account); **no
  report/Excel artifacts**; Actions logs print only sheet row counts (no ₹ amounts) — `IN_CI` check.
  Still public by design: code, strategy record, strategy messages in logs, `entry_decisions.json`.
* **Pages**: source = GitHub Actions (enabled via API). Unlock page supports "Remember on this device"
  (localStorage key `etfDashPw`).

---

## 10. Outputs

* **Dashboard sections** (order): header/status · data/market banners · strategy KPI cards · Live
  Portfolio Actions (order plan / executed / preview) · Current Live Positions · **My Account** (KPIs
  incl. XIRR, vs Nifty, execution vs close; holdings vs strategy; plan vs actual; growth-of-100 chart;
  execution-cost table) · Risk Monitor · Current Strategy Signals · **Ranking Trend** · Live
  Performance · Next Rebalance · Live Charts (allocation, P&L, value, drawdown, monthly returns) · Live
  Rebalance History · Live Closed Trades · Strategy Rules · collapsed Historical Backtest Reference.
* **Excel sheets**: Live_Summary, Live_Positions, Live_Actions, Live_Signals, Live_Data_Status,
  Rank_Trend, My_Account, My_Holdings, My_Plan_vs_Actual, My_Trades (backup copy of sheet rows),
  My_Performance, My_Execution_Cost, My_Cash, Live_Closed_Trades, Live_Equity, Live_Rebalances,
  Live_Rebalance_Log, Hist_Summary, Hist_Monthly_Portfolio, Hist_Closed_Trades, Hist_Open_Positions,
  Hist_Momentum_Scores (22 sheets). **Only produced locally now** (not uploaded from GitHub).
* **Telegram** (`notify_telegram.py`): 🛒 order plan, ⛔ data incomplete, ✅ entry-day rebalance
  recorded, ⚠️ exit risk (stop hit / within 5% / outside top 6), ⚠️ data delayed, 📊 daily summary (incl.
  👤 My account, 📈 since start, 🧾 execution vs close), 👤 your account (reference) stop levels from your
  fills hit/near, ❌ run failed; evening mode de-duplicates alerts
  against the previously published summary; manual runs always send. Reminder script sends 🔔.

---

## 11. Current state (2026-10-06)

* Code: `main` @ **`1be7362`** + the commit adding this handover; local = GitHub.
* Live strategy record: **Oct 2026 entry (2026-10-01 close)**: MONQ50 258, MAFANG 201, MON100 155,
  PHARMABEES 1844, METALIETF 3965, MODEFENCE 496; cash ≈ ₹501. Next entry day **Mon 2026-11-02**.
* Real account: private — see §8 for where to read it (the user entered all six October fills, plus a
  ₹3,00,000 deposit in the Cash tab).
* Historical backtest (reference, 2021-01 → 2026-09-30, same-day method, holiday fix): CAGR ≈ 29.5%,
  MaxDD ≈ −24.8%, Sharpe ≈ 1.24.
* Recent runs all green (last scheduled 20:15 run 2026-10-05). Page is encrypted (PAGE_PASSWORD set).
* Yahoo data quirk active: as of Oct 6 morning, Yahoo was missing the Oct 5 close for most ETFs (usually
  restored the same/next evening). The dashboard shows a data-delay warning in that case.
* Excluded ETFs (insufficient Yahoo history): MID150BEES (Yahoo has only days since 2026-09-22 —
  data problem; alternatives MID150CASE / MIDCAPIETF were tried and reverted at user request), CHEMICAL,
  MOCAPITAL (Yahoo missing older history), GROWWPOWER, GROWWHOSPI (new listings). INTERNET is borderline
  (~262 rows) and flips in/out with Yahoo gaps.

---

## 12. How to run & test

* Local run (Windows, PowerShell, Python 3.14 at `C:\Users\Harshil\AppData\Local\Python\pythoncore-3.14-64`):
  `python ETF_Momentum_Test_1.py --force` → `reports\`. Local runs read the sheet via
  `local_config.json` and **do not** write `live_state/` (only CI does).
* GitHub manual run: `gh workflow run run_strategy.yml -R ynmarket/LIVE_ETF_MOMENTUM --ref main -f telegram=false`
  (ask the user first — it changes the live page). Watch: `gh run watch <id> -R … --exit-status`.
* Telegram dry run: `python notify_telegram.py --summary reports/summary.json --mode manual --dry-run`
  (set `GITHUB_REPOSITORY=ynmarket/LIVE_ETF_MOMENTUM` for links).
* Reminder dry run: `python entry_day_reminder.py --today 2026-11-02 --dry-run`.
* Proven test techniques (reuse them):
  * **No-strategy-change proof**: copy the previous commit's script (`git show HEAD:ETF_Momentum_Test_1.py`)
    to a scratch dir, run both simultaneously, compare all Excel sheets with `pandas.testing.assert_frame_equal`.
  * **Scenario simulation**: copy the script to a scratch dir and patch lines, e.g.
    `MARKET_HOURS = …` (force open/closed), `_now = datetime(2026,10,1,14,0)`, `data_end_ts = pd.Timestamp("…")`,
    `TODAY = pd.Timestamp("…")`, or replace `read_trade_sheet()` with a stub returning sample rows.
    Run with `CI=true` to exercise CI-only paths (saving decisions, log masking).
  * Workflow shell steps can be run locally in **Git Bash** with dummy secrets.
  * Scratch dir: `C:\Users\Harshil\AppData\Local\Temp\claude\F--1-AI-Market-TEST-LIVE-ETF-MOMENTUM\<session>\scratchpad`.
* Windows gotchas: PowerShell 5.1 (no `&&`); `Set-Content`/`Out-File` add BOM/CRLF (prefer Python or the
  Write tool); inline Python containing `rm -f` or regex `\d` inside PowerShell commands can trip a safety
  filter → put such code in a `.py` file; `Select-String` may misread UTF-8 (use Python to search HTML).
  The in-app browser pane often does not repaint screenshots → verify via `javascript_tool` DOM checks.

---

## 13. Change history (what changed, in order)

1. Built the script (backtest + live portfolio + Excel + dark HTML report) from the user's spec.
2. Live dashboard redesign: KPI cards, actions, positions, risk monitor, signals, live charts; historical
   backtest moved to a collapsed section.
3. Previous-day signal + first-trading-day execution (backtest + live), look-ahead validation.
4. Data-delay warnings (stale prices per ETF).
5. GitHub deployment: `reports/`, IST clock, CI detection, requirements, workflow, Pages, mobile CSS fix.
6. Telegram + schedules (20:15 evening, 08:15 morning) + keep-alive.
7. Oct 1 incident fixes: market-hours guard, pre-close order plan, completeness check, frozen
   rebalances (`live_state/rebalances.json`); 08:15 morning run removed.
8. Mistake: removed the 20:15 schedule (misunderstanding) → **reverted** (`ea4be6b`).
9. **Timing redesign (current)**: entry day = signal & execution day (same-day close); order plan from
   live prices; saved entry-day decisions (`entry_decisions.json`); all freezing removed; October
   recomputed from the Oct 1 close.
10. My Account from Google Sheet; entry-day reminder workflow; holiday-placeholder fix.
11. K: MONQ50 → International. E: XIRR. F: account vs strategy vs Nifty (time-weighted). H: execution
    cost. M: ranking trend. R: password-protected page.
12. Privacy hardening (no artifacts with data, encrypted hand-off, public variant, masked logs, privacy check).
13. Live view: live prices for all current values during market hours (★ live points on charts).
14. Reference stop levels from the user's own fills in My Account + 👤 Telegram alert (info only; §8);
    strategy trailing-stop display now shows the stored engine peak (no today's price).
15. ~14:30 IST midday scheduled update (live prices, full SL detail; keeps a choice saved earlier that
    day); per-position SL lines in all summaries; reminder mentions the 14:30 fallback.

## 14. What was tried and did NOT work / was reverted (don't repeat)

* **Treating Yahoo's intraday bar as a close** → fake "EXECUTED" at 9:45 AM prices and wrong signals
  (Oct 1). Fixed by the market-hours guard; live prices only for *current* values/order plan.
* **08:15 morning scheduled run** → Yahoo withdraws the previous day's rows overnight → wrong signal
  (HDFCSML250 instead of MODEFENCE). Removed permanently.
* **Previous-day signal + frozen signals/fills** → user rejected; replaced by entry-day same-day method
  with saved 2 PM decisions.
* **Removing the 20:15 schedule / "manual-only" mode** → was a misunderstanding; reverted.
* **Universe swaps** MID150BEES→MID150CASE and MOCAPITAL→GROWWCAPM → reverted at user request.
* **Per-ETF zero-volume filter** for holidays → wrong (illiquid ETFs like INTERNET report volume 0 every
  day); use "any ETF traded that day" instead.
* **Plotly library inside the first chart div** → removing a section broke all charts; now in `<head>`.
* **Grid `auto-fit` for 2-column chart rows** → charts drawn full-width then clipped; fixed 2-column grid
  + resize on load.
* **First My Account Excel edit** inserted a loop at the wrong indentation and silently dropped sheets —
  always re-check the sheet list after editing `_write_excel()`.
* **Google Drive connector to find the Sheet ID** → user said ask instead.

---

## 15. Next steps

### Immediate / upcoming
1. **Mon 2026-11-02 (first November entry day)**: verify the 🔔 reminder arrives ~14:00 IST, the user's
   manual run shows **🛒 ORDER PLAN** (live), `live_state/entry_decisions.json` gets a `2026-11` entry
   ("entry-day run, live prices HH:MM IST"), and the 20:15 run shows **EXECUTED** with that choice at the
   close. If anything is off, check `gh run view <id> --log` first.
2. After the user enters November fills: confirm plan vs actual, execution cost and XIRR (XIRR appears
   once ≥ 30 days since 2026-10-01, i.e. from ~2026-10-31).
3. **GitHub notice**: `ubuntu-latest` migrates to Ubuntu 26 from 2026-10-19 — watch the next runs; pin
   `runs-on: ubuntu-24.04` if anything breaks (ask before changing).

### Open options offered to the user (not decided / not built)
* Move the entry-day reminder from 14:00 to **13:45 IST** (GitHub cron can be late).
* From the suggestion list (letters as presented to the user):
  **A** liquidity check in the order plan (e.g. INTERNET shows zero volume), **B** ETF premium/discount
  vs NAV (AMFI), **C** price sanity check (±15% moves), **D** "enter your fills" follow-up reminder,
  **G** tax-ready FIFO capital-gains ledger, **I** copy-ready order list with limit prices, **J**
  pre-entry-day preview, **L** monthly history cards, **N** lighter mobile page (~6.4 MB now),
  **O** backtest with real costs, **P** "run didn't happen" alert, **Q** weekly health message.
  Already done: K, E, F, H, M, R (+ real fills via Sheet, reminder, holiday fix, privacy, live view).
* Data-source resilience (NSE bhavcopy / broker API + local price store) — user wants this later as a
  **separate project**; do not mix in.

### How to start a new task
1. `git -C F:\1.AI_Market\TEST_LIVE_ETF_MOMENTUM pull --rebase` (the workflow commits to `main`).
2. Read the user's request; restate it if ambiguous; confirm before implementing.
3. Implement → `python ETF_Momentum_Test_1.py --force` (+ scenario tests) → commit locally →
   show results → **ask before pushing**.

---

## 16. Glossary

* **Entry day** — first trading day of the month; signal and execution day.
* **Order plan** — entry-day, market-hours view: what to buy/sell now with quantities from live prices.
* **Live view** — market-hours valuation of current values with live prices (not recorded).
* **Strategy record / live portfolio** — the model ₹3,00,000 portfolio from 2026-10-01 (closing prices).
* **My Account** — the user's real holdings from the Google Sheet.
* **Saved entry decision** — the ETF choice made on the entry day, stored in `live_state/entry_decisions.json`.
* **Holiday placeholder** — Yahoo's zero-volume fake bar on an NSE holiday (removed).
* **Public variant** — dashboard/summary without My Account (used only if `PAGE_PASSWORD` is missing).
