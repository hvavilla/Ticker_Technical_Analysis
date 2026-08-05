# Technical Analysis Report Generator

A learning tool for reading the key technical indicators across several companies at
once. It reads your tickers from `watchlist.xlsx`, pulls daily price history, computes a
standard indicator suite, and rolls each company up into four plain-English parameters, a
reversal flag, and one composite verdict — so you can compare names side by side and
see exactly how each conclusion was reached.

**Output: one HTML file per mode**, written by default:

```
Report_day.html      hourly bars, momentum-weighted
Report_swing.html    daily bars, balanced
Report_long.html     weekly bars, trend-weighted
```

Self-contained and colour-coded; double-click to open in any browser, works offline,
nothing to install to view them. **Hover any cell in
the summary table** to see the conditions that produced that label, with the actual
values substituted in.

Optionally reads `portfolio.xlsx` and adds a portfolio table with cost basis, market
value, gain/loss, weight, and days held.

**Three modes** — `day`, `swing`, `long` — switch the bar interval and reweight the
verdict for that horizon. See section 6.

**Two input workbooks, one HTML output.** Start from the supplied examples:

```bash
cp watchlist_example.xlsx watchlist.xlsx     # Windows: copy watchlist_example.xlsx watchlist.xlsx
cp portfolio_example.xlsx portfolio.xlsx     # optional
```

Each example has an **Instructions** sheet explaining its columns.

---

## The output parameters

### 1. Trend — where price sits relative to its own averages
`Uptrend` · `Downtrend` · `Sideways`

Scores four checks: price vs SMA-50, SMA-20 vs SMA-50, the 20-bar slope of SMA-50, and
price vs SMA-200. Two or more agreeing gives a verdict; otherwise Sideways.

### 2. Trend Strength — how much conviction is behind the move
`Strong bullish` · `Bullish (emerging)` · `Strong bearish` · `Bearish (emerging)` ·
`Ranging / no clear direction` — each qualified **strengthening**, **weakening**, or **stable**

Built on ADX, which measures strength *without* direction; the ±DI lines say which way.
Below ADX 20 nothing has enough conviction to call. The qualifier comes from ADX's own
20-bar slope, separating two situations that share a label: ADX at 27 rising from 21 is
a trend gathering force, ADX at 27 falling from 45 is one bleeding out.

### 3. Momentum — is the move speeding up or running out of steam
`Bullish & accelerating` · `Bullish but fading` · `Bullish, turning up/down (early)` ·
`Bearish & accelerating` · `Bearish but recovering` · `Bearish, turning up/down (early)` · `Flat`

Territory from the MACD histogram's sign; acceleration from least-squares slopes over 5
and 10 bars. Slopes agreeing means *confirmed*; disagreeing produces an *(early)* label
marking an unconfirmed inflection.

### 4. Overbought / Oversold — is price stretched too far, too fast
`Overbought (2/3 or 3/3)` · `Approaching overbought` · `Neutral` ·
`Approaching oversold` · `Oversold (2/3 or 3/3)` — with **how many of the last 60 sessions** were also stretched

Three indicators vote: RSI, Bollinger %B, Stochastic %K. The session count matters
because "overbought today" and "overbought 18 of the last 20 sessions" are different
situations wearing one label. The **RSI percentile** ranks today's reading against the
stock's own past year, which a fixed 70/30 threshold can't: RSI 65 is unremarkable on a
name that habitually runs hot and stretched on a quiet one.

### 5. Reversal Signals — are independent methods pointing at a turn
`Possible bullish/bearish reversal (2/3 or 3/3)` · `Watch — weak signal (1/3)` · `None`

Three methods vote: a moving-average crossover, a MACD crossover, and RSI divergence.
Requiring agreement stops one noisy signal producing a call alone.

### 6. Verdict — one composite rating
`STRONG BUY` · `BUY` · `HOLD` · `SELL` · `STRONG SELL`

A weighted sum of the five parameters above, normalised to a fixed **−10…+10** scale.
Raw component contributions:

| Input | Range | How |
|---|---|---|
| Trend | ±2 | Uptrend +2, Sideways 0, Downtrend −2 |
| Trend Strength | ±2.5 | ADX magnitude (2 / 1 / 0) × DI sign, scaled ×1.25 strengthening, ×0.75 weakening |
| Momentum | ±2 | territory ±1, short slope ±0.75, long slope ±0.25 |
| Overbought / Oversold | ±2 | **contrarian** — stretched high argues against buying |
| Reversal Signals | ±1 | ±1 corroborated, ±0.5 for a lone signal |

Each is multiplied by the **mode's weight** (section 6), summed, then divided by the
maximum achievable for that mode and scaled to ±10. Normalising is what lets the bands
mean the same thing in every mode, and it fixed an earlier problem where the extremes
were arithmetically almost unreachable and everything clustered in HOLD.

Bands: **≥ +6** STRONG BUY · **+2.5 to +6** BUY · **−2.5 to +2.5** HOLD ·
**−6 to −2.5** SELL · **≤ −6** STRONG SELL. All in `CFG["verdict_bands"]`.

The overbought/oversold term is halved inside a strong trend (ADX ≥ 25) and halved again
on names stretched 20+ of the last 60 bars — because overbought readings mislead exactly
there, where a hot stock keeps running. An RSI percentile ≥95 or ≤5 adjusts it ±0.5.

**The verdict ignores what you paid, deliberately.** The indicators are blind to your cost
basis, so anchoring a rating on it would report something about you while looking like a
statement about the stock — and would encode the *disposition effect* (Shefrin & Statman,
1985), the documented tendency to sell winners early and hold losers too long. The
technical picture is identical whether you bought at $90 or $190.

## Lookbacks are fixed, not something you set

Each is tied to the period of the indicator it reads, not a calendar number you pick:

| Parameter | What it examines |
|---|---|
| Trend | latest values + 20-bar slope of SMA-50 |
| Trend Strength | latest values + 20-bar slope of ADX |
| Momentum | histogram sign + 5-bar and 10-bar slopes |
| Overbought / Oversold | latest values + 60-bar persistence + 250-bar RSI percentile |
| Reversal Signals | 10-bar cross window + swing pivots within 120 bars |

About three years of history is pulled per company so every window has enough bars
behind it, including the 250-bar RSI percentile which sits on top of RSI's own settling
period. Only the final readings are reported.

**Why there's no `--months` option:** a lookback suiting SMA-200 doesn't suit the MACD
histogram, so one user-supplied window can't serve all five. Tying each to its own
indicator's period is more defensible than a number that appears to drive the analysis
but mostly wouldn't.

---

## 1. Prerequisites

**Python 3.9 or newer:**

```bash
python --version
```

If that errors or shows Python 2.x, try `python3 --version` and use `python3` throughout.
If Python isn't installed, get it from [python.org/downloads](https://www.python.org/downloads/)
— on Windows, tick **"Add Python to PATH"** during installation.

---

## 2. Set up a virtual environment

Keeps this project's packages separate from the rest of your system. Create once,
activate each time. Put the script, `requirements.txt`, and this README in one folder,
open a terminal, `cd` into it.

### macOS / Linux
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### Windows — Command Prompt
```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

### Windows — PowerShell
```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

> If PowerShell blocks it with an execution-policy error, run this once and activate again:
> ```powershell
> Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
> ```

Your prompt shows `(.venv)` when active. Leave it with `deactivate`.

---

## 3. Install dependencies

```bash
pip install -r requirements.txt
```

| Package | Why |
|---|---|
| `yfinance` | Price history, and resolving names to tickers |
| `pandas` | Time-series handling |
| `numpy` | Indicator maths |
| `openpyxl` | Reading the `.xlsx` input workbooks |

No compiled libraries — every indicator is implemented directly in pandas/numpy, so
there's no TA-Lib to build. `openpyxl` is needed only to *read* the two input workbooks;
the report itself is pure HTML.

---

## 4. Run it

```bash
python technical_analysis.py
```

No arguments needed. It reads `watchlist.xlsx` and `portfolio.xlsx` from the current
folder if they exist, and writes **all three reports**.

Each mode fetches its own bars — the intervals differ — so a default run makes three
passes over your tickers. Ten companies means thirty downloads and roughly three times
the runtime of a single mode. Use `--mode` when you only need one.

| Flag | Meaning | Default |
|---|---|---|
| `--mode` | Run a single mode: `day`, `swing`, or `long` | all three |
| `--interval` | Override the bar interval; requires `--mode` | mode's own |
| `--watchlist` | Watchlist workbook | `watchlist.xlsx` |
| `--portfolio` | Holdings workbook | `portfolio.xlsx` |
| `--output` | Output filename base, no extension | `report` |

### Examples

```bash
# Default: watchlist.xlsx + portfolio.xlsx in the current folder
python technical_analysis.py

# Point at different workbooks
python technical_analysis.py --watchlist chips.xlsx --portfolio holdings.xlsx

# Long-term mode, weekly bars
python technical_analysis.py --mode long

# Change the filename prefix (date, time and mode are always appended)
python technical_analysis.py --output chips
```

### The watchlist workbook

Sheet `Watchlist`, one ticker per row:

| ticker | note |
|---|---|
| NVDA | GPUs — dominant AI accelerator |
| AMD | Instinct accelerators |
| ASML | EUV lithography |

`ticker` is required; `note` is for your reference and ignored by the script. Headers are
matched case-insensitively, blank rows are skipped, and duplicates are read once.

**If `watchlist.xlsx` is missing, the script falls back to a built-in list** of ten AI
chip tickers (NVDA, AMD, AVGO, MRVL, ASML, MU, GOOGL, INTC, TSM, QCOM) and says so in the
console. It won't stop.

Tickers beat company names — unambiguous, and they skip the name-search step. Non-US
listings need their exchange suffix: `SAP.DE`, `BP.L`, `7203.T`, `RELIANCE.NS`. `ASML`
and `TSM` are US-listed ADRs, so they need no suffix.

## 5. The portfolio table

Copy `portfolio_example.xlsx` to `portfolio.xlsx` and it's picked up automatically.
Sheet `Portfolio`, **one row per purchase, not per company:**

| ticker | shares | price | date | note |
|---|---|---|---|---|
| NVDA | 25 | 118.40 | 2025-03-14 | initial position |
| NVDA | 10 | 142.05 | 2025-11-02 | added on the dip |
| AVGO | 15 | 178.20 | 2025-06-21 | |

`ticker`, `shares`, and `price` are required — `price` is what you paid **per share**,
not the total. `date` and `note` are optional; `date` drives the Days held column and
works either as a real Excel date or as text like `2025-01-15`.

Bought the same stock twice? Add another row. The script sums the shares, computes your
weighted average cost, and marks the row `(2 lots)` — the example file ships with two
NVDA rows so you can see it: 100 at 100.00 plus 50 at 125.00 becomes 150 shares at
108.33. Keeping lots separate keeps the history correct as you add.

The table shows shares, average cost, cost basis, last price, market value, gain/loss in
both dollars and percent, portfolio weight, days held, and the verdict — plus a TOTAL
row. Gains green, losses red.

Malformed rows are skipped individually, with the reason shown in the console and in an
amber banner above the table, so one bad cell doesn't lose the file. Holdings that fail
to resolve still appear, with `-` for price and verdict. **Anything held is analysed
automatically**, whether or not it's on the watchlist.

Values use the **last close**, not live prices, and there's no dividend, fee, or
corporate-action handling — a rough tracker, not a brokerage statement. Fold fees into
the price you paid if you want them counted.

## 6. Trading modes

```bash
python technical_analysis.py --mode long
```

Every period in `CFG` counts **bars, not calendar time**, so changing the interval
rescales the entire indicator suite at once — no periods need editing.

| Mode | Interval | Horizon | RSI-14 covers | SMA-200 covers |
|---|---|---|---|---|
| `day` | 60m | hours | ~2 days | ~28 days |
| `swing` *(default)* | 1d | days–weeks | ~3 weeks | ~9.5 months |
| `long` | 1wk | months–years | ~3 months | ~4 years |

Modes also **reweight the verdict**, which is what makes each one fit its horizon rather
than just run slower:

| Component | day | swing | long | Reasoning |
|---|---|---|---|---|
| Trend | 0.5 | 1.0 | **2.0** | Long MAs are near-meaningless intraday, decisive over years |
| Trend Strength | 1.0 | 1.0 | 1.5 | |
| Momentum | **2.0** | 1.0 | 0.5 | Intraday is largely a momentum game |
| Overbought / Oversold | **2.0** | 1.0 | 0.25 | Mean reversion pays intraday; over years it's entry timing at most |
| Reversal | 1.5 | 1.0 | 0.5 | |

`swing` reproduces the original behaviour exactly.

### Confirmation filter

`long` mode requires a verdict to **hold for 2 consecutive weekly bars** before it takes
effect. Until then the previous verdict stands and the new one shows as pending:

```
SELL (STRONG SELL pending)
```

This exists because over half the raw score sits in components that turn over in days —
momentum and overbought/oversold — and the bands are only 2.5 apart. Without the filter,
a single momentum sign flip could carry a name from BUY to SELL in a week, which is no
use over a multi-year horizon. `day` and `swing` report every bar (`confirm_bars: 1`).

The filter is honest rather than approximated: it re-runs every classifier on truncated
history to find what the verdict actually was on prior bars, instead of storing state
between runs.

### Overriding the interval

```bash
python technical_analysis.py --mode day --interval 15m
```

`--interval` applies to one mode's bars, so it requires `--mode`. Running it against the
three-mode default would be ambiguous, and the script says so rather than guessing.

**yfinance caps intraday history**, and the script clamps the request rather than letting
it return empty:

| Interval | Max history |
|---|---|
| 1m | 7 days |
| 2m–30m, 90m | 60 days |
| 60m / 1h | 730 days |
| 1d, 1wk, 1mo | unlimited |

At 15m you get ~60 days of bars, so the slow components — SMA-200 and the 250-bar RSI
percentile — will report `Insufficient data`. The console warns when you pick an interval
where that happens. 60m is the only intraday interval with enough history for everything.

### On day mode specifically

Worth stating plainly: **this is not a day-trading system.** The data is free and delayed
roughly 15 minutes, Yahoo's intraday bars are less reliable than its end-of-day, and the
tool ignores spreads, liquidity, and fees — which dominate outcomes at that timescale.
Day trading is also, consistently across studies, where most retail participants lose
money. Use `--mode day` to watch how the indicators behave on short timeframes; that's
what it's good for.

Put the mode in the output name so reports can't be confused:

```bash
python technical_analysis.py --mode long
```

### Output filenames

One file per mode, fixed names, no timestamp:

```
Report_day.html
Report_swing.html
Report_long.html
```

Each run overwrites the previous set, which is what makes them practical to bookmark or
sync — the path never changes. `--output chips` gives `chips_day.html` and so on.

If you want a history of past reports, use version control or copy them aside; the script
no longer stamps filenames.

### Only completed bars are used

yfinance returns the **in-progress** bar, and every period and threshold here assumes
completed bars — so the script drops the final bar unless its period has closed. That
matters most in `long` mode, where on four weekdays out of five the current weekly bar is
a two- or three-day stub. Scoring it wouldn't be less accurate, it would be simply wrong.

Where completeness can't be established — an unrecognised interval, or a missing timezone
database — the bar is dropped anyway. Losing one bar of recency is much cheaper than
scoring a fragment. The console says when this happens, and each company card shows the
**last complete bar** it actually used:

```
NVDA: BUY  (Uptrend, 517 bars)
        dropped in-progress bar 2026-07-29 00:00 - not yet closed
```

A consequence worth expecting: in swing mode before the close, the report reflects
*yesterday's* completed session. After the close (16:15 ET) today's bar is included.

### It refuses rather than guessing

If a ticker has fewer than **220 complete bars**, it isn't scored at all — it gets a row
in the summary explaining why, and no card:

```
Needs 220 complete day bars, has 150. Scoring this would use inputs the
parameters do not support.
```

This replaced a worse behaviour. Previously, short history meant SMA-200 was silently
absent, so Trend ran three checks instead of four while keeping the ±2 threshold — making
a verdict *easier* to reach on worse data. In testing, 150 weekly bars produced a
confident `STRONG BUY` with no warning at all. The trend threshold now also scales with
how many checks actually contributed: with three checks it requires ±3, not ±2.

Anything computed over a narrower window than designed is reported rather than hidden,
both in the console and as an amber banner on the company card:

```
degraded: RSI percentile ranked over 230 readings, not the full 250
```

### History is derived from the interval

The number of calendar days requested is computed from the interval, targeting ~450 bars,
then clamped to what yfinance will serve:

| Interval | Days requested | Bars |
|---|---|---|
| 15m | 28 | ~500 |
| 60m | 107 | ~516 |
| 1d | 750 | ~517 |
| 1wk | 3622 | ~517 |

Previously history was fixed per mode, so `--mode swing --interval 1wk` fetched 1,100 days
— about 157 weekly bars — and SMA-200 vanished silently. Deriving it from the interval
means an override can't under-fetch.

Long mode needs roughly ten years of weekly history, so recently listed companies will be
skipped in that mode. That's the data genuinely not existing, not a fault to fix.

## 7. Reading the output

Start with the **Summary** table. Hover any coloured cell to see the decision trace —
each condition with real values, what it contributed, and the final verdict on its own
line. For example:

```
ADX 14.62 < 20                       | no conviction -> ranging
ADX slope -0.6152 < -0.05            | weakening
```

**Colour follows category, not the exact label:** green bullish, red bearish, grey
neutral, amber overbought, blue oversold. `Bullish but fading` still shows green — the
territory sets the colour, the wording carries the nuance. A green cell isn't
automatically good news; read the label.

**Tooltips need a mouse.** Hover doesn't exist on phones or tablets, and tooltips don't
print. The per-company cards below carry the same numbers, so nothing is unreachable.

---

## 8. Customising

**`DEFAULTS`** — the fallback ticker list, the two workbook filenames, and the output name.

**`CFG`** — indicator periods, thresholds, classifier lookbacks, and verdict bands.
Indicator settings are the textbook standards (RSI 14, MACD 12/26/9, ADX 14, Bollinger
20/2, Stochastic 14/3, SMA 20/50/200).

Three settings worth understanding first:

- **`verdict_bands`** — the five cut points. Expect to move these; see the note above.
- **`mom_slope_long` (10)** — the MACD histogram with 12/26/9 inputs oscillates on a
  ~9–20 bar cycle, so a slope measured over much more than 10 bars can straddle a whole
  up-then-down swing, average to nearly flat, and hide a real turn. Lengthening it buys
  smoothness at the cost of the early warning the signal exists to give.
- **`pct_bars` (250)** — the RSI percentile window. Shorter makes it jumpy; longer than
  the history pulled silently returns no percentile. Raise `history_days` alongside it.

Change an indicator's period and scale its lookback with it — that's what keeps the
settings coherent rather than arbitrary.

---

## 9. Running it daily

Nothing in the script schedules itself. To automate, use your OS scheduler and
date-stamp the output so runs don't overwrite each other.

**macOS / Linux** — `crontab -e`, then (weekdays at 18:00, using absolute paths):

```
0 18 * * 1-5 cd /path/to/folder && .venv/bin/python technical_analysis.py
```

The script stamps its own filenames, so no date substitution is needed in cron.

**Windows** — Task Scheduler → Create Basic Task → Daily, with the action pointing at
`.venv\Scripts\python.exe` and the script as its argument, and "Start in" set to the
script's folder.

Markets are shut at weekends and on holidays, so a weekend run just repeats Friday's
close. Weekdays only, after the close, is the sensible cadence.

---

## 10. Cloud automation (GitHub Actions + Pages)

This repo also runs itself in the cloud, so it works even when your own computer is off.
A GitHub Actions workflow at
[`.github/workflows/daily-report.yml`](.github/workflows/daily-report.yml):

1. Checks out the repo on a temporary GitHub-hosted machine
2. Installs `requirements.txt`
3. Runs `python technical_analysis.py`, producing `Report_day.html`, `Report_swing.html`,
   `Report_long.html`
4. Commits those 3 files back to `main`
5. Publishes them to GitHub Pages

**Schedule:** weekdays at 21:30 UTC (4:30pm ET, ~30 minutes after the US market close —
matches the "weekdays after the close" guidance above). Cron time is fixed in UTC, so it
drifts by an hour relative to market close across DST transitions each spring/fall.

**Live reports** (always reflect the latest run):

```
https://hvavilla.github.io/Ticker_Technical_Analysis/Report_day.html
https://hvavilla.github.io/Ticker_Technical_Analysis/Report_swing.html
https://hvavilla.github.io/Ticker_Technical_Analysis/Report_long.html
```

### Triggering a run manually

**Option A — GitHub website, no installs needed:**

1. Go to the [Actions tab → "Daily technical analysis report"
   workflow](https://github.com/hvavilla/Ticker_Technical_Analysis/actions/workflows/daily-report.yml).
2. Click **"Run workflow"** (top right), then the green **"Run workflow"** button in the
   dropdown that appears.
3. Wait under a minute, then refresh one of the report URLs above.

**Option B — from your own computer, using the GitHub CLI.** These steps assume no prior
terminal experience.

#### 1. Install Git

Git downloads ("clones") the project files to your computer.

*macOS:* Open **Terminal** (`Cmd + Space`, type `Terminal`, Enter), then run:
```bash
git --version
```
If it isn't installed, a popup offers to install "Command Line Developer Tools" — click
**Install** and wait a few minutes.

*Windows:* Download the installer from [git-scm.com/downloads](https://git-scm.com/downloads),
run it, and click **Next** through the defaults. Then open **Command Prompt** (Windows key →
type `cmd` → Enter) — that's your terminal for the rest of these steps.

#### 2. Install GitHub CLI (`gh`)

This is what lets you trigger the workflow from a terminal.

*macOS:* If you don't have Homebrew yet, install it first:
```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```
Then install GitHub CLI:
```bash
brew install gh
```

*Windows:* Download and run the installer from [cli.github.com](https://cli.github.com/),
clicking **Next**/**Install** through the defaults, then close and reopen Command Prompt.

*Verify (both platforms):*
```bash
gh --version
```
A version number should print, not an error.

#### 3. Log in to GitHub

```bash
gh auth login
```
Answer the prompts (arrow keys + Enter):
- Account → **GitHub.com**
- Protocol → **HTTPS**
- Authenticate Git with your GitHub credentials? → **Yes**
- How would you like to authenticate? → **Login with a web browser**

Copy the one-time code it shows, press Enter, and your browser opens to a GitHub page —
paste the code, click **Continue**, then **Authorize**. The terminal should then show
"✓ Logged in as ...".

> The account you log in with needs write access to this repo, or the trigger command in
> step 5 will fail with a permissions error.

#### 4. Clone the repo

```bash
cd ~/Desktop
git clone https://github.com/hvavilla/Ticker_Technical_Analysis.git
cd Ticker_Technical_Analysis
```
(Windows: use `cd %USERPROFILE%\Desktop` instead of `cd ~/Desktop`.)

#### 5. Trigger the workflow

```bash
gh workflow run daily-report.yml
```

Check it's running:
```bash
gh run list --workflow=daily-report.yml --limit 1
```
The status column moves from `in_progress` to `completed`/`success`, usually within a
minute. Then refresh one of the report URLs above to see the new run.

---

## 11. Troubleshooting

**`Missing dependency 'yfinance'` or `'openpyxl'`** — environment not active, or install
skipped. Run `pip install -r requirements.txt`.

**"no watchlist.xlsx found — using the built-in list"** — expected if you haven't created
it. Copy `watchlist_example.xlsx` to `watchlist.xlsx`.

**`no 'ticker' column found`** — the header cell is spelled something else (`symbol`,
`stock`). Rename it to `ticker`; case and surrounding spaces don't matter.

**Wrong ticker resolved** — pass the ticker directly. Because tickers are tested first,
an input that's both a valid ticker and a company name resolves as the ticker: `ALL`
gives you Allstate.

**`Could not resolve to a ticker` / `Not enough price history`** — usually delisted,
newly listed, or non-equity. Others still report; the reason appears in the summary row.

**`RSI percentile` shows `-`** — under ~100 bars of history, so there's not enough of the
stock's own past to rank against. Expected for recent listings.

**Every verdict is HOLD** — the bands may be too wide for your names. Check
`Score (-10 to +10)` on the cards and narrow `CFG["verdict_bands"]`.

**A ticker was skipped for too few bars** — under 220 complete bars at that interval.
Expected for recent listings, and for most companies in `long` mode, which needs about ten
years of weekly history.

**One mode failed but the others worked** — by design. A failure in one mode doesn't
abort the run; the remaining reports are still written.

**"dropped in-progress bar"** — normal. The current bar hasn't closed, so it isn't scored.

**A verdict shows `(X pending)`** — long mode only. The new verdict hasn't held its two
bars yet, so the previous one still stands. Working as designed.

**Portfolio table missing** — no `portfolio.xlsx` in the folder. The console prints the
lot count when a file is found, so check that line first.

**Excel won't let you save over an open workbook** — close `watchlist.xlsx` in Excel
before running, or the read may fail on Windows.

---

## What this tool is and isn't

It reports what a standard set of indicators **currently says**, consistently and
without hand-calculation. That's genuinely useful for learning how the indicators behave
and interact — and the hover traces mean you can always see why a label came out as it did.

It is **not validated.** Thresholds, weights, and bands are conventions and judgment
calls, not values tested against outcomes. Nothing here has been backtested, so there is
no evidence that any label or verdict precedes any particular price move. A cell reading
`STRONG BUY` carries more authority than the underlying logic has earned.

Three limitations worth knowing as you learn:

- The three indicators voting on Overbought/Oversold (RSI, %B, Stochastic) largely
  measure the same thing — distance from a recent mean — so "3 of 3" is closer to one
  signal counted three times than to three independent confirmations. Partly true of the
  reversal vote too, where the MA crossover and MACD are both built on EMA 12/26.
- Overbought/oversold logic works in ranging markets and misleads in strong trends,
  where RSI can hold above 70 for months while price climbs. The dampers reduce this but
  don't remove it.
- Volume is not used at all. Volume confirmation is standard practice and its absence is
  a real gap.

**Not investment advice.** This can't account for your circumstances, timeframe, tax
position, or risk tolerance, and it is not a substitute for a licensed financial adviser.
Data comes from Yahoo Finance via `yfinance` — free, unofficial, occasionally delayed or
incomplete.
