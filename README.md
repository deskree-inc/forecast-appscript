# Deskree forecast — Google Sheets

Google Apps Script project used to **build and run financial forecasting for Deskree** inside Google Sheets. The spreadsheet model covers drivers, funding, headcount, revenue, P&L, cash flow, scenarios, and a **benchmark reality check**; a custom sidebar loads packaged scenarios into the input sheet.

## What’s in this repo

| File | Role |
|------|------|
| `SetupMain.gs` | `setupFinancialModel()` — orchestrates setup; logs each phase to **Logger** (Executions / View → Logs). Set `SETUP_PROGRESS_TOAST = true` for sheet toasts. |
| `ModelConstants.gs` | Shared maps: `DR`, `REVCOLS`, `PNL`, `CF`, `SUM`, etc. (used by setup + `benchmarks.gs`). |
| `SetupHelpers.gs` | Shared layout helpers (`hdr`, `colLetter`, …). |
| `SetupInstructions.gs`, `SetupDrivers.gs`, `SetupFunding.gs`, `SetupHeadcount.gs`, `SetupRevenue.gs`, `SetupPnL.gs`, `SetupCashFlow.gs`, `SetupSummary.gs`, `SetupBenchmarksTab.gs` | One tab (or tab pair) each; all called from `SetupMain.gs`. |
| `ScenarioSidebar.gs` | Custom menu, `applyScenario` / `getCurrentScenario` (**v2** Drivers layout), and opens the HTML sidebar. |
| `ScenarioSidebarView.html` | Sidebar UI (hosted as an HTML file in the Apps Script project). |
| `Benchmarks.gs` | `runBenchmarks()` — reads **Drivers**, **Headcount**, **P&L**, and **Cash flow**, scores key SaaS metrics, and writes results to **🚦 Benchmarks**. |

## Model overview

Running `setupFinancialModel()` builds these sheets:

- **Instructions** — how to use the model (inputs vs formulas) and workflow (including when to run benchmarks).
- **Drivers** — the **only** tab where you enter assumptions: funding rounds, ARR targets, ICP segments (mid-market & enterprise), logo ramp, maintenance ratios (AE / FDE / CSM), department defaults, individual roles, marketing, infrastructure, and sales comp.
- **Funding**, **Headcount**, **Revenue**, **P&L**, **Cash flow**, **Summary** — calculated views driven from Drivers.
- **Scenarios** — scenario comparison / framing.
- **Benchmarks** — populated when you run **Check Benchmarks** (not auto-updated on every edit).

The model follows **inputs → calculations → outputs**: keep assumptions in **Drivers** so downstream formulas stay intact.

## Tetrix menu

From `ScenarioSidebar.gs`, the **📊 Tetrix** menu provides:

1. **Open Scenario Loader** — HTML sidebar to apply a scenario or inspect **current model state** (`getCurrentScenario()`).
2. **Check Benchmarks** — runs `runBenchmarks()` and fills **🚦 Benchmarks** with traffic-light style checks (e.g. CAC payback, LTV:CAC, implied NRR vs churn/expansion, growth vs Bessemer-style heuristics, ARR vs capital raised, AE account load, gross margin, burn multiple when wired).

After loading a scenario, the script suggests running **Check Benchmarks** before sharing numbers externally.

## Scenario sidebar (v2)

The loader writes and reads the v2 Drivers layout: **funding rounds** (up to five), **ARR targets** and horizon, **segment** economics (MM/ENT), **logo ramp**, **maintenance ratios**, **headcount** department defaults and up to **ten named positions**, plus **marketing**, **infrastructure**, and **sales** fields. The HTML file name in Apps Script must stay **`ScenarioSidebarView`** (matching `createHtmlOutputFromFile("ScenarioSidebarView")`).

## Setup in Google Sheets

1. Open or create a Google Sheet for Deskree forecasting.
2. **Extensions → Apps Script** and create or open a project.
3. Add **all** `.gs` files from this repo (or push via clasp): `SetupMain`, `ModelConstants`, `SetupHelpers`, every `Setup*.gs` tab module, `ScenarioSidebar.gs`, and `Benchmarks.gs`.
4. **File → Add file → HTML**, name **`ScenarioSidebarView`**, paste `ScenarioSidebarView.html`.
5. Save.
6. Run **`setupFinancialModel`** once and authorize when prompted. If setup hangs, open **Executions** in the script editor and inspect **Logs** for the last completed phase.
7. Reload the spreadsheet; use **📊 Tetrix** for the scenario loader and benchmark check.

## Development notes

- Source lives here for copy/sync into Apps Script; add [clasp](https://github.com/google/clasp) locally if you want push/pull workflows.
- Re-running `setupFinancialModel()` **clears and rebuilds** the listed tabs; duplicate the sheet or export data before re-running.

---

*UI labeling (“Tetrix”) is the in-sheet name for this tooling; the business context is Deskree planning in Sheets.*
