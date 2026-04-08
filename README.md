# Deskree forecast — Google Sheets

Google Apps Script project used to **build and run financial forecasting for Deskree** inside Google Sheets. The spreadsheet model covers drivers, funding, headcount, revenue, P&L, cash flow, scenarios, and a **benchmark reality check**; a custom sidebar loads packaged scenarios into the input sheet.

## What’s in this repo

| File | Role |
|------|------|
| `SetupMain.gs` | `setupFinancialModel()` — orchestrates setup; logs each phase to **Logger** (Executions / View → Logs). Set `SETUP_PROGRESS_TOAST = true` for sheet toasts. |
| `ModelConstants.gs` | Shared maps: `DR`, `REVCOLS`, `PNL`, `CF`, `SUM`, etc. (used by setup + `benchmarks.gs`). |
| `SetupHelpers.gs` | Shared layout helpers (`hdr`, `colLetter`, …). |
| `SetupInstructions.gs`, `SetupDrivers.gs`, `SetupFunding.gs`, `SetupHeadcount.gs`, `SetupRevenue.gs`, `SetupPnL.gs`, `SetupCashFlow.gs`, `SetupSummary.gs`, `SetupBenchmarksTab.gs` | One tab (or tab pair) each; all called from `SetupMain.gs`. |
| `ScenarioSidebar.gs` | Custom menu, `applyScenario` / `getCurrentScenario` (**v3.1** Drivers layout, `DR`), and opens the HTML sidebar. |
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

## Scenario sidebar

The loader writes and reads the **v3.1** Drivers layout (see `ModelConstants.gs` → `DR`): **funding rounds** (up to three), **ARR targets** and horizon, **segment** economics (MM/ENT), **logo growth** (single cell `B39`; JSON still exposes a 4-slot ramp for compatibility), **maintenance ratios**, **headcount** department defaults and up to **ten named positions**, plus **marketing**, **infrastructure**, and **sales** fields. The HTML file name in Apps Script must stay **`ScenarioSidebarView`** (matching `createHtmlOutputFromFile("ScenarioSidebarView")`).

## Setup in Google Sheets

1. Open or create a Google Sheet for Deskree forecasting.
2. **Extensions → Apps Script** and create or open a project.
3. Add **all** `.gs` files from this repo (or push via clasp): `SetupMain`, `ModelConstants`, `SetupHelpers`, every `Setup*.gs` tab module, `ScenarioSidebar.gs`, and `Benchmarks.gs`.
4. **File → Add file → HTML**, name **`ScenarioSidebarView`**, paste `ScenarioSidebarView.html`.
5. Save.
6. Run **`setupFinancialModel`** once and authorize when prompted. If setup hangs, open **Executions** in the script editor and inspect **Logs** for the last completed phase.
7. Reload the spreadsheet; use **📊 Tetrix** for the scenario loader and benchmark check.

## Deploy with clasp (local → Google Sheets)

[clasp](https://github.com/google/clasp) syncs this folder to a **container-bound** Apps Script project (the script attached to your Google Sheet). Use it when you edit `.gs` / `.html` locally and want to upload without copy-paste.

### Prerequisites

- [Node.js](https://nodejs.org/) installed.
- Install clasp globally: `npm install -g @google/clasp` ([Google’s install line](https://developers.google.com/apps-script/guides/clasp) uses `npm install @google/clasp -g`).

### One-time Google login

```bash
clasp login
```

This opens a browser so clasp can act on your Google account. Use the same account that owns the spreadsheet.

### Connect this repo to your Sheet’s script

**If you already have the script** (this repo includes a `.clasp.json` with a `scriptId`):

1. Confirm `.clasp.json` points at **your** Apps Script project. The `scriptId` is the project ID from **Apps Script → Project Settings → Script ID**. If you created a new sheet/project, replace `scriptId` with yours or run `clasp clone <scriptId>` into a fresh folder and merge files.
2. From the repo root (where `.clasp.json` and `appsscript.json` live):

```bash
cd /path/to/forecast-appscript
clasp push
```

**If you are starting from scratch:**

1. In Google Sheets: **Extensions → Apps Script**, note the **Script ID** (Project Settings), or create a new spreadsheet and open its script project.
2. Either paste that ID into `.clasp.json` as `scriptId`, or clone the remote project: `clasp clone <scriptId>` (some CLI versions use `clasp clone-script`). Then align files with this repo or set `rootDir` as needed.
3. Ensure **`ScenarioSidebarView.html`** is present locally; clasp pushes all project files under `rootDir` (see `.clasp.json`).

**Optional — new Sheet + script from the CLI:** from an empty folder, run `clasp create "Your title" --type sheets` ([docs](https://developers.google.com/apps-script/guides/clasp)). That creates a spreadsheet, bound script, `.clasp.json`, and `appsscript.json`. Copy in this repo’s `.gs` / `.html` files, then `clasp push`. (Skip this if you already use the committed `.clasp.json`.)

### Enable the Apps Script API (first push)

If `clasp push` fails with an API error, turn on **Google Apps Script API** for your account: open [script.google.com/home/usersettings](https://script.google.com/home/usersettings) and enable it, then retry.

### Day-to-day commands

| Command | What it does |
|--------|----------------|
| `clasp push` | Upload local `.gs` and `.html` files to the Apps Script project (remote matches your disk). Use `clasp push --force` to skip the “remote has newer” prompt (CI and scripted deploys). |
| `clasp pull` | Download the remote project into the local folder (use when someone edited in the browser; merge carefully). |
| `clasp open-script` | Open this project in the Apps Script editor in your browser (see [Google’s clasp guide](https://developers.google.com/apps-script/guides/clasp)). Some older installs still expose `clasp open` as an alias. |
| `clasp open-container` | On newer clasp versions, may open the **parent** file (e.g. the bound spreadsheet) when the script is container-bound. If unavailable, open the Sheet from the script editor’s toolbar or Drive. |
| `clasp show-file-status` | List which local files differ from the server (some versions: `clasp status`). |

After **`clasp push`**, reload the Google Sheet so menus and the sidebar pick up changes. Run **`setupFinancialModel`** from the script editor when you need a full rebuild (it still clears/rebuilds tabs as documented above).

### Project metadata

- **`appsscript.json`** — runtime settings (e.g. `runtimeVersion: "V8"`, `timeZone`). Pushed with the project; edit locally if you need a different timezone.
- **`.clasp.json`** — `scriptId` (which Apps Script project) and `rootDir` (usually `"."` for this repo).

### Forks and copies

If you duplicate the spreadsheet or create a **new** Apps Script project, update **`scriptId`** in `.clasp.json` to that project’s ID before pushing, so you do not overwrite someone else’s script.

## Development notes

- Re-running `setupFinancialModel()` **clears and rebuilds** the listed tabs; duplicate the sheet or export data before re-running.

---

*UI labeling (“Tetrix”) is the in-sheet name for this tooling; the business context is Deskree planning in Sheets.*
