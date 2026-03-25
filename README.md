# Deskree forecast — Google Sheets

Google Apps Script project used to **build and run financial forecasting for Deskree** inside Google Sheets. The spreadsheet model covers drivers, funding, headcount, revenue, P&L, cash flow, and scenario comparison; a custom sidebar helps load named scenarios into the single input sheet.

## What’s in this repo

| File | Role |
|------|------|
| `Setup.gs` | One-shot setup: creates tabs, layouts, and formulas for the full financial model. |
| `ScenarioSidebar.gs` | Menu entry and server functions that read/write the **Drivers** sheet from the sidebar. |
| `ScenarioSidebarView.html` | Sidebar UI (hosted as an HTML file in the Apps Script project). |

## Model overview

Running `setupFinancialModel()` in Apps Script builds these sheets:

- **Instructions** — how to use the model (inputs vs formulas).
- **Drivers** — the **only** tab where you enter numbers; everything else flows from here.
- **Funding**, **Headcount**, **Revenue**, **P&L**, **Cash flow**, **Summary** — calculated views.
- **Scenarios** — scenario framing (e.g. multiple funding paths).

The model is built around **inputs → calculations → outputs**: edit **Drivers** only so formulas stay consistent.

## Scenario sidebar

The **Tetrix** custom menu (from `ScenarioSidebar.gs`) opens **“Open Scenario Loader”**, which shows the HTML sidebar. You can:

- Apply a packaged scenario (writes meta, segments, logo ramp, headcount, and cost fields on **Drivers**).
- Inspect **current model state** via `getCurrentScenario()`-backed behavior in the UI.

Ensure **Drivers** exists (run setup first) or the sidebar will report that the tab is missing.

## Setup in Google Sheets

1. Open or create a Google Sheet for Deskree forecasting.
2. **Extensions → Apps Script** and create a project (or open an existing one).
3. Add script files matching this repo:
   - Paste or sync `Setup.gs` as a `.gs` file (e.g. `Setup`).
   - Add `ScenarioSidebar.gs`.
   - **File → Add file → HTML**, name it **`ScenarioSidebarView`** (must match `createHtmlOutputFromFile("ScenarioSidebarView")`), and paste `ScenarioSidebarView.html` contents.
4. Save the project.
5. In the script editor, select **`setupFinancialModel`** and click **Run**. Authorize the script when prompted.
6. Reload the spreadsheet; use the **Tetrix** menu to open the scenario loader when needed.

## Development notes

- This repository holds source for copy/deploy into Apps Script; there is no local `clasp` configuration in the repo by default. You can use [clasp](https://github.com/google/clasp) to push these files if you prefer CLI workflows.
- Re-running `setupFinancialModel()` **clears and rebuilds** the configured tabs; back up data or use a copy of the sheet before re-running.

---

*Internal product naming in the UI (“Tetrix”) refers to this forecast tooling; the business context is Deskree planning in Sheets.*
