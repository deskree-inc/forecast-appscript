// ============================================================
// TETRIX FINANCIAL MODEL — GOOGLE APPS SCRIPT SETUP
// Run: setupFinancialModel() from the Apps Script editor
// ============================================================

function setupFinancialModel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const TABS = [
    { name: "📖 Instructions", color: "#FFFFFF" },
    { name: "🎛️ Drivers",     color: "#4A90D9" },
    { name: "💰 Funding",     color: "#27AE60" },
    { name: "👥 Headcount",   color: "#8E44AD" },
    { name: "📈 Revenue",     color: "#E67E22" },
    { name: "💸 P&L",         color: "#C0392B" },
    { name: "🏦 Cash Flow",   color: "#16A085" },
    { name: "📊 Summary",     color: "#2C3E50" },
    { name: "📋 Scenarios",   color: "#F39C12" },
  ];

  // Create or clear tabs
  TABS.forEach(t => {
    let sheet = ss.getSheetByName(t.name);
    if (!sheet) sheet = ss.insertSheet(t.name);
    else sheet.clearContents().clearFormats();
    sheet.setTabColor(t.color);
  });

  setupInstructions(ss);
  setupDrivers(ss);
  setupFunding(ss);
  setupHeadcount(ss);
  setupRevenue(ss);
  setupPnL(ss);
  setupSummary(ss);
  setupCashFlow(ss);
  setupScenarios(ss);

  ss.setActiveSheet(ss.getSheetByName("📖 Instructions"));
  SpreadsheetApp.getUi().alert("✅ Model built! Start in the 🎛️ Drivers tab.");
}

// ─── HELPERS ────────────────────────────────────────────────

function hdr(sheet, row, col, text, bg) {
  const cell = sheet.getRange(row, col);
  cell.setValue(text)
      .setFontWeight("bold")
      .setBackground(bg || "#2C3E50")
      .setFontColor("#FFFFFF");
}

function inp(sheet, row, col, value, format) {
  const cell = sheet.getRange(row, col);
  cell.setValue(value).setBackground("#EBF5FB").setFontColor("#1A5276");
  if (format) cell.setNumberFormat(format);
}

function label(sheet, row, col, text) {
  sheet.getRange(row, col).setValue(text).setFontWeight("bold");
}

function sectionHdr(sheet, row, text) {
  const cell = sheet.getRange(row, 1, 1, 8);
  cell.merge().setValue(text)
      .setBackground("#D6EAF8").setFontWeight("bold").setFontSize(10);
}

// ─── TAB 0: INSTRUCTIONS ────────────────────────────────────

function setupInstructions(ss) {
  const sh = ss.getSheetByName("📖 Instructions");
  sh.setColumnWidth(1, 30);   // left margin
  sh.setColumnWidth(2, 200);  // label col
  sh.setColumnWidth(3, 620);  // content col
  sh.setColumnWidth(4, 30);   // right margin

  function title(row, text) {
    sh.getRange(row, 1, 1, 4).merge()
      .setValue(text)
      .setBackground("#1A5276")
      .setFontColor("#FFFFFF")
      .setFontSize(16)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sh.setRowHeight(row, 40);
  }

  function sec(row, text) {
    sh.getRange(row, 2, 1, 2).merge()
      .setValue(text)
      .setBackground("#D6EAF8")
      .setFontWeight("bold")
      .setFontSize(11);
    sh.setRowHeight(row, 24);
  }

  function row(sh, r, label, content, bgLabel, bgContent) {
    sh.getRange(r, 2).setValue(label)
      .setBackground(bgLabel || "#F2F3F4")
      .setFontWeight("bold")
      .setVerticalAlignment("top")
      .setWrap(true);
    sh.getRange(r, 3).setValue(content)
      .setBackground(bgContent || "#FFFFFF")
      .setVerticalAlignment("top")
      .setWrap(true);
    sh.setRowHeight(r, 56);
  }

  function note(sh, r, text) {
    sh.getRange(r, 2, 1, 2).merge()
      .setValue(text)
      .setFontStyle("italic")
      .setFontColor("#717D7E")
      .setBackground("#FDFEFE")
      .setWrap(true);
    sh.setRowHeight(r, 36);
  }

  function blank(r) { sh.setRowHeight(r, 12); }

  let r = 1;

  // ── Title
  title(r, "📖 Tetrix Financial Model — How to Use"); r++;
  blank(r); r++;

  // ── Overview
  sec(r, "🗺️ Overview"); r++;
  sh.getRange(r, 2, 1, 2).merge()
    .setValue(
      "This model forecasts ARR, headcount, burn, and runway across up to 3 funding scenarios. " +
      "It follows a strict Inputs → Calculations → Outputs flow.\n\n" +
      "Rule #1: The ONLY tab you type numbers into is 🎛️ Drivers. " +
      "Every other tab is formulas only. Editing cells outside Drivers will break the model."
    )
    .setBackground("#EBF5FB")
    .setWrap(true)
    .setVerticalAlignment("top");
  sh.setRowHeight(r, 72); r++;
  blank(r); r++;

  // ── Color legend
  sec(r, "🎨 Color Legend"); r++;
  row(sh, r, "🔵 Blue cell", "Input — the only cells you should edit. Found only in 🎛️ Drivers.", "#D6EAF8", "#EBF5FB"); r++;
  row(sh, r, "⚫ Black cell", "Formula — do not edit. These auto-calculate from Drivers inputs.", "#F2F3F4", "#FFFFFF"); r++;
  row(sh, r, "🟡 Yellow cell", "Scenario toggle — the dropdown that switches between Seed / Series A scenarios.", "#FEF9E7", "#FDFEFE"); r++;
  row(sh, r, "🟢 Green cell", "Key output — ARR, ending cash, gross profit. Read-only summaries.", "#D5F5E3", "#EAFAF1"); r++;
  blank(r); r++;

  // ── Tab guide
  sec(r, "📑 Tab-by-Tab Guide"); r++;
  row(sh, r, "🎛️ Drivers", "START HERE. Set your funding scenario (dropdown), revenue assumptions, headcount triggers, and cost drivers. Every other tab reads from here.", "#D6EAF8"); r++;
  row(sh, r, "💰 Funding", "Defines the three raise scenarios (Seed $3M, Series A $10M, Series A $20M) with use-of-proceeds breakdown. The active scenario raise amount auto-populates row 12.", "#F2F3F4"); r++;
  row(sh, r, "👥 Headcount", "24-month headcount and payroll plan by department. Month 1 pulls starting HC from Drivers. Subsequent months carry forward — add hiring trigger formulas manually as needed.", "#F2F3F4"); r++;
  row(sh, r, "📈 Revenue", "ARR waterfall by segment (Enterprise, Mid-Market, SMB). Tracks new logos, new ARR, churn ARR, expansion ARR, and cumulative ARR month by month. New logo ramps come from Drivers section B2.", "#F2F3F4"); r++;
  row(sh, r, "💸 P&L", "Income statement: Revenue → COGS → Gross Profit → OpEx → EBITDA. Wire MRR (row 4, blue cells) from the Revenue tab once validated. All other rows auto-calculate.", "#F2F3F4"); r++;
  row(sh, r, "🏦 Cash Flow", "Monthly burn and runway. Runway turns red when < 6 months. Wire the blue MRR/cost cells to Revenue and P&L after validating. Capital raised auto-populates in Month 1.", "#F2F3F4"); r++;
  row(sh, r, "📊 Summary", "One-page investor KPI dashboard. All cells are formulas — it updates automatically when you change the scenario dropdown in Drivers.", "#F2F3F4"); r++;
  row(sh, r, "📋 Scenarios", "Side-by-side comparison of all three scenarios. Raise amounts are wired. Wire ARR/headcount/cash rows to Revenue and Cash Flow tabs once built out.", "#F2F3F4"); r++;
  blank(r); r++;

  // ── Step-by-step workflow
  sec(r, "🚀 Step-by-Step: First Time Setup"); r++;
  const steps = [
    ["Step 1 — Pick your scenario",
     "Go to 🎛️ Drivers. Click the yellow dropdown in cell B4 and select your active scenario (Seed $3M, Series A $10M, or Series A $20M). The raise amount in 💰 Funding row 12 will update automatically."],
    ["Step 2 — Set your ACV & segments",
     "In Drivers section B (rows 11–13), update the ACV, sales cycle, churn rate, and expansion rate for each segment to match your actual or target numbers."],
    ["Step 3 — Set your logo ramp",
     "In Drivers section B2 (rows 17–19), enter how many new logos per month you expect by segment, across four time bands (Mo 1–6, 7–12, 13–18, 19–24). These should reflect your hiring plan for sales."],
    ["Step 4 — Set headcount",
     "In Drivers section C (rows 23–26), confirm starting headcount and fully-loaded annual cost per department. Adjust Sales commission rate in cell B28."],
    ["Step 5 — Set cost drivers",
     "In Drivers section D (rows 31–34), set infra cost per customer, tooling per engineer, office/misc per employee, and marketing spend % of raise."],
    ["Step 6 — Validate Revenue tab",
     "Open 📈 Revenue and check the ARR waterfall looks reasonable. New logos should ramp per your Drivers inputs. Churn and expansion kick in from Month 2 onward."],
    ["Step 7 — Wire MRR into P&L",
     "In 💸 P&L, row 4 (MRR) contains blue input cells. For each month column, enter a formula pointing to the corresponding MRR cell in the Revenue tab summary row, e.g. ='📈 Revenue'!H5 for Month 1 total MRR."],
    ["Step 8 — Wire Cash Flow",
     "In 🏦 Cash Flow, wire the blue cells in rows 5, 7, 8, 9 (MRR collections, infra, S&M, G&A) to matching rows in P&L. Once done, the runway calculation and red alerts will be live."],
    ["Step 9 — Review Summary & Scenarios",
     "Open 📊 Summary to see your investor KPI dashboard. Then open 📋 Scenarios and wire the ARR/headcount/cash rows for all three scenarios. This becomes your investor-facing comparison view."],
    ["Step 10 — Share with investors",
     "Make a copy of the file. In the copy, hide tabs: 👥 Headcount, 📈 Revenue, 💸 P&L, 🏦 Cash Flow. Share only: 🎛️ Drivers, 📊 Summary, 📋 Scenarios with view-only access."],
  ];
  steps.forEach(([lbl, content]) => { row(sh, r, lbl, content); r++; });
  blank(r); r++;

  // ── Common mistakes
  sec(r, "⚠️ Common Mistakes to Avoid"); r++;
  const warnings = [
    ["Editing formula cells", "If a cell outside Drivers is black (no blue background), don't edit it. You'll silently break a formula with no obvious error. Use Ctrl+Z immediately if you do."],
    ["Hardcoding numbers outside Drivers", "If you type a number directly into Revenue, P&L, or Cash Flow instead of referencing Drivers, the scenario toggle will stop working for that cell."],
    ["Not wiring MRR into P&L", "Until you connect MRR row 4 in P&L to the Revenue tab, all COGS, gross margin, and EBITDA rows will show $0. This is expected — it's a deliberate manual step."],
    ["Confusing ARR and MRR", "The Revenue tab computes ARR (annual). P&L and Cash Flow operate on MRR (monthly = ARR/12). Make sure you're referencing MRR cells, not ARR, when wiring P&L row 4."],
    ["Changing forecast horizon", "If you change the forecast horizon in Drivers cell B7 from 24 to 36, the Revenue and Headcount tabs will not automatically extend — you'd need to add rows/columns manually."],
  ];
  warnings.forEach(([lbl, content]) => { row(sh, r, lbl, content, "#FDEDEC", "#FEF9E7"); r++; });
  blank(r); r++;

  // ── Key formulas
  sec(r, "🔧 Key Formula Patterns (for manual wiring)"); r++;
  const formulas = [
    ["Total MRR (Revenue tab)", "SUM of all three segment MRR rows for that month column. E.g. H28+H55+H82 where H = Month 1 column."],
    ["Cumulative ARR", "prior_month_ARR + net_new_ARR. Already built into Revenue tab column G per segment."],
    ["Runway", "ending_cash / ABS(net_burn). Already in Cash Flow row 12. Turns red when < 6."],
    ["NRR (Net Revenue Retention)", "(prior ARR - churned ARR + expansion ARR) / prior ARR. Add this to Summary once Revenue is wired."],
    ["ARR per Employee", "Total ARR / Total Headcount. Wire from Revenue summary ARR divided by sum of Headcount dept rows."],
  ];
  formulas.forEach(([lbl, content]) => { row(sh, r, lbl, content, "#EAF2FF", "#F5F8FF"); r++; });
  blank(r); r++;

  // ── AI Prompt section
  sec(r, "🤖 Generating Scenarios with AI (copy this prompt)"); r++;
  sh.getRange(r, 2, 1, 2).merge()
    .setValue(
      "Share the block below with any AI tool (Claude, ChatGPT, etc.) followed by your scenario description. " +
      "Paste the JSON output into the { } JSON tab of the Tetrix Scenario Loader sidebar."
    )
    .setBackground("#EBF5FB").setWrap(true).setVerticalAlignment("top");
  sh.setRowHeight(r, 48); r++;

  const aiPrompt =
    "Generate a Tetrix financial model scenario as a JSON object using EXACTLY this structure. " +
    "Return only raw JSON with no explanation, no markdown, no code fences.\n\n" +
    "Rules:\n" +
    "- activeScenario must be one of: \"Seed $3M\" | \"Series A $10M\" | \"Series A $20M\"\n" +
    "- closeDate must be ISO format: \"YYYY-MM-DD\"\n" +
    "- churnRate and expansionRate are decimals (e.g. 5% = 0.05)\n" +
    "- salesCommission is a decimal (e.g. 10% = 0.10)\n" +
    "- marketingPctOfRaise is a decimal (e.g. 5% = 0.05)\n" +
    "- logoRamp arrays have exactly 4 values: [mo1to6, mo7to12, mo13to18, mo19to24]\n\n" +
    "Required structure:\n" +
    "{\n" +
    "  \"meta\": {\n" +
    "    \"name\": \"string\",\n" +
    "    \"activeScenario\": \"Seed $3M\",\n" +
    "    \"closeDate\": \"2026-04-01\",\n" +
    "    \"runwayTarget\": 18,\n" +
    "    \"forecastHorizon\": 24\n" +
    "  },\n" +
    "  \"segments\": {\n" +
    "    \"enterprise\": { \"acv\": 60000, \"salesCycle\": 6, \"churnRate\": 0.05, \"expansionRate\": 0.15 },\n" +
    "    \"midMarket\":  { \"acv\": 18000, \"salesCycle\": 3, \"churnRate\": 0.08, \"expansionRate\": 0.10 },\n" +
    "    \"smb\":        { \"acv\": 4800,  \"salesCycle\": 1, \"churnRate\": 0.15, \"expansionRate\": 0.05 }\n" +
    "  },\n" +
    "  \"logoRamp\": {\n" +
    "    \"enterprise\": [0, 1, 2, 3],\n" +
    "    \"midMarket\":  [1, 2, 4, 6],\n" +
    "    \"smb\":        [2, 5, 10, 15]\n" +
    "  },\n" +
    "  \"headcount\": {\n" +
    "    \"engineering\": { \"startHC\": 4, \"hireTrigger\": 500000, \"annualCost\": 180000 },\n" +
    "    \"sales\":       { \"startHC\": 1, \"hireTrigger\": 200000, \"annualCost\": 120000 },\n" +
    "    \"csSupport\":   { \"startHC\": 1, \"hireTrigger\": 100000, \"annualCost\": 90000 },\n" +
    "    \"gAndA\":       { \"startHC\": 2, \"hireTrigger\": 500000, \"annualCost\": 100000 }\n" +
    "  },\n" +
    "  \"salesCommission\": 0.10,\n" +
    "  \"costs\": {\n" +
    "    \"infraPerCustomerPerMonth\": 50,\n" +
    "    \"toolingPerEngineerPerMonth\": 200,\n" +
    "    \"officeMiscPerEmployeePerMonth\": 100,\n" +
    "    \"marketingPctOfRaise\": 0.05\n" +
    "  }\n" +
    "}\n\n" +
    "My scenario: [DESCRIBE YOUR SCENARIO HERE]";

  sh.getRange(r, 2, 1, 2).merge()
    .setValue(aiPrompt)
    .setBackground("#F4F6F7")
    .setFontFamily("Courier New")
    .setFontSize(9)
    .setWrap(true)
    .setVerticalAlignment("top")
    .setBorder(true, true, true, true, false, false, "#AEB6BF", SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(r, 320); r++;

  note(sh, r, "Tip: you can also use the ⬆️ Export tab in the Scenario Loader sidebar to get your current model as JSON, then ask AI to modify it."); r++;
  blank(r); r++;
  note(sh, r, "Built with Apps Script. Re-run setupFinancialModel() to reset. All data will be cleared."); r++;
}

// ─── TAB 1: DRIVERS ─────────────────────────────────────────

function setupDrivers(ss) {
  const sh = ss.getSheetByName("🎛️ Drivers");
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 140);
  sh.setColumnWidth(4, 140);
  sh.setColumnWidth(5, 140);

  hdr(sh, 1, 1, "🎛️ DRIVERS — Control Panel (edit BLUE cells only)", "#1A5276");
  sh.getRange(1, 1, 1, 6).merge();

  // Section A — Funding
  sectionHdr(sh, 3, "A — Funding Scenario");
  label(sh, 4, 1, "Active Scenario");
  const scenarioCell = sh.getRange(4, 2);
  scenarioCell.setValue("Seed $3M").setBackground("#FEF9E7").setFontColor("#7D6608");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Seed $3M", "Series A $10M", "Series A $20M"], true)
    .build();
  scenarioCell.setDataValidation(rule);

  label(sh, 5, 1, "Close Date"); inp(sh, 5, 2, new Date(), "MMM YYYY");
  label(sh, 6, 1, "Runway Target (months)"); inp(sh, 6, 2, 18);
  label(sh, 7, 1, "Forecast Horizon (months)"); inp(sh, 7, 2, 24);

  // Section B — Revenue / Segments
  sectionHdr(sh, 9, "B — Revenue Drivers (by Segment)");
  ["Segment", "ACV ($)", "Sales Cycle (mo)", "Churn Rate", "Expansion Rate"]
    .forEach((h, i) => hdr(sh, 10, i + 1, h, "#1F618D"));
  [
    ["Enterprise", 60000, 6, 0.05, 0.15],
    ["Mid-Market", 18000, 3, 0.08, 0.10],
    ["SMB",         4800, 1, 0.15, 0.05],
  ].forEach((row, r) => {
    sh.getRange(11 + r, 1).setValue(row[0]).setFontWeight("bold");
    inp(sh, 11 + r, 2, row[1], "$#,##0");
    inp(sh, 11 + r, 3, row[2]);
    inp(sh, 11 + r, 4, row[3], "0%");
    inp(sh, 11 + r, 5, row[4], "0%");
  });

  // New logos ramp
  sectionHdr(sh, 15, "B2 — New Logos / Month Ramp");
  ["Segment", "Mo 1-6", "Mo 7-12", "Mo 13-18", "Mo 19-24"]
    .forEach((h, i) => hdr(sh, 16, i + 1, h, "#1F618D"));
  [
    ["Enterprise", 0, 1, 2, 3],
    ["Mid-Market", 1, 2, 4, 6],
    ["SMB",        2, 5, 10, 15],
  ].forEach((row, r) => {
    sh.getRange(17 + r, 1).setValue(row[0]).setFontWeight("bold");
    for (let c = 1; c < row.length; c++) inp(sh, 17 + r, c + 1, row[c]);
  });

  // Section C — Headcount
  sectionHdr(sh, 21, "C — Headcount Drivers");
  ["Dept", "Start HC", "Hire Trigger ($ARR)", "Fully-Loaded Cost/yr"]
    .forEach((h, i) => hdr(sh, 22, i + 1, h, "#1F618D"));
  [
    ["Engineering", 4, 500000, 180000],
    ["Sales",       1, 200000, 120000],
    ["CS/Support",  1, 100000,  90000],
    ["G&A",         2, 500000, 100000],
  ].forEach((row, r) => {
    sh.getRange(23 + r, 1).setValue(row[0]).setFontWeight("bold");
    inp(sh, 23 + r, 2, row[1]);
    inp(sh, 23 + r, 3, row[2], "$#,##0");
    inp(sh, 23 + r, 4, row[3], "$#,##0");
  });

  label(sh, 28, 1, "Sales Commission (% of ACV)"); inp(sh, 28, 2, 0.10, "0%");

  // Section D — Cost Drivers
  sectionHdr(sh, 30, "D — Cost Drivers");
  label(sh, 31, 1, "Infra cost / customer / mo"); inp(sh, 31, 2, 50, "$#,##0");
  label(sh, 32, 1, "Tooling / engineer / mo");    inp(sh, 32, 2, 200, "$#,##0");
  label(sh, 33, 1, "Office / misc / employee/mo"); inp(sh, 33, 2, 100, "$#,##0");
  label(sh, 34, 1, "Marketing spend (% of raise)");inp(sh, 34, 2, 0.05, "0%");

  // Legend
  sectionHdr(sh, 36, "Legend");
  sh.getRange(37, 1).setValue("🔵 Blue = Input (edit here)").setBackground("#EBF5FB");
  sh.getRange(37, 2).setValue("⚫ Black = Formula (do not edit)");
  sh.getRange(38, 1).setValue("🟡 Yellow = Scenario toggle").setBackground("#FEF9E7");
}

// ─── TAB 2: FUNDING ─────────────────────────────────────────

function setupFunding(ss) {
  const sh = ss.getSheetByName("💰 Funding");
  sh.setColumnWidth(1, 200);
  [2,3,4].forEach(c => sh.setColumnWidth(c, 160));

  hdr(sh, 1, 1, "💰 FUNDING — Round Scenarios", "#1A5276");
  sh.getRange(1, 1, 1, 4).merge();

  ["", "Seed $3M", "Series A $10M", "Series A $20M"]
    .forEach((h, i) => { if (i) hdr(sh, 2, i + 1, h, "#1F618D"); });

  const rows = [
    ["Raise Amount",       3000000,  10000000, 20000000, "$#,##0"],
    ["Equity Dilution",    0.15,     0.20,     0.22,     "0%"],
    ["→ Engineering %",    0.40,     0.35,     0.30,     "0%"],
    ["→ Sales & Mktg %",   0.30,     0.40,     0.45,     "0%"],
    ["→ G&A %",            0.15,     0.15,     0.15,     "0%"],
    ["→ Reserve %",        0.15,     0.10,     0.10,     "0%"],
  ];

  rows.forEach((row, r) => {
    label(sh, r + 3, 1, row[0]);
    [row[1], row[2], row[3]].forEach((v, c) => {
      inp(sh, r + 3, c + 2, v, row[4]);
    });
  });

  // Raise Amount lookup formula (used by other tabs)
  sectionHdr(sh, 11, "Active Scenario Lookup (auto)");
  label(sh, 12, 1, "Active Raise Amount");
  sh.getRange(12, 2).setFormula(
    `=INDEX(B3:D3,MATCH('🎛️ Drivers'!B4,{"Seed $3M","Series A $10M","Series A $20M"},0))`
  ).setNumberFormat("$#,##0").setBackground("#E8F8F5");
}

// ─── TAB 3: HEADCOUNT ────────────────────────────────────────

function setupHeadcount(ss) {
  const sh = ss.getSheetByName("👥 Headcount");
  sh.setColumnWidth(1, 160);

  hdr(sh, 1, 1, "👥 HEADCOUNT — Monthly Plan", "#1A5276");
  sh.getRange(1, 1, 1, 14).merge();

  const months = 24;
  const depts = ["Engineering", "Sales", "CS/Support", "G&A"];

  // Month headers
  hdr(sh, 2, 1, "Department / Month", "#1F618D");
  for (let m = 1; m <= months; m++) {
    hdr(sh, 2, m + 1, "Mo " + m, "#1F618D");
    sh.setColumnWidth(m + 1, 55);
  }

  // Starting HC from Drivers
  const startHC = [4, 1, 1, 2]; // Engineering, Sales, CS, G&A (matches Drivers rows 23-26)
  const driverRows = [23, 24, 25, 26];

  depts.forEach((dept, d) => {
    const row = d * 2 + 3;
    label(sh, row, 1, dept + " (HC)");
    label(sh, row + 1, 1, dept + " (Cost/mo)");

    for (let m = 1; m <= months; m++) {
      const col = m + 1;
      if (m === 1) {
        // Month 1 = starting HC from Drivers
        sh.getRange(row, col).setFormula(
          `='🎛️ Drivers'!B${driverRows[d]}`
        );
      } else {
        // Simple: prev month HC (users can add trigger logic manually)
        sh.getRange(row, col).setFormula(
          `=${colLetter(col - 1)}${row}`
        );
      }
      // Cost row: HC * monthly cost / 12
      sh.getRange(row + 1, col).setFormula(
        `=${colLetter(col)}${row}*'🎛️ Drivers'!D${driverRows[d]}/12`
      ).setNumberFormat("$#,##0");
    }
  });

  // Total cost row
  const totalRow = depts.length * 2 + 3;
  hdr(sh, totalRow, 1, "Total Monthly Payroll", "#2C3E50");
  for (let m = 1; m <= months; m++) {
    const col = m + 1;
    const costRows = [4, 6, 8, 10]; // rows with costs
    const sum = costRows.map(r => `${colLetter(col)}${r}`).join("+");
    sh.getRange(totalRow, col).setFormula(`=${sum}`)
      .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");
  }
}

// ─── TAB 4: REVENUE ─────────────────────────────────────────

function setupRevenue(ss) {
  const sh = ss.getSheetByName("📈 Revenue");
  sh.setColumnWidth(1, 160);

  hdr(sh, 1, 1, "📈 REVENUE — ARR Waterfall by Segment", "#1A5276");
  sh.getRange(1, 1, 1, 8).merge();

  const segments = ["Enterprise", "Mid-Market", "SMB"];
  const acvRow   = [11, 12, 13]; // Drivers tab rows for ACV
  const rampCols = [2, 3, 4, 5]; // Drivers B2 section cols for ramp periods

  // Ramp lookup helper (returns logo count for a given month)
  // Mo 1-6=col2, 7-12=col3, 13-18=col4, 19-24=col5 in Drivers rows 17-19
  const rampDriverRows = [17, 18, 19];

  let currentRow = 2;
  segments.forEach((seg, s) => {
    sectionHdr(sh, currentRow, `Segment: ${seg}`);
    currentRow++;

    ["Month", "New Logos", "New ARR", "Churn ARR", "Expansion ARR", "Net New ARR", "Cumul. ARR", "MRR"]
      .forEach((h, i) => hdr(sh, currentRow, i + 1, h, "#1F618D"));
    currentRow++;

    const dataStart = currentRow;
    for (let m = 1; m <= 24; m++) {
      const r = currentRow;
      sh.getRange(r, 1).setValue(m);

      // New logos: lookup ramp from Drivers based on month band
      const rampCol = m <= 6 ? "B" : m <= 12 ? "C" : m <= 18 ? "D" : "E";
      sh.getRange(r, 2).setFormula(
        `='🎛️ Drivers'!${rampCol}${rampDriverRows[s]}`
      );

      // New ARR = new logos * ACV
      sh.getRange(r, 3).setFormula(
        `=B${r}*'🎛️ Drivers'!B${acvRow[s]}`
      ).setNumberFormat("$#,##0");

      // Churn ARR = cumul ARR * churn rate / 12
      const churnCol = ["D", "D", "D"][s]; // col D = Churn Rate in Drivers row 11-13
      if (m === 1) {
        sh.getRange(r, 4).setValue(0).setNumberFormat("$#,##0");
      } else {
        sh.getRange(r, 4).setFormula(
          `=-G${r-1}*'🎛️ Drivers'!D${acvRow[s]}/12`
        ).setNumberFormat("$#,##0");
      }

      // Expansion ARR = cumul ARR * expansion rate / 12
      if (m === 1) {
        sh.getRange(r, 5).setValue(0).setNumberFormat("$#,##0");
      } else {
        sh.getRange(r, 5).setFormula(
          `=G${r-1}*'🎛️ Drivers'!E${acvRow[s]}/12`
        ).setNumberFormat("$#,##0");
      }

      // Net New ARR
      sh.getRange(r, 6).setFormula(`=C${r}+D${r}+E${r}`).setNumberFormat("$#,##0");

      // Cumulative ARR
      if (m === 1) {
        sh.getRange(r, 7).setFormula(`=F${r}`).setNumberFormat("$#,##0");
      } else {
        sh.getRange(r, 7).setFormula(`=G${r-1}+F${r}`).setNumberFormat("$#,##0");
      }

      // MRR
      sh.getRange(r, 8).setFormula(`=G${r}/12`).setNumberFormat("$#,##0");

      currentRow++;
    }
    currentRow += 2;
  });

  // Summary rows
  sectionHdr(sh, currentRow, "📊 Total ARR Summary (all segments)");
  currentRow++;
  ["Month", "Total ARR", "Total MRR", "Net New ARR"].forEach((h, i) =>
    hdr(sh, currentRow, i + 1, h, "#2C3E50")
  );
  currentRow++;

  // The 3 segments' Cumul ARR rows are at dataStart+0..23, dataStart+27..50, dataStart+54..77 (approx)
  // Easier: just reference G column of each segment's block
  // We'll output a note instead since row positions depend on gaps
  sh.getRange(currentRow, 1).setValue("→ Reference cumulative ARR rows per segment above to build totals.")
    .setFontStyle("italic").setFontColor("#888888");
  sh.getRange(currentRow, 1, 1, 4).merge();
}

// ─── TAB 5: P&L ─────────────────────────────────────────────

function setupPnL(ss) {
  const sh = ss.getSheetByName("💸 P&L");
  sh.setColumnWidth(1, 220);
  const months = 24;
  for (let m = 1; m <= months; m++) sh.setColumnWidth(m + 1, 80);

  hdr(sh, 1, 1, "💸 P&L — Income Statement (Monthly)", "#1A5276");
  sh.getRange(1, 1, 1, months + 1).merge();

  // Month headers
  hdr(sh, 2, 1, "Line Item", "#1F618D");
  for (let m = 1; m <= months; m++) hdr(sh, 2, m + 1, "Mo " + m, "#1F618D");

  const fmt = "$#,##0";
  const pct = "0.0%";

  // Revenue section
  sectionHdr(sh, 3, "REVENUE");
  label(sh, 4, 1, "MRR");
  label(sh, 5, 1, "ARR");
  label(sh, 6, 1, "YoY Growth");

  // COGS section
  sectionHdr(sh, 8, "COST OF REVENUE (COGS)");
  label(sh, 9,  1, "Infrastructure");
  label(sh, 10, 1, "Customer Success");
  label(sh, 11, 1, "Total COGS");
  label(sh, 12, 1, "Gross Profit");
  label(sh, 13, 1, "Gross Margin %");

  // OpEx section
  sectionHdr(sh, 15, "OPERATING EXPENSES");
  label(sh, 16, 1, "Engineering Payroll");
  label(sh, 17, 1, "Sales Payroll");
  label(sh, 18, 1, "G&A Payroll");
  label(sh, 19, 1, "Sales Commission");
  label(sh, 20, 1, "Marketing");
  label(sh, 21, 1, "Tooling / Misc");
  label(sh, 22, 1, "Total OpEx");

  // Bottom line
  sectionHdr(sh, 24, "BOTTOM LINE");
  label(sh, 25, 1, "EBITDA");
  label(sh, 26, 1, "EBITDA Margin %");
  label(sh, 27, 1, "Cumulative Burn");

  for (let m = 1; m <= months; m++) {
    const col = m + 1;
    const C = colLetter(col);

    // Revenue — placeholders pointing to Revenue tab (users wire MRR once Revenue tab is validated)
    sh.getRange(4, col).setValue(0).setNumberFormat(fmt).setBackground("#EBF5FB").setFontColor("#1A5276");
    sh.getRange(5, col).setFormula(`=${C}4*12`).setNumberFormat(fmt);
    sh.getRange(6, col).setFormula(
      m <= 12 ? `=""` : `=IFERROR(${C}5/${colLetter(col-12)}5-1,"")`
    ).setNumberFormat(pct);

    // COGS
    // Infra = customers * cost per customer (placeholder: MRR/ACV_blended * infra rate)
    sh.getRange(9, col).setFormula(`=${C}4/'🎛️ Drivers'!B11*'🎛️ Drivers'!B31`).setNumberFormat(fmt);
    // CS payroll from Headcount row 7 (CS/Support cost row)
    sh.getRange(10, col).setFormula(`='👥 Headcount'!${C}8`).setNumberFormat(fmt);
    sh.getRange(11, col).setFormula(`=${C}9+${C}10`).setNumberFormat(fmt).setFontWeight("bold");
    sh.getRange(12, col).setFormula(`=${C}4-${C}11`).setNumberFormat(fmt).setFontWeight("bold").setBackground("#D5F5E3");
    sh.getRange(13, col).setFormula(`=IFERROR(${C}12/${C}4,0)`).setNumberFormat(pct).setFontWeight("bold");

    // OpEx — pull from Headcount tab cost rows
    sh.getRange(16, col).setFormula(`='👥 Headcount'!${C}4`).setNumberFormat(fmt);  // Eng cost row
    sh.getRange(17, col).setFormula(`='👥 Headcount'!${C}6`).setNumberFormat(fmt);  // Sales cost row
    sh.getRange(18, col).setFormula(`='👥 Headcount'!${C}10`).setNumberFormat(fmt); // G&A cost row
    // Commission = new logos * blended ACV * commission rate (approximate via MRR delta)
    sh.getRange(19, col).setFormula(
      `=MAX(0,${C}4-${m > 1 ? colLetter(col-1)+"4" : "0"})*12*'🎛️ Drivers'!B28`
    ).setNumberFormat(fmt);
    // Marketing = % of raise / 12
    sh.getRange(20, col).setFormula(`='💰 Funding'!B12*'🎛️ Drivers'!B34/3/4`).setNumberFormat(fmt);
    // Tooling = engineers * tooling cost
    sh.getRange(21, col).setFormula(`='👥 Headcount'!${C}3*'🎛️ Drivers'!B32+'👥 Headcount'!${C}11*'🎛️ Drivers'!B33`).setNumberFormat(fmt);
    sh.getRange(22, col).setFormula(`=SUM(${C}16:${C}21)`).setNumberFormat(fmt).setFontWeight("bold");

    // Bottom line
    sh.getRange(25, col).setFormula(`=${C}12-${C}22`).setNumberFormat(fmt).setFontWeight("bold");
    sh.getRange(26, col).setFormula(`=IFERROR(${C}25/${C}4,0)`).setNumberFormat(pct);
    sh.getRange(27, col).setFormula(
      m === 1 ? `=${C}25` : `=${colLetter(col-1)}27+${C}25`
    ).setNumberFormat(fmt);

    // Highlight negative EBITDA red
    const ebitdaCell = sh.getRange(25, col);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor("#922B21")
      .setRanges([ebitdaCell])
      .build();
    const rules = sh.getConditionalFormatRules();
    rules.push(rule);
    sh.setConditionalFormatRules(rules);
  }

  sh.getRange(28, 1).setValue("🔵 Wire MRR (row 4) to the Revenue tab once validated.")
    .setFontStyle("italic").setFontColor("#888888");
  sh.getRange(28, 1, 1, 6).merge();
}

// ─── TAB 7: SUMMARY ─────────────────────────────────────────

function setupSummary(ss) {
  const sh = ss.getSheetByName("📊 Summary");
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(4, 160);

  hdr(sh, 1, 1, "📊 SUMMARY — Investor KPI Dashboard", "#1A5276");
  sh.getRange(1, 1, 1, 4).merge();

  // Active scenario banner
  label(sh, 2, 1, "Active Scenario:");
  sh.getRange(2, 2).setFormula(`='🎛️ Drivers'!B4`)
    .setBackground("#FEF9E7").setFontWeight("bold").setFontColor("#7D6608");
  sh.getRange(2, 3).setFormula(`="Raise: "&TEXT('💰 Funding'!B12,"$#,##0")`)
    .setBackground("#FEF9E7").setFontColor("#7D6608");

  // KPI sections
  sectionHdr(sh, 4, "📈 Revenue KPIs");
  [
    ["Current MRR",          `='💸 P&L'!B4`,          "$#,##0"],
    ["Current ARR",          `='💸 P&L'!B5`,           "$#,##0"],
    ["ARR at Month 12",      `='💸 P&L'!N5`,           "$#,##0"],
    ["ARR at Month 24",      `='💸 P&L'!Z5`,           "$#,##0"],
    ["Gross Margin (Mo 1)",  `='💸 P&L'!B13`,          "0%"],
  ].forEach(([lbl, formula, fmt], i) => {
    label(sh, 5 + i, 1, lbl);
    sh.getRange(5 + i, 2).setFormula(formula).setNumberFormat(fmt).setBackground("#E8F8F5");
  });

  sectionHdr(sh, 11, "🔥 Burn & Runway KPIs");
  [
    ["Capital Raised",         `='💰 Funding'!B12`,      "$#,##0"],
    ["Mo 1 Net Burn",          `='🏦 Cash Flow'!B10`,    "$#,##0"],
    ["Runway (months, Mo 1)",  `='🏦 Cash Flow'!B12`,    "0.0"],
    ["Cash at Month 12",       `='🏦 Cash Flow'!N11`,    "$#,##0"],
    ["Cash at Month 24",       `='🏦 Cash Flow'!Z11`,    "$#,##0"],
  ].forEach(([lbl, formula, fmt], i) => {
    label(sh, 12 + i, 1, lbl);
    sh.getRange(12 + i, 2).setFormula(formula).setNumberFormat(fmt).setBackground("#FEF9E7");
  });

  sectionHdr(sh, 18, "👥 Team KPIs");
  [
    ["Total HC (Month 1)",   `='👥 Headcount'!B11`,     "0"],
    ["Total HC (Month 12)",  `='👥 Headcount'!N11`,     "0"],
    ["Total HC (Month 24)",  `='👥 Headcount'!Z11`,     "0"],
    ["Mo 1 Payroll",         `='👥 Headcount'!B11`,     "$#,##0"],
  ].forEach(([lbl, formula, fmt], i) => {
    label(sh, 19 + i, 1, lbl);
    sh.getRange(19 + i, 2).setFormula(formula).setNumberFormat(fmt).setBackground("#EBF5FB");
  });

  // Fix HC rows — Headcount tab row 11 is total payroll, not headcount
  // HC rows are 3,5,7,9 (one per dept). Sum those for total HC.
  // Override the HC formulas with correct ones
  ["B","N","Z"].forEach((col, i) => {
    sh.getRange(19 + i, 2).setFormula(
      `='👥 Headcount'!${col}3+'👥 Headcount'!${col}5+'👥 Headcount'!${col}7+'👥 Headcount'!${col}9`
    ).setNumberFormat("0").setBackground("#EBF5FB");
  });
  sh.getRange(22, 2).setFormula(`='👥 Headcount'!B11`).setNumberFormat("$#,##0").setBackground("#EBF5FB");

  sectionHdr(sh, 24, "⚙️ Model Assumptions (reference)");
  [
    ["Enterprise ACV",      `='🎛️ Drivers'!B11`, "$#,##0"],
    ["Mid-Market ACV",      `='🎛️ Drivers'!B12`, "$#,##0"],
    ["SMB ACV",             `='🎛️ Drivers'!B13`, "$#,##0"],
    ["Runway Target",       `='🎛️ Drivers'!B6`,  "0 \"months\""],
    ["Forecast Horizon",    `='🎛️ Drivers'!B7`,  "0 \"months\""],
  ].forEach(([lbl, formula, fmt], i) => {
    label(sh, 25 + i, 1, lbl);
    sh.getRange(25 + i, 2).setFormula(formula).setNumberFormat(fmt);
  });

  sh.getRange(31, 1).setValue("All cells are formulas — change scenario in 🎛️ Drivers to refresh.")
    .setFontStyle("italic").setFontColor("#888888");
  sh.getRange(31, 1, 1, 4).merge();
}

// ─── TAB 6: CASH FLOW ───────────────────────────────────────

function setupCashFlow(ss) {
  const sh = ss.getSheetByName("🏦 Cash Flow");
  sh.setColumnWidth(1, 200);

  hdr(sh, 1, 1, "🏦 CASH FLOW — Monthly Burn & Runway", "#1A5276");
  sh.getRange(1, 1, 1, 13).merge();

  const months = 24;
  const headers = ["Line Item", ...Array.from({length: months}, (_, i) => "Mo " + (i+1))];
  headers.forEach((h, i) => hdr(sh, 2, i + 1, h, "#1F618D"));
  for (let m = 1; m <= months; m++) sh.setColumnWidth(m + 1, 75);

  const lines = [
    "Beginning Cash",
    "+ Capital Raised",
    "+ Cash Collections (MRR)",
    "- Payroll",
    "- Infra / COGS",
    "- Sales & Marketing",
    "- G&A Costs",
    "= Net Burn",
    "= Ending Cash",
    "Runway Remaining (mo)",
  ];

  lines.forEach((line, i) => {
    label(sh, i + 3, 1, line);
  });

  // Ending Cash row = row 11, Beginning Cash = row 3
  // Month 1: Beginning Cash = Capital Raised (from Funding!B12)
  sh.getRange(3, 2).setFormula(`='💰 Funding'!B12`).setNumberFormat("$#,##0");

  for (let m = 1; m <= months; m++) {
    const col = m + 1;
    const C = colLetter(col);

    if (m > 1) {
      // Beginning cash = prior ending cash
      sh.getRange(3, col).setFormula(`=${colLetter(col-1)}11`).setNumberFormat("$#,##0");
    }

    // Capital raised: only month 1
    sh.getRange(4, col).setValue(m === 1 ? `='💰 Funding'!B12` : 0).setNumberFormat("$#,##0");
    if (m === 1) sh.getRange(4, col).setFormula(`='💰 Funding'!B12`);

    // Cash collections — placeholder (link to Revenue tab MRR when built)
    sh.getRange(5, col).setValue(0).setNumberFormat("$#,##0").setBackground("#EBF5FB").setFontColor("#1A5276");

    // Payroll from Headcount tab row 11 (total payroll row)
    sh.getRange(6, col).setFormula(`=-'👥 Headcount'!${C}11`).setNumberFormat("$#,##0");

    // Infra placeholder
    sh.getRange(7, col).setValue(0).setNumberFormat("$#,##0").setBackground("#EBF5FB").setFontColor("#1A5276");

    // S&M placeholder
    sh.getRange(8, col).setValue(0).setNumberFormat("$#,##0").setBackground("#EBF5FB").setFontColor("#1A5276");

    // G&A placeholder
    sh.getRange(9, col).setValue(0).setNumberFormat("$#,##0").setBackground("#EBF5FB").setFontColor("#1A5276");

    // Net Burn
    sh.getRange(10, col).setFormula(`=SUM(${C}4:${C}9)`).setNumberFormat("$#,##0").setFontWeight("bold");

    // Ending Cash
    sh.getRange(11, col).setFormula(`=${C}3+${C}10`).setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");

    // Runway remaining = ending cash / avg monthly burn (rough)
    sh.getRange(12, col).setFormula(`=IFERROR(${C}11/ABS(${C}10),0)`).setNumberFormat("0.0");
  }

  // Conditional formatting: runway < 6 = red
  const runwayRange = sh.getRange(12, 2, 1, months);
  const cfRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(6)
    .setBackground("#FADBD8")
    .setFontColor("#922B21")
    .setRanges([runwayRange])
    .build();
  const cfRules = sh.getConditionalFormatRules();
  cfRules.push(cfRule);
  sh.setConditionalFormatRules(cfRules);

  // Note on blue cells
  sh.getRange(14, 1).setValue("🔵 Blue cells = connect to Revenue tab once built (MRR, Infra, S&M, G&A)")
    .setFontStyle("italic").setFontColor("#888888");
  sh.getRange(14, 1, 1, 6).merge();
}

// ─── TAB 8: SCENARIOS ───────────────────────────────────────

function setupScenarios(ss) {
  const sh = ss.getSheetByName("📋 Scenarios");
  sh.setColumnWidth(1, 220);
  [2,3,4].forEach(c => sh.setColumnWidth(c, 160));

  hdr(sh, 1, 1, "📋 SCENARIOS — Investor Comparison View", "#1A5276");
  sh.getRange(1, 1, 1, 4).merge();

  ["Metric", "Seed $3M", "Series A $10M", "Series A $20M"]
    .forEach((h, i) => hdr(sh, 2, i + 1, h, "#1F618D"));

  const metrics = [
    "Raise Amount",
    "ARR at Month 12",
    "ARR at Month 24",
    "Customers at Mo 24",
    "Headcount at Mo 24",
    "Gross Margin",
    "Runway (months)",
    "ARR / Employee",
  ];

  metrics.forEach((m, i) => {
    label(sh, i + 3, 1, m);
    // Placeholders — wire up after Revenue/Headcount tabs are complete
    for (let c = 2; c <= 4; c++) {
      sh.getRange(i + 3, c).setValue("→ TBD").setFontColor("#AAAAAA");
    }
  });

  // Raise amount auto-fill from Funding tab
  sh.getRange(3, 2).setFormula(`='💰 Funding'!B3`).setNumberFormat("$#,##0").setFontColor("#000000");
  sh.getRange(3, 3).setFormula(`='💰 Funding'!C3`).setNumberFormat("$#,##0").setFontColor("#000000");
  sh.getRange(3, 4).setFormula(`='💰 Funding'!D3`).setNumberFormat("$#,##0").setFontColor("#000000");

  sh.getRange(12, 1).setValue("→ Wire remaining rows to Revenue + Cash Flow tabs after building them out.")
    .setFontStyle("italic").setFontColor("#888888");
  sh.getRange(12, 1, 1, 4).merge();
}

// ─── UTILITY ─────────────────────────────────────────────────

function colLetter(col) {
  let letter = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}