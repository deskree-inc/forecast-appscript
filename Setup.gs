// ============================================================
// TETRIX FINANCIAL MODEL — SETUP v2
// Run: setupFinancialModel()
// ============================================================

// ─── DRIVERS ROW MAP (update here if layout shifts) ─────────
const DR = {
  // Funding rounds: rows 5–9, cols: 1=Name 2=Amount 3=Date 4=ARR 5=Notes
  ROUND_START: 5,
  // ARR Targets
  TARGET_ARR: "B12", MOM_GROWTH: "B13", HORIZON: "B14",
  // ICP Segments: row 18=MM, 19=ENT | cols: 2=BegACV 3=ExpACV 4=Churn 5=Exp% 6=ExpMo 7=CAC 8=LeadTime 9=CloseRate
  MM_ROW: 18, ENT_ROW: 19,
  // Logo ramp: row 23=MM, 24=ENT | cols B–E = 4 time bands
  MM_RAMP: 23, ENT_RAMP: 24,
  // Maintenance ratios: col 2 = max accounts/person
  AE_RATIO: "B28", FDE_RATIO: "B29", CSM_RATIO: "B30",
  // Dept defaults: rows 34–37 | cols: 2=StartHC 3=Salary 4=SW 5=HW 6=Insurance
  ENG: 34, SALES: 35, CS: 36, GA: 37,
  // Positions table: rows 41–50
  // Marketing
  EVENTS: "B53", DIGITAL: "B54",
  // Infrastructure
  INFRA: "B58", TOOLING: "B59",
  // Sales
  COMMISSION: "B62", ACCELERATOR: "B63",
};

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
    { name: "🚦 Benchmarks",  color: "#E74C3C" },
  ];
  TABS.forEach(t => {
    let s = ss.getSheetByName(t.name);
    if (!s) s = ss.insertSheet(t.name);
    else s.clearContents().clearFormats();
    s.setTabColor(t.color);
  });
  setupInstructions(ss);
  setupDrivers(ss);
  setupFunding(ss);
  setupHeadcount(ss);
  setupRevenue(ss);
  setupPnL(ss);
  setupCashFlow(ss);
  setupSummary(ss);
  setupScenarios(ss);
  setupBenchmarks(ss);
  ss.setActiveSheet(ss.getSheetByName("📖 Instructions"));
  SpreadsheetApp.getUi().alert("✅ Model built! Start in the 🎛️ Drivers tab.");
}

// ─── HELPERS ────────────────────────────────────────────────
function hdr(sheet, row, col, text, bg) {
  sheet.getRange(row, col).setValue(text).setFontWeight("bold")
    .setBackground(bg || "#2C3E50").setFontColor("#FFFFFF");
}
function inp(sheet, row, col, value, format) {
  const c = sheet.getRange(row, col);
  c.setValue(value).setBackground("#EBF5FB").setFontColor("#1A5276");
  if (format) c.setNumberFormat(format);
}
function label(sheet, row, col, text) {
  sheet.getRange(row, col).setValue(text).setFontWeight("bold");
}
function sectionHdr(sheet, row, text) {
  sheet.getRange(row, 1, 1, 9).merge().setValue(text)
    .setBackground("#D6EAF8").setFontWeight("bold").setFontSize(10);
}
function colLetter(col) {
  let l = "";
  while (col > 0) { const r = (col-1)%26; l = String.fromCharCode(65+r)+l; col = Math.floor((col-1)/26); }
  return l;
}

// ─── INSTRUCTIONS ───────────────────────────────────────────
function setupInstructions(ss) {
  const sh = ss.getSheetByName("📖 Instructions");
  sh.setColumnWidth(1, 30); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 620); sh.setColumnWidth(4, 30);

  function title(row, text) {
    sh.getRange(row,1,1,4).merge().setValue(text).setBackground("#1A5276").setFontColor("#FFFFFF")
      .setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sh.setRowHeight(row, 40);
  }
  function sec(row, text) {
    sh.getRange(row,2,1,2).merge().setValue(text).setBackground("#D6EAF8").setFontWeight("bold").setFontSize(11);
    sh.setRowHeight(row, 24);
  }
  function row(r, lbl, content, bgL, bgC) {
    sh.getRange(r,2).setValue(lbl).setBackground(bgL||"#F2F3F4").setFontWeight("bold").setVerticalAlignment("top").setWrap(true);
    sh.getRange(r,3).setValue(content).setBackground(bgC||"#FFFFFF").setVerticalAlignment("top").setWrap(true);
    sh.setRowHeight(r, 56);
  }
  function note(r, text) {
    sh.getRange(r,2,1,2).merge().setValue(text).setFontStyle("italic").setFontColor("#717D7E").setBackground("#FDFEFE").setWrap(true);
    sh.setRowHeight(r, 36);
  }
  function blank(r) { sh.setRowHeight(r, 12); }

  let r = 1;
  title(r, "📖 Tetrix Financial Model — How to Use"); r++;
  blank(r); r++;

  sec(r, "🗺️ Overview"); r++;
  sh.getRange(r,2,1,2).merge().setValue(
    "Inputs → Calculations → Outputs. The ONLY tab you type numbers into is 🎛️ Drivers. " +
    "Every other tab is formulas only.\n\nNew in v2: multiple funding rounds, Mid-Market & Enterprise segments only (with beginning + expanded ACV, CAC, lead time, close rate), " +
    "individual headcount positions, events marketing budget, and a 🚦 Benchmarks tab that acts as a reality check after every scenario load."
  ).setBackground("#EBF5FB").setWrap(true).setVerticalAlignment("top");
  sh.setRowHeight(r, 80); r++;
  blank(r); r++;

  sec(r, "🎨 Color Legend"); r++;
  row(r, "🔵 Blue", "Input — only cells you edit. Found only in 🎛️ Drivers.", "#D6EAF8", "#EBF5FB"); r++;
  row(r, "⚫ Black", "Formula — do not edit."); r++;
  row(r, "🟢 Green", "Key output — ARR, cash, gross profit."); r++;
  blank(r); r++;

  sec(r, "📑 Tab Guide"); r++;
  row(r, "🎛️ Drivers",    "START HERE. All inputs: funding rounds, ARR targets, ICP segments (MM + Enterprise), headcount, marketing, costs."); r++;
  row(r, "💰 Funding",    "Three raise scenarios with use-of-proceeds. Active scenario raise auto-populates."); r++;
  row(r, "👥 Headcount",  "24-month headcount + payroll plan. Month 1 pulls from Drivers dept defaults."); r++;
  row(r, "📈 Revenue",    "ARR waterfall by segment (MM + Enterprise). New logos, churn, expansion, cumulative ARR."); r++;
  row(r, "💸 P&L",        "Revenue → COGS → Gross Profit → OpEx → EBITDA. Wire MRR row 4 from Revenue tab."); r++;
  row(r, "🏦 Cash Flow",  "Burn + runway. Red when < 6 months. Wire blue cells from P&L."); r++;
  row(r, "📊 Summary",    "Investor KPI dashboard. All formulas, updates on scenario change."); r++;
  row(r, "📋 Scenarios",  "Side-by-side scenario comparison. Wire ARR/headcount rows once Revenue is built."); r++;
  row(r, "🚦 Benchmarks", "Reality check. Run 📊 Tetrix → Check Benchmarks after each scenario load to flag CAC payback, LTV:CAC, NRR, coverage ratios, and growth vs Bessemer heuristics.", "#FADBD8", "#FEF9E7"); r++;
  blank(r); r++;

  sec(r, "🚀 Step-by-Step Setup"); r++;
  [
    ["Step 1 — Funding rounds",   "Section A in Drivers. Add up to 5 rounds with amount, close date, and expected ARR at close."],
    ["Step 2 — ARR targets",      "Section B. Set your target ARR, MoM growth rate, and forecast horizon."],
    ["Step 3 — ICP segments",     "Section C. Set beginning ACV, expanded ACV, churn, expansion %, CAC, lead time, and close rate for Mid-Market and Enterprise."],
    ["Step 4 — Logo ramp",        "Section C2. Set new logos/month by time band. Should be consistent with close rate × sales capacity."],
    ["Step 5 — Ratios",           "Section D. Set max accounts per AE, FDE, and CSM. Benchmarks will flag if headcount is insufficient."],
    ["Step 6 — Headcount",        "Section E. Set dept defaults. Optionally add individual positions with start dates in E2."],
    ["Step 7 — Marketing",        "Section F. Set events and digital budgets separately (events are a key growth driver)."],
    ["Step 8 — Infrastructure",   "Section G. Set infra cost per customer and tooling cost per engineer."],
    ["Step 9 — Validate Revenue", "Open 📈 Revenue. Check ARR waterfall looks right per segment."],
    ["Step 10 — Wire MRR",        "In 💸 P&L row 4, wire each month to the corresponding MRR cell in Revenue. This drives all P&L and Cash Flow calculations."],
    ["Step 11 — Run Benchmarks",  "📊 Tetrix → Check Benchmarks. Review every flagged metric before sharing with investors."],
    ["Step 12 — Share",           "Make a copy, hide Headcount/Revenue/P&L/CashFlow tabs, share Summary + Scenarios + Benchmarks view-only."],
  ].forEach(([lbl, content]) => { row(r, lbl, content); r++; });
  blank(r); r++;

  sec(r, "⚠️ Common Mistakes"); r++;
  [
    ["Editing formula cells",       "Black cells outside Drivers are formulas. Ctrl+Z immediately if edited accidentally."],
    ["Not wiring MRR",              "Until P&L row 4 is connected to Revenue, COGS/EBITDA/CashFlow all show $0."],
    ["ARR vs MRR confusion",        "Revenue tab computes ARR. P&L/CashFlow use MRR = ARR/12. Reference MRR cells, not ARR."],
    ["Logo ramp vs close rate",     "C2 logo ramp is manual. Make sure it's consistent with your close rate × SDR capacity from section C."],
    ["Ignoring Benchmarks",         "The 🚦 Benchmarks tab won't auto-update — run it manually after each scenario load."],
  ].forEach(([lbl, content]) => { row(r, lbl, content, "#FDEDEC", "#FEF9E7"); r++; });
  blank(r); r++;

  sec(r, "🤖 AI Scenario Prompt (copy → paste into Claude or ChatGPT)"); r++;
  sh.getRange(r,2,1,2).merge().setValue(
    "Copy the block below into any AI tool and replace the last line with your scenario description. Paste the JSON into the { } JSON tab of the Tetrix Scenario Loader sidebar."
  ).setBackground("#EBF5FB").setWrap(true).setVerticalAlignment("top");
  sh.setRowHeight(r, 48); r++;

  const prompt =
    "Generate a Tetrix v2 financial model scenario as a JSON object using EXACTLY this structure.\n" +
    "Return only raw JSON — no explanation, no markdown, no code fences.\n\n" +
    "Rules:\n" +
    "- fundingRounds: array of up to 5 objects; omit empty rounds\n" +
    "- segments: midMarket and enterprise only (no SMB)\n" +
    "- churnRate, expansionRate, closeRate, commission, accelerator: decimals (5% = 0.05)\n" +
    "- logoRamp arrays: exactly 4 values [mo1to6, mo7to12, mo13to18, mo19to24]\n" +
    "- positions: optional array of individual hires; omit if not needed\n\n" +
    "{\n" +
    "  \"meta\": { \"name\": \"string\", \"forecastHorizon\": 24 },\n" +
    "  \"fundingRounds\": [\n" +
    "    { \"name\": \"Seed\", \"amount\": 3000000, \"closeDate\": \"2026-04-01\", \"expectedARR\": 200000, \"notes\": \"string\" },\n" +
    "    { \"name\": \"Series A\", \"amount\": 10000000, \"closeDate\": \"2027-06-01\", \"expectedARR\": 1500000, \"notes\": \"string\" }\n" +
    "  ],\n" +
    "  \"arrTargets\": { \"targetARR\": 2000000, \"momGrowthRate\": 0.20 },\n" +
    "  \"segments\": {\n" +
    "    \"midMarket\":  { \"begACV\": 80000, \"expACV\": 150000, \"churnRate\": 0.08, \"expansionRate\": 0.30, \"expansionMonth\": 12, \"cac\": 15000, \"leadTime\": 3, \"closeRate\": 0.20 },\n" +
    "    \"enterprise\": { \"begACV\": 250000, \"expACV\": 1000000, \"churnRate\": 0.05, \"expansionRate\": 0.40, \"expansionMonth\": 18, \"cac\": 50000, \"leadTime\": 6, \"closeRate\": 0.10 }\n" +
    "  },\n" +
    "  \"logoRamp\": { \"midMarket\": [1,2,4,6], \"enterprise\": [0,1,2,3] },\n" +
    "  \"maintenanceRatios\": { \"aePerAccounts\": 15, \"fdePerAccounts\": 10, \"csmPerAccounts\": 20 },\n" +
    "  \"headcount\": {\n" +
    "    \"deptDefaults\": {\n" +
    "      \"engineering\": { \"startHC\": 4, \"annualSalary\": 180000, \"swCostPerMo\": 300, \"hwCostOneTime\": 2500, \"insurancePerMo\": 200 },\n" +
    "      \"sales\":       { \"startHC\": 2, \"annualSalary\": 120000, \"swCostPerMo\": 150, \"hwCostOneTime\": 1500, \"insurancePerMo\": 200 },\n" +
    "      \"csSupport\":   { \"startHC\": 1, \"annualSalary\":  90000, \"swCostPerMo\": 100, \"hwCostOneTime\": 1500, \"insurancePerMo\": 200 },\n" +
    "      \"gAndA\":       { \"startHC\": 2, \"annualSalary\": 100000, \"swCostPerMo\":  80, \"hwCostOneTime\": 1000, \"insurancePerMo\": 200 }\n" +
    "    },\n" +
    "    \"positions\": [\n" +
    "      { \"title\": \"Sr. Engineer\", \"dept\": \"Engineering\", \"startDate\": \"2026-07-01\", \"annualSalary\": 200000, \"swCostPerMo\": 350 }\n" +
    "    ]\n" +
    "  },\n" +
    "  \"marketing\": { \"eventsAnnual\": 50000, \"digitalAnnual\": 30000 },\n" +
    "  \"infrastructure\": { \"infraPerCustomerPerMo\": 50, \"toolingPerEngineerPerMo\": 200 },\n" +
    "  \"sales\": { \"commission\": 0.10, \"accelerator\": 0.15 }\n" +
    "}\n\n" +
    "My scenario: [DESCRIBE YOUR SCENARIO HERE]";

  sh.getRange(r,2,1,2).merge().setValue(prompt)
    .setBackground("#F4F6F7").setFontFamily("Courier New").setFontSize(9)
    .setWrap(true).setVerticalAlignment("top")
    .setBorder(true,true,true,true,false,false,"#AEB6BF", SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(r, 380); r++;

  note(r, "Tip: use ⬆️ Export tab in the Scenario Loader sidebar to get the current model as JSON, then ask AI to modify it."); r++;
  blank(r); r++;
  note(r, "Built with Apps Script. Re-run setupFinancialModel() to reset. All data will be cleared."); r++;
}

// ─── DRIVERS ────────────────────────────────────────────────
function setupDrivers(ss) {
  const sh = ss.getSheetByName("🎛️ Drivers");
  sh.setColumnWidth(1, 220);
  [130,130,120,120,110,120,110,110].forEach((w,i) => sh.setColumnWidth(i+2, w));

  hdr(sh,1,1,"🎛️ DRIVERS — Control Panel (edit BLUE cells only)","#1A5276");
  sh.getRange(1,1,1,9).merge();

  // A: Funding Rounds (rows 3–9)
  sectionHdr(sh,3,"A — Funding Rounds (up to 5 within forecast period)");
  ["Round","Amount ($)","Close Date","ARR at Close ($)","Notes / Milestone"]
    .forEach((h,i) => hdr(sh,4,i+1,h,"#1F618D"));
  [
    ["Seed",     3000000, new Date("2026-04-01"),  200000, "Initial traction, product-market fit"],
    ["Series A", 10000000,new Date("2027-06-01"), 1500000, "$1.5M ARR, 3x YoY — Bessemer Series A bar"],
    ["","","","",""],["","","","",""],["","","","",""],
  ].forEach((round,i) => {
    const r = 5+i;
    inp(sh,r,1,round[0]);
    if (round[1]) { inp(sh,r,2,round[1],"$#,##0"); inp(sh,r,3,round[2],"MMM YYYY"); inp(sh,r,4,round[3],"$#,##0"); }
    else { [2,3,4].forEach(c => sh.getRange(r,c).setBackground("#EBF5FB")); }
    inp(sh,r,5,round[4]);
  });

  // B: ARR Targets (rows 11–14)
  sectionHdr(sh,11,"B — ARR Targets");
  label(sh,12,1,"Target ARR (end of forecast)"); inp(sh,12,2,2000000,"$#,##0");
  label(sh,13,1,"Target MoM Growth Rate");       inp(sh,13,2,0.20,"0%");
  label(sh,14,1,"Forecast Horizon (months)");    inp(sh,14,2,24);

  // C: ICP Segments (rows 16–19)
  sectionHdr(sh,16,"C — ICP Segments (Mid-Market & Enterprise)");
  ["Segment","Beg. ACV ($)","Exp. ACV ($)","Churn Rate","Expansion %","Expansion Mo","CAC ($)","Lead Time (mo)","Close Rate"]
    .forEach((h,i) => hdr(sh,17,i+1,h,"#1F618D"));
  [
    ["Mid-Market", 80000, 150000, 0.08, 0.30, 12, 15000, 3, 0.20],
    ["Enterprise", 250000,1000000,0.05, 0.40, 18, 50000, 6, 0.10],
  ].forEach((seg,i) => {
    const r = 18+i;
    sh.getRange(r,1).setValue(seg[0]).setFontWeight("bold");
    inp(sh,r,2,seg[1],"$#,##0"); inp(sh,r,3,seg[2],"$#,##0");
    inp(sh,r,4,seg[3],"0%");     inp(sh,r,5,seg[4],"0%");
    inp(sh,r,6,seg[5]);          inp(sh,r,7,seg[6],"$#,##0");
    inp(sh,r,8,seg[7]);          inp(sh,r,9,seg[8],"0%");
  });

  // C2: Logo Ramp (rows 21–24)
  sectionHdr(sh,21,"C2 — New Logo Ramp (logos/month — should align with close rate × sales capacity)");
  ["Segment","Mo 1-6","Mo 7-12","Mo 13-18","Mo 19-24"].forEach((h,i) => hdr(sh,22,i+1,h,"#1F618D"));
  [["Mid-Market",1,2,4,6],["Enterprise",0,1,2,3]].forEach((row,i) => {
    sh.getRange(23+i,1).setValue(row[0]).setFontWeight("bold");
    [1,2,3,4].forEach(c => inp(sh,23+i,c+1,row[c]));
  });

  // D: Maintenance Ratios (rows 26–30)
  sectionHdr(sh,26,"D — Customer Maintenance Ratios (flagged in 🚦 Benchmarks)");
  ["Role","Max Accounts / Person","Industry Benchmark"].forEach((h,i) => hdr(sh,27,i+1,h,"#1F618D"));
  [
    ["Account Executive (AE)",         15, "10–20 for MM/Enterprise SaaS"],
    ["Field Deploy Engineer (FDE)",     10, "8–12 for complex deployments"],
    ["Customer Success Manager (CSM)", 20, "15–25 for tech-touch model"],
  ].forEach((row,i) => {
    const r = 28+i;
    sh.getRange(r,1).setValue(row[0]).setFontWeight("bold");
    inp(sh,r,2,row[1]);
    sh.getRange(r,3).setValue(row[2]).setFontColor("#888").setFontStyle("italic");
  });

  // E: Dept Defaults (rows 32–37)
  sectionHdr(sh,32,"E — Headcount: Department Defaults");
  ["Department","Start HC","Annual Salary ($)","SW Cost/mo ($)","HW One-time ($)","Insurance/mo ($)"]
    .forEach((h,i) => hdr(sh,33,i+1,h,"#1F618D"));
  [
    ["Engineering",  4, 180000, 300, 2500, 200],
    ["Sales",        2, 120000, 150, 1500, 200],
    ["CS / Support", 1,  90000, 100, 1500, 200],
    ["G&A",          2, 100000,  80, 1000, 200],
  ].forEach((d,i) => {
    const r = 34+i;
    sh.getRange(r,1).setValue(d[0]).setFontWeight("bold");
    inp(sh,r,2,d[1]);
    [2,3,4,5].forEach(c => inp(sh,r,c+1,d[c+1],"$#,##0"));
  });

  // E2: Individual Positions (rows 39–50)
  sectionHdr(sh,39,"E2 — Individual Positions (optional — specific hires with start dates)");
  ["Title","Department","Start Date","Annual Salary ($)","SW Cost/mo ($)"]
    .forEach((h,i) => hdr(sh,40,i+1,h,"#1F618D"));
  for (let i = 0; i < 10; i++) {
    const r = 41+i;
    [1,2].forEach(c => sh.getRange(r,c).setBackground("#EBF5FB"));
    sh.getRange(r,3).setBackground("#EBF5FB").setNumberFormat("MMM YYYY");
    sh.getRange(r,4).setBackground("#EBF5FB").setNumberFormat("$#,##0");
    sh.getRange(r,5).setBackground("#EBF5FB").setNumberFormat("$#,##0");
  }

  // F: Marketing (rows 52–55)
  sectionHdr(sh,52,"F — Marketing Budget (annual)");
  label(sh,53,1,"Events Budget ($)");        inp(sh,53,2,50000,"$#,##0");
  label(sh,54,1,"Digital / Other ($)");      inp(sh,54,2,30000,"$#,##0");
  label(sh,55,1,"Events as % of Marketing");
  sh.getRange(55,2).setFormula("=B53/(B53+B54)").setNumberFormat("0%");

  // G: Infrastructure (rows 57–59)
  sectionHdr(sh,57,"G — Infrastructure Costs");
  label(sh,58,1,"Infra / customer / mo ($)");     inp(sh,58,2,50,"$#,##0");
  label(sh,59,1,"Tooling / engineer / mo ($)");   inp(sh,59,2,200,"$#,##0");

  // H: Sales (rows 61–63)
  sectionHdr(sh,61,"H — Sales Commission");
  label(sh,62,1,"Commission (% of first-year ACV)"); inp(sh,62,2,0.10,"0%");
  label(sh,63,1,"Accelerator (% above quota)");      inp(sh,63,2,0.15,"0%");

  // Legend
  sectionHdr(sh,65,"Legend");
  sh.getRange(66,1).setValue("🔵 Blue = Input").setBackground("#EBF5FB");
  sh.getRange(66,2).setValue("⚫ Black = Formula");
  sh.getRange(66,3).setValue("🚦 Run Tetrix → Check Benchmarks after every scenario load").setFontStyle("italic").setFontColor("#888");
}

// ─── FUNDING ────────────────────────────────────────────────
function setupFunding(ss) {
  const sh = ss.getSheetByName("💰 Funding");
  sh.setColumnWidth(1,200); [2,3,4].forEach(c => sh.setColumnWidth(c,160));

  hdr(sh,1,1,"💰 FUNDING — Round Scenarios","#1A5276");
  sh.getRange(1,1,1,4).merge();

  ["","Seed $3M","Series A $10M","Series A $20M"].forEach((h,i) => { if(i) hdr(sh,2,i+1,h,"#1F618D"); });

  const rows = [
    ["Raise Amount",    3000000,  10000000, 20000000, "$#,##0"],
    ["Equity Dilution", 0.15,     0.20,     0.22,     "0%"],
    ["→ Engineering %", 0.40,     0.35,     0.30,     "0%"],
    ["→ Sales & Mktg %",0.30,     0.40,     0.45,     "0%"],
    ["→ G&A %",         0.15,     0.15,     0.15,     "0%"],
    ["→ Reserve %",     0.15,     0.10,     0.10,     "0%"],
  ];
  rows.forEach((row,r) => {
    label(sh,r+3,1,row[0]);
    [row[1],row[2],row[3]].forEach((v,c) => inp(sh,r+3,c+2,v,row[4]));
  });

  sectionHdr(sh,11,"Active Scenario Lookup (auto)");
  label(sh,12,1,"Active Raise Amount");
  sh.getRange(12,2).setFormula(
    `=INDEX(B3:D3,MATCH('🎛️ Drivers'!B5,{"Seed $3M","Series A $10M","Series A $20M"},0))`
  ).setNumberFormat("$#,##0").setBackground("#E8F8F5");

  sh.getRange(14,1).setValue("Note: The Funding tab supports legacy scenario comparison. For multi-round planning, use Section A of 🎛️ Drivers.")
    .setFontStyle("italic").setFontColor("#888");
  sh.getRange(14,1,1,4).merge();
}

// ─── HEADCOUNT ──────────────────────────────────────────────
function setupHeadcount(ss) {
  const sh = ss.getSheetByName("👥 Headcount");
  sh.setColumnWidth(1,160);
  const months = 24;
  for (let m = 1; m <= months; m++) sh.setColumnWidth(m+1, 55);

  hdr(sh,1,1,"👥 HEADCOUNT — Monthly Plan","#1A5276");
  sh.getRange(1,1,1,months+1).merge();

  hdr(sh,2,1,"Department / Month","#1F618D");
  for (let m = 1; m <= months; m++) hdr(sh,2,m+1,"Mo "+m,"#1F618D");

  const depts = ["Engineering","Sales","CS / Support","G&A"];
  const driverRows = [DR.ENG, DR.SALES, DR.CS, DR.GA]; // 34,35,36,37

  depts.forEach((dept,d) => {
    const hcRow = d*2+3;
    const costRow = hcRow+1;
    label(sh,hcRow,1,dept+" (HC)");
    label(sh,costRow,1,dept+" (Cost/mo)");

    for (let m = 1; m <= months; m++) {
      const col = m+1; const C = colLetter(col);
      if (m === 1) {
        sh.getRange(hcRow,col).setFormula(`='🎛️ Drivers'!B${driverRows[d]}`);
      } else {
        sh.getRange(hcRow,col).setFormula(`=${colLetter(col-1)}${hcRow}`);
      }
      // Monthly cost = (salary/12) + sw_cost + insurance_cost per person
      sh.getRange(costRow,col).setFormula(
        `=${C}${hcRow}*('🎛️ Drivers'!C${driverRows[d]}/12+'🎛️ Drivers'!D${driverRows[d]}+'🎛️ Drivers'!F${driverRows[d]})`
      ).setNumberFormat("$#,##0");
    }
  });

  const totalRow = depts.length*2+3;
  hdr(sh,totalRow,1,"Total Monthly Payroll","#2C3E50");
  for (let m = 1; m <= months; m++) {
    const C = colLetter(m+1);
    const costRows = [4,6,8,10];
    sh.getRange(totalRow,m+1).setFormula(`=${costRows.map(r => `${C}${r}`).join("+")}`)
      .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");
  }
}

// ─── REVENUE ────────────────────────────────────────────────
function setupRevenue(ss) {
  const sh = ss.getSheetByName("📈 Revenue");
  sh.setColumnWidth(1,160);

  hdr(sh,1,1,"📈 REVENUE — ARR Waterfall by Segment","#1A5276");
  sh.getRange(1,1,1,8).merge();

  // Drivers refs: MM row 18, ENT row 19 | ACV=col2, Churn=col4, Exp%=col5
  // Ramp: MM row 23, ENT row 24 | cols B-E = 4 bands
  const segs = [
    { name: "Mid-Market", driverRow: DR.MM_ROW, rampRow: DR.MM_RAMP },
    { name: "Enterprise", driverRow: DR.ENT_ROW, rampRow: DR.ENT_RAMP },
  ];

  let currentRow = 2;
  segs.forEach(seg => {
    sectionHdr(sh,currentRow,`Segment: ${seg.name}`); currentRow++;
    ["Month","New Logos","New ARR","Churn ARR","Expansion ARR","Net New ARR","Cumul. ARR","MRR"]
      .forEach((h,i) => hdr(sh,currentRow,i+1,h,"#1F618D"));
    currentRow++;

    for (let m = 1; m <= 24; m++) {
      const r = currentRow;
      sh.getRange(r,1).setValue(m);

      const rampCol = m<=6 ? "B" : m<=12 ? "C" : m<=18 ? "D" : "E";
      sh.getRange(r,2).setFormula(`='🎛️ Drivers'!${rampCol}${seg.rampRow}`);
      sh.getRange(r,3).setFormula(`=B${r}*'🎛️ Drivers'!B${seg.driverRow}`).setNumberFormat("$#,##0");

      if (m === 1) {
        sh.getRange(r,4).setValue(0).setNumberFormat("$#,##0");
        sh.getRange(r,5).setValue(0).setNumberFormat("$#,##0");
      } else {
        sh.getRange(r,4).setFormula(`=-G${r-1}*'🎛️ Drivers'!D${seg.driverRow}/12`).setNumberFormat("$#,##0");
        sh.getRange(r,5).setFormula(`=G${r-1}*'🎛️ Drivers'!E${seg.driverRow}/12`).setNumberFormat("$#,##0");
      }
      sh.getRange(r,6).setFormula(`=C${r}+D${r}+E${r}`).setNumberFormat("$#,##0");
      sh.getRange(r,7).setFormula(m===1 ? `=F${r}` : `=G${r-1}+F${r}`).setNumberFormat("$#,##0");
      sh.getRange(r,8).setFormula(`=G${r}/12`).setNumberFormat("$#,##0");
      currentRow++;
    }
    currentRow += 2;
  });

  sectionHdr(sh,currentRow,"📊 Total ARR Summary — wire G column from each segment above");
}

// ─── P&L ────────────────────────────────────────────────────
function setupPnL(ss) {
  const sh = ss.getSheetByName("💸 P&L");
  sh.setColumnWidth(1,220);
  const months = 24;
  for (let m = 1; m <= months; m++) sh.setColumnWidth(m+1,80);

  hdr(sh,1,1,"💸 P&L — Income Statement (Monthly)","#1A5276");
  sh.getRange(1,1,1,months+1).merge();
  hdr(sh,2,1,"Line Item","#1F618D");
  for (let m = 1; m <= months; m++) hdr(sh,2,m+1,"Mo "+m,"#1F618D");

  const lines = [
    "MRR","ARR","YoY Growth","","COGS",
    "Infrastructure","CS Payroll","Total COGS","Gross Profit","Gross Margin %","",
    "OPEX","Engineering Payroll","Sales Payroll","G&A Payroll",
    "Sales Commission","Marketing","Tooling / Misc","Total OpEx","","EBITDA","EBITDA Margin %","Cumulative Burn"
  ];
  lines.forEach((l,i) => { if(l) label(sh,i+3,1,l); });

  for (let m = 1; m <= months; m++) {
    const col = m+1; const C = colLetter(col);
    // MRR row 3 — blue, user wires from Revenue
    sh.getRange(3,col).setValue(0).setBackground("#EBF5FB").setFontColor("#1A5276").setNumberFormat("$#,##0");
    sh.getRange(4,col).setFormula(`=${C}3*12`).setNumberFormat("$#,##0");
    sh.getRange(5,col).setFormula(m<=12 ? `=""` : `=IFERROR(${C}4/${colLetter(col-12)}4-1,"")`).setNumberFormat("0%");
    // COGS
    sh.getRange(7,col).setFormula(`=${C}3/'🎛️ Drivers'!B${DR.MM_ROW}*${DR.INFRA}`).setNumberFormat("$#,##0");
    sh.getRange(8,col).setFormula(`='👥 Headcount'!${C}8`).setNumberFormat("$#,##0");
    sh.getRange(9,col).setFormula(`=${C}7+${C}8`).setNumberFormat("$#,##0").setFontWeight("bold");
    sh.getRange(10,col).setFormula(`=${C}3-${C}9`).setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");
    sh.getRange(11,col).setFormula(`=IFERROR(${C}10/${C}3,0)`).setNumberFormat("0%").setFontWeight("bold");
    // OpEx
    sh.getRange(14,col).setFormula(`='👥 Headcount'!${C}4`).setNumberFormat("$#,##0");
    sh.getRange(15,col).setFormula(`='👥 Headcount'!${C}6`).setNumberFormat("$#,##0");
    sh.getRange(16,col).setFormula(`='👥 Headcount'!${C}10`).setNumberFormat("$#,##0");
    sh.getRange(17,col).setFormula(`=MAX(0,${C}3-${m>1?colLetter(col-1)+"3":"0"})*12*'🎛️ Drivers'!${DR.COMMISSION}`).setNumberFormat("$#,##0");
    sh.getRange(18,col).setFormula(`=('🎛️ Drivers'!${DR.EVENTS}+'🎛️ Drivers'!${DR.DIGITAL})/12`).setNumberFormat("$#,##0");
    sh.getRange(19,col).setFormula(`='👥 Headcount'!${C}3*'🎛️ Drivers'!${DR.TOOLING}`).setNumberFormat("$#,##0");
    sh.getRange(20,col).setFormula(`=SUM(${C}14:${C}19)`).setNumberFormat("$#,##0").setFontWeight("bold");
    // Bottom line
    sh.getRange(22,col).setFormula(`=${C}10-${C}20`).setNumberFormat("$#,##0").setFontWeight("bold");
    sh.getRange(23,col).setFormula(`=IFERROR(${C}22/${C}3,0)`).setNumberFormat("0%");
    sh.getRange(24,col).setFormula(m===1 ? `=${C}22` : `=${colLetter(col-1)}24+${C}22`).setNumberFormat("$#,##0");

    const eRule = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor("#922B21")
      .setRanges([sh.getRange(22,col)]).build();
    const rules = sh.getConditionalFormatRules(); rules.push(eRule); sh.setConditionalFormatRules(rules);
  }
  sh.getRange(26,1).setValue("🔵 Wire MRR (row 3) to Revenue tab MRR once validated.")
    .setFontStyle("italic").setFontColor("#888"); sh.getRange(26,1,1,6).merge();
}

// ─── CASH FLOW ──────────────────────────────────────────────
function setupCashFlow(ss) {
  const sh = ss.getSheetByName("🏦 Cash Flow");
  sh.setColumnWidth(1,200);
  const months = 24;
  for (let m = 1; m <= months; m++) sh.setColumnWidth(m+1,75);

  hdr(sh,1,1,"🏦 CASH FLOW — Monthly Burn & Runway","#1A5276");
  sh.getRange(1,1,1,months+1).merge();
  hdr(sh,2,1,"Line Item","#1F618D");
  for (let m = 1; m <= months; m++) hdr(sh,2,m+1,"Mo "+m,"#1F618D");

  ["Beginning Cash","+ Capital Raised","+ Cash Collections (MRR)",
   "- Payroll","- Infra / COGS","- Sales & Marketing","- G&A Costs",
   "= Net Burn","= Ending Cash","Runway Remaining (mo)"]
    .forEach((l,i) => label(sh,i+3,1,l));

  sh.getRange(3,2).setFormula(`='💰 Funding'!B12`).setNumberFormat("$#,##0");

  for (let m = 1; m <= months; m++) {
    const col = m+1; const C = colLetter(col);
    if (m>1) sh.getRange(3,col).setFormula(`=${colLetter(col-1)}11`).setNumberFormat("$#,##0");
    if (m===1) sh.getRange(4,col).setFormula(`='💰 Funding'!B12`).setNumberFormat("$#,##0");
    else sh.getRange(4,col).setValue(0).setNumberFormat("$#,##0");
    sh.getRange(5,col).setValue(0).setBackground("#EBF5FB").setFontColor("#1A5276").setNumberFormat("$#,##0");
    sh.getRange(6,col).setFormula(`=-'👥 Headcount'!${C}11`).setNumberFormat("$#,##0");
    sh.getRange(7,col).setValue(0).setBackground("#EBF5FB").setFontColor("#1A5276").setNumberFormat("$#,##0");
    sh.getRange(8,col).setValue(0).setBackground("#EBF5FB").setFontColor("#1A5276").setNumberFormat("$#,##0");
    sh.getRange(9,col).setValue(0).setBackground("#EBF5FB").setFontColor("#1A5276").setNumberFormat("$#,##0");
    sh.getRange(10,col).setFormula(`=SUM(${C}4:${C}9)`).setNumberFormat("$#,##0").setFontWeight("bold");
    sh.getRange(11,col).setFormula(`=${C}3+${C}10`).setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");
    sh.getRange(12,col).setFormula(`=IFERROR(${C}11/ABS(${C}10),0)`).setNumberFormat("0.0");
  }

  const runwayRange = sh.getRange(12,2,1,months);
  const cfRule = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(6)
    .setBackground("#FADBD8").setFontColor("#922B21").setRanges([runwayRange]).build();
  const cfRules = sh.getConditionalFormatRules(); cfRules.push(cfRule); sh.setConditionalFormatRules(cfRules);

  sh.getRange(14,1).setValue("🔵 Blue cells = wire to P&L once MRR row is connected.")
    .setFontStyle("italic").setFontColor("#888"); sh.getRange(14,1,1,6).merge();
}

// ─── SUMMARY ────────────────────────────────────────────────
function setupSummary(ss) {
  const sh = ss.getSheetByName("📊 Summary");
  sh.setColumnWidth(1,240); sh.setColumnWidth(2,160); sh.setColumnWidth(3,160);

  hdr(sh,1,1,"📊 SUMMARY — Investor KPI Dashboard","#1A5276");
  sh.getRange(1,1,1,3).merge();
  label(sh,2,1,"Target ARR:"); sh.getRange(2,2).setFormula(`='🎛️ Drivers'!${DR.TARGET_ARR}`).setNumberFormat("$#,##0").setBackground("#FEF9E7").setFontWeight("bold");
  label(sh,3,1,"MoM Growth Target:"); sh.getRange(3,2).setFormula(`='🎛️ Drivers'!${DR.MOM_GROWTH}`).setNumberFormat("0%").setBackground("#FEF9E7");

  const sectionHdrLocal = (r,t) => sh.getRange(r,1,1,3).merge().setValue(t).setBackground("#D6EAF8").setFontWeight("bold");
  sectionHdrLocal(5,"📈 Revenue KPIs");
  [["MRR (Mo 1)","='💸 P&L'!B3","$#,##0"],["ARR (Mo 1)","='💸 P&L'!B4","$#,##0"],
   ["ARR (Mo 12)","='💸 P&L'!N4","$#,##0"],["ARR (Mo 24)","='💸 P&L'!Z4","$#,##0"],
   ["Gross Margin (Mo 1)","='💸 P&L'!B11","0%"]]
    .forEach(([l,f,fmt],i) => {
      label(sh,6+i,1,l);
      sh.getRange(6+i,2).setFormula(f).setNumberFormat(fmt).setBackground("#E8F8F5");
    });

  sectionHdrLocal(12,"🔥 Burn & Runway");
  [["Mo 1 Net Burn","='🏦 Cash Flow'!B10","$#,##0"],["Runway (Mo 1)","='🏦 Cash Flow'!B12","0.0"],
   ["Cash (Mo 12)","='🏦 Cash Flow'!N11","$#,##0"],["Cash (Mo 24)","='🏦 Cash Flow'!Z11","$#,##0"]]
    .forEach(([l,f,fmt],i) => { label(sh,13+i,1,l); sh.getRange(13+i,2).setFormula(f).setNumberFormat(fmt).setBackground("#FEF9E7"); });

  sectionHdrLocal(18,"👥 Team");
  [DR.ENG,DR.SALES,DR.CS,DR.GA].reduce((sum,r) => sum+`+'🎛️ Drivers'!B${r}`, "").slice(1);
  const totalHCFormula = (col) =>
    `='👥 Headcount'!${col}3+'👥 Headcount'!${col}5+'👥 Headcount'!${col}7+'👥 Headcount'!${col}9`;
  [["Total HC (Mo 1)", totalHCFormula("B")],["Total HC (Mo 12)", totalHCFormula("N")],
   ["Total HC (Mo 24)", totalHCFormula("Z")],["Payroll (Mo 1)","='👥 Headcount'!B11"]]
    .forEach(([l,f],i) => { label(sh,19+i,1,l); sh.getRange(19+i,2).setFormula(f).setNumberFormat(i===3?"$#,##0":"0").setBackground("#EBF5FB"); });

  sectionHdrLocal(24,"⚙️ Key Assumptions");
  [["MM Beg. ACV",`='🎛️ Drivers'!B${DR.MM_ROW}`,"$#,##0"],
   ["ENT Beg. ACV",`='🎛️ Drivers'!B${DR.ENT_ROW}`,"$#,##0"],
   ["Events Budget",`='🎛️ Drivers'!${DR.EVENTS}`,"$#,##0"],
   ["Forecast Horizon",`='🎛️ Drivers'!${DR.HORIZON}`,"0 \"months\""]]
    .forEach(([l,f,fmt],i) => { label(sh,25+i,1,l); sh.getRange(25+i,2).setFormula(f).setNumberFormat(fmt); });

  sh.getRange(30,1).setValue("All cells are formulas. Run 🚦 Benchmarks after every scenario change.")
    .setFontStyle("italic").setFontColor("#888"); sh.getRange(30,1,1,3).merge();
}

// ─── SCENARIOS ──────────────────────────────────────────────
function setupScenarios(ss) {
  const sh = ss.getSheetByName("📋 Scenarios");
  sh.setColumnWidth(1,220); [2,3,4].forEach(c => sh.setColumnWidth(c,160));

  hdr(sh,1,1,"📋 SCENARIOS — Investor Comparison View","#1A5276");
  sh.getRange(1,1,1,4).merge();
  ["Metric","Seed $3M","Series A $10M","Series A $20M"].forEach((h,i) => { if(i) hdr(sh,2,i+1,h,"#1F618D"); else hdr(sh,2,1,h,"#1F618D"); });

  ["Raise Amount","ARR at Month 12","ARR at Month 24","Customers at Mo 24","Headcount at Mo 24","Gross Margin","Runway (months)","ARR / Employee"]
    .forEach((m,i) => {
      label(sh,i+3,1,m);
      [2,3,4].forEach(c => sh.getRange(i+3,c).setValue("→ TBD").setFontColor("#AAAAAA"));
    });
  sh.getRange(3,2).setFormula(`='💰 Funding'!B3`).setNumberFormat("$#,##0").setFontColor("#000");
  sh.getRange(3,3).setFormula(`='💰 Funding'!C3`).setNumberFormat("$#,##0").setFontColor("#000");
  sh.getRange(3,4).setFormula(`='💰 Funding'!D3`).setNumberFormat("$#,##0").setFontColor("#000");

  sh.getRange(12,1).setValue("→ Wire remaining rows to Revenue + Cash Flow after building them out.")
    .setFontStyle("italic").setFontColor("#888"); sh.getRange(12,1,1,4).merge();
}

// ─── BENCHMARKS (structure only — runBenchmarks() is in Benchmarks.gs) ─────
function setupBenchmarks(ss) {
  const sh = ss.getSheetByName("🚦 Benchmarks");
  sh.setColumnWidth(1,240); sh.setColumnWidth(2,140); sh.setColumnWidth(3,100); sh.setColumnWidth(4,400);

  hdr(sh,1,1,"🚦 BENCHMARKS — Reality Check","#922B21");
  sh.getRange(1,1,1,4).merge();
  sh.getRange(2,1).setValue("Run 📊 Tetrix → Check Benchmarks to populate this tab.")
    .setFontStyle("italic").setFontColor("#888");
  sh.getRange(2,1,1,4).merge();

  ["Metric","Value","Status","Notes / Benchmark"].forEach((h,i) => hdr(sh,4,i+1,h,"#1F618D"));

  const metrics = [
    "CAC Payback — Mid-Market","CAC Payback — Enterprise",
    "LTV:CAC — Mid-Market","LTV:CAC — Enterprise",
    "Est. NRR — Mid-Market","Est. NRR — Enterprise",
    "ARR Growth (MoM → YoY implied)","Bessemer Growth Check",
    "Funding → ARR Coverage Ratio",
    "AE Account Load (at target ARR)","FDE Account Load (at target ARR)",
    "Gross Margin (if P&L wired)","Burn Multiple (if Cash Flow wired)",
  ];
  metrics.forEach((m,i) => {
    sh.getRange(5+i,1).setValue(m).setFontWeight("bold");
    sh.getRange(5+i,2).setValue("→ Run check").setFontColor("#AAAAAA");
    sh.getRange(5+i,3).setValue("—").setHorizontalAlignment("center");
    sh.getRange(5+i,4).setValue("").setFontColor("#888");
  });

  sh.getRange(19,1).setValue("Legend:")
    .setFontWeight("bold"); 
  [["🟢 Green","On track — meets or exceeds benchmark"],["🟡 Yellow","Caution — approaching limit"],["🔴 Red","Flag — review before sharing with investors"]]
    .forEach(([s,d],i) => {
      sh.getRange(20+i,1).setValue(s);
      sh.getRange(20+i,2,1,3).merge().setValue(d).setFontColor("#888");
    });
}