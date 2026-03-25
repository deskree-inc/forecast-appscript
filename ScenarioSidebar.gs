// ============================================================
// TETRIX SCENARIO SIDEBAR v2
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Tetrix")
    .addItem("Open Scenario Loader", "openScenarioSidebar")
    .addItem("🚦 Check Benchmarks",  "runBenchmarks")
    .addToUi();
}

function openScenarioSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ScenarioSidebarView")
    .setTitle("Tetrix Scenario Loader")
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ─── APPLY SCENARIO ─────────────────────────────────────────

function applyScenario(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("🎛️ Drivers");
  if (!sh) throw new Error("'🎛️ Drivers' tab not found. Run setupFinancialModel() first.");

  // ── Meta
  sh.getRange("B14").setValue(data.meta.forecastHorizon);

  // ── Funding Rounds (rows 5–9, cols 1–5)
  const rounds = data.fundingRounds || [];
  for (let i = 0; i < 5; i++) {
    const r = 5 + i;
    const round = rounds[i];
    if (round) {
      sh.getRange(r,1).setValue(round.name || "");
      sh.getRange(r,2).setValue(round.amount || "");
      sh.getRange(r,3).setValue(round.closeDate ? new Date(round.closeDate) : "");
      sh.getRange(r,4).setValue(round.expectedARR || "");
      sh.getRange(r,5).setValue(round.notes || "");
    } else {
      [1,2,3,4,5].forEach(c => sh.getRange(r,c).clearContent());
    }
  }

  // ── ARR Targets
  sh.getRange("B12").setValue(data.arrTargets.targetARR);
  sh.getRange("B13").setValue(data.arrTargets.momGrowthRate);

  // ── ICP Segments (row 18 = MM, row 19 = ENT)
  [["midMarket", 18], ["enterprise", 19]].forEach(([key, row]) => {
    const seg = data.segments[key];
    sh.getRange(row,2).setValue(seg.begACV);
    sh.getRange(row,3).setValue(seg.expACV);
    sh.getRange(row,4).setValue(seg.churnRate);
    sh.getRange(row,5).setValue(seg.expansionRate);
    sh.getRange(row,6).setValue(seg.expansionMonth);
    sh.getRange(row,7).setValue(seg.cac);
    sh.getRange(row,8).setValue(seg.leadTime);
    sh.getRange(row,9).setValue(seg.closeRate);
  });

  // ── Logo Ramp (row 23 = MM, row 24 = ENT, cols B–E)
  [["midMarket", 23], ["enterprise", 24]].forEach(([key, row]) => {
    const ramp = data.logoRamp[key] || [0,0,0,0];
    ramp.forEach((v, c) => sh.getRange(row, c+2).setValue(v));
  });

  // ── Maintenance Ratios (rows 28–30, col 2)
  const mr = data.maintenanceRatios;
  sh.getRange("B28").setValue(mr.aePerAccounts);
  sh.getRange("B29").setValue(mr.fdePerAccounts);
  sh.getRange("B30").setValue(mr.csmPerAccounts);

  // ── Headcount Dept Defaults (rows 34–37)
  const deptMap = [
    ["engineering", 34], ["sales", 35], ["csSupport", 36], ["gAndA", 37]
  ];
  deptMap.forEach(([key, row]) => {
    const dept = data.headcount.deptDefaults[key];
    sh.getRange(row,2).setValue(dept.startHC);
    sh.getRange(row,3).setValue(dept.annualSalary);
    sh.getRange(row,4).setValue(dept.swCostPerMo);
    sh.getRange(row,5).setValue(dept.hwCostOneTime);
    sh.getRange(row,6).setValue(dept.insurancePerMo);
  });

  // ── Individual Positions (rows 41–50)
  const positions = data.headcount.positions || [];
  for (let i = 0; i < 10; i++) {
    const r = 41 + i;
    const pos = positions[i];
    if (pos) {
      sh.getRange(r,1).setValue(pos.title || "");
      sh.getRange(r,2).setValue(pos.dept || "");
      sh.getRange(r,3).setValue(pos.startDate ? new Date(pos.startDate) : "");
      sh.getRange(r,4).setValue(pos.annualSalary || "");
      sh.getRange(r,5).setValue(pos.swCostPerMo || "");
    } else {
      [1,2,3,4,5].forEach(c => sh.getRange(r,c).clearContent());
    }
  }

  // ── Marketing
  sh.getRange("B53").setValue(data.marketing.eventsAnnual);
  sh.getRange("B54").setValue(data.marketing.digitalAnnual);

  // ── Infrastructure
  sh.getRange("B58").setValue(data.infrastructure.infraPerCustomerPerMo);
  sh.getRange("B59").setValue(data.infrastructure.toolingPerEngineerPerMo);

  // ── Sales
  sh.getRange("B62").setValue(data.sales.commission);
  sh.getRange("B63").setValue(data.sales.accelerator);

  return `✅ "${data.meta.name}" loaded. Run 🚦 Check Benchmarks to validate.`;
}

// ─── EXPORT CURRENT STATE ────────────────────────────────────

function getCurrentScenario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("🎛️ Drivers");
  if (!sh) return null;
  function v(row, col) { return sh.getRange(row, col).getValue(); }

  const rounds = [];
  for (let i = 0; i < 5; i++) {
    const name = v(5+i, 1);
    if (name) rounds.push({
      name, amount: v(5+i,2),
      closeDate: Utilities.formatDate(v(5+i,3) || new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      expectedARR: v(5+i,4), notes: v(5+i,5),
    });
  }

  const positions = [];
  for (let i = 0; i < 10; i++) {
    const title = v(41+i, 1);
    if (title) positions.push({
      title, dept: v(41+i,2),
      startDate: Utilities.formatDate(v(41+i,3) || new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      annualSalary: v(41+i,4), swCostPerMo: v(41+i,5),
    });
  }

  return {
    meta: { name: "Current Model State", forecastHorizon: v(14,2) },
    fundingRounds: rounds,
    arrTargets: { targetARR: v(12,2), momGrowthRate: v(13,2) },
    segments: {
      midMarket:  { begACV: v(18,2), expACV: v(18,3), churnRate: v(18,4), expansionRate: v(18,5), expansionMonth: v(18,6), cac: v(18,7), leadTime: v(18,8), closeRate: v(18,9) },
      enterprise: { begACV: v(19,2), expACV: v(19,3), churnRate: v(19,4), expansionRate: v(19,5), expansionMonth: v(19,6), cac: v(19,7), leadTime: v(19,8), closeRate: v(19,9) },
    },
    logoRamp: {
      midMarket:  [v(23,2), v(23,3), v(23,4), v(23,5)],
      enterprise: [v(24,2), v(24,3), v(24,4), v(24,5)],
    },
    maintenanceRatios: { aePerAccounts: v(28,2), fdePerAccounts: v(29,2), csmPerAccounts: v(30,2) },
    headcount: {
      deptDefaults: {
        engineering: { startHC: v(34,2), annualSalary: v(34,3), swCostPerMo: v(34,4), hwCostOneTime: v(34,5), insurancePerMo: v(34,6) },
        sales:       { startHC: v(35,2), annualSalary: v(35,3), swCostPerMo: v(35,4), hwCostOneTime: v(35,5), insurancePerMo: v(35,6) },
        csSupport:   { startHC: v(36,2), annualSalary: v(36,3), swCostPerMo: v(36,4), hwCostOneTime: v(36,5), insurancePerMo: v(36,6) },
        gAndA:       { startHC: v(37,2), annualSalary: v(37,3), swCostPerMo: v(37,4), hwCostOneTime: v(37,5), insurancePerMo: v(37,6) },
      },
      positions,
    },
    marketing:      { eventsAnnual: v(53,2), digitalAnnual: v(54,2) },
    infrastructure: { infraPerCustomerPerMo: v(58,2), toolingPerEngineerPerMo: v(59,2) },
    sales:          { commission: v(62,2), accelerator: v(63,2) },
  };
}