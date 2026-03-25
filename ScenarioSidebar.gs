// ============================================================
// TETRIX SCENARIO SIDEBAR
// Add this as a new script file called "ScenarioSidebar"
// in the same Apps Script project as your model.
// ============================================================

// ─── MENU ───────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Tetrix")
    .addItem("Open Scenario Loader", "openScenarioSidebar")
    .addToUi();
}

function openScenarioSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ScenarioSidebarView")
    .setTitle("Tetrix Scenario Loader")
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ─── CALLED BY SIDEBAR ──────────────────────────────────────

function applyScenario(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("🎛️ Drivers");
  if (!sh) throw new Error("'🎛️ Drivers' tab not found. Run setupFinancialModel() first.");

  // Meta
  sh.getRange("B4").setValue(data.meta.activeScenario);
  sh.getRange("B5").setValue(new Date(data.meta.closeDate));
  sh.getRange("B6").setValue(data.meta.runwayTarget);
  sh.getRange("B7").setValue(data.meta.forecastHorizon);

  // Segments
  const segs = [data.segments.enterprise, data.segments.midMarket, data.segments.smb];
  segs.forEach((seg, i) => {
    const r = 11 + i;
    sh.getRange(r, 2).setValue(seg.acv);
    sh.getRange(r, 3).setValue(seg.salesCycle);
    sh.getRange(r, 4).setValue(seg.churnRate);
    sh.getRange(r, 5).setValue(seg.expansionRate);
  });

  // Logo ramp
  const ramps = [data.logoRamp.enterprise, data.logoRamp.midMarket, data.logoRamp.smb];
  ramps.forEach((ramp, i) => {
    ramp.forEach((val, c) => sh.getRange(17 + i, c + 2).setValue(val));
  });

  // Headcount
  const depts = [
    data.headcount.engineering, data.headcount.sales,
    data.headcount.csSupport,   data.headcount.gAndA
  ];
  depts.forEach((dept, i) => {
    sh.getRange(23 + i, 2).setValue(dept.startHC);
    sh.getRange(23 + i, 3).setValue(dept.hireTrigger);
    sh.getRange(23 + i, 4).setValue(dept.annualCost);
  });
  sh.getRange("B28").setValue(data.salesCommission);

  // Costs
  sh.getRange("B31").setValue(data.costs.infraPerCustomerPerMonth);
  sh.getRange("B32").setValue(data.costs.toolingPerEngineerPerMonth);
  sh.getRange("B33").setValue(data.costs.officeMiscPerEmployeePerMonth);
  sh.getRange("B34").setValue(data.costs.marketingPctOfRaise);

  return `✅ "${data.meta.name}" loaded successfully.`;
}

function getCurrentScenario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("🎛️ Drivers");
  if (!sh) return null;

  return {
    meta: {
      name:            "Current Model State",
      activeScenario:  sh.getRange("B4").getValue(),
      closeDate:       Utilities.formatDate(sh.getRange("B5").getValue() || new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      runwayTarget:    sh.getRange("B6").getValue(),
      forecastHorizon: sh.getRange("B7").getValue(),
    },
    segments: {
      enterprise: { acv: sh.getRange(11,2).getValue(), salesCycle: sh.getRange(11,3).getValue(), churnRate: sh.getRange(11,4).getValue(), expansionRate: sh.getRange(11,5).getValue() },
      midMarket:  { acv: sh.getRange(12,2).getValue(), salesCycle: sh.getRange(12,3).getValue(), churnRate: sh.getRange(12,4).getValue(), expansionRate: sh.getRange(12,5).getValue() },
      smb:        { acv: sh.getRange(13,2).getValue(), salesCycle: sh.getRange(13,3).getValue(), churnRate: sh.getRange(13,4).getValue(), expansionRate: sh.getRange(13,5).getValue() },
    },
    logoRamp: {
      enterprise: [sh.getRange(17,2).getValue(), sh.getRange(17,3).getValue(), sh.getRange(17,4).getValue(), sh.getRange(17,5).getValue()],
      midMarket:  [sh.getRange(18,2).getValue(), sh.getRange(18,3).getValue(), sh.getRange(18,4).getValue(), sh.getRange(18,5).getValue()],
      smb:        [sh.getRange(19,2).getValue(), sh.getRange(19,3).getValue(), sh.getRange(19,4).getValue(), sh.getRange(19,5).getValue()],
    },
    headcount: {
      engineering: { startHC: sh.getRange(23,2).getValue(), hireTrigger: sh.getRange(23,3).getValue(), annualCost: sh.getRange(23,4).getValue() },
      sales:       { startHC: sh.getRange(24,2).getValue(), hireTrigger: sh.getRange(24,3).getValue(), annualCost: sh.getRange(24,4).getValue() },
      csSupport:   { startHC: sh.getRange(25,2).getValue(), hireTrigger: sh.getRange(25,3).getValue(), annualCost: sh.getRange(25,4).getValue() },
      gAndA:       { startHC: sh.getRange(26,2).getValue(), hireTrigger: sh.getRange(26,3).getValue(), annualCost: sh.getRange(26,4).getValue() },
    },
    salesCommission: sh.getRange("B28").getValue(),
    costs: {
      infraPerCustomerPerMonth:      sh.getRange("B31").getValue(),
      toolingPerEngineerPerMonth:    sh.getRange("B32").getValue(),
      officeMiscPerEmployeePerMonth: sh.getRange("B33").getValue(),
      marketingPctOfRaise:           sh.getRange("B34").getValue(),
    }
  };
}