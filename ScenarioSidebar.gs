// ============================================================
// TETRIX SCENARIO SIDEBAR v2 (aligned with v3.1 Drivers + DR)
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

/** Cell values can be Date, string, or empty; Utilities.formatDate requires a Date. */
function scenarioCoerceDate_(val, fallback) {
  var fb = fallback || new Date();
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  if (val === null || val === undefined || val === "") return fb;
  if (typeof val === "string" && val.trim() === "") return fb;
  var d = new Date(val);
  return isNaN(d.getTime()) ? fb : d;
}

function scenarioFormatDateIso_(val) {
  var d = scenarioCoerceDate_(val, new Date());
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// ─── APPLY SCENARIO ─────────────────────────────────────────

function applyScenario(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("🎛️ Drivers");
  if (!sh) throw new Error("'🎛️ Drivers' tab not found. Run setupFinancialModel() first.");

  // ── Meta
  sh.getRange(DR.HORIZON).setValue(data.meta.forecastHorizon);

  // ── Funding Rounds (rows 121–123, cols 1–5) — v3 has up to 3 rounds
  const rounds = (data.fundingRounds || []).slice(0, 3);
  for (let i = 0; i < 3; i++) {
    const r = 121 + i;
    const round = rounds[i];
    if (round) {
      sh.getRange(r, 1).setValue(round.name || "");
      sh.getRange(r, 2).setValue(round.amount || "");
      sh.getRange(r, 3).setValue(round.closeDate ? new Date(round.closeDate) : "");
      sh.getRange(r, 4).setValue(round.expectedARR || "");
      sh.getRange(r, 5).setValue(round.notes || "");
    } else {
      [1, 2, 3, 4, 5].forEach(function (c) { sh.getRange(r, c).clearContent(); });
    }
  }

  // ── ARR Targets (B12 / B13 — B12 may be formula in v3; loading overwrites display target)
  sh.getRange(DR.TARGET_ARR).setValue(data.arrTargets.targetARR);
  sh.getRange(DR.MOM_GROWTH).setValue(data.arrTargets.momGrowthRate);

  // ── ICP Segments (row 18 = MM, row 19 = ENT)
  [["midMarket", DR.MM_ROW], ["enterprise", DR.ENT_ROW]].forEach(function (pair) {
    var key = pair[0];
    var row = pair[1];
    var seg = data.segments[key];
    sh.getRange(row, 2).setValue(seg.begACV);
    sh.getRange(row, 3).setValue(seg.expACV);
    sh.getRange(row, 4).setValue(seg.churnRate);
    sh.getRange(row, 5).setValue(seg.expansionRate);
    sh.getRange(row, 6).setValue(seg.expansionMonth);
    sh.getRange(row, 7).setValue(seg.cac);
    sh.getRange(row, 8).setValue(seg.leadTime);
    sh.getRange(row, 9).setValue(seg.closeRate);
  });

  // ── Logo MoM growth: v3 single input B39; JSON still uses 4 slots — use first non-zero or MM[0]
  var mmRamp = data.logoRamp.midMarket || [0, 0, 0, 0];
  var entRamp = data.logoRamp.enterprise || [0, 0, 0, 0];
  var growth = 0;
  for (var ri = 0; ri < mmRamp.length; ri++) {
    var x = Number(mmRamp[ri]);
    if (!isNaN(x) && x !== 0) { growth = x; break; }
  }
  if (!growth) {
    for (var ej = 0; ej < entRamp.length; ej++) {
      var y = Number(entRamp[ej]);
      if (!isNaN(y) && y !== 0) { growth = y; break; }
    }
  }
  sh.getRange(DR.LOGO_GROWTH).setValue(growth);

  // ── Maintenance Ratios (B24–B26)
  var mr = data.maintenanceRatios;
  sh.getRange(DR.AE_RATIO).setValue(mr.aePerAccounts);
  sh.getRange(DR.FDE_RATIO).setValue(mr.fdePerAccounts);
  sh.getRange(DR.CSM_RATIO).setValue(mr.csmPerAccounts);

  // ── Headcount Dept Defaults (rows 47–50)
  var deptMap = [
    ["engineering", DR.ENG],
    ["sales", DR.SALES],
    ["csSupport", DR.CS],
    ["gAndA", DR.GA]
  ];
  deptMap.forEach(function (pair) {
    var key = pair[0];
    var row = pair[1];
    var dept = data.headcount.deptDefaults[key];
    sh.getRange(row, 2).setValue(dept.startHC);
    sh.getRange(row, 3).setValue(dept.annualSalary);
    sh.getRange(row, 4).setValue(dept.swCostPerMo);
    sh.getRange(row, 5).setValue(dept.hwCostOneTime);
    sh.getRange(row, 6).setValue(dept.insurancePerMo);
  });

  // ── Individual Positions (rows 54–63; header row 53)
  var positions = data.headcount.positions || [];
  for (var pi = 0; pi < 10; pi++) {
    var pr = 54 + pi;
    var pos = positions[pi];
    if (pos) {
      sh.getRange(pr, 1).setValue(pos.title || "");
      sh.getRange(pr, 2).setValue(pos.dept || "");
      sh.getRange(pr, 3).setValue(pos.startDate ? new Date(pos.startDate) : "");
      sh.getRange(pr, 4).setValue(pos.annualSalary || "");
      sh.getRange(pr, 5).setValue(pos.swCostPerMo || "");
    } else {
      [1, 2, 3, 4, 5].forEach(function (c) { sh.getRange(pr, c).clearContent(); });
    }
  }

  // ── Marketing / Infrastructure / Sales
  sh.getRange(DR.EVENTS).setValue(data.marketing.eventsAnnual);
  sh.getRange(DR.DIGITAL).setValue(data.marketing.digitalAnnual);
  sh.getRange(DR.INFRA).setValue(data.infrastructure.infraPerCustomerPerMo);
  sh.getRange(DR.TOOLING).setValue(data.infrastructure.toolingPerEngineerPerMo);
  sh.getRange(DR.COMMISSION).setValue(data.sales.commission);
  sh.getRange(DR.ACCELERATOR).setValue(data.sales.accelerator);

  return '✅ "' + data.meta.name + '" loaded. Run 🚦 Check Benchmarks to validate.';
}

// ─── EXPORT CURRENT STATE ────────────────────────────────────

function getCurrentScenario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("🎛️ Drivers");
  if (!sh) return null;
  function v(row, col) { return sh.getRange(row, col).getValue(); }

  var rounds = [];
  for (var i = 0; i < 3; i++) {
    var name = v(121 + i, 1);
    if (name) {
      rounds.push({
        name: name,
        amount: v(121 + i, 2),
        closeDate: scenarioFormatDateIso_(v(121 + i, 3)),
        expectedARR: v(121 + i, 4),
        notes: v(121 + i, 5)
      });
    }
  }

  var positions = [];
  for (var j = 0; j < 10; j++) {
    var title = v(54 + j, 1);
    if (title) {
      positions.push({
        title: title,
        dept: v(54 + j, 2),
        startDate: scenarioFormatDateIso_(v(54 + j, 3)),
        annualSalary: v(54 + j, 4),
        swCostPerMo: v(54 + j, 5)
      });
    }
  }

  var lg = sh.getRange(DR.LOGO_GROWTH).getValue();
  if (lg === "" || lg === null || lg === undefined) lg = 0;

  return {
    meta: { name: "Current Model State", forecastHorizon: sh.getRange(DR.HORIZON).getValue() },
    fundingRounds: rounds,
    arrTargets: {
      targetARR: sh.getRange(DR.TARGET_ARR).getValue(),
      momGrowthRate: sh.getRange(DR.MOM_GROWTH).getValue()
    },
    segments: {
      midMarket: {
        begACV: v(DR.MM_ROW, 2), expACV: v(DR.MM_ROW, 3), churnRate: v(DR.MM_ROW, 4),
        expansionRate: v(DR.MM_ROW, 5), expansionMonth: v(DR.MM_ROW, 6), cac: v(DR.MM_ROW, 7),
        leadTime: v(DR.MM_ROW, 8), closeRate: v(DR.MM_ROW, 9)
      },
      enterprise: {
        begACV: v(DR.ENT_ROW, 2), expACV: v(DR.ENT_ROW, 3), churnRate: v(DR.ENT_ROW, 4),
        expansionRate: v(DR.ENT_ROW, 5), expansionMonth: v(DR.ENT_ROW, 6), cac: v(DR.ENT_ROW, 7),
        leadTime: v(DR.ENT_ROW, 8), closeRate: v(DR.ENT_ROW, 9)
      }
    },
    logoRamp: {
      midMarket: [lg, 0, 0, 0],
      enterprise: [lg, 0, 0, 0]
    },
    maintenanceRatios: {
      aePerAccounts: sh.getRange(DR.AE_RATIO).getValue(),
      fdePerAccounts: sh.getRange(DR.FDE_RATIO).getValue(),
      csmPerAccounts: sh.getRange(DR.CSM_RATIO).getValue()
    },
    headcount: {
      deptDefaults: {
        engineering: {
          startHC: v(DR.ENG, 2), annualSalary: v(DR.ENG, 3), swCostPerMo: v(DR.ENG, 4),
          hwCostOneTime: v(DR.ENG, 5), insurancePerMo: v(DR.ENG, 6)
        },
        sales: {
          startHC: v(DR.SALES, 2), annualSalary: v(DR.SALES, 3), swCostPerMo: v(DR.SALES, 4),
          hwCostOneTime: v(DR.SALES, 5), insurancePerMo: v(DR.SALES, 6)
        },
        csSupport: {
          startHC: v(DR.CS, 2), annualSalary: v(DR.CS, 3), swCostPerMo: v(DR.CS, 4),
          hwCostOneTime: v(DR.CS, 5), insurancePerMo: v(DR.CS, 6)
        },
        gAndA: {
          startHC: v(DR.GA, 2), annualSalary: v(DR.GA, 3), swCostPerMo: v(DR.GA, 4),
          hwCostOneTime: v(DR.GA, 5), insurancePerMo: v(DR.GA, 6)
        }
      },
      positions: positions
    },
    marketing: {
      eventsAnnual: sh.getRange(DR.EVENTS).getValue(),
      digitalAnnual: sh.getRange(DR.DIGITAL).getValue()
    },
    infrastructure: {
      infraPerCustomerPerMo: sh.getRange(DR.INFRA).getValue(),
      toolingPerEngineerPerMo: sh.getRange(DR.TOOLING).getValue()
    },
    sales: {
      commission: sh.getRange(DR.COMMISSION).getValue(),
      accelerator: sh.getRange(DR.ACCELERATOR).getValue()
    }
  };
}
