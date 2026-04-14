// ============================================================
// Scenario manager sidebar + custom menu (Drivers + DR). Labels: ModelConstants APP_*.
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP_MENU_LABEL)
    .addItem("Open Scenario Manager", "openScenarioSidebar")
    .addItem("🚦 Check Benchmarks",  "runBenchmarks")
    .addItem("🔧 Rebuild model (run setup)…", "runSetupFromMenu")
    .addSeparator()
    .addItem("Enter investor view (hide internal tabs)", "showInvestorView")
    .addItem("Show all internal sheets", "showInternalSheets")
    .addToUi();
}

/**
 * Runs setupFinancialModel() after confirmation (same as running it in the Apps Script editor).
 */
function runSetupFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert(
    "Rebuild financial model?",
    "This runs setupFinancialModel(). It clears and rebuilds the model tabs (Start here, Drivers, Revenue, etc.).\n\nContinue?",
    ui.ButtonSet.YES_NO
  );
  if (r !== ui.Button.YES) return;
  setupFinancialModel();
}

function openScenarioSidebar() {
  var tpl = HtmlService.createTemplateFromFile("ScenarioSidebarView");
  tpl.menuLabel = APP_MENU_LABEL;
  tpl.headingSuffix = APP_SIDEBAR_HEADING_SUFFIX;
  var html = tpl
    .evaluate()
    .setTitle(APP_MENU_LABEL + " — " + APP_SIDEBAR_HEADING_SUFFIX)
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Parses JSON / sheet values into a Date for writing to the sheet.
 * ISO date-only strings ("yyyy-MM-dd") are interpreted as calendar dates in the script
 * timezone — avoids UTC midnight shifting the day (e.g. forecast start showing wrong month).
 */
function scenarioParseDateForSheet_(val) {
  if (val === undefined || val === null || val === "") return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === "number" && !isNaN(val)) {
    var dn = new Date(val);
    return isNaN(dn.getTime()) ? null : dn;
  }
  var s = String(val).trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    var y = parseInt(iso[1], 10);
    var mo = parseInt(iso[2], 10);
    var day = parseInt(iso[3], 10);
    if (y >= 1900 && y <= 2200 && mo >= 1 && mo <= 12 && day >= 1 && day <= 31) {
      return new Date(y, mo - 1, day);
    }
  }
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** Cell values can be Date, string, or empty; Utilities.formatDate requires a Date. */
function scenarioCoerceDate_(val, fallback) {
  var fb = fallback || new Date();
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  if (val === null || val === undefined || val === "") return fb;
  if (typeof val === "string" && val.trim() === "") return fb;
  var parsed = scenarioParseDateForSheet_(val);
  return parsed || fb;
}

function scenarioFormatDateIso_(val) {
  var d = scenarioCoerceDate_(val, new Date());
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/** ISO date or "" when the cell is blank / invalid (export). */
function scenarioFormatDateIsoMaybeEmpty_(val) {
  if (val === "" || val === null || val === undefined) return "";
  if (val instanceof Date && isNaN(val.getTime())) return "";
  var d = val instanceof Date ? val : new Date(val);
  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function scenarioSetDateCell_(sh, a1, val) {
  if (val === undefined || val === null || val === "") return;
  var d = scenarioParseDateForSheet_(val);
  if (!d) return;
  sh.getRange(a1).setValue(d);
}

function scenarioSetIfDefined_(sh, a1, val) {
  if (val === undefined) return;
  sh.getRange(a1).setValue(val);
}

function scenarioSetRowColIfDefined_(sh, row, col, val) {
  if (val === undefined) return;
  sh.getRange(row, col).setValue(val);
}

// ─── APPLY SCENARIO ─────────────────────────────────────────

function applyScenario(data) {
  if (!data || typeof data !== "object") throw new Error("Invalid scenario: expected an object.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("🎛️ Drivers");
  if (!sh) throw new Error("'🎛️ Drivers' tab not found. Run setupFinancialModel() first.");

  var meta = data.meta || {};
  if (meta.forecastHorizon != null) sh.getRange(DR.HORIZON).setValue(meta.forecastHorizon);

  var tm = data.timing;
  if (tm) {
    if (tm.forecastStart != null) scenarioSetDateCell_(sh, DR.FORECAST_START, tm.forecastStart);
    if (tm.firstMMClientDate != null) scenarioSetDateCell_(sh, DR.FIRST_MM_CLIENT, tm.firstMMClientDate);
    if (tm.firstEntClientDate != null) scenarioSetDateCell_(sh, DR.FIRST_ENT_CLIENT, tm.firstEntClientDate);
  }
  // Top-level alias if timing block is missing or omits forecastStart
  if ((!tm || tm.forecastStart == null) && data.forecastStart != null) {
    scenarioSetDateCell_(sh, DR.FORECAST_START, data.forecastStart);
  }

  var years = data.annualArrTargets;
  if (years && years.length) {
    for (var yi = 0; yi < Math.min(5, years.length); yi++) {
      var yr = years[yi];
      if (!yr || typeof yr !== "object") continue;
      var row = 134 + yi;
      if (yr.targetARR != null) sh.getRange(row, 2).setValue(yr.targetARR);
      if (yr.targetDate != null) scenarioSetDateCell_(sh, "C" + row, yr.targetDate);
      if (yr.notes != null) sh.getRange(row, 5).setValue(yr.notes);
    }
  }

  var at = data.arrTargets || {};
  // B12 is a formula (target ARR from Section L); never overwrite — would break the model.
  if (at.momGrowthRate != null) sh.getRange(DR.MOM_GROWTH).setValue(at.momGrowthRate);

  var rounds = (data.fundingRounds || []).slice(0, 3);
  for (var i = 0; i < 3; i++) {
    var r = 123 + i;
    var round = rounds[i];
    if (round) {
      sh.getRange(r, 1).setValue(round.name || "");
      sh.getRange(r, 2).setValue(round.amount || "");
      sh.getRange(r, 3).setValue(round.closeDate ? (scenarioParseDateForSheet_(round.closeDate) || "") : "");
      sh.getRange(r, 4).setValue(round.notes || "");
    } else {
      [1, 2, 3, 4].forEach(function (c) { sh.getRange(r, c).clearContent(); });
    }
  }

  var fm = data.fundingMeta || {};
  if (fm.interestRate != null) sh.getRange(DR.INTEREST_RATE).setValue(fm.interestRate);
  if (fm.openingCash != null) sh.getRange(DR.OPENING_CASH).setValue(fm.openingCash);

  [["midMarket", DR.MM_ROW], ["enterprise", DR.ENT_ROW]].forEach(function (pair) {
    var key = pair[0];
    var row = pair[1];
    var seg = data.segments && data.segments[key];
    if (!seg) return;
    scenarioSetRowColIfDefined_(sh, row, 2, seg.begACV);
    scenarioSetRowColIfDefined_(sh, row, 3, seg.expACV);
    scenarioSetRowColIfDefined_(sh, row, 4, seg.churnRate);
    scenarioSetRowColIfDefined_(sh, row, 5, seg.expansionRate);
    scenarioSetRowColIfDefined_(sh, row, 6, seg.expansionMonth);
    scenarioSetRowColIfDefined_(sh, row, 7, seg.cac);
    scenarioSetRowColIfDefined_(sh, row, 8, seg.leadTime);
    scenarioSetRowColIfDefined_(sh, row, 9, seg.closeRate);
  });

  if (data.logoGrowth != null) {
    sh.getRange(DR.LOGO_GROWTH).setValue(data.logoGrowth);
  } else {
    var mmRamp = (data.logoRamp && data.logoRamp.midMarket) || [0, 0, 0, 0];
    var entRamp = (data.logoRamp && data.logoRamp.enterprise) || [0, 0, 0, 0];
    var growth = 0;
    var ri;
    for (ri = 0; ri < mmRamp.length; ri++) {
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
  }

  var fde = data.fdeCapacity;
  if (fde) {
    scenarioSetIfDefined_(sh, DR.FDE_MM_CAPACITY, fde.concurrentMmPerFde);
    scenarioSetIfDefined_(sh, DR.FDE_ENT_CAPACITY, fde.concurrentEntPerFde);
  }

  var lb = data.logoBackCalculation;
  if (lb) {
    scenarioSetIfDefined_(sh, DR.AE_QUOTA_MM, lb.aeQuotaMm);
    scenarioSetIfDefined_(sh, DR.AE_QUOTA_ENT, lb.aeQuotaEnt);
    scenarioSetIfDefined_(sh, DR.ATTAINMENT, lb.attainment);
    scenarioSetIfDefined_(sh, DR.MM_PCT_ARR, lb.mmPctOfTargetArr);
    scenarioSetIfDefined_(sh, DR.LOGO_GROWTH, lb.logoMomGrowth);
  }

  var mr = data.maintenanceRatios;
  if (mr) {
    scenarioSetIfDefined_(sh, DR.AE_RATIO, mr.aePerAccounts);
    scenarioSetIfDefined_(sh, DR.FDE_RATIO, mr.fdePerAccounts);
    scenarioSetIfDefined_(sh, DR.CSM_RATIO, mr.csmPerAccounts);
  }

  var deptMap = [
    ["engineering", DR.ENG],
    ["sales", DR.SALES],
    ["csSupport", DR.CS],
    ["marketing", DR.MKTG],
    ["gAndA", DR.GA]
  ];
  var hc = data.headcount;
  if (hc && hc.deptDefaults) {
    deptMap.forEach(function (pair) {
      var key = pair[0];
      var row = pair[1];
      var dept = hc.deptDefaults[key];
      if (!dept) return;
      scenarioSetRowColIfDefined_(sh, row, 2, dept.startHC);
      scenarioSetRowColIfDefined_(sh, row, 3, dept.annualSalary);
      scenarioSetRowColIfDefined_(sh, row, 4, dept.swCostPerMo);
      scenarioSetRowColIfDefined_(sh, row, 5, dept.hwCostOneTime);
      scenarioSetRowColIfDefined_(sh, row, 6, dept.insurancePerMo);
    });
  }

  if (hc && Array.isArray(hc.positions)) {
    var positions = hc.positions;
    for (var pi = 0; pi < 10; pi++) {
      var pr = 55 + pi;
      var pos = positions[pi];
      if (pos) {
        sh.getRange(pr, 1).setValue(pos.title || "");
        sh.getRange(pr, 2).setValue(pos.dept || "");
        sh.getRange(pr, 3).setValue(pos.startDate ? (scenarioParseDateForSheet_(pos.startDate) || "") : "");
        sh.getRange(pr, 4).setValue(pos.annualSalary || "");
        sh.getRange(pr, 5).setValue(pos.swCostPerMo || "");
      } else {
        [1, 2, 3, 4, 5].forEach(function (c) { sh.getRange(pr, c).clearContent(); });
      }
    }
  }

  var hs = data.headcountScaling;
  if (hs) {
    scenarioSetIfDefined_(sh, DR.ENG_MM_RATIO, hs.engMmRatio);
    scenarioSetIfDefined_(sh, DR.ENG_ENT_RATIO, hs.engEntRatio);
    scenarioSetIfDefined_(sh, DR.RND_RATIO, hs.rndRatio);
    scenarioSetIfDefined_(sh, DR.SALES_RAMP, hs.salesRampMonths);
    scenarioSetIfDefined_(sh, DR.SALES_REP_CAP, hs.salesRepCapacity);
    scenarioSetIfDefined_(sh, DR.ENT_SALES_WEIGHT, hs.entSalesWeight);
    scenarioSetIfDefined_(sh, DR.AE_RAMP, hs.aeRampMonths);
    scenarioSetIfDefined_(sh, DR.AE_SALARY, hs.aeSalary);
    scenarioSetIfDefined_(sh, DR.AE_SW, hs.aeSwCostMo);
    scenarioSetIfDefined_(sh, DR.GA_RATIO, hs.gaRatio);
    scenarioSetIfDefined_(sh, DR.LOADED_MULT, hs.loadedCostMult);
    scenarioSetIfDefined_(sh, DR.CSM_MM_RATIO, hs.csmMmRatio);
    scenarioSetIfDefined_(sh, DR.CSM_ENT_RATIO, hs.csmEntRatio);
  }

  var eb = data.existingBook;
  if (eb) {
    scenarioSetIfDefined_(sh, DR.EXIST_MM_LOGOS, eb.mmLogos);
    scenarioSetIfDefined_(sh, DR.EXIST_MM_ACV, eb.mmAcv);
    scenarioSetIfDefined_(sh, DR.EXIST_ENT_LOGOS, eb.entLogos);
    scenarioSetIfDefined_(sh, DR.EXIST_ENT_ACV, eb.entAcv);
  }

  var mk = data.marketing;
  if (mk) {
    scenarioSetIfDefined_(sh, DR.EVENTS, mk.eventsAnnual);
    scenarioSetIfDefined_(sh, DR.DIGITAL, mk.digitalAnnual);
    scenarioSetIfDefined_(sh, DR.MKTG_Y2, mk.mktgY2Multiplier);
    scenarioSetIfDefined_(sh, DR.MKTG_SPEND_PER_FTE, mk.mktgSpendPerFte);
  }

  var inf = data.infrastructure;
  if (inf) {
    scenarioSetIfDefined_(sh, DR.INFRA, inf.infraPerCustomerPerMo);
    scenarioSetIfDefined_(sh, DR.TOOLING, inf.toolingPerEngineerPerMo);
  }

  var ox = data.opex;
  if (ox) {
    scenarioSetIfDefined_(sh, DR.RECRUIT_PCT, ox.recruitPctOfSalary);
    scenarioSetIfDefined_(sh, DR.TRAVEL_ENT, ox.travelPerEntDeal);
    scenarioSetIfDefined_(sh, DR.PROF_FEES, ox.profFeesAnnual);
    scenarioSetIfDefined_(sh, DR.CO_SOFTWARE, ox.companySoftwarePerEmpMo);
    scenarioSetIfDefined_(sh, DR.HW_NEW_HIRE, ox.hardwarePerNewHire);
    scenarioSetIfDefined_(sh, DR.TRAVEL_EVENTS_PCT, ox.eventsTravelPct);
  }

  var sl = data.sales;
  if (sl && sl.commission != null) sh.getRange(DR.COMMISSION).setValue(sl.commission);

  var label = meta.name || "Scenario";
  return '✅ "' + label + '" loaded. Run 🚦 Check Benchmarks to validate.';
}

// ─── EXPORT CURRENT STATE ────────────────────────────────────

function getCurrentScenario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("🎛️ Drivers");
  if (!sh) return null;
  function v(row, col) { return sh.getRange(row, col).getValue(); }

  var rounds = [];
  var ri;
  for (ri = 0; ri < 3; ri++) {
    var name = v(123 + ri, 1);
    if (name) {
      rounds.push({
        name: name,
        amount: v(123 + ri, 2),
        closeDate: scenarioFormatDateIsoMaybeEmpty_(v(123 + ri, 3)),
        notes: v(123 + ri, 4)
      });
    }
  }

  var positions = [];
  var j;
  for (j = 0; j < 10; j++) {
    var title = v(55 + j, 1);
    if (title) {
      positions.push({
        title: title,
        dept: v(55 + j, 2),
        startDate: scenarioFormatDateIsoMaybeEmpty_(v(55 + j, 3)),
        annualSalary: v(55 + j, 4),
        swCostPerMo: v(55 + j, 5)
      });
    }
  }

  var lg = sh.getRange(DR.LOGO_GROWTH).getValue();
  if (lg === "" || lg === null || lg === undefined) lg = 0;

  var annualArrTargets = [];
  var yk;
  for (yk = 0; yk < 5; yk++) {
    var rr = 134 + yk;
    annualArrTargets.push({
      targetARR: v(rr, 2),
      targetDate: scenarioFormatDateIsoMaybeEmpty_(v(rr, 3)),
      notes: v(rr, 5) || ""
    });
  }

  return {
    meta: {
      name: "Current Model State",
      forecastHorizon: sh.getRange(DR.HORIZON).getValue()
    },
    timing: {
      forecastStart: scenarioFormatDateIsoMaybeEmpty_(sh.getRange(DR.FORECAST_START).getValue()),
      firstMMClientDate: scenarioFormatDateIsoMaybeEmpty_(sh.getRange(DR.FIRST_MM_CLIENT).getValue()),
      firstEntClientDate: scenarioFormatDateIsoMaybeEmpty_(sh.getRange(DR.FIRST_ENT_CLIENT).getValue())
    },
    annualArrTargets: annualArrTargets,
    arrTargets: {
      targetARR: sh.getRange(DR.TARGET_ARR).getValue(),
      momGrowthRate: sh.getRange(DR.MOM_GROWTH).getValue()
    },
    fundingRounds: rounds,
    fundingMeta: {
      interestRate: sh.getRange(DR.INTEREST_RATE).getValue(),
      openingCash: sh.getRange(DR.OPENING_CASH).getValue()
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
    logoGrowth: lg,
    logoRamp: {
      midMarket: [lg, 0, 0, 0],
      enterprise: [lg, 0, 0, 0]
    },
    fdeCapacity: {
      concurrentMmPerFde: sh.getRange(DR.FDE_MM_CAPACITY).getValue(),
      concurrentEntPerFde: sh.getRange(DR.FDE_ENT_CAPACITY).getValue()
    },
    logoBackCalculation: {
      aeQuotaMm: sh.getRange(DR.AE_QUOTA_MM).getValue(),
      aeQuotaEnt: sh.getRange(DR.AE_QUOTA_ENT).getValue(),
      attainment: sh.getRange(DR.ATTAINMENT).getValue(),
      mmPctOfTargetArr: sh.getRange(DR.MM_PCT_ARR).getValue(),
      logoMomGrowth: sh.getRange(DR.LOGO_GROWTH).getValue()
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
        marketing: {
          startHC: v(DR.MKTG, 2), annualSalary: v(DR.MKTG, 3), swCostPerMo: v(DR.MKTG, 4),
          hwCostOneTime: v(DR.MKTG, 5), insurancePerMo: v(DR.MKTG, 6)
        },
        gAndA: {
          startHC: v(DR.GA, 2), annualSalary: v(DR.GA, 3), swCostPerMo: v(DR.GA, 4),
          hwCostOneTime: v(DR.GA, 5), insurancePerMo: v(DR.GA, 6)
        }
      },
      positions: positions
    },
    headcountScaling: {
      engMmRatio: sh.getRange(DR.ENG_MM_RATIO).getValue(),
      engEntRatio: sh.getRange(DR.ENG_ENT_RATIO).getValue(),
      rndRatio: sh.getRange(DR.RND_RATIO).getValue(),
      salesRampMonths: sh.getRange(DR.SALES_RAMP).getValue(),
      salesRepCapacity: sh.getRange(DR.SALES_REP_CAP).getValue(),
      entSalesWeight: sh.getRange(DR.ENT_SALES_WEIGHT).getValue(),
      aeRampMonths: sh.getRange(DR.AE_RAMP).getValue(),
      aeSalary: sh.getRange(DR.AE_SALARY).getValue(),
      aeSwCostMo: sh.getRange(DR.AE_SW).getValue(),
      gaRatio: sh.getRange(DR.GA_RATIO).getValue(),
      loadedCostMult: sh.getRange(DR.LOADED_MULT).getValue(),
      csmMmRatio: sh.getRange(DR.CSM_MM_RATIO).getValue(),
      csmEntRatio: sh.getRange(DR.CSM_ENT_RATIO).getValue()
    },
    existingBook: {
      mmLogos: sh.getRange(DR.EXIST_MM_LOGOS).getValue(),
      mmAcv: sh.getRange(DR.EXIST_MM_ACV).getValue(),
      entLogos: sh.getRange(DR.EXIST_ENT_LOGOS).getValue(),
      entAcv: sh.getRange(DR.EXIST_ENT_ACV).getValue()
    },
    marketing: {
      eventsAnnual: sh.getRange(DR.EVENTS).getValue(),
      digitalAnnual: sh.getRange(DR.DIGITAL).getValue(),
      mktgY2Multiplier: sh.getRange(DR.MKTG_Y2).getValue(),
      mktgSpendPerFte: sh.getRange(DR.MKTG_SPEND_PER_FTE).getValue()
    },
    infrastructure: {
      infraPerCustomerPerMo: sh.getRange(DR.INFRA).getValue(),
      toolingPerEngineerPerMo: sh.getRange(DR.TOOLING).getValue()
    },
    opex: {
      recruitPctOfSalary: sh.getRange(DR.RECRUIT_PCT).getValue(),
      travelPerEntDeal: sh.getRange(DR.TRAVEL_ENT).getValue(),
      profFeesAnnual: sh.getRange(DR.PROF_FEES).getValue(),
      companySoftwarePerEmpMo: sh.getRange(DR.CO_SOFTWARE).getValue(),
      hardwarePerNewHire: sh.getRange(DR.HW_NEW_HIRE).getValue(),
      eventsTravelPct: sh.getRange(DR.TRAVEL_EVENTS_PCT).getValue()
    },
    sales: {
      commission: sh.getRange(DR.COMMISSION).getValue()
    }
  };
}
