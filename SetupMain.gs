// SetupMain.gs — entry point setupFinancialModel(). Other Setup*.gs modules + ModelConstants.gs.

/** Set true for spreadsheet toasts during each setup phase (interactive runs only). */
var SETUP_PROGRESS_TOAST = false;

function setupLog(phase, detail, elapsedMs) {
  var msg = "[setup] " + new Date().toISOString() + " " + phase;
  if (detail) msg += " | " + detail;
  if (elapsedMs != null) msg += " | " + elapsedMs + "ms";
  Logger.log(msg);
}

function setupFinancialModel() {
  var tAll = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupLog("begin", ss.getName ? ss.getName() : "(spreadsheet)");

  function runPhase(name, fn) {
    var t = Date.now();
    if (SETUP_PROGRESS_TOAST) {
      try {
        ss.toast("Building: " + name + "…", "Tetrix setup", 4);
      } catch (ignore) {}
    }
    setupLog("phase start", name);
    try {
      fn();
    } catch (e) {
      setupLog("FAILED", name + " | " + e.message);
      try {
        SpreadsheetApp.getUi().alert("Setup failed at: " + name + "\n\n" + e.message);
      } catch (e2) {
        Logger.log(e.stack || String(e));
      }
      throw e;
    }
    setupLog("phase end", name, Date.now() - t);
  }

  var TABS = [
    { name: "📖 Instructions", color: "#FFFFFF" },
    { name: "🎛️ Drivers",      color: "#4A90D9" },
    { name: "💰 Funding",      color: "#27AE60" },
    { name: "👥 Headcount",    color: "#8E44AD" },
    { name: "📈 Revenue",      color: "#E67E22" },
    { name: "💸 P&L",          color: "#C0392B" },
    { name: "🏦 Cash Flow",    color: "#16A085" },
    { name: "📊 Summary",      color: "#2C3E50" },
    { name: "🚦 Benchmarks",   color: "#E74C3C" }
  ];

  runPhase("tabs (create/clear)", function() {
    TABS.forEach(function(t) {
      var s = ss.getSheetByName(t.name);
      if (!s) s = ss.insertSheet(t.name);
      else s.clearContents().clearFormats();
      s.setTabColor(t.color);
    });
  });

  runPhase("setupInstructions", function() { setupInstructions(ss); });
  runPhase("setupDrivers", function() { setupDrivers(ss); });
  runPhase("setupFunding", function() { setupFunding(ss); });
  runPhase("setupHeadcount", function() { setupHeadcount(ss); });
  runPhase("setupRevenue_NEW", function() { setupRevenue_NEW(ss); });
  runPhase("setupPnL", function() { setupPnL(ss); });
  runPhase("setupCashFlow", function() { setupCashFlow(ss); });
  runPhase("setupSummary", function() { setupSummary(ss); });
  runPhase("setupBenchmarks", function() { setupBenchmarks(ss); });

  runPhase("reorder sheets", function() {
    ["📖 Instructions","📊 Summary","🎛️ Drivers","💰 Funding",
     "👥 Headcount","📈 Revenue","💸 P&L","🏦 Cash Flow","🚦 Benchmarks"]
      .forEach(function(name, i) {
        var s = ss.getSheetByName(name);
        if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); }
      });
    ss.setActiveSheet(ss.getSheetByName("📖 Instructions"));
  });

  runPhase("cleanup legacy sheets", function() {
    ["📋 Scenarios", "📈 Revenue_v2"].forEach(function(name) {
      var s = ss.getSheetByName(name);
      if (s) ss.deleteSheet(s);
    });
  });

  setupLog("complete", "all phases ok", Date.now() - tAll);
  try { SpreadsheetApp.getUi().alert("✅ Model built! Start in the 🎛️ Drivers tab."); }
  catch (e) { Logger.log("✅ Model built!"); }
}
