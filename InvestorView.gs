// InvestorView.gs — hide/show internal tabs for presentation sharing (same spreadsheet).

/** Tab names hidden when entering investor view (operational / sensitive). */
var INVESTOR_INTERNAL_SHEETS = [
  "🎛️ Drivers",
  "💰 Funding",
  "🚦 Benchmarks"
];

/**
 * Full list of tabs hidden in investor presentation mode: internal sheets plus
 * the intro tab (SHEET_INSTRUCTIONS), replaced by SHEET_INVESTOR_BRIEF for external readers.
 */
function investorPresentationHideList_() {
  var list = INVESTOR_INTERNAL_SHEETS.slice();
  list.push(SHEET_INSTRUCTIONS);
  return list;
}

/**
 * Hides internal + intro tab, ensures investor brief is visible, opens brief tab.
 */
function showInvestorView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hideNames = investorPresentationHideList_();
  var wouldRemain = ss.getSheets().filter(function (s) {
    return hideNames.indexOf(s.getName()) === -1;
  });
  if (wouldRemain.length === 0) {
    SpreadsheetApp.getUi().alert(
      "Investor view",
      "No tabs would stay visible. Check INVESTOR_INTERNAL_SHEETS and SHEET_INVESTOR_BRIEF in ModelConstants / InvestorView.gs.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var hidden = 0;
  hideNames.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh && !sh.isSheetHidden()) {
      sh.hideSheet();
      hidden++;
    }
  });

  var brief = ss.getSheetByName(SHEET_INVESTOR_BRIEF);
  if (brief) {
    if (brief.isSheetHidden()) brief.showSheet();
    ss.setActiveSheet(brief);
  } else {
    var target = ss.getSheetByName("📊 Summary");
    if (target && !target.isSheetHidden()) {
      ss.setActiveSheet(target);
    } else {
      var firstVisible = ss.getSheets().filter(function (s) {
        return !s.isSheetHidden();
      })[0];
      if (firstVisible) ss.setActiveSheet(firstVisible);
    }
  }

  var msg;
  if (hidden) {
    msg = "Investor view: hid " + hidden + " tab(s).";
    msg += brief
      ? " Open " + SHEET_INVESTOR_BRIEF + " for how to read the workbook."
      : " Run setupFinancialModel() to create " + SHEET_INVESTOR_BRIEF + ".";
  } else {
    msg = "Investor view: tabs were already hidden or not found.";
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, APP_SHORT_NAME, 6);
}

/**
 * Restores internal sheets and the intro tab (idempotent).
 */
function showInternalSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = investorPresentationHideList_();
  var shown = 0;
  names.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh && sh.isSheetHidden()) {
      sh.showSheet();
      shown++;
    }
  });
  var drv = ss.getSheetByName("🎛️ Drivers");
  if (drv) ss.setActiveSheet(drv);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    shown ? "Restored " + shown + " tab(s) including " + SHEET_INSTRUCTIONS + "." : "Internal tabs were already visible.",
    APP_SHORT_NAME,
    4
  );
}
