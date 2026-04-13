// SetupRevenue.gs — Revenue tab + _buildRevenue. Called from setupFinancialModel in SetupMain.gs.

function setupRevenue_NEW(ss) {
  var sh = ss.getSheetByName("📈 Revenue");
  if (!sh) sh = ss.insertSheet("📈 Revenue");
  else { sh.clearContents(); sh.clearFormats(); }
  sh.setTabColor("#E67E22");
  _buildRevenue(sh);
}

function _buildRevenue(sh) {
  var MAX  = 60;
  var HOR  = "'🎛️ Drivers'!" + DR.HORIZON;
  var FS   = "'🎛️ Drivers'!" + DR.FORECAST_START;
  var TCOL = REVCOLS.TOTAL;

  if (sh.getMaxColumns() < TCOL)
    sh.insertColumnsAfter(sh.getMaxColumns(), TCOL - sh.getMaxColumns());

  sh.setColumnWidth(REVCOLS.MO_NUM, 42);
  sh.setColumnWidth(REVCOLS.MONTH,  72);
  sh.setColumnWidth(REVCOLS.MM_LOGOS,  68);
  sh.setColumnWidth(REVCOLS.ENT_LOGOS, 68);
  for (var c = REVCOLS.MM_NEW_ARR;   c <= REVCOLS.MM_MRR;      c++) sh.setColumnWidth(c, 88);
  for (var c = REVCOLS.ENT_NEW_ARR;  c <= REVCOLS.ENT_MRR;     c++) sh.setColumnWidth(c, 88);
  for (var c = REVCOLS.EX_MM_CHURN;  c <= REVCOLS.EX_MM_MRR;   c++) sh.setColumnWidth(c, 88);
  for (var c = REVCOLS.EX_ENT_CHURN; c <= REVCOLS.EX_ENT_MRR;  c++) sh.setColumnWidth(c, 88);
  for (var c = REVCOLS.CB_NET_NEW;   c <= REVCOLS.CB_MRR;      c++) sh.setColumnWidth(c, 100);

  sh.getRange(REVROWS.HDR_TITLE, 1, 1, TCOL).merge()
    .setValue("📈 REVENUE — ARR Waterfall  |  All inputs in 🎛️ Drivers  |  Do not edit formulas")
    .setBackground("#1A5276").setFontColor("#FFFFFF")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
  sh.setRowHeight(REVROWS.HDR_TITLE, 28);

  sh.getRange(REVROWS.HDR_GROUPS, REVCOLS.MO_NUM, 1, 2).setBackground("#2C3E50");
  var groups = [
    { col: REVCOLS.MM_LOGOS,     span: 7, label: "📊 MID-MARKET — New Logos",  bg: "#1F618D" },
    { col: REVCOLS.ENT_LOGOS,    span: 7, label: "📊 ENTERPRISE — New Logos",  bg: "#1A5276" },
    { col: REVCOLS.EX_MM_CHURN,  span: 5, label: "📋 EXISTING MID-MARKET",     bg: "#6C3483" },
    { col: REVCOLS.EX_ENT_CHURN, span: 5, label: "📋 EXISTING ENTERPRISE",     bg: "#4A235A" },
    { col: REVCOLS.CB_NET_NEW,   span: 3, label: "✅ COMBINED",                bg: "#1D6A47" }
  ];
  groups.forEach(function(g) {
    sh.getRange(REVROWS.HDR_GROUPS, g.col, 1, g.span).merge()
      .setValue(g.label).setBackground(g.bg).setFontColor("#FFFFFF")
      .setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center");
  });
  sh.setRowHeight(REVROWS.HDR_GROUPS, 24);

  var subHdrs = [
    [REVCOLS.MO_NUM,       "Mo #",          "#2C3E50"],
    [REVCOLS.MONTH,        "Month",         "#2C3E50"],
    [REVCOLS.MM_LOGOS,     "New Logos",     "#1F618D"],
    [REVCOLS.MM_NEW_ARR,   "New ARR ($)",   "#1F618D"],
    [REVCOLS.MM_CHURN,     "Churn ARR ($)", "#1F618D"],
    [REVCOLS.MM_EXP,       "Expansion ($)", "#1F618D"],
    [REVCOLS.MM_NET_NEW,   "Net New ($)",   "#1F618D"],
    [REVCOLS.MM_CUMUL,     "Cumul ARR ($)", "#1F618D"],
    [REVCOLS.MM_MRR,       "MRR ($)",       "#1F618D"],
    [REVCOLS.ENT_LOGOS,    "New Logos",     "#1A5276"],
    [REVCOLS.ENT_NEW_ARR,  "New ARR ($)",   "#1A5276"],
    [REVCOLS.ENT_CHURN,    "Churn ARR ($)", "#1A5276"],
    [REVCOLS.ENT_EXP,      "Expansion ($)", "#1A5276"],
    [REVCOLS.ENT_NET_NEW,  "Net New ($)",   "#1A5276"],
    [REVCOLS.ENT_CUMUL,    "Cumul ARR ($)", "#1A5276"],
    [REVCOLS.ENT_MRR,      "MRR ($)",       "#1A5276"],
    [REVCOLS.EX_MM_CHURN,  "Churn ARR ($)", "#6C3483"],
    [REVCOLS.EX_MM_EXP,    "Expansion ($)", "#6C3483"],
    [REVCOLS.EX_MM_NET,    "Net ARR ($)",   "#6C3483"],
    [REVCOLS.EX_MM_CUMUL,  "Cumul ARR ($)", "#6C3483"],
    [REVCOLS.EX_MM_MRR,    "MRR ($)",       "#6C3483"],
    [REVCOLS.EX_ENT_CHURN, "Churn ARR ($)", "#4A235A"],
    [REVCOLS.EX_ENT_EXP,   "Expansion ($)", "#4A235A"],
    [REVCOLS.EX_ENT_NET,   "Net ARR ($)",   "#4A235A"],
    [REVCOLS.EX_ENT_CUMUL, "Cumul ARR ($)", "#4A235A"],
    [REVCOLS.EX_ENT_MRR,   "MRR ($)",       "#4A235A"],
    [REVCOLS.CB_NET_NEW,   "Net New ($)",   "#1D6A47"],
    [REVCOLS.CB_CUMUL,     "Cumul ARR ($)", "#1D6A47"],
    [REVCOLS.CB_MRR,       "MRR ($)",       "#1D6A47"]
  ];
  subHdrs.forEach(function(h) {
    sh.getRange(REVROWS.HDR_COLS, h[0])
      .setValue(h[1]).setBackground(h[2]).setFontColor("#FFFFFF")
      .setFontWeight("bold").setHorizontalAlignment("center").setWrap(true);
  });
  sh.setRowHeight(REVROWS.HDR_COLS, 32);

  for (var m = 1; m <= MAX; m++) {
    var row = REVROWS.DATA_START + m - 1;
    var bg  = m % 2 === 0 ? "#F8F9FA" : "#FFFFFF";
    sh.setRowHeight(row, 18);
    sh.getRange(row, REVCOLS.MO_NUM)
      .setFormula("=IF(" + m + ">" + HOR + ",\"\"," + m + ")")
      .setHorizontalAlignment("center").setBackground(bg);
    sh.getRange(row, REVCOLS.MONTH)
      .setFormula("=IF(" + m + ">" + HOR + ",\"\",TEXT(EDATE(" + FS + "," + (m-1) + "),\"MMM YY\"))")
      .setNumberFormat("@").setBackground(bg);
  }

  var segs3 = [
    {
      dRow: 18, firstDate: DR.FIRST_MM_DATE,
      pct: "'🎛️ Drivers'!" + DR.MM_PCT_ARR,
      logosCol: REVCOLS.MM_LOGOS,   newArrCol: REVCOLS.MM_NEW_ARR,
      churnCol: REVCOLS.MM_CHURN,   expCol:    REVCOLS.MM_EXP,
      netNewCol:REVCOLS.MM_NET_NEW, cumulCol:  REVCOLS.MM_CUMUL,
      mrrCol:   REVCOLS.MM_MRR,     exCumulCol:REVCOLS.EX_MM_CUMUL
    },
    {
      dRow: 19, firstDate: DR.FIRST_ENT_DATE,
      pct: "(1-'🎛️ Drivers'!" + DR.MM_PCT_ARR + ")",
      logosCol: REVCOLS.ENT_LOGOS,   newArrCol: REVCOLS.ENT_NEW_ARR,
      churnCol: REVCOLS.ENT_CHURN,   expCol:    REVCOLS.ENT_EXP,
      netNewCol:REVCOLS.ENT_NET_NEW, cumulCol:  REVCOLS.ENT_CUMUL,
      mrrCol:   REVCOLS.ENT_MRR,     exCumulCol:REVCOLS.EX_ENT_CUMUL
    }
  ];

  segs3.forEach(function(seg) {
    var off    = "'🎛️ Drivers'!" + seg.firstDate;
    var netRet = "(1-'🎛️ Drivers'!D" + seg.dRow + "/12+'🎛️ Drivers'!E" + seg.dRow + "/12)";
    var churn  = "'🎛️ Drivers'!D" + seg.dRow;
    var exp    = "'🎛️ Drivers'!E" + seg.dRow;
    var expMo  = "'🎛️ Drivers'!F" + seg.dRow;
    var acv    = "'🎛️ Drivers'!B" + seg.dRow;
    var grow   = "'🎛️ Drivers'!" + DR.LOGO_GROWTH;
    var exColLet = colLetter(seg.exCumulCol);
    var moColLet = colLetter(REVCOLS.MO_NUM);
    var dsRow    = REVROWS.DATA_START;
    var deRow    = REVROWS.DATA_START + MAX - 1;
    var existAtH = "IFERROR(INDEX(" + exColLet + dsRow + ":" + exColLet + deRow
                 + ",MATCH(" + HOR + "," + moColLet + dsRow + ":" + moColLet + deRow + ",0)),0)";
    var segTgt = "MAX(0,'🎛️ Drivers'!" + DR.TARGET_ARR + "*" + seg.pct + "-" + existAtH + ")";
    var actTot = "MAX(1," + HOR + "-(" + off + "))";
    var ratio  = "(1+" + grow + ")/(" + netRet + ")";
    var gsum   = "IF(ABS(" + ratio + "-1)<0.0001," + actTot
               + ",(POWER(" + ratio + "," + actTot + ")-1)/(" + ratio + "-1))";
    var base   = "IF('🎛️ Drivers'!B" + seg.dRow + ">0," + segTgt
               + "/('🎛️ Drivers'!B" + seg.dRow + "*POWER(" + netRet + "," + actTot + "-1)*" + gsum + "),0)";
    function cumF(n) {
      return "(" + base + ")*(IF(ABS(" + grow + ")<0.0001,(" + n + ")"
           + ",(POWER(1+(" + grow + "),(" + n + "))-1)/(" + grow + ")))";
    }
    var cL  = colLetter(seg.cumulCol);
    var lL  = colLetter(seg.logosCol);
    var nL  = colLetter(seg.newArrCol);
    var chL = colLetter(seg.churnCol);
    var eL  = colLetter(seg.expCol);
    var nnL = colLetter(seg.netNewCol);
    var mL  = colLetter(seg.mrrCol);

    for (var m = 1; m <= MAX; m++) {
      var row     = REVROWS.DATA_START + m - 1;
      var prevRow = REVROWS.DATA_START + m - 2;
      var mAct     = "MAX(0," + m + "-(" + off + "))";
      var mActPrev = "MAX(0," + (m-1) + "-(" + off + "))";
      var rndCur  = "MAX(0,CEILING(IF((" + mAct + ")<=0,0," + cumF(mAct) + "),1))";
      var rndPrev = "MAX(0,CEILING(IF((" + mActPrev + ")<=0,0," + cumF(mActPrev) + "),1))";
      sh.getRange(row, seg.logosCol)
        .setFormula("=IF(" + m + ">" + HOR + ",\"\",MAX(0," + rndCur + "-" + rndPrev + "))").setNumberFormat("0");
      sh.getRange(row, seg.newArrCol)
        .setFormula("=IF(" + m + ">" + HOR + ",\"\"," + lL + row + "*" + acv + ")").setNumberFormat("$#,##0");
      if (m === 1) {
        sh.getRange(row, seg.churnCol).setFormula("=IF(1>" + HOR + ",\"\",0)").setNumberFormat("$#,##0");
        sh.getRange(row, seg.expCol  ).setFormula("=IF(1>" + HOR + ",\"\",0)").setNumberFormat("$#,##0");
        sh.getRange(row, seg.cumulCol)
          .setFormula("=IF(1>" + HOR + ",\"\"," + nnL + row + ")")
          .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");
      } else {
        sh.getRange(row, seg.churnCol)
          .setFormula("=IF(" + m + ">" + HOR + ",\"\",-" + cL + prevRow + "*(" + churn + "/12))")
          .setNumberFormat("$#,##0").setFontColor("#922B21");
        sh.getRange(row, seg.expCol)
          .setFormula("=IF(" + m + ">" + HOR + ",\"\",IF(" + m + ">=" + expMo + "," + cL + prevRow + "*(" + exp + "/12),0))")
          .setNumberFormat("$#,##0").setFontColor("#1D9E75");
        sh.getRange(row, seg.cumulCol)
          .setFormula("=IF(" + m + ">" + HOR + ",\"\"," + cL + prevRow + "+" + nnL + row + ")")
          .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");
      }
      sh.getRange(row, seg.netNewCol)
        .setFormula("=IF(" + m + ">" + HOR + ",\"\"," + nL + row + "+" + chL + row + "+" + eL + row + ")").setNumberFormat("$#,##0");
      sh.getRange(row, seg.mrrCol)
        .setFormula("=IF(" + m + ">" + HOR + ",\"\"," + cL + row + "/12)").setNumberFormat("$#,##0");
    }
  });

  var existSegs4 = [
    {
      initArr: "'🎛️ Drivers'!" + DR.EXIST_MM_ARR, churnRate: "'🎛️ Drivers'!D" + DR.MM_ROW,
      expRate: "'🎛️ Drivers'!E" + DR.MM_ROW,
      churnCol: REVCOLS.EX_MM_CHURN, expCol: REVCOLS.EX_MM_EXP,
      netCol: REVCOLS.EX_MM_NET, cumulCol: REVCOLS.EX_MM_CUMUL, mrrCol: REVCOLS.EX_MM_MRR
    },
    {
      initArr: "'🎛️ Drivers'!" + DR.EXIST_ENT_ARR, churnRate: "'🎛️ Drivers'!D" + DR.ENT_ROW,
      expRate: "'🎛️ Drivers'!E" + DR.ENT_ROW,
      churnCol: REVCOLS.EX_ENT_CHURN, expCol: REVCOLS.EX_ENT_EXP,
      netCol: REVCOLS.EX_ENT_NET, cumulCol: REVCOLS.EX_ENT_CUMUL, mrrCol: REVCOLS.EX_ENT_MRR
    }
  ];
  existSegs4.forEach(function(seg) {
    var cL  = colLetter(seg.cumulCol);
    var chL = colLetter(seg.churnCol);
    var eL  = colLetter(seg.expCol);
    var nL  = colLetter(seg.netCol);
    for (var m = 1; m <= MAX; m++) {
      var row = REVROWS.DATA_START + m - 1; var prevRow = REVROWS.DATA_START + m - 2; var ms = String(m);
      if (m === 1) {
        sh.getRange(row, seg.churnCol).setFormula("=IF(1>" + HOR + ",\"\",0)").setNumberFormat("$#,##0");
        sh.getRange(row, seg.expCol  ).setFormula("=IF(1>" + HOR + ",\"\",0)").setNumberFormat("$#,##0");
        sh.getRange(row, seg.netCol  ).setFormula("=IF(1>" + HOR + ",\"\",0)").setNumberFormat("$#,##0");
        sh.getRange(row, seg.cumulCol)
          .setFormula("=IF(1>" + HOR + ",\"\"," + seg.initArr + ")")
          .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");
      } else {
        sh.getRange(row, seg.churnCol)
          .setFormula("=IF(" + ms + ">" + HOR + ",\"\",-" + cL + prevRow + "*(" + seg.churnRate + "/12))")
          .setNumberFormat("$#,##0").setFontColor("#922B21");
        sh.getRange(row, seg.expCol)
          .setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + cL + prevRow + "*(" + seg.expRate + "/12))")
          .setNumberFormat("$#,##0").setFontColor("#1D9E75");
        sh.getRange(row, seg.netCol)
          .setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + chL + row + "+" + eL + row + ")").setNumberFormat("$#,##0");
        sh.getRange(row, seg.cumulCol)
          .setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + cL + prevRow + "+" + nL + row + ")")
          .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");
      }
      sh.getRange(row, seg.mrrCol)
        .setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + cL + row + "/12)").setNumberFormat("$#,##0");
    }
  });

  var nnL_mm  = colLetter(REVCOLS.MM_NET_NEW);
  var nnL_ent = colLetter(REVCOLS.ENT_NET_NEW);
  var nnL_exM = colLetter(REVCOLS.EX_MM_NET);
  var nnL_exE = colLetter(REVCOLS.EX_ENT_NET);
  var cL_mm   = colLetter(REVCOLS.MM_CUMUL);
  var cL_ent  = colLetter(REVCOLS.ENT_CUMUL);
  var cL_exM  = colLetter(REVCOLS.EX_MM_CUMUL);
  var cL_exE  = colLetter(REVCOLS.EX_ENT_CUMUL);
  var cbNNL   = colLetter(REVCOLS.CB_NET_NEW);
  var cbCL    = colLetter(REVCOLS.CB_CUMUL);
  var cbML    = colLetter(REVCOLS.CB_MRR);
  var moL     = colLetter(REVCOLS.MO_NUM);
  var dsRow   = REVROWS.DATA_START;
  var deRow   = REVROWS.DATA_START + MAX - 1;

  for (var m = 1; m <= MAX; m++) {
    var row = REVROWS.DATA_START + m - 1; var ms = String(m);
    sh.getRange(row, REVCOLS.CB_NET_NEW)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\","
        + nnL_mm + row + "+" + nnL_ent + row + "+" + nnL_exM + row + "+" + nnL_exE + row + ")")
      .setNumberFormat("$#,##0");
    sh.getRange(row, REVCOLS.CB_CUMUL)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\","
        + cL_mm + row + "+" + cL_ent + row + "+" + cL_exM + row + "+" + cL_exE + row + ")")
      .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");
    sh.getRange(row, REVCOLS.CB_MRR)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + cbCL + row + "/12)")
      .setNumberFormat("$#,##0").setFontWeight("bold");
  }

  var lastH = "IFERROR(INDEX(" + cbCL + dsRow + ":" + cbCL + deRow
            + ",MATCH(MAX(" + moL + dsRow + ":" + moL + deRow + "),"
            + moL + dsRow + ":" + moL + deRow + ",0)),0)";

  sh.getRange(REVROWS.TV_HDR, 1, 1, TCOL).merge()
    .setValue("📊 TARGET vs ACTUAL — Combined Cumul ARR vs Target ARR  |  Gap < 5% is healthy")
    .setBackground("#D6EAF8").setFontWeight("bold").setFontSize(10);
  sh.setRowHeight(REVROWS.TV_HDR, 24);
  sh.getRange(REVROWS.TV_TARGET, 1).setValue("Target ARR (Drivers B12)").setFontWeight("bold");
  sh.getRange(REVROWS.TV_TARGET, REVCOLS.CB_NET_NEW).setValue("Target →").setFontColor("#888").setFontStyle("italic");
  sh.getRange(REVROWS.TV_TARGET, REVCOLS.CB_CUMUL)
    .setFormula("='🎛️ Drivers'!" + DR.TARGET_ARR).setNumberFormat("$#,##0").setBackground("#FEF9E7").setFontWeight("bold");
  sh.getRange(REVROWS.TV_ACTUAL, 1).setValue("Actual Final Cumul ARR").setFontWeight("bold");
  sh.getRange(REVROWS.TV_ACTUAL, REVCOLS.CB_NET_NEW).setValue("Actual →").setFontColor("#888").setFontStyle("italic");
  sh.getRange(REVROWS.TV_ACTUAL, REVCOLS.CB_CUMUL)
    .setFormula("=" + lastH).setNumberFormat("$#,##0").setBackground("#D5F5E3").setFontWeight("bold");
  sh.getRange(REVROWS.TV_ACTUAL, REVCOLS.CB_MRR)
    .setFormula("=IFERROR(" + lastH + "/'🎛️ Drivers'!" + DR.TARGET_ARR + "-1,\"\")").setNumberFormat("0.0%").setFontWeight("bold");
  sh.getRange(REVROWS.TV_GAP, 1).setValue("Gap (Actual − Target)").setFontWeight("bold");
  sh.getRange(REVROWS.TV_GAP, REVCOLS.CB_NET_NEW).setValue("Gap →").setFontColor("#888").setFontStyle("italic");
  sh.getRange(REVROWS.TV_GAP, REVCOLS.CB_CUMUL)
    .setFormula("=" + lastH + "-'🎛️ Drivers'!" + DR.TARGET_ARR).setNumberFormat("$#,##0").setFontWeight("bold");
  sh.getRange(REVROWS.TV_GAP, REVCOLS.CB_MRR)
    .setValue("Gap < 5% is normal. Adjust Logo Growth Rate or MM % in Drivers to minimize.")
    .setFontStyle("italic").setFontColor("#888").setWrap(true);

  var gapCell = sh.getRange(REVROWS.TV_GAP, REVCOLS.CB_CUMUL);
  var p5cfR   = sh.getConditionalFormatRules();
  p5cfR.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=ABS(" + cbCL + REVROWS.TV_GAP + ")<=0.05*" + cbCL + REVROWS.TV_TARGET)
    .setBackground("#D5F5E3").setRanges([gapCell]).build());
  p5cfR.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=ABS(" + cbCL + REVROWS.TV_GAP + ")>0.05*" + cbCL + REVROWS.TV_TARGET)
    .setBackground("#FADBD8").setFontColor("#922B21").setRanges([gapCell]).build());
  var cbCumulRange = sh.getRange(REVROWS.DATA_START, REVCOLS.CB_CUMUL, MAX, 1);
  p5cfR.push(SpreadsheetApp.newConditionalFormatRule()
    .whenCellNotEmpty().setBackground("#D5F5E3").setRanges([cbCumulRange]).build());
  sh.setConditionalFormatRules(p5cfR);

  sh.getRange(REVROWS.NOTE, 1, 1, TCOL).merge()
    .setValue("📈 Revenue — horizontal ARR waterfall. All formulas auto-calculate from 🎛️ Drivers — do not edit this tab.")
    .setFontStyle("italic").setFontColor("#1D9E75").setBackground("#FDFEFE").setWrap(true);
  sh.setRowHeight(REVROWS.NOTE, 28);

  sh.setRowHeight(REVROWS.HDR_TITLE,  28);
  sh.setRowHeight(REVROWS.HDR_GROUPS, 26);
  sh.setRowHeight(REVROWS.HDR_COLS,   34);
  for (var m = 1; m <= MAX; m++) sh.setRowHeight(REVROWS.DATA_START + m - 1, 18);
  sh.setRowHeight(REVROWS.TV_HDR,    26);
  sh.setRowHeight(REVROWS.TV_TARGET, 22);
  sh.setRowHeight(REVROWS.TV_ACTUAL, 22);
  sh.setRowHeight(REVROWS.TV_GAP,    22);
  sh.setRowHeight(REVROWS.NOTE,      28);

  var groupFirstCols = [
    REVCOLS.MM_LOGOS, REVCOLS.ENT_LOGOS,
    REVCOLS.EX_MM_CHURN, REVCOLS.EX_ENT_CHURN, REVCOLS.CB_NET_NEW
  ];
  groupFirstCols.forEach(function(col) {
    sh.getRange(REVROWS.HDR_GROUPS, col, 2, 1)
      .setBorder(false, true, false, false, false, false,
        "#2C3E50", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  var horizonVal = Math.min(
    Math.max(
      SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName("🎛️ Drivers")
        .getRange(DR.HORIZON).getValue() || MAX,
      1),
    MAX);
  var dataDividerCols = [
    REVCOLS.ENT_LOGOS,
    REVCOLS.EX_MM_CHURN,
    REVCOLS.EX_ENT_CHURN,
    REVCOLS.CB_NET_NEW
  ];
  dataDividerCols.forEach(function(col) {
    sh.getRange(REVROWS.DATA_START, col, horizonVal, 1)
      .setBorder(false, true, false, false, false, false,
        "#AAAAAA", SpreadsheetApp.BorderStyle.SOLID);
  });

  var p7Rules = sh.getConditionalFormatRules();

  var mrrCols = [
    REVCOLS.MM_MRR, REVCOLS.ENT_MRR,
    REVCOLS.EX_MM_MRR, REVCOLS.EX_ENT_MRR, REVCOLS.CB_MRR
  ];
  mrrCols.forEach(function(col) {
    var r = sh.getRange(REVROWS.DATA_START, col, MAX, 1);
    r.setBackground(null);
    p7Rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty().setBackground("#E8F8F5").setRanges([r]).build());
    sh.getRange(REVROWS.HDR_COLS, col).setBackground(
      col === REVCOLS.MM_MRR     ? "#1F618D" :
      col === REVCOLS.ENT_MRR    ? "#1A5276" :
      col === REVCOLS.EX_MM_MRR  ? "#6C3483" :
      col === REVCOLS.EX_ENT_MRR ? "#4A235A" : "#1D6A47"
    );
  });

  var cumulCols = [
    REVCOLS.MM_CUMUL, REVCOLS.ENT_CUMUL,
    REVCOLS.EX_MM_CUMUL, REVCOLS.EX_ENT_CUMUL, REVCOLS.CB_CUMUL
  ];
  cumulCols.forEach(function(col) {
    var r = sh.getRange(REVROWS.DATA_START, col, MAX, 1);
    r.setBackground(null).setFontWeight("bold");
    p7Rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty().setBackground("#D5F5E3").setRanges([r]).build());
  });

  sh.setConditionalFormatRules(p7Rules);

  var netNewCols = [
    REVCOLS.MM_NET_NEW, REVCOLS.ENT_NET_NEW,
    REVCOLS.EX_MM_NET, REVCOLS.EX_ENT_NET,
    REVCOLS.CB_NET_NEW
  ];
  netNewCols.forEach(function(col) {
    sh.getRange(REVROWS.DATA_START, col, MAX, 1).setFontWeight("bold");
  });

  [REVCOLS.MM_LOGOS, REVCOLS.ENT_LOGOS].forEach(function(col) {
    sh.getRange(REVROWS.DATA_START, col, MAX, 1).setHorizontalAlignment("center");
  });

  sh.getRange(REVROWS.DATA_START, REVCOLS.MO_NUM, MAX, 1)
    .setHorizontalAlignment("center");

  sh.getRange(REVROWS.TV_HDR,    1).setFontSize(10);
  sh.getRange(REVROWS.TV_TARGET, 1).setFontWeight("bold").setFontColor("#444444");
  sh.getRange(REVROWS.TV_ACTUAL, 1).setFontWeight("bold").setFontColor("#444444");
  sh.getRange(REVROWS.TV_GAP,    1).setFontWeight("bold").setFontColor("#444444");

  for (var m = 6; m <= MAX; m += 6) {
    sh.getRange(REVROWS.DATA_START + m - 1, 1, 1, TCOL)
      .setBorder(true, false, false, false, false, false,
        "#DDDDDD", SpreadsheetApp.BorderStyle.SOLID);
  }
}
