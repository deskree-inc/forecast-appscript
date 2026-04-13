// SetupSummary.gs — Called from setupFinancialModel in SetupMain.gs.

function setupSummary(ss) {
  var sh   = ss.getSheetByName("📊 Summary");
  var COLS = 6;

  sh.setColumnWidth(1, 260);
  [2,3,4,5].forEach(function(c){ sh.setColumnWidth(c, 115); });
  sh.setColumnWidth(6, 210);

  function secHdr(row, text) {
    sh.getRange(row, 1, 1, COLS).merge()
      .setValue(text).setBackground("#2C3E50").setFontColor("#FFFFFF")
      .setFontWeight("bold").setFontSize(10);
    sh.setRowHeight(row, 24);
  }
  function rowLbl(row, text, bold) {
    sh.getRange(row, 1).setValue(text)
      .setFontWeight(bold ? "bold" : "normal").setFontColor("#444444");
    sh.setRowHeight(row, 20);
  }
  function bench(row, text) {
    sh.getRange(row, 6).setValue(text)
      .setFontColor("#7D6608").setFontStyle("italic")
      .setBackground("#FEF9E7").setWrap(true);
  }

  sh.getRange(SUM.HDR_TITLE, 1, 1, COLS).merge()
    .setValue("📊 SUMMARY — Investor KPI Dashboard")
    .setBackground("#1A5276").setFontColor("#FFFFFF")
    .setFontWeight("bold").setFontSize(13)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(SUM.HDR_TITLE, 36);

  var HOR  = "'🎛️ Drivers'!" + DR.HORIZON;
  var FS   = "'🎛️ Drivers'!" + DR.FORECAST_START;
  var TARG = "'🎛️ Drivers'!" + DR.TARGET_ARR;

  sh.getRange(SUM.HDR_CONTEXT, 1, 1, COLS).merge()
    .setFormula("=TEXT("+FS+",\"MMM YYYY\")&\" → \"&TEXT(EDATE("+FS+","+HOR+"-1),\"MMM YYYY\")&\"   |   \"&"+HOR+"&\" months   |   Target ARR: \"&TEXT("+TARG+",\"$#,##0\")")
    .setBackground("#2C3E50").setFontColor("#BDC3C7")
    .setFontStyle("italic").setFontSize(9)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setNumberFormat("@");
  sh.setRowHeight(SUM.HDR_CONTEXT, 22);

  sh.getRange(SUM.HDR_MO, 1).setValue("Milestone →")
    .setFontColor("#AAAAAA").setFontStyle("italic").setFontSize(8);
  sh.getRange(SUM.HDR_MO, 2).setValue(1)
    .setBackground("#F2F3F4").setFontColor("#888888")
    .setFontSize(8).setHorizontalAlignment("center").setNumberFormat("0");
  sh.getRange(SUM.HDR_MO, 3)
    .setFormula("=MAX(1,MIN("+HOR+",IF(ISNUMBER('🎛️ Drivers'!"+DR.ARR_Y1_DATE+"),(YEAR('🎛️ Drivers'!"+DR.ARR_Y1_DATE+")-YEAR("+FS+"))*12+MONTH('🎛️ Drivers'!"+DR.ARR_Y1_DATE+")-MONTH("+FS+")+1,MAX(2,ROUND("+HOR+"*0.25,0)))))")
    .setBackground("#F2F3F4").setFontColor("#888888")
    .setFontSize(8).setHorizontalAlignment("center").setNumberFormat("0");
  sh.getRange(SUM.HDR_MO, 4)
    .setFormula("=MAX(1,MIN("+HOR+",IF(ISNUMBER('🎛️ Drivers'!"+DR.ARR_Y2_DATE+"),(YEAR('🎛️ Drivers'!"+DR.ARR_Y2_DATE+")-YEAR("+FS+"))*12+MONTH('🎛️ Drivers'!"+DR.ARR_Y2_DATE+")-MONTH("+FS+")+1,MAX(3,MIN("+HOR+"-1,ROUND("+HOR+"*0.5,0))))))")
    .setBackground("#F2F3F4").setFontColor("#888888")
    .setFontSize(8).setHorizontalAlignment("center").setNumberFormat("0");
  sh.getRange(SUM.HDR_MO, 5)
    .setFormula("="+HOR)
    .setBackground("#F2F3F4").setFontColor("#888888")
    .setFontSize(8).setHorizontalAlignment("center").setNumberFormat("0");
  sh.getRange(SUM.HDR_MO, 6).setBackground("#F2F3F4");
  sh.setRowHeight(SUM.HDR_MO, 18);

  hdr(sh, SUM.HDR_DATES, 1, "Line Item", "#1F618D");
  [2,3,4,5].forEach(function(col) {
    var moCell = colLetter(col) + SUM.HDR_MO;
    sh.getRange(SUM.HDR_DATES, col)
      .setFormula("=IF("+moCell+"=\"\",\"\",TEXT(EDATE("+FS+","+moCell+"-1),\"MMM YY\"))")
      .setBackground("#1F618D").setFontColor("#FFFFFF")
      .setFontWeight("bold").setHorizontalAlignment("center")
      .setNumberFormat("@");
  });
  hdr(sh, SUM.HDR_DATES, 6, "Benchmark / Note", "#1F618D");
  sh.setRowHeight(SUM.HDR_DATES, 22);

  secHdr(SUM.SEC_REVENUE, "📈 REVENUE");
  rowLbl(SUM.ARR,               "ARR ($)",                true); bench(SUM.ARR,               "Target ARR in Drivers B12");
  rowLbl(SUM.MRR,               "MRR ($)",                true);
  rowLbl(SUM.YOY_GROWTH,        "YoY ARR Growth");               bench(SUM.YOY_GROWTH,        "Bessemer: 3× Yr 1, 2× Yr 2");
  rowLbl(SUM.NET_NEW_ARR,       "Net New ARR ($)");
  rowLbl(SUM.MOM_GROWTH_IMPLIED,"Implied MoM ARR Growth");       bench(SUM.MOM_GROWTH_IMPLIED,"Compare to logo growth (B39) and MoM target (B13)");
  rowLbl(SUM.LOGOS,             "Total Active Customers", true);

  secHdr(SUM.SEC_QUALITY, "⭐ REVENUE QUALITY");
  rowLbl(SUM.NRR,              "Net Revenue Retention (NRR %)", true); bench(SUM.NRR,              "Best-in-class: > 120%");
  rowLbl(SUM.GROSS_MARGIN,     "Gross Margin %",                true); bench(SUM.GROSS_MARGIN,     "Series A benchmark: > 65%");
  rowLbl(SUM.CHURN_MM,         "  ↳ MM Annual Churn %");               bench(SUM.CHURN_MM,         "Driver — Section C");
  rowLbl(SUM.CHURN_ENT,        "  ↳ ENT Annual Churn %");              bench(SUM.CHURN_ENT,        "Driver — Section C");
  rowLbl(SUM.ARR_PER_LOGO_MM,  "  ↳ ARR per MM Logo ($)");             bench(SUM.ARR_PER_LOGO_MM,  "Cumul MM ARR ÷ active MM logos. Shows ACV expansion.");
  rowLbl(SUM.ARR_PER_LOGO_ENT, "  ↳ ARR per ENT Logo ($)");            bench(SUM.ARR_PER_LOGO_ENT, "Cumul ENT ARR ÷ active ENT logos. Shows ACV expansion.");

  secHdr(SUM.SEC_BURN, "🔥 BURN & RUNWAY");
  rowLbl(SUM.END_CASH,   "Ending Cash ($)",      true); bench(SUM.END_CASH,   "Must cover ≥ 18 months runway");
  rowLbl(SUM.BURN_RATE,  "Monthly Net Burn ($)", true);
  rowLbl(SUM.RUNWAY,     "Runway (months)",      true); bench(SUM.RUNWAY,     "Cash ÷ 3-mo avg |burn|. Red < 6mo  |  Yellow 6–12mo  |  Green > 12mo");
  rowLbl(SUM.CUMUL_BURN, "Cumulative Burn ($)");

  secHdr(SUM.SEC_FUNDING, "💰 FUNDING");
  rowLbl(SUM.CUMUL_CAP,    "Cumulative Capital Raised ($)", true);
  rowLbl(SUM.PLANNED_RAISE,"Planned Raise ($)");                  bench(SUM.PLANNED_RAISE,"Next round — from Drivers K");
  rowLbl(SUM.CASH_ARR,     "Cash as % of ARR");                   bench(SUM.CASH_ARR,     "Investors expect > 100% (≥ 12mo MRR)");

  secHdr(SUM.SEC_EFFICIENCY, "⚙️ EFFICIENCY");
  rowLbl(SUM.BURN_MULTIPLE, "Burn Multiple (x)",    true); bench(SUM.BURN_MULTIPLE, "< 1x efficient  |  1–2x watch  |  > 2x flag");
  rowLbl(SUM.RULE_OF_40,    "Rule of 40 (%)",       true); bench(SUM.RULE_OF_40,    "> 40% healthy. Live from Month 13.");
  rowLbl(SUM.ARR_PER_EMP,   "ARR per Employee ($)", true); bench(SUM.ARR_PER_EMP,   "Series A benchmark: $150K–$200K");
  rowLbl(SUM.LTV_CAC_MM,    "LTV:CAC — Mid-Market", true); bench(SUM.LTV_CAC_MM,    "Best-in-class: > 3×  |  = (ACV ÷ Churn) ÷ CAC");
  rowLbl(SUM.LTV_CAC_ENT,   "LTV:CAC — Enterprise", true); bench(SUM.LTV_CAC_ENT,   "Best-in-class: > 3×  |  = (ACV ÷ Churn) ÷ CAC");

  secHdr(SUM.SEC_TEAM, "👥 TEAM");
  rowLbl(SUM.HEADCOUNT,   "Total Headcount",          true);
  rowLbl(SUM.PAYROLL_PCT, "Payroll as % of ARR");            bench(SUM.PAYROLL_PCT, "Benchmark: < 40% at Series A");
  rowLbl(SUM.ENG_HC,      "  ↳ Engineering Headcount");      bench(SUM.ENG_HC,      "Product + R&D");
  rowLbl(SUM.CS_HC,       "  ↳ CS / FDE-CSE Headcount");     bench(SUM.CS_HC,       "Implementation + ongoing support");

  secHdr(SUM.SEC_ASSUMP, "🎛️ KEY ASSUMPTIONS");
  rowLbl(SUM.TARGET_ARR,  "Target ARR ($)");       bench(SUM.TARGET_ARR,  "Drivers B12");
  rowLbl(SUM.HORIZON_ROW, "Forecast Horizon");     bench(SUM.HORIZON_ROW, "Drivers B14");
  rowLbl(SUM.LOGO_GROWTH_ROW, "Logo acquisition growth (MoM)"); bench(SUM.LOGO_GROWTH_ROW, "Drivers B39");
  rowLbl(SUM.MM_ACV,      "MM Beg. ACV ($)");      bench(SUM.MM_ACV,      "Drivers C18");
  rowLbl(SUM.ENT_ACV,     "ENT Beg. ACV ($)");     bench(SUM.ENT_ACV,     "Drivers C19");
  rowLbl(SUM.MM_CHURN,    "MM Annual Churn %");     bench(SUM.MM_CHURN,    "Drivers D18");
  rowLbl(SUM.ENT_CHURN,   "ENT Annual Churn %");    bench(SUM.ENT_CHURN,   "Drivers D19");
  rowLbl(SUM.EXIST_ARR,   "Existing Book ARR ($)"); bench(SUM.EXIST_ARR,   "Drivers B107");
  [SUM.TARGET_ARR,SUM.HORIZON_ROW,SUM.LOGO_GROWTH_ROW,SUM.MM_ACV,SUM.ENT_ACV,
   SUM.MM_CHURN,SUM.ENT_CHURN,SUM.EXIST_ARR].forEach(function(r){sh.setRowHeight(r,20);});

  var moRow       = "$" + SUM.HDR_MO;
  var mmCulCol    = colLetter(REVCOLS.MM_CUMUL);
  var entCulCol   = colLetter(REVCOLS.ENT_CUMUL);
  var exMmCulCol  = colLetter(REVCOLS.EX_MM_CUMUL);
  var exEntCulCol = colLetter(REVCOLS.EX_ENT_CUMUL);
  var dsOff       = REVROWS.DATA_START - 1;

  function pnlF(pnlRow, col) {
    var m = colLetter(col) + moRow;
    return "=IF("+m+"=\"\",\"\",IFERROR(INDEX('💸 P&L'!$B$"+pnlRow+":$BQ$"+pnlRow+",1,"+m+"),\"\"))";
  }
  function hcF(hcRow, col) {
    var m = colLetter(col) + moRow;
    return "=IF("+m+"=\"\",\"\",IFERROR(INDEX('👥 Headcount'!$B$"+hcRow+":$BQ$"+hcRow+",1,"+m+"),\"\"))";
  }
  function setVal(sumRow, col, formula, fmt, bg, bold) {
    var r = sh.getRange(sumRow, col);
    r.setFormula(formula).setNumberFormat(fmt);
    if (bg)   r.setBackground(bg);
    if (bold) r.setFontWeight("bold");
  }

  [2,3,4,5].forEach(function(col) {
    var m = colLetter(col) + moRow;

    setVal(SUM.ARR, col, pnlF(PNL.ARR,col), "$#,##0", "#D5F5E3", true);
    setVal(SUM.MRR, col, pnlF(PNL.MRR,col), "$#,##0", "#D5F5E3", true);
    setVal(SUM.YOY_GROWTH, col,
      "=IF("+m+"=\"\",\"\",IF("+m+"<=12,\"—\",IFERROR(INDEX('💸 P&L'!$B$"+PNL.YOY_GROWTH+":$BQ$"+PNL.YOY_GROWTH+",1,"+m+"),\"—\")))",
      "0%");
    setVal(SUM.NET_NEW_ARR, col, pnlF(PNL.NET_NEW_ARR,col), "$#,##0");
    var arrM = "INDEX('💸 P&L'!$B$"+PNL.ARR+":$BQ$"+PNL.ARR+",1,"+m+")";
    var arr1 = "INDEX('💸 P&L'!$B$"+PNL.ARR+":$BQ$"+PNL.ARR+",1,1)";
    setVal(SUM.MOM_GROWTH_IMPLIED, col,
      "=IF("+m+"=\"\",\"\",IF("+m+"<=1,\"—\",IFERROR(POWER("+arrM+"/"+arr1+",1/MAX(1,"+m+"-1))-1,\"—\")))",
      "0.0%");
    setVal(SUM.LOGOS, col, hcF(17,col), "0", null, true);
    setVal(SUM.NRR, col,
      "=IF("+m+"=\"\",\"\",IF("+m+"<=1,\"—\",IFERROR(INDEX('💸 P&L'!$B$"+PNL.NRR+":$BQ$"+PNL.NRR+",1,"+m+"),\"—\")))",
      "0%", null, true);
    setVal(SUM.GROSS_MARGIN, col, pnlF(PNL.GROSS_MARGIN,col), "0%", null, true);
    var mmCumul  = "INDEX('📈 Revenue'!$"+mmCulCol+":$"+mmCulCol+","+m+"+"+dsOff+")";
    var mmExCum  = "INDEX('📈 Revenue'!$"+exMmCulCol+":$"+exMmCulCol+","+m+"+"+dsOff+")";
    var mmLogos  = "INDEX('👥 Headcount'!$B$15:$BQ$15,1,"+m+")";
    setVal(SUM.ARR_PER_LOGO_MM, col,
      "=IF("+m+"=\"\",\"\",IFERROR(("+mmCumul+"+"+mmExCum+")/MAX(1,"+mmLogos+"),\"—\"))",
      "$#,##0");
    var entCumul = "INDEX('📈 Revenue'!$"+entCulCol+":$"+entCulCol+","+m+"+"+dsOff+")";
    var entExCum = "INDEX('📈 Revenue'!$"+exEntCulCol+":$"+exEntCulCol+","+m+"+"+dsOff+")";
    var entLogos = "INDEX('👥 Headcount'!$B$16:$BQ$16,1,"+m+")";
    setVal(SUM.ARR_PER_LOGO_ENT, col,
      "=IF("+m+"=\"\",\"\",IFERROR(("+entCumul+"+"+entExCum+")/MAX(1,"+entLogos+"),\"—\"))",
      "$#,##0");
  });

  sh.getRange(SUM.CHURN_MM,  2).setFormula("='🎛️ Drivers'!D"+DR.MM_ROW).setNumberFormat("0%");
  sh.getRange(SUM.CHURN_ENT, 2).setFormula("='🎛️ Drivers'!D"+DR.ENT_ROW).setNumberFormat("0%");

  function cfF(cfRow, col) {
    var m = colLetter(col) + moRow;
    return "=IF("+m+"=\"\",\"\",IFERROR(INDEX('🏦 Cash Flow'!$B$"+cfRow+":$BQ$"+cfRow+",1,"+m+"),\"\"))";
  }

  [2,3,4,5].forEach(function(col) {
    var m = colLetter(col) + moRow;
    setVal(SUM.END_CASH,   col, cfF(CF.END_CASH,  col), "$#,##0", null, true);
    setVal(SUM.BURN_RATE,  col, cfF(CF.BURN_RATE, col), "$#,##0", null, true);
    setVal(SUM.RUNWAY,     col, cfF(CF.RUNWAY,    col), "0.0",    null, true);
    setVal(SUM.CUMUL_BURN, col, cfF(CF.CUMUL_BURN,col), "$#,##0");
    setVal(SUM.CUMUL_CAP,  col, cfF(CF.CUMUL_CAP, col), "$#,##0", null, true);
    setVal(SUM.CASH_ARR,   col, cfF(CF.CASH_ARR,  col), "0%");
  });

  sh.getRange(SUM.PLANNED_RAISE, 2)
    .setFormula("=IFERROR(SUMPRODUCT((LEN(TRIM('🎛️ Drivers'!B121:B123))>0)*IFERROR(VALUE('🎛️ Drivers'!B121:B123),0)),0)")
    .setNumberFormat("$#,##0").setFontWeight("bold");

  [2,3,4,5].forEach(function(col) {
    var m = colLetter(col) + moRow;
    setVal(SUM.BURN_MULTIPLE, col,
      "=IF("+m+"=\"\",\"\",IF("+m+"<=1,\"—\",IFERROR(INDEX('💸 P&L'!$B$"+PNL.BURN_MULTIPLE+":$BQ$"+PNL.BURN_MULTIPLE+",1,"+m+"),\"—\")))",
      "0.00", null, true);
    setVal(SUM.RULE_OF_40, col,
      "=IF("+m+"=\"\",\"\",IF("+m+"<=12,\"—\",IFERROR(INDEX('💸 P&L'!$B$"+PNL.RULE_OF_40+":$BQ$"+PNL.RULE_OF_40+",1,"+m+"),\"—\")))",
      "0%", null, true);
    setVal(SUM.ARR_PER_EMP, col, pnlF(PNL.ARR_PER_EMP, col), "$#,##0", null, true);
    setVal(SUM.HEADCOUNT, col, hcF(12, col), "0", null, true);
    setVal(SUM.PAYROLL_PCT, col,
      "=IF("+m+"=\"\",\"\",IFERROR(INDEX('👥 Headcount'!$B$22:$BQ$22,1,"+m+"),\"\"))",
      "0%");
    setVal(SUM.ENG_HC, col, hcF(3, col), "0");
    setVal(SUM.CS_HC,  col, hcF(7, col), "0");
    setVal(SUM.LTV_CAC_MM, col,
      "=IF("+m+"=\"\",\"\",IFERROR(('🎛️ Drivers'!B"+DR.MM_ROW+"/'🎛️ Drivers'!D"+DR.MM_ROW+")/'🎛️ Drivers'!G"+DR.MM_ROW+",\"—\"))",
      "0.0\"x\"", null, true);
    setVal(SUM.LTV_CAC_ENT, col,
      "=IF("+m+"=\"\",\"\",IFERROR(('🎛️ Drivers'!B"+DR.ENT_ROW+"/'🎛️ Drivers'!D"+DR.ENT_ROW+")/'🎛️ Drivers'!G"+DR.ENT_ROW+",\"—\"))",
      "0.0\"x\"", null, true);
  });

  sh.getRange(SUM.TARGET_ARR,  2).setFormula("='🎛️ Drivers'!"+DR.TARGET_ARR).setNumberFormat("$#,##0");
  sh.getRange(SUM.HORIZON_ROW, 2).setFormula("='🎛️ Drivers'!"+DR.HORIZON).setNumberFormat("0 \"months\"");
  sh.getRange(SUM.LOGO_GROWTH_ROW, 2).setFormula("='🎛️ Drivers'!"+DR.LOGO_GROWTH).setNumberFormat("0%");
  sh.getRange(SUM.MM_ACV,      2).setFormula("='🎛️ Drivers'!B"+DR.MM_ROW).setNumberFormat("$#,##0");
  sh.getRange(SUM.ENT_ACV,     2).setFormula("='🎛️ Drivers'!B"+DR.ENT_ROW).setNumberFormat("$#,##0");
  sh.getRange(SUM.MM_CHURN,    2).setFormula("='🎛️ Drivers'!D"+DR.MM_ROW).setNumberFormat("0%");
  sh.getRange(SUM.ENT_CHURN,   2).setFormula("='🎛️ Drivers'!D"+DR.ENT_ROW).setNumberFormat("0%");
  sh.getRange(SUM.EXIST_ARR,   2).setFormula("='🎛️ Drivers'!"+DR.EXIST_TOTAL_ARR).setNumberFormat("$#,##0");

  sh.getRange(SUM.NOTE, 1, 1, COLS).merge()
    .setValue("✅ All values pull from Drivers → P&L → Cash Flow. Change any Drivers input to update this dashboard instantly.")
    .setFontStyle("italic").setFontColor("#1D9E75").setBackground("#FDFEFE").setWrap(true);
  sh.setRowHeight(SUM.NOTE, 28);
}
