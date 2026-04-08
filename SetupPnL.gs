// SetupPnL.gs — P&L + _formatPnL. Called from setupFinancialModel in SetupMain.gs.

function setupPnL(ss) {
  var sh = ss.getSheetByName("💸 P&L"); var MONTHS = 60;
  var HOR     = "'🎛️ Drivers'!" + DR.HORIZON;
  var REV_TAB = "'📈 Revenue'";
  var cbMrrL   = colLetter(REVCOLS.CB_MRR);
  var cbCulL   = colLetter(REVCOLS.CB_CUMUL);
  var mmNAL    = colLetter(REVCOLS.MM_NEW_ARR);
  var mmChL    = colLetter(REVCOLS.MM_CHURN);
  var mmExL    = colLetter(REVCOLS.MM_EXP);
  var entNAL   = colLetter(REVCOLS.ENT_NEW_ARR);
  var entChL   = colLetter(REVCOLS.ENT_CHURN);
  var entExL   = colLetter(REVCOLS.ENT_EXP);
  var entLogL  = colLetter(REVCOLS.ENT_LOGOS);
  var exMmChL  = colLetter(REVCOLS.EX_MM_CHURN);
  var exMmExL  = colLetter(REVCOLS.EX_MM_EXP);
  var exEntChL = colLetter(REVCOLS.EX_ENT_CHURN);
  var exEntExL = colLetter(REVCOLS.EX_ENT_EXP);

  if(sh.getMaxColumns()<MONTHS+1) sh.insertColumnsAfter(sh.getMaxColumns(),MONTHS+1-sh.getMaxColumns());
  sh.setColumnWidth(1,240);
  for(var m=1;m<=MONTHS;m++) sh.setColumnWidth(m+1,80);

  hdr(sh,PNL.HDR_TITLE,1,"💸 P&L — Income Statement (Monthly)","#1A5276"); sh.getRange(PNL.HDR_TITLE,1,1,MONTHS+1).merge();
  hdr(sh,PNL.HDR_COLS,1,"Line Item","#1F618D");
  for(var m=1;m<=MONTHS;m++) hdr(sh,PNL.HDR_COLS,m+1,"Mo "+m,"#1F618D");

  _secHdr(sh,PNL.SEC_REVENUE,"REVENUE",MONTHS);
  sh.getRange(PNL.MRR,1).setValue("MRR ($)").setFontWeight("bold");
  sh.getRange(PNL.ARR,1).setValue("ARR ($)").setFontWeight("bold");
  sh.getRange(PNL.YOY_GROWTH,1).setValue("YoY ARR Growth");
  _subHdr(sh,PNL.SUB_ARR_MOVE,"ARR Movement",MONTHS);
  sh.getRange(PNL.NEW_ARR,1).setValue("  ↳ New ARR ($)");
  sh.getRange(PNL.CHURN_ARR,1).setValue("  ↳ Churned ARR ($)");
  sh.getRange(PNL.EXP_ARR,1).setValue("  ↳ Expansion ARR ($)");
  sh.getRange(PNL.NET_NEW_ARR,1).setValue("  ↳ Net New ARR ($)").setFontWeight("bold");
  sh.getRange(PNL.NRR,1).setValue("Net Revenue Retention (NRR %)").setFontWeight("bold");
  _secHdr(sh,PNL.SEC_COGS,"COST OF GOODS SOLD (COGS)",MONTHS);
  sh.getRange(PNL.INFRA,1).setValue("  Infrastructure");
  sh.getRange(PNL.CS_PAYROLL,1).setValue("  CS / FDE-CSE Payroll");
  sh.getRange(PNL.TOTAL_COGS,1).setValue("Total COGS ($)").setFontWeight("bold");
  sh.getRange(PNL.GROSS_PROFIT,1).setValue("Gross Profit ($)").setFontWeight("bold");
  sh.getRange(PNL.GROSS_MARGIN,1).setValue("Gross Margin %").setFontWeight("bold");
  _secHdr(sh,PNL.SEC_OPEX,"OPERATING EXPENSES (OpEx)",MONTHS);
  _subHdr(sh,PNL.SUB_RD,"R&D",MONTHS);
  sh.getRange(PNL.ENG_PAYROLL,1).setValue("  Engineering Payroll");
  sh.getRange(PNL.RD_SUBTOTAL,1).setValue("R&D Total ($)").setFontWeight("bold");
  _subHdr(sh,PNL.SUB_SM,"Sales & Marketing",MONTHS);
  sh.getRange(PNL.SALES_PAYROLL,1).setValue("  Sales Payroll");
  sh.getRange(PNL.COMMISSION,1).setValue("  Sales Commission");
  sh.getRange(PNL.MARKETING,1).setValue("  Marketing (Events + Digital)");
  sh.getRange(PNL.TRAVEL,1).setValue("  Travel (ENT deals + Events)");
  sh.getRange(PNL.SM_SUBTOTAL,1).setValue("S&M Total ($)").setFontWeight("bold");
  _subHdr(sh,PNL.SUB_GA,"General & Administrative (G&A)",MONTHS);
  sh.getRange(PNL.GA_PAYROLL,1).setValue("  G&A Payroll");
  sh.getRange(PNL.TOOLING,1).setValue("  Tooling / Misc (Engineering)");
  sh.getRange(PNL.PROF_FEES,1).setValue("  Professional Fees");
  sh.getRange(PNL.CO_SW,1).setValue("  Company Software");
  sh.getRange(PNL.RECRUITING,1).setValue("  Recruiting Costs");
  sh.getRange(PNL.HARDWARE,1).setValue("  Hardware (new hires)");
  sh.getRange(PNL.GA_SUBTOTAL,1).setValue("G&A Total ($)").setFontWeight("bold");
  sh.getRange(PNL.TOTAL_OPEX,1).setValue("Total OpEx ($)").setFontWeight("bold");
  _secHdr(sh,PNL.SEC_EBITDA,"EBITDA",MONTHS);
  sh.getRange(PNL.EBITDA,1).setValue("EBITDA ($)").setFontWeight("bold");
  sh.getRange(PNL.EBITDA_MARGIN,1).setValue("EBITDA Margin %").setFontWeight("bold");
  sh.getRange(PNL.CUMUL_BURN,1).setValue("Cumulative Burn ($)").setFontWeight("bold");
  _secHdr(sh,PNL.SEC_METRICS,"KEY METRICS",MONTHS);
  sh.getRange(PNL.HEADCOUNT,1).setValue("Total Headcount");
  sh.getRange(PNL.ARR_PER_EMP,1).setValue("ARR per Employee ($)").setFontWeight("bold");
  sh.getRange(PNL.MAGIC_NUMBER,1).setValue("Magic Number").setFontWeight("bold");
  sh.getRange(PNL.BURN_MULTIPLE,1).setValue("Burn Multiple (x)").setFontWeight("bold");
  sh.getRange(PNL.RULE_OF_40,1).setValue("Rule of 40 (%)").setFontWeight("bold");
  sh.getRange(PNL.RD_PCT_ARR,1).setValue("  R&D as % of ARR");
  sh.getRange(PNL.SM_PCT_ARR,1).setValue("  S&M as % of ARR");
  sh.getRange(PNL.GA_PCT_ARR,1).setValue("  G&A as % of ARR");
  sh.getRange(PNL.NOTE,1,1,6).merge().setValue("✅ All formulas live. Only edit 🎛️ Drivers — everything here is auto-calculated.").setFontStyle("italic").setFontColor("#1D9E75");

  for(var m=1;m<=MONTHS;m++){
    var col=m+1; var C=colLetter(col); var Cp=colLetter(col-1); var ms=String(m);
    var dRow     = REV.DATA_START + m - 1;
    var dRowPrev = REV.DATA_START + m - 2;

    sh.getRange(PNL.MRR,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+REV_TAB+"!"+cbMrrL+dRow+")")
      .setNumberFormat("$#,##0").setBackground("#D5F5E3").setFontWeight("bold");
    sh.getRange(PNL.ARR,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+C+PNL.MRR+"*12)")
      .setNumberFormat("$#,##0").setFontWeight("bold");
    if(m<=12) sh.getRange(PNL.YOY_GROWTH,col).setValue("").setNumberFormat("0%");
    else sh.getRange(PNL.YOY_GROWTH,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR("+C+PNL.ARR+"/"+colLetter(col-12)+PNL.ARR+"-1,\"\"))")
      .setNumberFormat("0%");

    sh.getRange(PNL.NEW_ARR,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+REV_TAB+"!"+mmNAL+dRow+"+"+REV_TAB+"!"+entNAL+dRow+")")
      .setNumberFormat("$#,##0");
    sh.getRange(PNL.CHURN_ARR,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+REV_TAB+"!"+mmChL+dRow+"+"+REV_TAB+"!"+entChL+dRow
        +"+IFERROR("+REV_TAB+"!"+exMmChL+dRow+",0)+IFERROR("+REV_TAB+"!"+exEntChL+dRow+",0))")
      .setNumberFormat("$#,##0").setFontColor("#922B21");
    sh.getRange(PNL.EXP_ARR,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+REV_TAB+"!"+mmExL+dRow+"+"+REV_TAB+"!"+entExL+dRow
        +"+IFERROR("+REV_TAB+"!"+exMmExL+dRow+",0)+IFERROR("+REV_TAB+"!"+exEntExL+dRow+",0))")
      .setNumberFormat("$#,##0").setFontColor("#1D9E75");
    sh.getRange(PNL.NET_NEW_ARR,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+C+PNL.NEW_ARR+"+"+C+PNL.CHURN_ARR+"+"+C+PNL.EXP_ARR+")")
      .setNumberFormat("$#,##0").setFontWeight("bold");

    if(m===1) sh.getRange(PNL.NRR,col).setValue("").setNumberFormat("0%");
    else sh.getRange(PNL.NRR,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR(POWER(("
        +REV_TAB+"!"+cbCulL+dRowPrev
        +"+"+REV_TAB+"!"+mmChL+dRow+"+"+REV_TAB+"!"+entChL+dRow
        +"+IFERROR("+REV_TAB+"!"+exMmChL+dRow+",0)+IFERROR("+REV_TAB+"!"+exEntChL+dRow+",0)"
        +"+"+REV_TAB+"!"+mmExL+dRow+"+"+REV_TAB+"!"+entExL+dRow
        +"+IFERROR("+REV_TAB+"!"+exMmExL+dRow+",0)+IFERROR("+REV_TAB+"!"+exEntExL+dRow+",0))"
        +"/"+REV_TAB+"!"+cbCulL+dRowPrev+",12),\"\"))")
      .setNumberFormat("0%").setFontWeight("bold").setBackground("#E8F8F5");

    sh.getRange(PNL.INFRA,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR('👥 Headcount'!"+C+"17,0)*'🎛️ Drivers'!"+DR.INFRA+")")
      .setNumberFormat("$#,##0");
    sh.getRange(PNL.CS_PAYROLL,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+"'👥 Headcount'!"+C+"8)").setNumberFormat("$#,##0");
    sh.getRange(PNL.TOTAL_COGS,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+C+PNL.INFRA+"+"+C+PNL.CS_PAYROLL+")")
      .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#F2F3F4");
    sh.getRange(PNL.GROSS_PROFIT,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+C+PNL.MRR+"-"+C+PNL.TOTAL_COGS+")")
      .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");
    sh.getRange(PNL.GROSS_MARGIN,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR("+C+PNL.GROSS_PROFIT+"/"+C+PNL.MRR+",0))")
      .setNumberFormat("0%").setFontWeight("bold").setBackground("#D5F5E3");

    sh.getRange(PNL.ENG_PAYROLL,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+"'👥 Headcount'!"+C+"4)").setNumberFormat("$#,##0");
    sh.getRange(PNL.RD_SUBTOTAL,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+C+PNL.ENG_PAYROLL+")")
      .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#F2F3F4");
    sh.getRange(PNL.SALES_PAYROLL,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+"'👥 Headcount'!"+C+"6)").setNumberFormat("$#,##0");
    var prevMRR=m===1?"0":colLetter(col-1)+PNL.MRR;
    sh.getRange(PNL.COMMISSION,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",MAX(0,"+C+PNL.MRR+"-"+prevMRR+")*12*'🎛️ Drivers'!"+DR.COMMISSION+")")
      .setNumberFormat("$#,##0");
    sh.getRange(PNL.MARKETING,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IF("+ms+"<=12,('🎛️ Drivers'!B66*(1-'🎛️ Drivers'!B117)+'🎛️ Drivers'!B67)/12,('🎛️ Drivers'!B66*(1-'🎛️ Drivers'!B117)+'🎛️ Drivers'!B67)*'🎛️ Drivers'!B69/12))")
      .setNumberFormat("$#,##0");
    sh.getRange(PNL.TRAVEL,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR("+REV_TAB+"!"+entLogL+dRow+",0)*'🎛️ Drivers'!B113"
        +"+IF("+ms+"<=12,'🎛️ Drivers'!B66*'🎛️ Drivers'!B117/12,'🎛️ Drivers'!B66*'🎛️ Drivers'!B117*'🎛️ Drivers'!B69/12))")
      .setNumberFormat("$#,##0");
    sh.getRange(PNL.SM_SUBTOTAL,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+C+PNL.SALES_PAYROLL+"+"+C+PNL.COMMISSION+"+"+C+PNL.MARKETING+"+"+C+PNL.TRAVEL+")")
      .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#F2F3F4");

    sh.getRange(PNL.GA_PAYROLL,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+"'👥 Headcount'!"+C+"10)").setNumberFormat("$#,##0");
    sh.getRange(PNL.TOOLING,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+"'👥 Headcount'!"+C+"3*'🎛️ Drivers'!"+DR.TOOLING+")").setNumberFormat("$#,##0");
    sh.getRange(PNL.PROF_FEES,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+"'🎛️ Drivers'!"+DR.PROF_FEES+"/12)").setNumberFormat("$#,##0");
    sh.getRange(PNL.CO_SW,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+"'👥 Headcount'!"+C+"12*'🎛️ Drivers'!"+DR.CO_SOFTWARE+")").setNumberFormat("$#,##0");
    if(m===1){
      sh.getRange(PNL.RECRUITING,col).setFormula("=IF(1>"+HOR+",\"\",0)").setNumberFormat("$#,##0");
      sh.getRange(PNL.HARDWARE,col).setFormula("=IF(1>"+HOR+",\"\",0)").setNumberFormat("$#,##0");
    } else {
      sh.getRange(PNL.RECRUITING,col)
        .setFormula("=IF("+ms+">"+HOR+",\"\","+"MAX(0,'👥 Headcount'!"+C+"12-'👥 Headcount'!"+Cp+"12)*IFERROR('👥 Headcount'!"+C+"11/'👥 Headcount'!"+C+"12/'🎛️ Drivers'!"+DR.LOADED_MULT+"*12,0)*'🎛️ Drivers'!"+DR.RECRUIT_PCT+")")
        .setNumberFormat("$#,##0");
      sh.getRange(PNL.HARDWARE,col)
        .setFormula("=IF("+ms+">"+HOR+",\"\","+"MAX(0,'👥 Headcount'!"+C+"12-'👥 Headcount'!"+Cp+"12)*'🎛️ Drivers'!"+DR.HW_NEW_HIRE+")")
        .setNumberFormat("$#,##0");
    }
    sh.getRange(PNL.GA_SUBTOTAL,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+C+PNL.GA_PAYROLL+"+"+C+PNL.TOOLING+"+"+C+PNL.PROF_FEES+"+"+C+PNL.CO_SW+"+"+C+PNL.RECRUITING+"+"+C+PNL.HARDWARE+")")
      .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#F2F3F4");
    sh.getRange(PNL.TOTAL_OPEX,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+C+PNL.RD_SUBTOTAL+"+"+C+PNL.SM_SUBTOTAL+"+"+C+PNL.GA_SUBTOTAL+")")
      .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#F2F3F4");

    sh.getRange(PNL.EBITDA,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+C+PNL.GROSS_PROFIT+"-"+C+PNL.TOTAL_OPEX+")")
      .setNumberFormat("$#,##0").setFontWeight("bold");
    sh.getRange(PNL.EBITDA_MARGIN,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR("+C+PNL.EBITDA+"/"+C+PNL.MRR+",0))")
      .setNumberFormat("0%").setFontWeight("bold");
    sh.getRange(PNL.INTEREST_INCOME,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR('💰 Funding'!D"+(FUND.MONTHLY_START+m-1)+",0))")
      .setNumberFormat("$#,##0").setFontColor("#1D9E75").setBackground("#E8F8F5");
    var prevBurn=m===1?"0":colLetter(col-1)+PNL.CUMUL_BURN;
    sh.getRange(PNL.CUMUL_BURN,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+prevBurn+"+"+C+PNL.EBITDA+"+"+C+PNL.INTEREST_INCOME+")")
      .setNumberFormat("$#,##0").setFontWeight("bold");

    sh.getRange(PNL.HEADCOUNT,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\","+"'👥 Headcount'!"+C+"12)").setNumberFormat("0");
    sh.getRange(PNL.ARR_PER_EMP,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR("+C+PNL.ARR+"/"+C+PNL.HEADCOUNT+",0))")
      .setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#E8F8F5");
    if(m===1){
      sh.getRange(PNL.MAGIC_NUMBER,col).setValue("").setNumberFormat("0.00");
      sh.getRange(PNL.BURN_MULTIPLE,col).setValue("").setNumberFormat("0.00");
    } else {
      sh.getRange(PNL.MAGIC_NUMBER,col)
        .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR(MAX(0,"+C+PNL.MRR+"-"+Cp+PNL.MRR+")/"+Cp+PNL.SM_SUBTOTAL+",\"\"))")
        .setNumberFormat("0.00").setFontWeight("bold").setBackground("#E8F8F5");
      sh.getRange(PNL.BURN_MULTIPLE,col)
        .setFormula("=IF("+ms+">"+HOR+",\"\",IF("+C+PNL.EBITDA+">=0,\"\",IFERROR(ABS("+C+PNL.EBITDA+")/MAX(1,"+C+PNL.NET_NEW_ARR+"),\"\")))")
        .setNumberFormat("0.00").setFontWeight("bold").setBackground("#E8F8F5");
    }
    if(m<=12) sh.getRange(PNL.RULE_OF_40,col).setValue("").setNumberFormat("0%");
    else sh.getRange(PNL.RULE_OF_40,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR("+C+PNL.YOY_GROWTH+"+"+C+PNL.EBITDA_MARGIN+",\"\"))")
      .setNumberFormat("0%").setFontWeight("bold").setBackground("#E8F8F5");
    sh.getRange(PNL.RD_PCT_ARR,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR("+C+PNL.RD_SUBTOTAL+"/"+C+PNL.ARR+",0))").setNumberFormat("0%");
    sh.getRange(PNL.SM_PCT_ARR,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR("+C+PNL.SM_SUBTOTAL+"/"+C+PNL.ARR+",0))").setNumberFormat("0%");
    sh.getRange(PNL.GA_PCT_ARR,col)
      .setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR("+C+PNL.GA_SUBTOTAL+"/"+C+PNL.ARR+",0))").setNumberFormat("0%");
  }

  var rules=sh.getConditionalFormatRules();
  var ebitdaRange=sh.getRange(PNL.EBITDA,2,1,MONTHS);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor("#922B21").setBackground("#FADBD8").setRanges([ebitdaRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setFontColor("#1D9E75").setBackground("#D5F5E3").setRanges([ebitdaRange]).build());
  var burnRange=sh.getRange(PNL.CUMUL_BURN,2,1,MONTHS);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor("#922B21").setBackground("#FADBD8").setRanges([burnRange]).build());
  var magicRange=sh.getRange(PNL.MAGIC_NUMBER,2,1,MONTHS);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.75).setBackground("#D5F5E3").setFontColor("#1D9E75").setRanges([magicRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.4,0.75).setBackground("#FEF9E7").setFontColor("#D4AC0D").setRanges([magicRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.4).setBackground("#FADBD8").setFontColor("#922B21").setRanges([magicRange]).build());
  var bmRange=sh.getRange(PNL.BURN_MULTIPLE,2,1,MONTHS);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(1).setBackground("#D5F5E3").setFontColor("#1D9E75").setRanges([bmRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(1,2).setBackground("#FEF9E7").setFontColor("#D4AC0D").setRanges([bmRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(2).setBackground("#FADBD8").setFontColor("#922B21").setRanges([bmRange]).build());
  var r40Range=sh.getRange(PNL.RULE_OF_40,2,1,MONTHS);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.40).setBackground("#D5F5E3").setFontColor("#1D9E75").setRanges([r40Range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.20,0.40).setBackground("#FEF9E7").setFontColor("#D4AC0D").setRanges([r40Range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.20).setBackground("#FADBD8").setFontColor("#922B21").setRanges([r40Range]).build());
  sh.setConditionalFormatRules(rules);
  _formatPnL(sh,MONTHS);
}

function _formatPnL(sh,MONTHS) {
  var HOR="'🎛️ Drivers'!"+DR.HORIZON;
  for(var m=1;m<=MONTHS;m++){
    sh.getRange(PNL.HDR_COLS,m+1).setFormula("=IF("+m+">"+HOR+",\"\",TEXT(EDATE('🎛️ Drivers'!"+DR.FORECAST_START+","+(m-1)+"),\"MMM YY\"))").setBackground("#1F618D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  }
  for(var r=1;r<=PNL.NOTE;r++) sh.setRowHeight(r,20);
  [PNL.SEC_REVENUE,PNL.SEC_COGS,PNL.SEC_OPEX,PNL.SEC_EBITDA,PNL.SEC_METRICS].forEach(function(r){sh.setRowHeight(r,26);});
  [PNL.SUB_ARR_MOVE,PNL.SUB_RD,PNL.SUB_SM,PNL.SUB_GA].forEach(function(r){sh.setRowHeight(r,21);});
  [PNL.MRR,PNL.ARR,PNL.NRR,PNL.GROSS_PROFIT,PNL.TOTAL_OPEX,PNL.EBITDA,PNL.MAGIC_NUMBER,PNL.BURN_MULTIPLE,PNL.RULE_OF_40].forEach(function(r){sh.setRowHeight(r,22);});
  [PNL.YOY_GROWTH,PNL.NEW_ARR,PNL.CHURN_ARR,PNL.EXP_ARR,PNL.INFRA,PNL.CS_PAYROLL,PNL.ENG_PAYROLL,
   PNL.SALES_PAYROLL,PNL.COMMISSION,PNL.MARKETING,PNL.TRAVEL,
   PNL.GA_PAYROLL,PNL.TOOLING,PNL.PROF_FEES,PNL.CO_SW,PNL.RECRUITING,PNL.HARDWARE,
   PNL.INTEREST_INCOME,PNL.HEADCOUNT,PNL.RD_PCT_ARR,PNL.SM_PCT_ARR,PNL.GA_PCT_ARR]
    .forEach(function(r){sh.getRange(r,1).setFontWeight("normal").setFontColor("#444444");});
  [PNL.MRR,PNL.ARR,PNL.NET_NEW_ARR,PNL.NRR,PNL.TOTAL_COGS,PNL.GROSS_PROFIT,PNL.GROSS_MARGIN,
   PNL.RD_SUBTOTAL,PNL.SM_SUBTOTAL,PNL.GA_SUBTOTAL,PNL.TOTAL_OPEX,
   PNL.EBITDA,PNL.EBITDA_MARGIN,PNL.CUMUL_BURN,PNL.ARR_PER_EMP,PNL.MAGIC_NUMBER,PNL.BURN_MULTIPLE,PNL.RULE_OF_40]
    .forEach(function(r){sh.getRange(r,1).setFontWeight("bold").setFontColor("#000000");});
  [PNL.TOTAL_COGS,PNL.GROSS_PROFIT,PNL.TOTAL_OPEX,PNL.EBITDA,PNL.SEC_METRICS].forEach(function(r){
    sh.getRange(r,1,1,MONTHS+1).setBorder(true,false,false,false,false,false,"#AAAAAA",SpreadsheetApp.BorderStyle.SOLID);
  });
  [PNL.GROSS_MARGIN,PNL.EBITDA_MARGIN].forEach(function(r){
    sh.getRange(r,1,1,MONTHS+1).setBorder(false,false,true,false,false,false,"#AAAAAA",SpreadsheetApp.BorderStyle.SOLID);
  });
  var stripeRows=[PNL.MRR,PNL.ARR,PNL.YOY_GROWTH,PNL.NEW_ARR,PNL.CHURN_ARR,PNL.EXP_ARR,PNL.NET_NEW_ARR,PNL.NRR,
    PNL.INFRA,PNL.CS_PAYROLL,PNL.TOTAL_COGS,PNL.GROSS_PROFIT,PNL.GROSS_MARGIN,
    PNL.ENG_PAYROLL,PNL.RD_SUBTOTAL,PNL.SALES_PAYROLL,PNL.COMMISSION,PNL.MARKETING,PNL.TRAVEL,PNL.SM_SUBTOTAL,
    PNL.GA_PAYROLL,PNL.TOOLING,PNL.PROF_FEES,PNL.CO_SW,PNL.RECRUITING,PNL.HARDWARE,PNL.GA_SUBTOTAL,PNL.TOTAL_OPEX,
    PNL.EBITDA,PNL.EBITDA_MARGIN,PNL.CUMUL_BURN,PNL.HEADCOUNT,PNL.ARR_PER_EMP,PNL.MAGIC_NUMBER,PNL.BURN_MULTIPLE,PNL.RULE_OF_40,PNL.RD_PCT_ARR,PNL.SM_PCT_ARR,PNL.GA_PCT_ARR];
  stripeRows.forEach(function(r,i){sh.getRange(r,1).setBackground(i%2===0?"#FFFFFF":"#F8F9FA");});
  sh.getRange(PNL.INTEREST_INCOME,1).setNote("Non-operating income. Based on prior month ending cash × annual rate from 🎛️ Drivers K.");
  sh.getRange(PNL.INTEREST_INCOME,1).setValue("Interest Income ($)").setFontColor("#1D9E75").setFontStyle("italic").setFontWeight("normal");
  sh.getRange(PNL.MAGIC_NUMBER,1).setNote("> 0.75 efficient  |  > 1.5 pour fuel  |  < 0.4 review S&M");
  sh.getRange(PNL.NRR,1).setNote("Annualized monthly retention. Best-in-class SaaS: > 120%");
  sh.getRange(PNL.ARR_PER_EMP,1).setNote("Bessemer Series A benchmark: $150K–$200K for vertical SaaS");
  sh.getRange(PNL.BURN_MULTIPLE,1).setNote("< 1x = efficient  |  1–2x = watch  |  > 2x = flag. Blank when profitable.");
  sh.getRange(PNL.RULE_OF_40,1).setNote("YoY ARR Growth % + EBITDA Margin %. > 40% = healthy. Blank for months 1–12.");
  sh.getRange(PNL.RD_PCT_ARR,1).setNote("Benchmark: R&D 20–30% of ARR at Series A");
  sh.getRange(PNL.SM_PCT_ARR,1).setNote("Benchmark: S&M 30–50% of ARR at Series A");
  sh.getRange(PNL.GA_PCT_ARR,1).setNote("Benchmark: G&A 10–20% of ARR at Series A");
}
