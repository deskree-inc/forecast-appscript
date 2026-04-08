// SetupFunding.gs — Called from setupFinancialModel in SetupMain.gs.

function setupFunding(ss) {
  var sh  = ss.getSheetByName("💰 Funding");
  var MAX = 60;
  var HOR = "'🎛️ Drivers'!"+DR.HORIZON;
  var FS  = "'🎛️ Drivers'!"+DR.FORECAST_START;

  sh.setColumnWidth(1,210); sh.setColumnWidth(2,160); sh.setColumnWidth(3,130);
  sh.setColumnWidth(4,160); sh.setColumnWidth(5,280);

  hdr(sh,1,1,"💰 FUNDING — Rounds & Monthly Cash Schedule","#1A5276"); sh.getRange(1,1,1,5).merge();
  sh.getRange(2,1,1,5).merge().setValue("All inputs in 🎛️ Drivers Section K. This tab is fully auto-calculated — do not edit.").setFontStyle("italic").setFontColor("#888");

  sh.getRange(3,1,1,5).merge().setValue("📋 FUNDING ROUNDS").setBackground("#2C3E50").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(10);
  ["Round Name","Amount Raised ($)","Close Date","ARR at Close ($)","Notes"].forEach(function(h,i){hdr(sh,4,i+1,h,"#1F618D");});
  [[DR.ROUND1_NAME,DR.ROUND1_AMT,DR.ROUND1_DATE,DR.ROUND1_ARR],
   [DR.ROUND2_NAME,DR.ROUND2_AMT,DR.ROUND2_DATE,DR.ROUND2_ARR],
   [DR.ROUND3_NAME,DR.ROUND3_AMT,DR.ROUND3_DATE,DR.ROUND3_ARR]]
    .forEach(function(round,i){
      var r=5+i;
      sh.getRange(r,1).setFormula("='🎛️ Drivers'!"+round[0]).setBackground("#F2F3F4");
      sh.getRange(r,2).setFormula("=IFERROR(IF('🎛️ Drivers'!"+round[1]+"=\"\",\"\",VALUE('🎛️ Drivers'!"+round[1]+")),\"\")").setNumberFormat("$#,##0").setBackground("#F2F3F4");
      sh.getRange(r,3).setFormula("=IFERROR(IF(ISNUMBER('🎛️ Drivers'!"+round[2]+"),'🎛️ Drivers'!"+round[2]+",\"\"),\"\")").setNumberFormat("MMM YYYY").setBackground("#F2F3F4");
      sh.getRange(r,4).setFormula("=IFERROR(IF('🎛️ Drivers'!"+round[3]+"=\"\",\"\",VALUE('🎛️ Drivers'!"+round[3]+")),\"\")").setNumberFormat("$#,##0").setBackground("#F2F3F4");
      sh.setRowHeight(r,24);
    });
  label(sh,8,1,"Total Capital Raised");
  sh.getRange(8,2).setFormula("=SUMPRODUCT((B5:B7<>\"\")*IFERROR(VALUE(B5:B7),0))").setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");
  label(sh,8,3,"Interest Rate (annual)");
  sh.getRange(8,4).setFormula("='🎛️ Drivers'!"+DR.INTEREST_RATE).setNumberFormat("0.0%").setBackground("#F2F3F4").setFontWeight("bold");
  label(sh,8,5,"Opening Cash");
  sh.getRange(8,5).setFormula("='🎛️ Drivers'!"+DR.OPENING_CASH).setNumberFormat("$#,##0").setBackground("#F2F3F4").setFontWeight("bold");

  sh.getRange(9,1,1,5).merge()
    .setFormula("=IF(OR("
      +"AND(ISNUMBER('🎛️ Drivers'!"+DR.ROUND1_DATE+"),MAX(0,(YEAR('🎛️ Drivers'!"+DR.ROUND1_DATE+")-YEAR('🎛️ Drivers'!"+DR.FORECAST_START+"))*12+MONTH('🎛️ Drivers'!"+DR.ROUND1_DATE+")-MONTH('🎛️ Drivers'!"+DR.FORECAST_START+"))>'🎛️ Drivers'!"+DR.HORIZON+"),"
      +"AND(ISNUMBER('🎛️ Drivers'!"+DR.ROUND2_DATE+"),MAX(0,(YEAR('🎛️ Drivers'!"+DR.ROUND2_DATE+")-YEAR('🎛️ Drivers'!"+DR.FORECAST_START+"))*12+MONTH('🎛️ Drivers'!"+DR.ROUND2_DATE+")-MONTH('🎛️ Drivers'!"+DR.FORECAST_START+"))>'🎛️ Drivers'!"+DR.HORIZON+"),"
      +"AND(ISNUMBER('🎛️ Drivers'!"+DR.ROUND3_DATE+"),MAX(0,(YEAR('🎛️ Drivers'!"+DR.ROUND3_DATE+")-YEAR('🎛️ Drivers'!"+DR.FORECAST_START+"))*12+MONTH('🎛️ Drivers'!"+DR.ROUND3_DATE+")-MONTH('🎛️ Drivers'!"+DR.FORECAST_START+"))>'🎛️ Drivers'!"+DR.HORIZON+")),"
      +"\"⚠️ Warning: one or more round close dates fall outside the forecast horizon.\","
      +"\"All round dates are within the forecast horizon.\")")
    .setWrap(true).setFontStyle("italic");
  sh.setRowHeight(9,44);

  var valRange = sh.getRange(9,1,1,5);
  var valRules = sh.getConditionalFormatRules();
  valRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=LEFT(A9,2)=\"⚠\"").setBackground("#FADBD8").setFontColor("#922B21").setRanges([valRange]).build());
  valRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=LEFT(A9,2)<>\"⚠\"").setBackground("#D5F5E3").setFontColor("#1D9E75").setRanges([valRange]).build());
  sh.setConditionalFormatRules(valRules);

  sh.getRange(10,1,1,5).merge()
    .setValue("📊 MONTHLY CAPITAL & INTEREST SCHEDULE  |  Col C feeds 🏦 Cash Flow  |  Col D feeds 💸 P&L Interest Income")
    .setBackground("#2C3E50").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(10);
  ["Mo #","Month","Capital Injected ($)","Interest Income ($)","Round"].forEach(function(h,i){hdr(sh,11,i+1,h,"#1F618D");});
  sh.getRange(12,1,1,5).merge().setValue("Note: each round must have ALL fields filled or ALL left blank.").setFontStyle("italic").setFontColor("#888").setWrap(true); sh.setRowHeight(12,36);
  sh.getRange(13,1,1,5).merge().setValue("Month 1 interest = Opening Cash × rate / 12. Subsequent months = prior ending cash × rate / 12.").setFontStyle("italic").setFontColor("#888").setWrap(true); sh.setRowHeight(13,30);

  for(var m=1;m<=MAX;m++){
    var r   = FUND.MONTHLY_START + m - 1;
    var ms  = String(m);
    var md  = "EDATE("+FS+","+(m-1)+")";
    function capCheck(dRef,aRef){
      return "IFERROR(IF(ISNUMBER('🎛️ Drivers'!"+dRef+"),IF(AND(YEAR("+md+")=YEAR('🎛️ Drivers'!"+dRef+"),MONTH("+md+")=MONTH('🎛️ Drivers'!"+dRef+")),IFERROR(VALUE('🎛️ Drivers'!"+aRef+"),0),0),0),0)";
    }
    sh.getRange(r,1).setFormula("=IF("+ms+">"+HOR+",\"\","+ms+")").setHorizontalAlignment("center");
    sh.getRange(r,2).setFormula("=IF("+ms+">"+HOR+",\"\",TEXT("+md+",\"MMM YY\"))").setNumberFormat("@");
    sh.getRange(r,3).setFormula("=IF("+ms+">"+HOR+",\"\","+capCheck(DR.ROUND1_DATE,DR.ROUND1_AMT)+"+"+capCheck(DR.ROUND2_DATE,DR.ROUND2_AMT)+"+"+capCheck(DR.ROUND3_DATE,DR.ROUND3_AMT)+")").setNumberFormat("$#,##0");
    if(m===1){
      sh.getRange(r,4).setFormula("=IF(1>"+HOR+",\"\",IFERROR('🎛️ Drivers'!"+DR.OPENING_CASH+"*'🎛️ Drivers'!"+DR.INTEREST_RATE+"/12,0))").setNumberFormat("$#,##0");
    } else {
      sh.getRange(r,4).setFormula("=IF("+ms+">"+HOR+",\"\",MAX(0,'🏦 Cash Flow'!"+colLetter(m)+CF.END_CASH+")*'🎛️ Drivers'!"+DR.INTEREST_RATE+"/12)").setNumberFormat("$#,##0");
    }
    sh.getRange(r,5).setFormula("=IF("+ms+">"+HOR+",\"\",IF(C"+r+">0,IF("+capCheck(DR.ROUND1_DATE,DR.ROUND1_AMT)+">0,'🎛️ Drivers'!"+DR.ROUND1_NAME+",IF("+capCheck(DR.ROUND2_DATE,DR.ROUND2_AMT)+">0,'🎛️ Drivers'!"+DR.ROUND2_NAME+",'🎛️ Drivers'!"+DR.ROUND3_NAME+")),\"\"))").setFontColor("#1D9E75").setFontStyle("italic");
  }

  var capRange = sh.getRange(FUND.MONTHLY_START,3,MAX,1);
  var intRange = sh.getRange(FUND.MONTHLY_START,4,MAX,1);
  var cfRules  = sh.getConditionalFormatRules();
  cfRules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground("#D5F5E3").setFontColor("#1D9E75").setRanges([capRange]).build());
  cfRules.push(SpreadsheetApp.newConditionalFormatRule().whenCellNotEmpty().setBackground("#E8F8F5").setRanges([intRange]).build());
  sh.setConditionalFormatRules(cfRules);
  sh.getRange(FUND.MONTHLY_START+MAX,1,1,5).merge().setValue("Interest income uses prior month ending cash from 🏦 Cash Flow — no circular dependency.").setFontStyle("italic").setFontColor("#888").setWrap(true); sh.setRowHeight(FUND.MONTHLY_START+MAX,36);
}
