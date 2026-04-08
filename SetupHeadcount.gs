// SetupHeadcount.gs — Called from setupFinancialModel in SetupMain.gs.

function setupHeadcount(ss) {
  var sh      = ss.getSheetByName("👥 Headcount");
  var MAX     = 60;
  var REV_TAB = "'📈 Revenue'";
  var DS      = REV.DATA_START;
  var MO_L    = colLetter(REVCOLS.MO_NUM);
  var MM_L    = colLetter(REVCOLS.MM_LOGOS);
  var ENT_L   = colLetter(REVCOLS.ENT_LOGOS);

  if(sh.getMaxColumns()<MAX+1) sh.insertColumnsAfter(sh.getMaxColumns(),MAX+1-sh.getMaxColumns());
  sh.setColumnWidth(1,220);
  for(var m=1;m<=MAX;m++) sh.setColumnWidth(m+1,55);
  hdr(sh,1,1,"👥 HEADCOUNT — Monthly Plan","#1A5276"); sh.getRange(1,1,1,MAX+1).merge();
  hdr(sh,2,1,"Department / Role","#1F618D");
  for(var m=1;m<=MAX;m++) sh.getRange(2,m+1)
    .setFormula("=IF("+m+">'🎛️ Drivers'!"+DR.HORIZON+",\"\",TEXT(EDATE('🎛️ Drivers'!"+DR.FORECAST_START+","+(m-1)+"),\"MMM YY\"))")
    .setBackground("#1F618D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setNumberFormat("@");
  var HOR = "'🎛️ Drivers'!"+DR.HORIZON;
  var LOD = "'🎛️ Drivers'!"+DR.LOADED_MULT;

  label(sh,3,1,"Engineering (product + R&D)  — HC");
  label(sh,4,1,"Engineering  — Cost / mo");
  for(var m=1;m<=MAX;m++){
    var col=m+1; var C=colLetter(col); var Cp=colLetter(col-1); var ms=String(m);
    var newMM="MAX(0,"+C+"15-'🎛️ Drivers'!"+DR.EXIST_MM_LOGOS+")";
    var newEN="MAX(0,"+C+"16-'🎛️ Drivers'!"+DR.EXIST_ENT_LOGOS+")";
    var mmR="'🎛️ Drivers'!"+DR.ENG_MM_RATIO; var enR="'🎛️ Drivers'!"+DR.ENG_ENT_RATIO;
    var rndR="'🎛️ Drivers'!"+DR.RND_RATIO; var seed="'🎛️ Drivers'!B"+DR.ENG;
    var prod="MAX("+seed+",IFERROR(CEILING("+newMM+"/"+mmR+",1),0)+IFERROR(CEILING("+newEN+"/"+enR+",1),0))";
    var rnd="IFERROR(FLOOR(("+prod+")/"+rndR+",1),0)"; var req="("+prod+")+("+rnd+")";
    if(m===1) sh.getRange(3,col).setFormula("=IF(1>"+HOR+",\"\",MAX("+seed+","+req+"))");
    else      sh.getRange(3,col).setFormula("=IF("+ms+">"+HOR+",\"\",MAX("+Cp+"3,"+req+"))");
    sh.getRange(3,col).setNumberFormat("0");
    sh.getRange(4,col).setFormula("=IF("+ms+"<="+HOR+","+C+"3*('🎛️ Drivers'!C"+DR.ENG+"/12*"+LOD+"+'🎛️ Drivers'!D"+DR.ENG+"+'🎛️ Drivers'!F"+DR.ENG+"),\"\")").setNumberFormat("$#,##0");
  }

  label(sh,5,1,"Sales (reps + AEs)  — HC");
  for(var m=1;m<=MAX;m++){var col=m+1;var C=colLetter(col);var ms=String(m);sh.getRange(5,col).setFormula("=IF("+ms+"<="+HOR+","+C+"18+"+C+"19,\"\")").setNumberFormat("0");}

  label(sh,6,1,"Sales (reps + AEs)  — Cost / mo");
  for(var m=1;m<=MAX;m++){
    var col=m+1; var C=colLetter(col); var ms=String(m);
    var rc=C+"18*('🎛️ Drivers'!C"+DR.SALES+"/12*"+LOD+"+'🎛️ Drivers'!D"+DR.SALES+"+'🎛️ Drivers'!F"+DR.SALES+")";
    var ac=C+"19*('🎛️ Drivers'!"+DR.AE_SALARY+"/12*"+LOD+"+'🎛️ Drivers'!"+DR.AE_SW+"+'🎛️ Drivers'!F"+DR.SALES+")";
    sh.getRange(6,col).setFormula("=IF("+ms+"<="+HOR+","+rc+"+"+ac+",\"\")").setNumberFormat("$#,##0");
  }

  label(sh,7,1,"CS / FDE-CSE (impl + ongoing)  — HC");
  label(sh,8,1,"CS / FDE-CSE  — Cost / mo");
  for(var m=1;m<=MAX;m++){
    var col=m+1; var C=colLetter(col); var Cp=colLetter(col-1); var ms=String(m);
    var newMM_cs="MAX(0,"+C+"15-'🎛️ Drivers'!"+DR.EXIST_MM_LOGOS+")";
    var newEN_cs="MAX(0,"+C+"16-'🎛️ Drivers'!"+DR.EXIST_ENT_LOGOS+")";
    var oMM="'🎛️ Drivers'!"+DR.CSM_MM_RATIO; var oENT="'🎛️ Drivers'!"+DR.CSM_ENT_RATIO;
    var fMM="'🎛️ Drivers'!"+DR.FDE_MM_CAPACITY; var fENT="'🎛️ Drivers'!"+DR.FDE_ENT_CAPACITY;
    var nMM ="IFERROR("+REV_TAB+"!"+MM_L+(DS+m-1)+",0)";
    var nENT="IFERROR("+REV_TAB+"!"+ENT_L+(DS+m-1)+",0)";
    var ongoing="IFERROR(CEILING("+newMM_cs+"/"+oMM+",1),0)+IFERROR(CEILING("+newEN_cs+"/"+oENT+",1),0)";
    var impl="IFERROR(CEILING("+nMM+"/"+fMM+",1),0)+IFERROR(CEILING("+nENT+"/"+fENT+",1),0)";
    var seed="'🎛️ Drivers'!B"+DR.CS; var needed="MAX(1,"+ongoing+"+"+impl+")";
    if(m===1) sh.getRange(7,col).setFormula("=IF(1>"+HOR+",\"\",MAX("+seed+","+needed+"))");
    else      sh.getRange(7,col).setFormula("=IF("+ms+">"+HOR+",\"\",MAX("+Cp+"7,MAX("+seed+","+needed+")))");
    sh.getRange(7,col).setNumberFormat("0");
    sh.getRange(8,col).setFormula("=IF("+ms+"<="+HOR+","+C+"7*('🎛️ Drivers'!C"+DR.CS+"/12*"+LOD+"+'🎛️ Drivers'!D"+DR.CS+"+'🎛️ Drivers'!F"+DR.CS+"),\"\")").setNumberFormat("$#,##0");
  }

  label(sh,9,1,"G&A  — HC"); label(sh,10,1,"G&A  — Cost / mo");
  for(var m=1;m<=MAX;m++){
    var col=m+1; var C=colLetter(col); var Cp=colLetter(col-1); var ms=String(m);
    var nonGA="("+C+"3+"+C+"5+"+C+"7)"; var gaR="'🎛️ Drivers'!"+DR.GA_RATIO;
    var seed="'🎛️ Drivers'!B"+DR.GA; var need="MAX(1,IFERROR(CEILING("+nonGA+"/"+gaR+",1),1))";
    if(m===1) sh.getRange(9,col).setFormula("=IF(1>"+HOR+",\"\","+seed+")");
    else      sh.getRange(9,col).setFormula("=IF("+ms+">"+HOR+",\"\",MAX("+Cp+"9,MAX("+seed+","+need+")))");
    sh.getRange(9,col).setNumberFormat("0");
    sh.getRange(10,col).setFormula("=IF("+ms+"<="+HOR+","+C+"9*('🎛️ Drivers'!C"+DR.GA+"/12*"+LOD+"+'🎛️ Drivers'!D"+DR.GA+"+'🎛️ Drivers'!F"+DR.GA+"),\"\")").setNumberFormat("$#,##0");
  }

  hdr(sh,11,1,"Total Monthly Payroll","#2C3E50");
  for(var m=1;m<=MAX;m++){var col=m+1;var C=colLetter(col);var ms=String(m);sh.getRange(11,col).setFormula("=IF("+ms+"<="+HOR+","+C+"4+"+C+"6+"+C+"8+"+C+"10,\"\")").setNumberFormat("$#,##0").setFontWeight("bold").setBackground("#D5F5E3");}

  hdr(sh,12,1,"Total Headcount (all depts)","#2C3E50");
  for(var m=1;m<=MAX;m++){var col=m+1;var C=colLetter(col);var ms=String(m);sh.getRange(12,col).setFormula("=IF("+ms+"<="+HOR+","+C+"3+"+C+"5+"+C+"7+"+C+"9,\"\")").setNumberFormat("0").setFontWeight("bold").setBackground("#EBF5FB");}

  sh.getRange(13,1,1,MAX+1).merge().setValue("Eng + CS scale with new logos only. AEs scale with total logos (existing + new).").setFontStyle("italic").setFontColor("#888").setWrap(true).setBackground("#FDFEFE"); sh.setRowHeight(13,36);
  sectionHdr(sh,14,"📊 Helper Rows — Customer counts + sub-dept HC");

  label(sh,15,1,"Active MM Logos (new + existing)");
  label(sh,16,1,"Active ENT Logos (new + existing)");
  hdr(sh,17,1,"Total Active Customers","#2C3E50");
  for(var m=1;m<=MAX;m++){
    var col=m+1; var C=colLetter(col); var ms=String(m);
    var mmC="IFERROR(SUMIF("+REV_TAB+"!"+MO_L+DS+":"+MO_L+(DS+MAX-1)+",\"<=\"&"+ms+","+REV_TAB+"!"+MM_L+DS+":"+MM_L+(DS+MAX-1)+"),0)+'🎛️ Drivers'!"+DR.EXIST_MM_LOGOS;
    var enC="IFERROR(SUMIF("+REV_TAB+"!"+MO_L+DS+":"+MO_L+(DS+MAX-1)+",\"<=\"&"+ms+","+REV_TAB+"!"+ENT_L+DS+":"+ENT_L+(DS+MAX-1)+"),0)+'🎛️ Drivers'!"+DR.EXIST_ENT_LOGOS;
    sh.getRange(15,col).setFormula("=IF("+ms+">"+HOR+",\"\","+mmC+")").setNumberFormat("0").setBackground("#F2F3F4");
    sh.getRange(16,col).setFormula("=IF("+ms+">"+HOR+",\"\","+enC+")").setNumberFormat("0").setBackground("#F2F3F4");
    sh.getRange(17,col).setFormula("=IF("+ms+">"+HOR+",\"\","+C+"15+"+C+"16)").setNumberFormat("0").setBackground("#D6EAF8").setFontWeight("bold");
  }

  label(sh,18,1,"  ↳ Sales Reps HC (helper)");
  for(var m=1;m<=MAX;m++){
    var col=m+1; var C=colLetter(col); var Cp=colLetter(col-1); var ms=String(m);
    var nMM ="IFERROR("+REV_TAB+"!"+MM_L+(DS+m-1)+",0)";
    var nENT="IFERROR("+REV_TAB+"!"+ENT_L+(DS+m-1)+",0)";
    var ewt="'🎛️ Drivers'!"+DR.ENT_SALES_WEIGHT; var cap="'🎛️ Drivers'!"+DR.SALES_REP_CAP; var att="'🎛️ Drivers'!"+DR.ATTAINMENT;
    var seed="'🎛️ Drivers'!B"+DR.SALES; var need="CEILING(("+nMM+"+"+nENT+"*"+ewt+")/("+cap+"*"+att+"),1)";
    if(m===1) sh.getRange(18,col).setFormula("=IF(1>"+HOR+",\"\",MAX("+seed+","+need+"))");
    else      sh.getRange(18,col).setFormula("=IF("+ms+">"+HOR+",\"\",MAX("+Cp+"18,MAX("+seed+","+need+")))");
    sh.getRange(18,col).setNumberFormat("0").setBackground("#F2F3F4");
  }

  label(sh,19,1,"  ↳ AE HC (helper)");
  for(var m=1;m<=MAX;m++){
    var col=m+1; var C=colLetter(col); var Cp=colLetter(col-1); var ms=String(m);
    var aeCap="'🎛️ Drivers'!"+DR.AE_RATIO; var att="'🎛️ Drivers'!"+DR.ATTAINMENT;
    var need="MAX(1,CEILING("+C+"17/("+aeCap+"*"+att+"),1))";
    if(m===1) sh.getRange(19,col).setFormula("=IF(1>"+HOR+",\"\","+need+")");
    else      sh.getRange(19,col).setFormula("=IF("+ms+">"+HOR+",\"\",MAX("+Cp+"19,"+need+"))");
    sh.getRange(19,col).setNumberFormat("0").setBackground("#F2F3F4");
  }

  sectionHdr(sh,20,"📊 Investor Metrics (auto)");
  hdr(sh,21,1,"ARR per Employee ($)","#2C3E50");
  for(var m=1;m<=MAX;m++){var col=m+1;var C=colLetter(col);var ms=String(m);sh.getRange(21,col).setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR('💸 P&L'!"+C+PNL.ARR+"/"+C+"12,\"\"))").setNumberFormat("$#,##0").setBackground("#E8F8F5").setFontWeight("bold");}
  hdr(sh,22,1,"Payroll as % of ARR","#2C3E50");
  for(var m=1;m<=MAX;m++){var col=m+1;var C=colLetter(col);var ms=String(m);sh.getRange(22,col).setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR("+C+"11*12/'💸 P&L'!"+C+PNL.ARR+",\"\"))").setNumberFormat("0%").setBackground("#FEF9E7").setFontWeight("bold");}
  sectionHdr(sh,23,"HC Mix by Department (% of total)");
  [{label:"Engineering %",r:3},{label:"Sales %",r:5},{label:"FDE/CSE %",r:7},{label:"G&A %",r:9}]
    .forEach(function(d,i){
      var row=24+i; label(sh,row,1,d.label);
      for(var m=1;m<=MAX;m++){var col=m+1;var C=colLetter(col);var ms=String(m);sh.getRange(row,col).setFormula("=IF("+ms+">"+HOR+",\"\",IFERROR("+C+d.r+"/"+C+"12,\"\"))").setNumberFormat("0%").setBackground("#F2F3F4");}
    });
}
