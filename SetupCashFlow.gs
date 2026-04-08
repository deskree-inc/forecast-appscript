// SetupCashFlow.gs — Called from setupFinancialModel in SetupMain.gs.

function setupCashFlow(ss) {
  var sh     = ss.getSheetByName("🏦 Cash Flow");
  var months = 60;
  var HOR    = "'🎛️ Drivers'!" + DR.HORIZON;
  var FS     = "'🎛️ Drivers'!" + DR.FORECAST_START;

  if (sh.getMaxColumns() < months + 1)
    sh.insertColumnsAfter(sh.getMaxColumns(), months + 1 - sh.getMaxColumns());
  sh.setColumnWidth(1, 280);
  for (var m = 1; m <= months; m++) sh.setColumnWidth(m + 1, 78);

  sh.getRange(CF.HDR_TITLE, 1, 1, months + 1).merge();
  hdr(sh, CF.HDR_TITLE, 1, "🏦 CASH FLOW — Direct Method  |  All inputs in 🎛️ Drivers  |  Do not edit formulas", "#1A5276");

  hdr(sh, CF.HDR_COLS, 1, "Line Item", "#1F618D");
  for (var m = 1; m <= months; m++) {
    sh.getRange(CF.HDR_COLS, m + 1)
      .setFormula("=IF(" + m + ">" + HOR + ",\"\",TEXT(EDATE(" + FS + "," + (m-1) + "),\"MMM YY\"))")
      .setBackground("#1F618D").setFontColor("#FFFFFF")
      .setFontWeight("bold").setHorizontalAlignment("center").setNumberFormat("@");
  }

  _secHdr(sh, CF.SEC_OPS, "OPERATING ACTIVITIES", months);
  sh.getRange(CF.CUST_CASH,  1).setValue("  Cash from Customers (MRR)").setFontColor("#1D9E75").setFontWeight("bold");
  sh.getRange(CF.INT_INCOME, 1).setValue("  Interest Income ($)").setFontColor("#1D9E75").setFontStyle("italic");
  sh.getRange(CF.PAY_ENG,    1).setValue("  Engineering Payroll");
  sh.getRange(CF.PAY_SALES,  1).setValue("  Sales Payroll");
  sh.getRange(CF.PAY_CS,     1).setValue("  CS / FDE-CSE Payroll");
  sh.getRange(CF.PAY_GA,     1).setValue("  G&A Payroll");
  sh.getRange(CF.COGS_INFRA, 1).setValue("  Infrastructure / COGS");
  sh.getRange(CF.SM_NONPAY,  1).setValue("  Sales & Marketing — non-payroll");
  sh.getRange(CF.GA_NONPAY,  1).setValue("  G&A — non-payroll (tools, fees, recruiting, hardware)");
  sh.getRange(CF.NET_OPS,    1).setValue("Net Operating Cash Flow ($)").setFontWeight("bold");

  _secHdr(sh, CF.SEC_FIN, "FINANCING ACTIVITIES", months);
  sh.getRange(CF.CAP_RAISED, 1).setValue("  Capital Raised ($)").setFontColor("#1D9E75").setFontWeight("bold");
  sh.getRange(CF.NET_FIN,    1).setValue("Net Financing Cash Flow ($)").setFontWeight("bold");

  _secHdr(sh, CF.SEC_INV, "INVESTING ACTIVITIES", months);
  sh.getRange(CF.CAPEX,   1).setValue("  CapEx — Hardware expensed in G&A (Investing = $0 until capitalised)").setFontColor("#AAAAAA").setFontStyle("italic");
  sh.getRange(CF.NET_INV, 1).setValue("Net Investing Cash Flow ($)").setFontWeight("bold");

  _secHdr(sh, CF.SEC_SUMMARY, "CASH SUMMARY", months);
  sh.getRange(CF.BEG_CASH,  1).setValue("Beginning Cash ($)");
  sh.getRange(CF.PLUS_OPS,  1).setValue("  + Operating Activities");
  sh.getRange(CF.PLUS_FIN,  1).setValue("  + Financing Activities");
  sh.getRange(CF.PLUS_INV,  1).setValue("  + Investing Activities");
  sh.getRange(CF.END_CASH,  1).setValue("Ending Cash ($)").setFontWeight("bold");

  _secHdr(sh, CF.SEC_METRICS, "KEY METRICS", months);
  sh.getRange(CF.BURN_RATE,  1).setValue("Monthly Net Burn ($)").setFontWeight("bold");
  sh.getRange(CF.RUNWAY,     1).setValue("Runway Remaining (months)").setFontWeight("bold");
  sh.getRange(CF.CASH_ARR,   1).setValue("Cash as % of ARR");
  sh.getRange(CF.CUMUL_CAP,  1).setValue("Cumulative Capital Deployed ($)");
  sh.getRange(CF.CUMUL_BURN, 1).setValue("Cumulative Net Burn ($)");

  sh.getRange(CF.CAPEX,      1).setNote("At seed/Series A stage, hardware per hire (~$2K) is expensed, not capitalised.");
  sh.getRange(CF.BURN_RATE,  1).setNote("Operating + Investing outflows only. Excludes capital raises and interest income.");
  sh.getRange(CF.RUNWAY,     1).setNote("Ending cash ÷ |Monthly Net Burn|. Red when < 6 months.");
  sh.getRange(CF.CASH_ARR,   1).setNote("Ending cash as a multiple of annualised MRR. Investors expect ≥ 12 months of runway.");
  sh.getRange(CF.INT_INCOME, 1).setNote("US GAAP: interest received = Operating Activity. Pulled from 💰 Funding schedule.");

  for (var m = 1; m <= months; m++) {
    var col     = m + 1;
    var C       = colLetter(col);
    var Cp      = colLetter(col - 1);
    var ms      = String(m);
    var fundRow = FUND.MONTHLY_START + m - 1;

    sh.getRange(CF.CUST_CASH, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",IFERROR('💸 P&L'!" + C + PNL.MRR + ",0))")
      .setNumberFormat("$#,##0").setBackground("#D5F5E3").setFontColor("#1D9E75");
    sh.getRange(CF.INT_INCOME, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",IFERROR('💰 Funding'!D" + fundRow + ",0))")
      .setNumberFormat("$#,##0").setBackground("#E8F8F5").setFontColor("#1D9E75");
    sh.getRange(CF.PAY_ENG, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",-'👥 Headcount'!" + C + "4)")
      .setNumberFormat("$#,##0").setBackground("#FFFFFF");
    sh.getRange(CF.PAY_SALES, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",-'👥 Headcount'!" + C + "6)")
      .setNumberFormat("$#,##0").setBackground("#FFFFFF");
    sh.getRange(CF.PAY_CS, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",-'👥 Headcount'!" + C + "8)")
      .setNumberFormat("$#,##0").setBackground("#FFFFFF");
    sh.getRange(CF.PAY_GA, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",-'👥 Headcount'!" + C + "10)")
      .setNumberFormat("$#,##0").setBackground("#FFFFFF");
    sh.getRange(CF.COGS_INFRA, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",IFERROR(-'💸 P&L'!" + C + PNL.INFRA + ",0))")
      .setNumberFormat("$#,##0").setBackground("#FFFFFF");
    sh.getRange(CF.SM_NONPAY, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\","
        + "IFERROR(-('💸 P&L'!" + C + PNL.SM_SUBTOTAL + "-'💸 P&L'!" + C + PNL.SALES_PAYROLL + "),0))")
      .setNumberFormat("$#,##0").setBackground("#FFFFFF");
    sh.getRange(CF.GA_NONPAY, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\","
        + "IFERROR(-('💸 P&L'!" + C + PNL.GA_SUBTOTAL + "-'💸 P&L'!" + C + PNL.GA_PAYROLL + "),0))")
      .setNumberFormat("$#,##0").setBackground("#FFFFFF");
    sh.getRange(CF.NET_OPS, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",SUM(" + C + CF.CUST_CASH + ":" + C + CF.GA_NONPAY + "))")
      .setNumberFormat("$#,##0").setFontWeight("bold");

    sh.getRange(CF.CAP_RAISED, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",IFERROR('💰 Funding'!C" + fundRow + ",0))")
      .setNumberFormat("$#,##0").setBackground("#D5F5E3").setFontColor("#1D9E75");
    sh.getRange(CF.NET_FIN, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + C + CF.CAP_RAISED + ")")
      .setNumberFormat("$#,##0").setFontWeight("bold");

    sh.getRange(CF.CAPEX, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",0)")
      .setNumberFormat("$#,##0").setFontColor("#AAAAAA");
    sh.getRange(CF.NET_INV, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + C + CF.CAPEX + ")")
      .setNumberFormat("$#,##0").setFontWeight("bold").setFontColor("#AAAAAA");

    if (m === 1) {
      sh.getRange(CF.BEG_CASH, col)
        .setFormula("=IF(1>" + HOR + ",\"\",IFERROR('🎛️ Drivers'!" + DR.OPENING_CASH + ",0))")
        .setNumberFormat("$#,##0").setBackground("#F2F3F4");
    } else {
      sh.getRange(CF.BEG_CASH, col)
        .setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + Cp + CF.END_CASH + ")")
        .setNumberFormat("$#,##0").setBackground("#F2F3F4");
    }
    sh.getRange(CF.PLUS_OPS, col).setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + C + CF.NET_OPS + ")").setNumberFormat("$#,##0").setBackground("#FFFFFF");
    sh.getRange(CF.PLUS_FIN, col).setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + C + CF.NET_FIN + ")").setNumberFormat("$#,##0").setBackground("#F8F9FA");
    sh.getRange(CF.PLUS_INV, col).setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + C + CF.NET_INV + ")").setNumberFormat("$#,##0").setBackground("#FFFFFF").setFontColor("#AAAAAA");
    sh.getRange(CF.END_CASH, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + C + CF.BEG_CASH + "+" + C + CF.PLUS_OPS + "+" + C + CF.PLUS_FIN + "+" + C + CF.PLUS_INV + ")")
      .setNumberFormat("$#,##0").setFontWeight("bold");

    sh.getRange(CF.BURN_RATE, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + C + CF.NET_OPS + "+" + C + CF.NET_INV + ")")
      .setNumberFormat("$#,##0").setFontWeight("bold");
    sh.getRange(CF.RUNWAY, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",IF(" + C + CF.BURN_RATE + ">=0,\"\",IFERROR(" + C + CF.END_CASH + "/ABS(" + C + CF.BURN_RATE + "),0)))")
      .setNumberFormat("0.0").setFontWeight("bold");
    sh.getRange(CF.CASH_ARR, col)
      .setFormula("=IF(" + ms + ">" + HOR + ",\"\",IFERROR(" + C + CF.END_CASH + "/IFERROR('💸 P&L'!" + C + PNL.ARR + ",1),0))")
      .setNumberFormat("0%");

    if (m === 1) {
      sh.getRange(CF.CUMUL_CAP, col).setFormula("=IF(1>" + HOR + ",\"\"," + C + CF.CAP_RAISED + ")").setNumberFormat("$#,##0");
      sh.getRange(CF.CUMUL_BURN, col).setFormula("=IF(1>" + HOR + ",\"\"," + C + CF.BURN_RATE + ")").setNumberFormat("$#,##0").setFontWeight("bold");
    } else {
      sh.getRange(CF.CUMUL_CAP, col).setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + Cp + CF.CUMUL_CAP + "+" + C + CF.CAP_RAISED + ")").setNumberFormat("$#,##0");
      sh.getRange(CF.CUMUL_BURN, col).setFormula("=IF(" + ms + ">" + HOR + ",\"\"," + Cp + CF.CUMUL_BURN + "+" + C + CF.BURN_RATE + ")").setNumberFormat("$#,##0").setFontWeight("bold");
    }
  }

  [CF.CUST_CASH, CF.INT_INCOME, CF.PAY_ENG, CF.PAY_SALES, CF.PAY_CS, CF.PAY_GA,
   CF.COGS_INFRA, CF.SM_NONPAY, CF.GA_NONPAY]
    .forEach(function(r, i) { sh.getRange(r, 1).setBackground(i % 2 === 0 ? "#FFFFFF" : "#F8F9FA"); });
  sh.getRange(CF.CAP_RAISED, 1).setBackground("#FFFFFF");
  sh.getRange(CF.CAPEX, 1).setBackground("#F8F9FA");
  sh.getRange(CF.PLUS_OPS, 1).setBackground("#FFFFFF");
  sh.getRange(CF.PLUS_FIN, 1).setBackground("#F8F9FA");
  sh.getRange(CF.PLUS_INV, 1).setBackground("#FFFFFF");
  sh.getRange(CF.CASH_ARR,   1).setBackground("#FFFFFF");
  sh.getRange(CF.CUMUL_CAP,  1).setBackground("#F8F9FA");
  sh.getRange(CF.CUMUL_BURN, 1).setBackground("#FFFFFF");

  [CF.NET_OPS, CF.NET_FIN, CF.NET_INV, CF.BEG_CASH, CF.END_CASH, CF.BURN_RATE, CF.RUNWAY]
    .forEach(function(r) { sh.getRange(r, 1).setFontWeight("bold").setFontColor("#000000"); });
  [CF.CUST_CASH, CF.INT_INCOME, CF.PAY_ENG, CF.PAY_SALES, CF.PAY_CS, CF.PAY_GA,
   CF.COGS_INFRA, CF.SM_NONPAY, CF.GA_NONPAY, CF.CAP_RAISED,
   CF.PLUS_OPS, CF.PLUS_FIN, CF.PLUS_INV, CF.CASH_ARR, CF.CUMUL_CAP, CF.CUMUL_BURN]
    .forEach(function(r) { sh.getRange(r, 1).setFontWeight("normal").setFontColor("#444444"); });
  sh.getRange(CF.CAPEX, 1).setFontWeight("normal").setFontColor("#AAAAAA").setFontStyle("italic");
  [CF.NET_OPS, CF.NET_FIN, CF.NET_INV, CF.BEG_CASH, CF.END_CASH, CF.BURN_RATE, CF.RUNWAY]
    .forEach(function(r) { sh.getRange(r, 1).setBackground("#F2F3F4"); });

  for (var r = 1; r <= CF.NOTE + 1; r++) sh.setRowHeight(r, 20);
  [CF.SEC_OPS, CF.SEC_FIN, CF.SEC_INV, CF.SEC_SUMMARY, CF.SEC_METRICS].forEach(function(r) { sh.setRowHeight(r, 24); });
  [CF.NET_OPS, CF.NET_FIN, CF.NET_INV, CF.END_CASH, CF.BURN_RATE, CF.RUNWAY].forEach(function(r) { sh.setRowHeight(r, 22); });
  sh.setRowHeight(CF.NOTE, 32);

  [CF.NET_OPS, CF.NET_FIN, CF.NET_INV].forEach(function(r) {
    sh.getRange(r, 1, 1, months + 1).setBorder(true,false,false,false,false,false,"#AAAAAA",SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.getRange(CF.END_CASH, 1, 1, months + 1).setBorder(true,false,true,false,false,false,"#AAAAAA",SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(CF.BURN_RATE, 1, 1, months + 1).setBorder(true,false,false,false,false,false,"#AAAAAA",SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(CF.NOTE, 1, 1, months + 1).merge()
    .setValue("✅ Direct-method cash flow. Hardware expensed in G&A (Operating) — Investing = $0 until the company capitalises CapEx. All formulas auto-calculate from 🎛️ Drivers — do not edit this tab.")
    .setFontStyle("italic").setFontColor("#1D9E75").setBackground("#FDFEFE").setWrap(true);

  var rules = sh.getConditionalFormatRules();
  var ranges = {
    netOps:    sh.getRange(CF.NET_OPS,    2, 1, months),
    netFin:    sh.getRange(CF.NET_FIN,    2, 1, months),
    endCash:   sh.getRange(CF.END_CASH,   2, 1, months),
    burnRate:  sh.getRange(CF.BURN_RATE,  2, 1, months),
    runway:    sh.getRange(CF.RUNWAY,     2, 1, months),
    cashArr:   sh.getRange(CF.CASH_ARR,   2, 1, months),
    cumulBurn: sh.getRange(CF.CUMUL_BURN, 2, 1, months)
  };
  function cfGR(val,bg,fc,rng){rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(val).setBackground(bg).setFontColor(fc).setRanges([rng]).build());}
  function cfLT(val,bg,fc,rng){rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(val).setBackground(bg).setFontColor(fc).setRanges([rng]).build());}
  function cfBT(a,b,bg,fc,rng){rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(a,b).setBackground(bg).setFontColor(fc).setRanges([rng]).build());}
  function cfEQ(val,bg,fc,rng){rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(val).setBackground(bg).setFontColor(fc).setRanges([rng]).build());}
  cfGR(0,"#D5F5E3","#1D9E75",ranges.netOps);   cfLT(0,"#FADBD8","#922B21",ranges.netOps);
  cfGR(0,"#D5F5E3","#1D9E75",ranges.netFin);   cfEQ(0,"#F2F3F4","#888888",ranges.netFin);
  cfGR(0,"#D5F5E3","#1D9E75",ranges.endCash);  cfLT(0,"#FADBD8","#922B21",ranges.endCash);
  cfGR(0,"#D5F5E3","#1D9E75",ranges.burnRate); cfLT(0,"#FADBD8","#922B21",ranges.burnRate);
  cfLT(6,"#FADBD8","#922B21",ranges.runway);
  cfBT(6,12,"#FEF9E7","#D4AC0D",ranges.runway);
  cfGR(12,"#D5F5E3","#1D9E75",ranges.runway);
  cfGR(1,"#D5F5E3","#1D9E75",ranges.cashArr);
  cfBT(0.5,1,"#FEF9E7","#D4AC0D",ranges.cashArr);
  cfLT(0.5,"#FADBD8","#922B21",ranges.cashArr);
  cfLT(0,"#FADBD8","#922B21",ranges.cumulBurn);
  sh.setConditionalFormatRules(rules);
}
