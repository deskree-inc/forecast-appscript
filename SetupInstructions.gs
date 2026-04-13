// SetupInstructions.gs — Called from setupFinancialModel in SetupMain.gs.

function setupInstructions(ss) {
  var sh = ss.getSheetByName(SHEET_INSTRUCTIONS);
  sh.setColumnWidth(1,30); sh.setColumnWidth(2,200);
  sh.setColumnWidth(3,620); sh.setColumnWidth(4,30);
  function title(row,text){sh.getRange(row,1,1,4).merge().setValue(text).setBackground("#1A5276").setFontColor("#FFFFFF").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");sh.setRowHeight(row,40);}
  function sec(row,text){sh.getRange(row,2,1,2).merge().setValue(text).setBackground("#D6EAF8").setFontWeight("bold").setFontSize(11);sh.setRowHeight(row,24);}
  function irow(r,lbl,content,bgL,bgC){sh.getRange(r,2).setValue(lbl).setBackground(bgL||"#F2F3F4").setFontWeight("bold").setVerticalAlignment("top").setWrap(true);sh.getRange(r,3).setValue(content).setBackground(bgC||"#FFFFFF").setVerticalAlignment("top").setWrap(true);sh.setRowHeight(r,56);}
  function note(r,text){sh.getRange(r,2,1,2).merge().setValue(text).setFontStyle("italic").setFontColor("#717D7E").setBackground("#FDFEFE").setWrap(true);sh.setRowHeight(r,36);}
  function blank(r){sh.setRowHeight(r,12);}
  var r=1;
  title(r,"📖 Financial model — how to use");r++;blank(r);r++;
  sec(r,"🗺️ Overview");r++;
  sh.getRange(r,2,1,2).merge().setValue("The ONLY tab you type into is 🎛️ Drivers. Everything else is formulas.\n\nHeadcount is fully automated. Revenue logos are back-calculated from Target ARR.").setBackground("#EBF5FB").setWrap(true).setVerticalAlignment("top");sh.setRowHeight(r,90);r++;
  blank(r);r++;sec(r,"🎨 Color Legend");r++;
  irow(r,"🔵 Blue","Input — only cells you edit. Found only in 🎛️ Drivers.","#D6EAF8","#EBF5FB");r++;
  irow(r,"⚫ Black","Formula — do not edit.");r++;
  irow(r,"🟢 Green","Key output — ARR, cash, gross profit.");r++;blank(r);r++;
  sec(r,"📑 Tab Guide");r++;
  irow(r,"🎛️ Drivers","START HERE. All inputs live here.");r++;
  irow(r,"💰 Funding","Funding rounds + monthly capital & interest schedule.");r++;
  irow(r,"👥 Headcount","Fully automated HC + payroll.");r++;
  irow(r,"📈 Revenue","Back-calculated ARR waterfall + existing book.");r++;
  irow(r,"💸 P&L","Full income statement: Revenue → COGS → OpEx → EBITDA → Key Metrics.");r++;
  irow(r,"🏦 Cash Flow","Direct-method statement: Operating / Financing / Investing + Key Metrics.");r++;
  irow(r,"📊 Summary","Investor KPI dashboard.");r++;
  irow(r,"📋 Scenarios","Side-by-side scenario comparison.");r++;
  irow(r,"🚦 Benchmarks","Reality check.","#FADBD8","#FEF9E7");r++;
  blank(r);r++;note(r,"Re-run setupFinancialModel() to reset. All data will be cleared.");
}
