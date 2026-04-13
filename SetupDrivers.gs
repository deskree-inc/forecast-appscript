// SetupDrivers.gs — Called from setupFinancialModel in SetupMain.gs.

function setupDrivers(ss) {
  var sh = ss.getSheetByName("🎛️ Drivers");
  sh.setColumnWidth(1,290);
  [140,130,120,120,110,130,110,110].forEach(function(w,i){sh.setColumnWidth(i+2,w);});

  hdr(sh,1,1,"🎛️ DRIVERS — Control Panel (edit BLUE cells only)","#1A5276");
  sh.getRange(1,1,1,9).merge()
    .setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1,36);

  function note3(row, text, span) {
    var rng = sh.getRange(row, 3, 1, span || 1);
    if (span && span > 1) rng.merge();
    rng.setValue(text)
      .setBackground("#FEF9E7").setFontColor("#7D6608")
      .setFontStyle("italic").setWrap(true);
  }
  function secHdrStyled(row, text) {
    sh.getRange(row, 1, 1, 9).merge()
      .setValue(text)
      .setBackground("#AED6F1").setFontWeight("bold")
      .setFontColor("#1A5276").setFontSize(10)
      .setBorder(false,false,false,false,false,false)
      .setBorder(true,false,false,false,false,false,
        "#1A5276", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.setRowHeight(row, 26);
  }
  function lbl(row, col, text) {
    sh.getRange(row, col).setValue(text)
      .setFontWeight("bold").setFontColor("#444444");
  }

  // B: ARR Targets
  secHdrStyled(11,"B — ARR Targets");
  lbl(12,1,"Target ARR (auto from Section L)");
  sh.getRange(12,2)
    .setFormula("=IFERROR(IF(IF(B14<=12,B132,IF(B14<=24,B133,IF(B14<=36,B134,IF(B14<=48,B135,B136))))=0,\"⚠️ Set target in Section L\",IF(B14<=12,B132,IF(B14<=24,B133,IF(B14<=36,B134,IF(B14<=48,B135,B136))))),\"⚠️ Set target in Section L\")")
    .setBackground("#E8F8F5").setFontWeight("bold").setNumberFormat("$#,##0").setFontColor("#1D9E75");
  var b12Rules = sh.getConditionalFormatRules();
  b12Rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("⚠️")
    .setBackground("#FADBD8").setFontColor("#922B21")
    .setRanges([sh.getRange(12,2)])
    .build());
  sh.setConditionalFormatRules(b12Rules);
  note3(12,"Auto-selected from Section L based on Forecast Horizon (B14). Do not edit this cell.",2);
  sh.setRowHeight(12,36);
  lbl(13,1,"Target MoM Growth Rate");       inp(sh,13,2,0.20,"0%");
  lbl(14,1,"Forecast Horizon (months)");    inp(sh,14,2,24);
  [13,14].forEach(function(r){sh.setRowHeight(r,22);});

  // C: ICP Segments
  secHdrStyled(16,"C — ICP Segments");
  ["Segment","Beg. ACV ($)","Exp. ACV ($)","Churn (ann.)","Expansion % (ann.)","Expansion Mo","CAC ($)","Lead Time (mo)","Close Rate"]
    .forEach(function(h,i){hdr(sh,17,i+1,h,"#1F618D");});
  sh.setRowHeight(17,22);
  [["Mid-Market",80000,150000,0.08,0.30,12,15000,3,0.20],
   ["Enterprise",250000,1000000,0.05,0.40,18,50000,6,0.10]]
    .forEach(function(seg,i){
      var r=18+i; sh.getRange(r,1).setValue(seg[0]).setFontWeight("bold").setFontColor("#444444");
      inp(sh,r,2,seg[1],"$#,##0"); inp(sh,r,3,seg[2],"$#,##0");
      inp(sh,r,4,seg[3],"0%");     inp(sh,r,5,seg[4],"0%");
      inp(sh,r,6,seg[5]);          inp(sh,r,7,seg[6],"$#,##0");
      inp(sh,r,8,seg[7]);          inp(sh,r,9,seg[8],"0%");
      sh.setRowHeight(r,22);
    });

  // C2
  secHdrStyled(22,"C2 — Customer Maintenance Ratios");
  ["Role","Max Accounts / Person","Industry Benchmark"].forEach(function(h,i){hdr(sh,23,i+1,h,"#1F618D");});
  sh.setRowHeight(23,22);
  [["Account Executive (AE)",15,"10–20 for MM/ENT SaaS"],
   ["Customer Success Engineer (FDE/CSE)",20,"15–25 tech-touch"],
   ["Field Deploy Engineer (FDE)",10,"3–6 concurrent deployments"]]
    .forEach(function(row,i){
      var r=24+i;
      sh.getRange(r,1).setValue(row[0]).setFontWeight("bold").setFontColor("#444444");
      inp(sh,r,2,row[1]);
      note3(r,row[2]);
      sh.setRowHeight(r,22);
    });

  // C3
  secHdrStyled(28,"C3 — FDE Deployment Capacity");
  ["FDE Capacity","Value","Notes"].forEach(function(h,i){hdr(sh,29,i+1,h,"#1F618D");});
  sh.setRowHeight(29,22);
  lbl(30,1,"Concurrent MM deployments / FDE"); inp(sh,30,2,4,"0"); note3(30,"MM deploy: 1–3 months");
  lbl(31,1,"Concurrent ENT deployments / FDE"); inp(sh,31,2,1,"0"); note3(31,"ENT deploy: 4–9 months");
  [30,31].forEach(function(r){sh.setRowHeight(r,22);});

  // C4
  secHdrStyled(33,"C4 — Logo Back-Calculation");
  ["Driver","Value","Notes"].forEach(function(h,i){hdr(sh,34,i+1,h,"#1F618D");});
  sh.setRowHeight(34,22);
  lbl(35,1,"AE Annual Quota — MM ($ ACV)"); inp(sh,35,2,800000,"$#,##0"); note3(35,"Benchmarks only");
  lbl(36,1,"AE Annual Quota — ENT ($ ACV)"); inp(sh,36,2,1500000,"$#,##0"); note3(36,"Benchmarks only");
  lbl(37,1,"Quota Attainment %"); inp(sh,37,2,0.75,"0%"); note3(37,"Capacity discount for reps and AEs");
  lbl(38,1,"MM % of Target ARR"); inp(sh,38,2,0.60,"0%"); note3(38,"60% = MM gets $1.2M of a $2M target",2); sh.setRowHeight(38,36);
  lbl(39,1,"Logo MoM Growth Rate"); inp(sh,39,2,0.05,"0%"); note3(39,"5% = each month closes 5% more logos than prior.",2); sh.setRowHeight(39,36);
  [35,36,37].forEach(function(r){sh.setRowHeight(r,22);});

  // C5 — ✏️ CHANGED: Forecast start updated to May 2026
  secHdrStyled(40,"C5 — First Client Dates & Forecast Start");
  ["Driver","Date","Notes"].forEach(function(h,i){hdr(sh,41,i+1,h,"#1F618D");});
  sh.setRowHeight(41,22);
  lbl(42,1,"Forecast Start Date (Month 1)");
  sh.getRange(42,2).setValue(new Date("2026-05-01")).setBackground("#EBF5FB").setFontColor("#1A5276").setNumberFormat("MMM YYYY");
  sh.getRange(42,2).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  note3(42,"Month 1 of forecast"); sh.setRowHeight(42,22);
  lbl(43,1,"First MM Client Date");
  sh.getRange(43,2).setValue(new Date("2026-05-01")).setBackground("#EBF5FB").setFontColor("#1A5276").setNumberFormat("MMM YYYY");
  sh.getRange(43,2).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  note3(43,"MM logos start from forecast month 1");
  sh.getRange(43,4).setFormula("=MAX(0,(YEAR(B43)-YEAR(B42))*12+MONTH(B43)-MONTH(B42))").setNumberFormat("0").setFontColor("#AAAAAA").setFontStyle("italic");
  sh.getRange(43,5).setValue("← auto offset (do not edit)").setFontColor("#AAAAAA").setFontStyle("italic");
  sh.setRowHeight(43,22);
  lbl(44,1,"First ENT Client Date");
  sh.getRange(44,2).setValue(new Date("2026-07-01")).setBackground("#EBF5FB").setFontColor("#1A5276").setNumberFormat("MMM YYYY");
  sh.getRange(44,2).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  note3(44,"ENT logos = 0 before this date");
  sh.getRange(44,4).setFormula("=MAX(0,(YEAR(B44)-YEAR(B42))*12+MONTH(B44)-MONTH(B42))").setNumberFormat("0").setFontColor("#AAAAAA").setFontStyle("italic");
  sh.getRange(44,5).setValue("← auto offset (do not edit)").setFontColor("#AAAAAA").setFontStyle("italic");
  sh.setRowHeight(44,22);

  // D
  secHdrStyled(45,"D — Headcount: Department Defaults (seed values)");
  ["Department","Start HC","Annual Salary ($)","SW Cost/mo ($)","HW One-time ($)","Insurance/mo ($)"].forEach(function(h,i){hdr(sh,46,i+1,h,"#1F618D");});
  sh.setRowHeight(46,22);
  [["Engineering",4,180000,300,2500,200],["Sales",2,120000,150,1500,200],
   ["CS / FDE-CSE",1,90000,100,1500,200],["G&A",2,100000,80,1000,200]]
    .forEach(function(d,i){
      var r=47+i;
      sh.getRange(r,1).setValue(d[0]).setFontWeight("bold").setFontColor("#444444");
      inp(sh,r,2,d[1]);inp(sh,r,3,d[2],"$#,##0");inp(sh,r,4,d[3],"$#,##0");
      inp(sh,r,5,d[4],"$#,##0");inp(sh,r,6,d[5],"$#,##0");
      sh.setRowHeight(r,22);
    });

  // D2
  secHdrStyled(52,"D2 — Individual Positions (optional)");
  ["Title","Department","Start Date","Annual Salary ($)","SW Cost/mo ($)"].forEach(function(h,i){hdr(sh,53,i+1,h,"#1F618D");});
  sh.setRowHeight(53,22);
  for(var i=0;i<10;i++){
    var r=54+i;
    [1,2].forEach(function(c){sh.getRange(r,c).setBackground("#EBF5FB");});
    sh.getRange(r,3).setBackground("#EBF5FB").setNumberFormat("MMM YYYY");
    sh.getRange(r,4).setBackground("#EBF5FB").setNumberFormat("$#,##0");
    sh.getRange(r,5).setBackground("#EBF5FB").setNumberFormat("$#,##0");
    sh.setRowHeight(r,22);
  }

  // E
  secHdrStyled(65,"E — Marketing Budget (annual)");
  lbl(66,1,"Events Budget ($)");   inp(sh,66,2,50000,"$#,##0");
  lbl(67,1,"Digital / Other ($)"); inp(sh,67,2,30000,"$#,##0");
  lbl(68,1,"Events as % of Marketing");
  sh.getRange(68,2).setFormula("=B66/(B66+B67)").setNumberFormat("0%");
  sh.getRange(68,1).setFontWeight("bold").setFontColor("#444444");
  lbl(69,1,"Marketing budget growth Y2 (x)"); inp(sh,69,2,1.5,"0.0");
  note3(69,"1.5 = 50% more spend in Year 2 vs Year 1.",2); sh.setRowHeight(69,36);
  [66,67,68].forEach(function(r){sh.setRowHeight(r,22);});

  // F
  secHdrStyled(70,"F — Infrastructure Costs");
  lbl(71,1,"Infra / customer / mo ($)");   inp(sh,71,2,50,"$#,##0");
  lbl(72,1,"Tooling / engineer / mo ($)"); inp(sh,72,2,200,"$#,##0");
  [71,72].forEach(function(r){sh.setRowHeight(r,22);});

  // G
  secHdrStyled(74,"G — Sales Commission");
  lbl(75,1,"Commission (% of new-logo ARR)"); inp(sh,75,2,0.10,"0%");
  sh.setRowHeight(75,22);

  // Legend
  secHdrStyled(78,"Legend");
  sh.getRange(79,1).setValue("🔵 Blue = Input").setBackground("#EBF5FB").setFontColor("#1A5276").setFontWeight("bold");
  sh.getRange(79,2).setValue("⚫ Black = Formula").setFontWeight("bold").setFontColor("#444444");
  sh.getRange(79,3).setValue("🚦 Run Benchmarks after every scenario change").setFontStyle("italic").setFontColor("#888");
  sh.setRowHeight(79,22);

  // H
  secHdrStyled(81,"H — Headcount Scaling Rules");
  ["Driver","Value","Industry Benchmark / Notes"].forEach(function(h,i){hdr(sh,82,i+1,h,"#1F618D");});
  sh.setRowHeight(82,22);
  [["Product engineers per active MM customer",15,"0","1 eng per 15 MM customers. Range: 10–20."],
   ["Product engineers per active ENT customer",3,"0","1 eng per 3 ENT customers. Range: 2–5."],
   ["R&D engineers: 1 per N product engineers",3,"0","~25% of eng capacity on platform."],
   ["Sales rep ramp time (months to 100%)",3,"0","3 months standard for MM SaaS."],
   ["Sales rep capacity (MM-equiv deals / month)",2,"0","A ramped rep closes ~2 MM deals/month."],
   ["ENT deal weight (x MM deal effort)",4,"0","1 ENT deal = 4x MM sales effort. Range: 3–6x."],
   ["AE ramp time (months to 100%)",3,"0","AEs manage existing accounts. 3-month ramp."],
   ["AE annual base salary ($)",90000,"$#,##0","Industry standard: $80–100K at seed/Series A."],
   ["AE SW cost / mo ($)",100,"$#,##0","CRM and tooling per AE per month."],
   ["G&A ratio: 1 hire per N total employees",35,"0","1 G&A per 30–40 headcount. Use 35 as midpoint."],
   ["Loaded cost multiplier (x base salary)",1.25,"0.00","1.25x = +25% for taxes, benefits, equipment."],
   ["FDE/CSE ongoing support: MM customers per person",20,"0","1 FDE/CSE per 20 active MM customers."],
   ["FDE/CSE ongoing support: ENT customers per person",4,"0","1 FDE/CSE per 4 active ENT customers."]]
    .forEach(function(row,i){
      var r=83+i;
      lbl(r,1,row[0]);
      inp(sh,r,2,row[1],row[2]);
      note3(r,row[3],3);
      sh.setRowHeight(r,36);
    });
  sh.getRange(97,1,1,3).merge()
    .setValue("Section H drives automated hiring in Headcount. These are ratios only.")
    .setFontStyle("italic").setFontColor("#888").setWrap(true);
  sh.setRowHeight(97,28);

  // I
  secHdrStyled(99,"I — Existing Book at Forecast Start (day 1 of forecast)");
  ["Metric","Value","Notes"].forEach(function(h,i){hdr(sh,100,i+1,h,"#1F618D");});
  sh.setRowHeight(100,22);
  lbl(101,1,"Existing MM Logos"); inp(sh,101,2,3); note3(101,"Number of MM customers live at forecast start");
  lbl(102,1,"Existing MM ACV ($)"); inp(sh,102,2,65000,"$#,##0"); note3(102,"Likely lower than projected new-deal ACV — reflect actuals");
  lbl(103,1,"Existing ENT Logos"); inp(sh,103,2,1); note3(103,"Number of ENT customers live at forecast start");
  lbl(104,1,"Existing ENT ACV ($)"); inp(sh,104,2,200000,"$#,##0"); note3(104,"Likely lower than projected new-deal ACV — reflect actuals");
  lbl(105,1,"Existing MM ARR ($)");
  sh.getRange(105,2).setFormula("=B101*B102").setNumberFormat("$#,##0").setBackground("#E8F8F5").setFontWeight("bold");
  note3(105,"Auto-calculated: Existing MM Logos × Existing MM ACV");
  lbl(106,1,"Existing ENT ARR ($)");
  sh.getRange(106,2).setFormula("=B103*B104").setNumberFormat("$#,##0").setBackground("#E8F8F5").setFontWeight("bold");
  note3(106,"Auto-calculated: Existing ENT Logos × Existing ENT ACV");
  lbl(107,1,"Total Existing ARR ($)");
  sh.getRange(107,2).setFormula("=B105+B106").setNumberFormat("$#,##0").setBackground("#D5F5E3").setFontWeight("bold");
  note3(107,"Auto-sum. ARR base on day 1 of the forecast.",2); sh.setRowHeight(107,28);
  [101,102,103,104,105,106].forEach(function(r){sh.setRowHeight(r,22);});
  sh.getRange(109,1,1,3).merge()
    .setValue("Churn and expansion for existing customers use the same rates as Sections C (ICP Segments).")
    .setFontStyle("italic").setFontColor("#888").setWrap(true);
  sh.setRowHeight(109,36);

  // J
  secHdrStyled(111,"J — Additional OpEx Drivers");
  [["Recruiting cost (% of annual salary per new hire)", "B112", 0.15, "0%",    "15% = agency fee. Standard for tech roles."],
   ["Travel per ENT deal closed ($)",                    "B113", 5000, "$#,##0","Flights + hotel per ENT close. Typical: $3K–$8K."],
   ["Professional fees — annual ($)",                    "B114", 60000,"$#,##0","Legal + accounting + audit. ~$5K/mo at seed stage."],
   ["Company software per employee per month ($)",        "B115", 50,   "$#,##0","Slack, Notion, HR tools etc. Scales with total HC."],
   ["Hardware per new hire — one-time ($)",               "B116", 2000, "$#,##0","Laptop + peripherals. Triggered on month of hire."],
   ["Events travel as % of events budget",                "B117", 0.40, "0%",   "40% = flights + hotels to attend events."]]
    .forEach(function(row){
      var rowNum = parseInt(row[1].replace("B",""));
      lbl(rowNum,1,row[0]);
      inp(sh,rowNum,2,row[2],row[3]);
      note3(rowNum,row[4],2);
      sh.setRowHeight(rowNum,36);
    });

  // ── K — Funding Rounds ─────────────────────────────────────
  // ✏️ CHANGED: Single $7M Seed round, May 2026. Rounds 2 & 3 cleared.
  secHdrStyled(119,"K — Funding Rounds (up to 3) & Interest Rate");
  ["Round Name","Amount Raised ($)","Close Date","Notes"].forEach(function(h,i){hdr(sh,120,i+1,h,"#1F618D");});
  sh.setRowHeight(120,22);
  [  ["Seed",  7000000, new Date("2026-05-01"), ""],
   ["",      "",      "",                    ""],
   ["",      "",      "",                    ""]]
    .forEach(function(round,i){
      var r=121+i;
      inp(sh,r,1,round[0]);
      if(round[1]){
        inp(sh,r,2,round[1],"$#,##0");
        inp(sh,r,3,round[2],"MMM YYYY");
        sh.getRange(r,3).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
      } else {
        [2,3].forEach(function(c){sh.getRange(r,c).setBackground("#EBF5FB");});
      }
      inp(sh,r,4,round[3]);
      sh.setRowHeight(r,22);
    });

  sh.getRange(124,1,1,5).merge()
    .setValue("Interest Income on Cash Balance")
    .setBackground("#AED6F1").setFontWeight("bold").setFontColor("#1A5276").setFontSize(10)
    .setBorder(true,false,false,false,false,false,"#1A5276",SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(124,26);

  lbl(125,1,"Annual interest rate on cash (%)"); inp(sh,125,2,0.045,"0.0%");
  note3(125,"Applied to prior month ending cash balance. 4–5% reflects current HYSA/T-bill rates.",3);
  sh.setRowHeight(125,28);

  lbl(126,1,"Cash on hand at forecast start ($)");
  inp(sh,126,2,0,"$#,##0");
  note3(126,"Existing cash in bank on day 1 of forecast. Add any pre-forecast raise here.",3);
  sh.setRowHeight(126,36);

  sh.getRange(128,1,1,4).merge()
    .setValue("Rounds feed the 💰 Funding tab automatically. All fields per round must be fully filled or completely blank.")
    .setFontStyle("italic").setFontColor("#888").setWrap(true);
  sh.setRowHeight(128,36);

  // ── L — Annual ARR Targets ─────────────────────────────────
  // ✏️ NEW: Y1–Y5 targets. B12 auto-selects the active year based on B14.
  secHdrStyled(130,"L — Annual ARR Targets (calendar year-end)");
  ["Year","Target ARR ($)","Target Date","Active?","Notes"].forEach(function(h,i){
    hdr(sh,131,i+1,h,"#1F618D");
  });
  sh.setRowHeight(131,22);

  var arrTargets = [
    ["Year 1",  4000000,   new Date("2026-12-01"), "CEO guidance: $3–5M by Dec 2026"],
    ["Year 2",  12000000,  new Date("2027-12-01"), "3× Year 1 minimum. CEO guidance: $9–15M by Dec 2027"],
    ["Year 3",  25000000,  new Date("2028-12-01"), "Industry: 2–2.5× Year 2 for high-growth SaaS"],
    ["Year 4",  50000000,  new Date("2029-12-01"), "Path to Series B / growth equity territory"],
    ["Year 5",  100000000, new Date("2030-12-01"), "Rule of 40 + path to profitability expected"]
  ];

  var horizonThresholds = [12, 24, 36, 48, 60];
  arrTargets.forEach(function(t, i) {
    var r = 132 + i;
    sh.getRange(r,1).setValue(t[0]).setFontWeight("bold").setFontColor("#444444");
    inp(sh,r,2,t[1],"$#,##0");
    sh.getRange(r,3).setValue(t[2]).setBackground("#EBF5FB").setFontColor("#1A5276").setNumberFormat("MMM YYYY");
    sh.getRange(r,3).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
    // Active? — highlights whichever year the current horizon falls in
    var loMo = i === 0 ? 1  : horizonThresholds[i-1] + 1;
    var hiMo = horizonThresholds[i];
    var activeFormula = i === 0
      ? "=IF(B14<=" + hiMo + ",\"✅ ACTIVE\",\"-\")"
      : i === 4
      ? "=IF(B14>" + (hiMo - 12) + ",\"✅ ACTIVE\",\"-\")"
      : "=IF(AND(B14>=" + loMo + ",B14<=" + hiMo + "),\"✅ ACTIVE\",\"-\")";
    sh.getRange(r,4).setFormula(activeFormula)
      .setFontColor("#1D9E75").setFontWeight("bold").setHorizontalAlignment("center");
    sh.getRange(r,5).setValue(t[3])
      .setBackground("#FEF9E7").setFontColor("#7D6608").setFontStyle("italic").setWrap(true);
    sh.setRowHeight(r,22);
  });

  // Conditional format — highlight active row green
  var lRules = sh.getConditionalFormatRules();
  for (var i = 0; i < 5; i++) {
    var r = 132 + i;
    var loMo = i === 0 ? 1 : horizonThresholds[i-1] + 1;
    var hiMo = horizonThresholds[i];
    var cfFormula = i === 0
      ? "=$B$14<=" + hiMo
      : i === 4
      ? "=$B$14>" + (hiMo - 12)
      : "=AND($B$14>=" + loMo + ",$B$14<=" + hiMo + ")";
    lRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(cfFormula)
      .setBackground("#D5F5E3").setFontColor("#1D9E75")
      .setRanges([sh.getRange(r, 1, 1, 5)])
      .build());
  }
  sh.setConditionalFormatRules(lRules);

  sh.getRange(138,1,1,5).merge()
    .setValue("B12 auto-selects the active year target based on B14 (horizon). Horizon 1–12 → Y1, 13–24 → Y2, 25–36 → Y3, 37–48 → Y4, 49–60 → Y5. Only edit blue cells.")
    .setFontStyle("italic").setFontColor("#888").setWrap(true);
  sh.setRowHeight(138,44);
}
