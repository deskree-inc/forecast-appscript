// ============================================================
// TETRIX BENCHMARKS — Reality Check
// Add to menu: .addItem("🚦 Check Benchmarks", "runBenchmarks")
// in the onOpen() function inside ScenarioSidebar.gs
// ============================================================

function runBenchmarks() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const drv = ss.getSheetByName("🎛️ Drivers");
  const bm  = ss.getSheetByName("🚦 Benchmarks");
  if (!drv || !bm) { SpreadsheetApp.getUi().alert("❌ Run setupFinancialModel() first."); return; }

  // ── Read Drivers ──────────────────────────────────────────
  const mmRow = 18, entRow = 19;
  function d(row, col) { return drv.getRange(row, col).getValue() || 0; }

  const mm = {
    begACV: d(mmRow,2), expACV: d(mmRow,3), churn: d(mmRow,4),
    exp: d(mmRow,5), expMo: d(mmRow,6), cac: d(mmRow,7),
    leadTime: d(mmRow,8), closeRate: d(mmRow,9),
  };
  const ent = {
    begACV: d(entRow,2), expACV: d(entRow,3), churn: d(entRow,4),
    exp: d(entRow,5), expMo: d(entRow,6), cac: d(entRow,7),
    leadTime: d(entRow,8), closeRate: d(entRow,9),
  };

  const targetARR   = drv.getRange("B12").getValue() || 0;
  const momGrowth   = drv.getRange("B13").getValue() || 0;
  const aeRatio     = drv.getRange("B28").getValue() || 15;
  const fdeRatio    = drv.getRange("B29").getValue() || 10;
  const totalFunding = [5,6,7,8,9].reduce((sum, r) => sum + (d(r,2)||0), 0);

  // Headcount from Headcount tab (month 1)
  const hcSheet = ss.getSheetByName("👥 Headcount");
  const totalHC_m1 = hcSheet ? (
    hcSheet.getRange(3,2).getValue() + hcSheet.getRange(5,2).getValue() +
    hcSheet.getRange(7,2).getValue() + hcSheet.getRange(9,2).getValue()
  ) : 0;
  const salesHC = hcSheet ? hcSheet.getRange(5,2).getValue() : 0;

  // P&L / CashFlow (if wired)
  const plSheet = ss.getSheetByName("💸 P&L");
  const grossMargin = plSheet ? plSheet.getRange(11,2).getValue() : null;
  const cfSheet = ss.getSheetByName("🏦 Cash Flow");
  const netBurn_m1 = cfSheet ? cfSheet.getRange(10,2).getValue() : null;
  const mrr_m1 = plSheet ? plSheet.getRange(3,2).getValue() : 0;

  // ── Compute Metrics ───────────────────────────────────────

  function cacPayback(acv, cac) { return acv > 0 ? cac / (acv / 12) : null; }
  function ltvcac(acv, churn, cac) { return cac > 0 && churn > 0 ? (acv / churn) / cac : null; }
  function nrr(churn, exp) { return 1 + exp - churn; }
  function yoyFromMoM(mom) { return Math.pow(1 + mom, 12) - 1; }
  // Bessemer growth endurance: at <$2M ARR expect ~200% YoY; $2-5M ~140%; $5M+ ~100%
  function bessemerTarget(arr) {
    if (arr < 2e6)  return 2.0;
    if (arr < 5e6)  return 1.4;
    return 1.0;
  }
  // Estimated accounts at target ARR
  const avgACV = (mm.begACV + ent.begACV) / 2;
  const estAccounts = avgACV > 0 ? Math.round(targetARR / avgACV) : 0;
  const neededAE  = aeRatio  > 0 ? Math.ceil(estAccounts / aeRatio)  : 0;
  const neededFDE = fdeRatio > 0 ? Math.ceil(estAccounts / fdeRatio) : 0;

  const mmPayback  = cacPayback(mm.begACV, mm.cac);
  const entPayback = cacPayback(ent.begACV, ent.cac);
  const mmLTVCAC   = ltvcac(mm.begACV, mm.churn, mm.cac);
  const entLTVCAC  = ltvcac(ent.begACV, ent.churn, ent.cac);
  const mmNRR      = nrr(mm.churn, mm.exp);
  const entNRR     = nrr(ent.churn, ent.exp);
  const impliedYoY = yoyFromMoM(momGrowth);
  const bessTarget = bessemerTarget(targetARR);
  const arrFundingRatio = totalFunding > 0 ? targetARR / totalFunding : null;
  const burnMultiple = (mrr_m1 > 0 && netBurn_m1 !== null && netBurn_m1 < 0)
    ? Math.abs(netBurn_m1) / mrr_m1 : null;

  // ── Flag logic ────────────────────────────────────────────
  // Returns { value, status, notes }
  function flagPayback(mo, shortOk, longWarn) {
    if (mo === null) return { value: "N/A", status: "—", notes: "Set CAC and ACV in Drivers section C" };
    const val = mo.toFixed(1) + " mo";
    if (mo <= shortOk) return { value: val, status: "🟢", notes: `Good — under ${shortOk} months` };
    if (mo <= longWarn) return { value: val, status: "🟡", notes: `Caution — aim for < ${shortOk} months` };
    return { value: val, status: "🔴", notes: `High — investors prefer < ${longWarn} months for SaaS` };
  }
  function flagLTVCAC(ratio, goodThresh, okThresh) {
    if (ratio === null) return { value: "N/A", status: "—", notes: "Set CAC, ACV and churn in Drivers section C" };
    const val = ratio.toFixed(1) + "x";
    if (ratio >= goodThresh) return { value: val, status: "🟢", notes: `Strong — benchmark is ${goodThresh}x+` };
    if (ratio >= okThresh)   return { value: val, status: "🟡", notes: `Caution — aim for ${goodThresh}x` };
    return { value: val, status: "🔴", notes: `Below ${okThresh}x — unit economics under pressure` };
  }
  function flagNRR(rate, seg) {
    const pct = (rate*100).toFixed(0)+"%";
    const isEnt = seg === "ent";
    if (rate >= (isEnt ? 1.30 : 1.20)) return { value: pct, status: "🟢", notes: `Top-decile for ${seg === "ent" ? "Enterprise" : "Mid-Market"}` };
    if (rate >= (isEnt ? 1.15 : 1.10)) return { value: pct, status: "🟡", notes: `Good — target is ${isEnt ? "130" : "120"}%+` };
    return { value: pct, status: "🔴", notes: "Below 110% — expansion/churn needs review" };
  }
  function flagGrowth() {
    const yoy = (impliedYoY*100).toFixed(0)+"%";
    if (impliedYoY >= bessTarget) return { value: yoy+" YoY", status: "🟢", notes: `Meets Bessemer target of ${(bessTarget*100).toFixed(0)}% at this ARR stage` };
    if (impliedYoY >= bessTarget*0.75) return { value: yoy+" YoY", status: "🟡", notes: `Below Bessemer target of ${(bessTarget*100).toFixed(0)}% — monitor` };
    return { value: yoy+" YoY", status: "🔴", notes: `Below Bessemer target of ${(bessTarget*100).toFixed(0)}% for $${(targetARR/1e6).toFixed(1)}M ARR stage` };
  }
  function flagBessemer() {
    // growth endurance: each year growth should decay ~30%
    const note = `At $${(targetARR/1e6).toFixed(1)}M ARR, Bessemer "growth endurance" heuristic: ~${(bessTarget*100).toFixed(0)}% YoY. ` +
      `Expect ~30% decay per year (e.g. 200% → 140% → 98%).`;
    return { value: `Target: ${(bessTarget*100).toFixed(0)}% YoY`, status: "ℹ️", notes: note };
  }
  function flagARRFunding() {
    if (arrFundingRatio === null) return { value: "N/A", status: "—", notes: "Add funding rounds in Drivers section A" };
    const pct = (arrFundingRatio*100).toFixed(0)+"%";
    if (arrFundingRatio >= 0.30) return { value: pct, status: "🟢", notes: "Good capital efficiency — ARR > 30% of total raised" };
    if (arrFundingRatio >= 0.20) return { value: pct, status: "🟡", notes: "Acceptable — aim for ARR > 30% of capital raised" };
    return { value: pct, status: "🔴", notes: "Low — less than $0.20 ARR per $1 raised; investors may flag burn efficiency" };
  }
  function flagCoverage(needed, current, role) {
    if (estAccounts === 0) return { value: "N/A", status: "—", notes: "Set target ARR and ACV to estimate" };
    const actual = current || 0;
    const load = actual > 0 ? (estAccounts / actual).toFixed(1) : "∞";
    if (actual >= needed) return { value: `${actual} ${role} for ~${estAccounts} accts`, status: "🟢", notes: `${load} accts/${role} — within ratio` };
    if (actual >= needed*0.8) return { value: `${actual} ${role} (need ~${needed})`, status: "🟡", notes: `${load} accts/${role} — approaching limit` };
    return { value: `${actual} ${role} (need ~${needed})`, status: "🔴", notes: `${load} accts/${role} — insufficient at target ARR; hire more ${role}s` };
  }
  function flagGrossMargin() {
    if (grossMargin === null || grossMargin === 0) return { value: "Not wired", status: "—", notes: "Wire P&L MRR row to Revenue tab to compute" };
    const pct = (grossMargin*100).toFixed(0)+"%";
    if (grossMargin >= 0.75) return { value: pct, status: "🟢", notes: "Strong SaaS gross margin (benchmark: 70–80%)" };
    if (grossMargin >= 0.65) return { value: pct, status: "🟡", notes: "Acceptable — aim for 75%+" };
    return { value: pct, status: "🔴", notes: "Below 65% — infra or CS costs may be too high" };
  }
  function flagBurnMultiple() {
    if (burnMultiple === null) return { value: "Not wired", status: "—", notes: "Wire Cash Flow and P&L to compute" };
    const val = burnMultiple.toFixed(2)+"x";
    if (burnMultiple <= 1.5) return { value: val, status: "🟢", notes: "Efficient — Series A benchmark is < 1.5x" };
    if (burnMultiple <= 2.5) return { value: val, status: "🟡", notes: "Caution — aim to get below 1.5x" };
    return { value: val, status: "🔴", notes: "High burn multiple — >2.5x will concern investors" };
  }

  // ── Write to Benchmarks tab ───────────────────────────────
  const results = [
    ["CAC Payback — Mid-Market",      flagPayback(mmPayback, 12, 18)],
    ["CAC Payback — Enterprise",      flagPayback(entPayback, 18, 24)],
    ["LTV:CAC — Mid-Market",          flagLTVCAC(mmLTVCAC, 5, 3)],
    ["LTV:CAC — Enterprise",          flagLTVCAC(entLTVCAC, 7, 4)],
    ["Est. NRR — Mid-Market",         flagNRR(mmNRR, "mm")],
    ["Est. NRR — Enterprise",         flagNRR(entNRR, "ent")],
    ["ARR Growth (MoM → YoY)",        flagGrowth()],
    ["Bessemer Growth Endurance",     flagBessemer()],
    ["Funding → ARR Coverage",        flagARRFunding()],
    ["AE Account Load",               flagCoverage(neededAE, salesHC, "AE")],
    ["FDE Account Load",              flagCoverage(neededFDE, 0, "FDE")],
    ["Gross Margin",                  flagGrossMargin()],
    ["Burn Multiple",                 flagBurnMultiple()],
  ];

  const colors = { "🟢": "#E6F4EA", "🟡": "#FEF9E7", "🔴": "#FADBD8", "—": "#F2F3F4", "ℹ️": "#EBF5FB" };

  results.forEach(([metric, result], i) => {
    const r = 5 + i;
    bm.getRange(r,1).setValue(metric).setFontWeight("bold").setBackground(colors[result.status]||"#FFFFFF");
    bm.getRange(r,2).setValue(result.value).setBackground(colors[result.status]||"#FFFFFF").setHorizontalAlignment("right");
    bm.getRange(r,3).setValue(result.status).setBackground(colors[result.status]||"#FFFFFF").setHorizontalAlignment("center").setFontSize(14);
    bm.getRange(r,4).setValue(result.notes).setBackground(colors[result.status]||"#FFFFFF").setFontColor("#555").setWrap(true);
    bm.setRowHeight(r, 36);
  });

  bm.getRange(18,1).setValue(`Last checked: ${new Date().toLocaleString()}`)
    .setFontStyle("italic").setFontColor("#888"); bm.getRange(18,1,1,4).merge();

  SpreadsheetApp.getUi().alert("✅ Benchmark check complete. Review the 🚦 Benchmarks tab.");
}