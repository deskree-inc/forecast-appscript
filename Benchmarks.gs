// ============================================================
// TETRIX BENCHMARKS — Reality Check (expanded)
// Menu: APP_MENU_LABEL → Check Benchmarks (see ModelConstants.gs)
// Depends on globals from ModelConstants.gs: DR, REV, REVCOLS, PNL, CF
// ============================================================

function runBenchmarks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var drv = ss.getSheetByName("🎛️ Drivers");
  var bm = ss.getSheetByName("🚦 Benchmarks");
  if (!drv || !bm) {
    SpreadsheetApp.getUi().alert("❌ Run setupFinancialModel() first.");
    return;
  }

  var rev = ss.getSheetByName("📈 Revenue");
  var pl = ss.getSheetByName("💸 P&L");
  var cf = ss.getSheetByName("🏦 Cash Flow");
  var hc = ss.getSheetByName("👥 Headcount");

  var mmRow = DR.MM_ROW;
  var entRow = DR.ENT_ROW;

  function d(row, col) {
    var v = drv.getRange(row, col).getValue();
    return v === "" || v === null ? 0 : v;
  }

  function num(v) {
    return typeof v === "number" && !isNaN(v) ? v : 0;
  }

  var mm = {
    begACV: num(d(mmRow, 2)),
    expACV: num(d(mmRow, 3)),
    churn: num(d(mmRow, 4)),
    exp: num(d(mmRow, 5)),
    expMo: num(d(mmRow, 6)),
    cac: num(d(mmRow, 7)),
    leadTime: num(d(mmRow, 8)),
    closeRate: num(d(mmRow, 9))
  };
  var ent = {
    begACV: num(d(entRow, 2)),
    expACV: num(d(entRow, 3)),
    churn: num(d(entRow, 4)),
    exp: num(d(entRow, 5)),
    expMo: num(d(entRow, 6)),
    cac: num(d(entRow, 7)),
    leadTime: num(d(entRow, 8)),
    closeRate: num(d(entRow, 9))
  };

  var targetARR = num(drv.getRange(DR.TARGET_ARR).getValue());
  if (String(drv.getRange(DR.TARGET_ARR).getDisplayValue()).indexOf("⚠") >= 0) {
    targetARR = 0;
  }
  var momGrowth = num(drv.getRange(DR.MOM_GROWTH).getValue());
  var horizon = Math.floor(num(drv.getRange(DR.HORIZON).getValue())) || 24;
  var mmPctArr = num(drv.getRange(DR.MM_PCT_ARR).getValue()) || 0.5;

  var aeRatio = num(drv.getRange(DR.AE_RATIO).getValue()) || 15;
  var fdeRatio = num(drv.getRange(DR.FDE_RATIO).getValue()) || 10;
  var csmRatio = num(drv.getRange(DR.CSM_RATIO).getValue()) || 10;

  var eventsAnn = num(drv.getRange(DR.EVENTS).getValue());
  var digitalAnn = num(drv.getRange(DR.DIGITAL).getValue());

  var forecastStart = drv.getRange(DR.FORECAST_START).getValue();
  var repCap = num(drv.getRange(DR.SALES_REP_CAP).getValue()) || 2;
  var attainment = num(drv.getRange(DR.ATTAINMENT).getValue()) || 0.75;

  function forecastMonthIndex(dt) {
    if (!(dt instanceof Date) || !(forecastStart instanceof Date)) return null;
    return (dt.getFullYear() - forecastStart.getFullYear()) * 12 +
      (dt.getMonth() - forecastStart.getMonth()) + 1;
  }

  var rounds = [];
  var r;
  for (r = 121; r <= 123; r++) {
    var nm = drv.getRange(r, 1).getValue();
    if (nm && String(nm).trim() !== "") {
      rounds.push({
        name: String(nm),
        amount: num(d(r, 2)),
        closeDate: drv.getRange(r, 3).getValue(),
        arrAtClose: d(r, 4)
      });
    }
  }
  var totalFunding = 0;
  for (r = 0; r < rounds.length; r++) totalFunding += rounds[r].amount;
  var firstRoundAmt = rounds.length ? rounds[0].amount : 0;

  function findSeriesRound(rx) {
    var i;
    for (i = 0; i < rounds.length; i++) {
      if (rx.test(rounds[i].name)) return rounds[i];
    }
    return null;
  }
  var seriesA = findSeriesRound(/series\s*a/i);

  function revRow(mo) {
    return REV.DATA_START + mo - 1;
  }
  function plCol(mo) {
    return mo + 1;
  }

  function getRev(mo, col) {
    if (!rev || mo < 1 || mo > 60) return 0;
    return num(rev.getRange(revRow(mo), col).getValue());
  }
  function getPl(mo, row) {
    if (!pl || mo < 1 || mo > 60) return null;
    return pl.getRange(row, plCol(mo)).getValue();
  }
  function getCf(mo, row) {
    if (!cf || mo < 1 || mo > 60) return null;
    return cf.getRange(row, plCol(mo)).getValue();
  }
  function getHc(mo, row) {
    if (!hc || mo < 1 || mo > 60) return 0;
    return num(hc.getRange(row, plCol(mo)).getValue());
  }

  function cacPayback(acv, cac) {
    return acv > 0 && cac > 0 ? cac / (acv / 12) : null;
  }
  function ltvcacFn(acv, churn, cac) {
    return cac > 0 && churn > 0 ? (acv / churn) / cac : null;
  }
  function nrrFn(churn, exp) {
    return 1 + exp - churn;
  }
  function yoyFromMoM(mom) {
    return Math.pow(1 + mom, 12) - 1;
  }
  function blendChurn() {
    return mmPctArr * mm.churn + (1 - mmPctArr) * ent.churn;
  }
  function bessemerTarget(arr) {
    if (arr < 2e6) return 2.0;
    if (arr < 5e6) return 1.4;
    return 1.0;
  }

  var mmPayback = cacPayback(mm.begACV, mm.cac);
  var entPayback = cacPayback(ent.begACV, ent.cac);
  var mmL12 = getRev(12, REVCOLS.MM_LOGOS);
  var entL12 = getRev(12, REVCOLS.ENT_LOGOS);
  var logoSum12 = mmL12 + entL12;
  var blendedPayback = null;
  if (logoSum12 > 0 && mmPayback !== null && entPayback !== null) {
    blendedPayback = (mmL12 * mmPayback + entL12 * entPayback) / logoSum12;
  } else if (mmPayback !== null && entPayback !== null) {
    blendedPayback = (mmPayback + entPayback) / 2;
  }

  var newArr12 = num(getPl(12, PNL.NEW_ARR));
  var expArr12 = num(getPl(12, PNL.EXP_ARR));
  var expPctNewArr = newArr12 > 0 ? expArr12 / newArr12 : null;

  var mmLTV = ltvcacFn(mm.begACV, mm.churn, mm.cac);
  var entLTV = ltvcacFn(ent.begACV, ent.churn, ent.cac);
  var mmNRRv = nrrFn(mm.churn, mm.exp);
  var entNRRv = nrrFn(ent.churn, ent.exp);
  var grr = blendChurn() < 1 ? 1 - blendChurn() : null;

  var entLogoM12 = getHc(12, 18);
  var impliedYoY = yoyFromMoM(momGrowth);
  var bessTarget = bessemerTarget(targetARR);
  var arrFundingRatio = totalFunding > 0 ? targetARR / totalFunding : null;

  var mrr1 = num(getPl(1, PNL.MRR));
  var netBurnM1 = null;
  if (cf) {
    var br1 = getCf(1, CF.BURN_RATE);
    if (typeof br1 === "number" && br1 < 0) netBurnM1 = br1;
  }
  if (netBurnM1 === null) {
    var e1 = num(getPl(1, PNL.EBITDA));
    if (e1 < 0) netBurnM1 = e1;
  }
  var burnMultiple = mrr1 > 0 && netBurnM1 !== null && netBurnM1 < 0
    ? Math.abs(netBurnM1) / mrr1 : null;

  var grossMargin1 = getPl(1, PNL.GROSS_MARGIN);
  grossMargin1 = typeof grossMargin1 === "number" ? grossMargin1 : null;
  var grossMargin12 = getPl(12, PNL.GROSS_MARGIN);
  grossMargin12 = typeof grossMargin12 === "number" ? grossMargin12 : null;

  var rawSAMonth = seriesA && seriesA.closeDate instanceof Date
    ? forecastMonthIndex(seriesA.closeDate) : null;
  var saMonth = rawSAMonth !== null && rawSAMonth >= 1 && rawSAMonth <= horizon
    ? rawSAMonth : null;
  var arrAtSAModel = saMonth && pl ? num(getPl(saMonth, PNL.ARR)) : null;

  var sbMonth = saMonth && saMonth + 21 <= horizon ? saMonth + 21 : null;
  var arrAtSBModel = sbMonth && pl ? num(getPl(sbMonth, PNL.ARR)) : null;

  function avgLogos(m0, m1) {
    var s = 0;
    var n = 0;
    var mx;
    for (mx = m0; mx <= m1; mx++) {
      s += getRev(mx, REVCOLS.MM_LOGOS) + getRev(mx, REVCOLS.ENT_LOGOS);
      n++;
    }
    return n ? s / n : 0;
  }
  var q1 = avgLogos(1, 3);
  var q2 = avgLogos(4, 6);
  var q3 = avgLogos(7, 9);
  var q4 = avgLogos(10, 12);

  var cash12 = cf ? num(getCf(12, CF.END_CASH)) : null;
  var cash18 = cf ? num(getCf(18, CF.END_CASH)) : null;
  var cashSA = saMonth && cf && saMonth <= horizon
    ? num(getCf(saMonth, CF.END_CASH)) : null;

  function avgBurn(m0, m1) {
    var s = 0;
    var n = 0;
    var mb;
    for (mb = m0; mb <= m1; mb++) {
      var b = getCf(mb, CF.BURN_RATE);
      if (typeof b === "number" && b < 0) {
        s += Math.abs(b);
        n++;
      }
    }
    return n ? s / n : 0;
  }

  var runwayAtSA = null;
  if (cashSA !== null && saMonth && saMonth <= horizon) {
    var abSA = avgBurn(Math.max(1, saMonth - 2), saMonth);
    runwayAtSA = abSA > 0 ? cashSA / abSA : null;
  }

  var runwayNext = null;
  if (horizon >= 18 && cash18 !== null) {
    var ab18 = avgBurn(13, 18);
    runwayNext = ab18 > 0 ? cash18 / ab18 : null;
  }

  var cashCushion12 = totalFunding > 0 && cash12 !== null ? cash12 / totalFunding : null;

  var spendMo6 = null;
  if (pl) {
    var sp = 0;
    var moSpend;
    for (moSpend = 1; moSpend <= 6; moSpend++) {
      var col = plCol(moSpend);
      sp += num(pl.getRange(PNL.INFRA, col).getValue());
      sp += num(pl.getRange(PNL.CS_PAYROLL, col).getValue());
      sp += num(pl.getRange(PNL.ENG_PAYROLL, col).getValue());
      sp += num(pl.getRange(PNL.SALES_PAYROLL, col).getValue());
      sp += num(pl.getRange(PNL.MKTG_PAYROLL, col).getValue());
      sp += num(pl.getRange(PNL.GA_PAYROLL, col).getValue());
      sp += num(pl.getRange(PNL.MARKETING, col).getValue());
    }
    spendMo6 = sp;
  }
  var capitalDeployPct = firstRoundAmt > 0 && spendMo6 !== null ? spendMo6 / firstRoundAmt : null;

  var arr12 = pl ? num(getPl(12, PNL.ARR)) : 0;
  var arr6 = pl ? num(getPl(6, PNL.ARR)) : 0;
  var smSum6 = 0;
  var smSum12 = 0;
  if (pl) {
    for (r = 1; r <= 6; r++) {
      smSum6 += num(pl.getRange(PNL.SM_SUBTOTAL, plCol(r)).getValue());
    }
    for (r = 1; r <= 12; r++) {
      smSum12 += num(pl.getRange(PNL.SM_SUBTOTAL, plCol(r)).getValue());
    }
  }
  var magicNum = smSum6 > 0 ? (arr12 - arr6) / (smSum6 * 4) : null;
  var smPctRev = arr12 > 0 ? smSum12 / arr12 : null;

  var mktTotal = eventsAnn + digitalAnn;
  var eventsPctMkt = mktTotal > 0 ? eventsAnn / mktTotal : null;

  var avgLogo712 = avgLogos(7, 12);
  var salesHC12 = getHc(12, 5);
  var capSales = salesHC12 * repCap * attainment;
  var pipelineRatio = capSales > 0 ? avgLogo712 / capSales : null;

  var firstEntLogoMo = null;
  var mi;
  for (mi = 1; mi <= 12; mi++) {
    if (getRev(mi, REVCOLS.ENT_LOGOS) > 0) {
      firstEntLogoMo = mi;
      break;
    }
  }

  var avgACV = mm.begACV > 0 && ent.begACV > 0
    ? (mm.begACV + ent.begACV) / 2 : Math.max(mm.begACV, ent.begACV);
  var estAccounts = avgACV > 0 ? Math.round(targetARR / avgACV) : 0;

  var fdeHcFromTitles = 0;
  for (r = 55; r <= 64; r++) {
    var title = drv.getRange(r, 1).getValue();
    var dept = drv.getRange(r, 2).getValue();
    var s = (" " + String(title) + " " + String(dept) + " ").toLowerCase();
    if (s.indexOf("fde") >= 0 || s.indexOf("deploy") >= 0) fdeHcFromTitles++;
  }
  var csHc12 = getHc(12, 7);
  var fdeHcEff = fdeHcFromTitles > 0 ? fdeHcFromTitles : csHc12;

  var engHC12 = getHc(12, 3);
  var totalHC1 = getHc(1, 14);
  var totalHC12 = getHc(12, 14);
  var engPct = totalHC12 > 0 ? engHC12 / totalHC12 : null;

  var arr1 = pl ? num(getPl(1, PNL.ARR)) : 0;
  var arrPerEmp12 = pl ? num(getPl(12, PNL.ARR_PER_EMP)) : 0;
  if (!arrPerEmp12 && totalHC12 > 0 && arr12 > 0) arrPerEmp12 = arr12 / totalHC12;

  var hcGrowth = totalHC1 > 0 ? totalHC12 / totalHC1 - 1 : null;
  var arrGrowth = arr1 > 0 ? arr12 / arr1 - 1 : null;

  var cogsPct1 = null;
  var cogsPct12 = null;
  if (pl && num(getPl(1, PNL.MRR)) > 0) {
    cogsPct1 = num(getPl(1, PNL.TOTAL_COGS)) / num(getPl(1, PNL.MRR));
  }
  if (pl && num(getPl(12, PNL.MRR)) > 0) {
    cogsPct12 = num(getPl(12, PNL.TOTAL_COGS)) / num(getPl(12, PNL.MRR));
  }

  var gaPct12 = null;
  if (pl && num(getPl(12, PNL.MRR)) > 0) {
    gaPct12 = num(getPl(12, PNL.GA_SUBTOTAL)) / num(getPl(12, PNL.MRR));
  }

  function ebitdaMg(mo) {
    var v = getPl(mo, PNL.EBITDA_MARGIN);
    return typeof v === "number" ? v : null;
  }
  var eb6 = horizon >= 6 ? ebitdaMg(6) : null;
  var eb12 = horizon >= 12 ? ebitdaMg(12) : null;
  var eb24 = horizon >= 24 ? ebitdaMg(24) : null;

  function R(value, status, notes) {
    return { value: value, status: status, notes: notes };
  }

  function flagPaybackMM(mo) {
    var note = "Best-in-class MM SaaS recovers CAC in < 12 months. > 18 months will concern Series A investors.";
    if (mo === null) return R("N/A", "—", note);
    var val = mo.toFixed(1) + " mo";
    if (mo < 12) return R(val, "🟢", note);
    if (mo <= 18) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagPaybackENT(mo) {
    var note = "Enterprise CAC payback up to 18 months is acceptable given ACV size. > 24 months requires strong NRR story.";
    if (mo === null) return R("N/A", "—", note);
    var val = mo.toFixed(1) + " mo";
    if (mo < 18) return R(val, "🟢", note);
    if (mo <= 24) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagLTVMM(ratio) {
    var note = "SaaS benchmark is 3x minimum. Top-decile MM companies show 5–7x. Below 3x means you're likely over-spending on acquisition.";
    if (ratio === null) return R("N/A", "—", note);
    var val = ratio.toFixed(1) + "x";
    if (ratio >= 5) return R(val, "🟢", note);
    if (ratio >= 3) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagLTVENT(ratio) {
    var note = "Enterprise LTV is high due to low churn, so ratio should be proportionally better. < 4x suggests CAC is too high relative to ACV.";
    if (ratio === null) return R("N/A", "—", note);
    var val = ratio.toFixed(1) + "x";
    if (ratio >= 7) return R(val, "🟢", note);
    if (ratio >= 4) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagBlendedPayback(mo) {
    var note = "Blended view investors use when evaluating mixed segment models. Weighted by expected logo volume.";
    if (mo === null) return R("N/A", "—", note);
    var val = mo.toFixed(1) + " mo";
    if (mo < 15) return R(val, "🟢", note);
    if (mo <= 20) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagExpPctNewArr(pct) {
    var note = "Best-in-class SaaS gets 30–50% of new ARR from expansion. Low expansion % means you're entirely dependent on new logos.";
    if (pct === null || !pl) return R("Not wired", "—", note);
    var val = (pct * 100).toFixed(0) + "%";
    if (pct >= 0.30) return R(val, "🟢", note);
    if (pct >= 0.15) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagNRRMM(rate) {
    var note = "120%+ NRR means your existing base grows without any new logos. Series A investors will ask for this number first.";
    var val = (rate * 100).toFixed(0) + "%";
    if (rate >= 1.20) return R(val, "🟢", note);
    if (rate >= 1.10) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagNRRENT(rate) {
    var note = "Enterprise benchmark is 130%+. Top-decile (Snowflake, Datadog) show 150%+. Below 115% is a red flag at this ACV level.";
    var val = (rate * 100).toFixed(0) + "%";
    if (rate >= 1.30) return R(val, "🟢", note);
    if (rate >= 1.15) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagGRR(rate) {
    var note = "GRR strips out expansion to show raw retention health. Below 85% means churn is structurally eating your base.";
    if (rate === null) return R("N/A", "—", note);
    var val = (rate * 100).toFixed(0) + "%";
    if (rate >= 0.90) return R(val, "🟢", note);
    if (rate >= 0.85) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagEntConcentration(logos) {
    var note = "If 1–2 enterprise customers represent > 40% of ARR, investors will flag key account risk. Target 5+ logos before Series A.";
    var val = String(Math.round(logos)) + " ENT logos (Mo 12)";
    if (logos > 5) return R(val, "🟢", note);
    if (logos >= 3) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagExpansionTiming() {
    var note = "Expansion timing drives NRR. If first expansion takes > 15 months for MM, your land-and-expand motion is too slow.";
    function mmTier() {
      if (mm.expMo <= 12) return 0;
      if (mm.expMo <= 15) return 1;
      return 2;
    }
    function entTier() {
      if (ent.expMo <= 18) return 0;
      if (ent.expMo <= 24) return 1;
      return 2;
    }
    var worst = Math.max(mmTier(), entTier());
    var val = "MM mo " + mm.expMo + " · ENT mo " + ent.expMo;
    if (worst === 0) return R(val, "🟢", note);
    if (worst === 1) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagImpliedYoY() {
    var note = "Bessemer growth endurance heuristic: at < $2M ARR expect 200% YoY; $2–5M expect 140%; $5M+ expect 100%.";
    var yoyPct = (impliedYoY * 100).toFixed(0) + "% YoY";
    if (impliedYoY >= bessTarget) return R(yoyPct, "🟢", note);
    if (impliedYoY >= bessTarget * 0.75) return R(yoyPct, "🟡", note);
    return R(yoyPct, "🔴", note);
  }

  function flagBessemer() {
    var note = "Each year growth should decay ~30%. If you're projecting flat growth rates year-over-year, the model is likely too optimistic.";
    var val = "Target ~" + (bessTarget * 100).toFixed(0) + "% YoY at this ARR stage";
    return R(val, "ℹ️", note);
  }

  function flagARRSeriesA() {
    var note = "2026 median Series A requires $1–2M ARR with 2–3x YoY. Top-decile shows $3M+ ARR. Adjust Series A close date or growth assumptions.";
    if (!seriesA) return R("No Series A round", "—", "Name a round \"Series A\" in Drivers Section K with close date within horizon.");
    if (saMonth === null || arrAtSAModel === null) {
      return R("Close outside horizon", "—", "Series A close must fall between Mo 1 and your forecast horizon (B14).");
    }
    var val = "$" + (arrAtSAModel / 1e6).toFixed(2) + "M ARR";
    if (arrAtSAModel >= 2e6) return R(val, "🟢", note);
    if (arrAtSAModel >= 1e6) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagARRSeriesB() {
    var note = "Series B threshold has risen to $4–8M ARR post-2021. You have a 40–50% graduation rate from A to B within 4 years industry-wide.";
    if (!seriesA || sbMonth === null) return R("N/A", "—", "Needs Series A date + horizon through ~21 mo post-close.");
    if (!pl) return R("Not wired", "—", note);
    var val = "$" + (arrAtSBModel / 1e6).toFixed(2) + "M ARR @ Mo " + sbMonth;
    if (arrAtSBModel >= 8e6) return R(val, "🟢", note);
    if (arrAtSBModel >= 4e6) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagLogoConsistency() {
    var note = "Investors want to see compounding growth, not a plateau. If new logos per month flatten after Mo 12, your sales capacity model needs revision.";
    var val = "Qavg " + q1.toFixed(1) + " → " + q2.toFixed(1) + " → " + q3.toFixed(1) + " → " + q4.toFixed(1);
    if (q2 < q1 * 0.92 && q3 < q2 * 0.92) return R(val, "🔴", note);
    if (q4 < q2 * 0.95 || (q3 < q2 * 0.95 && q4 <= q3 * 1.02)) return R(val, "🟡", note);
    return R(val, "🟢", note);
  }

  function flagBurnMult() {
    var note = "Bessemer burn multiple benchmark: < 1x is exceptional, 1–1.5x good, 1.5–2x acceptable for early stage. > 2.5x will be flagged at Series A diligence.";
    if (burnMultiple === null) return R("Not wired", "—", note);
    var val = burnMultiple.toFixed(2) + "x";
    if (burnMultiple < 1.5) return R(val, "🟢", note);
    if (burnMultiple <= 2.5) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagRunwaySA() {
    var note = "You should close your next round with at least 6 months of runway from the prior round remaining. < 3 months means you're fundraising from desperation, not strength.";
    if (runwayAtSA === null) return R("N/A", "—", "Set Series A in Drivers K with date in forecast.");
    var val = runwayAtSA.toFixed(1) + " mo";
    if (runwayAtSA >= 6) return R(val, "🟢", note);
    if (runwayAtSA >= 3) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagRunwayNext() {
    var note = "Rule of thumb: start fundraising when you have 9–12 months of runway. < 6 months entering a raise puts you in a weak negotiating position.";
    if (runwayNext === null) return R("N/A", "—", "Need horizon ≥ 18 mo and cash flow wired.");
    var val = runwayNext.toFixed(1) + " mo";
    if (runwayNext >= 12) return R(val, "🟢", note);
    if (runwayNext >= 6) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagCashCushion() {
    var note = "At Month 12 you should have spent no more than 60% of your raise. Burning > 75% by Mo 12 of a 24-month model leaves no buffer for delays.";
    if (cashCushion12 === null) return R("N/A", "—", note);
    var val = (cashCushion12 * 100).toFixed(0) + "% of raised";
    if (cashCushion12 >= 0.40) return R(val, "🟢", note);
    if (cashCushion12 >= 0.25) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagFundingARR() {
    var note = "A healthy SaaS company generates $0.40+ of ARR per $1 raised over its lifetime. Below $0.25 signals capital inefficiency that will concern Series B investors.";
    if (arrFundingRatio === null) return R("N/A", "—", note);
    var val = (arrFundingRatio * 100).toFixed(0) + "%";
    if (arrFundingRatio >= 0.40) return R(val, "🟢", note);
    if (arrFundingRatio >= 0.25) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagCapitalDeploy() {
    var note = "If you're burning > 40% of a raise in the first 6 months, you're front-loading cost without enough time to prove the metrics that justify the next round.";
    if (capitalDeployPct === null) return R("N/A", "—", "Fill first funding round in Section K.");
    var val = (capitalDeployPct * 100).toFixed(0) + "% of Round 1";
    if (capitalDeployPct <= 0.30) return R(val, "🟢", note);
    if (capitalDeployPct <= 0.40) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagMagic() {
    var note = "Magic Number > 0.75 means every $1 of S&M spend generates $0.75 of net new ARR. < 0.5 means your go-to-market is inefficient. Best-in-class shows > 1.0.";
    if (magicNum === null) return R("Not wired", "—", note);
    var val = magicNum.toFixed(2);
    if (magicNum >= 0.75) return R(val, "🟢", note);
    if (magicNum >= 0.5) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagSMPct() {
    var note = "Top-quartile SaaS spends 35–45% of revenue on S&M at the growth stage. > 60% is sustainable only if NRR is strong enough to offset it.";
    if (smPctRev === null) return R("Not wired", "—", note);
    var val = (smPctRev * 100).toFixed(0) + "%";
    if (smPctRev < 0.40) return R(val, "🟢", note);
    if (smPctRev <= 0.60) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagEventsMkt() {
    var note = "Events are high-ROI for enterprise pipeline but shouldn't dominate. > 75% means you have no scalable digital channel. < 15% means you're underinvesting in a proven channel for your ICP.";
    if (eventsPctMkt === null) return R("N/A", "—", note);
    var val = (eventsPctMkt * 100).toFixed(0) + "% events";
    if (eventsPctMkt >= 0.30 && eventsPctMkt <= 0.60) return R(val, "🟢", note);
    if (eventsPctMkt > 0.75 || eventsPctMkt < 0.15) return R(val, "🔴", note);
    return R(val, "🟡", note);
  }

  function flagPipeline() {
    var note = "If your logo ramp requires more closes than your team's close rate × headcount can support, the revenue forecast is structurally unreachable.";
    if (pipelineRatio === null || salesHC12 <= 0) return R("N/A", "—", "Need sales HC and logo ramp in model.");
    var val = pipelineRatio.toFixed(2) + " (logos / capacity)";
    if (pipelineRatio <= 0.8) return R(val, "🟢", note);
    if (pipelineRatio <= 1.0) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagLeadTimeVsRamp() {
    var note = "If your ramp shows enterprise logos in Mo 1–3 but lead time is 6 months, the model is physically impossible. First close can't happen before the sales cycle ends.";
    if (firstEntLogoMo === null) return R("No ENT logos in Y1", "—", note);
    var entStartLag = Math.ceil(num(drv.getRange(44, 4).getValue()));
    var needMo = entStartLag + Math.ceil(ent.leadTime);
    var val = "1st ENT logo Mo " + firstEntLogoMo + " vs need ≥ Mo " + needMo;
    if (firstEntLogoMo >= needMo) return R(val, "🟢", note);
    return R(val, "🔴", note);
  }

  function flagAELoad() {
    var note = "If each AE is carrying more accounts than your ratio allows, either revenue will suffer or you'll need unplanned hires. Adjust headcount or ARR target.";
    if (estAccounts <= 0 || salesHC12 <= 0) return R("N/A", "—", "Set target ARR, ACVs, and sales HC.");
    var load = estAccounts / salesHC12;
    var val = load.toFixed(1) + " acct/AE (max " + aeRatio + ")";
    if (load <= aeRatio) return R(val, "🟢", note);
    if (load <= aeRatio * 1.2) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagFDELoad() {
    var note = "FDEs are a constraint on your ability to deploy and retain enterprise customers. Understaffing here drives churn, not a lagging metric to fix later.";
    if (estAccounts <= 0) return R("N/A", "—", note);
    if (fdeHcEff <= 0 && estAccounts > 5) {
      return R("No FDE HC", "🔴", note);
    }
    if (fdeHcEff <= 0) return R("No FDE HC", "—", note);
    var load = estAccounts / fdeHcEff;
    var val = load.toFixed(1) + " acct/FDE (max " + fdeRatio + ")";
    if (load <= fdeRatio) return R(val, "🟢", note);
    if (load <= fdeRatio * 1.3) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagCSMLoad() {
    var note = "CSM coverage directly drives NRR. If CSMs are overloaded, expansion slows and churn rises — both will show up in your NRR benchmark.";
    if (estAccounts <= 0) return R("N/A", "—", note);
    if (csHc12 <= 0) return R("No CS HC", "🔴", note);
    var load = estAccounts / csHc12;
    var val = load.toFixed(1) + " acct/CSM (max " + csmRatio + ")";
    if (load <= csmRatio) return R(val, "🟢", note);
    if (load <= csmRatio * 1.15) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagEngPct() {
    var note = "For an AI infrastructure product, engineering should represent 40–60% of headcount at seed/Series A. < 30% suggests under-investment in product; > 70% suggests under-investment in GTM.";
    if (engPct === null) return R("N/A", "—", note);
    var val = (engPct * 100).toFixed(0) + "%";
    if (engPct >= 0.40 && engPct <= 0.60) return R(val, "🟢", note);
    if ((engPct >= 0.30 && engPct < 0.40) || (engPct > 0.60 && engPct <= 0.70)) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagRevPerEmp() {
    var note = "Top-decile Series A SaaS shows $200K+ ARR per employee. $150K is the floor for a credible pitch. Below $100K suggests you're hiring ahead of revenue.";
    if (!arrPerEmp12) return R("N/A", "—", note);
    var val = "$" + (arrPerEmp12 / 1000).toFixed(0) + "K";
    if (arrPerEmp12 >= 150000) return R(val, "🟢", note);
    if (arrPerEmp12 >= 100000) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagHiringVsARR() {
    var note = "Headcount should grow slower than ARR — that's how you get operating leverage. If you're hiring faster than you're growing revenue, margins will compress.";
    if (hcGrowth === null || arrGrowth === null) return R("N/A", "—", note);
    var val = "HC " + (hcGrowth * 100).toFixed(0) + "% vs ARR " + (arrGrowth * 100).toFixed(0) + "%";
    if (arrGrowth >= hcGrowth) return R(val, "🟢", note);
    if (hcGrowth <= arrGrowth * 1.2) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagGrossMargin() {
    var note = "SaaS benchmark is 70–80%. AI infrastructure products often start lower (60–70%) due to GPU/infra costs but should trend toward 75%+ as you optimize.";
    if (grossMargin1 === null) return R("Not wired", "—", note);
    var val = (grossMargin1 * 100).toFixed(0) + "% (Mo 1)";
    if (grossMargin1 >= 0.75) return R(val, "🟢", note);
    if (grossMargin1 >= 0.65) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagCOGSTrend() {
    var note = "COGS% should improve as infra costs spread over more customers. If it's rising, you have a unit economics problem that compounds at scale.";
    if (cogsPct1 === null || cogsPct12 === null) return R("Not wired", "—", note);
    var val = (cogsPct1 * 100).toFixed(0) + "% → " + (cogsPct12 * 100).toFixed(0) + "% of MRR";
    if (cogsPct12 < cogsPct1 - 0.005) return R(val, "🟢", note);
    if (cogsPct12 <= cogsPct1 + 0.01) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagGAPct() {
    var note = "G&A should be a shrinking % of revenue as you scale. > 30% at Mo 12 suggests operational overhead is too high relative to revenue.";
    if (gaPct12 === null) return R("Not wired", "—", note);
    var val = (gaPct12 * 100).toFixed(0) + "% of MRR";
    if (gaPct12 < 0.20) return R(val, "🟢", note);
    if (gaPct12 <= 0.30) return R(val, "🟡", note);
    return R(val, "🔴", note);
  }

  function flagEBITDATraj() {
    var note = "Investors don't expect profitability at seed/Series A, but they do expect a credible path. Worsening EBITDA margins over 24 months with increasing ARR is a red flag.";
    if (eb12 === null) return R("Not wired", "—", note);
    if (eb24 !== null) {
      var val3 = "Mo 6→12→24: " + (eb6 !== null ? (eb6 * 100).toFixed(0) : "?") + "% / " +
        (eb12 * 100).toFixed(0) + "% / " + (eb24 * 100).toFixed(0) + "%";
      if (eb24 > eb12 + 0.005 && eb12 > eb6 + 0.005) return R(val3, "🟢", note);
      if (eb24 < eb12 - 0.01 || eb12 < eb6 - 0.01) return R(val3, "🔴", note);
      return R(val3, "🟡", note);
    }
    var val2 = "Mo 6→12: " + (eb6 !== null ? (eb6 * 100).toFixed(0) : "?") + "% / " + (eb12 * 100).toFixed(0) + "%";
    if (eb6 !== null && eb12 > eb6 + 0.005) return R(val2, "🟢", note);
    if (eb6 !== null && eb12 < eb6 - 0.01) return R(val2, "🔴", note);
    return R(val2, "🟡", note);
  }

  var rows = [];

  function cat(title) {
    rows.push({ section: title });
  }

  function add(label, fn) {
    rows.push({ label: label, result: fn() });
  }

  cat("1 — Unit economics");
  add("CAC Payback — Mid-Market", function () { return flagPaybackMM(mmPayback); });
  add("CAC Payback — Enterprise", function () { return flagPaybackENT(entPayback); });
  add("LTV:CAC — Mid-Market", function () { return flagLTVMM(mmLTV); });
  add("LTV:CAC — Enterprise", function () { return flagLTVENT(entLTV); });
  add("Blended CAC Payback", function () { return flagBlendedPayback(blendedPayback); });
  add("Expansion Revenue as % of New ARR (Mo 12)", function () { return flagExpPctNewArr(expPctNewArr); });

  cat("2 — Revenue quality");
  add("Net Revenue Retention — Mid-Market", function () { return flagNRRMM(mmNRRv); });
  add("Net Revenue Retention — Enterprise", function () { return flagNRRENT(entNRRv); });
  add("Gross Revenue Retention (blended churn)", function () { return flagGRR(grr); });
  add("ARR Concentration Risk (ENT logos Mo 12)", function () { return flagEntConcentration(entLogoM12); });
  add("Time to First Expansion (Drivers)", function () { return flagExpansionTiming(); });

  cat("3 — Growth");
  add("Implied YoY Growth Rate", function () { return flagImpliedYoY(); });
  add("Bessemer Growth Endurance Check", function () { return flagBessemer(); });
  add("ARR at Series A Threshold", function () { return flagARRSeriesA(); });
  add("ARR at Series B Threshold (~21 mo post-A)", function () { return flagARRSeriesB(); });
  add("MoM Growth Consistency (logo quarters)", function () { return flagLogoConsistency(); });

  cat("4 — Burn & capital efficiency");
  add("Burn Multiple (|net burn Mo 1| / MRR Mo 1)", function () { return flagBurnMult(); });
  add("Runway at Series A Close", function () { return flagRunwaySA(); });
  add("Runway Before Next Round (Mo 18 / burn 13–18)", function () { return flagRunwayNext(); });
  add("Cash Cushion at Month 12", function () { return flagCashCushion(); });
  add("Funding → ARR Coverage Ratio", function () { return flagFundingARR(); });
  add("Capital Deployed vs Plan (Mo 6 vs Round 1)", function () { return flagCapitalDeploy(); });

  cat("5 — Sales & marketing efficiency");
  add("Magic Number", function () { return flagMagic(); });
  add("Sales & Marketing as % of Revenue (Y1)", function () { return flagSMPct(); });
  add("Events Budget as % of Total Marketing", function () { return flagEventsMkt(); });
  add("Pipeline Coverage (logos vs rep capacity)", function () { return flagPipeline(); });
  add("Selling Cycle vs Lead Time (first ENT logo)", function () { return flagLeadTimeVsRamp(); });

  cat("6 — Headcount & org health");
  add("AE Account Load at Target ARR", function () { return flagAELoad(); });
  add("FDE Account Load at Target ARR", function () { return flagFDELoad(); });
  add("CSM Account Load at Target ARR", function () { return flagCSMLoad(); });
  add("Engineering % of Total Headcount (Mo 12)", function () { return flagEngPct(); });
  add("Revenue per Employee (ARR / HC Mo 12)", function () { return flagRevPerEmp(); });
  add("Hiring Pace vs ARR Growth", function () { return flagHiringVsARR(); });

  cat("7 — Gross margin & P&L health");
  add("Gross Margin (Mo wired)", function () { return flagGrossMargin(); });
  add("COGS as % of Revenue Trend", function () { return flagCOGSTrend(); });
  add("G&A as % of Revenue (Mo 12)", function () { return flagGAPct(); });
  add("EBITDA Margin Trajectory", function () { return flagEBITDATraj(); });

  var counts = { "🟢": 0, "🟡": 0, "🔴": 0 };
  var i;
  for (i = 0; i < rows.length; i++) {
    if (rows[i].result && counts.hasOwnProperty(rows[i].result.status)) {
      counts[rows[i].result.status]++;
    }
  }

  var clearR = bm.getRange(2, 1, Math.max(bm.getLastRow(), 95), 4);
  try { clearR.breakApart(); } catch (e) { /* ignore */ }
  clearR.clearContent();
  clearR.clearFormat();

  bm.getRange(1, 1, 1, 4).merge()
    .setValue("🚦 BENCHMARKS — Reality Check")
    .setBackground("#922B21")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  bm.getRange(2, 1, 1, 4).merge()
    .setValue(counts["🟢"] + " green · " + counts["🟡"] + " yellow · " + counts["🔴"] + " red")
    .setBackground("#2C3E50")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  bm.getRange(3, 1, 1, 4).merge()
    .setValue("Excludes informational (ℹ️) and not-applicable (—) rows. Run after model recalculates.")
    .setFontStyle("italic")
    .setFontColor("#666666")
    .setFontSize(10);

  var colors = {
    "🟢": "#E6F4EA",
    "🟡": "#FEF9E7",
    "🔴": "#FADBD8",
    "—": "#F2F3F4",
    "ℹ️": "#EBF5FB"
  };

  hdr(bm, 4, 1, "Metric", "#1F618D");
  hdr(bm, 4, 2, "Value", "#1F618D");
  hdr(bm, 4, 3, "Status", "#1F618D");
  hdr(bm, 4, 4, "Notes / Benchmark", "#1F618D");

  var cur = 5;
  for (i = 0; i < rows.length; i++) {
    if (rows[i].section) {
      bm.getRange(cur, 1, 1, 4).merge()
        .setValue(rows[i].section)
        .setBackground("#D6EAF8")
        .setFontWeight("bold")
        .setFontColor("#1A5276")
        .setFontSize(10);
      bm.setRowHeight(cur, 24);
      cur++;
      continue;
    }
    var res = rows[i].result;
    var bg = colors[res.status] || "#FFFFFF";
    bm.getRange(cur, 1).setValue(rows[i].label).setFontWeight("bold").setBackground(bg);
    bm.getRange(cur, 2).setValue(res.value).setBackground(bg).setHorizontalAlignment("right");
    bm.getRange(cur, 3).setValue(res.status).setBackground(bg).setHorizontalAlignment("center").setFontSize(14);
    bm.getRange(cur, 4).setValue(res.notes).setBackground(bg).setFontColor("#555555").setWrap(true);
    bm.setRowHeight(cur, 36);
    cur++;
  }

  bm.getRange(cur + 1, 1, 1, 4).merge()
    .setValue("Last checked: " + new Date().toLocaleString())
    .setFontStyle("italic")
    .setFontColor("#888888");

  bm.setColumnWidth(1, 260);
  bm.setColumnWidth(2, 140);
  bm.setColumnWidth(3, 90);
  bm.setColumnWidth(4, 420);

  SpreadsheetApp.getUi().alert("✅ Benchmark check complete. Review the 🚦 Benchmarks tab.");
}
