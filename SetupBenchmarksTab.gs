// SetupBenchmarksTab.gs — empty Benchmarks sheet shell (not runBenchmarks). Called from SetupMain.gs.

function setupBenchmarks(ss) {
  var sh = ss.getSheetByName("🚦 Benchmarks");
  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 420);
  hdr(sh, 1, 1, "🚦 BENCHMARKS — Reality Check", "#922B21");
  sh.getRange(1, 1, 1, 4).merge();
  sh.getRange(2, 1, 1, 4).merge()
    .setValue("Run " + APP_MENU_LABEL + " → Check Benchmarks to populate categories, counts, and metric rows.")
    .setFontStyle("italic").setFontColor("#888").setWrap(true);
  hdr(sh, 4, 1, "Metric", "#1F618D");
  hdr(sh, 4, 2, "Value", "#1F618D");
  hdr(sh, 4, 3, "Status", "#1F618D");
  hdr(sh, 4, 4, "Notes / Benchmark", "#1F618D");
  sh.getRange(5, 1, 1, 4).merge()
    .setValue("Seven sections (unit economics → P&L). Summary row counts 🟢 / 🟡 / 🔴 after each run.")
    .setBackground("#F8F9FA").setFontColor("#555555").setFontSize(10).setWrap(true);
}