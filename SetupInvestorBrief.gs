// SetupInvestorBrief.gs — external-facing “how to read” tab. Called from SetupMain.gs.

function setupInvestorBrief(ss) {
  var sh = ss.getSheetByName(SHEET_INVESTOR_BRIEF);
  sh.setColumnWidth(1, 30);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 620);
  sh.setColumnWidth(4, 30);

  function title(row, text) {
    sh.getRange(row, 1, 1, 4).merge()
      .setValue(text)
      .setBackground("#1A5276")
      .setFontColor("#FFFFFF")
      .setFontSize(15)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sh.setRowHeight(row, 40);
  }
  function sec(row, text) {
    sh.getRange(row, 2, 1, 2).merge()
      .setValue(text)
      .setBackground("#D6EAF8")
      .setFontWeight("bold")
      .setFontSize(11);
    sh.setRowHeight(row, 24);
  }
  function irow(r, lbl, content, bgL, bgC) {
    sh.getRange(r, 2).setValue(lbl).setBackground(bgL || "#F2F3F4").setFontWeight("bold").setVerticalAlignment("top").setWrap(true);
    sh.getRange(r, 3).setValue(content).setBackground(bgC || "#FFFFFF").setVerticalAlignment("top").setWrap(true);
    sh.setRowHeight(r, 56);
  }
  function note(r, text) {
    sh.getRange(r, 2, 1, 2).merge()
      .setValue(text)
      .setFontStyle("italic")
      .setFontColor("#717D7E")
      .setBackground("#FDFEFE")
      .setWrap(true);
    sh.setRowHeight(r, 36);
  }
  function blank(r) {
    sh.setRowHeight(r, 12);
  }

  var r = 1;
  title(r, "Deskree — How to read this forecast");
  r++;
  blank(r);
  r++;

  sec(r, "What you’re looking at");
  r++;
  sh.getRange(r, 2, 1, 2).merge()
    .setValue(
      "This workbook shows Deskree’s financial forecast at a summary level. The tabs left visible here focus on outcomes: growth, profitability, and cash — not line-by-line operating assumptions.\n\n" +
        "Internal detail (assumptions, headcount mechanics, funding inputs, and internal benchmark checks) is hidden in the investor-friendly view so the story stays on the metrics that matter for the conversation."
    )
    .setBackground("#EBF5FB")
    .setWrap(true)
    .setVerticalAlignment("top");
  sh.setRowHeight(r, 110);
  r++;
  blank(r);
  r++;

  sec(r, "Suggested reading order");
  r++;
  irow(r, "📊 Summary", "Start here: headline KPIs and year-over-year context.", "#D6EAF8", "#EBF5FB");
  r++;
  irow(r, "📈 Revenue", "ARR path, segments, and revenue bridge — how the top line is expected to develop.");
  r++;
  irow(r, "💸 P&L", "Income statement view: revenue through EBITDA and key ratios.");
  r++;
  irow(r, "🏦 Cash Flow", "Liquidity, burn, runway, and cash impact of operations and financing.");
  r++;
  blank(r);
  r++;

  sec(r, "How this forecast is produced");
  r++;
  sh.getRange(r, 2, 1, 2).merge()
    .setValue(
      "The numbers are generated with Deskree’s internal financial planning model. The model ties together revenue, cost structure, and cash using a consistent set of management assumptions and segment drivers. " +
        "Before we share externally, we sanity-check outputs against our internal benchmarks and widely used SaaS industry reference ranges (where applicable) so the forecast is both internally coherent and directionally comparable to peers.\n\n" +
        "This is a planning and decision-support tool — not a promise of future results."
    )
    .setBackground("#E8F8F5")
    .setWrap(true)
    .setVerticalAlignment("top");
  sh.setRowHeight(r, 120);
  r++;
  blank(r);
  r++;

  note(
    r,
    'Questions on methodology or scenarios: contact the Deskree team.'
  );
}
