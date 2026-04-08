// SetupHelpers.gs — layout helpers. Called from setupFinancialModel in SetupMain.gs.

function hdr(sheet, row, col, text, bg) {
  sheet.getRange(row, col).setValue(text).setFontWeight("bold")
    .setBackground(bg || "#2C3E50").setFontColor("#FFFFFF");
}
function inp(sheet, row, col, value, format) {
  var c = sheet.getRange(row, col);
  c.setValue(value).setBackground("#EBF5FB").setFontColor("#1A5276");
  if (format) c.setNumberFormat(format);
}
function label(sheet, row, col, text) {
  sheet.getRange(row, col).setValue(text).setFontWeight("bold");
}
function sectionHdr(sheet, row, text) {
  sheet.getRange(row, 1, 1, 9).merge().setValue(text)
    .setBackground("#D6EAF8").setFontWeight("bold").setFontSize(10);
}
function colLetter(col) {
  var l = "";
  while (col > 0) {
    var r = (col - 1) % 26;
    l = String.fromCharCode(65 + r) + l;
    col = Math.floor((col - 1) / 26);
  }
  return l;
}
function _secHdr(sh, row, text, months) {
  sh.getRange(row, 1, 1, months + 1).merge()
    .setValue(text).setBackground("#2C3E50").setFontColor("#FFFFFF")
    .setFontWeight("bold").setFontSize(10);
  sh.setRowHeight(row, 24);
}
function _subHdr(sh, row, text, months) {
  sh.getRange(row, 1, 1, months + 1).merge()
    .setValue(text).setBackground("#D6EAF8").setFontWeight("bold").setFontSize(9);
  sh.setRowHeight(row, 18);
}
