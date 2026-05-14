function resetAllLabels() {//店名のサイズ設定
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ラベルA3_ゆ");
  if (!sheet) return;

  // ラベルの開始行3から、12行おきに10枚分（5段分×2列）
  const targetRows = [3, 15, 27, 39, 51,63]; 
  const targetCols = [2, 5]; // B列(2)とE列(5)

  targetCols.forEach(col => {
    targetRows.forEach(row => {
      sheet.getRange(row, col).setFontSize(48);
    });
  });
}
