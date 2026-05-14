function applyAssortFontColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  /** ABCDアソートコード → 文字色 */
  const ASSORT_FONT_COLORS = {
    'A': '#FF0000', // 赤
    'B': '#0000FF', // 青
    'C': '#ffa500', // オレンジ
    'D': '#008000', // 緑
  };

  // 各シートの個別設定
  // scanCols: 処理する列番号の配列, isHeader: 行全体を1行ずつ横にスキャンする場合はtrue
  const configs = [
    { name: "フリー入力用", startRow: 2, scanCols: [8],  isHeader: false }, // H列(8)
    { name: "ラベル集計",   startRow: 2, scanCols: [6],  isHeader: false }, // F列(6)
    { name: "ラベルA3_ゆ", startRow: 1, scanCols: [2, 5], isHeader: false }, // B列(2), E列(5)
    { name: "ラベルA3_エ", startRow: 1, scanCols: [2, 5], isHeader: false }, // B列(2), E列(5)
    { name: "クロス表YT",   startRow: 2, scanCols: [13], isHeader: false }, // M列(13) ※先頭文字判定
    { name: "クロス表ET",   startRow: 2, scanCols: [13], isHeader: false }  // M列(13) ※先頭文字判定
  ];

  // 1. 通常の縦方向スキャン（列単位の処理）
  configs.forEach(cfg => {
    const sheet = ss.getSheetByName(cfg.name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < cfg.startRow) return;
    const numRows = lastRow - cfg.startRow + 1;

    cfg.scanCols.forEach(col => {
      const range = sheet.getRange(cfg.startRow, col, numRows, 1);
      const values = range.getValues();
      const currentFontColors = range.getFontColors(); // タイトル行などの既存色保護用
      const newColors = [];

      for (let i = 0; i < numRows; i++) {
        const val = String(values[i][0]);
        
        // 先頭の1文字を抽出（A50ならA、B40ならBになる）
        const firstChar = val.charAt(0);

        if (ASSORT_FONT_COLORS.hasOwnProperty(firstChar)) {
          newColors.push([ASSORT_FONT_COLORS[firstChar]]);
        } else if (val === "" || val === "undefined") {
          newColors.push([null]); // 空欄ならデフォルト(黒)に戻す
        } else {
          newColors.push([currentFontColors[i][0]]); // キーにない文字（タイトル等）は元の色を維持
        }
      }
      range.setFontColors(newColors);
    });
  });

  // 2. 「仕入集計」シートの横方向スキャン（C1からF1の処理）
  const siireSheet = ss.getSheetByName("仕入集計");
  if (siireSheet) {
    const siireRange = siireSheet.getRange("C1:F1");
    const siireValues = siireRange.getValues()[0];
    const currentSiireColors = siireRange.getFontColors()[0];
    const newSiireColors = [];

    for (let j = 0; j < siireValues.length; j++) {
      const val = String(siireValues[j]);
      const firstChar = val.charAt(0);

      if (ASSORT_FONT_COLORS.hasOwnProperty(firstChar)) {
        newSiireColors.push(ASSORT_FONT_COLORS[firstChar]);
      } else if (val === "" || val === "undefined") {
        newSiireColors.push(null);
      } else {
        newSiireColors.push(currentSiireColors[j]);
      }
    }
    siireRange.setFontColors([newSiireColors]);
  }
}
