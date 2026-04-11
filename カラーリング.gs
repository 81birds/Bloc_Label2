/**
 * メイン処理：クロス表Aの特殊な範囲塗りつぶし
 */
function applyColoring() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- 1. シート「ラベル」の処理（変更なし） ---
  const labelSheet = ss.getSheetByName('ラベル');
  if (labelSheet) {
    const lastRow = labelSheet.getLastRow();
    if (lastRow >= 1) {
      const rangeB = labelSheet.getRange(1, 2, lastRow, 1);
      const rangeE = labelSheet.getRange(1, 5, lastRow, 1);
      [rangeB, rangeE].forEach(r => r.setBackground('#ffffff'));
      colorCellsByKeyword(rangeB);
      colorCellsByKeyword(rangeE);
    }
  }

    // --- 2. シート「クロス表A」の処理（新ロジック） ---
  const pvtSheet = ss.getSheetByName('クロス表A');
  if (pvtSheet) {
    const lastCol = pvtSheet.getLastColumn();
    const lastRow = pvtSheet.getLastRow();
    
    // 【修正点】シート全体の書式（背景色、文字色、交互の背景など）をクリア
    pvtSheet.clear({formatOnly: true});

    if (lastCol >= 1 && lastRow >= 1) {
      applyPivotBlockColoring(pvtSheet, lastCol, lastRow);
    }
  }

}

/**
 * クロス表A専用：店名から合計までのブロック塗りつぶし
 */
function applyPivotBlockColoring(sheet, lastCol, lastRow) {
  // 2行目のデータを取得
  const values = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rules = getLabelColorRules();
  
  let currentColor = null;
  let isColoring = false;

  for (let j = 0; j < lastCol; j++) {
    const cellValue = String(values[j]).trim();
    
    // 1. 新しい店名（キーワード）が見つかったか判定
    let foundRule = null;
    rules.forEach(rule => {
      if (cellValue.includes(rule.keyword) && !cellValue.includes('合計')) {
        foundRule = rule;
      }
    });

    if (foundRule) {
      currentColor = foundRule.color;
      isColoring = true; // 塗りつぶしモード開始
    }

    // 2. 塗りつぶしモード中の処理
    if (isColoring && currentColor) {
      // 2行目、3行目、4行目を同じ色で塗る
      sheet.getRange(1, j + 1, 3, 1).setBackground(currentColor);

      // 3. 「合計」が含まれるセルに到達した場合
      if (cellValue.includes('合計')) {
        // この列だけデータがある最下行まで塗る
        if (lastRow > 3) {
          sheet.getRange(4, j + 1, lastRow - 3, 1).setBackground(currentColor);
        }
        // この合計列を最後に、この店名の塗りつぶしモードを終了
        isColoring = false;
        currentColor = null;
      }
    }
  }
}

/**
 * ラベルシート用：単純キーワード判定
 */
function colorCellsByKeyword(range) {
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const rules = getLabelColorRules();

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const val = String(values[i][j]).trim();
      rules.forEach(rule => {
        if (val.includes(rule.keyword)) backgrounds[i][j] = rule.color;
      });
    }
  }
  range.setBackgrounds(backgrounds);
}

/**
 * 対応ルール
 */
function getLabelColorRules() {
  return [
    { keyword: 'コープ', color: '#c8fa7e' },
    { keyword: 'ラルズ', color: '#ff6666' },
    { keyword: 'アークス', color: '#ff9933' },
    { keyword: '北海市場', color: '#1ebbd9' },
    { keyword: 'サツドラ', color: '#3399ff' }
  ];
}
