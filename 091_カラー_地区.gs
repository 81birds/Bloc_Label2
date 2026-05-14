function applyDistrictColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  /** エリア名 → 背景色 */
  const DISTRICT_COLORS = {
    '旭川地区':    '#FFF9C4',
    '函館地区':    '#979fe3',
    '直納分':      '#add8e6',
    '室蘭地区':    '#d3d3d3',
    '苫小牧地区':  '#d3d3d3',
    'ラルズアークス': '#98fb98',
  };

  // 各シートの設定（targetCols の中に塗る列の番号を正しく指定しました）
  const sheetConfigs = [
    { name: "フリー入力用", startRow: 2, scanCol: 12, targetCols: [4, 12] }, // L列(12)スキャン、D(4)とL(12)を塗る
    { name: "クロス表YT",   startRow: 2, scanCol: 2,  targetCols: [2, 4]  }, // B列(2)スキャン、B(2)とD(4)を塗る
    { name: "クロス表ET",   startRow: 2, scanCol: 2,  targetCols: [2, 4]  }, // B列(2)スキャン、B(2)とD(4)を塗る
    { name: "ラベル集計",   startRow: 2, scanCol: 11, targetCols: [4, 11] } // K列(11)スキャン、D(4)とK(11)を塗る
  ];

  sheetConfigs.forEach(cfg => {
    const sheet = ss.getSheetByName(cfg.name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < cfg.startRow) return;

    const numRows = lastRow - cfg.startRow + 1;
    
    // スキャン対象の列データと、現在の各列の背景色をあらかじめ取得しておく
    const scanValues = sheet.getRange(cfg.startRow, cfg.scanCol, numRows, 1).getValues();
    
    // 列ごとに現在の背景色を取得（タイトル行の色を維持するために使用）
    const currentColors = {};
    cfg.targetCols.forEach(col => {
      currentColors[col] = sheet.getRange(cfg.startRow, col, numRows, 1).getBackgrounds();
    });

    // 新しく適用するカラーリストの配列を準備
    const newColorLists = {};
    cfg.targetCols.forEach(col => {
      newColorLists[col] = [];
    });

    for (let i = 0; i < numRows; i++) {
      const areaName = scanValues[i][0];

      // パレットのキーに登録されている地区名か、または完全に空欄であるか判定
      const isDistrict = DISTRICT_COLORS.hasOwnProperty(areaName);
      const isEmpty = (areaName === "" || areaName === null || areaName === undefined);

      cfg.targetCols.forEach(col => {
        if (isDistrict) {
          // 該当する地区名なら、パレットの色を塗る
          newColorLists[col].push([DISTRICT_COLORS[areaName]]);
        } else if (isEmpty) {
          // 空欄の行なら、色なし(透明)にする
          newColorLists[col].push([null]);
        } else {
          // 地区名でも空欄でもない場合（＝別の表のタイトル行など）は、現在の色をそのままキープして触らない
          const existingColor = currentColors[col][i][0];
          newColorLists[col].push([existingColor]);
        }
      });
    }

    // 各列に一括反映
    cfg.targetCols.forEach(col => {
      sheet.getRange(cfg.startRow, col, numRows, 1).setBackgrounds(newColorLists[col]);
    });
  });
}
