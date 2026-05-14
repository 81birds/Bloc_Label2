function resetDataAndColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 処理対象のシート設定
  const configs = [
    { name: "フリー入力用", startRow: 2, startCol: 1, numCols: 15 }, // A2からO列 (15列)
    { name: "仕入集計", startRow: 2, startCol: 1, numCols: 7 }, // A2からG列 (7列)
    { name: "クロス表YT",   startRow: 1, startCol: 1, numCols: 15 }, // AからO列 (15列)
    { name: "クロス表ET",   startRow: 1, startCol: 1, numCols: 15 }, // AからO列 (15列)
    { name: "ラベル集計", startRow: 2, startCol: 1, numCols: 11 }, // A2からK列 (11列)
    { name: "ラベルA3_ゆ", startRow: 1, startCol: 1, numCols: 5 },   // AからE列 (5列)
    { name: "ラベルA3_エ", startRow: 1, startCol: 1, numCols: 5 },   // AからE列 (5列)


  ];

  configs.forEach(cfg => {
    const sheet = ss.getSheetByName(cfg.name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    // 開始行が最終行より大きい、またはデータが1行もない場合はスキップ
    if (lastRow < cfg.startRow) return;

    const numRows = lastRow - cfg.startRow + 1;
    const range = sheet.getRange(cfg.startRow, cfg.startCol, numRows, cfg.numCols);

    // 1. 文字列（値）をクリア
    range.clearContent();

    // 2. セルの背景色を「色なし」にリセット
    range.setBackground(null);

    // 3. 文字の色を「自動（黒）」にリセット
    range.setFontColor(null);
  });
}
