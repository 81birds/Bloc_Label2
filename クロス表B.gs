function copyWithFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetA = ss.getSheetByName("クロス表A");
  const sheetB = ss.getSheetByName("クロス表B");
  
  // 1. 設定
  sheetB.clear(); // 既存の内容と書式をクリア
  const lastRowA = sheetA.getLastRow();
  const lastColA = sheetA.getLastColumn();
  
  // タイトル範囲 (2行目〜4行目)
  const headerRange = sheetA.getRange(1, 1, 3, lastColA);
  
  // 2. 品目行を1つずつループして「クロス表B」に作成
  let destRow = 1; // 書き込み開始行
  
  for (let i = 4; i <= lastRowA; i++) {
    const firstCellVal = sheetA.getRange(i, 1).getValue();
    
    // A列が「総計」なら終了
    if (firstCellVal === "合計") break;
    
    // --- タイトル3行をコピー ---
    headerRange.copyTo(sheetB.getRange(destRow, 1));
    destRow += 3;
    
    // --- 品目1行をコピー ---
    const itemRange = sheetA.getRange(i, 1, 1, lastColA);
    itemRange.copyTo(sheetB.getRange(destRow, 1));
    destRow += 1;
    
    // --- 空白2行分を飛ばす（destRowを2つ進める） ---
    destRow += 2;
    
    // ※ 1セット合計6行
  }
}
