/**
 * A5以下のセル結合を解除し、元の値をすべてのセルに埋める
 */
function unmergeAndFillValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('元データ');
  
  // A5からA列の最終行までの範囲を取得
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return;
  const range = sheet.getRange(5, 1, lastRow - 4, 1);
  
  // 1. 結合範囲をすべて取得
  const mergedRanges = range.getMergedRanges();
  
  mergedRanges.forEach(mergedRange => {
    // 結合されているセルの左上にある値を取得
    const value = mergedRange.getValue();
    
    // 2. 結合を解除
    mergedRange.breakApart();
    
    // 3. 解除されたすべてのセルに元の値を書き込む
    mergedRange.setValue(value);
  });
}

