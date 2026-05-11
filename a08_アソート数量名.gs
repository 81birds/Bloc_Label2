function updateConcatColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("フリー入力用");
  const lastRow = sheet.getLastRow();

  // データがない場合は終了
  if (lastRow < 2) return;

  // G列(7列目)とH列(8列目)のデータを取得
  const data = sheet.getRange(2, 7, lastRow - 1, 2).getValues();

  // H列 + G列 の順で連結した配列を作成
  const results = data.map(row => {
    const gVal = row[0]; // G列
    const hVal = row[1]; // H列
    return [String(hVal) + String(gVal)];
  });

  // M列(13列目)に一括書き出し
  sheet.getRange(2, 13, results.length, 1).setValues(results);
}
