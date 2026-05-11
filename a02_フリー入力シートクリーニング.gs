function freeSheetCleaningAllProcesses() {
  ///console.log('--- 処理を開始します ---');
  
deleteZeroOrEmptyRows();//空欄とゼロ値を削除
cleanShopNames();//、、、店をいったん削除

  ///console.log('--- すべての処理が正常に完了しました ---');
}

/**
 * フリー入力用シートのG列が0または空欄の行を削除する
 */
function deleteZeroOrEmptyRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('フリー入力用');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // G列（7列目）のデータを一括取得
  const range = sheet.getRange(1, 7, lastRow, 1);
  const values = range.getValues();

  // 下の行から順にループ（行の削除でインデックスがズレるのを防ぐため）
  for (let i = lastRow - 1; i >= 1; i--) {
    const val = values[i][0];
    
    // 0 または 空欄（null, 空文字）の場合に削除
    if (val === 0 || val === "" || val === null) {
      sheet.deleteRow(i + 1);
    }
  }
  
  //console.log('0または空欄の行の削除が完了しました。');
}




/**
 * D列の店名から末尾の「店」と空白を削除する
 */
function cleanShopNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('フリー入力用');
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // D列(2行目〜最終行)を取得
  const range = sheet.getRange(2, 4, lastRow - 1, 1);
  const values = range.getValues();

  // 1行ずつクリーニング処理
  const cleanedValues = values.map(row => {
    let name = row[0];
    if (typeof name !== 'string') return [name];

    // 1. 空白（半角・全角）をすべて削除
    name = name.replace(/[\s\u3000]/g, "");

    // 2. 末尾の「店」を削除
    // $ は末尾を意味します
    name = name.replace(/店$/, "");

    return [name];
  });

  // シートに書き戻し
  range.setValues(cleanedValues);
}

