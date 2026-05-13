function freeSheetCleaningAllProcesses() {
  ///console.log('--- 処理を開始します ---');

transferData();//元データからフリー入力用に転記  
deleteZeroOrEmptyRows();//空欄とゼロ値を削除
cleanShopNames();//、、、店をいったん削除
clearFreeInputBackgrounds();//背景色クリア

  ///console.log('--- すべての処理が正常に完了しました ---');
}



/**
 * MOTOシートからデータをスキャンし、SAKIシートに日付ごとに転記する
 * 日付(2行目・E列〜)と商品名(3行目・E列〜)である必要
 * 店舗名はC5からはじまっている必要
 * 
 */
function transferData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const motoSheet = ss.getSheetByName('元データ');
  const sakiSheet = ss.getSheetByName('フリー入力用');

  // MOTOシートの全データを配列として取得
  const motoValues = motoSheet.getDataRange().getValues();

  // 1. 日付(2行目・E列〜)と商品名(3行目・E列〜)を取得
  const dates = motoValues[1].slice(4); 
  const productNames = motoValues[2].slice(4);

  const outputData = [];

  // 2. 日付を起点にループ（1つの日付が終わったら次の日付へ）
  for (let j = 0; j < dates.length; j++) {
    const dateVal = dates[j];
    const productName = productNames[j]; // その列の商品名
    
    // 日付がなくなれば、その右側の列は処理せず終了
    if (!dateVal || dateVal === "") break;

    // その日付・商品に対して、全店舗（5行目以降）をスキャン
    for (let i = 4; i < motoValues.length; i++) {
      const shopName = motoValues[i][2]; // C列(インデックス2): 店名
      if (!shopName) continue; 

      const quantity = motoValues[i][4 + j]; // E列以降の数量
      
      // 配列の構造: [A, B(日付), C, D(店名), E, F(商品名), G(数量)]
      // インデックス: 0, 1, 2, 3, 4, 5, 6
      outputData.push(["", dateVal, "", shopName, "", productName, quantity]);
    }
  }

  // 3. SAKIシートに書き込み
  if (outputData.length > 0) {
    // 書き込み前にB2:Oの範囲をクリア
    const lastRow = sakiSheet.getLastRow();
    if (lastRow > 1) {
      sakiSheet.getRange("B2:O" + lastRow).clearContent();
    }

    // B2からoutputDataの列数(7列分)を一括書き込み
    // getRange(行, 列, 行数, 列数) -> (2, 1, 数, 7) で A列からG列までを対象にする
    sakiSheet.getRange(2, 1, outputData.length, 7).setValues(outputData);
    
    // SpreadsheetApp.getUi().alert("転記完了");
  } else {
    SpreadsheetApp.getUi().alert("転記対象のデータが見つかりませんでした。");
  }
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



 //D列の店名から末尾の「店」と空白を削除する
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




function clearFreeInputBackgrounds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("フリー入力用");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // A2セルからO列の最終行までの範囲を取得
  const range = sheet.getRange(2, 1, lastRow - 1, 15); // 15列目 = O列
  
  // 背景色をクリア（空文字を指定することで「色なし」になります）
  range.setBackground("");
}


