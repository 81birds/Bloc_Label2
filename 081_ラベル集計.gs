/**
 * すべてのプロセスを順番に実行する
 */
function runAllProcesses2() {
  transferToLabelSummary3();            
  deleteEmptyRowsInLabelSummary2(); 
  splitOver30Rows2(); 
  setFractionFormat2();
  addStoreSuffix();
  formatTextDatesInGAS();

}

/**
 * 1. フリー入力用(B:L)からラベル集計(A:K)へ転記
 */
function transferToLabelSummary3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('フリー入力用');
  const targetSheet = ss.getSheetByName('ラベル集計');
  
  if (!sourceSheet || !targetSheet) return;

  // ラベル集計シートの A2:K (11列分) をクリア
  const targetLastRow = targetSheet.getLastRow();
  if (targetLastRow >= 2) {
    targetSheet.getRange(2, 1, targetLastRow - 1, 11).clearContent(); 
  }

  const sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return;

  // 転記元 B列(2)〜L列(12)まで11列分取得,,14列ぶんまで修正
  const sourceData = sourceSheet.getRange(2, 2, sourceLastRow - 1, 14).getValues();
  
  // マッピング（転記先 A-K列）
  const outputData = sourceData.map(row => {
    return [
      row[13],  // 転記先のA列：転記元納品日(B)から表示納品日（O)に変更
      row[3],  // B列：コース番号(E)
      row[1],  // C列：クライアント名(C)
      row[12],  // D列：店名(D)から、Nの表示店名にする
      row[4],  // E列：製品名(F)
      row[6],  // F列：備考(H)
      row[5],  // G列：数量(G)分子
      '',      // H列：分母(J)分母用
      '',      // I列：分数表記用
      row[7],  // J列：予備/その他(I)
      row[10]  // K列：エリア(L)
    ];
  });

  targetSheet.getRange(2, 1, outputData.length, 11).setValues(outputData);
}

/**
 * 2. ラベル集計シートの A列(納品日)が空の行を削除
 */
function deleteEmptyRowsInLabelSummary2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ラベル集計');
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = lastRow - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === "") {
      sheet.deleteRow(i + 1);
    }
  }
}

/**
 * 3. G列(数量)を分割し、11列(K列)までデータを保持する
 */
function splitOver30Rows2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ラベル集計');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; 

  // K列(11列)まで取得範囲を拡大
  const range = sheet.getRange(2, 1, lastRow - 1, 11);
  const data = range.getValues();
  const newData = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const valueG = Number(row[6]); 
    let remaining = valueG;
    const limit = 50; 

    if (valueG > limit) {
      while (remaining > limit) {
        const rowPart = [...row];
        rowPart[6] = limit;    
        rowPart[7] = valueG;   
        newData.push(rowPart);
        remaining -= limit;
      }
      const rowLast = [...row];
      rowLast[6] = remaining;
      rowLast[7] = valueG;
      newData.push(rowLast);
    } else {
      const rowNormal = [...row];
      rowNormal[7] = valueG;
      newData.push(rowNormal);
    }
  }

  const finalLastRow = sheet.getLastRow();
  if (finalLastRow >= 2) {
    sheet.getRange(2, 1, finalLastRow - 1, 11).clearContent();
  }
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, 11).setValues(newData);
  }
}

/**
 * 4. 分数表記をI列(9列目)にセット
 */
function setFractionFormat2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ラベル集計');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 7, lastRow - 1, 2); 
  const data = range.getValues();
  const results = [];

  for (let i = 0; i < data.length; i++) {
    const valG = data[i][0]; 
    const valH = data[i][1]; 
    results.push(valG && valH ? [valG + "/" + valH] : [""]);
  }

  sheet.getRange(2, 9, results.length, 1).setValues(results);
}



/**
 * ラベル集計シートのD列（店名）の末尾に「店」を追記する
 */
function addStoreSuffix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ラベル集計');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // D列（4列目）のデータを取得
  const range = sheet.getRange(2, 4, lastRow - 1, 1);
  const values = range.getValues();

  const newValues = values.map(row => {
    let val = String(row[0]).trim();
    
    // 値が空でなく、かつ末尾が「店」で終わっていない場合のみ「店」を足す
    if (val !== "" && !val.endsWith("店")) {
      return [val + "店"];
    }
    return [val];
  });

  // 書き戻し
  range.setValues(newValues);
}



function formatTextDatesInGAS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ラベル集計");
  
  if (!sheet) {
    Browser.msgBox("「ラベル集計」シートが見つかりません。");
    return;
  }
  
  // A列の最終行を取得
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // データがなければ終了
  
  // A2セルから最終行までのデータ範囲を取得
  var range = sheet.getRange(2, 1, lastRow - 1, 1);
  var values = range.getValues(); // 2次元配列としてデータを取得
  
  // 1行ずつ文字をスキャンして変換
  for (var i = 0; i < values.length; i++) {
    var originalValue = String(values[i][0]).trim();
    
    // 「〇月〇日(曜日)」または「〇月〇日（曜日）」の形式にマッチする正規表現
    var match = originalValue.match(/^(\d{1,2})月(\d{1,2})日[\(（]([日月火水木金土])[\)）]$/);
    
    if (match) {
      var month = match[1]; // 月
      var day = match[2];   // 日
      var weekday = match[3]; // 曜日（「木」など）
      
      // 「5/14/木」の形式に組み立てて配列を書き換え
      values[i][0] = month + "/" + day + "/" + weekday;
    }
  }
  
  // セルの表示形式を「プレーンテキスト」にしてから、変換後のデータを一括で書き戻し
  range.setNumberFormat("@"); 
  range.setValues(values);
}




