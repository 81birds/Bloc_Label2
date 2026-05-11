/**
 * すべてのプロセスを順番に実行する
 */
function runAllProcesses2() {
  transferToLabelSummary2();            
  deleteEmptyRowsInLabelSummary2(); 
  splitOver30Rows2(); 
  setFractionFormat2();
  addStoreSuffix();
}

/**
 * 1. フリー入力用(B:L)からラベル集計(A:K)へ転記
 */
function transferToLabelSummary2() {
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

  // 転記元 B列(2)〜L列(12)まで11列分取得
  const sourceData = sourceSheet.getRange(2, 2, sourceLastRow - 1, 11).getValues();
  
  // マッピング（転記先 A-K列）
  const outputData = sourceData.map(row => {
    return [
      row[0],  // A列：納品日(B)
      row[3],  // B列：コース番号(E)
      row[1],  // C列：クライアント名(C)
      row[2],  // D列：店名(D)
      row[4],  // E列：製品名(F)
      row[6],  // F列：備考(H)
      row[5],  // G列：数量(G)
      row[8],  // H列：制作場所(J)
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




