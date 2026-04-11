function runAllProcesses() {
  ///console.log('--- 処理を開始します ---');
  
  transferToLabelSummary();            
  deleteEmptyRowsInLabelSummary(); // 2. 空行削除
  splitOver30Rows();            // 3. 数値の多段分割
  setFractionFormat();

  ///console.log('--- すべての処理が正常に完了しました ---');
}


/**
 * フリー入力用シートからラベル集計シートへ指定列を転記
 */
function transferToLabelSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('フリー入力用');
  const targetSheet = ss.getSheetByName('ラベル集計');
  
  if (!sourceSheet || !targetSheet) {
    console.error('シートが見つかりません。');
    return;
  }

  // 1. ラベル集計シートの A2:I (2行目以下) をクリア
  const targetLastRow = targetSheet.getLastRow();
  if (targetLastRow >= 2) {
    targetSheet.getRange(2, 1, targetLastRow - 1, 9).clear(); 
  }

  // 2. フリー入力用シートのデータ取得
  const sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return; // データがない場合は終了

  const sourceData = sourceSheet.getRange(2, 1, sourceLastRow - 1, 8).getValues();
  
  // 3. データの並べ替え（配列操作）
  // フリー入力用の列番号 (0始まり): B=1, C=2, D=3, E=4, F=5, G=6, H=7
  const outputData = sourceData.map(row => {
    return [
      row[1], // A列：納品日(B)
      row[4], // B列：コース番号(E)
      row[2], // C列：クライアント名(C)
      row[3], // D列：店名(D)
      row[5], // E列：製品名(F)
      row[7], // F列：備考(H)
      row[6], // G列：数量(G)
      '',     // H列：空欄
      ''      // I列：空欄
    ];
  });

  // 4. まとめて貼り付け
  targetSheet.getRange(2, 1, outputData.length, 9).setValues(outputData);
  
  ///console.log('転記が完了しました。');
}

/**
 * 2. A列が空の行を削除
 */
function deleteEmptyRowsInLabelSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ラベル集計');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(1, 1, lastRow, 1).getValues();

  // 下の行から順に削除
  for (let i = lastRow - 1; i >= 1; i--) { // 2行目(index 1)まで
    if (data[i][0] === "" || data[i][0] === null || String(data[i][0]).trim() === "") {
      sheet.deleteRow(i + 1);
    }
  }
}

/**
 * 3. G列が30を超える場合、30ずつ多段分割し、H列に元の数を記録
 * 例：G列が70の場合 → 「30, 30, 10」の3行に分割
 */
function splitOver30Rows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ラベル集計');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; 

  const range = sheet.getRange(2, 1, lastRow - 1, 7);
  const data = range.getValues();
  const newData = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const valueG = Number(row[6]); // 元のG列の数値
    let remaining = valueG;

    if (valueG > 30) {
      // 30を超える間、30ずつの行を生成
      while (remaining > 30) {
        const rowPart = [...row];
        rowPart[6] = 30;    // 分子（30固定）
        rowPart[7] = valueG; // 分母（H列）
        newData.push(rowPart);
        remaining -= 30;
      }
      // 最後に残った数値を1行追加
      const rowLast = [...row];
      rowLast[6] = remaining;
      rowLast[7] = valueG;
      newData.push(rowLast);
    } else {
      // 30以下の場合はそのまま、H列に分母をセット
      const rowNormal = [...row];
      rowNormal[7] = valueG;
      newData.push(rowNormal);
    }
  }

  // 既存データをH列(8列目)まで含めてクリア
  const finalLastRow = sheet.getLastRow();
  if (finalLastRow >= 2) {
    sheet.getRange(2, 1, finalLastRow - 1, 8).clearContent();
  }
  
  // A-H列(8列)を書き込む
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, 8).setValues(newData);
  }
}



function setFractionFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ラベル集計');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // ヘッダーのみなら終了

  // G列(7列目)とH列(8列目)のデータを取得
  const range = sheet.getRange(2, 7, lastRow - 1, 2);
  const data = range.getValues();
  const results = [];

  for (let i = 0; i < data.length; i++) {
    const valG = data[i][0]; // G列
    const valH = data[i][1]; // H列
    
    if (valG !== "" && valH !== "") {
      // 「30/55」のような文字列を作成
      results.push([valG + "/" + valH]);
    } else {
      results.push([""]);
    }
  }

  // I列(9列目)にまとめて書き込み
  sheet.getRange(2, 9, results.length, 1).setValues(results);
  
 // console.log('I列への分数表記の入力が完了しました。');
}



