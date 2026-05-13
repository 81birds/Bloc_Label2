function boxingAndABC_ALL() {


splitRowsByQuantityAdvanced2();
applyAssortPattern2();
setHToCharCIfUnique();
updateConcatColumn();//ABCアソート名と入数合体

}


function splitRowsByQuantityAdvanced2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('フリー入力用');
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // A列〜O列（15列分）を取得
  const range = sheet.getRange(1, 1, lastRow, 15);
  const values = range.getValues();
  
  const data = values.slice(1);
  const newValues = [];

  data.forEach(row => {
    let quantity = Number(row[6]); // G列: 数量
    
    if (isNaN(quantity) || quantity <= 0) {
      newValues.push(row);
      return;
    }

    // 分解後の数字を格納する配列
    let parts = [];

    // --- 分解ロジックの優先順位 ---
    
    // パターン1: 50と40の組み合わせで割り切れるか（90, 130, 140, 180など）
    let tempParts = findCombination(quantity, 50, 40);
    
    // パターン2: 上記がダメなら 40と30の組み合わせ（70, 110, 150など）
    if (tempParts.length === 0) {
      tempParts = findCombination(quantity, 40, 30);
    }

    // いずれかの組み合わせで見つかった場合
    if (tempParts.length > 0) {
      parts = tempParts;
    } else {
      // どちらの組み合わせでも分解できない（または50, 40, 30単体で割り切れる）場合
      if (quantity % 50 === 0) {
        parts = Array(quantity / 50).fill(50);
      } else if (quantity % 40 === 0) {
        parts = Array(quantity / 40).fill(40);
      } else if (quantity % 30 === 0) {
        parts = Array(quantity / 30).fill(30);
      } else {
        // 全く当てはまらない場合は分解しない
        parts = [quantity];
      }
    }

    // 分解されたパーツ分だけ行を複製
    parts.forEach(p => {
      const newRow = [...row];
      newRow[6] = p; // G列に分解した数値をセット
      newValues.push(newRow);
    });
  });

  // 書き込み処理
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).clearContent();
  }
  if (newValues.length > 0) {
    sheet.getRange(2, 1, newValues.length, 15).setValues(newValues);
  }
}

/**
 * 二つの数値の組み合わせで合計値が作れるか計算する補助関数
 * 補助関数は一括実行にはふくめない
 */
function findCombination(total, num1, num2) {
  // num1をできるだけ多く使うパターンから試行
  for (let i = Math.floor(total / num1); i >= 0; i--) {
    let remainder = total - (i * num1);
    if (remainder >= 0 && remainder % num2 === 0) {
      let res1 = Array(i).fill(num1);
      let res2 = Array(remainder / num2).fill(num2);
      return res1.concat(res2);
    }
  }
  return [];
}



function applyAssortPattern2() {//書き出し先のH列をスキャンしてブランクのセルのみに実行
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const free = ss.getSheetByName("フリー入力用");
  const ast = ss.getSheetByName("アソートDB");

  const astData = ast.getRange("B2:Q" + ast.getLastRow()).getValues();
  let masterMap = {};
  astData.forEach(row => {
    const dayOfWeek = row[0];
    const client = row[3];
    const shop = row[4];
    const pattern = row.slice(6, 16).filter(item => item !== "");
    if (dayOfWeek && client && shop && pattern.length > 0) {
      const key = `${dayOfWeek}|${client}|${shop}`;
      masterMap[key] = pattern;
    }
  });

  const lastRow = free.getLastRow();
  if (lastRow < 2) return;
  // B列からH列(7列分)を取得
  const freeRange = free.getRange("B2:H" + lastRow);
  const freeData = freeRange.getValues();
  
  const days = ["日", "月", "火", "水", "木", "金", "土"];
  // resultsを現在のH列の値で初期化しておく（上書き防止のため）
  let results = freeData.map(row => [row[6]]); 
  
  let i = 0;
  while (i < freeData.length) {
    const row = freeData[i];
    const date = new Date(row[0]);
    const dayStr = days[date.getDay()];
    const client = row[1];
    const shop = row[2];
    const key = `${dayStr}|${client}|${shop}`;

    let groupRows = [];
    const targetDateStr = Utilities.formatDate(date, "JST", "yyyyMMdd");

    while (i < freeData.length && 
           Utilities.formatDate(new Date(freeData[i][0]), "JST", "yyyyMMdd") === targetDateStr &&
           freeData[i][1] === client && 
           freeData[i][2] === shop) {
      groupRows.push(i);
      i++;
    }

    const pattern = masterMap[key];
    if (pattern) {
      groupRows.forEach((rowIndex, index) => {
        // ★ここが追加ポイント：H列が空の場合のみ、パターンを適用する
        if (results[rowIndex][0] === "") {
          const charToFill = pattern[index % pattern.length];
          results[rowIndex] = [charToFill];
        }
      });
    }
  }

  free.getRange(2, 8, results.length, 1).setValues(results);
}




function setHToCharCIfUnique() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("フリー入力用");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // B列(2列目)からD列(4列目)までのデータを取得
  const dataRange = sheet.getRange(2, 2, lastRow - 1, 3);
  const data = dataRange.getValues();
  
  // H列(8列目)の現在のデータを取得
  const hRange = sheet.getRange(2, 8, lastRow - 1, 1);
  const hresults = hRange.getValues();

  // 1. 「B列+D列」の組み合わせの出現回数をカウント
  const counts = {};
  for (let i = 0; i < data.length; i++) {
    // 日付オブジェクトを文字列に変換してキーを作成
    let bVal = data[i][0] instanceof Date ? data[i][0].getTime() : data[i][0];
    let dVal = data[i][2];
    let key = bVal + "_" + dVal;
    
    counts[key] = (counts[key] || 0) + 1;
  }

  // 2. カウントが1の場合のみ、H列の配列に "C" を入れる
  for (let i = 0; i < data.length; i++) {
    let bVal = data[i][0] instanceof Date ? data[i][0].getTime() : data[i][0];
    let dVal = data[i][2];
    let key = bVal + "_" + dVal;

    if (counts[key] === 1) {
      hresults[i] = ["C"];
    }
    // 1行でない（重複がある）場合は、既存のH列の値を保持します
  }

  // 3. H列に一括書き込み
  hRange.setValues(hresults);
}


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




