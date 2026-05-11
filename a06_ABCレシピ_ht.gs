function runAllProcessesThrough() {

applyAssortPattern2();
setHToCharCIfUnique();

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



// function applyAssortPattern() {////上のコードを使用中、以下はバックアップ
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const free = ss.getSheetByName("フリー入力用");
//   const ast = ss.getSheetByName("アソートDB");

//   // 1. アソートDB（マスター）を辞書化
//   const astData = ast.getRange("B2:Q" + ast.getLastRow()).getValues();
//   let masterMap = {};
//   astData.forEach(row => {
//     const dayOfWeek = row[0]; // B列: 曜日
//     const client = row[3];    // E列: クライアント
//     const shop = row[4];      // F列: 店名
//     // H列(索引6)からQ列(索引15)までのパターンを配列化（空白は除外）
//     const pattern = row.slice(6, 16).filter(item => item !== "");
    
//     if (dayOfWeek && client && shop && pattern.length > 0) {
//       const key = `${dayOfWeek}|${client}|${shop}`;
//       masterMap[key] = pattern;
//     }
//   });

//   // 2. フリー入力用のデータを取得
//   const lastRow = free.getLastRow();
//   if (lastRow < 2) return;
//   const freeRange = free.getRange("B2:H" + lastRow);
//   const freeData = freeRange.getValues();
  
//   const days = ["日", "月", "火", "水", "木", "金", "土"];
//   let results = [];
  
//   // 3. グループごとに処理（日付＋クライアント＋店名が同じ範囲を特定）
//   let i = 0;
//   while (i < freeData.length) {
//     const row = freeData[i];
//     const date = new Date(row[0]);
//     const dayStr = days[date.getDay()];
//     const client = row[1];
//     const shop = row[2];
//     const key = `${dayStr}|${client}|${shop}`;

//     // 同じグループが何行続くかカウント
//     let groupRows = [];
//     while (i < freeData.length && 
//            Utilities.formatDate(new Date(freeData[i][0]), "JST", "yyyyMMdd") === Utilities.formatDate(date, "JST", "yyyyMMdd") &&
//            freeData[i][1] === client && 
//            freeData[i][2] === shop) {
//       groupRows.push(i);
//       i++;
//     }

//     // パターンの割り当て
//     const pattern = masterMap[key];
//     groupRows.forEach((rowIndex, index) => {
//       if (pattern) {
//         // パターン数で割った余りを使うことで、先頭に戻るループを実現
//         const charToFill = pattern[index % pattern.length];
//         results[rowIndex] = [charToFill];
//       } else {
//         results[rowIndex] = [""]; // マスターになければ空欄
//       }
//     });
//   }

//   // 4. H列に一括書き出し
//   free.getRange(2, 8, results.length, 1).setValues(results);
// }


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




