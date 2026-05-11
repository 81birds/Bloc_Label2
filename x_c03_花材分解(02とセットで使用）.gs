// function aggregateKazai() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const kazai = ss.getSheetByName("花材分解");
  
//   const lastRow = kazai.getLastRow();
//   if (lastRow < 2) return;

//   // B2:F列のデータを取得
//   const data = kazai.getRange(2, 2, lastRow - 1, 5).getValues();
  
//   let summaryMap = {};

//   // 1. データの集計
//   data.forEach(row => {
//     const date = row[0];   // B列: 仕入日
//     const loc = row[1];    // C列: 制作場所
//     const flower = row[2]; // D列: 花材名
//     const qty = Number(row[3]) || 0; // E列: 数量

//     if (!flower) return;

//     // 「日付＋場所＋花材名」を合算のキーにする
//     const key = `${date}|${loc}|${flower}`;

//     if (!summaryMap[key]) {
//       summaryMap[key] = {
//         date: date,
//         loc: loc,
//         flower: flower,
//         qty: 0
//       };
//     }
//     summaryMap[key].qty += qty;
//   });

//   // 2. 出力用データの作成
//   let output = [];
//   for (let key in summaryMap) {
//     const item = summaryMap[key];
//     output.push([
//       item.date,
//       item.loc,
//       item.flower,
//       item.qty,
//       "合算" // F列に「合算」と入力
//     ]);
//   }

//   // 3. シートへの上書き（一度クリアしてから書き出す）
//   kazai.getRange(2, 2, lastRow - 1, 5).clearContent();
//   if (output.length > 0) {
//     // 日付順・場所順に並べ替えると見やすいのでソート
//     output.sort((a, b) => new Date(a[0]) - new Date(b[0]) || a[1].localeCompare(b[1]));
    
//     kazai.getRange(2, 2, output.length, 5).setValues(output);
//   }
// }
