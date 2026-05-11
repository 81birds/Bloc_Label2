// function decomposeRecipe() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const rcp = ss.getSheetByName("レシピDB");
//   const freein = ss.getSheetByName("フリー入力用");
//   const kazai = ss.getSheetByName("花材分解");

//   // 1. レシピDBの読み込み（D:製品名, E-I:花材名）
//   const rcpLastRow = rcp.getLastRow();
//   if (rcpLastRow < 2) return;
//   const rcpData = rcp.getRange("D2:I" + rcpLastRow).getValues();
  
//   let recipeMap = {};
//   rcpData.forEach(row => {
//     const productName = row[0]; // D列
//     // E(1), F(2), G(3), H(4), I(5) 列から空白以外の花材を抽出
//     const ingredients = [row[1], row[2], row[3], row[4], row[5]].filter(item => item !== "");
//     if (productName) {
//       recipeMap[productName] = ingredients;
//     }
//   });

//   // 2. フリー入力用（元データ）の読み込み（G:数量, I:制作場所, J:仕入日, K:製品名）
//   const freeLastRow = freein.getLastRow();
//   if (freeLastRow < 2) return;
//   const freeData = freein.getRange("G2:K" + freeLastRow).getValues();
  
//   let output = [];

//   // 3. データ分解ロジック
//   freeData.forEach(row => {
//     const qty = Number(row[0]) || 0; // G列: 数量
//     const loc = row[2];             // I列: 制作場所
//     const dateVal = row[3];          // J列: 仕入日
//     const productName = row[4];      // K列: 製品名

//     if (!productName || qty === 0) return;

//     const ingredients = recipeMap[productName];

//     if (ingredients) {
//       ingredients.forEach(flower => {
//         const formattedDate = dateVal instanceof Date 
//           ? Utilities.formatDate(dateVal, "JST", "yyyy/MM/dd") 
//           : dateVal;
        
//         // 出力列構成 [B:仕入日, C:制作場所, D:花材名, E:数量, F:製品名]
//         output.push([
//           formattedDate, // B列
//           loc,           // C列
//           flower,        // D列
//           qty,           // E列 ★ここに追加
//           productName    // F列
//         ]);
//       });
//     }
//   });

//   // 4. 花材分解シートへの書き出し
//   if (output.length > 0) {
//     const lastRow = kazai.getLastRow();
//     if (lastRow > 1) {
//       // B2からF列の最終行までをクリア
//       kazai.getRange(2, 2, lastRow - 1, 5).clearContent();
//     }
//     // まとめて書き込み
//     kazai.getRange(2, 2, output.length, 5).setValues(output);
//   } else {
//     SpreadsheetApp.getUi().alert("集計対象のデータが見つかりませんでした。");
//   }
// }


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

