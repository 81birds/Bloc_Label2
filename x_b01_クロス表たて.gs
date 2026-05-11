// function runAllProcessesCrossTable() {
//   ///console.log('--- 処理を開始します ---');
  
// createVerticalProductCrossTables();
// applyTableLayout();
// formatCrossTablesByBlockFinal();

//   ///console.log('--- すべての処理が正常に完了しました ---');
// }




// /**
//  * 前半：クロス表（A, B, C, D固定）、後半：箱数集計
//  */
// function createVerticalProductCrossTables() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sourceSheet = ss.getSheetByName('フリー入力用');
//   if (!sourceSheet) return;

//   const data = sourceSheet.getDataRange().getValues();
//   const header = data.shift();
  
//   const col = {
//     date: header.indexOf('納品日'),
//     location: header.indexOf('制作場所'),
//     product: header.indexOf('アソート'),
//     client: header.indexOf('クライアント名'),
//     area: header.indexOf('エリア'),
//     course: header.indexOf('コース番号'),
//     shop: header.indexOf('店名'),
//     qty: header.indexOf('数量'),
//     note: header.indexOf('備考')
//   };

//   // ★固定で確保する製品アソート名
//   const fixedProducts = ['A', 'B', 'C', 'D'];
//   // 固定で確保する箱タイプ
//   const fixedBoxTypes = [50, 40, 30, 20, 10];

//   const targets = [
//     { loc: "豊ビル", sheetName: "クロス表YT" },
//     { loc: "エルブ奥", sheetName: "クロス表ET" }
//   ];

//   targets.forEach(target => {
//     const targetSheet = ss.getSheetByName(target.sheetName);
//     if (!targetSheet) return;
//     targetSheet.clear();

//     const filteredData = data.filter(row => row[col.location] === target.loc && row[col.date]);
//     if (filteredData.length === 0) return;

//     // ヘッダー作成
//     const headerBase = ['納品日', '制作場所', 'クライアント名', 'エリア', 'コース番号', '店名', '備考'];
//     // 製品列は常に A, B, C, D
//     const leftHeader = [...headerBase, ...fixedProducts, '合計'];
//     const emptyCols = ["", ""]; 
//     const rightHeader = ['箱タイプ', '箱数', '数量小計'];
//     const fullHeader = [...leftHeader, ...emptyCols, ...rightHeader];
//     const startBoxIdx = leftHeader.length + emptyCols.length;

//     const dates = [...new Set(filteredData.map(row => 
//       Utilities.formatDate(new Date(row[col.date]), "JST", "yyyy年M月d日")
//     ))].sort();

//     let allOutputRows = [];

//     dates.forEach(dStr => {
//       const dateRows = filteredData.filter(row => 
//         Utilities.formatDate(new Date(row[col.date]), "JST", "yyyy年M月d日") === dStr
//       );

//       // 左：製品集計
//       const leftSummary = {};
//       dateRows.forEach(row => {
//         const key = `${row[col.client]}|${row[col.area]}|${row[col.course]}|${row[col.shop]}|${row[col.note]}`;
//         if (!leftSummary[key]) leftSummary[key] = {};
//         const pName = String(row[col.product]).trim();
//         leftSummary[key][pName] = (leftSummary[key][pName] || 0) + Number(row[col.qty]);
//       });

//       // 右：箱数集計
//       const rightSummary = {};
//       fixedBoxTypes.forEach(type => rightSummary[type] = 0);
//       dateRows.forEach(row => {
//         const bType = Number(row[col.qty]);
//         if (rightSummary[bType] === undefined) rightSummary[bType] = 0;
//         rightSummary[bType] += 1;
//       });

//       const leftRowKeys = Object.keys(leftSummary);
//       const rightRowKeys = Object.keys(rightSummary).map(Number).sort((a,b)=>b-a);
//       const maxRows = Math.max(leftRowKeys.length + 1, rightRowKeys.length + 1);

//       for (let i = 0; i < maxRows; i++) {
//         let row = new Array(fullHeader.length).fill("");

//         // --- 左側（A, B, C, D固定） ---
//         if (i < leftRowKeys.length) {
//           const [client, area, course, shop, note] = leftRowKeys[i].split('|');
//           const pData = leftSummary[leftRowKeys[i]];
//           let rowTotal = 0;
//           const leftVals = [dStr, target.loc, client, area, course, shop, note];
          
//           // 固定された A, B, C, D の順に値を抽出
//           fixedProducts.forEach(p => {
//             const val = pData[p] || 0;
//             leftVals.push(val || "");
//             rowTotal += val;
//           });
//           leftVals.push(rowTotal);
//           row.splice(0, leftVals.length, ...leftVals);
//         } else if (i === leftRowKeys.length) {
//           row[0] = "日付計";
//           let dateTotal = 0;
//           fixedProducts.forEach((p, idx) => {
//             const colSum = leftRowKeys.reduce((sum, k) => sum + (leftSummary[k][p] || 0), 0);
//             row[headerBase.length + idx] = colSum || "";
//             dateTotal += colSum;
//           });
//           row[leftHeader.length - 1] = dateTotal;
//         }

//         // --- 右側（箱数集計） ---
//         if (i < rightRowKeys.length) {
//           const bType = rightRowKeys[i];
//           const bCount = rightSummary[bType];
//           row[startBoxIdx] = bType;
//           row[startBoxIdx + 1] = bCount || 0;
//           row[startBoxIdx + 2] = bType * (bCount || 0);
//         } else if (i === rightRowKeys.length) {
//           row[startBoxIdx] = "総合計";
//           const totalBoxes = Object.values(rightSummary).reduce((a, b) => a + b, 0);
//           const totalQty = Object.keys(rightSummary).reduce((sum, type) => sum + (Number(type) * rightSummary[type]), 0);
//           row[startBoxIdx + 1] = totalBoxes;
//           row[startBoxIdx + 2] = totalQty;
//         }

//         allOutputRows.push(row);
//       }
//       allOutputRows.push(new Array(fullHeader.length).fill(""));
//       allOutputRows.push(fullHeader);
//     });

//     const finalData = [fullHeader, ...allOutputRows];
//     targetSheet.getRange(1, 1, finalData.length, fullHeader.length).setValues(finalData);
    
//     // 装飾・レイアウト
//     const range = targetSheet.getRange(1, 1, finalData.length, fullHeader.length);
//     range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
//     range.setFontSize(10).setVerticalAlignment("middle");
//     targetSheet.autoResizeColumns(1, fullHeader.length);
//     // 店名と備考列を少し広めに
//     targetSheet.setColumnWidth(6, 150); // 店名
//     targetSheet.setColumnWidth(7, 120); // 備考
//   });
// }

// /**
//  * クロス表YT/ETの全テーブルに対してレイアウトを適用する
//  */
// function applyTableLayout() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const targetSheetNames = ["クロス表YT", "クロス表ET"];
  
//   targetSheetNames.forEach(sheetName => {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) return;

//     const lastRow = sheet.getLastRow();
//     const lastCol = sheet.getLastColumn();
//     if (lastRow < 1) return;

//     // 1. 全体の書式をリセット
//     const allRange = sheet.getRange(1, 1, lastRow, lastCol);
//     allRange.setBorder(false, false, false, false, false, false)
//             .setBackground(null)
//             .setFontWeight("normal")
//             .setHorizontalAlignment("left");

//     const data = allRange.getValues();

//     // 2. 行をスキャンして表を特定
//     data.forEach((row, i) => {
//       const rIdx = i + 1;

//       // --- 前半：ABCD集計（A列〜L列）の処理 ---
//       // A列〜L列の範囲で何かしらデータがある行に罫線を引く
//       const leftRange = sheet.getRange(rIdx, 1, 1, 12);
//       const leftRowValues = row.slice(0, 12);
      
//       if (leftRowValues.join("").trim() !== "") {
//         leftRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
        
//         // ヘッダー（納品日）の場合は薄ピンク
//         if (row[0] === "納品日") {
//           leftRange.setBackground("#fce4ec").setFontWeight("bold").setHorizontalAlignment("center");
//         }
//         // 集計行（日付計）の場合は薄グレー
//         if (row[0] === "日付計") {
//           leftRange.setBackground("#f3f3f3").setFontWeight("bold");
//         }
//       }

//       // --- 後半：箱数計算（O列〜Q列）の処理 ---
//       // L列から2列空けると O(15), P(16), Q(17) 列
//       const boxStartCol = 15; 
//       const rightRange = sheet.getRange(rIdx, boxStartCol, 1, 3);
//       const rightRowValues = row.slice(boxStartCol - 1, boxStartCol + 2);

//       if (rightRowValues.join("").trim() !== "") {
//         rightRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

//         // ヘッダー（箱タイプ）の場合は水色
//         if (row[boxStartCol - 1] === "箱タイプ") {
//           rightRange.setBackground("#e1f5fe").setFontWeight("bold").setHorizontalAlignment("center");
//         }
//         // 総合計行の場合は水色（またはグレー）
//         if (row[boxStartCol - 1] === "総合計") {
//           rightRange.setBackground("#e1f5fe").setFontWeight("bold");
//         }
//       }
//     });

//     // 列幅の微調整（必要に応じて数値を変更してください）
//     sheet.setColumnWidths(1, 12, 60); // A〜Lを一律60px
//     sheet.setColumnWidth(6, 120);    // 店名は広めに
//     sheet.setColumnWidths(13, 2, 20); // 空白2列を狭く
//     sheet.setColumnWidths(15, 3, 60); // O〜Qを60px
//   });
// }

// function formatCrossTablesByBlockFinal() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheetNames = ['クロス表YT', 'クロス表ET'];

//   sheetNames.forEach(name => {
//     const sheet = ss.getSheetByName(name);
//     if (!sheet) return;

//     let lastRow = sheet.getLastRow();
//     if (lastRow < 1) return;

//     let data = sheet.getRange(1, 1, lastRow, 2).getValues();
//     let insertPositions = [];

//     for (let i = 0; i < data.length; i++) {
//       // 「納品日」という見出し行を検知
//       if (String(data[i][0]).trim() === "納品日") {
//         let targetRowIdx = i + 1;
        
//         if (targetRowIdx < data.length) {
//           let dateVal = data[targetRowIdx][0]; // A列：日付
//           let locVal = data[targetRowIdx][1];  // B列：制作場所

//           let dateStr = "";
//           if (dateVal !== "") {
//             // 日付を「5/7木」の形式に整形
//             if (dateVal instanceof Date) {
//               dateStr = Utilities.formatDate(dateVal, "JST", "M/d") + "日月火水木金土"[dateVal.getDay()];
//             } else {
//               dateStr = String(dateVal);
//             }
//           }

//           insertPositions.push({
//             rowIndex: i + 1,
//             title: `${locVal}  ${dateStr}`
//           });
//         }
//       }
//     }

//     // 下から順に空行挿入とタイトル書き込み
//     for (let j = insertPositions.length - 1; j >= 0; j--) {
//       let pos = insertPositions[j];
//       sheet.insertRowBefore(pos.rowIndex);
//       // 旧C列（3列目）の真上にタイトルをセット（背景色なし）
//       sheet.getRange(pos.rowIndex, 3)
//            .setValue(pos.title)
//            .setFontWeight("bold");
//     }

//     // A・B列を削除（これでタイトルが新しいA列の真上に来る）
//     sheet.deleteColumns(1, 2);

//     // A列の幅を広げる（150ピクセル程度）
//     sheet.setColumnWidth(1, 150);
//   });
// }
