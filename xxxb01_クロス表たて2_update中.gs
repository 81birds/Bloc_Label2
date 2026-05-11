// /**
//  * メイン実行関数
//  */
// function runAllProcessesCrossTable2555() {
//   createVerticalProductCrossTables555();
//   applyTableLayout555();
//   formatCrossTablesByBlockFinal555();
// }

// /**
//  * 1. データ集計と書き出し
//  */
// function createVerticalProductCrossTables555() {
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
//     note: header.indexOf('備考'),
//     assortM: 12 // M列
//   };

//   const fixedProducts = ['A', 'B', 'C', 'D'];
//   const fixedBoxTypes = [50, 40, 30, 20, ,15,10];
//   const targets = [
//     { loc: "豊ビル", sheetName: "クロス表YT" },
//     { loc: "エルブ奥", sheetName: "クロス表ET" }
//   ];

//   targets.forEach(target => {
//     const targetSheet = ss.getSheetByName(target.sheetName);
//     if (!targetSheet) return;

//     targetSheet.clear();

//     const filteredData = data.filter(row =>
//       row[col.location] === target.loc && row[col.date]
//     );

//     if (filteredData.length === 0) return;

//     const headerBase = ['納品日', '制作場所', 'クライアント名', 'エリア', 'コース番号', '店名', '備考'];
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

//       // 左側集計
//       const leftSummary = {};
//       dateRows.forEach(row => {
//         const key = `${row[col.client]}|${row[col.area]}|${row[col.course]}|${row[col.shop]}|${row[col.note]}`;
//         if (!leftSummary[key]) leftSummary[key] = {};
//         const pName = String(row[col.product]).trim();
//         leftSummary[key][pName] = (leftSummary[key][pName] || 0) + Number(row[col.qty]);
//       });

//       // 右側集計
//       const rightSummary = {};
//       fixedBoxTypes.forEach(type => { rightSummary[type] = 0; });
//       dateRows.forEach(row => {
//         const bType = Number(row[col.qty]);
//         if (rightSummary[bType] !== undefined) rightSummary[bType] += 1;
//       });

//       // M列集計
//       const assortSummary = {};
//       dateRows.forEach(row => {
//         const assort = String(row[col.assortM]).trim();
//         if (assort) assortSummary[assort] = (assortSummary[assort] || 0) + 1;
//       });

//       const leftRowKeys = Object.keys(leftSummary);
//       const rightRowKeys = Object.keys(rightSummary).map(Number).sort((a, b) => b - a);
//       const assortKeys = Object.keys(assortSummary).sort((a, b) => {
//         const matchA = String(a).match(/^([A-Za-z]+)(\d+)$/);
//         const matchB = String(b).match(/^([A-Za-z]+)(\d+)$/);
//         if (!matchA || !matchB) return String(a).localeCompare(String(b));
//         return (matchA[1] !== matchB[1]) ? matchA[1].localeCompare(matchB[1]) : Number(matchB[2]) - Number(matchA[2]);
//       });

//       const maxRows = Math.max(leftRowKeys.length + 1, rightRowKeys.length + 4 + assortKeys.length + 2);

//       for (let i = 0; i < maxRows; i++) {
//         let row = new Array(fullHeader.length).fill("");

//         // 左側生成
//         if (i < leftRowKeys.length) {
//           const [client, area, course, shop, note] = leftRowKeys[i].split('|');
//           const pData = leftSummary[leftRowKeys[i]];
//           let rowTotal = 0;
//           const leftVals = [dStr, target.loc, client, area, course, shop, note];
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

//         // 右側（箱数・アソート）生成
//         if (i < rightRowKeys.length) {
//           const bType = rightRowKeys[i];
//           row[startBoxIdx] = bType;
//           row[startBoxIdx + 1] = rightSummary[bType] || 0;
//           row[startBoxIdx + 2] = bType * (rightSummary[bType] || 0);
//         } else if (i === rightRowKeys.length) {
//           row[startBoxIdx] = "総合計";
//           row[startBoxIdx + 1] = Object.values(rightSummary).reduce((a, b) => a + b, 0);
//           row[startBoxIdx + 2] = Object.keys(rightSummary).reduce((sum, t) => sum + (Number(t) * rightSummary[t]), 0);
//         }

//         const assortStartRow = rightRowKeys.length + 4;
//         if (i === assortStartRow - 1) {
//           row[startBoxIdx] = "アソート名";
//           row[startBoxIdx + 1] = "箱数";
//           row[startBoxIdx + 2] = "箱数合計";
//         } else if (i >= assortStartRow && i < assortStartRow + assortKeys.length) {
//           const assort = assortKeys[i - assortStartRow];
//           row[startBoxIdx] = assort;
//           row[startBoxIdx + 1] = row[startBoxIdx + 2] = assortSummary[assort];
//         } else if (i === assortStartRow + assortKeys.length) {
//           row[startBoxIdx] = "総合計";
//           row[startBoxIdx + 1] = row[startBoxIdx + 2] = Object.values(assortSummary).reduce((a, b) => a + b, 0);
//         }
//         allOutputRows.push(row);
//       }
//       allOutputRows.push(new Array(fullHeader.length).fill(""));
//       allOutputRows.push(fullHeader);
//     });

//     if (allOutputRows.length > 0) allOutputRows.pop();
//     const finalData = [fullHeader, ...allOutputRows];
//     targetSheet.getRange(1, 1, finalData.length, fullHeader.length).setValues(finalData);

//     // 装飾
//     const range = targetSheet.getRange(1, 1, finalData.length, fullHeader.length);
//     range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
//          .setFontSize(10).setVerticalAlignment("middle");
//     targetSheet.autoResizeColumns(1, fullHeader.length);
//     targetSheet.setColumnWidth(6, 150);
//     targetSheet.setColumnWidth(7, 120);
//   });
// }

// /**
//  * 2. レイアウト整形
//  */
// function applyTableLayout555() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   ["クロス表YT", "クロス表ET"].forEach(sheetName => {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) return;
//     const lastRow = sheet.getLastRow();
//     const lastCol = sheet.getLastColumn();
//     if (lastRow < 1) return;

//     const allRange = sheet.getRange(1, 1, lastRow, lastCol);
//     allRange.setBorder(false, false, false, false, false, false).setBackground(null).setFontWeight("normal").setHorizontalAlignment("left");

//     const data = allRange.getValues();
//     data.forEach((row, i) => {
//       const rIdx = i + 1;
//       if (row.slice(0, 12).join("").trim() !== "") {
//         const leftRange = sheet.getRange(rIdx, 1, 1, 12);
//         leftRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
//         if (row[0] === "納品日") leftRange.setBackground("#fce4ec").setFontWeight("bold").setHorizontalAlignment("center");
//         if (row[0] === "日付計") leftRange.setBackground("#f3f3f3").setFontWeight("bold");
//       }
//       const boxStartCol = 15;
//       if (row.slice(boxStartCol - 1, boxStartCol + 2).join("").trim() !== "") {
//         const rightRange = sheet.getRange(rIdx, boxStartCol, 1, 3);
//         rightRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
//         if (["箱タイプ", "アソート名", "総合計"].includes(row[boxStartCol - 1])) {
//           rightRange.setBackground("#e1f5fe").setFontWeight("bold");
//           if (row[boxStartCol - 1] !== "総合計") rightRange.setHorizontalAlignment("center");
//         }
//       }
//     });
//     sheet.setColumnWidths(1, 12, 60);
//     sheet.setColumnWidth(6, 120);
//     sheet.setColumnWidths(13, 2, 20);
//     sheet.setColumnWidths(15, 3, 60);
//   });
// }

// /**
//  * 3. ブロックタイトル挿入と列整理
//  */
// function formatCrossTablesByBlockFinal555() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   ['クロス表YT', 'クロス表ET'].forEach(name => {
//     const sheet = ss.getSheetByName(name);
//     if (!sheet) return;
//     const lastRow = sheet.getLastRow();
//     if (lastRow < 1) return;

//     const data = sheet.getRange(1, 1, lastRow, 2).getValues();
//     let insertPositions = [];

//     for (let i = 0; i < data.length; i++) {
//       if (String(data[i][0]).trim() === "納品日") {
//         let targetRowIdx = i + 1;
//         if (targetRowIdx < data.length) {
//           let dateVal = data[targetRowIdx][0];
//           let locVal = data[targetRowIdx][1];
//           let dateStr = "";
//           if (dateVal instanceof Date) {
//             dateStr = Utilities.formatDate(dateVal, "JST", "M/d") + "日月火水木金土"[dateVal.getDay()];
//           } else {
//             dateStr = String(dateVal);
//           }
//           insertPositions.push({ rowIndex: i + 1, title: `${locVal}  ${dateStr}` });
//         }
//       }
//     }

//     for (let j = insertPositions.length - 1; j >= 0; j--) {
//       sheet.insertRowBefore(insertPositions[j].rowIndex);
//       sheet.getRange(insertPositions[j].rowIndex, 3).setValue(insertPositions[j].title).setFontWeight("bold");
//     }
//     sheet.deleteColumns(1, 2);
//     sheet.setColumnWidth(1, 150);
//   });
// }
