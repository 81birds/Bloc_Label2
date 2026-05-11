// /**
//  * 指定シートのA列とD列をスキャンし、特定の文字に合わせて塗りつぶす
//  * 同時に、見つかったセルの「2つ上のセル」を白で塗りつぶす（結合セル対応）
//  */
// function colorizeLabels() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const targetSheets = ["ラベルA3_ゆ", "ラベルA3_エ"];
  
//   targetSheets.forEach(sheetName => {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) return;

//     const lastRow = sheet.getLastRow();
//     if (lastRow < 1) return;

//     const rangeA = sheet.getRange(1, 1, lastRow); // A列
//     const rangeD = sheet.getRange(1, 4, lastRow); // D列
    
//     const ranges = [rangeA, rangeD];

//     ranges.forEach(range => {
//       const values = range.getValues();
//       const backgrounds = range.getBackgrounds();
//       const fontColors = range.getFontColors();

//       for (let i = 0; i < values.length; i++) {
//         const cellValue = String(values[i][0]);

//         const isElb = cellValue.includes("エルブ奥");
//         const isToyo = cellValue.includes("豊ビル");

//         if (isElb || isToyo) {
//           // --- 2つ上のセルを白にする処理 ---
//           if (i >= 2) { // 3行目以降（インデックス2以降）であれば実行
//             backgrounds[i - 2][0] = "#ffffff";
//             fontColors[i - 2][0] = "#ffffff";
//           }

//           // --- 対象セルの色付け ---
//           if (isElb) {
//             backgrounds[i][0] = "#ffc0cb"; // エルブ奥：ピンク
//             fontColors[i][0] = "#ffc0cb";
//           } else if (isToyo) {
//             backgrounds[i][0] = "#93c47d"; // 豊ビル：みどり
//             fontColors[i][0] = "#93c47d";
//           }
//         }
//       }
      
//       // まとめて反映
//       range.setBackgrounds(backgrounds);
//       range.setFontColors(fontColors);
//     });
//   });
// }
