// /**
//  * ラベルシートにページ区切りのグレー行とページ番号を挿入する
//  */
// function xxxsetPageDividers() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const targetSheetNames = ['ラベルA3_ゆ', 'ラベルA3_エ'];
//   const grayColor = '#a9a9a9'; // 

//   targetSheetNames.forEach(sheetName => {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) return;

//     let pageNum = 1;
//     // 30行目から1200行目まで、30行おきに処理
//     for (let row = 30; row <= 1200; row += 30) {
//       const range = sheet.getRange(row, 1, 1, 5); // A列〜E列
      
//       // 背景色をグレーにし、A列にページ番号を入力
//       range.setBackground(grayColor);
//       sheet.getRange(row, 1).setValue('page' + String(pageNum).padStart(3, '0'));
      
//       pageNum++;
//     }
//   });
// }
