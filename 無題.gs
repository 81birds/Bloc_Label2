// function setColumnMFormat() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheetNames = ["クロス表ET", "クロス表YT"];
  
//   // 設定する条件と色のペア (A:赤, B:青, C:黄, D:緑)
//   const rules = [
//     { text: "A", color: "#FF0000" }, // 赤
//     { text: "B", color: "#0000FF" }, // 青
//     { text: "C", color: "#FFFF00" }, // 黄
//     { text: "D", color: "#008000" }  // 緑
//   ];

//   sheetNames.forEach(name => {
//     const sheet = ss.getSheetByName(name);
//     if (!sheet) return;

//     const range = sheet.getRange("M2:M" + sheet.getMaxRows());
//     const newRules = [];

//     rules.forEach(r => {
//       const rule = SpreadsheetApp.newConditionalFormatRule()
//         .whenTextStartsWith(r.text)
//         .setFontColor(r.color)
//         .setRanges([range])
//         .build();
//       newRules.push(rule);
//     });

//     // 既存の条件付き書式に今回のルールを追加
//     const currentRules = sheet.getConditionalFormatRules();
//     sheet.setConditionalFormatRules(currentRules.concat(newRules));
//   });
// }
