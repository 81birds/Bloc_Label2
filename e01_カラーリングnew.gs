// // =============================================================================
// //  定数・設定
// // =============================================================================

// /** エリア名 → 背景色 */
// const DISTRICT_COLORS = {
//   '旭川地区':    '#FFF9C4',
//   '函館地区':    '#979fe3',
//   '直納分':      '#add8e6',
//   '室蘭地区':    '#d3d3d3',
//   '苫小牧地区':  '#d3d3d3',
//   'ラルズアークス': '#98fb98',
// };

// /** ABCDアソートコード → 文字色 */
// const ASSORT_FONT_COLORS = {
//   'A': '#FF0000',
//   'B': '#0000FF',
//   'C': '#ffa500',
//   'D': '#008000',
// };

// /** 制作場所キーワード → 背景色・文字色 */
// const LABEL_PLACE_COLORS = {
//   'エルブ奥': '#ffc0cb',
//   '豊ビル':   '#000080',
// };

// /** ページ区切りの背景色 */
// const PAGE_DIVIDER_COLOR = '#a9a9a9';

// /** A3ラベルシート名 */
// const A3_SHEETS = ['ラベルA3_ゆ', 'ラベルA3_エ'];


// // =============================================================================
// //  メイン関数
// // =============================================================================

// /**
//  * 全シートの背景色・文字色設定をまとめて実行する
//  */
// function coloring4AllSheets() {
//   colorizeFreeInputSheet();
//   colorizeOtherSheetsFixed();
//   colorizeLabelA3Sheets();
//   setFontColorByCode();
//   setPageDividers();
//   colorizeLabels();
//   applyManualColors();
// }


// // =============================================================================
// //  1. フリー入力用シート
// // =============================================================================

// function colorizeFreeInputSheet() {
//   const sheet = getSheet('フリー入力用');
//   if (!sheet) return;

//   const lastRow = sheet.getLastRow();
//   if (lastRow < 2) return;

//   // D列(4)・L列(12) をまとめて塗る
//   const numRows = lastRow - 1;
//   const values  = sheet.getRange(2, 1, numRows, 12).getValues();

//   const colD_bg = [];
//   const colL_bg = [];

//   for (const row of values) {
//     const color = getDistrictColor(row[11]); // L列
//     colD_bg.push([color]);
//     colL_bg.push([color]);
//   }

//   sheet.getRange(2, 4,  numRows, 1).setBackgrounds(colD_bg);
//   sheet.getRange(2, 12, numRows, 1).setBackgrounds(colL_bg);
// }


// // =============================================================================
// //  2. ラベル集計・クロス表YT/ET シート
// // =============================================================================

// function colorizeOtherSheetsFixed() {
//   _colorizeDistrictSheet('ラベル集計', { dataWidth: 11, districtCol: 10, paintCols: [3, 10] });

//   for (const name of ['クロス表YT', 'クロス表ET']) {
//     _colorizeDistrictSheet(name, { dataWidth: 6, districtCol: 3, paintCols: [3, 5] });
//   }
// }

// /**
//  * エリア色をまとめて塗るユーティリティ
//  * @param {string} sheetName
//  * @param {{ dataWidth: number, districtCol: number, paintCols: number[] }} opts
//  *   districtCol: 0-based index of the column used for color lookup
//  *   paintCols:   0-based indexes of columns to paint
//  */
// function _colorizeDistrictSheet(sheetName, { dataWidth, districtCol, paintCols }) {
//   const sheet = getSheet(sheetName);
//   if (!sheet) return;

//   const lastRow = sheet.getLastRow();
//   if (lastRow < 2) return;

//   const numRows = lastRow - 1;
//   const values  = sheet.getRange(2, 1, numRows, dataWidth).getValues();

//   // 塗る列ごとに背景色配列を作成
//   const bgArrays = paintCols.map(() => []);

//   for (const row of values) {
//     const color = getDistrictColor(row[districtCol]);
//     paintCols.forEach((_, i) => bgArrays[i].push([color]));
//   }

//   // 1列ずつまとめて適用（1行1callで済む）
//   paintCols.forEach((col0, i) => {
//     sheet.getRange(2, col0 + 1, numRows, 1).setBackgrounds(bgArrays[i]);
//   });
// }


// // =============================================================================
// //  3. ラベルA3_ゆ / ラベルA3_エ シート — 背景色
// // =============================================================================

// function colorizeLabelA3Sheets() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();

//   for (const sheetName of A3_SHEETS) {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) continue;

//     const maxRows = sheet.getMaxRows();
//     const maxCols = sheet.getMaxColumns();

//     // シート全体をクリア（1回のAPIコール）
//     sheet.getRange(1, 1, maxRows, maxCols).setBackground(null);

//     const lastRow = sheet.getLastRow();
//     if (lastRow < 8) continue;

//     const labelCount = Math.ceil(lastRow / 12) * 2;

//     // 塗る位置を収集してからまとめて適用
//     // key = "row,col", value = color
//     const paintMap = {};

//     for (let i = 0; i < labelCount; i++) {
//       const isRight   = i % 2 === 1;
//       const rowOffset = Math.floor(i / 2) * 12;
//       const colA      = isRight ? 4 : 1;
//       const colB      = isRight ? 5 : 2;

//       const district = sheet.getRange(8 + rowOffset, colA).getValue();
//       const color    = getDistrictColor(district);
//       if (!color) continue;

//       // A2/A3（または D2/D3）と B2/B3（または E2/E3）を塗る
//       for (const r of [2 + rowOffset, 3 + rowOffset]) {
//         paintMap[`${r},${colA}`] = color;
//         paintMap[`${r},${colB}`] = color;
//       }
//     }

//     // 同色の連続セルをまとめて setBackground できると理想的だが、
//     // ラベルごとに色が変わる可能性があるため1セルずつ適用する
//     // （件数は多くないので許容範囲）
//     for (const [key, color] of Object.entries(paintMap)) {
//       const [r, c] = key.split(',').map(Number);
//       sheet.getRange(r, c).setBackground(color);
//     }
//   }
// }


// // =============================================================================
// //  4. ABCDアソート 文字色
// // =============================================================================

// function setFontColorByCode() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();

//   for (const sheetName of A3_SHEETS) {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) continue;

//     const lastRow = sheet.getLastRow();
//     if (lastRow < 1) continue;

//     // B列(2) と E列(5) を対象にする
//     for (const col of [2, 5]) {
//       const range  = sheet.getRange(1, col, lastRow, 1);
//       const values = range.getValues();

//       const fontColors   = values.map(([v]) => [ASSORT_FONT_COLORS[String(v).trim()] ?? null]);
//       const fontWeights  = values.map(([v]) => [ASSORT_FONT_COLORS[String(v).trim()] ? 'bold' : null]);

//       range.setFontColors(fontColors);
//       range.setFontWeights(fontWeights);
//     }
//   }
// }


// // =============================================================================
// //  5. ページ区切り
// // =============================================================================

// function setPageDividers() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();

//   for (const sheetName of A3_SHEETS) {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) continue;

//     const lastRow = sheet.getLastRow();
//     if (lastRow < 1) continue;

//     // 60行ごとにページ区切り行を配置
//     for (let row = 60, page = 1; row <= lastRow + 60; row += 60, page++) {
//       const range = sheet.getRange(row, 1, 1, 5);
//       range.setBackground(PAGE_DIVIDER_COLOR);
//       sheet.getRange(row, 1).setValue(`--- Page ${String(page).padStart(2, '0')} ---`);
//     }
//   }
// }


// // =============================================================================
// //  6. 制作場所の色付け（エルブ奥・豊ビル）
// // =============================================================================

// function colorizeLabels() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();

//   for (const sheetName of A3_SHEETS) {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) continue;

//     const lastRow = sheet.getLastRow();
//     if (lastRow < 1) continue;

//     // A列(1) と D列(4) を処理
//     for (const col of [1, 4]) {
//       const range      = sheet.getRange(1, col, lastRow);
//       const values     = range.getValues();
//       const backgrounds = range.getBackgrounds();
//       const fontColors  = range.getFontColors();

//       for (let i = 0; i < values.length; i++) {
//         const cell = String(values[i][0]);

//         const matchKey = Object.keys(LABEL_PLACE_COLORS).find(k => cell.includes(k));
//         if (!matchKey) continue;

//         const color = LABEL_PLACE_COLORS[matchKey];

//         // 2つ上のセルを白で隠す
//         if (i >= 2) {
//           backgrounds[i - 2][0] = '#ffffff';
//           fontColors[i - 2][0]  = '#ffffff';
//         }

//         backgrounds[i][0] = color;
//         fontColors[i][0]  = color;
//       }

//       range.setBackgrounds(backgrounds);
//       range.setFontColors(fontColors);
//     }
//   }
// }


// // =============================================================================
// //  7. クロス表 M列 アソートコード文字色
// // =============================================================================

// function applyManualColors() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();

//   for (const name of ['クロス表ET', 'クロス表YT']) {
//     const sheet = ss.getSheetByName(name);
//     if (!sheet) continue;

//     const lastRow = sheet.getLastRow();
//     if (lastRow < 2) return;

//     const range  = sheet.getRange(2, 13, lastRow - 1, 1); // M列
//     const values = range.getValues();

//     const colors = values.map(([v]) => {
//       const text = String(v);
//       if (text.startsWith('A')) return ['#FF0000'];
//       if (text.startsWith('B')) return ['#0000FF'];
//       if (text.startsWith('C')) return ['#a52a2a'];
//       if (text.startsWith('D')) return ['#008000'];
//       return [null];
//     });

//     range.setFontColors(colors);
//   }
// }


// // =============================================================================
// //  ユーティリティ
// // =============================================================================

// /**
//  * シートを名前で取得する（見つからなければ null）
//  */
// function getSheet(name) {
//   return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
// }

// /**
//  * エリア名から背景色を返す
//  * @param {*} district
//  * @returns {string|null}
//  */
// function getDistrictColor(district) {
//   if (!district) return null;
//   return DISTRICT_COLORS[String(district).trim()] ?? null;
// }