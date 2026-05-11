// /**
//  * 全てのシート（フリー入力用、ラベル集計、クロス表、ラベルA3）の
//  * 背景色設定を順番に実行するメイン関数
//  */
// function coloring4AllSheets() {
//   //console.log('背景色の更新を開始します...');
  
//   colorizeFreeInputSheet();       // 1. フリー入力用
//   colorizeOtherSheetsFixed();      // 2. ラベル集計、クロス表YT/ET
//   colorizeLabelA3Sheets();         // 3. ラベルA3_ゆ/エ
 
//   setFontColorByCode();//ABCDアソート色分け
//   // colorizeSupermarketBackground();//クライアント名色分け
//   setPageDividers();//ページ区切り（印刷設定用）
//   colorizeLabels();//制作場所色付けとエリアみえなくする
//   applyManualColors();



//   //console.log('全てのシートの更新が完了しました');
// }

// /**
//  * 1. 「フリー入力用」シートの背景色設定
//  */
// function colorizeFreeInputSheet() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getSheetByName('フリー入力用');
//   if (!sheet) return;
  
//   const lastRow = sheet.getLastRow();
//   if (lastRow < 2) return;

//   const range = sheet.getRange(2, 1, lastRow - 1, 12);
//   const values = range.getValues();
  
//   for (let i = 0; i < values.length; i++) {
//     const district = values[i][11]; // L列
//     const color = getDistrictColor2(district);
//     const rowNum = i + 2;

//     // D列とL列を塗る（条件に合わない場合は白/クリア）
//     sheet.getRange(rowNum, 4).setBackground(color);
//     sheet.getRange(rowNum, 12).setBackground(color);
//   }
// }

// /**
//  * 2. 「ラベル集計」「クロス表YT」「クロス表ET」の背景色設定
//  */
// function colorizeOtherSheetsFixed() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
  
//   // ラベル集計シート
//   const labelSheet = ss.getSheetByName('ラベル集計');
//   if (labelSheet) {
//     const lastRow = labelSheet.getLastRow();
//     if (lastRow >= 2) {
//       const data = labelSheet.getRange(2, 1, lastRow - 1, 11).getValues();
//       for (let i = 0; i < data.length; i++) {
//         const color = getDistrictColor2(data[i][10]); // K列
//         labelSheet.getRange(i + 2, 4).setBackground(color);  // D列
//         labelSheet.getRange(i + 2, 11).setBackground(color); // K列
//       }
//     }
//   }

//   // クロス表YT/ETシート
//   ['クロス表YT', 'クロス表ET'].forEach(name => {
//     const sheet = ss.getSheetByName(name);
//     if (!sheet) return;
//     const lastRow = sheet.getLastRow();
//     if (lastRow >= 2) {
//       const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
//       for (let i = 0; i < data.length; i++) {
//         const color = getDistrictColor2(data[i][3]); // D列
//         sheet.getRange(i + 2, 4).setBackground(color); // D列
//         sheet.getRange(i + 2, 6).setBackground(color); // F列
//       }
//     }
//   });
// }

// // 3. 「ラベルA3_ゆ」「ラベルA3_エ」の背景色設定
// function colorizeLabelA3Sheets() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const targetSheetNames = ['ラベルA3_ゆ', 'ラベルA3_エ'];

//   targetSheetNames.forEach(sheetName => {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) return;

//     // 1. シート全体の色をクリア
//     sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setBackground(null);

//     const lastRow = sheet.getLastRow();
//     if (lastRow < 8) return; 

//     // 2. ラベルの枚数を計算
//     const labelCount = Math.ceil(lastRow / 12) * 2;

//     for (let i = 0; i < labelCount; i++) {
//       const isRightSide = i % 2 === 1;
//       const rowOffset = Math.floor(i / 2) * 12;
      
//       const colA = isRightSide ? 4 : 1; // A列(1) または D列(4)
//       const colB = isRightSide ? 5 : 2; // B列(2) または E列(5)
      
//       // 判定に使用するエリア名のセル (A8, D8, A20, D20...)
//       const district = sheet.getRange(8 + rowOffset, colA).getValue();
//       const color = getDistrictColor2(district);

//       if (color) {
//         // ① エリア名のセル自体 (A8, D8等)
//        // sheet.getRange(8 + rowOffset, colA).setBackground(color);
        
//         // ② 最初のラベル例でいう B2 と B3 の位置に色を塗る
//         // rowOffset=0 の時、(2,2)と(3,2) ⇒ B2, B3
//         // rowOffset=12 の時、(14,2)と(15,2) ⇒ B14, B15
//         sheet.getRange(2 + rowOffset, colB).setBackground(color);
//         sheet.getRange(3 + rowOffset, colB).setBackground(color);
// 　　　　　
// 　　　　　// ② そのひとつ左側（A2, A3 / D2, D3）を追加
//         sheet.getRange(2 + rowOffset, colA).setBackground(color);
//         sheet.getRange(3 + rowOffset, colA).setBackground(color);

//       }
//     }
//   });
// }


// /**
//  * カラーコード共通判定ルール
//  */
// function getDistrictColor2(district) {
//   if (!district) return null;
//   const target = String(district).trim();
//   switch (target) {
//     case '旭川地区': return '#FFF9C4';
//     case '函館地区': return '#979fe3';
//     case '直納分':   return '#add8e6';
//     case '室蘭地区': return '#d3d3d3';
//     case '苫小牧地区': return '#d3d3d3';
// case 'ラルズアークス': return '#98fb98';

//     default: return null; 
//   }
// }


// /**
//  * ラベルシート（A3_ゆ / A3_エ）の文字色を設定
//  * 対象：B列とD列をスキャン
//  * ルール：A=赤, B=青, C=黄, D=緑
//  */
// function setFontColorByCode() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const targetSheetNames = ['ラベルA3_ゆ', 'ラベルA3_エ'];

//   targetSheetNames.forEach(sheetName => {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) return;

//     const lastRow = sheet.getLastRow();
//     if (lastRow < 1) return;

//     // B列(2)とE列(5)を含む範囲を取得
//     const range = sheet.getRange(1, 1, lastRow, 5);
//     const values = range.getValues();

//     for (let r = 0; r < values.length; r++) {
//       // B列(index 1) と E列(index 4) をスキャン
//       [1, 4].forEach(colIdx => {
//         const text = String(values[r][colIdx]).trim();
//         let fontColor = '#000000'; // 基本は黒

//         switch (text) {
//           case 'A': fontColor = '#FF0000'; break; // 赤
//           case 'B': fontColor = '#0000FF'; break; // 青
//           case 'C': fontColor = '#ffa500'; break; // 黄
//           case 'D': fontColor = '#008000'; break; // 緑
//           default: return; // A-D以外は何もしない
//         }

//         // 文字色を適用
//         sheet.getRange(r + 1, colIdx + 1).setFontColor(fontColor)
//              .setFontWeight('bold'); // はっきりさせるため太字に設定
//       });
//     }
//   });
// }





// /**
//  * ラベルシート（A3_ゆ / A3_エ）の特定のスーパー名に基づき背景色を設定する
//  * スキャン対象：B列とE列
//  */
// function colorizeSupermarketBackground() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const targetSheetNames = ['ラベルA3_ゆ', 'ラベルA3_エ'];

//   targetSheetNames.forEach(sheetName => {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) return;

//     const lastRow = sheet.getLastRow();
//     if (lastRow < 1) return;

//     // B列(2)からE列(5)までを含む範囲を取得
//     const range = sheet.getRange(1, 1, lastRow, 5);
//     const values = range.getValues();

//     for (let r = 0; r < values.length; r++) {
//       // B列(index 1) と E列(index 4) をチェック
//       [1, 4].forEach(colIdx => {
//         const text = String(values[r][colIdx]).trim();
//         let bgColor = null;

//         // スーパー名による条件判定（部分一致でも反応するように設定）
//         if (text.includes('コープ')) {
//           bgColor = '#E8F5E9'; // うすみどり (Pale Green)
//         } else if (text.includes('ラルズ')) {
//           bgColor = '#FFF3E0'; // 薄オレンジ (Light Orange)
//         } else if (text.includes('アークス')) {
//           bgColor = '#FBE9E7'; // うすいレンガいろ (Pale Deep Orange/Terra Cotta)
//         } else if (text.includes('サツドラ')) {
//           bgColor = '#E1F5FE'; // みずいろ (Light Blue)
//         } else if (text.includes('北海市場')) {
//           bgColor = '#ec8eaf'; // 薄い赤色 (Pale Red)
//         }

//         if (bgColor) {
//           // 背景色を適用
//           sheet.getRange(r + 1, colIdx + 1).setBackground(bgColor);
//         }
//       });
//     }
//   });
// }



// function setPageDividers() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const targetSheetNames = ['ラベルA3_ゆ', 'ラベルA3_エ'];
//   const grayColor = '#a9a9a9'; 

//   targetSheetNames.forEach(sheetName => {
//     const sheet = ss.getSheetByName(sheetName);
//     if (!sheet) return;

//     const lastRow = sheet.getLastRow();
//     if (lastRow < 1) return;

//     let pageNum = 1;
//     // 60行目から60行おきに処理（60, 120, 180...）
//     for (let row = 60; row <= lastRow + 60; row += 60) {
//       const range = sheet.getRange(row, 1, 1, 5); // A列〜E列
      
//       // ページ番号が出る行だけ色を塗る
//       range.setBackground(grayColor);
//       sheet.getRange(row, 1).setValue('--- Page ' + String(pageNum).padStart(2, '0') + ' ---');
      
//       pageNum++;
//     }
//   });
// }


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
//             backgrounds[i][0] = "#000080"; // 豊ビル：みどり
//             fontColors[i][0] = "#000080";
//           }
//         }
//       }
      
//       // まとめて反映
//       range.setBackgrounds(backgrounds);
//       range.setFontColors(fontColors);
//     });
//   });
// }
// function applyManualColors() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheetNames = ["クロス表ET", "クロス表YT"];
  
//   sheetNames.forEach(name => {
//     const sheet = ss.getSheetByName(name);
//     if (!sheet) return;

//     const lastRow = sheet.getLastRow();
//     if (lastRow < 2) return;

//     const range = sheet.getRange(2, 13, lastRow - 1, 1); // M列(13列目)
//     const values = range.getValues();
//     const colors = [];

//     for (let i = 0; i < values.length; i++) {
//       let text = String(values[i][0]);
//       let color = null; // デフォルト（自動/黒）

//       // 先頭の文字で判定
//       if (text.startsWith("A")) {
//         color = "#FF0000"; // 赤
//       } else if (text.startsWith("B")) {
//         color = "#0000FF"; // 青
//       } else if (text.startsWith("C")) {
//         color = "#a52a2a"; // 黄（見やすいように少し濃い金にしています）
//       } else if (text.startsWith("D")) {
//         color = "#008000"; // 緑
//       }
      
//       colors.push([color]);
//     }

//     // 文字色を一括適用（条件付き書式ではないので後で変更可能）
//     range.setFontColors(colors);
//   });
// }


