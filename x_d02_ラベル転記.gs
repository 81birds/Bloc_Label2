// function transferToLabels2() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const shk = ss.getSheetByName('ラベル集計');
//   if (!shk) { console.error('「ラベル集計」シートが見つかりません'); return; }

//   const lastRow = shk.getLastRow();
//   const lastCol = shk.getLastColumn();
//   console.log('診断: ラベル集計の最終行は ' + lastRow + '、最終列は ' + lastCol);
  
//   if (lastRow < 2) { console.warn('データが2行目以降にありません'); return; }

//   // 全データを取得（K列まで含めるため、最低11列以上を想定）
//   const data = shk.getRange(2, 1, lastRow - 1, Math.max(lastCol, 11)).getValues();

//   const sheets = {
//     '豊ビル': ss.getSheetByName('ラベルA3_ゆ'),
//     'エルブ奥': ss.getSheetByName('ラベルA3_エ')
//   };

//   // 出力先クリア
//   for (let key in sheets) {
//     if (sheets[key]) {
//       sheets[key].clearContents();
//       console.log('診断: シート「' + key + '」をクリアしました');
//     } else {
//       console.error('エラー: シート「' + key + '」がこのスプレッドシート内に存在しません');
//     }
//   }

//   const counts = { '豊ビル': 0, 'エルブ奥': 0 };

//   for (let i = 0; i < data.length; i++) {
//     const rowData = data[i];
    
//     // J列(10番目): 制作場所
//     let rawLocation = rowData[9] || ""; 
//     let location = String(rawLocation).trim();

//     const targetSheet = sheets[location];
//     if (!targetSheet) {
//       continue;
//     }

//     // 転記データの割り当て
//     const deliveryDate = rowData[0]; // A列(納品日)
//     const routeNum     = rowData[1]; // B列（ルート）
//     const clientName   = rowData[2]; // C列（クライアント）
//     const shopName     = rowData[3]; // D列（店名）
//     const productName  = rowData[4]; // E列（製品名）
//     const remarks      = rowData[5]; // F列（アソート名）
//     const countDisplay = rowData[8]; // I列（入数表示）
//     const areaTag      = rowData[10]; // K列 (エリア) ★追加項目

//     const currentIdx = counts[location];
//     const isRightSide = currentIdx % 2 === 1; // 0,2,4...は左(1列目)、1,3,5...は右(4列目)
//     const rowOffset = Math.floor(currentIdx / 2) * 6;
//     const colBase = isRightSide ? 4 : 1; // 左なら1(A列)、右なら4(D列)
//     const colSub = colBase + 1;         // 左なら2(B列)、右なら5(E列)

//     // 書き込み（各ラベルの構成）
//     targetSheet.getRange(1 + rowOffset, colBase).setValue(deliveryDate);
//     targetSheet.getRange(1 + rowOffset, colSub).setValue(clientName);
//     targetSheet.getRange(2 + rowOffset, colBase).setValue(routeNum);
//     targetSheet.getRange(2 + rowOffset, colSub).setValue(shopName);
//     targetSheet.getRange(3 + rowOffset, colSub).setValue(productName);
    
//     // エリア項目の転記 ★追加（左ラベルならA4, A10... / 右ラベルならD4, D10...）
//     targetSheet.getRange(4 + rowOffset, colBase).setValue(areaTag);
    
//     targetSheet.getRange(4 + rowOffset, colSub).setValue(remarks);
//     targetSheet.getRange(5 + rowOffset, colSub).setValue(countDisplay);
//     targetSheet.getRange(5 + rowOffset, colBase).setValue(location);

//     counts[location]++;
//   }

//   console.log('最終結果: ', counts);
//   if (counts['豊ビル'] === 0 && counts['エルブ奥'] === 0) {
//     console.error('【重要】1件も転記されませんでした。J列の文字を確認してください。');
//   }
// }
