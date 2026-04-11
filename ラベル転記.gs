function transferToLabels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shk = ss.getSheetByName('ラベル集計');
  const label = ss.getSheetByName('ラベル');

  if (!shk || !label) {
    throw new Error('シート名「ラベル集計」または「ラベル」が見つかりません。');
  }

  // 1. 転記元データの取得 (A2:I列)
  const lastRow = shk.getLastRow();
  if (lastRow < 2) return;
  const data = shk.getRange(2, 1, lastRow - 1, 9).getValues();

  // 2. 転記先シートの既存データをクリア（必要に応じて）
  // 1行目の項目名を残すなら A1:Z は避けて A2:Z などに調整してください
  label.clearContents(); 

  // 3. データの転記ループ
  for (let i = 0; i < data.length; i++) {
    const rowData = data[i];
    
    // shkの列定義 (0始まりのインデックス)
    const deliveryDate = rowData[0]; // A列:納品日
    const routeNum     = rowData[1]; // B列:コース番号
    const clientName   = rowData[2]; // C列:クライアント名
    const shopName     = rowData[3]; // D列:店名
    const productName  = rowData[4]; // E列:製品名
    const remarks      = rowData[5]; // F列:備考
    const countDisplay = rowData[8]; // I列:入数表示

    // 転記先の位置計算
    // i=0(1件目) → 左、i=1(2件目) → 右、i=2(3件目) → 左(下段)
    const isRightSide = i % 2 === 1; // 奇数番目が右側
    const rowOffset = Math.floor(i / 2) * 7; // 2件ごとに7行ずつ下にずらす

    // 列の起点 (左ならA列=1, 右ならD列=4)
    const colBase = isRightSide ? 4 : 1;
    // B列やE列（右隣の列）の起点
    const colSub = colBase + 1;

    // 4. 指定のセルに書き込み
    // A2 / D2 にコース番号
    label.getRange(2 + rowOffset, colBase).setValue(routeNum);
    
    // B1 / E1 に納品日
    label.getRange(1 + rowOffset, colSub).setValue(deliveryDate);
    
    // B2 / E2 にクライアント名
    label.getRange(2 + rowOffset, colSub).setValue(clientName);
    
    // B3 / E3 に店名
    label.getRange(3 + rowOffset, colSub).setValue(shopName);
    
    // B4 / E4 に製品名
    label.getRange(4 + rowOffset, colSub).setValue(productName);
    
    // B5 / E5 に備考
    label.getRange(5 + rowOffset, colSub).setValue(remarks);
    
    // B6 / E6 に入数表示
    label.getRange(6 + rowOffset, colSub).setValue(countDisplay);
  }

  console.log('ラベルシートへの転記が完了しました。');
}

