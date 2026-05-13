function transferToLabels3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shk = ss.getSheetByName('ラベル集計');
  if (!shk) return;

  const lastRow = shk.getLastRow();
  if (lastRow < 2) return;

  // 1. データを一括取得（11列分）
  const data = shk.getRange(2, 1, lastRow - 1, 11).getValues();

  const sheets = {
    '豊ビル': ss.getSheetByName('ラベルA3_ゆ'),
    'エルブ奥': ss.getSheetByName('ラベルA3_エ')
  };

  // 2. 初期化（クリア）とデータ格納用配列の準備
  const outputData = {};
  for (let key in sheets) {
    if (sheets[key]) {
      sheets[key].clearContents();
      outputData[key] = []; // 各シート用のデータを溜める場所
    }
  }

  const counts = { '豊ビル': 0, 'エルブ奥': 0 };

  // 3. データをループ処理
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    // 制作場所の取得（J列：インデックス9）
    const location = String(row[9] || "").trim();
    if (!sheets[location]) continue;

    // 各項目の割り当て（他と同じ場所で定義）
    const deliveryDate = row[0];  // A列
    const routeNum     = row[1];  // B列
    const clientName   = row[2];  // C列
    const shopName     = row[3];  // D列
    const productName  = row[4];  // E列
    const remarks      = row[5];  // F列
    const countDisplay = row[8];  // I列
    const areaTag      = row[10]; // K列

    const currentIdx = counts[location];
    const isRightSide = currentIdx % 2 === 1;
    const rowIndex = Math.floor(currentIdx / 2) * 12;

    // ラベル1枚分の2次元配列（12行 × 2列分）を作成
    // 初期値がすべて空文字列の配列を用意
    let labelValues = Array.from({length: 12}, () => ["", ""]);
    
    // 配列の相対位置（0から数えるので、行番号-1）にセット
    labelValues[0][0] = deliveryDate; // A1
    labelValues[0][1] = clientName;   // B1
    // labelValues[2][0] = routeNum;     // A3(ルートを表示する場合はオンにする）
    labelValues[2][1] = shopName;     // B3
    labelValues[5][1] = productName;  // B6
    labelValues[7][0] = areaTag;      // A8
    labelValues[7][1] = remarks;      // B8
    labelValues[9][0] = location;     // A10（制作場所）
    labelValues[9][1] = countDisplay; // B10

    // 書き込み処理：ここが高速化のポイント
    // A,B列（1,2列目）か D,E列（4,5列目）に一気にセット
    const startCol = isRightSide ? 4 : 1;
    sheets[location].getRange(rowIndex + 1, startCol, 12, 2).setValues(labelValues);

    counts[location]++;
  }

  console.log('高速転記完了: ', counts);
}
