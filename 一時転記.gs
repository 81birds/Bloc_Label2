/**
 * MOTOシートからデータをスキャンし、SAKIシートに日付ごとに転記する
 */
function transferData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const motoSheet = ss.getSheetByName('元データ');
  const sakiSheet = ss.getSheetByName('フリー入力用');

  // MOTOシートの全データを配列として取得
  const motoValues = motoSheet.getDataRange().getValues();

  // 1. 日付(2行目・E列〜)と商品名(3行目・E列〜)を取得
  const dates = motoValues[1].slice(4); 
  const productNames = motoValues[2].slice(4);

  const outputData = [];

  // 2. 日付を起点にループ（1つの日付が終わったら次の日付へ）
  for (let j = 0; j < dates.length; j++) {
    const dateVal = dates[j];
    const productName = productNames[j]; // その列の商品名
    
    // 日付がなくなれば、その右側の列は処理せず終了
    if (!dateVal || dateVal === "") break;

    // その日付・商品に対して、全店舗（5行目以降）をスキャン
    for (let i = 4; i < motoValues.length; i++) {
      const shopName = motoValues[i][2]; // C列(インデックス2): 店名
      if (!shopName) continue; 

      const quantity = motoValues[i][4 + j]; // E列以降の数量
      
      // 配列の構造: [A, B(日付), C, D(店名), E, F(商品名), G(数量)]
      // インデックス: 0, 1, 2, 3, 4, 5, 6
      outputData.push(["", dateVal, "", shopName, "", productName, quantity]);
    }
  }

  // 3. SAKIシートに書き込み
  if (outputData.length > 0) {
    // 書き込み前にB2:Gの範囲をクリア
    const lastRow = sakiSheet.getLastRow();
    if (lastRow > 1) {
      sakiSheet.getRange("B2:G" + lastRow).clearContent();
    }

    // B2からoutputDataの列数(7列分)を一括書き込み
    // getRange(行, 列, 行数, 列数) -> (2, 1, 数, 7) で A列からG列までを対象にする
    sakiSheet.getRange(2, 1, outputData.length, 7).setValues(outputData);
    
    // SpreadsheetApp.getUi().alert("転記完了");
  } else {
    SpreadsheetApp.getUi().alert("転記対象のデータが見つかりませんでした。");
  }
}
