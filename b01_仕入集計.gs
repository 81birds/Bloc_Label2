function generatePivotSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("フリー入力用");
  const targetSheet = ss.getSheetByName("仕入集計");
  
  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) return; 

  const header = data[0];
  const rows = data.slice(1);
  
  // 列インデックスの取得
  const idxPlace = header.indexOf("制作場所");
  const idxDate = header.indexOf("仕入日");
  const idxAssort = header.indexOf("アソート");
  const idxQty = header.indexOf("数量");

  if (idxPlace === -1 || idxDate === -1 || idxAssort === -1 || idxQty === -1) {
    SpreadsheetApp.getUi().alert("列名が見つかりません。タイトル行を確認してください。");
    return;
  }

  let pivotData = {};
  let assortNames = new Set();
  let rowKeys = [];

  rows.forEach(row => {
    const place = row[idxPlace];
    const dateVal = row[idxDate];
    const assort = row[idxAssort];
    const qty = Number(row[idxQty]) || 0;
    
    if (!place || !dateVal || !assort) return; 

    // 日付のフォーマット
    const date = dateVal instanceof Date ? Utilities.formatDate(dateVal, "JST", "yyyy/MM/dd") : String(dateVal);
    
    // キーを「日付|場所」の順に変更
    const rowKey = `${date}|${place}`;
    
    if (!pivotData[rowKey]) {
      pivotData[rowKey] = {};
      rowKeys.push({date, place, key: rowKey});
    }
    
    pivotData[rowKey][assort] = (pivotData[rowKey][assort] || 0) + qty;
    assortNames.add(assort);
  });

  // 日付順に並べ替え
  rowKeys.sort((a, b) => new Date(a.date) - new Date(b.date));

  const sortedAssorts = Array.from(assortNames).sort();
  
  // ヘッダーを「仕入日」「制作場所」の順に設定
  let output = [["仕入日", "制作場所", ...sortedAssorts, "行合計"]];
  
  rowKeys.forEach(r => {
    let rowLine = [r.date, r.place]; // ここを入れ替え
    let rowTotal = 0;
    sortedAssorts.forEach(assort => {
      const val = pivotData[r.key][assort] || 0;
      rowLine.push(val);
      rowTotal += val;
    });
    rowLine.push(rowTotal);
    output.push(rowLine);
  });

  targetSheet.clear(); 
  if (output.length > 0) {
    const rowCount = output.length;
    const colCount = output[0].length; // 正しい列数を取得
    
    targetSheet.getRange(1, 1, rowCount, colCount).setValues(output);
    
    // スタイル調整
    targetSheet.getRange(1, 1, 1, colCount).setFontWeight("bold").setBackground("#EFEFEF");
    targetSheet.getRange(2, 3, rowCount - 1, colCount - 2).setHorizontalAlignment("center");
    targetSheet.setFrozenRows(1);
  }
}
