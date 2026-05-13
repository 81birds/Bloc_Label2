function runAllProcessesSHIRE() {

convertAndFillDates();//仕入日設定
generatePivotSummary();//仕入集計表

}


function convertAndFillDates() {//仕入日設定
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("フリー入力用");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // B列(2列目)の表示形式を「4月9日(水)」に一括変更
  const bRange = sheet.getRange(2, 2, lastRow - 1, 1);
  bRange.setNumberFormat('M"月"d"日"(ddd)');

  const bData = bRange.getValues();
  const results = [];
  const weeks = ["日", "月", "火", "水", "木", "金", "土"];

  for (let i = 0; i < bData.length; i++) {
    let date = new Date(bData[i]);
    
    if (isNaN(date.getTime())) {
      results.push([""]);
      continue;
    }

    let dayNum = date.getDay(); // 0:日, 1:月, 2:火, 3:水, 4:木, 5:金, 6:土
    let targetDate = new Date(date);

    // // 条件判定(まちがってる？)
    // if (dayNum === 4 || dayNum === 6 || dayNum === 0) {
    //   // 木(4), 土(6), 日(0) の場合 → 直前の月曜日
    //   let diff = (dayNum === 0) ? 6 : dayNum - 1;
    //   targetDate.setDate(date.getDate() - diff);
    // } else {
    //   // 月(1), 火(2), 水(3), 金(5) の場合 → 直前の金曜日
    //   let diff = (dayNum <= 3) ? dayNum + 2 : 7;
    //   targetDate.setDate(date.getDate() - diff);
    // }

   // 条件判定（修正版）
    if (dayNum === 4 || dayNum === 5 || dayNum === 6 || dayNum === 0) {
      // 木(4), 金(5), 土(6), 日(0) の場合 → 直前の月曜日
      let diff = (dayNum === 0) ? 6 : dayNum - 1;
      targetDate.setDate(date.getDate() - diff);
    } else {
      // 月(1), 火(2), 水(3) の場合 → 直前の金曜日
      let diff = dayNum + 2;
      targetDate.setDate(date.getDate() - diff);
    }



    // J列用のフォーマット作成
    let formattedDate = (targetDate.getMonth() + 1) + "月" + 
                        targetDate.getDate() + "日" + 
                        "(" + weeks[targetDate.getDay()] + ")";
    results.push([formattedDate]);
  }

  // J列(10列目)に一括書き込み
  sheet.getRange(2, 10, results.length, 1).setValues(results);
}



function generatePivotSummary() {//仕入集計表作成
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





