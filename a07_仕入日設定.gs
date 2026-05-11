function convertAndFillDates() {
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
