function modifySourceDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("フリー入力用");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // B列(日付)、L列(地区)のデータを取得
  const bRange = sheet.getRange(2, 2, lastRow - 1, 1);
  const bData = bRange.getValues();
  const lData = sheet.getRange(2, 12, lastRow - 1, 1).getValues();
  
  const resultsJ = []; // J列用
  const updatedB = []; // B列用(上書き用)
  const bBackgrounds = []; // B列背景用
  const weeks = ["日", "月", "火", "水", "木", "金", "土"];

  for (let i = 0; i < bData.length; i++) {
    let originalDate = new Date(bData[i]);
    let area = String(lData[i]).trim();
    
    if (isNaN(originalDate.getTime())) {
      resultsJ.push([""]);
      updatedB.push([bData[i][0]]); // そのまま
      bBackgrounds.push([null]);
      continue;
    }

    let dayNum = originalDate.getDay();
    let isModified = false;

    // --- 1. まずB列自体のマイナス1日判定 ---
    let bDate = new Date(originalDate);
    if (((dayNum === 4 || dayNum === 6) && area === "旭川地区") ||
        (dayNum === 2 && (area === "旭川地区" || area === "函館地区"))) {
      bDate.setDate(bDate.getDate() - 1);
      isModified = true;
    }
    updatedB.push([bDate]);
    bBackgrounds.push([isModified ? "#fff0f5" : null]);

    // --- 2. J列用の計算 (基本ルールに基づき、更新後のB列から算出) ---
    let targetDateJ = new Date(bDate);
    let newDayNum = bDate.getDay();

    if (newDayNum === 4 || newDayNum === 5 || newDayNum === 6 || newDayNum === 0) {
      let diff = (newDayNum === 0) ? 6 : newDayNum - 1;
      targetDateJ.setDate(bDate.getDate() - diff);
    } else {
      let diff = newDayNum + 2;
      targetDateJ.setDate(bDate.getDate() - diff);
    }

    let formattedDateJ = (targetDateJ.getMonth() + 1) + "月" + 
                         targetDateJ.getDate() + "日" + 
                         "(" + weeks[targetDateJ.getDay()] + ")";
    resultsJ.push([formattedDateJ]);
  }

  // B列(日付と背景色)を更新
  bRange.setValues(updatedB);
  bRange.setBackgrounds(bBackgrounds);

  // J列に結果を書き込み
  sheet.getRange(2, 10, resultsJ.length, 1).setValues(resultsJ);
}
