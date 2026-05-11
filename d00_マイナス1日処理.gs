function modifySourceDates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フリー入力用");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const bRange  = sheet.getRange(2, 2, numRows, 1);
  const bData   = bRange.getValues();
  const lData   = sheet.getRange(2, 12, numRows, 1).getValues();

  const updatedB     = [];
  const bBackgrounds = [];

  for (let i = 0; i < numRows; i++) {
    const date = new Date(bData[i][0]);
    const area = String(lData[i][0]).trim();

    // 無効な日付はそのまま
    if (isNaN(date.getTime())) {
      updatedB.push([bData[i][0]]);
      bBackgrounds.push([null]);
      continue;
    }

    const day = date.getDay(); // 0=日 1=月 2=火 3=水 4=木 5=金 6=土

    const needsMinus1 =
      (day === 4 && area === '旭川地区') ||  // 木曜 × 旭川
      (day === 6 && area === '旭川地区') ||  // 土曜 × 旭川
      (day === 2 && (area === '旭川地区' || area === '函館地区')); // 火曜 × 旭川・函館

    if (needsMinus1) date.setDate(date.getDate() - 1);

    updatedB.push([date]);
    bBackgrounds.push([needsMinus1 ? '#fff0f5' : null]);
  }
　bRange.setBackgrounds(bBackgrounds);
  bRange.setValues(updatedB);
  
}