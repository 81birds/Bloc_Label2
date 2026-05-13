function transferShopName4Label() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 各シートの取得
  const srcSheet = ss.getSheetByName("コースDB");
  const destSheet = ss.getSheetByName("フリー入力用");
  if (!srcSheet || !destSheet) return;

  const srcLastRow = srcSheet.getLastRow();
  const destLastRow = destSheet.getLastRow();
  if (srcLastRow < 2 || destLastRow < 2) return;

  // 2. コースDBのデータを配列に取得して連想配列（マップ）を作成
  const srcData = srcSheet.getRange(2, 3, srcLastRow - 1, 5).getValues(); // C列(3)からG列(7)まで
  const courseMap = {};
  
  for (let i = 0; i < srcData.length; i++) {
    const keyC = String(srcData[i][0]).trim(); // C列
    const valG = srcData[i][4];               // G列
    if (keyC !== "") {
      courseMap[keyC] = valG; // キーに対応するG列の値を保存
    }
  }

  // 3. フリー入力用のD列を取得してマッピング
  const destDData = destSheet.getRange(2, 4, destLastRow - 1, 1).getValues(); // D列(4)
  const resultN = [];

  for (let i = 0; i < destDData.length; i++) {
    const keyD = String(destDData[i][0]).trim(); // D列
    
    // キーが一致するものがあればその値を、なければ空欄をセット
    if (courseMap.hasOwnProperty(keyD)) {
      resultN.push([courseMap[keyD]]);
    } else {
      resultN.push([""]); 
    }
  }

  // 4. フリー入力用のN列(14列目)に一括書き込み
  destSheet.getRange(2, 14, resultN.length, 1).setValues(resultN);
}
