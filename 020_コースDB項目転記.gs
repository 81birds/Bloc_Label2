function courseDBtoFreeNsheet() {
  ///console.log('--- 処理を開始します ---');

lookupCourseData2();//メイン転記
transferShopName4Label();//ラベル表示名転記


  ///console.log('--- すべての処理が正常に完了しました ---');
}




/**
 * コースDBからフリー入力用シートへデータを転記する
 * キー：フリー入力用 D列 ＝ コースDB C列
 * 追加：L列へコースDBのA列を転記
 */
function lookupCourseData2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const furiSheet = ss.getSheetByName('フリー入力用');
  const dbSheet = ss.getSheetByName('コースDB');

  if (!furiSheet || !dbSheet) return;

  // 1. コースDBのデータをMapに格納
  // A列(インデックス0)〜F列(インデックス5)を取得
  const dbLastRow = dbSheet.getLastRow();
  if (dbLastRow < 2) return;
  const dbValues = dbSheet.getRange("A2:F" + dbLastRow).getValues();
  const dbMap = new Map();

  dbValues.forEach(row => {
    const key = row[2]; // C列: キー
    if (key) {
      // keyに対して [D列, E列, F列, A列] の配列を紐付ける
      dbMap.set(key, [row[3], row[4], row[5], row[0]]);
    }
  });

  // 2. フリー入力用のデータを取得（A列〜L列まで扱うため12列分）
  const furiLastRow = furiSheet.getLastRow();
  if (furiLastRow < 2) return;

  // 書き込み範囲をA2:L(12列)に広げる
  const furiRange = furiSheet.getRange(2, 1, furiLastRow - 1, 12); 
  const furiValues = furiRange.getValues();

  // 3. データを突き合わせる
  for (let i = 0; i < furiValues.length; i++) {
    const searchKey = furiValues[i][3]; // D列 (インデックス3)
    
    if (dbMap.has(searchKey)) {
      const res = dbMap.get(searchKey);
      
      furiValues[i][2]  = res[0]; // C列 ← コースDB D列
      furiValues[i][4]  = res[1]; // E列 ← コースDB E列
      furiValues[i][8]  = res[2]; // I列 ← コースDB F列
      furiValues[i][11] = res[3]; // L列 (インデックス11) ← コースDB A列
    }
  }

  // 4. まとめてシートに書き戻す
  furiRange.setValues(furiValues);
}


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



