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
