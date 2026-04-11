/**
 * コースDBからフリー入力用シートへデータを転記する
 * キー：フリー入力用 D列 ＝ コースDB C列
 */
function lookupCourseData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const furiSheet = ss.getSheetByName('フリー入力用');
  const dbSheet = ss.getSheetByName('コースDB');

  // 1. コースDBのデータをMap（辞書）に格納する
  // C列(キー), D列(転記1), E列(転記2), F列(転記3) を取得
  const dbValues = dbSheet.getRange("C2:F" + dbSheet.getLastRow()).getValues();
  const dbMap = new Map();

  dbValues.forEach(row => {
    const key = row[0]; // C列: キー
    if (key) {
      // keyに対して [D列, E列, F列] の配列を紐付ける
      dbMap.set(key, [row[1], row[2], row[3]]);
    }
  });

  // 2. フリー入力用のデータを取得（D列を基準にするためA〜I列まで取得）
  const furiLastRow = furiSheet.getLastRow();
  if (furiLastRow < 2) return; // データがなければ終了

  const furiRange = furiSheet.getRange(2, 1, furiLastRow - 1, 9); // A2:I
  const furiValues = furiRange.getValues();

  // 3. データを突き合わせる
  for (let i = 0; i < furiValues.length; i++) {
    const searchKey = furiValues[i][3]; // D列 (インデックス3)
    
    if (dbMap.has(searchKey)) {
      const res = dbMap.get(searchKey);
      
      furiValues[i][2] = res[0]; // C列 (インデックス2) ← コースDB D列
      furiValues[i][4] = res[1]; // E列 (インデックス4) ← コースDB E列
      furiValues[i][8] = res[2]; // I列 (インデックス8) ← コースDB F列
    }
  }

  // 4. まとめてシートに書き戻す
  furiRange.setValues(furiValues);
}
