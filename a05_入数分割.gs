function splitRowsByQuantityAdvanced2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('フリー入力用');
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // A列〜O列（15列分）を取得
  const range = sheet.getRange(1, 1, lastRow, 15);
  const values = range.getValues();
  
  const data = values.slice(1);
  const newValues = [];

  data.forEach(row => {
    let quantity = Number(row[6]); // G列: 数量
    
    if (isNaN(quantity) || quantity <= 0) {
      newValues.push(row);
      return;
    }

    // 分解後の数字を格納する配列
    let parts = [];

    // --- 分解ロジックの優先順位 ---
    
    // パターン1: 50と40の組み合わせで割り切れるか（90, 130, 140, 180など）
    let tempParts = findCombination(quantity, 50, 40);
    
    // パターン2: 上記がダメなら 40と30の組み合わせ（70, 110, 150など）
    if (tempParts.length === 0) {
      tempParts = findCombination(quantity, 40, 30);
    }

    // いずれかの組み合わせで見つかった場合
    if (tempParts.length > 0) {
      parts = tempParts;
    } else {
      // どちらの組み合わせでも分解できない（または50, 40, 30単体で割り切れる）場合
      if (quantity % 50 === 0) {
        parts = Array(quantity / 50).fill(50);
      } else if (quantity % 40 === 0) {
        parts = Array(quantity / 40).fill(40);
      } else if (quantity % 30 === 0) {
        parts = Array(quantity / 30).fill(30);
      } else {
        // 全く当てはまらない場合は分解しない
        parts = [quantity];
      }
    }

    // 分解されたパーツ分だけ行を複製
    parts.forEach(p => {
      const newRow = [...row];
      newRow[6] = p; // G列に分解した数値をセット
      newValues.push(newRow);
    });
  });

  // 書き込み処理
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).clearContent();
  }
  if (newValues.length > 0) {
    sheet.getRange(2, 1, newValues.length, 15).setValues(newValues);
  }
}

/**
 * 二つの数値の組み合わせで合計値が作れるか計算する補助関数
 */
function findCombination(total, num1, num2) {
  // num1をできるだけ多く使うパターンから試行
  for (let i = Math.floor(total / num1); i >= 0; i--) {
    let remainder = total - (i * num1);
    if (remainder >= 0 && remainder % num2 === 0) {
      let res1 = Array(i).fill(num1);
      let res2 = Array(remainder / num2).fill(num2);
      return res1.concat(res2);
    }
  }
  return [];
}




