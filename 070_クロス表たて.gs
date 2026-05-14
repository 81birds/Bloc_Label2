/**
 * メイン実行関数
 */
function runAllProcessesCrossTable2555() {
  createVerticalProductCrossTables555();
  applyTableLayout555();
  formatCrossTablesByBlockFinal555();
}

/**
 * クライアント名がアークス・ラルズ系かどうか判定
 */
function isArcusRalse(clientName) {
  const name = String(clientName);
  return name.includes('アークス') || name.includes('ラルズ');
}

/**
 * 日付オブジェクトを「4/1 土」形式に変換
 */
function formatShortDate(dateObj) {
  const dayNames = ["日", "月", "火", "水", "木", "金", "土"];
  return `${dateObj.getMonth() + 1}/${dateObj.getDate()} ${dayNames[dateObj.getDay()]}`;
}

/**
 * アソート名（例: A50, C40）から数字を抽出
 */
function extractBoxSize(assortName) {
  const match = String(assortName).match(/(\d+)/);
  return match ? Number(match[1]) : 0;
}

/**
 * 1グループ分の出力行を生成（左表＋右2表）
 */
function buildGroupRows(groupData, leftSummary, dStr, locName, fixedProducts, headerBase, leftHeader, startBoxIdx, fullHeader, fixedBoxTypes) {
  const leftRowKeys = Object.keys(leftSummary);

  // 右側（箱タイプ）集計
  const rightSummary = {};
  fixedBoxTypes.forEach(type => { rightSummary[type] = 0; });
  groupData.forEach(row => {
    const bType = Number(row._qty);
    if (rightSummary[bType] !== undefined) rightSummary[bType] += 1;
  });

  // アソート集計
  const assortSummary = {};
  groupData.forEach(row => {
    const assort = String(row._assortM).trim();
    if (assort && assort !== 'undefined') assortSummary[assort] = (assortSummary[assort] || 0) + 1;
  });

  const rightRowKeys = Object.keys(rightSummary).map(Number).sort((a, b) => b - a);
  const assortKeys = Object.keys(assortSummary).sort((a, b) => {
    const matchA = String(a).match(/^([A-Za-z]+)(\d+)$/);
    const matchB = String(b).match(/^([A-Za-z]+)(\d+)$/);
    if (!matchA || !matchB) return String(a).localeCompare(String(b));
    return (matchA[1] !== matchB[1]) ? matchA[1].localeCompare(matchB[1]) : Number(matchB[2]) - Number(matchA[2]);
  });

  const assortStartRow = rightRowKeys.length + 4;
  const maxRows = Math.max(leftRowKeys.length + 1, assortStartRow + assortKeys.length + 1);

  const rows = [];
  for (let i = 0; i < maxRows; i++) {
    let row = new Array(fullHeader.length).fill("");

    // 左側
    if (i < leftRowKeys.length) {
      const [client, area, course, shop, note] = leftRowKeys[i].split('|');
      const pData = leftSummary[leftRowKeys[i]];
      let rowTotal = 0;
      const leftVals = [dStr, locName, client, area, course, shop, note];
      fixedProducts.forEach(p => {
        const val = pData[p] || 0;
        leftVals.push(val || "");
        rowTotal += val;
      });
      leftVals.push(rowTotal);
      row.splice(0, leftVals.length, ...leftVals);
    } else if (i === leftRowKeys.length) {
      row[0] = "日付計";
      let dateTotal = 0;
      fixedProducts.forEach((p, idx) => {
        const colSum = leftRowKeys.reduce((sum, k) => sum + (leftSummary[k][p] || 0), 0);
        row[headerBase.length + idx] = colSum || "";
        dateTotal += colSum;
      });
      row[leftHeader.length - 1] = dateTotal;
    }

    // 右側（箱タイプ）
    if (i < rightRowKeys.length) {
      const bType = rightRowKeys[i];
      row[startBoxIdx] = bType;
      row[startBoxIdx + 1] = rightSummary[bType] || 0;
      row[startBoxIdx + 2] = bType * (rightSummary[bType] || 0);
    } else if (i === rightRowKeys.length) {
      row[startBoxIdx] = "総合計";
      row[startBoxIdx + 1] = Object.values(rightSummary).reduce((a, b) => a + b, 0);
      row[startBoxIdx + 2] = Object.keys(rightSummary).reduce((sum, t) => sum + (Number(t) * rightSummary[t]), 0);
    }

    // アソート（ヘッダー行）
    if (i === assortStartRow - 1) {
      row[startBoxIdx]     = "アソート名";
      row[startBoxIdx + 1] = "箱数";
      row[startBoxIdx + 2] = "束数合計";  // 変更

    // アソート（データ行）：箱数合計 → アソート名から箱数を抽出して箱数×個数
    } else if (i >= assortStartRow && i < assortStartRow + assortKeys.length) {
      const assort = assortKeys[i - assortStartRow];
      const count = assortSummary[assort];
      const boxSize = extractBoxSize(assort);  // A50→50, C40→40
      row[startBoxIdx]     = assort;
      row[startBoxIdx + 1] = count;
      row[startBoxIdx + 2] = boxSize * count;  // 束数合計

    // アソート（総合計行）
    } else if (i === assortStartRow + assortKeys.length) {
      const totalCount = Object.values(assortSummary).reduce((a, b) => a + b, 0);
      const totalSoku = assortKeys.reduce((sum, assort) => {
        return sum + extractBoxSize(assort) * assortSummary[assort];
      }, 0);
      row[startBoxIdx]     = "総合計";
      row[startBoxIdx + 1] = totalCount;
      row[startBoxIdx + 2] = totalSoku;  // 束数合計の総計
    }

    rows.push(row);
  }
  return rows;
}

/**
 * 1. データ集計と書き出し
 */
function createVerticalProductCrossTables555() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('フリー入力用');
  if (!sourceSheet) return;

  const data = sourceSheet.getDataRange().getValues();
  const header = data.shift();

  // const col = {
  //   date: header.indexOf('納品日'),
  //   location: header.indexOf('制作場所'),
  //   product: header.indexOf('アソート'),
  //   client: header.indexOf('クライアント名'),
  //   area: header.indexOf('エリア'),
  //   course: header.indexOf('コース番号'),
  //   shop: header.indexOf('店名'),
  //   qty: header.indexOf('数量'),
  //   note: header.indexOf('備考'),
  //   assortM: 12
  // };


  const col = {
  date:     header.indexOf('納品日'),
  location: header.indexOf('制作場所'),
  product:  header.indexOf('アソート'),
  client:   header.indexOf('クライアント名'),
  area:     header.indexOf('エリア'),
  course:   header.indexOf('コース番号'),
  shop:     header.indexOf('店名'),
  qty:      header.indexOf('数量'),
  note:     header.indexOf('備考'),
  assortM:  header.indexOf('アソート数量名')  // ← ハードコードをindexOfに修正
};










  const fixedProducts = ['A', 'B', 'C', 'D'];
  const fixedBoxTypes = [50, 40, 30, 20, 15, 10];
  const targets = [
    { loc: "豊ビル", sheetName: "クロス表YT" },
    { loc: "エルブ奥", sheetName: "クロス表ET" }
  ];

  const headerBase = ['納品日', '制作場所', 'クライアント名', 'エリア', 'コース番号', '店名', '備考'];
  const leftHeader = [...headerBase, ...fixedProducts, '合計'];
  const emptyCols = ["", ""];
  const rightHeader = ['箱タイプ', '箱数', '数量小計'];
  const fullHeader = [...leftHeader, ...emptyCols, ...rightHeader];
  const startBoxIdx = leftHeader.length + emptyCols.length;

  targets.forEach(target => {
    const targetSheet = ss.getSheetByName(target.sheetName);
    if (!targetSheet) return;
    targetSheet.clear();

    const filteredData = data.filter(row =>
      row[col.location] === target.loc && row[col.date]
    ).map(row => ({
      _date: row[col.date],
      _location: row[col.location],
      _product: String(row[col.product]).trim(),
      _client: row[col.client],
      _area: row[col.area],
      _course: row[col.course],
      _shop: row[col.shop],
      _qty: row[col.qty],
      _note: row[col.note],
      _assortM: row[col.assortM]
    }));

    if (filteredData.length === 0) return;

    const dates = [...new Set(filteredData.map(row =>
      Utilities.formatDate(new Date(row._date), "JST", "yyyy年M月d日")
    ))].sort();

    let allOutputRows = [];

    dates.forEach(dStr => {
      const dateRows = filteredData.filter(row =>
        Utilities.formatDate(new Date(row._date), "JST", "yyyy年M月d日") === dStr
      );

      const dateObj = new Date(dateRows[0]._date);
      const shortDate = formatShortDate(dateObj);

      const groups = [
        { label: "アークス・ラルズ", rows: dateRows.filter(r => isArcusRalse(r._client)) },
        { label: "コープ",           rows: dateRows.filter(r => !isArcusRalse(r._client)) }
      ];

      groups.forEach(group => {
        if (group.rows.length === 0) return;

        // 空行
        allOutputRows.push(new Array(fullHeader.length).fill(""));

        // タイトル行：C列（index=2）に入れる（後で1・2列目を削除するためA列になる）
        const titleRow = new Array(fullHeader.length).fill("");
        titleRow[2] = `${target.loc}　${shortDate}　${group.label}`;
        allOutputRows.push(titleRow);

        // ヘッダー行
        allOutputRows.push(fullHeader);

        // 左側集計
        const leftSummary = {};
        group.rows.forEach(row => {
          const key = `${row._client}|${row._area}|${row._course}|${row._shop}|${row._note}`;
          if (!leftSummary[key]) leftSummary[key] = {};
          leftSummary[key][row._product] = (leftSummary[key][row._product] || 0) + Number(row._qty);
        });

        // グループ行を生成して追加
        const groupRows = buildGroupRows(
          group.rows, leftSummary, dStr, target.loc,
          fixedProducts, headerBase, leftHeader, startBoxIdx, fullHeader, fixedBoxTypes
        );
        groupRows.forEach(r => allOutputRows.push(r));
      });
    });

    if (allOutputRows.length === 0) return;

    // 先頭の余分な空行を除去
    while (allOutputRows.length > 0 && allOutputRows[0].join("").trim() === "") {
      allOutputRows.shift();
    }

    targetSheet.getRange(1, 1, allOutputRows.length, fullHeader.length).setValues(allOutputRows);

    targetSheet.getRange(1, 1, allOutputRows.length, fullHeader.length)
      .setFontSize(10).setVerticalAlignment("middle");
    targetSheet.autoResizeColumns(1, fullHeader.length);
    targetSheet.setColumnWidth(6, 150);
    targetSheet.setColumnWidth(7, 120);
  });
}

/**
 * 2. レイアウト整形
 */
function applyTableLayout555() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ["クロス表YT", "クロス表ET"].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1) return;

    const allRange = sheet.getRange(1, 1, lastRow, lastCol);
    allRange.setBorder(false, false, false, false, false, false)
            .setBackground(null).setFontWeight("normal").setHorizontalAlignment("left");

    const data = allRange.getValues();
    data.forEach((row, i) => {
      const rIdx = i + 1;
      const cellA = String(row[0]).trim();
      const cellC = String(row[2]).trim();

      // タイトル行（C列に制作場所名を含む）
      if (cellC.includes('豊ビル') || cellC.includes('エルブ奥')) {
        sheet.getRange(rIdx, 1, 1, lastCol)
          .setBackground("#F0F0F0")
          .setFontWeight("bold")
          .setFontSize(12);
        return;
      }

      // 左側テーブル
      if (row.slice(0, 12).join("").trim() !== "") {
        const leftRange = sheet.getRange(rIdx, 1, 1, 12);
        leftRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
        if (cellA === "納品日") leftRange.setBackground("#fce4ec").setFontWeight("bold").setHorizontalAlignment("center");
        if (cellA === "日付計") leftRange.setBackground("#f3f3f3").setFontWeight("bold");
      }

      // 右側テーブル
      const boxStartCol = 15;
      if (row.slice(boxStartCol - 1, boxStartCol + 2).join("").trim() !== "") {
        const rightRange = sheet.getRange(rIdx, boxStartCol, 1, 3);
        rightRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
        if (["箱タイプ", "アソート名", "総合計"].includes(row[boxStartCol - 1])) {
          rightRange.setBackground("#e1f5fe").setFontWeight("bold");
          if (row[boxStartCol - 1] !== "総合計") rightRange.setHorizontalAlignment("center");
        }
      }
    });

    sheet.setColumnWidths(1, 12, 60);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidths(13, 2, 20);
    sheet.setColumnWidths(15, 3, 60);
  });
}

/**
 * 3. 列整理（1・2列目削除後、タイトルがA列に来る）
 */
function formatCrossTablesByBlockFinal555() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['クロス表YT', 'クロス表ET'].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return;

    sheet.deleteColumns(1, 2);
    sheet.setColumnWidth(1, 150);
  });
}