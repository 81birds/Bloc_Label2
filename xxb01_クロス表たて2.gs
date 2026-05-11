function runAllProcessesCrossTable2() {
  createVerticalProductCrossTables();
  applyTableLayout();
  formatCrossTablesByBlockFinal();
}

/**
 * 前半：クロス表（A, B, C, D固定）
 * 後半：箱数集計
 * さらに下：M列アソート集計
 */
function createVerticalProductCrossTables() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('フリー入力用');
  if (!sourceSheet) return;

  const data = sourceSheet.getDataRange().getValues();
  const header = data.shift();

  const col = {
    date: header.indexOf('納品日'),
    location: header.indexOf('制作場所'),
    product: header.indexOf('アソート'),
    client: header.indexOf('クライアント名'),
    area: header.indexOf('エリア'),
    course: header.indexOf('コース番号'),
    shop: header.indexOf('店名'),
    qty: header.indexOf('数量'),
    note: header.indexOf('備考'),
    assortM: 12 // M列
  };

  // 固定製品
  const fixedProducts = ['A', 'B', 'C', 'D'];

  // 固定箱タイプ
  const fixedBoxTypes = [50, 40, 30, 20, 10];

  const targets = [
    { loc: "豊ビル", sheetName: "クロス表YT" },
    { loc: "エルブ奥", sheetName: "クロス表ET" }
  ];

  targets.forEach(target => {

    const targetSheet = ss.getSheetByName(target.sheetName);
    if (!targetSheet) return;

    targetSheet.clear();

    const filteredData = data.filter(row =>
      row[col.location] === target.loc && row[col.date]
    );

    if (filteredData.length === 0) return;

    // ヘッダー
    const headerBase = [
      '納品日',
      '制作場所',
      'クライアント名',
      'エリア',
      'コース番号',
      '店名',
      '備考'
    ];

    const leftHeader = [...headerBase, ...fixedProducts, '合計'];

    const emptyCols = ["", ""];

    const rightHeader = ['箱タイプ', '箱数', '数量小計'];

    const fullHeader = [
      ...leftHeader,
      ...emptyCols,
      ...rightHeader
    ];

    const startBoxIdx = leftHeader.length + emptyCols.length;

    // 日付一覧
    const dates = [...new Set(filteredData.map(row =>
      Utilities.formatDate(
        new Date(row[col.date]),
        "JST",
        "yyyy年M月d日"
      )
    ))].sort();

    let allOutputRows = [];

    dates.forEach(dStr => {

      const dateRows = filteredData.filter(row =>
        Utilities.formatDate(
          new Date(row[col.date]),
          "JST",
          "yyyy年M月d日"
        ) === dStr
      );

      // =========================
      // 左側：製品集計
      // =========================

      const leftSummary = {};

      dateRows.forEach(row => {

        const key =
          `${row[col.client]}|` +
          `${row[col.area]}|` +
          `${row[col.course]}|` +
          `${row[col.shop]}|` +
          `${row[col.note]}`;

        if (!leftSummary[key]) {
          leftSummary[key] = {};
        }

        const pName = String(row[col.product]).trim();

        leftSummary[key][pName] =
          (leftSummary[key][pName] || 0)
          + Number(row[col.qty]);
      });

      // =========================
      // 右側：箱数集計
      // =========================

      const rightSummary = {};

      fixedBoxTypes.forEach(type => {
        rightSummary[type] = 0;
      });

      dateRows.forEach(row => {

        const bType = Number(row[col.qty]);

        if (rightSummary[bType] === undefined) {
          rightSummary[bType] = 0;
        }

        rightSummary[bType] += 1;
      });

      // =========================
      // 下段：M列アソート集計
      // =========================

      const assortSummary = {};

      dateRows.forEach(row => {

        const assort = String(row[col.assortM]).trim();

        if (!assort) return;

        assortSummary[assort] =
          (assortSummary[assort] || 0) + 1;
      });

      // =========================

      const leftRowKeys = Object.keys(leftSummary);

      const rightRowKeys =
        Object.keys(rightSummary)
        .map(Number)
        .sort((a, b) => b - a);

      const assortKeys = Object.keys(assortSummary).sort((a, b) => {

  const matchA = String(a).match(/^([A-Za-z]+)(\d+)$/);
  const matchB = String(b).match(/^([A-Za-z]+)(\d+)$/);

  if (!matchA || !matchB) {
    return String(a).localeCompare(String(b));
  }

  const letterA = matchA[1];
  const numA = Number(matchA[2]);

  const letterB = matchB[1];
  const numB = Number(matchB[2]);

  // A→B→C順
  if (letterA !== letterB) {
    return letterA.localeCompare(letterB);
  }

  // 同じAなら 50→40→30
  return numB - numA;
});

      const assortRowsNeeded =
        assortKeys.length + 2;

      const maxRows = Math.max(
        leftRowKeys.length + 1,
        rightRowKeys.length + 4 + assortRowsNeeded
      );

      // =========================
      // 行生成
      // =========================

      for (let i = 0; i < maxRows; i++) {

        let row = new Array(fullHeader.length).fill("");

        // ===================================
        // 左側：ABCD集計
        // ===================================

        if (i < leftRowKeys.length) {

          const [client, area, course, shop, note] =
            leftRowKeys[i].split('|');

          const pData = leftSummary[leftRowKeys[i]];

          let rowTotal = 0;

          const leftVals = [
            dStr,
            target.loc,
            client,
            area,
            course,
            shop,
            note
          ];

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

            const colSum =
              leftRowKeys.reduce((sum, k) =>
                sum + (leftSummary[k][p] || 0), 0);

            row[headerBase.length + idx] =
              colSum || "";

            dateTotal += colSum;
          });

          row[leftHeader.length - 1] = dateTotal;
        }

        // ===================================
        // 右側：箱数集計
        // ===================================

        if (i < rightRowKeys.length) {

          const bType = rightRowKeys[i];

          const bCount = rightSummary[bType];

          row[startBoxIdx] = bType;
          row[startBoxIdx + 1] = bCount || 0;
          row[startBoxIdx + 2] =
            bType * (bCount || 0);

        } else if (i === rightRowKeys.length) {

          row[startBoxIdx] = "総合計";

          const totalBoxes =
            Object.values(rightSummary)
            .reduce((a, b) => a + b, 0);

          const totalQty =
            Object.keys(rightSummary)
            .reduce((sum, type) =>
              sum + (
                Number(type)
                * rightSummary[type]
              ), 0);

          row[startBoxIdx + 1] = totalBoxes;
          row[startBoxIdx + 2] = totalQty;
        }

        // ===================================
        // M列アソート集計
        // ===================================

        const assortStartRow =
          rightRowKeys.length + 4;

        // タイトル行
        if (i === assortStartRow - 1) {

          row[startBoxIdx] = "アソート名";
          row[startBoxIdx + 1] = "箱数";
          row[startBoxIdx + 2] = "箱数合計";
        }

        // データ行
        if (
          i >= assortStartRow &&
          i < assortStartRow + assortKeys.length
        ) {

          const assort =
            assortKeys[i - assortStartRow];

          const count =
            assortSummary[assort];

          row[startBoxIdx] = assort;
          row[startBoxIdx + 1] = count;
          row[startBoxIdx + 2] = count;
        }

        // 総合計行
        if (
          i === assortStartRow + assortKeys.length
        ) {

          const total =
            Object.values(assortSummary)
            .reduce((a, b) => a + b, 0);

          row[startBoxIdx] = "総合計";
          row[startBoxIdx + 1] = total;
          row[startBoxIdx + 2] = total;
        }

        allOutputRows.push(row);
      }

      // 空行
      allOutputRows.push(
        new Array(fullHeader.length).fill("")
      );

      // 次ブロック用ヘッダー
      allOutputRows.push(fullHeader);
    });

    const finalData = [
      fullHeader,
      ...allOutputRows
    ];

    targetSheet
      .getRange(
        1,
        1,
        finalData.length,
        fullHeader.length
      )
      .setValues(finalData);

    // ===================================
    // 装飾
    // ===================================

    const range = targetSheet.getRange(
      1,
      1,
      finalData.length,
      fullHeader.length
    );

    range
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        '#000000',
        SpreadsheetApp.BorderStyle.SOLID
      )
      .setFontSize(10)
      .setVerticalAlignment("middle");

    targetSheet.autoResizeColumns(
      1,
      fullHeader.length
    );

    targetSheet.setColumnWidth(6, 150);
    targetSheet.setColumnWidth(7, 120);
  });
}

/**
 * レイアウト
 */
function applyTableLayout() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const targetSheetNames = [
    "クロス表YT",
    "クロス表ET"
  ];

  targetSheetNames.forEach(sheetName => {

    const sheet =
      ss.getSheetByName(sheetName);

    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 1) return;

    const allRange =
      sheet.getRange(1, 1, lastRow, lastCol);

    allRange
      .setBorder(false, false, false, false, false, false)
      .setBackground(null)
      .setFontWeight("normal")
      .setHorizontalAlignment("left");

    const data = allRange.getValues();

    data.forEach((row, i) => {

      const rIdx = i + 1;

      // 左
      const leftRange =
        sheet.getRange(rIdx, 1, 1, 12);

      const leftRowValues =
        row.slice(0, 12);

      if (leftRowValues.join("").trim() !== "") {

        leftRange.setBorder(
          true, true, true, true, true, true,
          "black",
          SpreadsheetApp.BorderStyle.SOLID
        );

        if (row[0] === "納品日") {

          leftRange
            .setBackground("#fce4ec")
            .setFontWeight("bold")
            .setHorizontalAlignment("center");
        }

        if (row[0] === "日付計") {

          leftRange
            .setBackground("#f3f3f3")
            .setFontWeight("bold");
        }
      }

      // 右
      const boxStartCol = 15;

      const rightRange =
        sheet.getRange(rIdx, boxStartCol, 1, 3);

      const rightRowValues =
        row.slice(boxStartCol - 1, boxStartCol + 2);

      if (rightRowValues.join("").trim() !== "") {

        rightRange.setBorder(
          true, true, true, true, true, true,
          "black",
          SpreadsheetApp.BorderStyle.SOLID
        );

        if (
          row[boxStartCol - 1] === "箱タイプ" ||
          row[boxStartCol - 1] === "アソート名"
        ) {

          rightRange
            .setBackground("#e1f5fe")
            .setFontWeight("bold")
            .setHorizontalAlignment("center");
        }

        if (
          row[boxStartCol - 1] === "総合計"
        ) {

          rightRange
            .setBackground("#e1f5fe")
            .setFontWeight("bold");
        }
      }
    });

    sheet.setColumnWidths(1, 12, 60);

    sheet.setColumnWidth(6, 120);

    sheet.setColumnWidths(13, 2, 20);

    sheet.setColumnWidths(15, 3, 60);
  });
}

function formatCrossTablesByBlockFinal() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNames = [
    'クロス表YT',
    'クロス表ET'
  ];

  sheetNames.forEach(name => {

    const sheet = ss.getSheetByName(name);

    if (!sheet) return;

    let lastRow = sheet.getLastRow();

    if (lastRow < 1) return;

    let data =
      sheet.getRange(1, 1, lastRow, 2)
      .getValues();

    let insertPositions = [];

    for (let i = 0; i < data.length; i++) {

      if (
        String(data[i][0]).trim()
        === "納品日"
      ) {

        let targetRowIdx = i + 1;

        if (targetRowIdx < data.length) {

          let dateVal =
            data[targetRowIdx][0];

          let locVal =
            data[targetRowIdx][1];

          let dateStr = "";

          if (dateVal !== "") {

            if (dateVal instanceof Date) {

              dateStr =
                Utilities.formatDate(
                  dateVal,
                  "JST",
                  "M/d"
                ) +
                "日月火水木金土"[dateVal.getDay()];

            } else {

              dateStr = String(dateVal);
            }
          }

          insertPositions.push({
            rowIndex: i + 1,
            title: `${locVal}  ${dateStr}`
          });
        }
      }
    }

    // 下から挿入
    for (
      let j = insertPositions.length - 1;
      j >= 0;
      j--
    ) {

      let pos = insertPositions[j];

      sheet.insertRowBefore(pos.rowIndex);

      sheet
        .getRange(pos.rowIndex, 3)
        .setValue(pos.title)
        .setFontWeight("bold");
    }

    // A,B削除
    sheet.deleteColumns(1, 2);

    sheet.setColumnWidth(1, 150);
  });
}