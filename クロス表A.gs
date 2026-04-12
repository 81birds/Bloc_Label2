/**
 * カスタムクロス表作成：ヘッダー整形・罫線・太字装飾付き//
 */
function createFinalCustomTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('フリー入力用');
  const targetSheet = ss.getSheetByName('クロス表A');
  
  if (!sourceSheet || !targetSheet) return;

  const data = sourceSheet.getDataRange().getValues();
  const header = data.shift();
  const col = {
    date: header.indexOf('納品日'),
    product: header.indexOf('製品名'),
    note: header.indexOf('備考'),
    client: header.indexOf('クライアント名'),
    course: header.indexOf('コース番号'),
    shop: header.indexOf('店名'),
    qty: header.indexOf('数量')
  };

  const summary = {};      
  const clients = new Set(); 
  const columnMap = {};    

  // 1. 集計処理
  data.forEach(row => {
    const rawDate = row[col.date];
    const formattedDate = Utilities.formatDate(new Date(rawDate), "JST", "yyyy年M月d日");
    const rowKey = `${formattedDate}|${row[col.product]}|${row[col.note]}`;
    const colKey = `${row[col.client]}|${row[col.course]}|${row[col.shop]}`;
    const val = Number(row[col.qty]) || 0;

    if (!summary[rowKey]) summary[rowKey] = {};
    summary[rowKey][colKey] = (summary[rowKey][colKey] || 0) + val;
    
    clients.add(row[col.client]);
    columnMap[colKey] = { client: row[col.client], course: row[col.course], shop: row[col.shop] };
  });

  const sortedClients = Array.from(clients).sort();
  const sortedColKeys = Object.keys(columnMap).sort();

  // 2. ヘッダー作成（3行）
  // A3:C3にラベルを配置するため、1-2行目のA:Cは空欄
  let h1 = ['', '', ''], h2 = ['', '', ''], h3 = ['納品日', '製品名', '備考'];
  const clientTotalCols = []; // 合計列のインデックス記録用

  sortedClients.forEach(client => {
    const clientCols = sortedColKeys.filter(k => columnMap[k].client === client);
    clientCols.forEach(k => {
      h1.push(columnMap[k].client); h2.push(columnMap[k].course); h3.push(columnMap[k].shop);
    });
    // クライアント合計列：1行目に「XX合計」、2-3行目は空欄
    h1.push(client + "合計"); h2.push(""); h3.push("");
    clientTotalCols.push(h1.length - 1);
  });
  h1.push('総合計'); h2.push(''); h3.push('');
  const grandTotalColIdx = h1.length - 1;

  // 3. データ行と「合計」行の作成
  const outputRows = [];
  const subtotalRowIndices = []; // 「合計」行の行番号記録用
  const sortedRowKeys = Object.keys(summary).sort();
  
  let lastProcessedDate = "";
  let dateSubtotal = null;

  sortedRowKeys.forEach((rowKey, index) => {
    const [date, product, note] = rowKey.split('|');
    if (lastProcessedDate !== "" && lastProcessedDate !== date) {
      subtotalRowIndices.push(outputRows.length + 4); // +4はヘッダー3行分+次行
      outputRows.push(dateSubtotal);
    }
    if (lastProcessedDate !== date) {
      lastProcessedDate = date;
      dateSubtotal = ["合計", "", ""]; 
      for (let i = 0; i < h1.length - 3; i++) dateSubtotal.push(0);
    }

    const rowData = [date, product, note];
    let rowGrandTotal = 0;
    let currentColIdx = 3;

    sortedClients.forEach(client => {
      let clientSum = 0;
      const clientCols = sortedColKeys.filter(k => columnMap[k].client === client);
      clientCols.forEach(colKey => {
        const val = summary[rowKey][colKey] || 0;
        rowData.push(val);
        dateSubtotal[currentColIdx] += val;
        clientSum += val;
        currentColIdx++;
      });
      rowData.push(clientSum);
      dateSubtotal[currentColIdx] += clientSum;
      rowGrandTotal += clientSum;
      currentColIdx++;
    });
    rowData.push(rowGrandTotal);
    dateSubtotal[currentColIdx] += rowGrandTotal;
    outputRows.push(rowData);

    if (index === sortedRowKeys.length - 1) {
      subtotalRowIndices.push(outputRows.length + 4);
      outputRows.push(dateSubtotal);
    }
  });

  // 4. 書き出し
  targetSheet.clear();
  const finalOutput = [h1, h2, h3, ...outputRows];
  const fullRange = targetSheet.getRange(1, 1, finalOutput.length, h1.length);
  fullRange.setValues(finalOutput);

  // 5. 装飾（罫線・太字・色塗り）
  // 全体に細い実線の罫線
  fullRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  // 縦方向の合計列（クライアント合計と総合計）を太字に
  clientTotalCols.concat(grandTotalColIdx + 1).forEach(colIdx => {
    targetSheet.getRange(4, colIdx + 1, outputRows.length, 1).setFontWeight("bold");
  });

  // 横方向の合計行（「合計」行）を太字に
  subtotalRowIndices.forEach(rowIdx => {
    targetSheet.getRange(rowIdx, 1, 1, h1.length).setFontWeight("bold");
  });

  // 既存の色塗り実行
  //applyPivotBlockColoring(targetSheet, h1.length, finalOutput.length);
}
