// =============================================================================
//  定数・設定
// =============================================================================

/** エリア名 → 背景色 */
const DISTRICT_COLORS = {
  '旭川地区':    '#FFF9C4',
  '函館地区':    '#979fe3',
  '直納分':      '#add8e6',
  '室蘭地区':    '#d3d3d3',
  '苫小牧地区':  '#d3d3d3',
  'ラルズアークス': '#98fb98',
};

/** ABCDアソートコード → 文字色 */
const ASSORT_FONT_COLORS = {
  'A': '#FF0000',
  'B': '#0000FF',
  'C': '#ffa500',
  'D': '#008000',
};

/** 制作場所キーワード → 背景色・文字色 */
const LABEL_PLACE_COLORS = {
  'エルブ奥': '#ffc0cb',
  '豊ビル':   '#000080',
};

/** ページ区切りの背景色 */
const PAGE_DIVIDER_COLOR = '#a9a9a9';

/** A3ラベルシート名 */
const A3_SHEETS = ['ラベルA3_ゆ', 'ラベルA3_エ'];


// =============================================================================
//  メイン関数
// =============================================================================

/**
 * 全シートの背景色・文字色設定をまとめて実行する
 */
function coloring4AllSheets() {
  colorizeFreeInputSheet();
  colorizeOtherSheetsFixed();
  colorizeLabelA3Sheets();
  setFontColorByCode();
  setPageDividers();
  colorizeLabels();
  applyManualColors();
}


// =============================================================================
//  1. フリー入力用シート
// =============================================================================

function colorizeFreeInputSheet() {
  const sheet = getSheet('フリー入力用');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // D列(4)・L列(12) をまとめて塗る
  const numRows = lastRow - 1;
  const values  = sheet.getRange(2, 1, numRows, 12).getValues();

  const colD_bg = [];
  const colL_bg = [];

  for (const row of values) {
    const color = getDistrictColor(row[11]); // L列
    colD_bg.push([color]);
    colL_bg.push([color]);
  }

  sheet.getRange(2, 4,  numRows, 1).setBackgrounds(colD_bg);
  sheet.getRange(2, 12, numRows, 1).setBackgrounds(colL_bg);
}


// =============================================================================
//  2. ラベル集計・クロス表YT/ET シート
// =============================================================================

function colorizeOtherSheetsFixed() {
  _colorizeDistrictSheet('ラベル集計', { dataWidth: 11, districtCol: 10, paintCols: [3, 10] });

  for (const name of ['クロス表YT', 'クロス表ET']) {
    _colorizeDistrictSheet(name, { dataWidth: 6, districtCol: 3, paintCols: [3, 5] });
  }
}

/**
 * エリア色をまとめて塗るユーティリティ
 * @param {string} sheetName
 * @param {{ dataWidth: number, districtCol: number, paintCols: number[] }} opts
 *   districtCol: 0-based index of the column used for color lookup
 *   paintCols:   0-based indexes of columns to paint
 */
function _colorizeDistrictSheet(sheetName, { dataWidth, districtCol, paintCols }) {
  const sheet = getSheet(sheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const values  = sheet.getRange(2, 1, numRows, dataWidth).getValues();

  // 塗る列ごとに背景色配列を作成
  const bgArrays = paintCols.map(() => []);

  for (const row of values) {
    const color = getDistrictColor(row[districtCol]);
    paintCols.forEach((_, i) => bgArrays[i].push([color]));
  }

  // 1列ずつまとめて適用（1行1callで済む）
  paintCols.forEach((col0, i) => {
    sheet.getRange(2, col0 + 1, numRows, 1).setBackgrounds(bgArrays[i]);
  });
}


// =============================================================================
//  3. ラベルA3_ゆ / ラベルA3_エ シート — 背景色
// =============================================================================

function colorizeLabelA3Sheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const sheetName of A3_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = getLabelLastRow(sheet);
    if (lastRow < 8) continue;

    // A〜E列・有効行のみクリア（1回のAPIコール）
    sheet.getRange(1, 1, lastRow, 5).setBackground(null);

    const labelCount = Math.ceil(lastRow / 12) * 2;

    // シート全体を一括読み込み → ループ内のgetValue()を排除
    const allValues = sheet.getRange(1, 1, lastRow, 5).getValues();

    // 塗る列・行・色を列ごとに配列で管理
    // A〜E列の背景色を lastRow 分まとめて保持し最後に一括書き込み
    const bgGrid = Array.from({ length: lastRow }, () => Array(5).fill(null));

    for (let i = 0; i < labelCount; i++) {
      const isRight   = i % 2 === 1;
      const rowOffset = Math.floor(i / 2) * 12;
      const colA      = isRight ? 4 : 1; // 1-based
      const colB      = isRight ? 5 : 2; // 1-based

      // allValues は 0-based: [row][col]
      const district = allValues[7 + rowOffset]?.[colA - 1]; // 8行目 = index 7
      const color    = getDistrictColor(district);
      if (!color) continue;

      for (const r of [2 + rowOffset, 3 + rowOffset]) {
        if (r - 1 < lastRow) {
          bgGrid[r - 1][colA - 1] = color;
          bgGrid[r - 1][colB - 1] = color;
        }
      }
    }

    // A〜E列まとめて1回で書き込み
    sheet.getRange(1, 1, lastRow, 5).setBackgrounds(bgGrid);
  }
}


// =============================================================================
//  4. ABCDアソート 文字色
// =============================================================================

function setFontColorByCode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const sheetName of A3_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = getLabelLastRow(sheet);
    if (lastRow < 1) continue;

    // A〜E列を一括取得（B列とE列だけ使う）
    const values = sheet.getRange(1, 1, lastRow, 5).getValues();

    // B列(index 1) と E列(index 4) 分の配列を用意
    const fontColorsB  = [];
    const fontWeightsB = [];
    const fontColorsE  = [];
    const fontWeightsE = [];

    for (const row of values) {
      for (const [arr_fc, arr_fw, colIdx] of [
        [fontColorsB, fontWeightsB, 1],
        [fontColorsE, fontWeightsE, 4],
      ]) {
        const key   = String(row[colIdx]).trim();
        const color = ASSORT_FONT_COLORS[key] ?? null;
        arr_fc.push([color]);
        arr_fw.push([color ? 'bold' : null]);
      }
    }

    // B列・E列それぞれ1回で書き込み
    sheet.getRange(1, 2, lastRow, 1).setFontColors(fontColorsB).setFontWeights(fontWeightsB);
    sheet.getRange(1, 5, lastRow, 1).setFontColors(fontColorsE).setFontWeights(fontWeightsE);
  }
}

// =============================================================================
//  5. ページ区切り
// =============================================================================

function setPageDividers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const sheetName of A3_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = getLabelLastRow(sheet);
    if (lastRow < 1) continue;

    // 60行ごとにページ区切り行を配置
    for (let row = 60, page = 1; row <= lastRow + 60; row += 60, page++) {
      sheet.getRange(row, 1, 1, 5).setBackground(PAGE_DIVIDER_COLOR);
      sheet.getRange(row, 1).setValue(`--- Page ${String(page).padStart(2, '0')} ---`);
    }
  }
}


// =============================================================================
//  6. 制作場所の色付け（エルブ奥・豊ビル）
// =============================================================================

function colorizeLabels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const sheetName of A3_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = getLabelLastRow(sheet);
    if (lastRow < 1) continue;

    // A〜E列を一括取得（A列とD列だけ使う）
    const fullRange   = sheet.getRange(1, 1, lastRow, 5);
    const values      = fullRange.getValues();
    const backgrounds = fullRange.getBackgrounds();
    const fontColors  = fullRange.getFontColors();

    // A列(index 0) と D列(index 3) を処理
    for (const colIdx of [0, 3]) {
      for (let i = 0; i < values.length; i++) {
        const cell     = String(values[i][colIdx]);
        const matchKey = Object.keys(LABEL_PLACE_COLORS).find(k => cell.includes(k));
        if (!matchKey) continue;

        const color = LABEL_PLACE_COLORS[matchKey];

        // 2つ上のセルを白で隠す
        if (i >= 2) {
          backgrounds[i - 2][colIdx] = '#ffffff';
          fontColors[i - 2][colIdx]  = '#ffffff';
        }

        backgrounds[i][colIdx] = color;
        fontColors[i][colIdx]  = color;
      }
    }

    // A〜E列まとめて1回で書き込み
    fullRange.setBackgrounds(backgrounds);
    fullRange.setFontColors(fontColors);
  }
}


// =============================================================================
//  7. クロス表 M列 アソートコード文字色
// =============================================================================

function applyManualColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const name of ['クロス表ET', 'クロス表YT']) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const range  = sheet.getRange(2, 13, lastRow - 1, 1); // M列
    const values = range.getValues();

    const colors = values.map(([v]) => {
      const text = String(v);
      if (text.startsWith('A')) return ['#FF0000'];
      if (text.startsWith('B')) return ['#0000FF'];
      if (text.startsWith('C')) return ['#a52a2a'];
      if (text.startsWith('D')) return ['#008000'];
      return [null];
    });

    range.setFontColors(colors);
  }
}


// =============================================================================
//  ユーティリティ
// =============================================================================

/**
 * シートを名前で取得する（見つからなければ null）
 */
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

/**
 * ラベルA3シートの有効な最終行を返す
 * B列を下からスキャンして最初にデータが見つかった行を返す（1-based）
 * データなしなら 0
 */
function getLabelLastRow(sheet) {
  const maxRow = sheet.getMaxRows();
  const colB   = sheet.getRange(1, 2, maxRow, 1).getValues(); // B列を一括取得
  for (let i = maxRow - 1; i >= 0; i--) {
    if (colB[i][0] !== '' && colB[i][0] !== null) return i + 1;
  }
  return 0;
}

/**
 * エリア名から背景色を返す
 * @param {*} district
 * @returns {string|null}
 */
function getDistrictColor(district) {
  if (!district) return null;
  return DISTRICT_COLORS[String(district).trim()] ?? null;
}