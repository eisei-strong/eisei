// ===== ホープ数・プッシュ数シート作成・管理 =====

var HOPE_SHEET_NAME = '【4月】ホープ数';
var PUSH_SHEET_NAME = '【4月】プッシュ数';
var HOPE_META_COLS = 3;       // A:ID, B:名前, C:合計
var HOPE_DATE_START_COL = 4;  // D列から日別開始
var HOPE_COLS_PER_DAY = 3;    // YT, IG, TT
var HOPE_MONTH_DAYS = 30;     // 4月

/**
 * 投稿数シートを部分一致で検索
 */
function findPostSheet_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf('投稿') !== -1) return sheets[i];
  }
  return null;
}

/**
 * 投稿数シートからメンバーリスト+行背景色を読み取る
 */
function readPostSheetData_(ss) {
  var postSheet = findPostSheet_(ss);
  if (!postSheet) { Logger.log('投稿数シートが見つかりません'); return null; }
  Logger.log('参照シート: ' + postSheet.getName());

  var lastRow = postSheet.getLastRow();
  if (lastRow < 2) { Logger.log('投稿数シートにデータがありません'); return null; }

  var srcData = postSheet.getRange(2, 1, lastRow - 1, POST_APP_NAME_COL).getValues();

  // データ行の背景色を読み取り（メタ列の背景=行の背景色として使う）
  var srcBgs = postSheet.getRange(2, POST_APP_NAME_COL, lastRow - 1, 1).getBackgrounds();
  // 日付列の背景も読み取り（1列目だけでOK）
  var dateBgs = postSheet.getRange(2, POST_APP_DATE_START_COL, lastRow - 1, 1).getBackgrounds();
  // フォントサイズ
  var srcSizes = postSheet.getRange(2, POST_APP_NAME_COL, lastRow - 1, 1).getFontSizes();
  // 行の高さ
  var rowHeights = [];
  for (var r = 2; r <= lastRow; r++) {
    rowHeights.push(postSheet.getRowHeight(r));
  }
  // ヘッダー行の高さ
  var headerHeight1 = postSheet.getRowHeight(1);

  var members = [];
  for (var i = 0; i < srcData.length; i++) {
    var id = String(srcData[i][POST_APP_ID_COL - 1] || '').trim();
    var name = String(srcData[i][POST_APP_NAME_COL - 1] || '').trim();
    if (!id || !/^\d+$/.test(id)) continue;
    members.push({
      id: id,
      name: name,
      nameBg: srcBgs[i][0],
      dateBg: dateBgs[i][0],
      fontSize: srcSizes[i][0],
      rowHeight: rowHeights[i]
    });
  }

  return { members: members, headerHeight: headerHeight1 };
}

/**
 * 共通: SNS3列シートを作成し装飾を適用
 * @param {Sheet} sheet - 作成済みの空シート
 * @param {Array} members - readPostSheetData_の結果
 * @param {number} headerHeight - ヘッダー行高さ
 * @param {string} metaBgColor - メタ列ヘッダーの背景色
 * @param {string} metaFontColor - メタ列ヘッダーの文字色
 * @param {string} dateBgEven - 日付ヘッダー偶数日背景
 * @param {string} dateBgOdd  - 日付ヘッダー奇数日背景
 */
function buildSnsSheet_(sheet, members, headerHeight, metaBgColor, metaFontColor, dateBgEven, dateBgOdd) {
  var totalCols = HOPE_META_COLS + HOPE_MONTH_DAYS * HOPE_COLS_PER_DAY;

  // === ヘッダー1行目 ===
  var h1 = ['ID', '名前', '合計'];
  for (var d = 1; d <= HOPE_MONTH_DAYS; d++) {
    h1.push('4/' + d, '', '');
  }
  sheet.getRange(1, 1, 1, totalCols).setValues([h1]);

  // === ヘッダー2行目 ===
  var h2 = ['', '', ''];
  for (var d = 1; d <= HOPE_MONTH_DAYS; d++) {
    h2.push('YT', 'IG', 'TT');
  }
  sheet.getRange(2, 1, 1, totalCols).setValues([h2]);

  // 日付セルを3列結合
  for (var d = 0; d < HOPE_MONTH_DAYS; d++) {
    var col = HOPE_DATE_START_COL + d * HOPE_COLS_PER_DAY;
    sheet.getRange(1, col, 1, 3).merge().setHorizontalAlignment('center');
  }

  // ヘッダー書式
  sheet.getRange(1, 1, 2, totalCols)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // メタ列ヘッダー背景（シートごとに違う色）
  sheet.getRange(1, 1, 2, HOPE_META_COLS).setBackground(metaBgColor).setFontColor(metaFontColor);

  // SNSサブヘッダー色（各日付）
  for (var d = 0; d < HOPE_MONTH_DAYS; d++) {
    var col = HOPE_DATE_START_COL + d * HOPE_COLS_PER_DAY;
    sheet.getRange(1, col, 1, 3).setBackground(d % 2 === 0 ? dateBgEven : dateBgOdd);
    sheet.getRange(2, col).setBackground('#FF0000').setFontColor('#FFFFFF');       // YT
    sheet.getRange(2, col + 1).setBackground('#C13584').setFontColor('#FFFFFF');   // IG
    sheet.getRange(2, col + 2).setBackground('#25F4EE').setFontColor('#000000');   // TT
  }

  // ヘッダー行高さを投稿数シートに合わせる
  if (headerHeight) {
    sheet.setRowHeight(1, headerHeight);
  }

  // === メンバー行 ===
  var rows = [];
  for (var m = 0; m < members.length; m++) {
    var row = new Array(totalCols).fill('');
    row[0] = members[m].id;
    row[1] = members[m].name;
    rows.push(row);
  }

  if (rows.length > 0) {
    var dataRange = sheet.getRange(3, 1, rows.length, totalCols);
    dataRange.setValues(rows);

    // A列テキスト形式（先頭0対応）
    sheet.getRange(3, 1, rows.length, 1).setNumberFormat('@');

    // C列(合計)数式
    var lastLetter = colToLetter_(totalCols);
    var formulas = [];
    for (var r = 0; r < rows.length; r++) {
      formulas.push(['=SUM(D' + (r + 3) + ':' + lastLetter + (r + 3) + ')']);
    }
    sheet.getRange(3, 3, formulas.length, 1).setFormulas(formulas);

    // === 交互背景色（ゼブラストライプ） ===
    var allBgs = [];
    for (var m = 0; m < members.length; m++) {
      var bg = (m % 2 === 0) ? '#FFFFFF' : '#F3F4F6';
      var bgRow = [];
      for (var c = 0; c < totalCols; c++) {
        bgRow.push(bg);
      }
      allBgs.push(bgRow);
    }
    dataRange.setBackgrounds(allBgs);

    // フォントサイズをコピー
    var sizes = [];
    for (var m = 0; m < members.length; m++) {
      var sizeRow = [];
      for (var c = 0; c < totalCols; c++) {
        sizeRow.push(members[m].fontSize);
      }
      sizes.push(sizeRow);
    }
    dataRange.setFontSizes(sizes);

    // 行高さをコピー
    for (var m = 0; m < members.length; m++) {
      if (members[m].rowHeight) {
        sheet.setRowHeight(m + 3, members[m].rowHeight);
      }
    }

    // データ領域の中央揃え
    sheet.getRange(3, HOPE_DATE_START_COL, rows.length, HOPE_MONTH_DAYS * HOPE_COLS_PER_DAY)
      .setHorizontalAlignment('center');
  }

  // === 列幅調整 ===
  sheet.setColumnWidth(1, 55);   // ID
  sheet.setColumnWidth(2, 100);  // 名前
  sheet.setColumnWidth(3, 45);   // 合計
  for (var c = HOPE_DATE_START_COL; c <= totalCols; c++) {
    sheet.setColumnWidth(c, 30);
  }

  // 固定行列
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(HOPE_META_COLS);

  // 罫線: 日付区切り（3列ごと）
  for (var d = 0; d < HOPE_MONTH_DAYS; d++) {
    var col = HOPE_DATE_START_COL + d * HOPE_COLS_PER_DAY;
    sheet.getRange(1, col, rows.length + 2, 1)
      .setBorder(null, true, null, null, null, null, '#B0B0B0', SpreadsheetApp.BorderStyle.SOLID);
  }

  // 全データ行に薄い罫線
  if (rows.length > 0) {
    sheet.getRange(3, 1, rows.length, totalCols)
      .setBorder(true, true, true, true, true, true, '#D0D0D0', SpreadsheetApp.BorderStyle.SOLID);
  }

  return rows.length;
}

// ===== ホープ数シート作成 =====

/**
 * ホープ数シートを作成（投稿数シートと同じメンバー並び・装飾）
 * GASエディタから手動実行
 */
function createHopeSheet() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);

  if (ss.getSheetByName(HOPE_SHEET_NAME)) {
    Logger.log('既に存在します: ' + HOPE_SHEET_NAME);
    return;
  }

  var data = readPostSheetData_(ss);
  if (!data) return;

  var sheet = ss.insertSheet(HOPE_SHEET_NAME);
  var count = buildSnsSheet_(sheet, data.members, data.headerHeight,
    '#4A90D9', '#FFFFFF',   // メタヘッダー: 青
    '#E8F0FE', '#F3F4F6'   // 日付ヘッダー: 青系交互
  );
  Logger.log('ホープ数シートを作成しました: ' + count + '人');
}

// ===== プッシュ数シート作成 =====

/**
 * プッシュ数シートを作成（投稿数シートと同じメンバー並び・装飾）
 * GASエディタから手動実行
 */
function createPushSheet() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);

  if (ss.getSheetByName(PUSH_SHEET_NAME)) {
    Logger.log('既に存在します: ' + PUSH_SHEET_NAME);
    return;
  }

  var data = readPostSheetData_(ss);
  if (!data) return;

  var sheet = ss.insertSheet(PUSH_SHEET_NAME);
  var count = buildSnsSheet_(sheet, data.members, data.headerHeight,
    '#2E7D32', '#FFFFFF',   // メタヘッダー: 緑
    '#E8F5E9', '#F1F8E9'   // 日付ヘッダー: 緑系交互
  );
  Logger.log('プッシュ数シートを作成しました: ' + count + '人');
}

// ===== 既存シートに装飾だけ再適用 =====

/**
 * ホープ数・プッシュ数の装飾を投稿数シートに合わせて再適用
 * GASエディタから手動実行
 */
function reformatHopeAndPushSheets() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var data = readPostSheetData_(ss);
  if (!data) return;

  var sheets = [
    { name: HOPE_SHEET_NAME, metaBg: '#4A90D9', evenBg: '#E8F0FE', oddBg: '#F3F4F6' },
    { name: PUSH_SHEET_NAME, metaBg: '#2E7D32', evenBg: '#E8F5E9', oddBg: '#F1F8E9' }
  ];

  for (var s = 0; s < sheets.length; s++) {
    var sheet = ss.getSheetByName(sheets[s].name);
    if (!sheet) { Logger.log(sheets[s].name + ' が見つかりません'); continue; }

    var lastRow = sheet.getLastRow();
    if (lastRow < 3) continue;
    var totalCols = HOPE_META_COLS + HOPE_MONTH_DAYS * HOPE_COLS_PER_DAY;
    var numRows = lastRow - 2;

    // IDでマッチしてメンバーの背景色を取得
    var ids = sheet.getRange(3, 1, numRows, 1).getValues();
    var memberMap = {};
    for (var m = 0; m < data.members.length; m++) {
      memberMap[data.members[m].id] = data.members[m];
    }

    // 交互背景色
    var allBgs = [];
    for (var i = 0; i < numRows; i++) {
      var bg = (i % 2 === 0) ? '#FFFFFF' : '#F3F4F6';
      var bgRow = [];
      for (var c = 0; c < totalCols; c++) bgRow.push(bg);
      allBgs.push(bgRow);
    }
    sheet.getRange(3, 1, numRows, totalCols).setBackgrounds(allBgs);
    sheet.getRange(3, HOPE_DATE_START_COL, numRows, HOPE_MONTH_DAYS * HOPE_COLS_PER_DAY)
      .setHorizontalAlignment('center');

    // 罫線再適用
    sheet.getRange(3, 1, numRows, totalCols)
      .setBorder(true, true, true, true, true, true, '#D0D0D0', SpreadsheetApp.BorderStyle.SOLID);

    Logger.log(sheets[s].name + ' の装飾を更新しました（' + numRows + '行）');
  }
}

/**
 * 列番号→列文字変換（1-based）
 */
function colToLetter_(col) {
  var letter = '';
  while (col > 0) {
    var mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
