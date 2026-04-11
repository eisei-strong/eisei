// ============================================
// Setup.js — シート自動生成・書式設定 (v2)
// ============================================

// --- 配色定義 ---
var COLOR_HEADER_BG    = '#0F172A';
var COLOR_HEADER_TEXT  = '#F1F5F9';
var COLOR_TOTAL_BG     = '#1E293B';
var COLOR_ORANGE_DARK  = '#7C2D12';
var COLOR_ORANGE_MID   = '#EA580C';
var COLOR_ORANGE_LIGHT = '#FFF7ED';
var COLOR_ORANGE_BG    = '#FED7AA';
var COLOR_WHITE        = '#FFFFFF';
var COLOR_GRAY_LIGHT   = '#F8FAFC';
var COLOR_GRAY_BORDER  = '#E2E8F0';
var COLOR_TEXT_DARK     = '#1E293B';
var COLOR_GOLD_BG      = '#FEF3C7';
var COLOR_GOLD_TEXT     = '#92400E';
var COLOR_GREEN        = '#059669';
var COLOR_YELLOW       = '#CA8A04';
var COLOR_RED          = '#DC2626';

/**
 * 全シートを自動生成（メニューから実行）
 */
function setupAllSheets() {
  var ss = getSpreadsheet_();
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    'シート生成',
    '設定・日別入力・サマリー・アーカイブ・CO残管理シートを生成します。\n既存の同名シートは上書きされます。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  createSettingsSheet_(ss);
  createDailySheets_(ss);
  createSummarySheet_(ss);
  createArchiveSheet_(ss);
  createCOManageSheet_(ss);
  applyFormatting_(ss);

  ui.alert('完了', '全シートの生成が完了しました。', ui.ButtonSet.OK);
}

/**
 * 設定シートを作成 (v2)
 */
function createSettingsSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_SETTINGS);
  sheet.clear();

  // メンバー情報ヘッダー (A-E)
  sheet.getRange(1, 1, 1, SETTINGS_MEMBER_COL_COUNT).setValues([
    ['メンバー名', '表示名', '着金ランキング順位', 'ステータス', '配色コード']
  ]);

  // メンバーデータ
  var members = [
    ['AをAでやる',                 'AをAでやる',   3, 'アクティブ', '#3B82F6'],
    ['ドライ',                     'ドライ',       2, 'アクティブ', '#10B981'],
    ['ヒトコト',                   'ヒトコト',     1, 'アクティブ', '#F59E0B'],
    ['ビッグマウス',               'ビッグマウス', 3, 'アクティブ', '#EF4444'],
    ['けつだん',                   'けつだん',     3, 'アクティブ', '#A78BFA'],
    ['ぜんぶり',                   'ぜんぶり',     3, 'アクティブ', '#22D3EE'],
    ['スクリプト通りに営業するくん', 'スクリプトくん', 3, 'アクティブ', '#10B981'],
    ['ワントーン',                 'ワントーン',   3, 'アクティブ', '#3B82F6'],
    ['ゴン',                       'ゴン',         3, 'アクティブ', '#F97316'],
    ['ガロウ',                     'ガロウ',       '-', '退職済', '#6B7280'],
    ['リヴァイ',                   'リヴァイ',     '-', '退職済', '#6B7280']
  ];
  sheet.getRange(2, 1, members.length, SETTINGS_MEMBER_COL_COUNT).setValues(members);

  // 空行スペーサー (行12-13)

  // グローバル設定 (行14以降)
  var now = new Date();
  sheet.getRange(SETTINGS_ROW_GLOBAL_START, 1, 1, 2).setValues([['--- グローバル設定 ---', '']]);
  sheet.getRange(SETTINGS_ROW_GLOBAL_START, 1).setFontWeight('bold').setFontColor(COLOR_ORANGE_DARK);

  var globals = [
    ['年間着金目標', TEAM_GOAL],
    ['対象年度', now.getFullYear()],
    ['対象月', now.getMonth() + 1]
  ];
  sheet.getRange(SETTINGS_ROW_GLOBAL_START + 1, 1, globals.length, 2).setValues(globals);

  // 書式
  sheet.getRange(1, 1, 1, SETTINGS_MEMBER_COL_COUNT)
    .setBackground(COLOR_HEADER_BG).setFontColor(COLOR_HEADER_TEXT).setFontWeight('bold');

  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 100);
  sheet.setTabColor(COLOR_ORANGE_MID);

  // ステータスのドロップダウン
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['アクティブ', '退職済'], true)
    .build();
  sheet.getRange(2, SETTINGS_COL_STATUS, members.length, 1).setDataValidation(statusRule);

  sheet.setFrozenRows(1);
}

/**
 * 日別入力シートをアクティブメンバーごとに作成
 */
function createDailySheets_(ss) {
  var members = getActiveMembers_(ss);
  var settings = getGlobalSettings_(ss);

  for (var i = 0; i < members.length; i++) {
    createDailySheet_(ss, members[i].name, settings.year, settings.month);
  }
}

/**
 * 単一メンバーの日別入力シートを作成 (v2)
 */
function createDailySheet_(ss, memberName, year, month) {
  var sheetName = SHEET_DAILY_PREFIX + memberName;
  var sheet = getOrCreateSheet_(ss, sheetName);
  sheet.clear();

  // ヘッダー行 (Row 1)
  sheet.getRange(DAILY_ROW_HEADER, 1, 1, DAILY_COL_COUNT).setValues([DAILY_HEADERS]);

  // 前月繰越行 (Row 2)
  var carryoverRow = new Array(DAILY_COL_COUNT).fill('');
  carryoverRow[0] = '前月繰越';
  sheet.getRange(DAILY_ROW_CARRYOVER, 1, 1, DAILY_COL_COUNT).setValues([carryoverRow]);

  // 日付自動生成 (Row 3-33)
  var daysInMonth = new Date(year, month, 0).getDate();
  var dateData = [];
  for (var d = 1; d <= 31; d++) {
    var row = new Array(DAILY_COL_COUNT).fill('');
    if (d <= daysInMonth) {
      row[COL_DATE] = new Date(year, month - 1, d);
    }
    dateData.push(row);
  }
  sheet.getRange(DAILY_ROW_DATA_START, 1, 31, DAILY_COL_COUNT).setValues(dateData);

  // 空行 (Row 34)

  // 合計行 (Row 35) - SUM数式
  var totalRow = ['合計'];
  for (var c = 1; c < DAILY_COL_COUNT; c++) {
    var colLetter = columnToLetter_(c + 1);
    totalRow.push('=SUM(' + colLetter + DAILY_ROW_CARRYOVER + ':' + colLetter + DAILY_ROW_DATA_END + ')');
  }
  sheet.getRange(DAILY_ROW_TOTAL, 1, 1, DAILY_COL_COUNT).setValues([totalRow]);

  // CBS / ライフティ (Row 37-40)
  sheet.getRange(DAILY_ROW_CBS_APPROVED, 1).setValue('CBS承認数');
  sheet.getRange(DAILY_ROW_CBS_APPLIED, 1).setValue('CBS申請数');
  sheet.getRange(DAILY_ROW_LF_APPROVED, 1).setValue('ライフティ承認数');
  sheet.getRange(DAILY_ROW_LF_APPLIED, 1).setValue('ライフティ申請数');

  // クレカ・信販合計 (Row 42-43)
  sheet.getRange(DAILY_ROW_CREDIT_TOTAL, 1).setValue('クレカ着金合計');
  sheet.getRange(DAILY_ROW_SHINPAN_TOTAL, 1).setValue('信販着金合計');

  // 列幅設定
  sheet.setColumnWidth(1, 100);
  for (var w = 2; w <= DAILY_COL_COUNT; w++) {
    sheet.setColumnWidth(w, 120);
  }

  // 日付フォーマット
  sheet.getRange(DAILY_ROW_DATA_START, 1, 31, 1).setNumberFormat('M/d');

  // 数値フォーマット（万円列: D,E,G）
  var moneyFormat = '#,##0.0';
  sheet.getRange(DAILY_ROW_CARRYOVER, COL_REVENUE + 1, 32, 1).setNumberFormat(moneyFormat);
  sheet.getRange(DAILY_ROW_CARRYOVER, COL_SALES + 1, 32, 1).setNumberFormat(moneyFormat);
  sheet.getRange(DAILY_ROW_CARRYOVER, COL_CO_AMOUNT + 1, 32, 1).setNumberFormat(moneyFormat);
  sheet.getRange(DAILY_ROW_TOTAL, 2, 1, DAILY_COL_COUNT - 1).setNumberFormat('#,##0.0');

  // タブ色
  sheet.setTabColor('#2563EB');
}

/**
 * サマリーシートを作成 (v2)
 */
function createSummarySheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_SUMMARY);
  sheet.clear();

  var settings = getGlobalSettings_(ss);

  // セクション1: チームKPI (行1-8)
  sheet.getRange(SM_ROW_TITLE, 1, 1, 2).setValues([
    ['⚔️ Namaka ウォーリアーズ', settings.year + '年' + settings.month + '月']
  ]);
  sheet.getRange(SM_ROW_UPDATED, 1).setValue('最終更新');

  var kpiLabels = [
    ['合計着金額（繰越含）', ''],
    ['合計商談数（当月）', ''],
    ['合計成約数', ''],
    ['合計売上', ''],
    ['合計CO', ''],
    ['年間目標進捗率', '']
  ];
  sheet.getRange(SM_ROW_KPI_REVENUE, 1, kpiLabels.length, 2).setValues(kpiLabels);

  // セクション2: メンバー別成績テーブル (行10-)
  sheet.getRange(SM_ROW_MEMBER_HEADER, 1).setValue('▶ メンバー別成績');
  var memberLabels = ['メンバー', 'ランク', '資金有商談', '資金なし商談', '合計商談',
    '成約数', '成約率', '売上', '着金額(繰越含)', 'CO', 'ライフティ承認率',
    '前月着金額', '前月比', 'トップとの差', '商談差', '成約差', '成約率差'];
  sheet.getRange(SM_ROW_MEMBER_HEADER, 1, 1, SM_MEMBER_COL_COUNT).setValues([memberLabels]);

  // 合計行
  sheet.getRange(SM_ROW_MEMBER_TOTAL, 1).setValue('合計');

  // セクション3: 期間比較 (行22-)
  sheet.getRange(SM_ROW_PERIOD_HEADER, 1, 1, 4).setValues([['▶ 期間比較', '当月商談数', '先月商談数', '対比']]);
  sheet.getRange(SM_ROW_PERIOD_1_10, 1).setValue('1日〜10日');
  sheet.getRange(SM_ROW_PERIOD_11_20, 1).setValue('11日〜20日');
  sheet.getRange(SM_ROW_PERIOD_21_END, 1).setValue('21日〜末日');

  // セクション4: 着金内訳 (行28-)
  sheet.getRange(SM_ROW_BREAKDOWN_HEADER, 1).setValue('▶ 着金内訳');
  sheet.getRange(SM_ROW_CREDIT_TOTAL, 1).setValue('クレカ着金');
  sheet.getRange(SM_ROW_SHINPAN_TOTAL, 1).setValue('信販会社着金');
  sheet.getRange(SM_ROW_CURRENT_REVENUE, 1).setValue('当月着金のみ');
  sheet.getRange(SM_ROW_CARRYOVER_REVENUE, 1).setValue('繰り越し着金額');
  sheet.getRange(SM_ROW_CUMULATIVE_CLOSED, 1).setValue('8月からの累計成約数');

  // セクション5: 退職者CO残 (行35-)
  sheet.getRange(SM_ROW_CO_HEADER, 1).setValue('▶ 退職者CO残サマリー');

  // 列幅
  sheet.setColumnWidth(1, 160);
  for (var col = 2; col <= 11; col++) {
    sheet.setColumnWidth(col, 110);
  }
  sheet.setColumnWidth(12, 110);  // L: 前月着金額
  sheet.setColumnWidth(13, 100);  // M: 前月比
  sheet.setColumnWidth(14, 110);  // N: トップとの差
  sheet.setColumnWidth(15, 80);   // O: 商談差
  sheet.setColumnWidth(16, 80);   // P: 成約差
  sheet.setColumnWidth(17, 80);   // Q: 成約率差

  sheet.setTabColor(COLOR_GREEN);
}

/**
 * 月次アーカイブシートを作成
 */
function createArchiveSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_ARCHIVE);
  if (sheet.getLastRow() > 0) {
    var existing = sheet.getRange(1, 1, 1, 1).getValue();
    if (String(existing).trim() === ARCHIVE_HEADERS[0]) return;
  }

  sheet.getRange(1, 1, 1, ARCHIVE_HEADERS.length).setValues([ARCHIVE_HEADERS]);
  sheet.getRange(1, 1, 1, ARCHIVE_HEADERS.length)
    .setBackground(COLOR_HEADER_BG).setFontColor(COLOR_HEADER_TEXT).setFontWeight('bold');

  sheet.setTabColor('#7C3AED');
}

/**
 * CO残管理シートを作成 (v2)
 */
function createCOManageSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_CO_MANAGE);
  sheet.clear();

  // 行1: タイトル
  sheet.getRange(CO_ROW_SUMMARY_TITLE, 1).setValue('退職者CO残 サマリー');

  // 行2: サマリーヘッダー
  sheet.getRange(CO_ROW_SUMMARY_HEADER, 1, 1, 4).setValues([
    ['メンバー', 'CO総額', '回収済', '未回収残高']
  ]);

  // 行3-4: 退職者サマリー（初期値）
  sheet.getRange(CO_ROW_SUMMARY_START, 1, 2, 4).setValues([
    ['ガロウ', '', '', ''],
    ['リヴァイ', '', '', '']
  ]);

  // 行5: 合計
  sheet.getRange(CO_ROW_SUMMARY_TOTAL, 1, 1, 4).setValues([
    ['合計', '', '', '']
  ]);

  // 行7: データテーブルヘッダー
  var dataHeaders = [
    'メンバー名', '成約日', '顧客名/案件ID', '成約金額（万円）',
    'CO発生日', 'CO金額（万円）', '回収ステータス', '請求日',
    '回収日', '回収金額（万円）', '未回収残高（万円）', '備考'
  ];
  sheet.getRange(CO_ROW_DATA_HEADER, 1, 1, CO_COL_COUNT).setValues([dataHeaders]);

  // 行8以降: 未回収残高の自動計算数式（50行分）
  for (var r = CO_ROW_DATA_START; r <= CO_ROW_DATA_START + 49; r++) {
    sheet.getRange(r, CO_COL_REMAINING).setFormula(
      '=IF(F' + r + '="","",F' + r + '-IF(J' + r + '="",0,J' + r + '))'
    );
  }

  // SUMIF集計数式（サマリー）
  var summaryMembers = ['ガロウ', 'リヴァイ'];
  for (var i = 0; i < summaryMembers.length; i++) {
    var row = CO_ROW_SUMMARY_START + i;
    // CO総額 = SUMIF(メンバー名列, メンバー, CO金額列)
    sheet.getRange(row, 2).setFormula(
      '=SUMIF(A' + CO_ROW_DATA_START + ':A200,A' + row + ',F' + CO_ROW_DATA_START + ':F200)'
    );
    // 回収済 = SUMIF(メンバー名列, メンバー, 回収金額列)
    sheet.getRange(row, 3).setFormula(
      '=SUMIF(A' + CO_ROW_DATA_START + ':A200,A' + row + ',J' + CO_ROW_DATA_START + ':J200)'
    );
    // 未回収残高 = CO総額 - 回収済
    sheet.getRange(row, 4).setFormula('=B' + row + '-C' + row);
  }

  // 合計行数式
  sheet.getRange(CO_ROW_SUMMARY_TOTAL, 2).setFormula('=SUM(B' + CO_ROW_SUMMARY_START + ':B' + (CO_ROW_SUMMARY_START + 1) + ')');
  sheet.getRange(CO_ROW_SUMMARY_TOTAL, 3).setFormula('=SUM(C' + CO_ROW_SUMMARY_START + ':C' + (CO_ROW_SUMMARY_START + 1) + ')');
  sheet.getRange(CO_ROW_SUMMARY_TOTAL, 4).setFormula('=SUM(D' + CO_ROW_SUMMARY_START + ':D' + (CO_ROW_SUMMARY_START + 1) + ')');

  // 回収ステータスのドロップダウン
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CO_STATUSES, true)
    .build();
  sheet.getRange(CO_ROW_DATA_START, CO_COL_STATUS, 50, 1).setDataValidation(statusRule);

  // 書式
  sheet.getRange(CO_ROW_SUMMARY_TITLE, 1).setFontSize(14).setFontWeight('bold').setFontColor(COLOR_ORANGE_DARK);
  sheet.getRange(CO_ROW_SUMMARY_HEADER, 1, 1, 4)
    .setBackground(COLOR_HEADER_BG).setFontColor(COLOR_HEADER_TEXT).setFontWeight('bold');
  sheet.getRange(CO_ROW_SUMMARY_TOTAL, 1, 1, 4)
    .setBackground(COLOR_TOTAL_BG).setFontColor(COLOR_HEADER_TEXT).setFontWeight('bold');
  sheet.getRange(CO_ROW_DATA_HEADER, 1, 1, CO_COL_COUNT)
    .setBackground(COLOR_ORANGE_MID).setFontColor(COLOR_WHITE).setFontWeight('bold');

  // 列幅
  sheet.setColumnWidth(1, 120);  // メンバー名
  sheet.setColumnWidth(2, 100);  // 成約日
  sheet.setColumnWidth(3, 180);  // 顧客名
  sheet.setColumnWidth(4, 120);  // 成約金額
  sheet.setColumnWidth(5, 100);  // CO発生日
  sheet.setColumnWidth(6, 120);  // CO金額
  sheet.setColumnWidth(7, 100);  // 回収ステータス
  sheet.setColumnWidth(8, 100);  // 請求日
  sheet.setColumnWidth(9, 100);  // 回収日
  sheet.setColumnWidth(10, 120); // 回収金額
  sheet.setColumnWidth(11, 120); // 未回収残高
  sheet.setColumnWidth(12, 200); // 備考

  // 日付フォーマット
  sheet.getRange(CO_ROW_DATA_START, CO_COL_DEAL_DATE, 50, 1).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(CO_ROW_DATA_START, CO_COL_CO_DATE, 50, 1).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(CO_ROW_DATA_START, CO_COL_CLAIM_DATE, 50, 1).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(CO_ROW_DATA_START, CO_COL_COLLECT_DATE, 50, 1).setNumberFormat('yyyy/MM/dd');

  // 金額フォーマット
  sheet.getRange(CO_ROW_DATA_START, CO_COL_DEAL_AMOUNT, 50, 1).setNumberFormat('#,##0.0');
  sheet.getRange(CO_ROW_DATA_START, CO_COL_CO_AMOUNT, 50, 1).setNumberFormat('#,##0.0');
  sheet.getRange(CO_ROW_DATA_START, CO_COL_COLLECTED, 50, 1).setNumberFormat('#,##0.0');
  sheet.getRange(CO_ROW_DATA_START, CO_COL_REMAINING, 50, 1).setNumberFormat('#,##0.0');

  sheet.setTabColor(COLOR_RED);
}

// ============================================
// 書式設定
// ============================================

/**
 * サマリーシートのヘッダーと書式を再適用（API用）
 */
function refreshSummaryFormat() {
  var ss = getSpreadsheet_();
  var summarySheet = getSummarySheet_(ss);
  if (!summarySheet) return { error: 'サマリーシートなし' };

  // ヘッダーラベルを更新
  var memberLabels = ['メンバー', 'ランク', '資金有商談', '資金なし商談', '合計商談',
    '成約数', '成約率', '売上', '着金額(繰越含)', 'CO', 'ライフティ承認率',
    '前月着金額', '前月比', 'トップとの差', '商談差', '成約差', '成約率差'];
  summarySheet.getRange(SM_ROW_MEMBER_HEADER, 1, 1, SM_MEMBER_COL_COUNT).setValues([memberLabels]);

  // 列幅
  summarySheet.setColumnWidth(1, 160);
  for (var col = 2; col <= 11; col++) {
    summarySheet.setColumnWidth(col, 110);
  }
  summarySheet.setColumnWidth(12, 110);
  summarySheet.setColumnWidth(13, 100);
  summarySheet.setColumnWidth(14, 110);
  summarySheet.setColumnWidth(15, 80);
  summarySheet.setColumnWidth(16, 80);
  summarySheet.setColumnWidth(17, 80);

  // 条件付き書式クリア＆再適用
  summarySheet.clearConditionalFormatRules();
  formatSummarySheet_(summarySheet);

  return { success: true };
}

function applyFormatting_(ss) {
  var settingsSheet = getSettingsSheet_(ss);
  if (settingsSheet) settingsSheet.setFrozenRows(1);

  var summarySheet = getSummarySheet_(ss);
  if (summarySheet) formatSummarySheet_(summarySheet);

  var members = getActiveMembers_(ss);
  for (var i = 0; i < members.length; i++) {
    var dailySheet = getDailySheet_(ss, members[i].name);
    if (dailySheet) formatDailySheet_(dailySheet);
  }
}

function formatSummarySheet_(sheet) {
  var memberRows = SM_ROW_MEMBER_END - SM_ROW_MEMBER_START + 1;
  var borderRows = memberRows + 1; // ヘッダー含む

  // タイトル
  sheet.getRange(SM_ROW_TITLE, 1).setFontSize(16).setFontWeight('bold').setFontColor(COLOR_ORANGE_DARK);

  // KPIラベル列
  sheet.getRange(SM_ROW_KPI_REVENUE, 1, 6, 1)
    .setBackground(COLOR_ORANGE_LIGHT).setFontWeight('bold').setFontColor(COLOR_TEXT_DARK);

  // メンバーテーブルヘッダー
  sheet.getRange(SM_ROW_MEMBER_HEADER, 1, 1, SM_MEMBER_COL_COUNT)
    .setBackground(COLOR_HEADER_BG).setFontColor(COLOR_HEADER_TEXT).setFontWeight('bold').setFontSize(10);

  // 着金額列を金色強調
  sheet.getRange(SM_ROW_MEMBER_START, SM_COL_REVENUE, memberRows, 1)
    .setBackground(COLOR_GOLD_BG).setFontWeight('bold').setFontColor('#000000');

  // 合計行
  sheet.getRange(SM_ROW_MEMBER_TOTAL, 1, 1, SM_MEMBER_COL_COUNT)
    .setBackground(COLOR_TOTAL_BG).setFontColor(COLOR_HEADER_TEXT).setFontWeight('bold');

  // セクションヘッダー
  var sectionRows = [SM_ROW_PERIOD_HEADER, SM_ROW_BREAKDOWN_HEADER, SM_ROW_CO_HEADER];
  for (var i = 0; i < sectionRows.length; i++) {
    sheet.getRange(sectionRows[i], 1, 1, SM_MEMBER_COL_COUNT)
      .setBackground(COLOR_ORANGE_LIGHT).setFontWeight('bold').setFontColor(COLOR_ORANGE_DARK);
  }

  // 成約率の条件付き書式（60%以上=緑, 30-59%=黄, 30%未満=赤）— 文字は黒
  var rateRange = sheet.getRange(SM_ROW_MEMBER_START, SM_COL_CLOSE_RATE, memberRows, 1);
  var greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(60)
    .setBackground('#D1FAE5').setFontColor('#000000')
    .setRanges([rateRange]).build();
  var yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(30, 59.9)
    .setBackground('#FEF3C7').setFontColor('#000000')
    .setRanges([rateRange]).build();
  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(30)
    .setBackground('#FEE2E2').setFontColor('#000000')
    .setRanges([rateRange]).build();
  sheet.setConditionalFormatRules([greenRule, yellowRule, redRule]);

  // ランキング列の条件付き書式 — 文字は黒
  var rankRange = sheet.getRange(SM_ROW_MEMBER_START, SM_COL_RANK, memberRows, 1);
  var goldRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1).setBackground('#FEF3C7').setFontColor('#000000')
    .setRanges([rankRange]).build();
  var silverRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(2).setBackground('#F1F5F9').setFontColor('#000000')
    .setRanges([rankRange]).build();
  var bronzeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(3).setBackground('#FED7AA').setFontColor('#000000')
    .setRanges([rankRange]).build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(goldRule, silverRule, bronzeRule);

  // === 前月比・ランキング列の書式（L-Q列） ===

  // 前月着金額列（L）: 控えめなグレー背景、文字は黒
  sheet.getRange(SM_ROW_MEMBER_START, SM_COL_PREV_REVENUE, memberRows, 1)
    .setBackground('#F8FAFC').setFontColor('#000000');

  // L列の左に太い区切り線（既存データと比較データの境界）
  sheet.getRange(SM_ROW_MEMBER_HEADER, SM_COL_PREV_REVENUE, borderRows + 1, 1)
    .setBorder(null, true, null, null, null, null, '#7C2D12', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 数値フォーマット
  sheet.getRange(SM_ROW_MEMBER_START, SM_COL_PREV_REVENUE, memberRows, 3)
    .setNumberFormat('#,##0.0');
  sheet.getRange(SM_ROW_MEMBER_START, SM_COL_DIFF_DEALS, memberRows, 2)
    .setNumberFormat('+#;-#;0');
  sheet.getRange(SM_ROW_MEMBER_START, SM_COL_DIFF_RATE, memberRows, 1)
    .setNumberFormat('+#,##0.0;-#,##0.0;0');

  // 前月比（着金M列）: 正=緑背景, 負=赤背景、文字は黒
  var diffRevRange = sheet.getRange(SM_ROW_MEMBER_START, SM_COL_DIFF_REVENUE, memberRows, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setBackground('#D1FAE5').setFontColor('#000000')
    .setRanges([diffRevRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setBackground('#FEE2E2').setFontColor('#000000')
    .setRanges([diffRevRange]).build());

  // トップとの差（N列）: 0=金色背景(トップ), >0=赤背景、文字は黒
  var gapRange = sheet.getRange(SM_ROW_MEMBER_START, SM_COL_GAP_TO_TOP, memberRows, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0).setBackground('#FEF3C7').setFontColor('#000000')
    .setRanges([gapRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setBackground('#FEE2E2').setFontColor('#000000')
    .setRanges([gapRange]).build());

  // 商談差(O), 成約差(P), 成約率差(Q): 正=緑背景, 負=赤背景、文字は黒
  var diffCols = [SM_COL_DIFF_DEALS, SM_COL_DIFF_CLOSED, SM_COL_DIFF_RATE];
  for (var dc = 0; dc < diffCols.length; dc++) {
    var dRange = sheet.getRange(SM_ROW_MEMBER_START, diffCols[dc], memberRows, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setBackground('#D1FAE5').setFontColor('#000000')
      .setRanges([dRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setBackground('#FEE2E2').setFontColor('#000000')
      .setRanges([dRange]).build());
  }

  sheet.setConditionalFormatRules(rules);

  sheet.setFrozenRows(SM_ROW_MEMBER_HEADER);
}

function formatDailySheet_(sheet) {
  // ヘッダー行
  sheet.getRange(DAILY_ROW_HEADER, 1, 1, DAILY_COL_COUNT)
    .setBackground(COLOR_HEADER_BG).setFontColor(COLOR_HEADER_TEXT).setFontWeight('bold').setFontSize(10);

  // 前月繰越行
  sheet.getRange(DAILY_ROW_CARRYOVER, 1, 1, DAILY_COL_COUNT)
    .setBackground(COLOR_ORANGE_BG).setFontColor(COLOR_ORANGE_DARK).setFontWeight('bold');

  // 合計行
  sheet.getRange(DAILY_ROW_TOTAL, 1, 1, DAILY_COL_COUNT)
    .setBackground(COLOR_GOLD_BG).setFontWeight('bold').setFontColor(COLOR_GOLD_TEXT);

  // CBS/ライフティ行ラベル
  var specialRows = [DAILY_ROW_CBS_APPROVED, DAILY_ROW_CBS_APPLIED, DAILY_ROW_LF_APPROVED, DAILY_ROW_LF_APPLIED];
  for (var i = 0; i < specialRows.length; i++) {
    sheet.getRange(specialRows[i], 1).setBackground(COLOR_ORANGE_LIGHT).setFontWeight('bold');
  }

  // クレカ・信販行ラベル
  sheet.getRange(DAILY_ROW_CREDIT_TOTAL, 1).setBackground(COLOR_ORANGE_LIGHT).setFontWeight('bold');
  sheet.getRange(DAILY_ROW_SHINPAN_TOTAL, 1).setBackground(COLOR_ORANGE_LIGHT).setFontWeight('bold');

  // データ行の交互色
  for (var r = DAILY_ROW_DATA_START; r <= DAILY_ROW_DATA_END; r++) {
    var bg = ((r - DAILY_ROW_DATA_START) % 2 === 0) ? COLOR_WHITE : COLOR_GRAY_LIGHT;
    sheet.getRange(r, 1, 1, DAILY_COL_COUNT).setBackground(bg);
  }

  // 枠線
  sheet.getRange(DAILY_ROW_HEADER, 1, DAILY_ROW_DATA_END - DAILY_ROW_HEADER + 1, DAILY_COL_COUNT)
    .setBorder(true, true, true, true, true, true, COLOR_GRAY_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // 入力制限: 数値のみ (B-K列のデータ部分)
  var numRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(DAILY_ROW_DATA_START, 2, 31, DAILY_COL_COUNT - 1).setDataValidation(numRule);

  sheet.setFrozenRows(1);
}

// ============================================
// ユーティリティ
// ============================================

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function columnToLetter_(col) {
  var letter = '';
  while (col > 0) {
    var mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ============================================
// 2026年4月 新メンバー追加マイグレーション
// ============================================

/**
 * 新メンバー6人を追加（設定シート・サマリーシート・日別入力シート）
 * Apps Scriptエディタから手動実行
 */
function addNewMembers_April2026() {
  var ss = getSpreadsheet_();

  // 1. 設定シート: グローバル設定を行22に移動 + 新メンバー追加
  migrateSettingsForNewMembers_(ss);

  // 2. サマリーシート: メンバー行を6行拡張
  migrateSummaryForNewMembers_(ss);

  // 3. 新メンバーの日別入力シート作成
  var settings = getGlobalSettings_(ss);
  var newMembers = ['長谷部', 'ゴジータ', 'L', '悟空', 'やまと', '夜神月'];
  for (var i = 0; i < newMembers.length; i++) {
    createDailySheet_(ss, newMembers[i], settings.year, settings.month);
  }

  Logger.log('完了: 新メンバー6人追加 + 設定シート移行 + サマリー拡張 + 日別入力シート作成');
}

/**
 * 日別入力シートが未作成の新メンバー分だけ作成（タイムアウト対策）
 * Apps Scriptエディタから手動実行
 */
function createMissingDailySheets() {
  var ss = getSpreadsheet_();
  var settings = getGlobalSettings_(ss);
  var newMembers = ['長谷部', 'ゴジータ', 'L', '悟空', 'やまと', '夜神月'];

  for (var i = 0; i < newMembers.length; i++) {
    var sheetName = SHEET_DAILY_PREFIX + newMembers[i];
    if (ss.getSheetByName(sheetName)) {
      Logger.log('スキップ（既存）: ' + sheetName);
      continue;
    }
    createDailySheet_(ss, newMembers[i], settings.year, settings.month);
    Logger.log('作成完了: ' + sheetName);
  }
  Logger.log('全シート作成完了');
}

/**
 * 設定シート: グローバル設定を行14→行22に移動 + 新メンバー行追加
 */
function migrateSettingsForNewMembers_(ss) {
  var sheet = getSettingsSheet_(ss);
  if (!sheet) return;

  // 旧グローバル設定を読み取り（行14〜）
  var oldGlobalStart = 14;
  var lastRow = sheet.getLastRow();
  var globalData = [];
  if (lastRow >= oldGlobalStart) {
    globalData = sheet.getRange(oldGlobalStart, 1, lastRow - oldGlobalStart + 1, 2).getValues();
    // 旧グローバル設定をクリア
    sheet.getRange(oldGlobalStart, 1, lastRow - oldGlobalStart + 1, 2).clearContent();
  }

  // 新メンバー6人を追加（既存メンバーの後に）
  var newMembers = [
    ['長谷部',   '長谷部',   '', 'アクティブ', '#F472B6'],
    ['ゴジータ', 'ゴジータ', '', 'アクティブ', '#34D399'],
    ['L',        'L',        '', 'アクティブ', '#60A5FA'],
    ['悟空',     '悟空',     '', 'アクティブ', '#FBBF24'],
    ['やまと',   'やまと',   '', 'アクティブ', '#A78BFA'],
    ['夜神月',   '夜神月',   '', 'アクティブ', '#F87171']
  ];

  // 既存メンバーの最終行を検出
  var memberLastRow = oldGlobalStart - 1;
  for (var r = oldGlobalStart - 1; r >= 2; r--) {
    var val = sheet.getRange(r, 1).getValue();
    if (String(val).trim() !== '') {
      memberLastRow = r;
      break;
    }
  }

  // 新メンバーを挿入
  var insertRow = memberLastRow + 1;
  sheet.getRange(insertRow, 1, newMembers.length, 5).setValues(newMembers);

  // ステータスのドロップダウン設定
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['アクティブ', '退職済'], true)
    .build();
  sheet.getRange(insertRow, 4, newMembers.length, 1).setDataValidation(statusRule);

  // グローバル設定を行22に書き戻し
  if (globalData.length > 0) {
    sheet.getRange(22, 1, globalData.length, 2).setValues(globalData);
  }

  Logger.log('設定シート移行完了: 新メンバー' + newMembers.length + '人追加, グローバル設定を行22に移動');
}

/**
 * サマリーシート: 行21の前に6行挿入してメンバー枠を拡張
 */
function migrateSummaryForNewMembers_(ss) {
  var sheet = getSummarySheet_(ss);
  if (!sheet) return;

  // 旧合計行(21)の前に6行挿入 → メンバー行が11-26に拡張される
  sheet.insertRowsBefore(21, 6);

  // 挿入後の確認: 合計行ラベルを再設定（行27）
  sheet.getRange(27, 1).setValue('合計');

  Logger.log('サマリーシート拡張完了: 行21の前に6行挿入（メンバー枠 10→16人）');
}
