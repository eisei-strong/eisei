// ============================================
// CelebrateNotifier.js
// 当月シート（PA_CURRENT_MONTH_SHEET_ID）のG列「着金額」が
// 0/空 → 正の数 に変わったら、line-ai-bot の /api/celebrate を叩いて
// すべかがくんが営業グループに祝福メッセージをpushする。
//
// セットアップ手順（一度だけ実行）：
//   1. Script Properties に CELEBRATE_SECRET を登録
//      （line-ai-bot の Vercel環境変数と完全一致させる）
//   2. setupCelebrateTrigger() を実行（onEditトリガーを当月シートに対して作成）
//
// 検出ロジック：
//   - 編集対象 = 当月シート（PA_CURRENT_MONTH_SHEET_ID）
//   - 編集タブ = 営業タブ（PA_CURRENT_MONTH_SKIP_TABS は除外）
//   - 編集セル = G列（着金額、1-indexed=7）
//   - 行 >= PA_CURRENT_MONTH_DATA_START_ROW
//   - 旧値が空/0、新値が正の数（既登録の修正は通知しない）
// ============================================

var CELEBRATE_API_URL = 'https://line-ai-bot2.vercel.app/api/celebrate';

// 当月シートの列構成（1-indexed）:
// A=1 No. / B=2 担当者 / C=3 初回商談日 / D=4 本名 / E=5 LINE名
// F=6 成約状況 / G=7 成約金額 / H=8 着金額 / I=9 クレカ/銀振 ...
var CELEBRATE_AMOUNT_COL = 8;    // H列：着金額
var CELEBRATE_CUSTOMER_COL = 4;  // D列：本名

/**
 * onEdit トリガーから呼ばれる本体。
 * 当月シートのG列に着金額が新規入力された時のみ祝福APIを叩く。
 */
function onEdit_celebrateCheck(e) {
  if (!e || !e.range) return;

  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();

  // 営業タブ以外（全体数値、テンプレ等）はスキップ
  if (PA_CURRENT_MONTH_SKIP_TABS && PA_CURRENT_MONTH_SKIP_TABS[sheetName]) return;

  // 着金額セル以外の編集はスキップ
  if (range.getColumn() !== CELEBRATE_AMOUNT_COL) return;

  // ヘッダー行・サマリ行への編集はスキップ
  var row = range.getRow();
  if (row < PA_CURRENT_MONTH_DATA_START_ROW) return;

  // 範囲編集（複数セル一括ペースト等）はスキップ（e.value が undefined になる）
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;

  var newAmount = parseFloat(e.value);
  if (!Number.isFinite(newAmount) || newAmount <= 0) return;

  // 既存値があった場合（修正）は通知しない
  var oldAmount = parseFloat(e.oldValue);
  if (Number.isFinite(oldAmount) && oldAmount > 0) {
    Logger.log('[celebrate] skip update old=' + oldAmount + ' new=' + newAmount);
    return;
  }

  // 顧客名（本名 = C列）
  var customerName = '';
  try {
    var cv = sheet.getRange(row, CELEBRATE_CUSTOMER_COL).getValue();
    customerName = cv ? String(cv) : '';
  } catch (err) {
    Logger.log('[celebrate] customer read failed: ' + err);
  }

  // 営業名（タブ名 → 表示名マッピング）
  var salesperson = (PA_CURRENT_MONTH_TAB_TO_DISPLAY && PA_CURRENT_MONTH_TAB_TO_DISPLAY[sheetName]) || sheetName;

  // 万円 → 円
  var amountYen = Math.round(newAmount * 10000);

  Logger.log('[celebrate] fire salesperson=' + salesperson + ' amount=' + amountYen + ' customer=' + customerName + ' row=' + row);

  postCelebration_({
    salesperson: salesperson,
    amount: amountYen,
    client: customerName
  });
}

/**
 * /api/celebrate へPOST。失敗してもログだけ残してスプシ操作は継続。
 */
function postCelebration_(payload) {
  var secret = PropertiesService.getScriptProperties().getProperty('CELEBRATE_SECRET');
  if (!secret) {
    Logger.log('[celebrate] CELEBRATE_SECRET is not set in Script Properties');
    return;
  }

  try {
    var res = UrlFetchApp.fetch(CELEBRATE_API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + secret },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log('[celebrate] response status=' + res.getResponseCode() + ' body=' + res.getContentText().slice(0, 300));
  } catch (err) {
    Logger.log('[celebrate] POST failed: ' + err);
  }
}

/**
 * onEditトリガーを当月シートに対してセットアップする（一度だけ実行）。
 * 同じ関数の既存トリガーは削除してから新規作成する（重複回避）。
 */
function setupCelebrateTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEdit_celebrateCheck') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }

  var ss = SpreadsheetApp.openById(PA_CURRENT_MONTH_SHEET_ID);
  ScriptApp.newTrigger('onEdit_celebrateCheck')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log('[celebrate] trigger setup complete. removed=' + removed + ' created=1');
}

/**
 * 動作確認用：実際のスプシ編集なしに祝福APIを叩いてみる。
 * 引数は適当な営業名と金額。
 */
function testCelebrationManual() {
  postCelebration_({
    salesperson: 'ありのまま',
    amount: 1498000,
    client: 'テスト株式会社'
  });
}
