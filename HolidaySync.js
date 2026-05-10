// ============================================
// HolidaySync.js
// Googleカレンダーから営業の休日を取得し、
// マスター「休日データ」タブに書き出し + ダッシュボード PHP に POST送信
// ============================================

var HOLIDAY_CAL_ID     = 'namakabackoffice@gmail.com';   // 要確認: 営業休日が入ってるカレンダーID
var HOLIDAY_TAB        = '休日データ';
var HOLIDAY_PUSH_SECRET = 'gas_holiday_push_2026';
var HOLIDAY_PHP_URL    = 'https://giver.work/sales-dashboard/api-proxy.php';

/**
 * メインエントリ: 1時間トリガー想定
 *
 * 動作:
 *   1. CalendarApp で当月の休日イベント取得
 *   2. タイトル解析（メンバー名 + 半休/全休判定）
 *   3. マスター「休日データ」タブに書き出し（全クリア→当月のみ）
 *   4. PHP api-proxy.php に updateHoliday POST → holiday-data.json 更新
 */
function writeHolidayDataToSheet() {
  Logger.log('=== writeHolidayDataToSheet 開始: ' + new Date().toISOString());

  var cal = CalendarApp.getCalendarById(HOLIDAY_CAL_ID);
  if (!cal) {
    Logger.log('カレンダー取得失敗: ' + HOLIDAY_CAL_ID);
    return;
  }

  var now = new Date();
  var month = now.getMonth() + 1;
  var year  = now.getFullYear();

  var monthStart = new Date(year, month - 1, 1);
  var monthEnd   = new Date(year, month, 0, 23, 59, 59);
  var events = cal.getEvents(monthStart, monthEnd);
  Logger.log('当月イベント数: ' + events.length);

  // パース → byDate / counts
  var byDate = {};
  var counts = {};
  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    var parsed = parseHolidayTitle_(ev.getTitle());
    if (!parsed) continue;
    var dateKey = Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', 'yyyy-MM-dd');
    if (!byDate[dateKey]) byDate[dateKey] = [];
    byDate[dateKey].push({ name: parsed.name, type: parsed.type });
    var weight = (parsed.type === 'half') ? 0.5 : 1;
    counts[parsed.name] = (counts[parsed.name] || 0) + weight;
  }

  // マスター「休日データ」タブに書き出し
  var ss = SpreadsheetApp.openById(PA_MASTER_ID);
  var sheet = ss.getSheetByName(HOLIDAY_TAB);
  if (!sheet) sheet = ss.insertSheet(HOLIDAY_TAB);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['date', 'name', 'type']]);
  var rows = [];
  for (var dk in byDate) {
    for (var j = 0; j < byDate[dk].length; j++) {
      var h = byDate[dk][j];
      // 既存仕様: ISO形式で書き出し（dk + 'T15:00:00Z' = JST 0時）
      rows.push([new Date(dk + 'T15:00:00Z'), h.name, h.type]);
    }
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  Logger.log('「休日データ」タブに ' + rows.length + ' 行書き出し');

  // PHP POST
  var payload = {
    action: 'updateHoliday',
    secret: HOLIDAY_PUSH_SECRET,
    month: month,
    year: year,
    byDate: byDate,
    counts: counts
  };
  try {
    var resp = UrlFetchApp.fetch(HOLIDAY_PHP_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log('PHP POST status=' + resp.getResponseCode() + ' / body=' + resp.getContentText().substring(0, 200));
  } catch (e) {
    Logger.log('PHP POST例外: ' + e.message);
  }

  Logger.log('=== writeHolidayDataToSheet 完了');
}

/**
 * イベントタイトルから「メンバー名」と「タイプ(full/half)」を抽出
 * 想定タイトル例:
 *   - "1日1more 休み"     → name=1日1more, type=full
 *   - "ヒトコト 半休"      → name=ヒトコト, type=half
 *   - "ぜんぶり"          → name=ぜんぶり, type=full
 */
function parseHolidayTitle_(title) {
  var t = String(title || '').trim();
  if (!t) return null;

  var type = 'full';
  if (/半休|半日|half/i.test(t)) type = 'half';

  // 「休み」「半休」「半日」「休」「休暇」を除去 → メンバー名抽出
  var name = t.replace(/[\s　]*(半休|半日|休み|休暇|休)[\s　]*/g, '').trim();
  if (!name) return null;
  return { name: name, type: type };
}

/**
 * 1時間トリガーを設定（既存の同名トリガーを削除してから登録）
 */
function installHolidayTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var deleted = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'writeHolidayDataToSheet') {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted++;
    }
  }
  Logger.log('既存 writeHolidayDataToSheet トリガーを ' + deleted + ' 件削除');
  ScriptApp.newTrigger('writeHolidayDataToSheet')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('writeHolidayDataToSheet を1時間ごとに設定しました');
}

/**
 * デバッグ用: カレンダーの当月イベント一覧をログ出力
 * カレンダーIDが正しいか・タイトル形式を確認するため
 */
function debugHolidayCalendar() {
  var cal = CalendarApp.getCalendarById(HOLIDAY_CAL_ID);
  if (!cal) {
    Logger.log('カレンダー見つかりません: ' + HOLIDAY_CAL_ID);
    // 全アクセス可能カレンダー一覧出力（IDの確認用）
    var allCals = CalendarApp.getAllCalendars();
    Logger.log('=== アクセス可能なカレンダー一覧 ===');
    for (var i = 0; i < allCals.length; i++) {
      Logger.log('  name=' + allCals[i].getName() + ' / id=' + allCals[i].getId());
    }
    return;
  }
  var now = new Date();
  var monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  var monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
  var events = cal.getEvents(monthStart, monthEnd);
  Logger.log('=== 当月イベント (' + events.length + '件) ===');
  for (var j = 0; j < events.length; j++) {
    Logger.log('  ' + events[j].getTitle() + ' | ' + events[j].getStartTime());
  }
}

/**
 * PHP連携を確認するためのGASダッシュボードPHP取得テスト
 */
function debugHolidayPhpFetch() {
  var resp = UrlFetchApp.fetch(HOLIDAY_PHP_URL + '?action=api', { muteHttpExceptions: true });
  var data = JSON.parse(resp.getContentText());
  Logger.log('現在のholiday counts:');
  if (data.members) {
    for (var i = 0; i < data.members.length; i++) {
      Logger.log('  ' + data.members[i].name + ': ' + (data.members[i].holidays || 0) + '日');
    }
  }
}
