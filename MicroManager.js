// ============================================
// MicroManager.js — 営業マイクロマネジメント本体
// ============================================
// 専用ルームに提出される 翌朝計画 / 日報 / 週次振り返り を監視
// AI 判定 → 不合格 or 締切超過なら上司口調で詰めDM + ペナ加算
// 月末に集計レポート

// ============================================
// メインポーリング（5分間隔）
// ============================================
/**
 * 専用ルームの新着メッセージを取得し、提出物を判定する
 * AI判定で不合格なら詰め文面を全員ルームに投げる
 */
function pollMicroManager() {
  if (!microRoomReady_()) {
    microBotLog_('INFO', 'MICRO_ROOM_ID 未設定のためスキップ');
    return;
  }

  var token = (typeof getChatworkToken_ === 'function') ? getChatworkToken_() : null;
  if (!token) {
    microBotLog_('ERROR', 'CHATWORK_API_TOKEN 未設定');
    return;
  }

  var ss = getMicroSpreadsheet_();
  var processedIds = microGetProcessedIds_(ss);

  var messages = microFetchMessages_(MICRO_ROOM_ID, token);
  if (!messages || messages.length === 0) {
    microBotLog_('INFO', '新着なし');
    return;
  }

  var targets = messages.filter(function (msg) {
    var body = msg.body || '';
    // 自分（bot）の投稿スキップ
    if (body.indexOf(MICRO_BOT_LABEL) === 0) return false;
    // 既処理スキップ
    if (processedIds[String(msg.message_id)]) return false;
    // 監視対象メンバー以外スキップ
    if (!MICRO_MEMBERS[String(msg.account.account_id)]) return false;
    return true;
  });

  microBotLog_('INFO', '取得 ' + messages.length + '件 / 判定対象 ' + targets.length + '件');

  for (var i = 0; i < targets.length; i++) {
    try {
      microProcessSubmission_(targets[i], token, ss);
    } catch (e) {
      microBotLog_('ERROR', '判定処理例外 msg_id=' + targets[i].message_id + ': ' + e.message);
    }
  }
}

// ============================================
// 1メッセージの判定処理
// ============================================
function microProcessSubmission_(msg, token, ss) {
  var body = msg.body || '';
  var accountId = String(msg.account.account_id);
  var memberName = MICRO_MEMBERS[accountId] || msg.account.name;

  // 種別判定
  var type = microDetectType_(body);
  if (!type) {
    // 種別不明 → 雑談として処理済みマークだけ
    microMarkProcessed_(ss, msg, '(種別不明スキップ)', null, null);
    return;
  }

  // AI 判定
  var verdict = microJudgeSubmission_(type, memberName, body);
  if (!verdict) {
    microBotLog_('ERROR', 'AI判定失敗 msg_id=' + msg.message_id + ' member=' + memberName);
    microMarkProcessed_(ss, msg, '(AI判定失敗)', type, null);
    return;
  }

  // 投稿（[To:] + bot本文）
  var toPrefix = '[To:' + accountId + '] ' + memberName + 'さん\n';
  var fullBody = MICRO_BOT_LABEL + ' ' + toPrefix + verdict.message;

  Utilities.sleep(Math.floor(Math.random() * 90 * 1000));  // 0〜90秒ジッター（bot感低減）
  microPostMessage_(MICRO_ROOM_ID, fullBody, token);

  // 不合格ならペナ加算
  if (verdict.verdict === '不合格') {
    microAddPenalty_(accountId, type);
  }

  // 提出ログ
  microLogSubmission_(ss, msg, type, verdict);
  microMarkProcessed_(ss, msg, verdict.verdict, type, verdict);
}

// ============================================
// 種別判定（ルール式）
// ============================================
function microDetectType_(body) {
  if (!body) return null;
  var b = body.replace(/\[[^\]]+\]/g, '');  // [To:][toall]タグ除去

  if (/翌朝(計画|やる|タスク)|明朝|明日(やる|の.{0,5}(計画|目標|タスク))/.test(b)) return MICRO_TYPE_MORNING;
  if (/【?翌朝計画】?/.test(b)) return MICRO_TYPE_MORNING;
  if (/【?日報】?|本日の(報告|振り返り|結果)|今日の(結果|報告)/.test(b)) return MICRO_TYPE_DAILY;
  if (/【?週次(振り返り|報告)?】?|週報|今週の(振り返り|結果)/.test(b)) return MICRO_TYPE_WEEKLY;

  // 時刻フォールバック（タグなし投稿の救済）
  var hour = new Date().getHours();
  if (hour >= 22 || hour < 3)  return MICRO_TYPE_MORNING;  // 22時〜2時 → 翌朝計画想定
  if (hour >= 18 && hour < 22) return MICRO_TYPE_DAILY;    // 18時〜22時 → 日報想定

  return null;
}

// ============================================
// 締切超過チェック（時刻トリガーから呼ぶ）
// ============================================
/**
 * 翌朝計画の未提出をチェック（毎日 2:00 起動）
 * 締切は当日 1:59:59 → 過去24時間に翌朝計画の合格投稿がない者をペナ
 */
function microCheckMorningPlanDeadline() {
  microCheckDeadline_(MICRO_TYPE_MORNING, function (now) {
    var start = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    return { from: start, to: microMorningPlanDeadline_(now) };
  });
}

/**
 * 日報の未提出をチェック（毎日 21:30 起動）
 * 締切は当日 21:00:00 → 当日中の日報合格投稿がない者をペナ
 */
function microCheckDailyReportDeadline() {
  microCheckDeadline_(MICRO_TYPE_DAILY, function (now) {
    var start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
    return { from: start, to: microDailyReportDeadline_(now) };
  });
}

/**
 * 週次振り返りの未提出をチェック（金曜 18:30 起動）
 */
function microCheckWeeklyReviewDeadline() {
  var dow = new Date().getDay();
  if (dow !== MICRO_DEADLINES.WEEKLY_REVIEW_DOW) {
    microBotLog_('INFO', '週次チェック: 金曜以外のためスキップ dow=' + dow);
    return;
  }
  microCheckDeadline_(MICRO_TYPE_WEEKLY, function (now) {
    // 当週月曜0:00〜金曜18:00
    var start = new Date(now.getTime());
    start.setDate(start.getDate() - (start.getDay() - 1));
    start.setHours(0, 0, 0, 0);
    return { from: start, to: microWeeklyReviewDeadline_(now) };
  });
}

/**
 * 共通: 期間内に type の合格投稿がない者を抽出してペナ + 詰めDM
 */
function microCheckDeadline_(type, windowFn) {
  if (!microRoomReady_()) {
    microBotLog_('INFO', 'MICRO_ROOM_ID 未設定のためスキップ (' + type + ')');
    return;
  }
  var token = (typeof getChatworkToken_ === 'function') ? getChatworkToken_() : null;
  if (!token) {
    microBotLog_('ERROR', 'CHATWORK_API_TOKEN 未設定 (' + type + ')');
    return;
  }

  var now = new Date();
  var range = windowFn(now);
  var ss = getMicroSpreadsheet_();
  var submitted = microFindSubmittersInRange_(ss, type, range.from, range.to);

  var missed = [];
  for (var accountId in MICRO_MEMBERS) {
    if (!submitted[accountId]) {
      missed.push({ accountId: accountId, name: MICRO_MEMBERS[accountId] });
    }
  }

  microBotLog_('INFO', type + ' 締切チェック: 提出 ' + Object.keys(submitted).length +
    '人 / 未提出 ' + missed.length + '人');

  if (missed.length === 0) return;

  for (var i = 0; i < missed.length; i++) {
    var m = missed[i];
    var msgBody = microBuildMissedMessage_(type, m.name);
    var fullBody = MICRO_BOT_LABEL + ' [To:' + m.accountId + '] ' + m.name + 'さん\n' + msgBody;
    microPostMessage_(MICRO_ROOM_ID, fullBody, token);
    microAddPenalty_(m.accountId, type);
    Utilities.sleep(2000);  // 連投で詰まらないように
  }
}

// ============================================
// 月次ペナ集計レポート
// ============================================
function microMonthlyReport() {
  if (!microRoomReady_()) {
    microBotLog_('INFO', 'MICRO_ROOM_ID 未設定のため月次レポートスキップ');
    return;
  }
  var token = (typeof getChatworkToken_ === 'function') ? getChatworkToken_() : null;
  if (!token) { microBotLog_('ERROR', 'CHATWORK_API_TOKEN 未設定'); return; }

  var now = new Date();
  var monthKey = microMonthKey_(now);
  var monthLabel = Utilities.formatDate(now, 'Asia/Tokyo', 'M月');

  var props = PropertiesService.getScriptProperties();
  var penaltyData = JSON.parse(props.getProperty('MICRO_PENA_' + monthKey) || '{}');

  var body = microBuildMonthlyReport_(monthLabel, penaltyData);
  microPostMessage_(MICRO_ROOM_ID, body, token);
  microBotLog_('INFO', '月次ペナレポート送信 ' + monthKey);
}

/**
 * 月末判定: 翌日が1日ならレポート送信
 */
function microMonthlyReportWrapper() {
  var tomorrow = new Date(Date.now() + 24 * 60 * 60 * 1000);
  if (parseInt(Utilities.formatDate(tomorrow, 'Asia/Tokyo', 'd'), 10) === 1) {
    microMonthlyReport();
  }
}

// ============================================
// ペナルティ加算
// ============================================
function microAddPenalty_(accountId, type) {
  // 1日に同種ペナは最大1回（連投でも多重カウントしない）
  var dateKey = microDateKey_();
  var props = PropertiesService.getScriptProperties();
  var dailyKey = 'MICRO_PENA_DAY_' + dateKey + '_' + accountId + '_' + type;
  if (props.getProperty(dailyKey)) {
    microBotLog_('INFO', '同日同種ペナ既加算スキップ ' + accountId + ' ' + type);
    return;
  }
  props.setProperty(dailyKey, '1');

  var monthKey = microMonthKey_();
  var monthlyKey = 'MICRO_PENA_' + monthKey;
  var data = JSON.parse(props.getProperty(monthlyKey) || '{}');
  if (!data[accountId]) data[accountId] = { morning: 0, daily: 0, weekly: 0 };

  if (type === MICRO_TYPE_MORNING) data[accountId].morning++;
  else if (type === MICRO_TYPE_DAILY) data[accountId].daily++;
  else if (type === MICRO_TYPE_WEEKLY) data[accountId].weekly++;

  props.setProperty(monthlyKey, JSON.stringify(data));
}

// ============================================
// Chatwork API
// ============================================
function microFetchMessages_(roomId, token) {
  var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages?force=0';
  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'X-ChatWorkToken': token },
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code === 204) return [];
    if (code !== 200) {
      microBotLog_('ERROR', 'GET messages失敗 code=' + code);
      return [];
    }
    return JSON.parse(res.getContentText());
  } catch (e) {
    microBotLog_('ERROR', 'GET messages例外: ' + e.message);
    return [];
  }
}

function microPostMessage_(roomId, body, token) {
  var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages';
  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { 'X-ChatWorkToken': token },
      payload: { body: body },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      microBotLog_('ERROR', 'POST失敗 code=' + res.getResponseCode() + ' body=' + res.getContentText().substring(0, 200));
    }
    return res;
  } catch (e) {
    microBotLog_('ERROR', 'POST例外: ' + e.message);
    return null;
  }
}

// ============================================
// ログ / 状態管理（スクリプトSPS書込みOK ※マスターではない）
// ============================================
function microGetProcessedIds_(ss) {
  var sheet = ss.getSheetByName(MICRO_SHEET_STATE);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var startRow = Math.max(2, lastRow - 999);
  var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1).getValues();
  var ids = {};
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) ids[String(data[i][0])] = true;
  }
  return ids;
}

function microMarkProcessed_(ss, msg, status, type, verdict) {
  var sheet = ss.getSheetByName(MICRO_SHEET_STATE);
  if (!sheet) {
    sheet = ss.insertSheet(MICRO_SHEET_STATE);
    sheet.getRange(1, 1, 1, 6).setValues([['message_id', 'date', 'account_id', 'name', 'status', 'type']]);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    String(msg.message_id),
    new Date(),
    String(msg.account.account_id),
    msg.account.name,
    status,
    type || ''
  ]);
}

function microLogSubmission_(ss, msg, type, verdict) {
  var sheet = ss.getSheetByName(MICRO_SHEET_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(MICRO_SHEET_LOG);
    sheet.getRange(1, 1, 1, 8).setValues([
      ['timestamp', 'date', 'account_id', 'name', 'type', 'verdict', 'missing', 'body_excerpt']
    ]);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    new Date(),
    microDateKey_(),
    String(msg.account.account_id),
    msg.account.name,
    type,
    verdict.verdict,
    (verdict.missing || []).join(' / '),
    (msg.body || '').substring(0, 300)
  ]);
}

/**
 * 期間内に type の合格投稿をした account_id を抽出
 */
function microFindSubmittersInRange_(ss, type, from, to) {
  var sheet = ss.getSheetByName(MICRO_SHEET_LOG);
  var submitted = {};
  if (!sheet) return submitted;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return submitted;

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  for (var i = 0; i < data.length; i++) {
    var ts = data[i][0];
    var accountId = String(data[i][2]);
    var rowType = data[i][4];
    var verdict = data[i][5];
    if (rowType !== type) continue;
    if (verdict !== '合格') continue;
    if (!(ts instanceof Date)) continue;
    if (ts.getTime() < from.getTime() || ts.getTime() > to.getTime()) continue;
    submitted[accountId] = true;
  }
  return submitted;
}

function microBotLog_(level, message) {
  try {
    var ss = getMicroSpreadsheet_();
    var sheet = ss.getSheetByName(MICRO_SHEET_BOTLOG);
    if (!sheet) {
      sheet = ss.insertSheet(MICRO_SHEET_BOTLOG);
      sheet.getRange(1, 1, 1, 3).setValues([['timestamp', 'level', 'message']]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), level, message]);
    var lastRow = sheet.getLastRow();
    if (lastRow > 1100) sheet.deleteRows(2, 100);
  } catch (e) {
    Logger.log('MicroManager log write error: ' + e.message);
  }
}

// ============================================
// トリガー管理
// ============================================
/**
 * マイクロマネジメント Bot のトリガーを一括設定
 */
function installMicroManagerTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var targetHandlers = [
    'pollMicroManager',
    'microCheckMorningPlanDeadline',
    'microCheckDailyReportDeadline',
    'microCheckWeeklyReviewDeadline',
    'microMonthlyReportWrapper'
  ];
  for (var i = 0; i < triggers.length; i++) {
    if (targetHandlers.indexOf(triggers[i].getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 5分ポーリング
  ScriptApp.newTrigger('pollMicroManager')
    .timeBased().everyMinutes(MICRO_POLL_MIN).create();

  // 翌朝計画 締切チェック: 毎日 2:00
  ScriptApp.newTrigger('microCheckMorningPlanDeadline')
    .timeBased().atHour(2).nearMinute(0).everyDays(1)
    .inTimezone('Asia/Tokyo').create();

  // 日報 締切チェック: 毎日 21:30
  ScriptApp.newTrigger('microCheckDailyReportDeadline')
    .timeBased().atHour(21).nearMinute(30).everyDays(1)
    .inTimezone('Asia/Tokyo').create();

  // 週次 締切チェック: 毎日 18:30（金曜だけ実行）
  ScriptApp.newTrigger('microCheckWeeklyReviewDeadline')
    .timeBased().atHour(18).nearMinute(30).everyDays(1)
    .inTimezone('Asia/Tokyo').create();

  // 月次レポート: 毎日 23:58（月末だけ実行）
  ScriptApp.newTrigger('microMonthlyReportWrapper')
    .timeBased().atHour(23).nearMinute(58).everyDays(1)
    .inTimezone('Asia/Tokyo').create();

  microBotLog_('INFO', 'トリガー設定完了');
  Logger.log('MicroManager triggers installed');
}

/**
 * 停止: 全関連トリガー削除
 */
function uninstallMicroManagerTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var targetHandlers = [
    'pollMicroManager',
    'microCheckMorningPlanDeadline',
    'microCheckDailyReportDeadline',
    'microCheckWeeklyReviewDeadline',
    'microMonthlyReportWrapper'
  ];
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (targetHandlers.indexOf(triggers[i].getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  microBotLog_('INFO', 'トリガー停止: ' + removed + '個削除');
  Logger.log('MicroManager triggers removed: ' + removed);
}

/**
 * 動作確認: 手動1回ポーリング
 */
function testMicroManager() {
  Logger.log('=== MicroManager テスト実行 ===');
  pollMicroManager();
  Logger.log('=== 完了。' + MICRO_SHEET_BOTLOG + ' / ' + MICRO_SHEET_LOG + ' を確認 ===');
}
