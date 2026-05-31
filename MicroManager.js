// ============================================
// MicroManager.js — 営業マイクロマネジメント本体
// ============================================
// 専用ルームの投稿を監視（キーワード判定のみ、AI判定なし）
// 毎日 2:00 に「翌朝計画」「日報」の提出有無をチェック → 未提出ならペナ+詰めDM
// 提出記録は ScriptProperties (MICRO_SUBMISSIONS) に保持（3日分のみ）

// ============================================
// ポーリング（5分間隔）: 投稿を提出記録に保存
// ============================================
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
    if (body.indexOf(MICRO_BOT_LABEL) === 0) return false;
    if (processedIds[String(msg.message_id)]) return false;
    if (!MICRO_MEMBERS[String(msg.account.account_id)]) return false;
    return true;
  });

  microBotLog_('INFO', '取得 ' + messages.length + '件 / 対象 ' + targets.length + '件');

  for (var i = 0; i < targets.length; i++) {
    try {
      microRecordSubmission_(targets[i], ss);
    } catch (e) {
      microBotLog_('ERROR', '記録例外 msg_id=' + targets[i].message_id + ': ' + e.message);
    }
  }
}

/**
 * 投稿1件をキーワード判定 → 提出記録に保存
 */
function microRecordSubmission_(msg, ss) {
  var body = msg.body || '';
  var accountId = String(msg.account.account_id);
  var ts = new Date((msg.send_time || (Date.now() / 1000)) * 1000);
  var workDate = microWorkDateKey_(ts);

  var hasMorning = MICRO_KEYWORDS_MORNING.some(function (kw) { return body.indexOf(kw) >= 0; });
  var hasDaily   = MICRO_KEYWORDS_DAILY.some(function (kw)   { return body.indexOf(kw) >= 0; });

  var detected = [];
  if (hasMorning) detected.push('翌朝計画');
  if (hasDaily)   detected.push('日報');

  if (hasMorning || hasDaily) {
    var subs = microLoadSubmissions_();
    if (!subs[workDate]) subs[workDate] = { morning: {}, daily: {} };
    if (hasMorning) subs[workDate].morning[accountId] = true;
    if (hasDaily)   subs[workDate].daily[accountId] = true;
    microSaveSubmissions_(subs);
    microBotLog_('INFO', '提出記録: ' + (MICRO_MEMBERS[accountId] || accountId) +
      ' [' + detected.join(',') + '] workDate=' + workDate);
  }

  microMarkProcessed_(ss, msg, detected.length ? detected.join(',') : '(キーワード無し)');
}

// ============================================
// 提出記録の永続化（ScriptProperties）
// ============================================
function microLoadSubmissions_() {
  var raw = PropertiesService.getScriptProperties().getProperty('MICRO_SUBMISSIONS') || '{}';
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

function microSaveSubmissions_(subs) {
  // 古い分（3日以上前）を削除して肥大化防止
  var cutoff = new Date(Date.now() - 4 * 24 * 60 * 60 * 1000);
  var cutoffKey = Utilities.formatDate(cutoff, 'Asia/Tokyo', 'yyyy-MM-dd');
  for (var k in subs) {
    if (k < cutoffKey) delete subs[k];
  }
  PropertiesService.getScriptProperties().setProperty('MICRO_SUBMISSIONS', JSON.stringify(subs));
}

// ============================================
// 締切チェック（毎日 2:00 トリガー）
// ============================================
/**
 * 直前1営業日の「翌朝計画」未提出をチェック → ペナ + 詰めDM
 */
function microCheckMorningPlanDeadline() {
  microCheckDeadline_(MICRO_TYPE_MORNING);
}

/**
 * 直前1営業日の「日報」未提出をチェック → ペナ + 詰めDM
 */
function microCheckDailyReportDeadline() {
  microCheckDeadline_(MICRO_TYPE_DAILY);
}

function microCheckDeadline_(type) {
  if (!microRoomReady_()) {
    microBotLog_('INFO', 'MICRO_ROOM_ID 未設定のためスキップ (' + type + ')');
    return;
  }
  var token = (typeof getChatworkToken_ === 'function') ? getChatworkToken_() : null;
  if (!token) {
    microBotLog_('ERROR', 'CHATWORK_API_TOKEN 未設定 (' + type + ')');
    return;
  }

  // チェック対象営業日 = 4時間前の workDate（例: 2:00 起動 → 2:00 - 4h = 22:00 前日 → 前日キー）
  var targetKey = microWorkDateKey_(new Date(Date.now() - 4 * 60 * 60 * 1000));
  var subs = microLoadSubmissions_();
  var bucket = subs[targetKey] || { morning: {}, daily: {} };
  var submitted = (type === MICRO_TYPE_MORNING) ? bucket.morning : bucket.daily;

  var missed = [];
  for (var accountId in MICRO_MEMBERS) {
    if (!submitted[accountId]) {
      missed.push({ accountId: accountId, name: MICRO_MEMBERS[accountId] });
    }
  }

  microBotLog_('INFO', type + ' 締切チェック ' + targetKey + ': 提出 ' +
    Object.keys(submitted).length + '人 / 未提出 ' + missed.length + '人');

  if (missed.length === 0) return;

  for (var i = 0; i < missed.length; i++) {
    var m = missed[i];
    var msgBody = microBuildMissedMessage_(type, m.name);
    var fullBody = MICRO_BOT_LABEL + ' [To:' + m.accountId + '] ' + m.name + 'さん\n' + msgBody;
    microPostMessage_(MICRO_ROOM_ID, fullBody, token);
    microAddPenalty_(m.accountId, type);
    Utilities.sleep(2000);
  }
}

function microBuildMissedMessage_(type, memberName) {
  if (type === MICRO_TYPE_MORNING) {
    return [
      '2:00。翌朝計画の提出なし',
      '締切は深夜1:59',
      '今日のアポ件数・最優先タスク・気合入れる商談、3点セットで即出せ',
      '※「翌朝計画」のキーワードを含めて投稿すること',
      'ペナ1（1万円）'
    ].join('\n');
  }
  if (type === MICRO_TYPE_DAILY) {
    return [
      '2:00。日報の提出なし',
      '締切は深夜1:59',
      '今日の数字も明日の打ち手も投げっぱなしか',
      '実数・原因・明日の打ち手・朝の宣言への結果、4点で出せ',
      '※「日報」のキーワードを含めて投稿すること',
      'ペナ1（1万円）'
    ].join('\n');
  }
  return '提出物なし\nペナ1（1万円）';
}

// ============================================
// ペナルティ加算
// ============================================
function microAddPenalty_(accountId, type) {
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
  if (!data[accountId]) data[accountId] = { morning: 0, daily: 0 };
  if (type === MICRO_TYPE_MORNING) data[accountId].morning++;
  else if (type === MICRO_TYPE_DAILY) data[accountId].daily++;
  props.setProperty(monthlyKey, JSON.stringify(data));
}

// ============================================
// ワンショットセットアップ
// ============================================
/**
 * 1. Bot参加ルームを取得 → 「マイクロ」を含むルームを自動検出して MICRO_ROOM_ID を保存
 * 2. 営業全員ルーム(rid349937583)からメンバー一覧を取得 → 未登録 account_id をログ出力
 * 3. installMicroManagerTriggers() でトリガー一括設定
 * 4. testMicroManager() でテスト投稿
 */
function microSetupAll() {
  var token = (typeof getChatworkToken_ === 'function') ? getChatworkToken_() : null;
  if (!token) {
    Logger.log('CHATWORK_API_TOKEN が ScriptProperties に未設定。先に設定してください。');
    microBotLog_('ERROR', 'setup: token未設定');
    return;
  }

  var rooms = microFetchBotRooms_(token);
  if (!rooms) {
    Logger.log('rooms取得失敗。ネットワーク or token を確認');
    return;
  }
  var candidates = rooms.filter(function (r) {
    var name = String(r.name || '');
    return name.indexOf('マイクロ') >= 0 || name.indexOf('micro') >= 0 || name.indexOf('Micro') >= 0;
  });

  var current = PropertiesService.getScriptProperties().getProperty('MICRO_ROOM_ID') || '';
  var resolvedRoomId = '';

  if (candidates.length === 1) {
    resolvedRoomId = String(candidates[0].room_id);
    PropertiesService.getScriptProperties().setProperty('MICRO_ROOM_ID', resolvedRoomId);
    Logger.log('✅ ルーム自動検出: ' + candidates[0].name + ' (roomId=' + resolvedRoomId + ')');
    microBotLog_('INFO', 'setup: roomId自動検出=' + resolvedRoomId + ' (' + candidates[0].name + ')');
  } else if (candidates.length === 0) {
    Logger.log('⚠️ 「マイクロ」を含むルームが見つかりません。');
    Logger.log('--- 直近30ルーム一覧（手動で選んでください） ---');
    rooms.sort(function (a, b) { return (b.last_update_time || 0) - (a.last_update_time || 0); });
    for (var i = 0; i < Math.min(rooms.length, 30); i++) {
      Logger.log('  microSetRoomId(\'' + rooms[i].room_id + '\')  // ' + rooms[i].name);
    }
    if (!current) {
      microBotLog_('WARN', 'setup: roomId検出失敗。ログから microSetRoomId() を実行してください');
      return;
    }
    resolvedRoomId = current;
    Logger.log('既存ScriptProperty MICRO_ROOM_ID=' + current + ' を使用して継続');
  } else {
    Logger.log('⚠️ 「マイクロ」を含むルームが複数 (' + candidates.length + '件)。下記から選択:');
    for (var j = 0; j < candidates.length; j++) {
      Logger.log('  microSetRoomId(\'' + candidates[j].room_id + '\')  // ' + candidates[j].name);
    }
    if (!current) return;
    resolvedRoomId = current;
  }

  MICRO_ROOM_ID = resolvedRoomId;

  try {
    var teamMembers = microFetchRoomMembers_(349937583, token);
    if (teamMembers) {
      var unmapped = [];
      for (var k = 0; k < teamMembers.length; k++) {
        var m = teamMembers[k];
        if (!MICRO_MEMBERS[String(m.account_id)]) unmapped.push(m);
      }
      if (unmapped.length > 0) {
        Logger.log('--- 営業全員ルームに居て MICRO_MEMBERS 未登録 ---');
        for (var l = 0; l < unmapped.length; l++) {
          Logger.log("  '" + unmapped[l].account_id + "': '" + (unmapped[l].name || '?') + "',");
        }
        microBotLog_('INFO', 'setup: 未登録メンバー ' + unmapped.length + '人');
      } else {
        Logger.log('✅ 営業全員ルームのメンバー全員が MICRO_MEMBERS に登録済み');
      }
    }
  } catch (e) {
    Logger.log('メンバー抽出スキップ: ' + e.message);
  }

  try {
    installMicroManagerTriggers();
    Logger.log('✅ トリガー設定完了');
  } catch (e) {
    Logger.log('❌ トリガー設定失敗: ' + e.message);
    return;
  }

  try {
    testMicroManager();
    Logger.log('✅ テスト実行');
  } catch (e) {
    Logger.log('テストスキップ: ' + e.message);
  }

  Logger.log('===== microSetupAll 完了 =====');
  Logger.log('MICRO_ROOM_ID = ' + resolvedRoomId);
}

function microSetRoomId(roomId) {
  PropertiesService.getScriptProperties().setProperty('MICRO_ROOM_ID', String(roomId));
  MICRO_ROOM_ID = String(roomId);
  Logger.log('✅ MICRO_ROOM_ID を保存: ' + roomId);
  microBotLog_('INFO', 'roomId手動設定=' + roomId);
}

function microFetchBotRooms_(token) {
  try {
    var res = UrlFetchApp.fetch('https://api.chatwork.com/v2/rooms', {
      method: 'get',
      headers: { 'X-ChatWorkToken': token },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      microBotLog_('ERROR', 'rooms取得失敗 code=' + res.getResponseCode());
      return null;
    }
    return JSON.parse(res.getContentText());
  } catch (e) {
    microBotLog_('ERROR', 'rooms取得例外: ' + e.message);
    return null;
  }
}

function microFetchRoomMembers_(roomId, token) {
  try {
    var res = UrlFetchApp.fetch('https://api.chatwork.com/v2/rooms/' + roomId + '/members', {
      method: 'get',
      headers: { 'X-ChatWorkToken': token },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return null;
    return JSON.parse(res.getContentText());
  } catch (e) {
    return null;
  }
}

function microListBotRooms_() {
  var token = (typeof getChatworkToken_ === 'function') ? getChatworkToken_() : null;
  if (!token) { microBotLog_('ERROR', 'token未設定'); return; }
  var rooms = microFetchBotRooms_(token);
  if (!rooms) return;
  rooms.sort(function (a, b) { return (b.last_update_time || 0) - (a.last_update_time || 0); });
  for (var i = 0; i < Math.min(rooms.length, 30); i++) {
    Logger.log('roomId=' + rooms[i].room_id + '  ' + (rooms[i].name || '(no name)'));
  }
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
      microBotLog_('ERROR', 'POST失敗 code=' + res.getResponseCode() +
        ' body=' + res.getContentText().substring(0, 200));
    }
    return res;
  } catch (e) {
    microBotLog_('ERROR', 'POST例外: ' + e.message);
    return null;
  }
}

// ============================================
// 状態管理 / ログ
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

function microMarkProcessed_(ss, msg, status) {
  var sheet = ss.getSheetByName(MICRO_SHEET_STATE);
  if (!sheet) {
    sheet = ss.insertSheet(MICRO_SHEET_STATE);
    sheet.getRange(1, 1, 1, 5).setValues([['message_id', 'date', 'account_id', 'name', 'status']]);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    String(msg.message_id),
    new Date(),
    String(msg.account.account_id),
    msg.account.name,
    status
  ]);
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
function installMicroManagerTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var targetHandlers = [
    'pollMicroManager',
    'microCheckMorningPlanDeadline',
    'microCheckDailyReportDeadline',
    // 旧バージョンの残骸も削除対象に
    'microCheckWeeklyReviewDeadline',
    'microKpiCheck1700',
    'microKpiCheckWednesday',
    'microKpiCheckMonthly',
    'microMonthlyReportWrapper'
  ];
  for (var i = 0; i < triggers.length; i++) {
    if (targetHandlers.indexOf(triggers[i].getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 5分ポーリング（提出記録）
  ScriptApp.newTrigger('pollMicroManager')
    .timeBased().everyMinutes(MICRO_POLL_MIN).create();

  // 翌朝計画 締切チェック: 毎日 2:00
  ScriptApp.newTrigger('microCheckMorningPlanDeadline')
    .timeBased().atHour(2).nearMinute(0).everyDays(1)
    .inTimezone('Asia/Tokyo').create();

  // 日報 締切チェック: 毎日 2:00（翌朝計画と同タイミング）
  ScriptApp.newTrigger('microCheckDailyReportDeadline')
    .timeBased().atHour(2).nearMinute(5).everyDays(1)
    .inTimezone('Asia/Tokyo').create();

  microBotLog_('INFO', 'トリガー設定完了 (3個)');
  Logger.log('MicroManager triggers installed (3 triggers)');
}

function uninstallMicroManagerTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var targetHandlers = [
    'pollMicroManager',
    'microCheckMorningPlanDeadline',
    'microCheckDailyReportDeadline',
    'microCheckWeeklyReviewDeadline',
    'microKpiCheck1700',
    'microKpiCheckWednesday',
    'microKpiCheckMonthly',
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
  Logger.log('=== 完了。' + MICRO_SHEET_BOTLOG + ' を確認 ===');
}
