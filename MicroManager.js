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
// KPI リアルタイム警告（ペナなし・詰めDMのみ）
// ============================================
/**
 * 17:00: 当日アポ0のメンバーに警告
 * トリガー: 毎日 17:00
 */
function microKpiCheck1700() {
  if (!MICRO_KPI_ENABLED) return;
  if (!microRoomReady_()) return;
  var token = (typeof getChatworkToken_ === 'function') ? getChatworkToken_() : null;
  if (!token) { microBotLog_('ERROR', '17時警告: token未設定'); return; }

  var api = microFetchDashboardApi_();
  if (!api || !api.dailyPushes || !api.dailyPushes.byMember) {
    microBotLog_('ERROR', '17時警告: API取得失敗');
    return;
  }

  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var todayArr = api.dailyPushes.byMember[today] || [];
  var countMap = {};
  for (var i = 0; i < todayArr.length; i++) {
    countMap[todayArr[i].name] = todayArr[i].count;
  }

  var triggered = 0;
  for (var accountId in MICRO_MEMBERS) {
    var displayName = MICRO_MEMBERS[accountId];
    var v2Name = MICRO_DISPLAY_TO_V2[displayName] || displayName;
    var apoCount = countMap[v2Name] || 0;
    if (apoCount > MICRO_KPI_DAILY_THRESHOLD) continue;

    var body = MICRO_BOT_LABEL + ' [To:' + accountId + '] ' + displayName + 'さん\n' +
      '17時 今日のアポ ' + apoCount + '\n' +
      '残り5時間で何やる\n' +
      '今日の最終アポ数 + 明日リカバリの動き、30分以内に返事\n' +
      '※KPI警告（ペナなし）';
    microPostMessage_(MICRO_ROOM_ID, body, token);
    triggered++;
    Utilities.sleep(2000);
  }
  microBotLog_('INFO', '17時KPI警告: ' + triggered + '人に発火');
}

/**
 * 水曜12:00: 週初〜火曜のアポ合計が閾値以下の人に警告
 */
function microKpiCheckWednesday() {
  if (!MICRO_KPI_ENABLED) return;
  if (!microRoomReady_()) return;
  if (new Date().getDay() !== 3) {  // 水曜=3
    microBotLog_('INFO', '週次KPI: 水曜以外のためスキップ');
    return;
  }
  var token = (typeof getChatworkToken_ === 'function') ? getChatworkToken_() : null;
  if (!token) { microBotLog_('ERROR', '週次KPI: token未設定'); return; }

  var api = microFetchDashboardApi_();
  if (!api || !api.dailyPushes || !api.dailyPushes.byMember) {
    microBotLog_('ERROR', '週次KPI: API取得失敗'); return;
  }

  var now = new Date();
  // 月曜0:00〜火曜23:59のキー2日分を合算
  var monday = new Date(now.getTime());
  monday.setDate(monday.getDate() - (monday.getDay() === 0 ? 6 : monday.getDay() - 1));
  monday.setHours(0, 0, 0, 0);
  var tuesday = new Date(monday.getTime() + 24 * 60 * 60 * 1000);
  var monKey = Utilities.formatDate(monday,  'Asia/Tokyo', 'yyyy-MM-dd');
  var tueKey = Utilities.formatDate(tuesday, 'Asia/Tokyo', 'yyyy-MM-dd');

  var sumByName = {};
  [monKey, tueKey].forEach(function (k) {
    var arr = api.dailyPushes.byMember[k] || [];
    for (var i = 0; i < arr.length; i++) {
      sumByName[arr[i].name] = (sumByName[arr[i].name] || 0) + (arr[i].count || 0);
    }
  });

  var triggered = 0;
  for (var accountId in MICRO_MEMBERS) {
    var displayName = MICRO_MEMBERS[accountId];
    var v2Name = MICRO_DISPLAY_TO_V2[displayName] || displayName;
    var weekCount = sumByName[v2Name] || 0;
    if (weekCount > MICRO_KPI_WEEKLY_THRESHOLD) continue;

    var body = MICRO_BOT_LABEL + ' [To:' + accountId + '] ' + displayName + 'さん\n' +
      '水曜昼 今週ここまで月火合計アポ ' + weekCount + '\n' +
      'このペースだと週末詰む\n' +
      '今週どう巻き返すか、本数とどこに架けるか、30分以内に返事\n' +
      '※KPI警告（ペナなし）';
    microPostMessage_(MICRO_ROOM_ID, body, token);
    triggered++;
    Utilities.sleep(2000);
  }
  microBotLog_('INFO', '水曜KPI警告: ' + triggered + '人に発火');
}

/**
 * 月の21日以降: 当月着金が閾値以下の人に警告
 * トリガー: 毎日 12:00（21日以降のみ実行）
 */
function microKpiCheckMonthly() {
  if (!MICRO_KPI_ENABLED) return;
  if (!microRoomReady_()) return;
  var day = new Date().getDate();
  if (day < 21) {
    microBotLog_('INFO', '月次KPI: 21日未満のためスキップ day=' + day);
    return;
  }
  var token = (typeof getChatworkToken_ === 'function') ? getChatworkToken_() : null;
  if (!token) { microBotLog_('ERROR', '月次KPI: token未設定'); return; }

  var api = microFetchDashboardApi_();
  if (!api || !api.members) { microBotLog_('ERROR', '月次KPI: API取得失敗'); return; }

  var revByName = {};
  for (var i = 0; i < api.members.length; i++) {
    revByName[api.members[i].name] = api.members[i].revenue || 0;
  }

  var triggered = 0;
  for (var accountId in MICRO_MEMBERS) {
    var displayName = MICRO_MEMBERS[accountId];
    var v2Name = MICRO_DISPLAY_TO_V2[displayName] || displayName;
    var rev = revByName[v2Name] || 0;
    if (rev > MICRO_KPI_MONTHLY_REVENUE_THRESHOLD) continue;

    var body = MICRO_BOT_LABEL + ' [To:' + accountId + '] ' + displayName + 'さん\n' +
      day + '日 今月着金 ' + rev + '万円\n' +
      '残り10日で何取りに行く\n' +
      '残り商談の確度と着金見込み、1時間以内に返事\n' +
      '※KPI警告（ペナなし）';
    microPostMessage_(MICRO_ROOM_ID, body, token);
    triggered++;
    Utilities.sleep(2000);
  }
  microBotLog_('INFO', '月次KPI警告(' + day + '日): ' + triggered + '人に発火');
}

/**
 * ダッシュボードAPIを取得
 */
function microFetchDashboardApi_() {
  var url = 'https://giver.work/sales-dashboard/api-proxy.php?action=api';
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      microBotLog_('ERROR', 'API HTTP code=' + res.getResponseCode());
      return null;
    }
    return JSON.parse(res.getContentText());
  } catch (e) {
    microBotLog_('ERROR', 'API取得例外: ' + e.message);
    return null;
  }
}

// ============================================
// ワンショットセットアップ
// ============================================
/**
 * これ1つでセットアップ完了:
 *   1. Bot参加ルームを取得 → 「マイクロ」を含むルームを自動検出して MICRO_ROOM_ID を ScriptProperties に保存
 *   2. 営業全員ルーム(rid349937583)からメンバー一覧を取得 → MICRO_MEMBERS に未登録の account_id をログ出力
 *   3. installMicroManagerTriggers() でトリガー一括設定
 *   4. testMicroManager() でテスト投稿
 * 失敗時はログで指示を出すので、それに従って microSetRoomId('xxx') を実行
 */
function microSetupAll() {
  var token = (typeof getChatworkToken_ === 'function') ? getChatworkToken_() : null;
  if (!token) {
    var msg1 = 'CHATWORK_API_TOKEN が ScriptProperties に未設定。先に設定してください。';
    Logger.log(msg1); microBotLog_('ERROR', 'setup: ' + msg1);
    return;
  }

  // --- Step 1: roomId 自動検出 ---
  var rooms = microFetchBotRooms_(token);
  if (!rooms) {
    Logger.log('rooms取得失敗。ネットワーク or token を確認');
    return;
  }
  // 名前に「マイクロ」を含むルームを候補に
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
    Logger.log('⚠️ 「マイクロ」を含むルームが見つかりません。bot がルームに招待されていない可能性。');
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
    Logger.log('⚠️ 「マイクロ」を含むルームが複数 (' + candidates.length + '件)。下記から選んで microSetRoomId() を実行:');
    for (var j = 0; j < candidates.length; j++) {
      Logger.log('  microSetRoomId(\'' + candidates[j].room_id + '\')  // ' + candidates[j].name);
    }
    if (!current) return;
    resolvedRoomId = current;
  }

  // GAS グローバル var の再評価のため、MICRO_ROOM_ID を上書き
  MICRO_ROOM_ID = resolvedRoomId;

  // --- Step 2: 新メンバー account_id 抽出 ---
  try {
    var teamMembers = microFetchRoomMembers_(349937583, token);
    if (teamMembers) {
      var known = MICRO_MEMBERS;
      var unmapped = [];
      for (var k = 0; k < teamMembers.length; k++) {
        var m = teamMembers[k];
        if (!known[String(m.account_id)]) {
          unmapped.push(m);
        }
      }
      if (unmapped.length > 0) {
        Logger.log('--- 営業全員ルームに居て MICRO_MEMBERS 未登録のメンバー ---');
        Logger.log('必要なら MicroManagerConfig.js の MICRO_MEMBERS に追記してください:');
        for (var l = 0; l < unmapped.length; l++) {
          Logger.log("  '" + unmapped[l].account_id + "': '" + (unmapped[l].name || '?') + "',");
        }
        microBotLog_('INFO', 'setup: 未登録メンバー ' + unmapped.length + '人 (詳細はLogger)');
      } else {
        Logger.log('✅ 営業全員ルームのメンバー全員が MICRO_MEMBERS に登録済み');
      }
    }
  } catch (e) {
    Logger.log('メンバー抽出スキップ: ' + e.message);
  }

  // --- Step 3: トリガー設定 ---
  try {
    installMicroManagerTriggers();
    Logger.log('✅ トリガー設定完了');
  } catch (e) {
    Logger.log('❌ トリガー設定失敗: ' + e.message);
    microBotLog_('ERROR', 'setup: トリガー設定失敗 ' + e.message);
    return;
  }

  // --- Step 4: テスト投稿 ---
  try {
    if (typeof testMicroManager === 'function') {
      testMicroManager();
      Logger.log('✅ テスト投稿実行');
    }
  } catch (e) {
    Logger.log('テスト投稿スキップ: ' + e.message);
  }

  Logger.log('===== microSetupAll 完了 =====');
  Logger.log('MICRO_ROOM_ID = ' + resolvedRoomId);
  Logger.log('Chatworkルームを確認してテストメッセージが投稿されていればOK');
}

/**
 * 手動で MICRO_ROOM_ID を設定（コード変更不要）
 */
function microSetRoomId(roomId) {
  PropertiesService.getScriptProperties().setProperty('MICRO_ROOM_ID', String(roomId));
  MICRO_ROOM_ID = String(roomId);
  Logger.log('✅ MICRO_ROOM_ID を ScriptProperties に保存: ' + roomId);
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

// ============================================
// ルームID解決ヘルパー（手動用）
// ============================================
/**
 * Bot が参加している全ルームを一覧表示（Logger + ログシート）
 * MICRO_ROOM_ID 設定時の参照用。bot を招待後にこれを実行 → ログから roomId を取得
 */
function microListBotRooms_() {
  var token = (typeof getChatworkToken_ === 'function') ? getChatworkToken_() : null;
  if (!token) { microBotLog_('ERROR', 'token未設定'); return; }

  try {
    var res = UrlFetchApp.fetch('https://api.chatwork.com/v2/rooms', {
      method: 'get',
      headers: { 'X-ChatWorkToken': token },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      microBotLog_('ERROR', 'rooms取得失敗 code=' + res.getResponseCode());
      return;
    }
    var rooms = JSON.parse(res.getContentText());
    rooms.sort(function (a, b) { return (b.last_update_time || 0) - (a.last_update_time || 0); });
    var lines = ['===== Bot参加ルーム一覧（更新順） ====='];
    for (var i = 0; i < Math.min(rooms.length, 30); i++) {
      var r = rooms[i];
      lines.push('roomId=' + r.room_id + '  ' + (r.name || '(no name)'));
    }
    var msg = lines.join('\n');
    Logger.log(msg);
    microBotLog_('INFO', msg);
  } catch (e) {
    microBotLog_('ERROR', 'rooms取得例外: ' + e.message);
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

  // KPI: 17:00 当日アポ0警告（毎日）
  ScriptApp.newTrigger('microKpiCheck1700')
    .timeBased().atHour(17).nearMinute(0).everyDays(1)
    .inTimezone('Asia/Tokyo').create();

  // KPI: 水曜12:00 週次警告（毎日実行・水曜のみ発火）
  ScriptApp.newTrigger('microKpiCheckWednesday')
    .timeBased().atHour(12).nearMinute(0).everyDays(1)
    .inTimezone('Asia/Tokyo').create();

  // KPI: 月次警告（毎日12:30、21日以降のみ発火）
  ScriptApp.newTrigger('microKpiCheckMonthly')
    .timeBased().atHour(12).nearMinute(30).everyDays(1)
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
  Logger.log('=== 完了。' + MICRO_SHEET_BOTLOG + ' / ' + MICRO_SHEET_LOG + ' を確認 ===');
}
