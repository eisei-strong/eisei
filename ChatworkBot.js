// ============================================
// ChatworkBot.js — Chatwork FBぼっと メインロジック
// ============================================

/**
 * メインポーリング関数（2分トリガーで呼ばれる）
 */
function pollAndRespond() {
  var token = getChatworkToken_();
  if (!token) {
    logBotError_('CHATWORK_API_TOKEN未設定');
    return;
  }

  var ss = getSpreadsheet_();
  var processedIds = getProcessedMessageIds_(ss);

  for (var r = 0; r < MONITORED_ROOMS.length; r++) {
    var room = MONITORED_ROOMS[r];
    var messages = getNewMessages_(room.roomId, token);

    if (!messages || messages.length === 0) continue;

    var targets = filterFeedbackWorthy_(messages, processedIds);
    logBotActivity_('Room ' + room.name + ': ' + messages.length + '件取得, ' + targets.length + '件FB対象');

    for (var i = 0; i < targets.length; i++) {
      generateAndPost_(targets[i], room.roomId, token, ss);
    }
  }
}

/**
 * Chatwork APIからメッセージ取得
 * force=0: 前回取得以降の新着のみ
 */
function getNewMessages_(roomId, token) {
  var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages?force=0';
  var options = {
    method: 'get',
    headers: { 'X-ChatWorkToken': token },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code === 204) return [];  // 新着なし
    if (code !== 200) {
      logBotError_('GET messages失敗: room=' + roomId + ' code=' + code);
      return [];
    }

    // レートリミットチェック
    var remaining = parseInt(response.getHeaders()['x-ratelimit-remaining'] || '999');
    if (remaining < CW_RATE_STOP) {
      logBotError_('レートリミット危険: remaining=' + remaining + ' → サイクル中断');
      return [];
    }
    if (remaining < CW_RATE_WARN) {
      logBotActivity_('レートリミット警告: remaining=' + remaining);
    }

    return JSON.parse(response.getContentText());
  } catch (e) {
    logBotError_('GET messages例外: ' + e.message);
    return [];
  }
}

/**
 * フィードバック対象のメッセージをフィルタ
 */
function filterFeedbackWorthy_(messages, processedIds) {
  return messages.filter(function(msg) {
    // ほし自身のメッセージスキップ
    if (String(msg.account.account_id) === BOT_ACCOUNT_ID) return false;

    // 処理済みスキップ
    if (processedIds[String(msg.message_id)]) return false;

    var body = msg.body || '';

    // BOTラベル付きスキップ
    if (body.indexOf(BOT_LABEL) === 0) return false;

    // 嬴政へのTO（メンション）がない場合スキップ
    if (body.indexOf('[To:' + BOT_ACCOUNT_ID + ']') === -1) return false;

    // 無視パターンスキップ
    for (var i = 0; i < IGNORE_PATTERNS_BOT.length; i++) {
      if (IGNORE_PATTERNS_BOT[i].test(body)) return false;
    }

    return true;
  });
}

/**
 * メッセージの種別を判定
 * @param {string} body
 * @returns {string|null}
 */
function detectFeedbackType_(body) {
  if (TRIGGER_PATTERNS.SHOUDAN_REPORT.test(body)) return 'SHOUDAN_REPORT';
  if (TRIGGER_PATTERNS.QUESTION.test(body))       return 'QUESTION';
  if (TRIGGER_PATTERNS.PROPOSAL.test(body))       return 'PROPOSAL';
  if (TRIGGER_PATTERNS.PROBLEM.test(body))        return 'PROBLEM';
  if (TRIGGER_PATTERNS.REQUEST.test(body))         return 'REQUEST';
  if (TRIGGER_PATTERNS.REPORT.test(body))          return 'REPORT';
  if (TRIGGER_PATTERNS.VIDEO.test(body))           return 'VIDEO';
  return null;
}

/**
 * RAG検索→Claude生成→Chatwork投稿
 */
function generateAndPost_(msg, roomId, token, ss) {
  // TO部分を除去してから判定
  var cleanBody = msg.body.replace(/\[To:\d+\]\s*[^\n]*/g, '').trim();
  var feedbackType = detectFeedbackType_(cleanBody) || 'REQUEST';

  // 1. RAGでKBから類似FB検索
  var examples = searchKnowledgeBase_(msg.body, feedbackType, 5);

  // 1.5. YouTube動画があれば字幕を取得
  var videoTranscripts = getVideoTranscripts_(msg.body);
  if (videoTranscripts.length > 0) {
    logBotActivity_('動画検出: ' + videoTranscripts.length + '本の字幕取得 msg_id=' + msg.message_id);
  }

  // 2. Claude APIでFB生成（動画字幕も渡す）
  var feedback = callClaudeForFeedback_(msg, feedbackType, examples, roomId, videoTranscripts);

  if (!feedback || feedback.length < 10) {
    logBotActivity_('FB生成なし: msg_id=' + msg.message_id + ' type=' + feedbackType + ' examples=' + examples.length + ' feedback=' + (feedback || 'null'));
    // 処理済みにはマークする（次回またスキップ）
    logProcessedMessage_(ss, msg, roomId, '(生成スキップ: ' + (feedback || 'null') + ')', feedbackType);
    return;
  }

  // 3. Chatworkに投稿
  var toPrefix = '[To:' + msg.account.account_id + '] ' + msg.account.name + 'さん\n';
  var fullBody = BOT_LABEL + ' ' + toPrefix + '\n' + feedback;

  var postResult = postChatworkMessage_(roomId, fullBody, token);

  // 4. 処理済み記録
  logProcessedMessage_(ss, msg, roomId, feedback, feedbackType);

  // 5. フォローアップ必要なら登録
  if (needsFollowUp_(feedback, feedbackType)) {
    registerFollowUp_(msg, roomId, feedback);
  }

  logBotActivity_('FB投稿完了: to=' + msg.account.name + ' type=' + feedbackType);
}

/**
 * Chatworkにメッセージ投稿
 */
function postChatworkMessage_(roomId, body, token) {
  var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages';
  var options = {
    method: 'post',
    headers: { 'X-ChatWorkToken': token },
    payload: { body: body },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      logBotError_('POST message失敗: room=' + roomId + ' code=' + response.getResponseCode());
    }
    return response;
  } catch (e) {
    logBotError_('POST message例外: ' + e.message);
    return null;
  }
}

// ============================================
// トリガー管理
// ============================================

/**
 * Botポーリング開始
 */
function setupBotTrigger() {
  // 既存のpollAndRespondトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollAndRespond') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 新規トリガー作成（2分間隔）
  ScriptApp.newTrigger('pollAndRespond')
    .timeBased()
    .everyMinutes(CW_POLL_INTERVAL_MIN)
    .create();

  logBotActivity_('Bot開始: ' + CW_POLL_INTERVAL_MIN + '分間隔ポーリング');
  Logger.log('Bot polling trigger set: every ' + CW_POLL_INTERVAL_MIN + ' min');
}

/**
 * Bot停止
 */
function stopBot() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'pollAndRespond' || fn === 'checkFollowUps') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }

  logBotActivity_('Bot停止: ' + removed + '個のトリガー削除');
  Logger.log('Bot stopped: ' + removed + ' triggers removed');
}

/**
 * Bot稼働状況取得
 */
function getBotStatus() {
  var triggers = ScriptApp.getProjectTriggers();
  var botTriggers = [];
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'pollAndRespond' || fn === 'checkFollowUps') {
      botTriggers.push(fn);
    }
  }

  var ss = getSpreadsheet_();
  var logSheet = ss.getSheetByName(SHEET_BOT_LOG);
  var lastLog = '';
  if (logSheet && logSheet.getLastRow() > 1) {
    var row = logSheet.getRange(logSheet.getLastRow(), 1, 1, 3).getValues()[0];
    lastLog = row[0] + ' [' + row[1] + '] ' + row[2];
  }

  return {
    running: botTriggers.length > 0,
    triggers: botTriggers,
    lastLog: lastLog
  };
}

/**
 * フォローアップトリガー設定（毎朝9時）
 */
function setupFollowUpTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkFollowUps') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('checkFollowUps')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  Logger.log('Follow-up trigger set: daily at 9:00');
}

/**
 * フォローアップチェック（毎朝9時）
 */
function checkFollowUps() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_FOLLOWUPS);
  if (!sheet || sheet.getLastRow() < 2) return;

  var token = getChatworkToken_();
  if (!token) return;

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var followUpDate = data[i][3];
    var status = String(data[i][5]);

    if (status === '完了' || status === '送信済') continue;

    if (followUpDate instanceof Date) {
      var fDate = new Date(followUpDate);
      fDate.setHours(0, 0, 0, 0);
      if (fDate <= today) {
        var roomId = String(data[i][1]);
        var accountId = String(data[i][2]);
        var topic = String(data[i][4]).substring(0, 100);

        var reminder = BOT_LABEL + ' [To:' + accountId + ']\n';
        reminder += '⏰ フォローアップリマインド\n\n';
        reminder += '以前の件の進捗はいかがですか？\n';
        reminder += '概要: ' + topic + '\n\n';
        reminder += '報告をお願いします！^^';

        postChatworkMessage_(roomId, reminder, token);
        sheet.getRange(i + 1, 6).setValue('送信済');
      }
    }
  }
}

// ============================================
// ログ管理
// ============================================

/**
 * 処理済みメッセージIDを取得（直近500件）
 */
function getProcessedMessageIds_(ss) {
  var sheet = ss.getSheetByName(SHEET_PROCESSED);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var startRow = Math.max(2, lastRow - 499);
  var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1).getValues();
  var ids = {};
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) ids[String(data[i][0])] = true;
  }
  return ids;
}

/**
 * 処理済みメッセージ記録
 */
function logProcessedMessage_(ss, msg, roomId, feedback, feedbackType) {
  var sheet = ss.getSheetByName(SHEET_PROCESSED);
  if (!sheet) return;

  sheet.appendRow([
    String(msg.message_id),
    new Date(),
    roomId,
    String(msg.account.account_id),
    msg.account.name,
    (msg.body || '').substring(0, 500),
    (feedback || '').substring(0, 500),
    feedbackType || ''
  ]);
}

/** エラーログ */
function logBotError_(message) {
  writeBotLog_('ERROR', message);
}

/** アクティビティログ */
function logBotActivity_(message) {
  writeBotLog_('INFO', message);
}

/** ログ書込み */
function writeBotLog_(level, message) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(SHEET_BOT_LOG);
    if (!sheet) return;

    sheet.appendRow([new Date(), level, message]);

    // 1000行超えたらトリム
    var lastRow = sheet.getLastRow();
    if (lastRow > 1100) {
      sheet.deleteRows(2, 100);
    }
  } catch (e) {
    Logger.log('ログ書込みエラー: ' + e.message);
  }
}
