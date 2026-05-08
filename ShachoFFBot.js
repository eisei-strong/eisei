// ============================================
// ShachoFFBot.js — 社長FFbot メインロジック
// ============================================
// rid382597791 (SMC営業チーム FF部屋) を5分間隔で監視し、
// AI政 (10144671) へのメンションに対して、 14,006件のKnowledge を元に
// GPTs同等品質の社長FFbot として自動応答する

/**
 * メイン ポーリング関数 (5分間隔トリガーで実行)
 */
function pollShachoFFBot() {
  var token = getShachoChatworkToken_();
  if (!token) {
    logShachoBotError_('CHATWORK_BOT_TOKEN 未設定');
    return;
  }

  var apiKey = getClaudeApiKey_();
  if (!apiKey) {
    logShachoBotError_('CLAUDE_API_KEY 未設定');
    return;
  }

  var ss = getShachoSpreadsheet_();
  var processedIds = getShachoProcessedIds_(ss);

  var roomId = SHACHO_FF_TARGET_ROOM;
  var messages = getNewMessages_(roomId, token);

  if (!messages || messages.length === 0) {
    logShachoBotActivity_('新着なし');
    return;
  }

  var targets = filterShachoTargets_(messages, processedIds);
  logShachoBotActivity_('取得 ' + messages.length + '件 / 応答対象 ' + targets.length + '件');

  for (var i = 0; i < targets.length; i++) {
    try {
      generateAndPostShachoResponse_(targets[i], roomId, token, ss);
    } catch (e) {
      logShachoBotError_('応答生成エラー msg_id=' + targets[i].message_id + ': ' + e.message);
    }
  }
}

/** 応答対象メッセージをフィルタ */
function filterShachoTargets_(messages, processedIds) {
  return messages.filter(function (msg) {
    // 自分自身のメッセージはスキップ
    if (String(msg.account.account_id) === BOT_ACCOUNT_ID) return false;
    // 処理済みスキップ
    if (processedIds[String(msg.message_id)]) return false;

    var body = msg.body || '';

    // BOTラベル付きスキップ
    if (body.indexOf(BOT_LABEL) === 0) return false;

    // AI政 (10144671) へのTOがない場合スキップ
    if (body.indexOf('[To:' + BOT_ACCOUNT_ID + ']') === -1) return false;

    // 無視パターン
    for (var i = 0; i < IGNORE_PATTERNS_BOT.length; i++) {
      if (IGNORE_PATTERNS_BOT[i].test(body)) return false;
    }

    return true;
  });
}

/** 応答生成 + Chatwork 投稿 */
function generateAndPostShachoResponse_(msg, roomId, token, ss) {
  // 質問本文の抽出 (TO 部分を除去)
  var questionBody = msg.body.replace(/\[To:\d+\]\s*[^\n]*/g, '').trim();

  if (!questionBody || questionBody.length < 5) {
    logShachoBotActivity_('質問短すぎスキップ msg_id=' + msg.message_id);
    logShachoProcessed_(ss, msg, roomId, '(短い質問スキップ)');
    return;
  }

  // 関連 Knowledge MD を取得
  var knowledge = getRelevantKnowledge_(questionBody);
  logShachoBotActivity_('タグ判定: ' + knowledge.tags.join(', '));

  // Claude API 呼び出し
  var response = callShachoClaudeAPI_(msg, questionBody, knowledge);

  if (!response || response.length < 20) {
    logShachoBotActivity_('応答生成失敗 msg_id=' + msg.message_id);
    logShachoProcessed_(ss, msg, roomId, '(応答失敗: ' + (response || 'null') + ')');
    return;
  }

  // Chatwork 投稿
  var toPrefix = '[To:' + msg.account.account_id + '] ' + msg.account.name + 'さん\n';
  var fullBody = BOT_LABEL + ' ' + toPrefix + '\n' + response;

  postChatworkMessage_(roomId, fullBody, token);

  // 処理済み記録
  logShachoProcessed_(ss, msg, roomId, response);

  logShachoBotActivity_('応答投稿完了: to=' + msg.account.name);
}

/** Claude API 呼び出し */
function callShachoClaudeAPI_(msg, questionBody, knowledge) {
  var apiKey = getClaudeApiKey_();
  if (!apiKey) return null;

  var systemPrompt = buildShachoSystemPrompt_();
  var knowledgeBlock = '## 関連Knowledge MD（質問から判定: ' + knowledge.tags.join(', ') + '）\n' + knowledge.content;

  var userContent = ''
    + '## 質問者\n' + (msg.account.name || '不明') + '\n\n'
    + '## 質問本文\n' + questionBody + '\n\n'
    + knowledgeBlock + '\n\n'
    + '## 回答指示\n'
    + 'この質問に対し、上記 Knowledge MD の中から最も該当する原則を1個選んで引用し、 6セクション完結の最新フォーマット (👑プッシュFF（嬴政AI） or 👑嬴政FF（嬴政AI） から始める) で回答してください。\n'
    + '- 商談・営業の質問なら 👑プッシュFF（嬴政AI）\n'
    + '- それ以外（ライティング・マネジメント等）なら 👑嬴政FF（嬴政AI）\n'
    + '- 「例）」 ラベルを必ず書いて独立セクションにする\n'
    + '- スクリプトは [info][/info] で囲み、「相手：〜」テンプレで会話パターン全体を完結\n'
    + '- アクションプランは「〜しろ」命令調、 句点「。」不要、 番号「1.」付けない\n'
    + '- Knowledge MD「動画内の章」 リストから source_quote と最も近い章のURL?t=秒数 を選び、 章範囲を URL 下に併記';

  var payload = {
    model: SHACHO_FF_MODEL,
    max_tokens: SHACHO_FF_MAX_TOKENS,
    system: [
      { type: 'text', text: systemPrompt, cache_control: { type: 'ephemeral' } }
    ],
    messages: [
      { role: 'user', content: userContent }
    ]
  };

  var options = {
    method: 'post',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    var code = response.getResponseCode();
    if (code !== 200) {
      logShachoBotError_('Claude API エラー: ' + code + ' ' + response.getContentText().substring(0, 300));
      return null;
    }
    var result = JSON.parse(response.getContentText());
    return result.content[0].text;
  } catch (e) {
    logShachoBotError_('Claude API 例外: ' + e.message);
    return null;
  }
}

// ============================================
// トリガー管理
// ============================================

/** 社長FFbot 起動 */
function setupShachoFFBotTrigger() {
  // 既存の pollShachoFFBot トリガー削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollShachoFFBot') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 新規 5分トリガー
  ScriptApp.newTrigger('pollShachoFFBot')
    .timeBased()
    .everyMinutes(5)
    .create();

  logShachoBotActivity_('社長FFbot 起動: 5分間隔ポーリング');
  Logger.log('Shacho FFbot trigger set: every 5 min');
}

/** 社長FFbot 停止 */
function stopShachoFFBot() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollShachoFFBot') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  logShachoBotActivity_('社長FFbot 停止: ' + removed + '個のトリガー削除');
  Logger.log('Shacho FFbot stopped: ' + removed + ' triggers removed');
}

/** 動作確認用: 1回だけ手動実行 */
function testShachoFFBot() {
  Logger.log('=== 社長FFbot テスト実行 ===');
  pollShachoFFBot();
  Logger.log('=== テスト完了、 ログシートを確認してください ===');
}

// ============================================
// ログ管理
// ============================================

function getShachoProcessedIds_(ss) {
  var sheet = ss.getSheetByName(SHEET_SHACHO_PROCESSED);
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

function logShachoProcessed_(ss, msg, roomId, response) {
  var sheet = ss.getSheetByName(SHEET_SHACHO_PROCESSED);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SHACHO_PROCESSED);
    sheet.getRange(1, 1, 1, 6).setValues([['message_id', 'date', 'roomId', 'account_id', 'name', 'body+response']]);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    String(msg.message_id),
    new Date(),
    roomId,
    String(msg.account.account_id),
    msg.account.name,
    (msg.body || '').substring(0, 300) + ' === ' + (response || '').substring(0, 300)
  ]);
}

function logShachoBotError_(message) {
  writeShachoBotLog_('ERROR', message);
}

function logShachoBotActivity_(message) {
  writeShachoBotLog_('INFO', message);
}

function writeShachoBotLog_(level, message) {
  try {
    var ss = getShachoSpreadsheet_();
    var sheet = ss.getSheetByName(SHEET_SHACHO_BOT_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_SHACHO_BOT_LOG);
      sheet.getRange(1, 1, 1, 3).setValues([['timestamp', 'level', 'message']]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), level, message]);
    var lastRow = sheet.getLastRow();
    if (lastRow > 1100) sheet.deleteRows(2, 100);
  } catch (e) {
    Logger.log('社長FFbotログ書込みエラー: ' + e.message);
  }
}
