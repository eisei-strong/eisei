// ============================================
// BugReportHandler.js — 講師ルームのバグ報告を Claude で解析→GitHub PR 自動作成
// ============================================

var BUG_REPORT_ROOM_ID = '434019583';
var BUG_REPORT_KEYWORDS = /バグ|エラー|動かない|表示されない|おかしい|不具合|変な|うまくいかない|出ない|できない|消えた|見れない|落ちる/;
var BUG_REPORT_TARGET_FILES = ['post-app.html', 'コード.js', 'PostApp.js'];
var BUG_REPORT_MIN_LENGTH = 30;

/**
 * メインポーリング関数（5分トリガーで呼ばれる）
 * 講師ルーム(rid=434019583)の新規メッセージを確認 → バグ報告判別 → PR作成
 */
function pollBugReportRoom() {
  var token = getChatworkToken_();
  if (!token) {
    logBotError_('CHATWORK_API_TOKEN 未設定');
    return;
  }
  var ss = getSpreadsheet_();
  var processedIds = getProcessedMessageIds_(ss);

  var messages = getNewMessages_(BUG_REPORT_ROOM_ID, token);
  if (!messages || messages.length === 0) {
    logBotActivity_('バグ報告ルーム: 新着なし');
    return;
  }

  var processed = 0;
  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];
    if (processedIds[msg.message_id]) continue;
    if (msg.account.account_id == BOT_ACCOUNT_ID) continue;

    // 無視パターン
    var skip = false;
    for (var j = 0; j < IGNORE_PATTERNS_BOT.length; j++) {
      if (IGNORE_PATTERNS_BOT[j].test(msg.body)) { skip = true; break; }
    }
    if (skip) continue;

    if (msg.body.length < BUG_REPORT_MIN_LENGTH) continue;
    if (!BUG_REPORT_KEYWORDS.test(msg.body)) continue;

    handleBugReport_(msg, BUG_REPORT_ROOM_ID, token, ss);
    processed++;
  }
  logBotActivity_('バグ報告ルーム: ' + messages.length + '件取得, ' + processed + '件処理');
}

function handleBugReport_(msg, roomId, token, ss) {
  try {
    Logger.log('=== バグ報告処理: msg_id=' + msg.message_id + ' ===');

    // 関連ファイル取得
    var fileContents = {};
    for (var i = 0; i < BUG_REPORT_TARGET_FILES.length; i++) {
      var filename = BUG_REPORT_TARGET_FILES[i];
      try {
        var f = githubGetFile_(filename);
        fileContents[filename] = f.content;
      } catch (e) {
        Logger.log('ファイル取得失敗: ' + filename + ' - ' + e.message);
      }
    }

    if (Object.keys(fileContents).length === 0) {
      throw new Error('GitHub から対象ファイルを取得できませんでした');
    }

    // Claude API で原因解析+修正案生成
    var analysis = analyzeBugForFix_(msg.body, fileContents);
    if (!analysis || !analysis.targetFile || !analysis.newContent) {
      var reason = (analysis && analysis.reason) || '修正対象を特定できませんでした';
      postChatworkMessage_(roomId,
        '[BOT][rp aid=' + msg.account.account_id + ' to=' + roomId + '-' + msg.message_id + ']\n' +
        '解析できませんでした:\n' + reason + '\n\n具体的な機能名・操作手順を教えてもらえると修正できます。',
        token);
      logProcessedMessage_(ss, msg, roomId, 'fail: ' + reason, 'BUG_REPORT_FAIL');
      return;
    }

    // GitHub: ブランチ作成 → ファイル更新 → PR作成
    var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
    var branchName = 'bugfix/cw-' + ts;
    githubCreateBranch_(branchName, 'main');
    githubUpdateFile_(analysis.targetFile, analysis.newContent, '[bot] ' + analysis.summary, branchName);

    var prBody = '## Chatwork バグ報告から自動生成 🤖\n\n' +
      '**メッセージID**: ' + msg.message_id + '\n' +
      '**送信者**: ' + (msg.account.name || ('account_id ' + msg.account.account_id)) + '\n' +
      '**ルーム**: rid=' + roomId + '\n\n' +
      '## 報告内容\n' +
      '```\n' + msg.body + '\n```\n\n' +
      '## 解析結果\n' +
      (analysis.reasoning || '-') + '\n\n' +
      '## 修正概要\n' +
      analysis.summary + '\n\n' +
      '## 対象ファイル\n' +
      '`' + analysis.targetFile + '`\n\n' +
      '⚠️ **Bot自動生成のため、必ずレビュー後にマージしてください。**';

    var pr = githubCreatePR_(
      'fix: ' + analysis.summary + ' (Chatworkバグ報告)',
      prBody,
      branchName,
      'main'
    );

    postChatworkMessage_(roomId,
      '[BOT][rp aid=' + msg.account.account_id + ' to=' + roomId + '-' + msg.message_id + ']\n' +
      '解析完了 ✅\n\n' +
      '**原因**: ' + analysis.summary + '\n' +
      '**PR**: ' + pr.html_url + '\n\n' +
      '⚠️ Bot自動生成。レビュー後にマージしてください。',
      token
    );

    logProcessedMessage_(ss, msg, roomId, pr.html_url, 'BUG_REPORT_PR');
    Logger.log('✅ PR作成完了: ' + pr.html_url);

  } catch (e) {
    logBotError_('handleBugReport: ' + e.message);
    try {
      postChatworkMessage_(roomId,
        '[BOT][rp aid=' + msg.account.account_id + ' to=' + roomId + '-' + msg.message_id + ']\n' +
        '⚠️ 処理中にエラーが発生しました: ' + e.message + '\n手動対応をお願いします。',
        token
      );
    } catch (e2) {}
  }
}

/**
 * Claude API でバグ報告を解析し、修正後のファイル全文を生成
 */
function analyzeBugForFix_(reportBody, fileContents) {
  var apiKey = getClaudeApiKey_();
  if (!apiKey) {
    logBotError_('CLAUDE_API_KEY 未設定');
    return null;
  }

  // ファイル内容を結合（各ファイル最大5万文字）
  var fileSection = '';
  var keys = Object.keys(fileContents);
  for (var k = 0; k < keys.length; k++) {
    var filename = keys[k];
    var content = fileContents[filename];
    if (content.length > 50000) {
      content = content.substring(0, 50000) + '\n... (truncated, ' + (fileContents[filename].length - 50000) + '文字省略)';
    }
    fileSection += '\n\n=== ' + filename + ' ===\n' + content + '\n';
  }

  var systemPrompt = 'あなたは日本語で対応するソフトウェア修正アシスタント。\n' +
    'Chatworkで報告された講師からのバグ報告を読み、対象ファイルの中身を確認して修正案を提案する。\n\n' +
    '出力は厳密にJSON形式のみ（前後の説明文・コードブロックの```禁止）:\n' +
    '{\n' +
    '  "targetFile": "修正対象のファイル名（提供されたファイルのいずれか）",\n' +
    '  "newContent": "修正後のファイル全文（diffではなく完全な内容）",\n' +
    '  "summary": "修正概要を1行で",\n' +
    '  "reasoning": "原因と修正の説明（複数行可）"\n' +
    '}\n\n' +
    '解析できない場合 / 修正不要な場合:\n' +
    '{"reason": "理由を日本語で"}\n\n' +
    '注意:\n' +
    '- newContent は対象ファイルの全文を返すこと\n' +
    '- 既存のコードスタイルを維持すること（var/function宣言、インデント）\n' +
    '- 確証がない場合は推測せず {"reason": "..."}を返すこと';

  var userPrompt = '## バグ報告\n```\n' + reportBody + '\n```\n\n' +
    '## 対象ファイル群\n' + fileSection + '\n\n' +
    '上記の報告から原因を特定し、必要なら修正後のファイル全文を生成してください。';

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 16000,
      system: systemPrompt,
      messages: [{ role: 'user', content: userPrompt }]
    }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    logBotError_('Claude API failed: ' + res.getContentText().substring(0, 500));
    return null;
  }

  var body = JSON.parse(res.getContentText());
  var text = body.content[0].text;

  // JSON抽出（前後の説明文があっても拾う）
  var match = text.match(/\{[\s\S]*\}/);
  if (!match) {
    logBotError_('Claude応答にJSONなし: ' + text.substring(0, 200));
    return null;
  }

  try {
    return JSON.parse(match[0]);
  } catch (e) {
    logBotError_('JSON parse 失敗: ' + e.message + ' / 応答: ' + text.substring(0, 500));
    return null;
  }
}

/**
 * バグ報告監視トリガーを設定（5分ごと）
 * GASエディタから1回実行
 */
function setupBugReportTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var deleted = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollBugReportRoom') {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted++;
    }
  }
  ScriptApp.newTrigger('pollBugReportRoom')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('✅ バグ報告トリガー設定（5分ごと、旧トリガー' + deleted + '個削除）');
}

/** バグ報告監視トリガー停止 */
function stopBugReportTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var deleted = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollBugReportRoom') {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted++;
    }
  }
  Logger.log('バグ報告トリガー停止: ' + deleted + '個削除');
}
