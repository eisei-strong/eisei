// ============================================
// RuleCheckBot.js — チャットルール違反チェックぼっと
// ============================================

// --- 監視対象ルーム ---
var RULE_CHECK_ROOMS = [
  { roomId: '349937583', name: 'ウォーリアーズ' },
  { roomId: '400352653', name: '受講生勝たせるチャット' },
  { roomId: '402148031', name: 'ガーディアンズ' },
  { roomId: '415414086', name: '緊急｜営業', urgent: true },
  { roomId: '412572337', name: '緊急｜講師', urgent: true }
];

var SHEET_RULE_PROCESSED = 'ルールチェック済み';
var SHEET_RULE_LOG       = 'ルールチェックログ';
var RULE_BOT_LABEL       = '[ルールチェック]';
var RULE_CHECK_MIN_LENGTH = 5;  // これ以下のメッセージはスキップ
var RULE_LOG_ROOM_ID      = '434014865';  // 指摘記録を残すルーム
var SHEET_RULE_PENDING    = 'ルール修正待ち';
var RULE_REMIND_INTERVAL_HOURS = 1;  // リマインド間隔（時間）

// ============================================
// メインポーリング
// ============================================

/**
 * ルールチェックポーリング（1分トリガーで呼ばれる）
 */
function pollAndCheckRules() {
  var token = getChatworkToken_();
  if (!token) {
    logRuleCheck_('ERROR', 'CHATWORK_API_TOKEN未設定');
    return;
  }

  var ss = getSpreadsheet_();
  var processedIds = getRuleProcessedIds_(ss);

  for (var r = 0; r < RULE_CHECK_ROOMS.length; r++) {
    var room = RULE_CHECK_ROOMS[r];
    var messages = getRuleMessages_(room.roomId, token);

    if (!messages || messages.length === 0) continue;

    // 完了報告を検知
    checkRuleCompletions_(messages, room.roomId, ss);

    var targets = filterRuleTargets_(messages, processedIds);
    logRuleCheck_('INFO', 'Room ' + room.name + '(' + room.roomId + '): ' + messages.length + '件取得, ' + targets.length + '件チェック対象');

    for (var i = 0; i < targets.length; i++) {
      checkAndPost_(targets[i], room.roomId, token, ss);
    }
  }
}

// ============================================
// Chatwork API
// ============================================

/**
 * メッセージ取得（force=0で新着のみ）
 */
function getRuleMessages_(roomId, token) {
  var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages?force=0';
  var options = {
    method: 'get',
    headers: { 'X-ChatWorkToken': token },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    if (code === 204) return [];
    if (code !== 200) {
      logRuleCheck_('ERROR', 'GET失敗: room=' + roomId + ' code=' + code);
      return [];
    }

    var remaining = parseInt(response.getHeaders()['x-ratelimit-remaining'] || '999');
    if (remaining < CW_RATE_STOP) {
      logRuleCheck_('ERROR', 'レートリミット危険: remaining=' + remaining);
      return [];
    }

    return JSON.parse(response.getContentText());
  } catch (e) {
    logRuleCheck_('ERROR', 'GET例外: ' + e.message);
    return [];
  }
}

/**
 * チェック対象のフィルタ
 */
function filterRuleTargets_(messages, processedIds) {
  return messages.filter(function(msg) {
    // BOT自身・AI政・嬴政ver3.0のメッセージスキップ
    if (String(msg.account.account_id) === BOT_ACCOUNT_ID) return false;
    if (msg.account.name === 'AI政') return false;
    if (msg.account.name === '嬴政ver3.0') return false;

    // 処理済みスキップ
    if (processedIds[String(msg.message_id)]) return false;

    var body = msg.body || '';

    // BOTラベル付きスキップ
    if (body.indexOf(BOT_LABEL) === 0) return false;
    if (body.indexOf('⚠️文章ルール違反') !== -1) return false;

    // 短すぎるメッセージスキップ
    var cleanBody = body.replace(/\[To:\d+\]\s*[^\n]*/g, '').replace(/\[引用[^\]]*\][\s\S]*?\[\/引用\]/g, '').trim();
    if (cleanBody.length < RULE_CHECK_MIN_LENGTH) return false;

    // システムメッセージスキップ
    for (var i = 0; i < IGNORE_PATTERNS_BOT.length; i++) {
      if (IGNORE_PATTERNS_BOT[i].test(body)) return false;
    }

    return true;
  });
}

// ============================================
// ルールチェック → Claude API → 投稿
// ============================================

/**
 * メッセージをチェックして違反があれば投稿
 */
function checkAndPost_(msg, roomId, token, ss) {
  var body = msg.body || '';

  // Claude APIでルールチェック
  var result = callClaudeForRuleCheck_(body, msg.account.name);

  // 処理済み記録（違反有無に関わらず）
  logRuleProcessed_(ss, msg, roomId, result);

  if (!result || result === 'OK') {
    logRuleCheck_('INFO', 'ルール違反なし: msg_id=' + msg.message_id + ' from=' + msg.account.name);
    return;
  }

  // 違反あり → 違反文章を引用して指摘
  var isUrgent = false;
  for (var ui = 0; ui < RULE_CHECK_ROOMS.length; ui++) {
    if (RULE_CHECK_ROOMS[ui].roomId === roomId && RULE_CHECK_ROOMS[ui].urgent) {
      isUrgent = true;
      break;
    }
  }
  var cleanBody = (msg.body || '').substring(0, 2000);
  var reply;
  if (isUrgent) {
    reply = '⚠️文章ルール違反\n\n';
  } else {
    reply = '[To:' + msg.account.account_id + '] ' + msg.account.name + 'さん\n';
    reply += '⚠️文章ルール違反\n\n';
  }
  reply += '[qt][qtmeta aid=' + msg.account.account_id + ' time=' + (msg.send_time || '') + ']' + cleanBody + '[/qt]\n\n';
  reply += result.replace(/^⚠️\s*ルール違反があるよ！\s*\n*/m, '');

  postChatworkMessage_(roomId, reply, token);

  // 指摘記録を記録ルームに投稿
  var roomName = '';
  for (var ri = 0; ri < RULE_CHECK_ROOMS.length; ri++) {
    if (RULE_CHECK_ROOMS[ri].roomId === roomId) { roomName = RULE_CHECK_ROOMS[ri].name; break; }
  }
  var logMsg = '[info][title]📝 文章ルール指摘記録[/title]';
  logMsg += '対象者: ' + msg.account.name + '\n';
  logMsg += 'ルーム: ' + roomName + '(' + roomId + ')\n';
  logMsg += '日時: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n';
  logMsg += '【元のメッセージ】\n' + (msg.body || '').substring(0, 2000) + '\n\n';
  logMsg += '【指摘内容】\n' + result.replace(/^⚠️ ルール違反があるよ！\n*/m, '');
  logMsg += '[/info]';
  postChatworkMessage_(RULE_LOG_ROOM_ID, logMsg, token);

  // 修正待ちリストに登録
  registerRulePending_(ss, msg, roomId);
  logRuleCheck_('INFO', 'ルール指摘投稿: to=' + msg.account.name + ' msg_id=' + msg.message_id);
}

/**
 * Claude APIでチャットルール違反チェック
 */
function callClaudeForRuleCheck_(messageBody, senderName) {
  var apiKey = getClaudeApiKey_();
  if (!apiKey) {
    logRuleCheck_('ERROR', 'CLAUDE_API_KEY未設定');
    return null;
  }

  var systemPrompt = buildRuleCheckSystemPrompt_();
  var userPrompt = '送信者: ' + senderName + '\n\nメッセージ:\n```\n' + messageBody.substring(0, 2000) + '\n```';

  var payload = {
    model: CLAUDE_MODEL,
    max_tokens: 512,
    system: systemPrompt,
    messages: [
      { role: 'user', content: userPrompt }
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
      logRuleCheck_('ERROR', 'Claude API エラー: ' + code + ' ' + response.getContentText().substring(0, 300));
      return null;
    }

    var result = JSON.parse(response.getContentText());
    return result.content[0].text;
  } catch (e) {
    logRuleCheck_('ERROR', 'Claude API 例外: ' + e.message);
    return null;
  }
}

/**
 * ルールチェック用システムプロンプト
 */
function buildRuleCheckSystemPrompt_() {
  return [
    'あなたは社内チャットのルールチェッカーです。',
    'メンバーが送信したメッセージが以下のチャットルールに違反していないかチェックしてください。',
    '',
    '## チャットルール（9個）',
    '',
    '1. 【主語明確】主語を必ず書く',
    '   - ダメ: 「受講生にメールします」',
    '   - 良い: 「"僕が"受講生にメールします」',
    '',
    '2. 【形式文章禁止】「お世話になっております」「お疲れ様です」等の形式的な挨拶は全て不要',
    '',
    '3. 【感情論禁止】「迷惑かけてごめんなさい」「すみません」等の感情論の文章は全て不要',
    '',
    '4. 【タメ語】敬語は使わない。タメ口で書く',
    '   - ダメ: 「〜いたします」「〜でございます」「〜させていただきます」',
    '   - 良い: 「〜する」「〜だよ」「〜するね」',
    '',
    '5. 【句点で改行すんな】。で2行改行しろ',
    '   - 「。」を使った場合、その後に必ず2行分の改行を入れる',
    '   - 読点「、」で文を繋ぎ続けて改行するのはNG',
    '   - ダメ: 「〜した。次に〜」（句点の後に改行がない）',
    '   - ダメ: 「〜したから、\\n〜して、\\n〜した」（読点で改行して文を繋ぎ続ける）',
    '   - 良い: 「〜した。」の後に空行を1つ入れてから次の文',
    '   - 「。」を使わない短文（「修正済み」等）は問題ない',
    '   - 「！」「？」で終わる文は句点「。」不要。「！」「？」も文の終わりとして扱う',
    '   - 「！」「？」の後も2行改行が必要',
    '   - 指摘時は「【句点で改行すんな】。で2行改行しろ」と書くこと',
    '',
    '',
    '7. 【冒頭ラベル】メッセージの1行目に以下のいずれかを単独で付ける（ラベルの後に内容を続けない）',
    '   - ⭕️確認',
    '   - ⭕️質問',
    '   - ⭕️依頼',
    '   - ⭕️回答',
    '   - ⭕️報告',
    '   - ⭕️指示',
    '   - ⭕️信託',
    '   - ダメ: 「⭕️報告該当者修正済み」（ラベルと内容がくっついてる）',
    '   - 良い: 1行目「⭕️報告」、改行して2行目以降に内容を書く',
    '   - 1つのメッセージにつき、メインの内容（依頼・質問等）は1つに絞れ。2つあるなら2回に分けろ',
    '',
    '8. 【指示語禁止】「それ」「あの」「これ」「あれ」などの指示語を使わない',
    '   - ダメ: 「それについて確認した」',
    '   - 良い: 「〇〇の件について確認した」',
    '',
    '9. 【引用返信】返事をする場合は必ず引用（[qt]）を使って、何に対する返事か明確にする',
    '',
    '10. 【URL説明】URLを貼る場合、その上に何のURLか説明を書け',
    '   - ダメ: いきなりURLだけ貼る',
    '   - 良い: 「入力スプシ」「マニュアル」「解説動画」等の説明を書いてからURLを貼る',
    '',
    '11. 【理由の書き方】理由を書く場合は結論のすぐ下に「理由：」を記載する',
    '   - 「理由：」は結論のすぐ下の行に書く。結論と理由の間に空行を絶対に入れるな',
    '   - 理由が1つの場合: 「理由：○○だから。」',
    '   - 理由が2つ以上の場合: 「◎理由」でもOK',
    '   - ⭕️指示 の場合は理由を必ず書くこと（なぜその指示を出すのか）',
    '   - ⭕️指示 以外では理由がないこと自体は違反ではない',
    '',
    '## 判定ルール',
    '',
    '- Chatworkの[To:][qt][qtmeta][info]等のタグ部分はルールチェックの対象外',
    '- [To:]タグの後の改行はルール違反ではない。Toの羅列で改行するのは正常',
    '- スタンプ、リアクション、ファイルアップロード通知等はチェック対象外',
    '- [qt]ブロック内のテキストはチェック対象外（他人の発言の引用なので）',
    '- URL自体の内容はチェック対象外だが、URLの前に何のリンクか説明がない場合は指摘する',
    '- 「了解」「OK」「ありがとう」等の一言だけの返事はチェック対象外',
    '- 2行以上のメッセージは必ず冒頭ラベルが必要',
    '- 理由を書いている場合は必ず「理由：」のフォーマットを使うこと',
    '',
    '## 絶対に守れ',
    '- 元のメッセージに存在しない言葉・表現を指摘するな（ハルシネーション厳禁）',
    '- 指摘する場合は必ず元のメッセージから該当箇所を正確に引用しろ',
    '- 存在しないラベルや言葉を捏造して指摘するな',
    '',
    '## 出力フォーマット',
    '',
    '違反がない場合: 「OK」とだけ出力',
    '',
    '違反がある場合: 以下のフォーマットで出力',
    '- 違反ルールごとに1行で簡潔に指摘',
    '- 修正例を具体的に示す',
    '- タメ口で、フレンドリーだけど的確に指摘する',
    '- 最大3つまでの違反を指摘（多すぎると読まない）',
    '',
    '出力例:',
    '---',
    '1. 【主語明確】「メール送ります」→ 「"俺が"メール送る」に直して',
    '',
    '2. 【冒頭ラベル】冒頭に ⭕️確認 / ⭕️質問 / ⭕️依頼 / ⭕️回答 のどれかをつけて',
    '',
    '3. 【タメ語】「送信いたします」→ 「送信する」に直して',
    '---',
    '',
    '## 重要ルール',
    '- 必ず「OK」か違反指摘のどちらかで始めること',
    '- 違反がある場合、前置き（「⚠️ ルール違反があるよ！」等）は不要。いきなり指摘項目（1. 2. 3.）から始めること',
    '- 違反がある場合、指摘の最後に必ず以下を追加すること:',
    '',
    '## 正解例',
    '指摘の後に「✅正しい文章例」として、元のメッセージをルール通りに修正した完成形を[info]タグで囲んで書くこと。',
    'メンバーがそのままコピペして再投稿できるレベルで書け。',
    'フォーマット:',
    '✅正しい文章例',
    '[info](修正した文章全文)[/info]',
    '',
    '正解例のルール:',
    '- 元のメッセージの内容は一切削らない。全文を残す',
    '- 余計な文を足さない',
    '- 理由がある場合は「理由：」を付ける',
    '- 冒頭ラベル（⭕️報告 等）は1行目に単独で書く',
    '- 句点「。」の後は2行改行',
    '',
    '例:',
    '元: 「上記なしで！10時希望で返信があったため。俺が9時プッシュ入れる！」',
    '✅正しい文章例',
    '[info]⭕️報告',
    '上記なしで！',
    '理由：10時希望で返信があったため。',
    '',
    '俺が9時プッシュ入れる！[/info]',
    '',
    '## アクションプラン',
    '正解例の後に以下を必ず追加:',
    '',
    '🔸アクションプラン',
    '1. 上の指摘を踏まえて、元のメッセージを正しく書き直してこのチャットに再投稿しろ',
    '2. 完了したら「完了！」と伝えろ'
  ].join('\n');
}

// ============================================
// トリガー管理
// ============================================

/**
 * ルールチェックBot開始
 */
function setupRuleCheckTrigger() {
  // 既存トリガー削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollAndCheckRules') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 1分間隔トリガー作成
  ScriptApp.newTrigger('pollAndCheckRules')
    .timeBased()
    .everyMinutes(1)
    .create();

  logRuleCheck_('INFO', 'ルールチェックBot開始: 1分間隔ポーリング');
  Logger.log('RuleCheckBot trigger set: every 1 min');
}

/**
 * ルールチェックBot停止
 */
function stopRuleCheckBot() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollAndCheckRules') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }

  logRuleCheck_('INFO', 'ルールチェックBot停止: ' + removed + '個のトリガー削除');
  Logger.log('RuleCheckBot stopped: ' + removed + ' triggers removed');
}

/**
 * ルールチェックBot稼働状況
 */
function getRuleCheckStatus() {
  var triggers = ScriptApp.getProjectTriggers();
  var active = [];
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollAndCheckRules') {
      active.push('pollAndCheckRules');
    }
  }

  return {
    running: active.length > 0,
    triggers: active,
    rooms: RULE_CHECK_ROOMS.length
  };
}

// ============================================
// ログ・処理済み管理
// ============================================

/**
 * 処理済みメッセージID取得（直近500件）
 */
function getRuleProcessedIds_(ss) {
  var sheet = ss.getSheetByName(SHEET_RULE_PROCESSED);
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
function logRuleProcessed_(ss, msg, roomId, result) {
  var sheet = ss.getSheetByName(SHEET_RULE_PROCESSED);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RULE_PROCESSED);
    sheet.appendRow(['message_id', '日時', 'roomId', 'account_id', '送信者', 'メッセージ', '判定結果']);
  }

  sheet.appendRow([
    String(msg.message_id),
    new Date(),
    roomId,
    String(msg.account.account_id),
    msg.account.name,
    (msg.body || '').substring(0, 2000),
    (result || 'null').substring(0, 2000)
  ]);
}

/**
 * ログ書込み
 */
function logRuleCheck_(level, message) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(SHEET_RULE_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_RULE_LOG);
      sheet.appendRow(['日時', 'レベル', 'メッセージ']);
    }

    sheet.appendRow([new Date(), level, message]);

    // 1000行超えたらトリム
    var lastRow = sheet.getLastRow();
    if (lastRow > 1100) {
      sheet.deleteRows(2, 100);
    }
  } catch (e) {
    Logger.log('ルールチェックログ書込みエラー: ' + e.message);
  }
}

// ============================================
// 修正待ち管理 & リマインド
// ============================================

/**
 * 修正待ちリストに登録
 */
function registerRulePending_(ss, msg, roomId) {
  var sheet = ss.getSheetByName(SHEET_RULE_PENDING);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RULE_PENDING);
    sheet.appendRow(['account_id', '送信者', 'roomId', '指摘日時', '最終リマインド', 'ステータス', 'message_id']);
  }

  sheet.appendRow([
    String(msg.account.account_id),
    msg.account.name,
    roomId,
    new Date(),
    new Date(),
    '未完了',
    String(msg.message_id)
  ]);
}

/**
 * 完了報告を検知してペンディングを消化
 * pollAndCheckRules内で呼ばれる
 */
function checkRuleCompletions_(messages, roomId, ss) {
  var sheet = ss.getSheetByName(SHEET_RULE_PENDING);
  if (!sheet || sheet.getLastRow() < 2) return;

  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];
    var body = msg.body || '';

    // AI政へのToで「完了」を含むメッセージを検知
    if (body.indexOf('[To:' + BOT_ACCOUNT_ID + ']') === -1) continue;
    if (!/完了/.test(body)) continue;

    // この送信者の未完了を完了にする
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(msg.account.account_id) &&
          String(data[r][2]) === roomId &&
          data[r][5] === '未完了') {
        sheet.getRange(r + 1, 6).setValue('完了');
        logRuleCheck_('INFO', '修正完了: ' + msg.account.name + ' room=' + roomId);
      }
    }
  }
}

/**
 * 未完了の指摘に対してリマインド送信
 */
function remindRulePending() {
  var token = getChatworkToken_();
  if (!token) return;

  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_RULE_PENDING);
  if (!sheet || sheet.getLastRow() < 2) return;

  var now = new Date();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][5] !== '未完了') continue;

    var lastRemind = data[i][4];
    if (!(lastRemind instanceof Date)) continue;

    var hoursSince = (now - lastRemind) / (1000 * 60 * 60);
    if (hoursSince < RULE_REMIND_INTERVAL_HOURS) continue;

    var accountId = String(data[i][0]);
    var senderName = data[i][1];
    var roomId = String(data[i][2]);

    var reminder = '[To:' + accountId + '] ' + senderName + 'さん\n';
    reminder += '⏰ 文章ルール修正がまだ完了してないよ！\n\n';
    reminder += '1. 指摘された文章を正しく書き直して再投稿しろ\n';
    reminder += '2. 完了したら「完了！」と伝えろ';

    postChatworkMessage_(roomId, reminder, token);

    // 最終リマインド日時を更新
    sheet.getRange(i + 1, 5).setValue(now);
    logRuleCheck_('INFO', 'リマインド送信: ' + senderName + ' room=' + roomId);
  }
}

/**
 * リマインドトリガー設定（1時間おき）
 */
function setupRuleRemindTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'remindRulePending') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('remindRulePending')
    .timeBased()
    .everyHours(1)
    .create();

  logRuleCheck_('INFO', 'リマインドトリガー設定: 1時間間隔');
  Logger.log('Rule remind trigger set: every 1 hour');
}

// ============================================
// テスト用
// ============================================

/**
 * テスト: 任意のメッセージでルールチェック（投稿はしない）
 */
function testRuleCheck(testMessage) {
  testMessage = testMessage || 'お世話になっております。それについて確認いたしました。明日対応します。';

  var result = callClaudeForRuleCheck_(testMessage, 'テストユーザー');
  Logger.log('=== ルールチェック結果 ===');
  Logger.log('入力: ' + testMessage);
  Logger.log('判定: ' + result);
  return result;
}
