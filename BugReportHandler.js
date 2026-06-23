// ============================================
// BugReportHandler.js — 講師ルームのバグ報告を Claude で解析→GitHub PR 自動作成
//                     + 既知パターン即対応 + 受講生データ自動診断
// ============================================

var BUG_REPORT_ROOM_ID = '434019583';
var BUG_REPORT_KEYWORDS = /バグ|エラー|動かない|表示されない|おかしい|不具合|変な|うまくいかない|出ない|できない|消えた|見れない|落ちる|反応しない|公開設定/;
var BUG_REPORT_TARGET_FILES = ['post-app.html', 'コード.js', 'PostApp.js'];
var BUG_REPORT_MIN_LENGTH = 30;

/**
 * 既知パターン辞書: Claude API使わずに即返信できるケース
 * - testRegex: メッセージにマッチするか判定する正規表現
 * - generateReply: 引数 (msg, memberData) → 返信文字列
 *   memberData は extractStudentIdFromMessage_(msg.body) で取れたIDの inspectMemberDataAsObject_ 結果（取れなかったらnull）
 */
var KNOWN_PATTERNS = [
  // ===== ① リスト数/商談数 公開設定の確認系 =====
  {
    name: 'list_push_visibility_check',
    testRegex: /(リスト数|商談数|アポ数|公開設定).{0,30}(見られる|見れる|表示|完了|お願い|反映)/,
    generateReply: function(msg, memberData) {
      if (!memberData || !memberData.postSheet.exists) {
        return '📋 リスト数/商談数 公開設定確認\n\n' +
          '受講生IDが特定できなかったか、投稿数シートに該当IDが見つかりませんでした。\n' +
          'メッセージに【受講生ID:xxxx】の形式で記載されているか確認してください。';
      }
      var info = memberData.postSheet.name + '（ID: ' + memberData.id + '、row=' + memberData.postSheet.row + '）';
      if (memberData.allowed) {
        return '📋 リスト数/商談数 公開設定確認\n\n' +
          info + '\n' +
          '✅ 公開設定はONになっています（D列背景色: 水色 #00ffff）\n\n' +
          'もし受講生側でまだリスト数/商談数タブが見えない場合は以下を案内してください:\n' +
          '・ブラウザを完全に閉じてから再度開く\n' +
          '・iPhone Safari: タブ一覧から該当タブを左スワイプで閉じ、再オープン\n' +
          '・PC: Cmd+Shift+R（Mac）or Ctrl+Shift+R（Win）で強制リロード\n\n' +
          '今月のホープ数(リスト数)合計: ' + memberData.hopeSheet.currentMonthTotal + '\n' +
          '今月のプッシュ数(商談数)合計: ' + memberData.pushSheet.currentMonthTotal;
      } else {
        return '📋 リスト数/商談数 公開設定確認\n\n' +
          info + '\n' +
          '❌ 公開設定OFF（D列背景色: ' + memberData.postSheet.bgColor + '）\n\n' +
          'D列の背景色を「水色 #00ffff」に変更すると公開されます。';
      }
    }
  },

  // ===== ② パスワードリセット「マスタデータが見つかりません」 =====
  {
    name: 'password_reset_master_not_found',
    testRegex: /パスワード.{0,10}リセット|マスタデータが見つかりません/,
    generateReply: function(msg, memberData) {
      if (!memberData || !memberData.postSheet.exists) {
        return '🔑 パスワードリセット診断\n\n' +
          '受講生IDが特定できなかったか、投稿数シートに該当IDが見つかりませんでした。\n' +
          'シートにIDを追加するか、メッセージに【受講生ID:xxxx】の形式でID記載してください。';
      }
      return '🔑 パスワードリセット診断: ' + memberData.postSheet.name + '（ID: ' + memberData.id + '）\n\n' +
        '・投稿数シート: ✅ 存在（row=' + memberData.postSheet.row + '、名前=' + (memberData.postSheet.name || '空欄') + '）\n' +
        '・postapp_auth: ' + (memberData.hasPassword ? '✅ パスワード設定済み' : '❌ パスワード未設定') + '\n\n' +
        (memberData.hasPassword
          ? '→ パスワード再設定の流れが正常に動いていない可能性。post-app側で「メールアドレス」入力を求めている画面で「マスタデータが見つかりません」が出る場合、メールアドレス検証ロジックの調査が必要です。'
          : '→ 受講生は初回ログイン状態。「新規パスワード登録」リンクから登録してもらってください。');
    }
  },

  // ===== ③ 「ログインできない」「通信エラー」(現状はAPIヘルスチェック相当) =====
  {
    name: 'communication_error_login',
    testRegex: /通信エラー|ログインできな|ログイン.{0,5}できない|ログインしようと|エラーが.{0,3}出/,
    generateReply: function(msg, memberData) {
      var memberInfo = '';
      if (memberData && memberData.postSheet.exists) {
        memberInfo = '受講生情報: ' + memberData.postSheet.name + '（ID: ' + memberData.id + '）\n' +
          '・投稿数シート: ✅ 登録あり\n' +
          '・postapp_auth: ' + (memberData.hasPassword ? '✅ パスワード設定済み' : '❌ パスワード未設定（新規登録案内が必要）') + '\n\n';
      }
      return '📡 通信エラー診断\n\n' +
        memberInfo +
        '対応案内（受講生に伝えてください）:\n' +
        '1. ブラウザを完全に閉じて再度 https://giver.work/post-app/ を開く\n' +
        '2. iPhone Safari: タブを左スワイプで閉じて再オープン\n' +
        '3. LINEアプリ内ブラウザの場合は外部ブラウザで開く: https://giver.work/post-app/?openExternalBrowser=1\n' +
        '4. PC: Cmd+Shift+R（Mac） / Ctrl+Shift+R（Win）で強制リロード\n\n' +
        '⚠️ サーバー側API（postCheckId等）は別途ヘルスチェックで確認してください。';
    }
  },

  // ===== ④ ID見つからない =====
  {
    name: 'id_not_found',
    testRegex: /IDが.{0,3}見つかりません|IDを.{0,3}入力.{0,3}しても次に進めない/,
    generateReply: function(msg, memberData) {
      if (!memberData) {
        return '🆔 ID未登録診断\n\n' +
          '受講生IDが特定できませんでした。メッセージにID記載があれば再投稿してください。';
      }
      if (memberData.postSheet.exists) {
        return '🆔 ID登録状況: ' + memberData.postSheet.name + '（ID: ' + memberData.id + '）\n\n' +
          '✅ 投稿数シートには登録あり（row=' + memberData.postSheet.row + '）\n' +
          'postapp_auth: ' + (memberData.hasPassword ? '✅ パスワード設定済み' : '❌ パスワード未設定') + '\n\n' +
          '→ 受講生が「IDが見つかりません」エラーになるのは異常です。具体的に入力したID値を確認してください（先頭ゼロの有無、半角全角等）。';
      } else {
        return '🆔 ID未登録: 受講生ID ' + memberData.id + '\n\n' +
          '❌ 投稿数シートに該当IDが見つかりません。\n' +
          '→ 投稿数シートに該当受講生を追加する運用が必要です。';
      }
    }
  }
];

/**
 * メインポーリング関数（5分トリガーで呼ばれる）
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

/**
 * 1件のバグ報告を処理。
 * フロー:
 *   1. メッセージから受講生ID抽出 → inspectMemberDataAsObject_ で診断データ取得
 *   2. 既知パターン辞書とマッチング → ヒットしたら即返信
 *   3. ヒットしなければClaude API へ。コード修正が必要なら従来通りPR作成、
 *      設定/操作系なら Chatwork に診断返信
 */
function handleBugReport_(msg, roomId, token, ss) {
  try {
    Logger.log('=== バグ報告処理: msg_id=' + msg.message_id + ' ===');

    // ① 受講生IDを抽出して現状取得
    var memberData = null;
    var studentId = extractStudentIdFromMessage_(msg.body);
    if (studentId) {
      try {
        memberData = inspectMemberDataAsObject_(studentId);
        Logger.log('受講生ID=' + studentId + ' の診断データ取得: postSheet=' + memberData.postSheet.exists);
      } catch (e) {
        Logger.log('inspectMemberDataAsObject_ エラー: ' + e.message);
      }
    }

    // ② 既知パターンマッチング
    for (var p = 0; p < KNOWN_PATTERNS.length; p++) {
      var pattern = KNOWN_PATTERNS[p];
      if (pattern.testRegex.test(msg.body)) {
        Logger.log('既知パターンヒット: ' + pattern.name);
        var reply = pattern.generateReply(msg, memberData);
        postChatworkMessage_(roomId,
          '[BOT][rp aid=' + msg.account.account_id + ' to=' + roomId + '-' + msg.message_id + ']\n' +
          reply + '\n\n' +
          '（自動診断 - pattern: ' + pattern.name + '）',
          token);
        logProcessedMessage_(ss, msg, roomId, pattern.name, 'BUG_REPORT_PATTERN');
        return;
      }
    }

    // ③ Claude API で原因解析。
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

    var analysis = analyzeBugForFix_(msg.body, fileContents, memberData);

    // 解析失敗
    if (!analysis || (!analysis.category && !analysis.targetFile)) {
      var reason = (analysis && analysis.reason) || '修正対象を特定できませんでした';
      postChatworkMessage_(roomId,
        '[BOT][rp aid=' + msg.account.account_id + ' to=' + roomId + '-' + msg.message_id + ']\n' +
        '🤖 自動診断: ' + reason + '\n\n' +
        '具体的な機能名・操作手順・受講生IDを教えてもらえると修正できます。',
        token);
      logProcessedMessage_(ss, msg, roomId, 'fail: ' + reason, 'BUG_REPORT_FAIL');
      return;
    }

    // カテゴリ別分岐
    var category = analysis.category || 'code_bug';

    // コード修正以外（設定確認 / データ確認 / 認証 / 操作問題）→ 返信のみ
    if (category !== 'code_bug') {
      var replyBody = analysis.chatworkReply || analysis.summary || '解析しましたが詳細が不明です';
      postChatworkMessage_(roomId,
        '[BOT][rp aid=' + msg.account.account_id + ' to=' + roomId + '-' + msg.message_id + ']\n' +
        '🔍 自動診断結果 [' + category + ']\n\n' +
        replyBody + '\n\n' +
        '（Claude AI 解析 - コード修正不要と判定）',
        token);
      logProcessedMessage_(ss, msg, roomId, category + ': ' + (analysis.summary || ''), 'BUG_REPORT_DIAG');
      return;
    }

    // コード修正系 → 従来通りPR作成
    if (!analysis.targetFile || !analysis.newContent) {
      postChatworkMessage_(roomId,
        '[BOT][rp aid=' + msg.account.account_id + ' to=' + roomId + '-' + msg.message_id + ']\n' +
        '🤖 コード修正系と判定されましたが、修正案を生成できませんでした。\n手動対応をお願いします。',
        token);
      logProcessedMessage_(ss, msg, roomId, 'code_bug_no_fix', 'BUG_REPORT_FAIL');
      return;
    }

    var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
    var branchName = 'bugfix/cw-' + ts;
    githubCreateBranch_(branchName, 'main');
    githubUpdateFile_(analysis.targetFile, analysis.newContent, '[bot] ' + analysis.summary, branchName);

    var prBody = '## Chatwork バグ報告から自動生成 🤖\n\n' +
      '**メッセージID**: ' + msg.message_id + '\n' +
      '**送信者**: ' + (msg.account.name || ('account_id ' + msg.account.account_id)) + '\n' +
      '**ルーム**: rid=' + roomId + '\n' +
      (studentId ? ('**抽出した受講生ID**: ' + studentId + '\n') : '') +
      '\n## 報告内容\n' +
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
      '解析完了 ✅ コード修正PRを作成しました\n\n' +
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
 * Claude API でバグ報告を解析。
 * カテゴリ分類して、コード修正系ならファイル全文を生成、
 * それ以外（設定確認/データ確認/認証/操作）なら chatworkReply のみ生成。
 */
function analyzeBugForFix_(reportBody, fileContents, memberData) {
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

  // 受講生データ（取れた場合のみ）
  var memberDataSection = '';
  if (memberData) {
    memberDataSection = '\n\n## 該当受講生の現在のスプシ状態\n```json\n' +
      JSON.stringify({
        id: memberData.id,
        allowed: memberData.allowed,
        hasPassword: memberData.hasPassword,
        postSheetExists: memberData.postSheet.exists,
        postSheetBg: memberData.postSheet.bgColor,
        postSheetName: memberData.postSheet.name,
        hopeCurrentMonthTotal: memberData.hopeSheet.currentMonthTotal,
        pushCurrentMonthTotal: memberData.pushSheet.currentMonthTotal
      }, null, 2) + '\n```\n';
  }

  var systemPrompt = 'あなたは日本語で対応するソフトウェア診断アシスタント。\n' +
    'Chatworkで報告された講師からのバグ報告を読み、対象ファイル群と受講生スプシデータから、原因を特定し対応を提案する。\n\n' +
    '【まずカテゴリ分類すること】\n' +
    '・"code_bug": 明らかなコードのバグで、ファイル修正が必要\n' +
    '・"config_check": スプシD列の色など設定の確認・案内で済む\n' +
    '・"data_check": スプシデータの中身（未入力/0/不整合等）の確認・説明で済む\n' +
    '・"auth_check": パスワード/アカウント認証の問題で、コード修正不要\n' +
    '・"user_op": 受講生の操作問題（キャッシュリロード等の案内で済む）\n\n' +
    '出力は厳密にJSON形式のみ（前後の説明文・コードブロックの```禁止）:\n\n' +
    '【code_bug の場合】\n' +
    '{\n' +
    '  "category": "code_bug",\n' +
    '  "targetFile": "修正対象のファイル名（提供されたファイルのいずれか）",\n' +
    '  "newContent": "修正後のファイル全文（diffではなく完全な内容）",\n' +
    '  "summary": "修正概要を1行で",\n' +
    '  "reasoning": "原因と修正の説明（複数行可）"\n' +
    '}\n\n' +
    '【config_check / data_check / auth_check / user_op の場合】\n' +
    '{\n' +
    '  "category": "config_check" など,\n' +
    '  "summary": "1行サマリ",\n' +
    '  "reasoning": "原因の説明",\n' +
    '  "chatworkReply": "講師に返す日本語の対応案内（具体的な手順含む）"\n' +
    '}\n\n' +
    '【解析できない場合】\n' +
    '{"reason": "理由を日本語で"}\n\n' +
    '注意:\n' +
    '- newContent は対象ファイルの全文を返すこと\n' +
    '- 既存のコードスタイルを維持（var/function宣言、インデント）\n' +
    '- 確証がない場合は推測せず {"reason": "..."}を返すこと\n' +
    '- 受講生スプシデータが提供された場合は、それを根拠にした診断を優先せよ';

  var userPrompt = '## バグ報告\n```\n' + reportBody + '\n```\n' +
    memberDataSection +
    '\n## 対象ファイル群\n' + fileSection + '\n\n' +
    '上記の報告を分析し、カテゴリ判定の上で適切なJSONを出力してください。';

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
