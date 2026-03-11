// ============================================
// FeedbackKB.js — フィードバックKB管理 & RAG検索
// ============================================

/**
 * DriveにアップロードしたCSVスプレッドシートからKBにインポート
 * GAS Editorから実行: importFromDrive()
 */
function importFromDrive() {
  var SOURCE_FILE_ID = '13bam45AKzCs1ThmAXvogfqxmcKbGqz91uggya-ZvQhQ';
  var ss = getSpreadsheet_();

  // KBシート準備
  var kb = ss.getSheetByName(SHEET_KB);
  if (!kb) kb = createKBSheet_(ss);

  // ソースを開く
  var src = SpreadsheetApp.openById(SOURCE_FILE_ID);
  var srcSheet = src.getSheets()[0];
  var data = srcSheet.getDataRange().getValues();

  Logger.log('ソース行数: ' + data.length);

  // ヘッダースキップ、バッチで書込み
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var body = String(data[i][6] || '');
    if (body.length < 80) continue;

    rows.push([
      data[i][0],                              // A: 日時
      String(data[i][1] || ''),                // B: ルームID
      String(data[i][3] || ''),                // C: カテゴリ
      String(data[i][4] || ''),                // D: キーワード
      String(data[i][5] || '').substring(0, 500), // E: トリガーメッセージ
      body.substring(0, 2000),                 // F: ほしの返答
      body.length                              // G: 文字数
    ]);
  }

  // バッチ書込み（500行ずつ）
  var BATCH = 500;
  var startRow = kb.getLastRow() + 1;
  for (var b = 0; b < rows.length; b += BATCH) {
    var chunk = rows.slice(b, b + BATCH);
    kb.getRange(startRow, 1, chunk.length, 7).setValues(chunk);
    startRow += chunk.length;
    SpreadsheetApp.flush();
    Logger.log('書込み: ' + (b + chunk.length) + '/' + rows.length);
  }

  Logger.log('KBインポート完了: ' + rows.length + '件');
  return { imported: rows.length };
}

/**
 * ステージングシートからKBシートにインポート
 * 使い方: CSVをGoogle Sheetsにアップロード後に実行
 * @param {string} stagingName - ステージングシート名（デフォルト: 'KB_Staging'）
 * @returns {Object} {imported: number}
 */
function processStaging(stagingName) {
  var ss = getSpreadsheet_();
  var staging = ss.getSheetByName(stagingName || 'KB_Staging');
  if (!staging) { Logger.log('ステージングシートが見つかりません'); return { imported: 0 }; }

  var kb = ss.getSheetByName(SHEET_KB);
  if (!kb) { kb = createKBSheet_(ss); }

  var data = staging.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var body = String(data[i][6] || '');       // G列: ほしの返答
    if (body.length < 80) continue;

    // 無視パターンチェック
    var skip = false;
    for (var p = 0; p < IGNORE_PATTERNS_BOT.length; p++) {
      if (IGNORE_PATTERNS_BOT[p].test(body)) { skip = true; break; }
    }
    if (skip) continue;

    rows.push([
      data[i][0],                              // A: 日時
      String(data[i][1] || ''),                // B: ルームID
      String(data[i][3] || ''),                // C: カテゴリ
      String(data[i][4] || ''),                // D: キーワード
      String(data[i][5] || '').substring(0, 500), // E: トリガーメッセージ
      body.substring(0, 2000),                 // F: ほしの返答
      body.length                              // G: 文字数
    ]);
  }

  if (rows.length > 0) {
    var startRow = kb.getLastRow() + 1;
    kb.getRange(startRow, 1, rows.length, 7).setValues(rows);
  }

  Logger.log('KB インポート完了: ' + rows.length + '件');
  return { imported: rows.length };
}

/**
 * KBシートを新規作成（ヘッダー付き）
 */
function createKBSheet_(ss) {
  var sheet = ss.insertSheet(SHEET_KB);
  sheet.getRange(1, 1, 1, 7).setValues([[
    '日時', 'ルームID', 'カテゴリ', 'キーワード',
    'トリガーメッセージ', 'ほしの返答', '文字数'
  ]]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  return sheet;
}

/**
 * KBからキーワード+カテゴリでスコアリング検索（簡易RAG）
 * @param {string} messageBody - 入力メッセージ
 * @param {string} feedbackType - 検出されたタイプ
 * @param {number} maxResults - 最大取得件数
 * @returns {Array<{trigger: string, response: string, category: string}>}
 */
function searchKnowledgeBase_(messageBody, feedbackType, maxResults) {
  var ss = getSpreadsheet_();
  var kb = ss.getSheetByName(SHEET_KB);
  if (!kb || kb.getLastRow() < 2) return [];

  maxResults = maxResults || 5;

  // 入力メッセージからキーワード抽出
  var queryKw = extractKeywords_(messageBody);
  var queryCat = feedbackType ? mapTypeToCat_(feedbackType) : '';

  // KB読み込み（C-F列のみ: カテゴリ, キーワード, トリガー, 返答）
  var lastRow = kb.getLastRow();
  var data = kb.getRange(2, 3, lastRow - 1, 4).getValues();

  var scored = [];
  for (var i = 0; i < data.length; i++) {
    var cat = String(data[i][0]);
    var kw  = String(data[i][1]);
    var trigger = String(data[i][2]);
    var response = String(data[i][3]);

    if (!response || response.length < 50) continue;

    var score = 0;
    // カテゴリ一致で+3
    if (queryCat && cat === queryCat) score += 3;

    // キーワード一致
    for (var j = 0; j < queryKw.length; j++) {
      if (kw.indexOf(queryKw[j]) !== -1) score += 1;
      if (trigger.indexOf(queryKw[j]) !== -1) score += 0.5;
    }

    if (score > 0) {
      scored.push({ score: score, trigger: trigger, response: response, category: cat });
    }
  }

  scored.sort(function(a, b) { return b.score - a.score; });
  return scored.slice(0, maxResults);
}

/**
 * メッセージからキーワード抽出
 * @param {string} body
 * @returns {Array<string>}
 */
function extractKeywords_(body) {
  var patterns = [
    [/商談/, '商談'], [/成約/, '成約'], [/改善/, '改善'],
    [/提案/, '提案'], [/依頼/, '依頼'], [/クロージング/, 'クロージング'],
    [/リスケ/, 'リスケ'], [/返金/, '返金'], [/LINE/, 'LINE'],
    [/エルメ/, 'エルメ'], [/TikTok/, 'TikTok'], [/データ/, 'データ'],
    [/報告/, '報告'], [/質問/, '質問'], [/フォロー/, 'フォロー'],
    [/アポ/, 'アポ'], [/見送り/, '見送り'], [/保留/, '保留'],
    [/戦略/, '戦略'], [/着金/, '着金']
  ];

  var keywords = [];
  for (var i = 0; i < patterns.length; i++) {
    if (patterns[i][0].test(body)) keywords.push(patterns[i][1]);
  }
  return keywords;
}

/**
 * フォローアップが必要か判定
 */
function needsFollowUp_(feedback, feedbackType) {
  return /フォロー|報告.*まで|確認.*まで|〆切|締切|日まで/.test(feedback);
}

/**
 * フォローアップ登録
 */
function registerFollowUp_(msg, roomId, feedback) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_FOLLOWUPS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FOLLOWUPS);
    sheet.getRange(1, 1, 1, 6).setValues([[
      '作成日', 'ルームID', '対象者ID', 'フォロー日', 'トピック', 'ステータス'
    ]]);
    sheet.setFrozenRows(1);
  }

  // フォロー日を抽出（デフォルト3日後）
  var followUpDate = new Date();
  followUpDate.setDate(followUpDate.getDate() + 3);

  var dateMatch = feedback.match(/(\d{1,2})\/(\d{1,2})まで|(\d{1,2})日まで/);
  if (dateMatch) {
    var now = new Date();
    if (dateMatch[1] && dateMatch[2]) {
      followUpDate = new Date(now.getFullYear(), parseInt(dateMatch[1]) - 1, parseInt(dateMatch[2]));
    } else if (dateMatch[3]) {
      followUpDate = new Date(now.getFullYear(), now.getMonth(), parseInt(dateMatch[3]));
    }
  }

  sheet.appendRow([
    new Date(),
    roomId,
    String(msg.account.account_id),
    followUpDate,
    feedback.substring(0, 200),
    '未送信'
  ]);
}

/**
 * Botシートを1つずつ作成（タイムアウト対策）
 * step: 1=KB, 2=処理済み, 3=ログ, 4=テンプレ, 5=フォロー
 */
function setupBotSheets(step) {
  step = step || 1;
  var ss = getSpreadsheet_();

  if (step === 1) {
    if (!ss.getSheetByName(SHEET_KB)) createKBSheet_(ss);
    Logger.log('Step1: KB作成完了');
  } else if (step === 2) {
    if (!ss.getSheetByName(SHEET_PROCESSED)) {
      var s1 = ss.insertSheet(SHEET_PROCESSED);
      s1.getRange(1, 1, 1, 8).setValues([[
        'message_id', '処理日時', 'ルームID', '送信者ID',
        '送信者名', '元メッセージ', '生成FB', 'タイプ'
      ]]);
      s1.setFrozenRows(1);
    }
    Logger.log('Step2: 処理済みメッセージ作成完了');
  } else if (step === 3) {
    if (!ss.getSheetByName(SHEET_BOT_LOG)) {
      var s2 = ss.insertSheet(SHEET_BOT_LOG);
      s2.getRange(1, 1, 1, 3).setValues([['日時', 'レベル', 'メッセージ']]);
      s2.setFrozenRows(1);
    }
    Logger.log('Step3: Bot動作ログ作成完了');
  } else if (step === 4) {
    if (!ss.getSheetByName(SHEET_TEMPLATES)) {
      var s3 = ss.insertSheet(SHEET_TEMPLATES);
      s3.getRange(1, 1, 1, 5).setValues([[
        'テンプレートID', 'カテゴリ', 'トリガーパターン', 'テンプレート本文', '使用回数'
      ]]);
      s3.setFrozenRows(1);
    }
    Logger.log('Step4: FBテンプレート作成完了');
  } else if (step === 5) {
    if (!ss.getSheetByName(SHEET_FOLLOWUPS)) {
      var s4 = ss.insertSheet(SHEET_FOLLOWUPS);
      s4.getRange(1, 1, 1, 6).setValues([[
        '作成日', 'ルームID', '対象者ID', 'フォロー日', 'トピック', 'ステータス'
      ]]);
      s4.setFrozenRows(1);
    }
    Logger.log('Step5: フォローアップ作成完了');
  }

  return { step: step, done: step >= 5 };
}

/** 全Botシートを順番に作成（ヘルパー） */
function setupBotSheet1() { return setupBotSheets(1); }
function setupBotSheet2() { return setupBotSheets(2); }
function setupBotSheet3() { return setupBotSheets(3); }
function setupBotSheet4() { return setupBotSheets(4); }
function setupBotSheet5() { return setupBotSheets(5); }
