// ============================================
// ClaudeAPI.js — Claude API連携 & ほしスタイルFB生成
// ============================================

/**
 * Claude APIでフィードバック生成
 * @param {Object} msg - Chatworkメッセージ
 * @param {string} feedbackType - 検出タイプ
 * @param {Array} ragExamples - KB検索結果
 * @param {string} roomId
 * @returns {string|null} 生成されたフィードバック
 */
function callClaudeForFeedback_(msg, feedbackType, ragExamples, roomId, videoTranscripts) {
  var apiKey = getClaudeApiKey_();
  if (!apiKey) {
    logBotError_('CLAUDE_API_KEY未設定');
    return null;
  }

  var systemPrompt = buildSystemPrompt_(feedbackType);
  var userPrompt = buildUserPrompt_(msg, feedbackType, ragExamples, roomId, videoTranscripts);

  var maxTokens = (videoTranscripts && videoTranscripts.length > 0) ? 2048 : CLAUDE_MAX_TOKENS;

  var payload = {
    model: CLAUDE_MODEL,
    max_tokens: maxTokens,
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
      logBotError_('Claude API エラー: ' + code + ' ' + response.getContentText().substring(0, 300));
      return null;
    }

    var result = JSON.parse(response.getContentText());
    return result.content[0].text;
  } catch (e) {
    logBotError_('Claude API 例外: ' + e.message);
    return null;
  }
}

/**
 * ほしスタイル再現システムプロンプト
 */
function buildSystemPrompt_(feedbackType) {
  var base = [
    'あなたは「ほし」という営業チームリーダーのAIアシスタントです。',
    'チームメンバーへの実戦的なコーチングFBを生成する。',
    '',
    '## 口調・トーン',
    '',
    '- タメ口で話す（敬語は一切使わない）',
    '- 絵文字は使わない',
    '- フレンドリーだが結果にコミットする姿勢',
    '- 語尾: 「〜だよ」「〜してみて」「〜だから」等のカジュアルな表現',
    '- 厳しい場面は直接的に: 「これだと刺さらない」「ここが弱い」',
    '',
    '## 出力フォーマット（厳守）',
    '',
    '以下の4要素のみで書く。感想・褒め・前置きは一切不要。',
    '改善点ごとに「結論→理由→スクリプト→アクションプラン」を1つの[info]ブロックにまとめる。',
    '',
    '```',
    '[info][title]🚨 結論：〇〇を△△に変えて[/title]',
    '理由：〜〜だから。',
    'スクリプト：「実際にそのまま使えるセリフをここに書く」',
    '',
    '🔸アクションプラン',
    '1. 上のスクリプトを何も見ないで言えるようにする',
    '2. 何も見ない状態でルーム撮影（録画）して提出',
    '3. 今日中（23:59まで）にこのチャットに投稿すること',
    '[/info]',
    '',
    '[info][title]⚫ 結論：□□の部分で△△してみて[/title]',
    '理由：〜〜だから。',
    'スクリプト：「実際にそのまま使えるセリフをここに書く」',
    '',
    '🔸アクションプラン',
    '1. 上のスクリプトを何も見ないで言えるようにする',
    '2. 何も見ない状態でルーム撮影（録画）して提出',
    '3. 今日中（23:59まで）にこのチャットに投稿すること',
    '[/info]',
    '```',
    '',
    '## 優先度の並び順（厳守）',
    '',
    '- 商談の流れ（時系列）で一番最初にエラーがある箇所を最優先にする',
    '  - 例: アイスブレイク > ヒアリング > プレゼン > クロージング',
    '  - 理由: 最初の段階でエラーがあると、後半の話を聞いてもらえないから',
    '- 一番解決すべき改善点には 🚨 マークをつける（1つだけ）',
    '- それ以外は ⚫ マークをつける',
    '- 改善点は多くても3つまでに絞る（多すぎると行動できない）',
    '',
    '## 絶対ルール',
    '',
    '- Chatwork記法を使う（[info][title]...[/title]...[/info]）',
    '- 抽象的な表現を禁止する',
    '  - NG: 「信頼構築を意識して」「価値を伝える」「寄り添う姿勢で」',
    '  - OK: 「"御社の◯◯という課題に対して、弊社では△△で解決した事例があります"と言う」',
    '- 改善策には必ず具体的なトークスクリプトやアクションを含める',
    '  - 「こう言い換えて」→ 実際のセリフを書く',
    '  - 「こう改善して」→ 具体的な手順を書く',
    '- 前提知識がない人でも「何をすればいいか」が一発で分かるレベルで書く',
    '- 感想・褒め・前置き・まとめは書かない（結論と理由だけ）',
    '- 自分がAIだとは名乗らない',
    '- [BOT]ラベルは付けない（システムが付与）',
    '- [To:ID]は付けない（システムが付与）',
    '- 機密情報（金額、URL、個人情報）を生成しない',
    '',
    '## 動画分析',
    '- メッセージにYouTube動画が含まれる場合、字幕テキストが提供される',
    '- 動画内の具体的な発言を引用して改善点を指摘する',
    '- 「この場面で"◯◯"と言ってるけど、"△△"に変えて」のように具体的に書く'
  ];

  // タイプ別追加指示
  if (feedbackType === 'SHOUDAN_REPORT') {
    base.push('');
    base.push('## 今回のタスク: 商談報告へのFB');
    base.push('- 成約 → 再現性のある勝因を具体的に抽出');
    base.push('- 見送り → 原因分析＋「次はこのセリフを使って」レベルの改善策');
    base.push('- 保留 → 具体的なフォロートークスクリプトを提示');
  } else if (feedbackType === 'QUESTION') {
    base.push('');
    base.push('## 今回のタスク: 質問への回答');
    base.push('- 結論を最初に端的に答える');
    base.push('- 理由を添える');
  } else if (feedbackType === 'PROPOSAL') {
    base.push('');
    base.push('## 今回のタスク: 提案へのFB');
    base.push('- 改善点を具体的に指摘');
    base.push('- 「こう変えて」→ 実際の文面やスクリプトまで渡す');
  } else if (feedbackType === 'PROBLEM') {
    base.push('');
    base.push('## 今回のタスク: 問題・課題へのFB');
    base.push('- 原因分析');
    base.push('- 具体的な解決手順を書く');
  } else if (feedbackType === 'VIDEO') {
    base.push('');
    base.push('## 今回のタスク: 動画付きメッセージへのFB');
    base.push('- 動画内の具体的な発言を引用して改善点を指摘');
    base.push('- 「◯分◯秒あたりで"〜"と言ってるけど、"〜"に変えて」等');
    base.push('- 改善版トークスクリプトをそのまま使える形で渡す');
  }

  return base.join('\n');
}

/**
 * ユーザープロンプト構築（RAG付き）
 */
function buildUserPrompt_(msg, feedbackType, ragExamples, roomId, videoTranscripts) {
  var roomName = '';
  for (var i = 0; i < MONITORED_ROOMS.length; i++) {
    if (MONITORED_ROOMS[i].roomId === roomId) {
      roomName = MONITORED_ROOMS[i].name;
      break;
    }
  }

  var parts = [
    '## 状況',
    'ルーム: ' + roomName,
    '送信者: ' + msg.account.name,
    'メッセージタイプ: ' + (feedbackType || '一般'),
    '',
    '## チームメンバーのメッセージ',
    '```',
    msg.body.substring(0, 2000),
    '```'
  ];

  // 動画トランスクリプトがあれば追加
  if (videoTranscripts && videoTranscripts.length > 0) {
    parts.push('');
    parts.push('## 📹 メッセージ内の動画内容');
    for (var v = 0; v < videoTranscripts.length; v++) {
      parts.push('');
      parts.push('--- 動画' + (v + 1) + ' (ID: ' + videoTranscripts[v].videoId + ') ---');
      parts.push(videoTranscripts[v].transcript);
    }
    parts.push('');
    parts.push('※ 上記は動画の字幕テキストです。動画の内容も踏まえてフィードバックしてください。');
  }

  if (ragExamples && ragExamples.length > 0) {
    parts.push('');
    parts.push('## 参考: ほしの過去の類似フィードバック');
    for (var j = 0; j < ragExamples.length; j++) {
      parts.push('');
      parts.push('--- 例' + (j + 1) + ' (カテゴリ: ' + ragExamples[j].category + ') ---');
      if (ragExamples[j].trigger) {
        parts.push('メンバーの発言: ' + ragExamples[j].trigger.substring(0, 200));
      }
      parts.push('ほしの返答: ' + ragExamples[j].response.substring(0, 400));
    }
  }

  parts.push('');
  parts.push('## 指示');
  parts.push('上記メッセージに対して、ほしのスタイルでフィードバックを生成してください。');
  parts.push('参考例の口調やトーンを踏まえつつ、この具体的な状況に合ったオリジナルの返答を作成してください。');

  return parts.join('\n');
}

// ============================================
// YouTube動画分析
// ============================================

/**
 * メッセージからYouTube URLを抽出
 * @param {string} body
 * @returns {Array<string>} video IDs
 */
function extractYouTubeIds_(body) {
  var patterns = [
    /(?:https?:\/\/)?(?:www\.)?youtube\.com\/watch\?[^\s]*v=([a-zA-Z0-9_-]{11})/g,
    /(?:https?:\/\/)?youtu\.be\/([a-zA-Z0-9_-]{11})/g,
    /(?:https?:\/\/)?(?:www\.)?youtube\.com\/shorts\/([a-zA-Z0-9_-]{11})/g
  ];

  var ids = [];
  var seen = {};
  for (var p = 0; p < patterns.length; p++) {
    var match;
    while ((match = patterns[p].exec(body)) !== null) {
      if (!seen[match[1]]) {
        ids.push(match[1]);
        seen[match[1]] = true;
      }
    }
  }
  return ids;
}

/**
 * YouTube動画のトランスクリプト（字幕）を取得
 * @param {string} videoId
 * @returns {string|null} トランスクリプトテキスト
 */
function fetchYouTubeTranscript_(videoId) {
  try {
    // 1. YouTube動画ページを取得
    var pageUrl = 'https://www.youtube.com/watch?v=' + videoId;
    var response = UrlFetchApp.fetch(pageUrl, {
      muteHttpExceptions: true,
      headers: {
        'Accept-Language': 'ja,en;q=0.9'
      }
    });

    if (response.getResponseCode() !== 200) {
      logBotActivity_('YouTube取得失敗: ' + videoId + ' code=' + response.getResponseCode());
      return null;
    }

    var html = response.getContentText();

    // 2. captionTracksからURLを抽出
    var captionMatch = html.match(/"captionTracks":\s*(\[.*?\])/);
    if (!captionMatch) {
      logBotActivity_('YouTube字幕なし: ' + videoId);
      return null;
    }

    var tracks;
    try {
      tracks = JSON.parse(captionMatch[1]);
    } catch (e) {
      logBotActivity_('YouTube字幕パース失敗: ' + videoId);
      return null;
    }

    if (!tracks || tracks.length === 0) return null;

    // 3. 日本語字幕を優先、なければ最初の字幕
    var trackUrl = null;
    for (var i = 0; i < tracks.length; i++) {
      if (tracks[i].languageCode === 'ja') {
        trackUrl = tracks[i].baseUrl;
        break;
      }
    }
    if (!trackUrl) trackUrl = tracks[0].baseUrl;

    // 4. 字幕XMLを取得
    var captionResp = UrlFetchApp.fetch(trackUrl, { muteHttpExceptions: true });
    if (captionResp.getResponseCode() !== 200) return null;

    var xml = captionResp.getContentText();

    // 5. XMLからテキスト抽出
    var textParts = [];
    var textMatches = xml.match(/<text[^>]*>([\s\S]*?)<\/text>/g);
    if (!textMatches) return null;

    for (var j = 0; j < textMatches.length; j++) {
      var content = textMatches[j].replace(/<[^>]+>/g, '');
      // HTML entities decode
      content = content.replace(/&amp;/g, '&')
                       .replace(/&lt;/g, '<')
                       .replace(/&gt;/g, '>')
                       .replace(/&quot;/g, '"')
                       .replace(/&#39;/g, "'")
                       .replace(/&apos;/g, "'");
      content = content.replace(/\n/g, ' ').trim();
      if (content) textParts.push(content);
    }

    var transcript = textParts.join(' ');

    // 動画タイトルも取得
    var titleMatch = html.match(/<title>(.*?)<\/title>/);
    var title = titleMatch ? titleMatch[1].replace(' - YouTube', '').trim() : '';

    logBotActivity_('YouTube字幕取得: ' + videoId + ' (' + transcript.length + '文字) title=' + title);

    // 長すぎる場合は切り詰め（Claudeのトークン節約）
    if (transcript.length > 5000) {
      transcript = transcript.substring(0, 5000) + '...（以下省略）';
    }

    return (title ? '【動画タイトル】' + title + '\n\n' : '') + transcript;
  } catch (e) {
    logBotError_('YouTube取得例外: ' + videoId + ' ' + e.message);
    return null;
  }
}

/**
 * メッセージ内のYouTube動画からトランスクリプトを一括取得
 * @param {string} body - メッセージ本文
 * @returns {Array<{videoId: string, transcript: string}>}
 */
function getVideoTranscripts_(body) {
  var videoIds = extractYouTubeIds_(body);
  if (videoIds.length === 0) return [];

  var results = [];
  // 最大2動画まで（API負荷対策）
  var limit = Math.min(videoIds.length, 2);
  for (var i = 0; i < limit; i++) {
    var transcript = fetchYouTubeTranscript_(videoIds[i]);
    if (transcript) {
      results.push({ videoId: videoIds[i], transcript: transcript });
    }
  }
  return results;
}

/**
 * テスト用: フィードバック生成テスト（投稿はしない）
 * @param {string} roomId - テスト対象ルームID
 * @param {string} msgBody - テストメッセージ本文
 * @returns {Object} {type, examples, feedback}
 */
function testFeedbackGeneration(roomId, msgBody) {
  roomId = roomId || MONITORED_ROOMS[0].roomId;
  msgBody = msgBody || '商談報告です。本日の商談は見送りになりました。相手の方は興味はあったのですが、資金面で厳しいとのことでした。';

  var feedbackType = detectFeedbackType_(msgBody);
  var examples = searchKnowledgeBase_(msgBody, feedbackType, 3);

  var mockMsg = {
    body: msgBody,
    account: { account_id: '99999', name: 'テストユーザー' }
  };

  var videoTranscripts = getVideoTranscripts_(msgBody);
  var feedback = callClaudeForFeedback_(mockMsg, feedbackType, examples, roomId, videoTranscripts);

  var result = {
    type: feedbackType,
    examplesCount: examples.length,
    videosFound: videoTranscripts.length,
    feedback: feedback
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
