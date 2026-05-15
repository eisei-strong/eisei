// ============================================
// MicroManagerAI.js — 上司キャラ AI判定 + 文面生成
// ============================================
// Claude API を使って提出物の内容を判定し、上司口調の詰めDMを生成する

// ============================================
// システムプロンプト（上司キャラ）
// ============================================
function buildMicroBossPrompt_() {
  return [
    'あなたは営業チームの直属上司（社長 Namaka／嬴政）です。',
    '営業マイクロマネジメント Bot の本体として、メンバーの提出物（翌朝計画/日報/週次振り返り）を判定し、不合格なら全員ルームで本人を名指しで詰めます。',
    '',
    '## 役割',
    '- 提出物が「合格基準」を満たしているかを判定する',
    '- 不合格なら、その場で短く・直接的に詰める',
    '- 合格なら、簡潔に承認する',
    '',
    '## 口調（厳守）',
    '- タメ口（敬語禁止）',
    '- 命令調「〜しろ」「〜出せ」「〜書け」',
    '- 句点「。」は付けない',
    '- 1メッセージ 3〜5行に収める（長文NG）',
    '- 精神論は受け取らない。「頑張ります」「気合入れます」等は具体性ゼロとして必ず指摘',
    '- 人格攻撃NG。行動と数字に対してだけ詰める',
    '',
    '## 出力フォーマット（JSON のみ・他の文字は一切出さない）',
    '{',
    '  "verdict": "合格" または "不合格",',
    '  "missing": ["欠落項目1", "欠落項目2"],',
    '  "message": "本人に送るChatwork本文（[To:xxx]プレフィックスは付けない、本文のみ）"',
    '}',
    '',
    '## 不合格メッセージの作り方',
    '1. 1行目: なにが欠落してるか1行で',
    '2. 2行目: なぜダメか1行で（精神論/数字なし/抽象的など）',
    '3. 3行目: いつまでに修正出すか命令調で',
    '4. 末尾: 「ペナ1（1万円）」と明記',
    '',
    '## 合格メッセージの作り方',
    '- 1〜2行で簡潔に承認',
    '- 「OK」「確認」「数字明確」など短い言葉',
    '- 称賛しすぎない（毎日合格が普通）'
  ].join('\n');
}

// ============================================
// 判定 API 呼び出し
// ============================================
/**
 * 提出物を判定
 * @param {string} type - MICRO_TYPE_MORNING / DAILY / WEEKLY
 * @param {string} memberName - 表示名
 * @param {string} body - 投稿本文
 * @returns {object|null} - { verdict, missing, message } または null
 */
function microJudgeSubmission_(type, memberName, body) {
  var apiKey = (typeof getClaudeApiKey_ === 'function') ? getClaudeApiKey_() : null;
  if (!apiKey) {
    microBotLog_('ERROR', 'CLAUDE_API_KEY 未取得');
    return null;
  }

  var requirements = MICRO_REQUIREMENTS[type] || [];
  var typeLabel = (
    type === MICRO_TYPE_MORNING ? '翌朝計画（前日深夜1:59締切）' :
    type === MICRO_TYPE_DAILY   ? '日報（当日21:00締切）' :
    type === MICRO_TYPE_WEEKLY  ? '週次振り返り（金曜18:00締切）' : type
  );

  var userPrompt = [
    '## 提出種別',
    typeLabel,
    '',
    '## 提出者',
    memberName,
    '',
    '## 合格基準（全て満たすこと）',
    requirements.map(function (r, i) { return (i + 1) + '. ' + r; }).join('\n'),
    '',
    '## 提出本文',
    body,
    '',
    '## 指示',
    '上記合格基準と提出本文を照合し、JSONフォーマットで判定を返してください。',
    '- 1つでも欠落があれば "不合格"',
    '- "missing" には欠落項目を日本語で配列で列挙',
    '- "message" は本人を詰める/承認する本文（タメ口・命令調・句点なし・3〜5行）'
  ].join('\n');

  var payload = {
    model: MICRO_AI_MODEL,
    max_tokens: MICRO_AI_MAX_TOKENS,
    system: [
      { type: 'text', text: buildMicroBossPrompt_(), cache_control: { type: 'ephemeral' } }
    ],
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
      microBotLog_('ERROR', 'Claude API エラー code=' + code + ' body=' + response.getContentText().substring(0, 300));
      return null;
    }
    var result = JSON.parse(response.getContentText());
    var text = (result.content && result.content[0] && result.content[0].text) || '';

    // JSON抽出（前後の余計な文字があっても拾えるように）
    var jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      microBotLog_('ERROR', 'AI応答からJSON抽出失敗: ' + text.substring(0, 300));
      return null;
    }

    var parsed = JSON.parse(jsonMatch[0]);
    if (!parsed.verdict || !parsed.message) {
      microBotLog_('ERROR', 'AI応答に verdict/message なし: ' + jsonMatch[0].substring(0, 300));
      return null;
    }
    return parsed;
  } catch (e) {
    microBotLog_('ERROR', 'Claude API 例外: ' + e.message);
    return null;
  }
}

// ============================================
// ルール式の詰め文面（締切超過・未提出向け）
// ※ AI を使わずローカルで生成（コスト/レイテンシゼロ）
// ============================================
function microBuildMissedMessage_(type, memberName) {
  if (type === MICRO_TYPE_MORNING) {
    return [
      '2:00。翌朝計画の提出なし',
      '締切は深夜1:59',
      '今日のアポ件数・最優先タスク・気合入れる商談、3点セットで即出せ',
      'ペナ1（1万円）'
    ].join('\n');
  }
  if (type === MICRO_TYPE_DAILY) {
    return [
      '21:30。日報なし',
      '今日の数字も明日の打ち手も投げっぱなしか',
      '今すぐ実数・原因・明日の打ち手・朝の宣言への結果、4点で出せ',
      'ペナ1（1万円）'
    ].join('\n');
  }
  if (type === MICRO_TYPE_WEEKLY) {
    return [
      '金曜18:30。週次振り返りなし',
      '今週なにやって何が動いたか言語化しないと来週も同じ',
      'KPI実績・上手くいったこと・失敗・来週の数字目標、即出せ',
      'ペナ1（1万円）'
    ].join('\n');
  }
  return '提出物なし\nペナ1（1万円）';
}

// ============================================
// 月次ペナ集計の文面（ルール式）
// ============================================
function microBuildMonthlyReport_(monthLabel, penaltyData) {
  var lines = [];
  var total = 0;
  var perfect = [];

  for (var accountId in MICRO_MEMBERS) {
    var name = MICRO_MEMBERS[accountId];
    var rec = penaltyData[accountId] || { morning: 0, daily: 0, weekly: 0 };
    var count = (rec.morning || 0) + (rec.daily || 0) + (rec.weekly || 0);
    var amount = count * MICRO_PENALTY_AMOUNT;
    total += amount;

    if (count === 0) {
      perfect.push(name);
    } else {
      lines.push(
        name + 'さん：ペナ' + count + '回（翌朝' + (rec.morning || 0) +
        '/日報' + (rec.daily || 0) +
        '/週次' + (rec.weekly || 0) + '）→ ' + (amount / 10000) + '万円'
      );
    }
  }

  var body = '[info][title]📊 ' + monthLabel + ' マイクロマネジメント集計[/title]';
  if (lines.length === 0) {
    body += '今月はペナルティ該当者なし\n全員えらい\n';
  } else {
    body += lines.join('\n') + '\n\n合計 ' + (total / 10000) + '万円';
  }
  if (perfect.length > 0) {
    body += '\n\n🛡️ 無傷の侍: ' + perfect.join('、');
  }
  body += '[/info]\n\n※自動メッセージ';
  return body;
}
