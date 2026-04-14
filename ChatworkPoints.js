// ============================================
// ChatworkPoints.js — 共創ポイント（More moreChatから集計）
// ============================================

var CW_API_TOKEN = 'c1a87be06bce046831a9de0c73b4b093';
var CW_ROOM_ID = '419910408';

// Chatwork account_id → v2メンバー名
var CW_NAME_MAP = {
  10258043: 'ヒトコト',
  9398311: 'スマイル',
  10750441: 'ぜんぶり',
  10751140: '意思決定',
  10751530: 'ポジティブ',
  10751652: 'スクリプト',
  10754430: '意思決定',
  11109913: 'ゴン',
  11105287: 'トニー',
  3602185: 'ゴール',
  9855353: 'ハヤ',
  7528759: 'スマイル',
  11159019: '長谷部'
};

/**
 * Chatwork APIからRoom 419910408のメッセージを取得し、
 * 発言者ごとにポイントを集計する。
 * - 1発言 = 1ポイント（累計、月リセットなし）
 * - 返信（[rp ...]で始まる）は除外
 * - 自分自身へのTo（[To:自分ID]のみ含む）は除外
 */
/**
 * メッセージが共創対象かフィルタ判定
 * 対象: [To:xxx]が含まれる（自分以外宛）
 * 対象外: 返信([rp)、アクションプラン(🔸🔶)
 */
function isKyosoMessage_(body, senderId) {
  // 返信は除外
  if (body.indexOf('[rp ') === 0 || body.indexOf('[rp\n') >= 0) return false;
  // アクションプランは除外
  if (body.indexOf('\uD83D\uDD38') >= 0 || body.indexOf('\uD83D\uDD36') >= 0) return false;
  if (body.indexOf('アクションプラン') >= 0) return false;
  // To必須（[To:xxx]タグ or テキストベースの "To〜"）
  var toMatches = body.match(/\[To:(\d+)\]/g);
  var hasToTag = toMatches && toMatches.length > 0;
  var hasTextTo = /(?:^|\n)\s*[Tt][Oo]\s*[^\[\]\s\n]/m.test(body);
  if (!hasToTag && !hasTextTo) return false;
  // 自分自身へのToのみは除外（タグToの場合のみ）
  if (hasToTag && !hasTextTo && toMatches.length === 1) {
    var toId = body.match(/\[To:(\d+)\]/);
    if (toId && parseInt(toId[1]) === senderId) return false;
  }
  // 短い感想・リアクション系は除外（Toタグ・引用・URL除去後の本文が短い）
  var cleanText = body
    .replace(/\[To:\d+\][^\n]*/g, '')
    .replace(/\[qt\][\s\S]*?\[\/qt\]/g, '')
    .replace(/(?:^|\n)\s*[Tt][Oo]\s*[^\n]*/g, '')
    .replace(/https?:\/\/\S+/g, '')
    .replace(/\[info\]|\[\/info\]|\[title\]|\[\/title\]/g, '')
    .replace(/[\s\n！!？?。、.,…]+/g, '')
    .trim();
  if (cleanText.length < 30 && body.indexOf('http') === -1) return false;
  return true;
}

function getChatworkPoints_() {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'cw_points_v2';
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var points = {};
  var messages = fetchChatworkMessages_();

  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];
    var body = msg.body || '';
    var senderId = msg.account ? msg.account.account_id : 0;
    var senderName = CW_NAME_MAP[senderId] || '';

    if (!senderName) continue;
    if (!isKyosoMessage_(body, senderId)) continue;

    points[senderName] = (points[senderName] || 0) + 1;
  }

  // 10分キャッシュ
  try { cache.put(cacheKey, JSON.stringify(points), 600); } catch(e) {}

  return points;
}

/**
 * 共創メッセージ一覧を取得（全件）
 */
function getChatworkMessages_() {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'cw_messages';
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var rawMsgs = fetchChatworkMessages_();
  var result = [];

  for (var i = rawMsgs.length - 1; i >= 0; i--) {
    var msg = rawMsgs[i];
    var body = msg.body || '';
    var senderId = msg.account ? msg.account.account_id : 0;
    var v2Name = CW_NAME_MAP[senderId] || '';
    var displayName = v2Name || (msg.account ? msg.account.name : '');

    if (!displayName) continue;
    if (!isKyosoMessage_(body, senderId)) continue;

    // To先を解決
    var toIds = body.match(/\[To:(\d+)\]/g) || [];
    var toMembers = [];
    for (var ti = 0; ti < toIds.length; ti++) {
      var tid = parseInt(toIds[ti].match(/\d+/)[0]);
      if (tid === senderId) continue;
      var toName = CW_NAME_MAP[tid] || '';
      if (toName) {
        var toIcon = '';
        try { toIcon = iconUrl_(toName); } catch(e) {}
        toMembers.push({ name: toName, icon: toIcon });
      }
    }

    // Chatworkタグを除去して表示用テキストを生成
    var cleanBody = body
      .replace(/\[To:\d+\][^\n]*/g, '')
      .replace(/\[info\]|\[\/info\]|\[title\]|\[\/title\]/g, '')
      .replace(/\[qt\]\[qtmeta[^\]]*\]/g, '<div class="cw-quote">')
      .replace(/\[\/qt\]/g, '</div>')
      .replace(/\[hr\]/g, '───')
      .replace(/\[code\]|\[\/code\]/g, '')
      .trim();

    var sendTime = msg.send_time ? Utilities.formatDate(new Date(msg.send_time * 1000), 'Asia/Tokyo', 'MM/dd HH:mm') : '';

    var icon = '';
    if (v2Name) {
      try { icon = iconUrl_(v2Name); } catch(e) {}
    }

    result.push({
      name: displayName,
      icon: icon,
      to: toMembers,
      body: cleanBody,
      time: sendTime
    });
  }

  // CacheServiceは100KBまでなので、大きすぎる場合は切り詰め
  var jsonStr = JSON.stringify(result);
  if (jsonStr.length < 90000) {
    try { cache.put(cacheKey, jsonStr, 600); } catch(e) {}
  }
  return result;
}

/**
 * Chatwork APIからメッセージを取得
 */
function fetchChatworkMessages_() {
  var url = 'https://api.chatwork.com/v2/rooms/' + CW_ROOM_ID + '/messages?force=1';
  var options = {
    method: 'get',
    headers: { 'X-ChatWorkToken': CW_API_TOKEN },
    muteHttpExceptions: true
  };

  try {
    var resp = UrlFetchApp.fetch(url, options);
    var code = resp.getResponseCode();
    if (code !== 200) return [];
    return JSON.parse(resp.getContentText()) || [];
  } catch(e) {
    return [];
  }
}

/**
 * テスト用: 共創ポイントをログに出力
 */
function testChatworkPoints() {
  var points = getChatworkPoints_();
  Logger.log(JSON.stringify(points, null, 2));
}

// ============================================
// ボトルネック通知（Chatwork送信）
// ============================================

// v2メンバー名 → Chatwork account_id（To送信用）
var V2_TO_CW_ID = {
  'ヒトコト': 10258043,
  '意思決定': 10751140,
  'スマイル': 9398311,
  'ぜんぶり': 10750441,
  'ポジティブ': 10751530,
  'けつだん': 9418659,
  'スクリプト': 10751652,
  'スクリプト通りに営業するくん': 10751652,
  'スクリプトくん': 10751652,
  'トニー': 11105287,
  'ありのまま': 4415237,
  'ゴン': 11109913
};

/**
 * ボトルネック診断を実行してChatworkに通知
 * トリガーまたはAPI経由で呼び出し
 */
function sendBottleneckNotification() {
  // 2026-04-12 停止：ボトルネック通知を無効化
  return { sent: 0, message: 'ボトルネック通知は停止中' };
  var data = getDashboardData();
  var members = data.members || [];
  if (!members.length) return { sent: 0 };

  var cwPoints = getChatworkPoints_();

  // 各指標のランク計算
  var metricDefs = [
    { key: 'revenue',   label: '着金額',   get: function(m) { return m.revenue || 0; } },
    { key: 'closeRate', label: '成約率',   get: function(m) { return m.closeRate || 0; } },
    { key: 'avgPrice',  label: '平均単価', get: function(m) { return m.avgPrice || 0; } },
    { key: 'deals',     label: '商談数',   get: function(m) { return m.deals || 0; } },
    { key: 'closed',    label: '成約数',   get: function(m) { return m.closed || 0; } },
    { key: 'kyoso',     label: '共創pt',   get: function(m) { return cwPoints[m.name] || 0; } }
  ];

  var rankMap = {};
  for (var i = 0; i < members.length; i++) { rankMap[members[i].name] = {}; }
  for (var mi = 0; mi < metricDefs.length; mi++) {
    var met = metricDefs[mi];
    var sorted = members.slice().sort(function(a, b) { return met.get(b) - met.get(a); });
    for (var si = 0; si < sorted.length; si++) {
      rankMap[sorted[si].name][met.key] = si + 1;
    }
  }

  var topHolidays = members.length > 0 ? (members[0].holidays || 0) : 0;
  var n = members.length;

  // 各メンバーの診断
  var notifications = [];
  for (var i = 0; i < members.length; i++) {
    var m = members[i];
    var r = rankMap[m.name];
    var diags = [];

    // 成約数多い + 平均単価低い
    if (r.closed <= Math.ceil(n * 0.4) && r.avgPrice > Math.ceil(n * 0.6)) {
      diags.push({ fact: '成約数' + r.closed + '位(' + (m.closed || 0) + '件)に対して平均単価が' + r.avgPrice + '位(' + (m.avgPrice || 0) + '万円)', action: 'アンカリングを見直せ。提示価格を上げろ。安売りするな。お前の成約力があれば単価を上げるだけで一気にトップ狙える。' });
    }
    // 着金3位以内で平均単価150万未満
    var avgP = m.avgPrice || 0;
    if (r.revenue <= 3 && avgP < 150 && avgP > 0) {
      var already = diags.some(function(d) { return d.fact.indexOf('平均単価') >= 0; });
      if (!already) {
        diags.push({ fact: '着金' + r.revenue + '位(' + (m.revenue || 0) + '万円)なのに平均単価が' + round1_(avgP) + '万円(150万未満)', action: 'アンカリングを上げろ。お前の実力なら単価150万以上いける。安売りで数を稼ぐな。' });
      }
    }
    // 商談数多い + 成約率低い
    if (r.deals <= Math.ceil(n * 0.4) && r.closeRate > Math.ceil(n * 0.6)) {
      diags.push({ fact: '商談数' + r.deals + '位(' + (m.deals || 0) + '件)に対して成約率が' + r.closeRate + '位(' + (m.closeRate || 0) + '%)', action: 'ヒアリングを徹底しろ。相手の理想と現実のギャップを明確にしろ。お前の商談数でヒアリングが噛み合えば一気にブレイクする。' });
    }
    // 成約率高い + 商談数少ない
    if (r.closeRate <= Math.ceil(n * 0.4) && r.deals > Math.ceil(n * 0.6)) {
      diags.push({ fact: '成約率' + r.closeRate + '位(' + (m.closeRate || 0) + '%)に対して商談数が' + r.deals + '位(' + (m.deals || 0) + '件)', action: '商談数が少ない。顧客選ぶな、休みのタイミングを賢く調整しろ（周りが休んでる時出勤しろ）。お前の成約率なら商談数を増やすだけで着金は爆伸びする。' });
    }
    // 商談数多い + 着金低い
    if (r.deals <= Math.ceil(n * 0.3) && r.revenue > Math.ceil(n * 0.6)) {
      diags.push({ fact: '商談数' + r.deals + '位(' + (m.deals || 0) + '件)に対して着金が' + r.revenue + '位(' + (m.revenue || 0) + '万円)', action: '単価を上げろ。アンカリングとクロージングの質を見直せ。お前の行動量があれば質を上げるだけで結果は変わる。' });
    }
    // 共創ptが平均以下
    var myKyosoPt = cwPoints[m.name] || 0;
    var totalKyoso = 0;
    for (var ki = 0; ki < members.length; ki++) { totalKyoso += cwPoints[members[ki].name] || 0; }
    var avgKyoso = n > 0 ? totalKyoso / n : 0;
    if (myKyosoPt < avgKyoso) {
      diags.push({ fact: '共創ptが' + r.kyoso + '位(' + myKyosoPt + 'pt)でチーム平均(' + Math.round(avgKyoso) + 'pt)を下回っている', action: '今の俺じゃ人に言えることはない？ちげえだろ、人に言うから自分が磨かれて言えるようになるんだろ。自分のためだけに生きてるからいつまで経っても自分の人生すら豊かにできねえんだよ。仲間に発信しろ。お前の言葉がお前と仲間を強くする。' });
    }
    // 出勤日数
    var myH = m.holidays || 0;
    if (myH > topHolidays + 2 && r.revenue > 1) {
      var diff = myH - topHolidays;
      diags.push({ fact: '着金1位(' + topHolidays + '日休み)より' + diff + '日多く休んでいる(' + myH + '日休み)', action: '休みを減らせ。あと' + diff + '日稼働を増やせ。お前のスキルなら稼働日が増えるだけで数字は確実に伸びる。' });
    }
    // ライフティ承認率50%未満
    var lfStr = m.lifety || '';
    if (lfStr && lfStr !== '-') {
      var lfParts = String(lfStr).split('/');
      if (lfParts.length === 2) {
        var lfA = parseInt(lfParts[0]), lfT = parseInt(lfParts[1]);
        if (lfT > 0 && !isNaN(lfA)) {
          var lfRate = Math.round((lfA / lfT) * 100);
          if (lfRate < 50) {
            diags.push({ fact: 'ライフティ承認率が' + lfRate + '%(' + lfStr + ')で50%を切っている', action: 'ライフティの審査通過率を上げろ。申込前の事前確認を徹底しろ。お前の営業力があれば通る案件を見極められる。' });
          }
        }
      }
    }
    // 汎用
    for (var mi = 0; mi < metricDefs.length; mi++) {
      var met = metricDefs[mi];
      if (met.key === 'revenue') continue;
      var rank = r[met.key];
      if (rank - r.revenue >= 4) {
        var dup = diags.some(function(d) { return d.fact.indexOf(met.label) >= 0; });
        if (!dup) {
          var genAction = met.label + 'を改善しろ。ここを伸ばせばお前の数字は確実に上がる。';
          if (met.key === 'closeRate') genAction = 'ヒアリングを徹底しろ。相手の理想と現実のギャップを明確にしろ。お前の商談数でヒアリングが噛み合えば一気にブレイクする。';
          if (met.key === 'avgPrice') genAction = 'アンカリングを見直せ。提示価格を上げろ。お前の成約力があれば単価を上げるだけで一気にトップ狙える。';
          if (met.key === 'deals') genAction = '商談数が少ない。顧客選ぶな、休みのタイミングを賢く調整しろ（周りが休んでる時出勤しろ）。お前の成約率なら商談数を増やすだけで着金は爆伸びする。';
          diags.push({ fact: '着金' + r.revenue + '位に対して' + met.label + 'が' + rank + '位', action: genAction });
        }
      }
    }

    // フォールバック: ボトルネックが見つからない場合、最も弱い指標を検出
    if (diags.length === 0) {
      var worstKey = '', worstRank = 0, worstLabel = '';
      for (var mi = 0; mi < metricDefs.length; mi++) {
        var met = metricDefs[mi];
        if (met.key === 'revenue') continue;
        var rank = r[met.key];
        if (rank > worstRank) { worstRank = rank; worstKey = met.key; worstLabel = met.label; }
      }
      if (worstLabel) {
        var worstVal = metricDefs.filter(function(md) { return md.key === worstKey; })[0].get(m);
        diags.push({ fact: worstLabel + 'が' + worstRank + '位（全指標で最も低い順位）', action: worstLabel + 'を改善しろ。ここを伸ばせばお前の数字は確実に上がる。伸びしろがあるのは武器だ。' });
      }
    }
    notifications.push({ member: m, rank: r.revenue, diags: diags });
  }

  if (notifications.length === 0) return { sent: 0, message: 'ボトルネックなし' };

  // Chatworkメッセージ構築
  var body = '[info][title]🚀 👑 ボトルネック通知[/title]';
  for (var ni = 0; ni < notifications.length; ni++) {
    var notif = notifications[ni];
    var cwId = V2_TO_CW_ID[notif.member.name];
    var toTag = cwId ? '[To:' + cwId + ']' : '';

    body += toTag + notif.member.name + '（着金' + notif.rank + '位）\n';
    for (var di = 0; di < notif.diags.length; di++) {
      var d = notif.diags[di];
      body += '\n🌀事実：' + d.fact + '\n';
      body += '\n🔸アクションプラン\n';
      body += d.action + '\n';
    }
    if (ni < notifications.length - 1) body += '\n───────────────\n';
  }
  body += '\n[toall]\n\n👑50億指示\n1.ボトルネックは、解決したら一番インパクトでかい（着金増える）から死ぬ気で改善しろ！\n\n2.メッセージを見たら必ず具体的なアクションプランをこのチャットに書け！\n[/info]';

  // Chatwork API送信
  var url = 'https://api.chatwork.com/v2/rooms/' + CW_ROOM_ID + '/messages';
  var options = {
    method: 'post',
    headers: { 'X-ChatWorkToken': CW_API_TOKEN },
    payload: { body: body },
    muteHttpExceptions: true
  };

  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  Logger.log('ボトルネック通知送信: HTTP ' + code + ' / ' + notifications.length + '人');

  return { sent: notifications.length, httpCode: code };
}

/**
 * ボトルネック通知の毎日21:30トリガーを設定
 * GASエディタから1回だけ実行すればOK
 */
function setupBottleneckTrigger() {
  // 既存トリガー削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendBottleneckNotification') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // 毎日21:30に実行
  ScriptApp.newTrigger('sendBottleneckNotification')
    .timeBased()
    .atHour(21)
    .nearMinute(30)
    .everyDays(1)
    .create();
  Logger.log('ボトルネック通知トリガー設定完了: 毎日21:30');
}
