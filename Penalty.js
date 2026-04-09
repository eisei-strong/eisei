// ============================================
// 幹部チャット ペナルティシステム
// Chatworkルーム425591871への当日投稿を監視
// ============================================

var PENALTY_CW_ROOM_ID = '425591871';
var PENALTY_AMOUNT = 10000; // 1ペナ = 1万円

// CWアカウントID → 名前マッピング（ペナ対象者）
var PENALTY_MEMBERS = {
  '7528759': 'さかね【魯迅】',
  '10541954': '押切｜50億',
  '10733308': '【万バズ】月1000本万再生（ルルーシュ）',
  '11169713': '代表:ほし（ブレスル）'
};

/**
 * 毎日23:55にトリガーで実行
 * Chatworkルームのメッセージを確認し、当日未投稿の対象者にペナルティ通知
 */
function penaltyCheck() {
  var cwToken = PropertiesService.getScriptProperties().getProperty('CHATWORK_POSTAPP_TOKEN');
  if (!cwToken) { Logger.log('CHATWORK_POSTAPP_TOKEN未設定'); return; }

  var now = new Date();
  var monthKey = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM');

  // Chatworkルームのメッセージを取得（当日分）
  var todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime() / 1000;
  var postedIds = penaltyGetPostedIds_(cwToken, PENALTY_CW_ROOM_ID, todayStart);

  // ペナルティ記録を取得
  var props = PropertiesService.getScriptProperties();
  var penaltyData = JSON.parse(props.getProperty('PENALTY_' + monthKey) || '{}');

  var penaltyNames = [];

  for (var cwId in PENALTY_MEMBERS) {
    // 当日投稿済みならスキップ
    if (postedIds[cwId]) continue;

    // ペナルティ加算
    if (!penaltyData[cwId]) penaltyData[cwId] = 0;
    penaltyData[cwId]++;
    penaltyNames.push({ cwId: cwId, name: PENALTY_MEMBERS[cwId], count: penaltyData[cwId] });
  }

  // ペナルティ記録を保存
  props.setProperty('PENALTY_' + monthKey, JSON.stringify(penaltyData));

  // 該当者がいればChatworkに通知
  if (penaltyNames.length > 0) {
    var msg = '';
    for (var p = 0; p < penaltyNames.length; p++) {
      msg += '[To:' + penaltyNames[p].cwId + ']' + penaltyNames[p].name + 'さん\n';
    }
    msg += '\n今日の投稿確認できてないからペナ1（1万円）追加ね。\n\n※自動メッセージ';

    penaltySendChatwork_(cwToken, PENALTY_CW_ROOM_ID, msg);
    Logger.log('ペナルティ通知送信: ' + penaltyNames.length + '名');
  } else {
    Logger.log('本日のペナルティ該当者なし');
  }
}

/**
 * Chatworkルームのメッセージを取得し、投稿済みアカウントIDのセットを返す
 */
function penaltyGetPostedIds_(token, roomId, sinceTimestamp) {
  var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages?force=1';
  var posted = {};
  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'X-ChatWorkToken': token },
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code !== 200) {
      Logger.log('CWメッセージ取得エラー: ' + code + ' ' + res.getContentText());
      return posted;
    }
    var messages = JSON.parse(res.getContentText());
    for (var i = 0; i < messages.length; i++) {
      var msg = messages[i];
      if (msg.send_time >= sinceTimestamp) {
        posted[String(msg.account.account_id)] = true;
      }
    }
  } catch (e) {
    Logger.log('CWメッセージ取得エラー: ' + e.message);
  }
  return posted;
}

/**
 * チャットワークにメッセージ送信
 */
function penaltySendChatwork_(token, roomId, message) {
  var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages';
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { 'X-ChatWorkToken': token },
      payload: { body: message },
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('CW送信エラー room=' + roomId + ': ' + e.message);
  }
}

/**
 * 月末に自動実行：ペナルティ集計を全員に通達
 */
function penaltyMonthlyReport() {
  var cwToken = PropertiesService.getScriptProperties().getProperty('CHATWORK_POSTAPP_TOKEN');
  if (!cwToken) { Logger.log('CHATWORK_POSTAPP_TOKEN未設定'); return; }

  var now = new Date();
  var monthKey = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM');
  var monthLabel = Utilities.formatDate(now, 'Asia/Tokyo', 'M月');

  var props = PropertiesService.getScriptProperties();
  var penaltyData = JSON.parse(props.getProperty('PENALTY_' + monthKey) || '{}');

  var lines = [];
  var totalAll = 0;
  for (var cwId in PENALTY_MEMBERS) {
    var count = penaltyData[cwId] || 0;
    var amount = count * PENALTY_AMOUNT;
    totalAll += amount;
    lines.push(PENALTY_MEMBERS[cwId] + 'さん：ペナ' + count + '回 → ' + (amount / 10000) + '万円');
  }

  if (totalAll === 0) {
    var msg = '[info][title]📊 ' + monthLabel + ' ペナルティ集計[/title]今月はペナルティ該当者なし！全員えらい！🎉[/info]\n\n※自動メッセージ';
  } else {
    var msg = '[info][title]📊 ' + monthLabel + ' ペナルティ集計[/title]' + lines.join('\n') + '[/info]\n\n※自動メッセージ';
  }

  penaltySendChatwork_(cwToken, PENALTY_CW_ROOM_ID, msg);
  Logger.log('月次ペナルティレポート送信完了');
}

/**
 * ペナルティ系トリガーを設定
 * - 毎日23:55にペナルティチェック
 * - 毎月末にレポート送信
 */
function installPenaltyTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'penaltyCheck' || fn === 'penaltyMonthlyReportWrapper') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 毎日23:55にペナルティチェック
  ScriptApp.newTrigger('penaltyCheck')
    .timeBased()
    .atHour(23)
    .nearMinute(55)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  // 毎月末レポート（毎日実行して月末だけ発火）
  ScriptApp.newTrigger('penaltyMonthlyReportWrapper')
    .timeBased()
    .atHour(23)
    .nearMinute(58)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  Logger.log('ペナルティトリガー設定完了');
}

/**
 * 月末判定ラッパー：翌日が1日なら月末レポートを実行
 */
function penaltyMonthlyReportWrapper() {
  var now = new Date();
  var tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  var tomorrowDay = parseInt(Utilities.formatDate(tomorrow, 'Asia/Tokyo', 'dd'));
  if (tomorrowDay === 1) {
    penaltyMonthlyReport();
  }
}
