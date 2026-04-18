// ============================================
// GuardianCheck.js — ガーディアン稼働チェック
// CW rid404641188 の日報投稿を監視
// ============================================

// --- ガーディアン除外メンバー（部分一致でルームメンバーから除外） ---
var GUARDIAN_EXCLUDE = [
  'ルルーシュ', '万バズ',
  'ほし', 'ブレスル',
  '押切', '50億',
  'ハヤ',
  '嬴政'
];

var GUARDIAN_LEADER = '星野';
var GUARDIAN_SUB_LEADER = 'レイ';

var GUARDIAN_ROOM_ID = '404641188';

// --- ルームメンバーから除外リスト以外を動的取得（キャッシュ付き） ---
var _guardianMembersCache = null;

function isExcludedMember_(name) {
  for (var j = 0; j < GUARDIAN_EXCLUDE.length; j++) {
    if (name.indexOf(GUARDIAN_EXCLUDE[j]) >= 0) return true;
  }
  return false;
}

function getGuardianMembersInfo_() {
  if (_guardianMembersCache) return _guardianMembersCache;

  var roomMembers = getGuardianRoomMembers_();
  var filtered = [];
  for (var i = 0; i < roomMembers.length; i++) {
    if (!isExcludedMember_(roomMembers[i].name)) {
      filtered.push(roomMembers[i]);
    }
  }
  _guardianMembersCache = filtered;
  return filtered;
}

function getGuardianMemberNames_() {
  var info = getGuardianMembersInfo_();
  var names = [];
  for (var i = 0; i < info.length; i++) {
    names.push(info[i].name);
  }
  return names;
}

// --- CWメッセージ取得 ---
function getGuardianMessages_(date) {
  var token = getChatworkToken_();
  if (!token) throw new Error('CW token not found');

  var url = 'https://api.chatwork.com/v2/rooms/' + GUARDIAN_ROOM_ID + '/messages?force=1';
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-ChatWorkToken': token },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() === 204) return []; // メッセージなし
  if (res.getResponseCode() !== 200) {
    Logger.log('CW API error: ' + res.getResponseCode() + ' ' + res.getContentText());
    return [];
  }

  var messages = JSON.parse(res.getContentText());
  var targetDate = date || new Date();
  var dateStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy-MM-dd');

  // 対象日のメッセージだけフィルタ
  var filtered = [];
  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];
    var msgDate = Utilities.formatDate(
      new Date(msg.send_time * 1000), 'Asia/Tokyo', 'yyyy-MM-dd'
    );
    if (msgDate === dateStr) {
      filtered.push(msg);
    }
  }
  return filtered;
}

// --- CWルームメンバー取得（名前→account_id マッピング用） ---
function getGuardianRoomMembers_() {
  var token = getChatworkToken_();
  var url = 'https://api.chatwork.com/v2/rooms/' + GUARDIAN_ROOM_ID + '/members';
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-ChatWorkToken': token },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return [];
  return JSON.parse(res.getContentText());
}

// --- ガーディアンのaccount_id→名前・アバターマッピングを構築 ---
function buildGuardianIdMap_() {
  var membersInfo = getGuardianMembersInfo_();
  var idToName = {};
  var idToAvatar = {};
  var nameToId = {};
  for (var i = 0; i < membersInfo.length; i++) {
    var m = membersInfo[i];
    var aid = String(m.account_id);
    idToName[aid] = m.name;
    idToAvatar[aid] = m.avatar_image_url || '';
    nameToId[m.name] = aid;
  }
  return { idToName: idToName, idToAvatar: idToAvatar, nameToId: nameToId };
}

// --- 稼働チェック: 指定日に投稿した人/してない人を返す ---
function checkGuardianActivity_(date) {
  var messages = getGuardianMessages_(date);
  var map = buildGuardianIdMap_();
  var memberNames = getGuardianMemberNames_();

  // 投稿者を集計
  var posted = {};
  var messageCount = {};
  var firstPostTime = {};
  var lastPostTime = {};
  for (var j = 0; j < messages.length; j++) {
    var msg = messages[j];
    var aid = String(msg.account.account_id);
    if (map.idToName[aid]) {
      var gn = map.idToName[aid];
      posted[gn] = true;
      messageCount[gn] = (messageCount[gn] || 0) + 1;
      var t = new Date(msg.send_time * 1000);
      if (!firstPostTime[gn] || t < firstPostTime[gn]) {
        firstPostTime[gn] = t;
      }
      if (!lastPostTime[gn] || t > lastPostTime[gn]) {
        lastPostTime[gn] = t;
      }
    }
  }

  // 結果構築
  var result = [];
  for (var k = 0; k < memberNames.length; k++) {
    var name = memberNames[k];
    var isActive = !!posted[name];
    var count = messageCount[name] || 0;
    var firstTime = firstPostTime[name]
      ? Utilities.formatDate(firstPostTime[name], 'Asia/Tokyo', 'HH:mm')
      : null;
    var lastTime = lastPostTime[name]
      ? Utilities.formatDate(lastPostTime[name], 'Asia/Tokyo', 'HH:mm')
      : null;
    result.push({
      name: name,
      active: isActive,
      messageCount: count,
      firstPostTime: firstTime,
      lastPostTime: lastTime,
      avatar: map.idToAvatar[map.nameToId[name]] || '',
      isLeader: name.indexOf(GUARDIAN_LEADER) >= 0,
      isSubLeader: name.indexOf(GUARDIAN_SUB_LEADER) >= 0
    });
  }

  return result;
}

// --- CW全メッセージから日別・投稿者を集計 ---
function getAllGuardianMessagesByDate_() {
  var token = getChatworkToken_();
  if (!token) return {};

  var url = 'https://api.chatwork.com/v2/rooms/' + GUARDIAN_ROOM_ID + '/messages?force=1';
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-ChatWorkToken': token },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() === 204 || res.getResponseCode() !== 200) return {};

  var messages = JSON.parse(res.getContentText());
  var map = buildGuardianIdMap_();

  // 日別・名前でグループ化
  var result = {}; // { dateStr: { name: true } }
  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];
    var aid = msg.account ? String(msg.account.account_id) : '';
    if (!map.idToName[aid]) continue;
    var dateStr = Utilities.formatDate(new Date(msg.send_time * 1000), 'Asia/Tokyo', 'yyyy-MM-dd');
    if (!result[dateStr]) result[dateStr] = {};
    result[dateStr][map.idToName[aid]] = true;
  }
  return result;
}

// --- 月間データ取得（1日起算、CW + スプレッドシート） ---
function getGuardianWeeklyData_() {
  var ss = SpreadsheetApp.openById(GUARDIAN_SS_ID);
  var sheet = ss.getSheetByName(GUARDIAN_WORK_SHEET);
  var memberNames = getGuardianMemberNames_();

  var today = new Date();
  var year = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy');
  var month = Utilities.formatDate(today, 'Asia/Tokyo', 'MM');
  var todayDay = parseInt(Utilities.formatDate(today, 'Asia/Tokyo', 'd'), 10);
  var prefix = year + '-' + month;

  var activeDays = {}; // { name: { dateStr: true } }

  // 1) スプレッドシートから稼働日を取得
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var raw = data[i][0];
      var dateStr = (raw instanceof Date)
        ? Utilities.formatDate(raw, 'Asia/Tokyo', 'yyyy-MM-dd')
        : String(raw);
      if (dateStr.indexOf(prefix) !== 0) continue;
      var name = data[i][1];
      if (!activeDays[name]) activeDays[name] = {};
      activeDays[name][dateStr] = true;
    }
  }

  // 2) CWメッセージ履歴から稼働日を補完
  try {
    var cwDays = getAllGuardianMessagesByDate_();
    for (var cwDate in cwDays) {
      if (cwDate.indexOf(prefix) !== 0) continue;
      for (var cwName in cwDays[cwDate]) {
        if (!activeDays[cwName]) activeDays[cwName] = {};
        activeDays[cwName][cwDate] = true;
      }
    }
  } catch (e) {
    Logger.log('CW backfill error: ' + e.message);
  }

  // 1日〜今日まで結果構築
  var results = {};
  for (var d = 1; d <= todayDay; d++) {
    var date = new Date(parseInt(year), parseInt(month) - 1, d);
    var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
    var youbi = ['日','月','火','水','木','金','土'][date.getDay()];
    var dayLabel = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d') + '(' + youbi + ')';

    var members = [];
    for (var k = 0; k < memberNames.length; k++) {
      var gn = memberNames[k];
      var isActive = !!(activeDays[gn] && activeDays[gn][dateStr]);
      members.push({ name: gn, active: isActive, messageCount: 0, lastPostTime: null });
    }
    results[dateStr] = { label: dayLabel, members: members };
  }

  return results;
}

// --- 業務名正規化: キーワードベースで統一カテゴリに変換 ---
var TASK_KEYWORD_MAP = [
  { keywords: ['LINE対応', '受講者LINE', '受講生LINE', 'チャット対応', 'チャット作成', 'チャット移行', '1:1チャット', '返事対応'], category: 'LINE・チャット対応' },
  { keywords: ['プッシュ割り振り', '漏れ確認', '信託業務確認'], category: 'プッシュ割り振り' },
  { keywords: ['数値入力', '数値続き', '数値報告', '数値記入'], category: '数値入力・報告' },
  { keywords: ['プッシュ報告'], category: 'プッシュ報告入力' },
  { keywords: ['流入経路'], category: '流入経路数値記入' },
  { keywords: ['着金'], category: '着金ランキング' },
  { keywords: ['マニュアル'], category: 'マニュアル作成' },
  { keywords: ['SDZOOM', 'SD zoom', 'SDzoom'], category: 'SDZOOM対応' },
  { keywords: ['炭治郎'], category: '炭治郎部屋関連' },
  { keywords: ['添削'], category: '添削確認' },
  { keywords: ['成長シェア'], category: '成長シェア確認' },
  { keywords: ['契約書'], category: '契約書業務' },
  { keywords: ['万バズ', 'リサーチ'], category: '万バズリサーチ' },
  { keywords: ['ID付与'], category: 'ID付与' },
  { keywords: ['リンク発行'], category: 'リンク発行' },
  { keywords: ['アマギフ'], category: 'アマギフ送付' },
  { keywords: ['郵送'], category: '郵送確認' },
  { keywords: ['入金確認'], category: '入金確認' },
  { keywords: ['決済'], category: '決済対応' },
  { keywords: ['報酬計算'], category: '報酬計算' },
  { keywords: ['タスク確認'], category: 'タスク確認' }
];

function normalizeTaskName_(task) {
  for (var i = 0; i < TASK_KEYWORD_MAP.length; i++) {
    var entry = TASK_KEYWORD_MAP[i];
    for (var k = 0; k < entry.keywords.length; k++) {
      if (task.indexOf(entry.keywords[k]) >= 0) {
        return entry.category;
      }
    }
  }
  return task;
}

// --- 日報パーサー: ✅HH:MM~HH:MM：業務名 を抽出 ---
function parseWorkEntries_(body) {
  var lines = String(body || '').split('\n');
  var entries = [];
  // ✅7:10~7:55：数値入力/報告 のパターン
  var re = /[✅☑️◻️●▶️・\-]?\s*(\d{1,2}):(\d{2})\s*[~〜\-]\s*(\d{1,2}):(\d{2})\s*[：:]\s*(.+)/;
  for (var i = 0; i < lines.length; i++) {
    var match = lines[i].match(re);
    if (match) {
      var startH = parseInt(match[1], 10);
      var startM = parseInt(match[2], 10);
      var endH = parseInt(match[3], 10);
      var endM = parseInt(match[4], 10);
      var task = match[5].replace(/\s+$/,'');
      var minutes = (endH * 60 + endM) - (startH * 60 + startM);
      if (minutes < 0) minutes += 24 * 60; // 日跨ぎ対応
      if (minutes > 0 && minutes <= 720) { // 12時間以内のみ有効
        entries.push({
          start: match[1] + ':' + match[2],
          end: match[3] + ':' + match[4],
          task: normalizeTaskName_(task),
          minutes: minutes
        });
      }
    }
  }
  return entries;
}

// --- 業務時間集計（指定日） ---
function getGuardianWorkDetail_(date) {
  var messages = getGuardianMessages_(date);
  var map = buildGuardianIdMap_();
  var memberNames = getGuardianMemberNames_();

  var result = {};
  for (var k = 0; k < memberNames.length; k++) {
    var n = memberNames[k];
    result[n] = { entries: [], totalMinutes: 0, avatar: map.idToAvatar[map.nameToId[n]] || '' };
  }

  for (var j = 0; j < messages.length; j++) {
    var msg = messages[j];
    var aid = msg.account ? String(msg.account.account_id) : '';
    if (!map.idToName[aid]) continue;
    var gn = map.idToName[aid];
    var entries = parseWorkEntries_(msg.body);
    for (var e = 0; e < entries.length; e++) {
      result[gn].entries.push(entries[e]);
      result[gn].totalMinutes += entries[e].minutes;
    }
  }

  return result;
}

// --- 日次業務時間をスプレッドシートに保存 ---
var GUARDIAN_SS_ID = '1k_x3aNRTbojmhJZGMS6JGNiTNJLQR4sD5zyJCBh1YqY';
var GUARDIAN_WORK_SHEET = 'ガーディアン稼働ログ';

function saveGuardianDailyWork() {
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var dateStr = Utilities.formatDate(yesterday, 'Asia/Tokyo', 'yyyy-MM-dd');

  var detail = getGuardianWorkDetail_(yesterday);
  var ss = SpreadsheetApp.openById(GUARDIAN_SS_ID);
  var sheet = ss.getSheetByName(GUARDIAN_WORK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GUARDIAN_WORK_SHEET);
    sheet.getRange(1, 1, 1, 4).setValues([['日付', '名前', '合計分', '詳細JSON']]);
  }

  // 同日の既存データを確認（重複防止）
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rawDate = data[i][0];
    var cellDate = (rawDate instanceof Date)
      ? Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(rawDate);
    if (cellDate === dateStr) {
      Logger.log(dateStr + ' は保存済み');
      return;
    }
  }

  var rows = [];
  for (var name in detail) {
    var d = detail[name];
    if (d && d.totalMinutes > 0) {
      rows.push([dateStr, name, d.totalMinutes, JSON.stringify(d.entries)]);
    }
  }
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }
  Logger.log(dateStr + ': ' + rows.length + '件保存');
}

// --- 今日の分もライブで保存 ---
function saveGuardianTodayWork() {
  var today = new Date();
  var dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');

  var detail = getGuardianWorkDetail_(today);
  var ss = SpreadsheetApp.openById(GUARDIAN_SS_ID);
  var sheet = ss.getSheetByName(GUARDIAN_WORK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GUARDIAN_WORK_SHEET);
    sheet.getRange(1, 1, 1, 4).setValues([['日付', '名前', '合計分', '詳細JSON']]);
  }

  // 今日の既存データを削除して再書込み
  var data = sheet.getDataRange().getValues();
  var deleteRows = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var rawDate = data[i][0];
    var cellDate = (rawDate instanceof Date)
      ? Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(rawDate);
    if (cellDate === dateStr) {
      deleteRows.push(i + 1);
    }
  }
  for (var d = 0; d < deleteRows.length; d++) {
    sheet.deleteRow(deleteRows[d]);
  }

  var rows = [];
  for (var name in detail) {
    var det = detail[name];
    if (det && det.totalMinutes > 0) {
      rows.push([dateStr, name, det.totalMinutes, JSON.stringify(det.entries)]);
    }
  }
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }
}

// --- 月間集計を取得 ---
function getGuardianMonthlyTotal_(year, month) {
  var ss = SpreadsheetApp.openById(GUARDIAN_SS_ID);
  var sheet = ss.getSheetByName(GUARDIAN_WORK_SHEET);
  if (!sheet) return {};

  var map = buildGuardianIdMap_();
  var memberNames = getGuardianMemberNames_();

  var prefix = year + '-' + ('0' + month).slice(-2);
  var data = sheet.getDataRange().getValues();
  var totals = {};
  var dailyData = {};

  for (var i = 1; i < data.length; i++) {
    var raw = data[i][0];
    var dateStr = (raw instanceof Date)
      ? Utilities.formatDate(raw, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(raw);
    if (dateStr.indexOf(prefix) !== 0) continue;
    var name = data[i][1];
    var minutes = Number(data[i][2]) || 0;
    totals[name] = (totals[name] || 0) + minutes;
    if (!dailyData[name]) dailyData[name] = {};
    dailyData[name][dateStr] = minutes;
  }

  var result = {};
  for (var k = 0; k < memberNames.length; k++) {
    var n = memberNames[k];
    result[n] = {
      totalMinutes: totals[n] || 0,
      days: dailyData[n] || {},
      avatar: map.idToAvatar[map.nameToId[n]] || ''
    };
  }
  return result;
}

// --- 業務別月間集計 ---
function getGuardianMonthlyTaskTotal_(year, month) {
  var ss = SpreadsheetApp.openById(GUARDIAN_SS_ID);
  var sheet = ss.getSheetByName(GUARDIAN_WORK_SHEET);
  if (!sheet) return {};

  var map = buildGuardianIdMap_();

  var prefix = year + '-' + ('0' + month).slice(-2);
  var data = sheet.getDataRange().getValues();
  var result = {}; // { name: { tasks: {}, avatar: '' } }

  for (var i = 1; i < data.length; i++) {
    var raw = data[i][0];
    var dateStr = (raw instanceof Date)
      ? Utilities.formatDate(raw, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(raw);
    if (dateStr.indexOf(prefix) !== 0) continue;
    var name = data[i][1];
    var entriesJson = data[i][3];
    if (!entriesJson) continue;
    try {
      var entries = JSON.parse(entriesJson);
      if (!result[name]) result[name] = { tasks: {}, avatar: map.idToAvatar[map.nameToId[name]] || '' };
      for (var e = 0; e < entries.length; e++) {
        var task = entries[e].task;
        result[name].tasks[task] = (result[name].tasks[task] || 0) + entries[e].minutes;
      }
    } catch (err) {}
  }

  return result;
}

// --- トリガー設定に日次保存を追加 ---
function setupGuardianWorkTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fname = triggers[i].getHandlerFunction();
    if (fname === 'saveGuardianDailyWork' || fname === 'saveGuardianTodayWork') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 毎日23:50に前日分保存
  ScriptApp.newTrigger('saveGuardianDailyWork')
    .timeBased()
    .atHour(23)
    .nearMinute(50)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  // 毎時、今日の分を更新
  ScriptApp.newTrigger('saveGuardianTodayWork')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('稼働ログトリガー設定完了');
}

// --- API: ガーディアン稼働データ取得 ---
function getGuardianData_(params) {
  var type = params.type || 'today';

  if (type === 'today') {
    var activity = checkGuardianActivity_(new Date());
    return { date: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'), members: activity };
  }

  if (type === 'weekly') {
    return getGuardianWeeklyData_();
  }

  if (type === 'date') {
    var dateStr = params.date;
    if (!dateStr) return { error: 'date param required' };
    var targetDate = new Date(dateStr + 'T00:00:00+09:00');
    var activity = checkGuardianActivity_(targetDate);
    return { date: dateStr, members: activity };
  }

  if (type === 'monthly') {
    var now = new Date();
    var year = params.year || Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy');
    var month = params.month || Utilities.formatDate(now, 'Asia/Tokyo', 'M');
    // 今日の分もライブ保存してから集計
    saveGuardianTodayWork();
    var monthly = getGuardianMonthlyTotal_(year, month);
    return { year: year, month: month, members: monthly };
  }

  if (type === 'monthlytasks') {
    var now = new Date();
    var year = params.year || Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy');
    var month = params.month || Utilities.formatDate(now, 'Asia/Tokyo', 'M');
    saveGuardianTodayWork();
    var taskTotals = getGuardianMonthlyTaskTotal_(year, month);
    return { year: year, month: month, tasks: taskTotals };
  }

  if (type === 'workdetail') {
    var dateStr = params.date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    var targetDate = new Date(dateStr + 'T00:00:00+09:00');
    var detail = getGuardianWorkDetail_(targetDate);
    return { date: dateStr, work: detail };
  }

  if (type === 'debug') {
    var members = getGuardianRoomMembers_();
    var memberNames = [];
    for (var mi = 0; mi < members.length; mi++) {
      memberNames.push({ name: members[mi].name, id: members[mi].account_id });
    }
    var messages = getGuardianMessages_(new Date());
    var posters = [];
    for (var pi = 0; pi < messages.length; pi++) {
      var body = String(messages[pi].body || '').substring(0, 50);
      var aid = messages[pi].account ? messages[pi].account.account_id : null;
      var matchedName = '';
      for (var mk = 0; mk < memberNames.length; mk++) {
        if (memberNames[mk].id === aid) { matchedName = memberNames[mk].name; break; }
      }
      posters.push({ id: aid, name: matchedName, body: body, time: messages[pi].send_time });
    }
    return { roomMembers: memberNames, todayMessages: posters, guardianNames: getGuardianMemberNames_(), excludeList: GUARDIAN_EXCLUDE };
  }

  return { error: 'unknown type: ' + type };
}

// ============================================
// CW通知: 朝7時 — 前日未報告者
// ============================================
function guardianMorningAlert() {
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var dateLabel = Utilities.formatDate(yesterday, 'Asia/Tokyo', 'M月d日');

  var activity = checkGuardianActivity_(yesterday);
  var missing = [];
  for (var i = 0; i < activity.length; i++) {
    if (!activity[i].active) {
      missing.push(activity[i].name);
    }
  }

  if (missing.length === 0) {
    Logger.log('全員報告済み（' + dateLabel + '）');
    return;
  }

  var body = '[info][title]⚠️ ガーディアン稼働アラート（' + dateLabel + '）[/title]'
    + '昨日、業務報告がなかったメンバー:\n\n';
  for (var j = 0; j < missing.length; j++) {
    body += '❌ ' + missing[j] + '\n';
  }
  body += '\n確認をお願いします。[/info]';

  sendGuardianNotification_(body);
}

// ============================================
// CW通知: 夕方18時 — 当日未稼働者
// ============================================
function guardianEveningAlert() {
  var today = new Date();
  var dateLabel = Utilities.formatDate(today, 'Asia/Tokyo', 'M月d日');

  var activity = checkGuardianActivity_(today);
  var missing = [];
  for (var i = 0; i < activity.length; i++) {
    if (!activity[i].active) {
      missing.push(activity[i].name);
    }
  }

  if (missing.length === 0) {
    Logger.log('全員稼働確認済み（' + dateLabel + '）');
    return;
  }

  var body = '[info][title]⚠️ 本日の未稼働アラート（' + dateLabel + '）[/title]'
    + '本日まだ業務報告がないメンバー:\n\n';
  for (var j = 0; j < missing.length; j++) {
    body += '❌ ' + missing[j] + '\n';
  }
  body += '\n確認をお願いします。[/info]';

  sendGuardianNotification_(body);
}

// ============================================
// CW通知: 1時間無更新チェック（毎時実行、9〜21時）
// ============================================
var NUDGE_MESSAGES = [
  'おーい！1時間経ってるよ！更新して〜！',
  'ちょっと！報告止まってるよ？大丈夫？',
  '1時間サボってない？笑　更新よろしく！',
  'そろそろ更新しよ！止まってるよ〜',
  'おい！手止まってるぞ！報告して！',
  '1時間経過！何やってるか教えて〜！',
  'ストップしてるよ！更新お願い！'
];

// --- Googleカレンダーで休みメンバーを取得 ---
var GUARDIAN_CALENDAR_ID = ''; // ← 後でカレンダーID設定

function getOffDutyMembers_() {
  if (!GUARDIAN_CALENDAR_ID) return [];
  try {
    var today = new Date();
    var startOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0);
    var endOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
    var cal = CalendarApp.getCalendarById(GUARDIAN_CALENDAR_ID);
    if (!cal) return [];
    var events = cal.getEvents(startOfDay, endOfDay);
    var offMembers = [];
    var memberNames = getGuardianMemberNames_();
    for (var i = 0; i < events.length; i++) {
      var title = events[i].getTitle();
      for (var gi = 0; gi < memberNames.length; gi++) {
        if (title.indexOf(memberNames[gi]) >= 0 && (title.indexOf('休') >= 0 || title.indexOf('OFF') >= 0 || title.indexOf('off') >= 0)) {
          offMembers.push(memberNames[gi]);
        }
      }
    }
    return offMembers;
  } catch (e) {
    Logger.log('Calendar error: ' + e.message);
    return [];
  }
}

function guardianHourlyNudge() {
  var now = new Date();
  var hour = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'H'), 10);

  // 9時〜20時のみ
  if (hour < 9 || hour > 20) return;

  var messages = getGuardianMessages_(now);
  var map = buildGuardianIdMap_();
  var memberNames = getGuardianMemberNames_();

  // 各メンバーの最終投稿時刻を取得
  var lastPost = {};
  var hasPosted = {};
  for (var j = 0; j < messages.length; j++) {
    var msg = messages[j];
    var aid = msg.account ? String(msg.account.account_id) : '';
    if (!map.idToName[aid]) continue;
    var gn = map.idToName[aid];
    var t = new Date(msg.send_time * 1000);
    hasPosted[gn] = true;
    if (!lastPost[gn] || t > lastPost[gn]) {
      lastPost[gn] = t;
    }
  }

  // 休みメンバーを取得
  var offDuty = getOffDutyMembers_();

  // 1時間以上経過しているメンバーを検出（投稿実績がある人のみ、休みは除外）
  var nudgeTargets = [];
  var nowMs = now.getTime();
  for (var k = 0; k < memberNames.length; k++) {
    var name = memberNames[k];
    if (offDuty.indexOf(name) >= 0) continue; // 休みの人はスキップ
    if (!hasPosted[name]) continue; // 今日まだ投稿してない人はスキップ（朝・夕アラートで対応）
    var elapsed = nowMs - lastPost[name].getTime();
    if (elapsed >= 60 * 60 * 1000) { // 1時間以上
      nudgeTargets.push(name);
    }
  }

  if (nudgeTargets.length === 0) {
    Logger.log('全員1時間以内に更新済み');
    return;
  }

  // タメ口でTo付きメッセージ送信
  var body = '';
  for (var n = 0; n < nudgeTargets.length; n++) {
    var tName = nudgeTargets[n];
    var tId = map.nameToId[tName];
    var randMsg = NUDGE_MESSAGES[Math.floor(Math.random() * NUDGE_MESSAGES.length)];
    body += '[To:' + tId + '] ' + tName + '\n' + randMsg + '\n\n';
  }

  sendGuardianNotification_(body.trim());
  Logger.log('ナッジ送信: ' + nudgeTargets.join(', '));
}

// --- CW通知送信 ---
function sendGuardianNotification_(body) {
  var token = getChatworkToken_();
  if (!token) { Logger.log('CW token not found'); return; }

  var url = 'https://api.chatwork.com/v2/rooms/' + GUARDIAN_ROOM_ID + '/messages';
  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'X-ChatWorkToken': token },
    payload: { body: body },
    muteHttpExceptions: true
  });
}

// ============================================
// トリガー設定（1回だけ実行）
// ============================================
function setupGuardianTriggers() {
  // 既存のガーディアントリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fname = triggers[i].getHandlerFunction();
    if (fname === 'guardianMorningAlert' || fname === 'guardianEveningAlert' || fname === 'guardianNudge9' || fname === 'guardianNudge12' || fname === 'guardianNudge15' || fname === 'guardianNudge18') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 朝7時トリガー
  ScriptApp.newTrigger('guardianMorningAlert')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  // 夕方18時トリガー
  ScriptApp.newTrigger('guardianEveningAlert')
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  // 3時間おきナッジ（9, 12, 15, 18時）
  var nudgeHours = [9, 12, 15, 18];
  for (var h = 0; h < nudgeHours.length; h++) {
    ScriptApp.newTrigger('guardianHourlyNudge')
      .timeBased()
      .atHour(nudgeHours[h])
      .everyDays(1)
      .inTimezone('Asia/Tokyo')
      .create();
  }

  Logger.log('ガーディアントリガー設定完了（7時・18時・ナッジ9/12/15/18時）');
}
