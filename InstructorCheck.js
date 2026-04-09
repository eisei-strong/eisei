// ============================================
// InstructorCheck.js — 講師稼働チェック
// CW rid348824475 の日報投稿を監視
// ============================================

// --- 講師メンバー ---
var INSTRUCTOR_MEMBERS = [
  '一歩', '炭治郎', '一撃', 'エース', 'デク', '鬼スピ', 'ハイエフィ'
];

var INSTRUCTOR_ROOM_ID = '348824475';
var INSTRUCTOR_CALENDAR_ID = 'allforone.namaku@gmail.com';

// --- account_id → 講師名 直接マッピング ---
var INSTRUCTOR_ACCOUNT_MAP = {
  '10933943': '一歩',
  '9333978': '炭治郎',
  '11180179': '一撃',
  '10578812': 'エース',
  '10572112': 'デク',
  '9939321': '鬼スピ',
  '5264121': 'ハイエフィ'
};

// --- アバター取得用の追加ルーム ---
var INSTRUCTOR_AVATAR_ROOMS = ['400352653'];

// --- スプレッドシート保存先 ---
var INSTRUCTOR_SS_ID = '1k_x3aNRTbojmhJZGMS6JGNiTNJLQR4sD5zyJCBh1YqY';
var INSTRUCTOR_WORK_SHEET = '講師稼働ログ';

// --- CWメッセージ取得 ---
function getInstructorMessages_(date) {
  var token = getChatworkToken_();
  if (!token) throw new Error('CW token not found');

  var url = 'https://api.chatwork.com/v2/rooms/' + INSTRUCTOR_ROOM_ID + '/messages?force=1';
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-ChatWorkToken': token },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() === 204) return [];
  if (res.getResponseCode() !== 200) {
    Logger.log('CW API error: ' + res.getResponseCode() + ' ' + res.getContentText());
    return [];
  }

  var messages = JSON.parse(res.getContentText());
  var targetDate = date || new Date();
  var dateStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy-MM-dd');

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

// --- CWルームメンバー取得 ---
function getInstructorRoomMembers_() {
  var token = getChatworkToken_();
  var url = 'https://api.chatwork.com/v2/rooms/' + INSTRUCTOR_ROOM_ID + '/members';
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-ChatWorkToken': token },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return [];
  return JSON.parse(res.getContentText());
}

// --- 全ルームからアバター解決 ---
function resolveInstructorAvatars_() {
  var avatarMap = {};
  var idMap = {}; // account_id → instructor name

  // 1) メインルームから名前マッチ
  var members = getInstructorRoomMembers_();
  for (var i = 0; i < members.length; i++) {
    var m = members[i];
    for (var gi = 0; gi < INSTRUCTOR_MEMBERS.length; gi++) {
      var gName = INSTRUCTOR_MEMBERS[gi];
      if (m.name.indexOf(gName) >= 0) {
        idMap[String(m.account_id)] = gName;
        avatarMap[gName] = m.avatar_image_url || '';
      }
    }
  }

  // 2) account_id直接マッピング
  for (var aid in INSTRUCTOR_ACCOUNT_MAP) {
    idMap[aid] = INSTRUCTOR_ACCOUNT_MAP[aid];
  }

  // 3) 追加ルームからアバター取得（未解決メンバー用）
  var needAvatar = [];
  for (var ni = 0; ni < INSTRUCTOR_MEMBERS.length; ni++) {
    if (!avatarMap[INSTRUCTOR_MEMBERS[ni]]) needAvatar.push(INSTRUCTOR_MEMBERS[ni]);
  }

  if (needAvatar.length > 0) {
    var token = getChatworkToken_();
    for (var ri = 0; ri < INSTRUCTOR_AVATAR_ROOMS.length; ri++) {
      var url = 'https://api.chatwork.com/v2/rooms/' + INSTRUCTOR_AVATAR_ROOMS[ri] + '/members';
      try {
        var res = UrlFetchApp.fetch(url, {
          method: 'get',
          headers: { 'X-ChatWorkToken': token },
          muteHttpExceptions: true
        });
        if (res.getResponseCode() === 200) {
          var extraMembers = JSON.parse(res.getContentText());
          for (var ei = 0; ei < extraMembers.length; ei++) {
            var em = extraMembers[ei];
            var emAid = String(em.account_id);
            // account_idマッピングで名前解決
            if (INSTRUCTOR_ACCOUNT_MAP[emAid] && !avatarMap[INSTRUCTOR_ACCOUNT_MAP[emAid]]) {
              avatarMap[INSTRUCTOR_ACCOUNT_MAP[emAid]] = em.avatar_image_url || '';
            }
            // 名前部分一致でも解決
            for (var nn = 0; nn < needAvatar.length; nn++) {
              if (em.name.indexOf(needAvatar[nn]) >= 0 && !avatarMap[needAvatar[nn]]) {
                avatarMap[needAvatar[nn]] = em.avatar_image_url || '';
                if (!idMap[emAid]) idMap[emAid] = needAvatar[nn];
              }
            }
          }
        }
      } catch (e) {
        Logger.log('Extra room avatar fetch error: ' + e.message);
      }
    }
  }

  return { idMap: idMap, avatarMap: avatarMap };
}

// --- 稼働チェック ---
function checkInstructorActivity_(date) {
  var messages = getInstructorMessages_(date);
  var resolved = resolveInstructorAvatars_();
  var instructorIds = resolved.idMap;
  var instructorAvatars = resolved.avatarMap;

  var posted = {};
  var messageCount = {};
  var firstPostTime = {};
  var lastPostTime = {};
  for (var j = 0; j < messages.length; j++) {
    var msg = messages[j];
    var aid = String(msg.account.account_id);
    if (instructorIds[aid]) {
      var gn = instructorIds[aid];
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

  // 休みメンバー取得
  var offDuty = getInstructorOffDutyMembers_(date);

  var result = [];
  for (var k = 0; k < INSTRUCTOR_MEMBERS.length; k++) {
    var name = INSTRUCTOR_MEMBERS[k];
    var isOff = offDuty.indexOf(name) >= 0;
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
      offDuty: isOff,
      messageCount: count,
      firstPostTime: firstTime,
      lastPostTime: lastTime,
      avatar: instructorAvatars[name] || ''
    });
  }

  return result;
}

// --- Googleカレンダーで休みメンバーを取得 ---
function getInstructorOffDutyMembers_(date) {
  if (!INSTRUCTOR_CALENDAR_ID) return [];
  try {
    var targetDate = date || new Date();
    var startOfDay = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 0, 0, 0);
    var endOfDay = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 23, 59, 59);
    var cal = CalendarApp.getCalendarById(INSTRUCTOR_CALENDAR_ID);
    if (!cal) return [];
    var events = cal.getEvents(startOfDay, endOfDay);
    var offMembers = [];
    for (var i = 0; i < events.length; i++) {
      var title = events[i].getTitle();
      for (var gi = 0; gi < INSTRUCTOR_MEMBERS.length; gi++) {
        if (title.indexOf(INSTRUCTOR_MEMBERS[gi]) >= 0 && (title.indexOf('休') >= 0 || title.indexOf('OFF') >= 0 || title.indexOf('off') >= 0)) {
          offMembers.push(INSTRUCTOR_MEMBERS[gi]);
        }
      }
    }
    return offMembers;
  } catch (e) {
    Logger.log('Calendar error: ' + e.message);
    return [];
  }
}

// --- CW全メッセージから日別・投稿者を集計 ---
function getAllInstructorMessagesByDate_() {
  var token = getChatworkToken_();
  if (!token) return {};

  var url = 'https://api.chatwork.com/v2/rooms/' + INSTRUCTOR_ROOM_ID + '/messages?force=1';
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-ChatWorkToken': token },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() === 204 || res.getResponseCode() !== 200) return {};

  var messages = JSON.parse(res.getContentText());
  var resolved = resolveInstructorAvatars_();
  var idToInstructor = resolved.idMap;

  var result = {};
  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];
    var aid = msg.account ? String(msg.account.account_id) : '';
    if (!idToInstructor[aid]) continue;
    var dateStr = Utilities.formatDate(new Date(msg.send_time * 1000), 'Asia/Tokyo', 'yyyy-MM-dd');
    if (!result[dateStr]) result[dateStr] = {};
    result[dateStr][idToInstructor[aid]] = true;
  }
  return result;
}

// --- 月間データ取得（1日起算） ---
function getInstructorWeeklyData_() {
  var ss = SpreadsheetApp.openById(INSTRUCTOR_SS_ID);
  var sheet = ss.getSheetByName(INSTRUCTOR_WORK_SHEET);

  var today = new Date();
  var year = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy');
  var month = Utilities.formatDate(today, 'Asia/Tokyo', 'MM');
  var todayDay = parseInt(Utilities.formatDate(today, 'Asia/Tokyo', 'd'), 10);
  var prefix = year + '-' + month;

  var activeDays = {};

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

  try {
    var cwDays = getAllInstructorMessagesByDate_();
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

  var results = {};
  for (var d = 1; d <= todayDay; d++) {
    var date = new Date(parseInt(year), parseInt(month) - 1, d);
    var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
    var youbi = ['日','月','火','水','木','金','土'][date.getDay()];
    var dayLabel = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d') + '(' + youbi + ')';

    var members = [];
    for (var k = 0; k < INSTRUCTOR_MEMBERS.length; k++) {
      var gn = INSTRUCTOR_MEMBERS[k];
      var isActive = !!(activeDays[gn] && activeDays[gn][dateStr]);
      members.push({ name: gn, active: isActive, messageCount: 0, lastPostTime: null });
    }
    results[dateStr] = { label: dayLabel, members: members };
  }

  return results;
}

// --- 業務名正規化 ---
var INSTRUCTOR_TASK_KEYWORD_MAP = [
  { keywords: ['添削', 'FB', 'フィードバック'], category: '添削・FB' },
  { keywords: ['ZOOM', 'zoom', 'Zoom', 'ズーム'], category: 'ZOOM対応' },
  { keywords: ['プッシュ', 'push'], category: 'プッシュ対応' },
  { keywords: ['チャット', 'DM', 'LINE'], category: 'チャット・LINE対応' },
  { keywords: ['教材', 'コンテンツ', '台本'], category: '教材・コンテンツ作成' },
  { keywords: ['MTG', 'ミーティング', '会議'], category: 'MTG' },
  { keywords: ['研修', '勉強会'], category: '研修・勉強会' },
  { keywords: ['数値', '報告'], category: '数値・報告' },
  { keywords: ['シェア', '共有'], category: 'シェア・共有' }
];

function normalizeInstructorTaskName_(task) {
  for (var i = 0; i < INSTRUCTOR_TASK_KEYWORD_MAP.length; i++) {
    var entry = INSTRUCTOR_TASK_KEYWORD_MAP[i];
    for (var k = 0; k < entry.keywords.length; k++) {
      if (task.indexOf(entry.keywords[k]) >= 0) {
        return entry.category;
      }
    }
  }
  return task;
}

// --- 日報パーサー ---
function parseInstructorWorkEntries_(body) {
  var lines = String(body || '').split('\n');
  var entries = [];
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
      if (minutes < 0) minutes += 24 * 60;
      if (minutes > 0 && minutes <= 720) {
        entries.push({
          start: match[1] + ':' + match[2],
          end: match[3] + ':' + match[4],
          task: normalizeInstructorTaskName_(task),
          minutes: minutes
        });
      }
    }
  }
  return entries;
}

// --- 業務時間集計（指定日） ---
function getInstructorWorkDetail_(date) {
  var messages = getInstructorMessages_(date);
  var resolved = resolveInstructorAvatars_();
  var instructorIds = resolved.idMap;
  var instructorAvatars = resolved.avatarMap;

  var result = {};
  for (var k = 0; k < INSTRUCTOR_MEMBERS.length; k++) {
    result[INSTRUCTOR_MEMBERS[k]] = { entries: [], totalMinutes: 0, avatar: instructorAvatars[INSTRUCTOR_MEMBERS[k]] || '' };
  }

  for (var j = 0; j < messages.length; j++) {
    var msg = messages[j];
    var aid = msg.account ? String(msg.account.account_id) : '';
    if (!instructorIds[aid]) continue;
    var gn = instructorIds[aid];
    var entries = parseInstructorWorkEntries_(msg.body);
    for (var e = 0; e < entries.length; e++) {
      result[gn].entries.push(entries[e]);
      result[gn].totalMinutes += entries[e].minutes;
    }
  }

  return result;
}

// --- 日次業務時間をスプレッドシートに保存 ---
function saveInstructorDailyWork() {
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var dateStr = Utilities.formatDate(yesterday, 'Asia/Tokyo', 'yyyy-MM-dd');

  var detail = getInstructorWorkDetail_(yesterday);
  var ss = SpreadsheetApp.openById(INSTRUCTOR_SS_ID);
  var sheet = ss.getSheetByName(INSTRUCTOR_WORK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(INSTRUCTOR_WORK_SHEET);
    sheet.getRange(1, 1, 1, 4).setValues([['日付', '名前', '合計分', '詳細JSON']]);
  }

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
  for (var k = 0; k < INSTRUCTOR_MEMBERS.length; k++) {
    var name = INSTRUCTOR_MEMBERS[k];
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
function saveInstructorTodayWork() {
  var today = new Date();
  var dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');

  var detail = getInstructorWorkDetail_(today);
  var ss = SpreadsheetApp.openById(INSTRUCTOR_SS_ID);
  var sheet = ss.getSheetByName(INSTRUCTOR_WORK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(INSTRUCTOR_WORK_SHEET);
    sheet.getRange(1, 1, 1, 4).setValues([['日付', '名前', '合計分', '詳細JSON']]);
  }

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
  for (var k = 0; k < INSTRUCTOR_MEMBERS.length; k++) {
    var name = INSTRUCTOR_MEMBERS[k];
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
function getInstructorMonthlyTotal_(year, month) {
  var ss = SpreadsheetApp.openById(INSTRUCTOR_SS_ID);
  var sheet = ss.getSheetByName(INSTRUCTOR_WORK_SHEET);
  if (!sheet) return {};

  var resolved = resolveInstructorAvatars_();
  var avatarMap = resolved.avatarMap;

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
  for (var k = 0; k < INSTRUCTOR_MEMBERS.length; k++) {
    var n = INSTRUCTOR_MEMBERS[k];
    result[n] = {
      totalMinutes: totals[n] || 0,
      days: dailyData[n] || {},
      avatar: avatarMap[n] || ''
    };
  }
  return result;
}

// --- 業務別月間集計 ---
function getInstructorMonthlyTaskTotal_(year, month) {
  var ss = SpreadsheetApp.openById(INSTRUCTOR_SS_ID);
  var sheet = ss.getSheetByName(INSTRUCTOR_WORK_SHEET);
  if (!sheet) return {};

  var resolved = resolveInstructorAvatars_();
  var avatarMap = resolved.avatarMap;

  var prefix = year + '-' + ('0' + month).slice(-2);
  var data = sheet.getDataRange().getValues();
  var result = {};

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
      if (!result[name]) result[name] = { tasks: {}, avatar: avatarMap[name] || '' };
      for (var e = 0; e < entries.length; e++) {
        var task = entries[e].task;
        result[name].tasks[task] = (result[name].tasks[task] || 0) + entries[e].minutes;
      }
    } catch (err) {}
  }

  return result;
}

// --- 添削回数集計（業務ログから解析、追加API呼び出しなし） ---
// 稼働チャット（room 348824475）の投稿から「添削」「1:1」キーワードと件数を解析
function getInstructorCorrectionStats_() {
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var messages = getInstructorMessages_(new Date());
  var resolved = resolveInstructorAvatars_();
  var instructorIds = resolved.idMap;

  var counts = {};
  for (var k = 0; k < INSTRUCTOR_MEMBERS.length; k++) {
    counts[INSTRUCTOR_MEMBERS[k]] = 0;
  }

  // 添削関連キーワード
  var correctionKeywords = ['添削', '1:1', 'FB', 'フィードバック'];
  // 件数抽出パターン: "3件", "×5", "x3" など
  var countRe = /(\d+)\s*件/;
  var countRe2 = /[×xX]\s*(\d+)/;

  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];
    if (!msg.account) continue;
    var aid = String(msg.account.account_id);
    var name = instructorIds[aid];
    if (!name) continue;

    var body = String(msg.body || '');
    var lines = body.split('\n');

    for (var li = 0; li < lines.length; li++) {
      var line = lines[li];
      // 添削関連キーワードを含む行をチェック
      var isCorrectionLine = false;
      for (var ci = 0; ci < correctionKeywords.length; ci++) {
        if (line.indexOf(correctionKeywords[ci]) >= 0) {
          isCorrectionLine = true;
          break;
        }
      }
      if (!isCorrectionLine) continue;

      // 件数を抽出（"3件" or "×5"）
      var countMatch = line.match(countRe) || line.match(countRe2);
      if (countMatch) {
        counts[name] += parseInt(countMatch[1], 10);
      } else {
        // 件数なし → 1件としてカウント
        counts[name] += 1;
      }
    }
  }

  var totalCorrections = 0;
  for (var n in counts) {
    totalCorrections += counts[n];
  }

  return {
    counts: counts,
    total: totalCorrections,
    date: todayStr,
    updatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm')
  };
}

// --- API: 講師稼働データ取得 ---
function getInstructorData_(params) {
  var type = params.type || 'today';

  if (type === 'today') {
    var activity = checkInstructorActivity_(new Date());
    return { date: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'), members: activity };
  }

  if (type === 'weekly') {
    return getInstructorWeeklyData_();
  }

  if (type === 'date') {
    var dateStr = params.date;
    if (!dateStr) return { error: 'date param required' };
    var targetDate = new Date(dateStr + 'T00:00:00+09:00');
    var activity = checkInstructorActivity_(targetDate);
    return { date: dateStr, members: activity };
  }

  if (type === 'monthly') {
    var now = new Date();
    var year = params.year || Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy');
    var month = params.month || Utilities.formatDate(now, 'Asia/Tokyo', 'M');
    saveInstructorTodayWork();
    var monthly = getInstructorMonthlyTotal_(year, month);
    return { year: year, month: month, members: monthly };
  }

  if (type === 'monthlytasks') {
    var now = new Date();
    var year = params.year || Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy');
    var month = params.month || Utilities.formatDate(now, 'Asia/Tokyo', 'M');
    saveInstructorTodayWork();
    var taskTotals = getInstructorMonthlyTaskTotal_(year, month);
    return { year: year, month: month, tasks: taskTotals };
  }

  if (type === 'corrections') {
    var stats = getInstructorCorrectionStats_();
    return stats;
  }

  if (type === 'workdetail') {
    var dateStr = params.date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    var targetDate = new Date(dateStr + 'T00:00:00+09:00');
    var detail = getInstructorWorkDetail_(targetDate);
    return { date: dateStr, work: detail };
  }

  if (type === 'debug') {
    var roomId = params.room || INSTRUCTOR_ROOM_ID;
    var token = getChatworkToken_();
    var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/members';
    var res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'X-ChatWorkToken': token },
      muteHttpExceptions: true
    });
    var roomMembers = res.getResponseCode() === 200 ? JSON.parse(res.getContentText()) : [];
    var memberNames = [];
    for (var mi = 0; mi < roomMembers.length; mi++) {
      memberNames.push({ name: roomMembers[mi].name, id: roomMembers[mi].account_id, avatar: roomMembers[mi].avatar_image_url || '' });
    }

    if (roomId !== INSTRUCTOR_ROOM_ID) {
      return { roomId: roomId, roomMembers: memberNames };
    }

    var messages = getInstructorMessages_(new Date());
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
    return { roomMembers: memberNames, todayMessages: posters, instructorNames: INSTRUCTOR_MEMBERS };
  }

  return { error: 'unknown type: ' + type };
}

// ============================================
// CW通知: 朝7時 — 前日未報告者
// ============================================
function instructorMorningAlert() {
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var dateLabel = Utilities.formatDate(yesterday, 'Asia/Tokyo', 'M月d日');

  var activity = checkInstructorActivity_(yesterday);
  var offDuty = getInstructorOffDutyMembers_(yesterday);
  var missing = [];
  for (var i = 0; i < activity.length; i++) {
    if (!activity[i].active && offDuty.indexOf(activity[i].name) < 0) {
      missing.push(activity[i].name);
    }
  }

  if (missing.length === 0) {
    Logger.log('全員報告済み（' + dateLabel + '）');
    return;
  }

  var body = '[info][title]⚠️ 講師稼働アラート（' + dateLabel + '）[/title]'
    + '昨日、業務報告がなかったメンバー:\n\n';
  for (var j = 0; j < missing.length; j++) {
    body += '❌ ' + missing[j] + '\n';
  }
  body += '\n確認をお願いします。[/info]';

  sendInstructorNotification_(body);
}

// ============================================
// CW通知: 夕方18時 — 当日未稼働者
// ============================================
function instructorEveningAlert() {
  var today = new Date();
  var dateLabel = Utilities.formatDate(today, 'Asia/Tokyo', 'M月d日');

  var activity = checkInstructorActivity_(today);
  var offDuty = getInstructorOffDutyMembers_(today);
  var missing = [];
  for (var i = 0; i < activity.length; i++) {
    if (!activity[i].active && offDuty.indexOf(activity[i].name) < 0) {
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

  sendInstructorNotification_(body);
}

// ============================================
// CW通知: 1時間無更新チェック
// ============================================
var INSTRUCTOR_NUDGE_MESSAGES = [
  'おーい！1時間経ってるよ！更新して〜！',
  'ちょっと！報告止まってるよ？大丈夫？',
  '1時間サボってない？笑　更新よろしく！',
  'そろそろ更新しよ！止まってるよ〜',
  'おい！手止まってるぞ！報告して！',
  '1時間経過！何やってるか教えて〜！',
  'ストップしてるよ！更新お願い！'
];

function instructorHourlyNudge() {
  var now = new Date();
  var hour = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'H'), 10);
  if (hour < 9 || hour > 20) return;

  var messages = getInstructorMessages_(now);
  var resolved = resolveInstructorAvatars_();
  var idToInstructor = resolved.idMap;

  // 逆引き: 講師名 → account_id
  var instructorToId = {};
  for (var aid in idToInstructor) {
    instructorToId[idToInstructor[aid]] = aid;
  }

  var lastPost = {};
  var hasPosted = {};
  for (var j = 0; j < messages.length; j++) {
    var msg = messages[j];
    var aid = msg.account ? String(msg.account.account_id) : '';
    if (!idToInstructor[aid]) continue;
    var gn = idToInstructor[aid];
    var t = new Date(msg.send_time * 1000);
    hasPosted[gn] = true;
    if (!lastPost[gn] || t > lastPost[gn]) {
      lastPost[gn] = t;
    }
  }

  var offDuty = getInstructorOffDutyMembers_(now);

  var nudgeTargets = [];
  var nowMs = now.getTime();
  for (var k = 0; k < INSTRUCTOR_MEMBERS.length; k++) {
    var name = INSTRUCTOR_MEMBERS[k];
    if (offDuty.indexOf(name) >= 0) continue;
    if (!hasPosted[name]) continue;
    var elapsed = nowMs - lastPost[name].getTime();
    if (elapsed >= 60 * 60 * 1000) {
      nudgeTargets.push(name);
    }
  }

  if (nudgeTargets.length === 0) {
    Logger.log('講師: 全員1時間以内に更新済み');
    return;
  }

  var body = '';
  for (var n = 0; n < nudgeTargets.length; n++) {
    var tName = nudgeTargets[n];
    var tId = instructorToId[tName];
    var randMsg = INSTRUCTOR_NUDGE_MESSAGES[Math.floor(Math.random() * INSTRUCTOR_NUDGE_MESSAGES.length)];
    body += '[To:' + tId + '] ' + tName + '\n' + randMsg + '\n\n';
  }

  sendInstructorNotification_(body.trim());
  Logger.log('講師ナッジ送信: ' + nudgeTargets.join(', '));
}

// --- CW通知送信 ---
function sendInstructorNotification_(body) {
  var token = getChatworkToken_();
  if (!token) { Logger.log('CW token not found'); return; }

  var url = 'https://api.chatwork.com/v2/rooms/' + INSTRUCTOR_ROOM_ID + '/messages';
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
function setupInstructorTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fname = triggers[i].getHandlerFunction();
    if (fname === 'instructorMorningAlert' || fname === 'instructorEveningAlert' ||
        fname === 'instructorHourlyNudge' || fname === 'saveInstructorDailyWork' ||
        fname === 'saveInstructorTodayWork') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 毎日23:50に前日分保存
  ScriptApp.newTrigger('saveInstructorDailyWork')
    .timeBased()
    .atHour(23)
    .nearMinute(50)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  // 毎時、今日の分を更新
  ScriptApp.newTrigger('saveInstructorTodayWork')
    .timeBased()
    .everyHours(1)
    .create();

  // 朝7時トリガー
  ScriptApp.newTrigger('instructorMorningAlert')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  // 夕方18時トリガー
  ScriptApp.newTrigger('instructorEveningAlert')
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  // 3時間おきナッジ
  var nudgeHours = [9, 12, 15, 18];
  for (var h = 0; h < nudgeHours.length; h++) {
    ScriptApp.newTrigger('instructorHourlyNudge')
      .timeBased()
      .atHour(nudgeHours[h])
      .everyDays(1)
      .inTimezone('Asia/Tokyo')
      .create();
  }

  // 添削キャッシュは saveGuardianTodayWork（毎時）に相乗り

  Logger.log('講師トリガー設定完了');
}
