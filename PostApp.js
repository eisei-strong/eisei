// ============================================
// 投稿本数入力アプリ API
// ============================================

// 4月ホープ数スプレッドシートID
var POST_APP_SS_ID = '1rQSoM2zu38aPXJHHD6ILEgyM0SEDOXHXZXz0qK_nQuk';

// シート構造定数
// ===== 統合シート方式（2026-05-02 移行） =====
// 1シートで全月分の列を保持。月初トリガーで右側に列追加していく。
var POST_APP_SHEET_NAME = '投稿数';
var POST_APP_HOPE_SHEET_NAME = 'ホープ数';
var POST_APP_PUSH_SHEET_NAME = 'プッシュ数';
var POST_APP_AUTH_SHEET = '認証';
var POST_APP_ID_COL = 1;      // A列: ID
var POST_APP_NAME_COL = 4;    // D列: 名前 (背景色=リスト数権限判定)
var POST_APP_TOTAL_COL = 5;   // E列: 全期間合計（月別合計列追加後 2026-05-04）
var POST_APP_MONTHLY_TOTAL_START_COL = 6;  // F列: 4月合計（月別合計の起点）
var POST_APP_DATE_START_COL = 15;  // O列: 運用開始月の1日（月別合計列9個分シフト後）

// 運用開始年月（この月の1日からF列に記録開始）
var POST_APP_RUN_START_YEAR = 2026;
var POST_APP_RUN_START_MONTH = 4;

// 当月日数（後方互換）
var POST_APP_MONTH_DAYS = (function(){
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
})();

/**
 * 投稿数シートの指定月の列範囲を返す
 * @param {number} [year] 省略時は当年
 * @param {number} [month] 省略時は当月（1-12）
 */
function getMonthColRange_(year, month) {
  var now = new Date();
  if (!year) year = now.getFullYear();
  if (!month) month = now.getMonth() + 1;
  var startCol = POST_APP_DATE_START_COL;
  var startYm = POST_APP_RUN_START_YEAR * 12 + (POST_APP_RUN_START_MONTH - 1);
  var ym = year * 12 + (month - 1);
  for (var i = startYm; i < ym; i++) {
    var y = Math.floor(i / 12);
    var m = (i % 12) + 1;
    startCol += new Date(y, m, 0).getDate();
  }
  var monthDays = new Date(year, month, 0).getDate();
  return { year: year, month: month, startCol: startCol, monthDays: monthDays, endCol: startCol + monthDays - 1 };
}

/** 後方互換: 当月の列範囲 */
function getCurrentMonthColRange_() {
  return getMonthColRange_();
}

/**
 * ホープ数/プッシュ数シートの当月列範囲を返す（1日3列 = YT/IG/TT）
 * シート種別ごとにデータ起点列が違う（重要）：
 *   ホープ数: L列(12) - メタは A=ID, B=名前, C=4月合計, D=5月合計, ..., K=12月合計
 *   プッシュ数: L列(12) - メタは A=ID, B〜J=名前(マージ等), K=合計
 *
 * 月別合計列追加（2026-05-03）後はホープ数もL列起点に統一。
 */
function getCurrentMonthHopeColRange_(sheetName, year, month) {
  // ホープ数: 全期間合計列追加後 M列(13)起点 / プッシュ数: L列(12)起点
  var dataStartCol = (sheetName === POST_APP_PUSH_SHEET_NAME) ? 12 : 13;
  var COLS_PER_DAY = 3;
  var now = new Date();
  if (!year) year = now.getFullYear();
  if (!month) month = now.getMonth() + 1;
  var startCol = dataStartCol;
  var startYm = POST_APP_RUN_START_YEAR * 12 + (POST_APP_RUN_START_MONTH - 1);
  var ym = year * 12 + (month - 1);
  for (var i = startYm; i < ym; i++) {
    var y = Math.floor(i / 12);
    var m = (i % 12) + 1;
    startCol += new Date(y, m, 0).getDate() * COLS_PER_DAY;
  }
  var monthDays = new Date(year, month, 0).getDate();
  var totalCols = monthDays * COLS_PER_DAY;
  return { startCol: startCol, monthDays: monthDays, totalCols: totalCols, endCol: startCol + totalCols - 1, colsPerDay: COLS_PER_DAY };
}

/** 列番号 → 列文字（A,B,...AA,AB...） */
function postAppColToLetter_(col) {
  var s = '';
  while (col > 0) {
    var m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - m) / 26);
  }
  return s;
}

// LP投稿本数シート（受講生マスター）
var LP_SS_ID = '1LP_eye2PMswK1OuGfCpzALJkiRE7gsAvrdFjrj5zoik';
var LP_POST_SHEET_GID = 655724533;
var LP_POST_ID_COL = 13;    // M列: ID
var LP_POST_NAME_COL = 16;  // P列: 名前

// SMCマスター（契約日取得用）
var SMC_MASTER_ID_COL = 2;       // B列: ID (1-indexed)
var SMC_MASTER_CONTRACT_COL = 36; // AJ列: 契約日 (1-indexed)
var SMC_MASTER_EMAIL_COL = 35;   // AI列: 契約メールアドレス (1-indexed)

// ---- 認証ヘルパー（PropertiesService版） ----

function getAuthData_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('postapp_auth');
  return raw ? JSON.parse(raw) : {};
}

function saveAuthData_(data) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('postapp_auth', JSON.stringify(data));
}

// 旧認証シートからPropertiesServiceへ移行（初回のみ）
function migrateAuthToProps_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('postapp_auth_migrated')) return;
  try {
    var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
    var sheet = ss.getSheetByName(POST_APP_AUTH_SHEET);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        var auth = {};
        for (var i = 0; i < data.length; i++) {
          var id = String(data[i][0]).trim();
          if (id) auth[id] = String(data[i][1]);
        }
        saveAuthData_(auth);
      }
    }
  } catch (e) { /* 認証シートが読めなくても続行 */ }
  props.setProperty('postapp_auth_migrated', 'true');
}

function hashPassword_(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw);
  var hex = '';
  for (var i = 0; i < raw.length; i++) {
    var b = (raw[i] + 256) % 256;
    hex += ('0' + b.toString(16)).slice(-2);
  }
  return hex;
}

function verifyToken_(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  return cache.get('postapp_token_' + token);
}

function postAppResetAuth_() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('postapp_auth');
  props.deleteProperty('postapp_auth_migrated');
  return { ok: true };
}

// ---- ID確認 ----

function postAppCheckId_(id) {
  if (!id) return { error: 'IDを入力してください' };
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) return { error: 'シートが見つかりません' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: '受講生データがありません' };

  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) {
      var name = sheet.getRange(i + 2, POST_APP_NAME_COL).getValue();
      migrateAuthToProps_();
      var authMap = getAuthData_();
      var hasPassword = !!authMap[String(id).trim()];
      var contract = sheet.getRange(i + 2, 2).getValue(); // B列: 契約日
      var contractStr = '';
      if (contract instanceof Date) {
        contractStr = Utilities.formatDate(contract, 'Asia/Tokyo', 'yyyy/MM/dd');
      } else if (contract) {
        contractStr = String(contract);
      }
      return { ok: true, id: String(id).trim(), name: String(name), hasPassword: hasPassword, contract: contractStr };
    }
  }
  return { error: 'IDが見つかりません。正しい受講生IDを入力してください。' };
}

// ---- パスワード登録 ----

function postAppRegister_(id, password) {
  if (!id || !password) return { error: 'IDとパスワードを入力してください' };
  if (password.length < 4) return { error: 'パスワードは4文字以上で設定してください' };

  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  var found = false;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) { found = true; break; }
  }
  if (!found) return { error: 'IDが見つかりません' };

  migrateAuthToProps_();
  var authMap = getAuthData_();
  authMap[String(id).trim()] = hashPassword_(password);
  saveAuthData_(authMap);
  return { ok: true };
}

// ---- パスワードリセット（メールアドレス照合） ----

function postAppResetPassword_(id, email) {
  if (!id || !email) return { error: 'IDとメールアドレスを入力してください' };

  // SMCマスターでメールアドレスを照合
  var smcSs = SpreadsheetApp.openById(SMC_SS_ID);
  var smcSheet = smcSs.getSheetByName(SMC_MASTER_SHEET);
  if (!smcSheet) return { error: 'マスターデータが見つかりません' };

  var smcLastRow = smcSheet.getLastRow();
  var smcData = smcSheet.getRange(2, SMC_MASTER_ID_COL, smcLastRow - 1, SMC_MASTER_EMAIL_COL - SMC_MASTER_ID_COL + 1).getValues();
  var matched = false;
  for (var i = 0; i < smcData.length; i++) {
    var sid = String(smcData[i][0] || '').trim();
    var semail = String(smcData[i][SMC_MASTER_EMAIL_COL - SMC_MASTER_ID_COL] || '').trim().toLowerCase();
    if (sid === String(id).trim() && semail === email.trim().toLowerCase()) {
      matched = true;
      break;
    }
  }
  if (!matched) return { error: 'メールアドレスが一致しません' };

  // PropertiesServiceからパスワードを削除
  migrateAuthToProps_();
  var authMap = getAuthData_();
  delete authMap[String(id).trim()];
  saveAuthData_(authMap);

  return { ok: true };
}

// ---- ログイン（トークン＋全月データ一括返却） ----

function postAppLogin_(id, password) {
  if (!id) return { error: 'IDを入力してください' };
  if (!password) return { error: 'パスワードを入力してください' };

  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) return { error: 'シートが見つかりません' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: '受講生データがありません' };

  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  var memberRow = -1;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) { memberRow = i + 2; break; }
  }
  if (memberRow < 0) return { error: 'IDが見つかりません' };

  migrateAuthToProps_();
  var authMap = getAuthData_();
  var storedHash = authMap[String(id).trim()];
  if (!storedHash) return { error: 'パスワードが未設定です。先にパスワードを登録してください。' };
  var hash = hashPassword_(password);
  if (storedHash !== hash) return { error: 'パスワードが正しくありません' };

  // トークン生成
  var token = hashPassword_(id + '_' + new Date().getTime() + '_' + Math.random());
  var cache = CacheService.getScriptCache();
  cache.put('postapp_token_' + token, id, 86400);

  // ログイン時に全月データも返す（API呼び出し削減）
  var name = sheet.getRange(memberRow, POST_APP_NAME_COL).getValue();
  var monthData = getMonthData_(sheet, memberRow);

  // ログイン日を記録
  recordLogin_(String(id).trim());
  var loginData = getLoginStreakData_(String(id).trim());

  return { ok: true, id: String(id).trim(), name: String(name), token: token, month: monthData, login: loginData };
}

// ---- 全月データ取得（ログイン後のリフレッシュ用） ----

function postAppGet_(token, year, month) {
  var id = verifyToken_(token);
  if (!id) return { error: 'セッション切れです。再ログインしてください。' };

  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) return { error: 'シートが見つかりません' };

  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  var rowIdx = -1;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) { rowIdx = i; break; }
  }
  if (rowIdx < 0) return { error: 'IDが見つかりません' };

  var row = rowIdx + 2;
  var name = sheet.getRange(row, POST_APP_NAME_COL).getValue();
  var monthData = getMonthData_(sheet, row, year, month);

  // ログイン記録は当月のときのみ
  var now = new Date();
  var isCurrent = !year || (parseInt(year) === now.getFullYear() && parseInt(month) === now.getMonth() + 1);
  var loginData = null;
  if (isCurrent) {
    recordLogin_(String(id).trim());
    loginData = getLoginStreakData_(String(id).trim());
  }

  return { ok: true, id: String(id).trim(), name: String(name), month: monthData, login: loginData };
}

// ---- 月データ一括取得ヘルパー ----

function getMonthData_(sheet, row, year, month) {
  var range = getMonthColRange_(year ? parseInt(year) : null, month ? parseInt(month) : null);
  var headers = sheet.getRange(1, range.startCol, 1, range.monthDays).getValues()[0];
  var values = sheet.getRange(row, range.startCol, 1, range.monthDays).getValues()[0];

  // total: 当月はE列(リアルタイム数式)、過去月は範囲から集計
  var now = new Date();
  var isCurrent = (range.year === now.getFullYear() && range.month === now.getMonth() + 1);
  var total;
  if (isCurrent) {
    total = sheet.getRange(row, POST_APP_TOTAL_COL).getValue();
  } else {
    total = 0;
    for (var v = 0; v < values.length; v++) {
      var s = String(values[v] || '').trim();
      if (s && s !== '❌') total += parseInt(s) || 0;
    }
  }

  var days = [];
  for (var i = 0; i < headers.length; i++) {
    days.push({
      col: i,
      label: (headers[i] instanceof Date) ? (headers[i].getMonth() + 1) + '/' + headers[i].getDate() : String(headers[i]),
      value: String(values[i] || '❌')
    });
  }
  var contract = sheet.getRange(row, 2).getValue(); // B列: 契約日
  var contractStr = '';
  if (contract instanceof Date) {
    contractStr = Utilities.formatDate(contract, 'Asia/Tokyo', 'yyyy/MM/dd');
  } else if (contract) {
    contractStr = String(contract);
  }
  return { year: range.year, month: range.month, days: days, total: total, contract: contractStr };
}

// ---- 保存（任意の月の日付を指定可能） ----

function postAppSave_(token, value, col, year, month) {
  var id = verifyToken_(token);
  if (!id) return { error: 'セッション切れです。再ログインしてください。' };

  var validValues = ['❌', '1本', '2本', '3本', '4本', '5本', '6本', '7本', '8本', '9本', '10本'];
  if (validValues.indexOf(value) < 0) return { error: '無効な値です: ' + value };

  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) return { error: 'シートが見つかりません' };

  // col = 指定月の日数内の0始まりインデックス
  var colNum = parseInt(col);
  var range = getMonthColRange_(year ? parseInt(year) : null, month ? parseInt(month) : null);
  if (isNaN(colNum) || colNum < 0 || colNum >= range.monthDays) return { error: '日付が無効です' };
  var targetCol = range.startCol + colNum;

  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  var rowIdx = -1;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) { rowIdx = i; break; }
  }
  if (rowIdx < 0) return { error: 'IDが見つかりません' };

  var row = rowIdx + 2;
  var cell = sheet.getRange(row, targetCol);
  cell.clearDataValidations();
  cell.setValue(value);
  SpreadsheetApp.flush();
  var total = sheet.getRange(row, POST_APP_TOTAL_COL).getValue();

  return { ok: true, saved: value, col: colNum, total: total };
}

// ============================================
// 連続ログイン記録
// ============================================

// ログイン日を記録（PropertiesService）
function recordLogin_(id) {
  var props = PropertiesService.getScriptProperties();
  var key = 'postapp_logins_' + id;
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  var raw = props.getProperty(key);
  var dates = raw ? JSON.parse(raw) : [];

  if (dates.indexOf(today) < 0) {
    dates.push(today);
    // 直近90日分だけ保持
    if (dates.length > 90) {
      dates = dates.slice(dates.length - 90);
    }
    props.setProperty(key, JSON.stringify(dates));
  }
}

// 連続ログインデータを返す
function getLoginStreakData_(id) {
  var props = PropertiesService.getScriptProperties();
  var key = 'postapp_logins_' + id;
  var raw = props.getProperty(key);
  var dates = raw ? JSON.parse(raw) : [];

  // 今日の日付（JST）
  var now = new Date();
  var todayStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');

  // 日付をソート
  dates.sort();

  // 現在の連続日数（今日から遡る）
  var streak = 0;
  // todayStrのインデックスから遡る（日付文字列ベースで計算、タイムゾーン二重変換を回避）
  var todayParts = todayStr.split('-');
  var checkY = parseInt(todayParts[0]);
  var checkM = parseInt(todayParts[1]) - 1;
  var checkD = parseInt(todayParts[2]);

  while (true) {
    var cd = new Date(checkY, checkM, checkD);
    var ds = checkY + '-' + ('0' + (cd.getMonth() + 1)).slice(-2) + '-' + ('0' + cd.getDate()).slice(-2);
    if (dates.indexOf(ds) >= 0) {
      streak++;
      checkD--;
    } else {
      break;
    }
  }

  // 最長連続日数
  var maxStreak = 0;
  var cur = 1;
  for (var i = 1; i < dates.length; i++) {
    var prev = new Date(dates[i - 1] + 'T00:00:00+09:00');
    var curr = new Date(dates[i] + 'T00:00:00+09:00');
    var diff = (curr.getTime() - prev.getTime()) / 86400000;
    if (diff === 1) {
      cur++;
    } else {
      if (cur > maxStreak) maxStreak = cur;
      cur = 1;
    }
  }
  if (cur > maxStreak) maxStreak = cur;

  // 今月のログイン日
  var year = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy'));
  var month = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'MM'));
  var prefix = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM');
  var monthDates = [];
  for (var j = 0; j < dates.length; j++) {
    if (dates[j].indexOf(prefix) === 0) {
      monthDates.push(parseInt(dates[j].split('-')[2]));
    }
  }

  // 月の日数
  var daysInMonth = new Date(year, month, 0).getDate();
  var todayDay = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'dd'));

  return {
    streak: streak,
    maxStreak: maxStreak,
    totalThisMonth: monthDates.length,
    monthDates: monthDates,
    daysInMonth: daysInMonth,
    year: year,
    month: month,
    today: todayDay
  };
}

// 連続ログインデータ取得API
function postAppLoginStreak_(token) {
  var id = verifyToken_(token);
  if (!id) return { error: 'セッション切れです。再ログインしてください。' };

  recordLogin_(id);
  var loginData = getLoginStreakData_(id);
  return { ok: true, login: loginData };
}

// ============================================
// LP投稿本数 → 4月ホープ数 受講生同期
// ============================================

/**
 * LP投稿本数シートから受講生を4月ホープ数シートに同期
 * GASエディタから手動実行
 */
function syncMembersToPostApp() {
  // LP投稿本数シートを開く
  var lpSs = SpreadsheetApp.openById(LP_SS_ID);
  var lpSheet = null;
  var sheets = lpSs.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === LP_POST_SHEET_GID) { lpSheet = sheets[i]; break; }
  }
  if (!lpSheet) { Logger.log('投稿本数シートが見つかりません'); return; }

  // LP投稿本数からID+名前を取得（重複排除）
  var lpLastRow = lpSheet.getLastRow();
  var lpData = lpSheet.getRange(2, LP_POST_ID_COL, lpLastRow - 1, LP_POST_NAME_COL - LP_POST_ID_COL + 1).getValues();
  var members = {};
  for (var r = 0; r < lpData.length; r++) {
    var id = String(lpData[r][0] || '').trim(); // M列（ID）
    var name = String(lpData[r][LP_POST_NAME_COL - LP_POST_ID_COL] || '').trim(); // P列（名前）
    // IDが数字のみの行だけ取り込む（チーム名・ティア名を除外）
    if (id && name && /^\d+$/.test(id) && !members[id]) {
      members[id] = name;
    }
  }
  Logger.log('LP投稿本数のユニーク受講生数: ' + Object.keys(members).length);

  // 4月ホープ数シートを開く
  var postSs = SpreadsheetApp.openById(POST_APP_SS_ID);
  var postSheet = postSs.getSheetByName(POST_APP_SHEET_NAME);
  if (!postSheet) { Logger.log('ホープ数シートが見つかりません'); return; }

  // 既存IDを取得
  var postLastRow = postSheet.getLastRow();
  var existingIds = {};
  if (postLastRow >= 2) {
    var postIds = postSheet.getRange(2, POST_APP_ID_COL, postLastRow - 1, 1).getValues();
    for (var p = 0; p < postIds.length; p++) {
      var pid = String(postIds[p][0] || '').trim();
      if (pid) existingIds[pid] = true;
    }
  }
  Logger.log('既存の受講生数: ' + Object.keys(existingIds).length);

  // SMCマスターから契約日を取得（ID→契約日マップ）
  var smcSs = SpreadsheetApp.openById(SMC_SS_ID);
  var smcSheet = smcSs.getSheetByName(SMC_MASTER_SHEET);
  var contractMap = {};
  if (smcSheet) {
    var smcLastRow = smcSheet.getLastRow();
    if (smcLastRow >= 2) {
      var smcIds = smcSheet.getRange(2, SMC_MASTER_ID_COL, smcLastRow - 1, 1).getValues();
      var smcDates = smcSheet.getRange(2, SMC_MASTER_CONTRACT_COL, smcLastRow - 1, 1).getValues();
      for (var s = 0; s < smcIds.length; s++) {
        var sid = String(smcIds[s][0] || '').trim();
        var sdate = smcDates[s][0];
        if (sid && sdate) {
          if (sdate instanceof Date) {
            contractMap[sid] = Utilities.formatDate(sdate, 'Asia/Tokyo', 'yyyy/MM/dd');
          } else {
            contractMap[sid] = String(sdate);
          }
        }
      }
    }
  }
  Logger.log('マスターの契約日データ: ' + Object.keys(contractMap).length + '件');

  // --- 既存行に契約日を補完 ---
  var POST_APP_CONTRACT_COL = 2; // B列: 契約日
  if (postLastRow >= 2) {
    var existIds = postSheet.getRange(2, POST_APP_ID_COL, postLastRow - 1, 1).getValues();
    var existContracts = postSheet.getRange(2, POST_APP_CONTRACT_COL, postLastRow - 1, 1).getValues();
    for (var ec = 0; ec < existIds.length; ec++) {
      var ecId = String(existIds[ec][0] || '').trim();
      var ecVal = String(existContracts[ec][0] || '').trim();
      if (ecId && !ecVal && contractMap[ecId]) {
        postSheet.getRange(ec + 2, POST_APP_CONTRACT_COL).setValue(contractMap[ecId]);
      }
    }
    Logger.log('既存行の契約日を補完しました');
  }

  // 不足者を追加
  var toAdd = [];
  for (var mid in members) {
    if (!existingIds[mid]) {
      toAdd.push({ id: mid, name: members[mid], contract: contractMap[mid] || '' });
    }
  }

  if (toAdd.length === 0) {
    Logger.log('追加対象なし。全員揃っています。');
    // 不要行削除 + ソートだけ実行
    cleanupPostApp_(postSheet);
    sortPostAppByContract_(postSheet);
    return;
  }

  Logger.log('追加する受講生: ' + toAdd.length + '人');

  // ヘッダー確認（1行目にない場合は作成）
  var header = postSheet.getRange(1, 1, 1, POST_APP_DATE_START_COL + POST_APP_MONTH_DAYS - 1).getValues()[0];
  if (!header[0]) {
    var h = new Array(POST_APP_DATE_START_COL + POST_APP_MONTH_DAYS - 1).fill('');
    h[POST_APP_ID_COL - 1] = 'ID';
    h[POST_APP_CONTRACT_COL - 1] = '契約日';
    h[POST_APP_NAME_COL - 1] = '名前';
    h[POST_APP_TOTAL_COL - 1] = '合計';
    for (var d = 0; d < POST_APP_MONTH_DAYS; d++) {
      h[POST_APP_DATE_START_COL - 1 + d] = '4/' + (d + 1);
    }
    postSheet.getRange(1, 1, 1, h.length).setValues([h]);
  }
  // B1ヘッダーが空なら追加
  if (!postSheet.getRange(1, POST_APP_CONTRACT_COL).getValue()) {
    postSheet.getRange(1, POST_APP_CONTRACT_COL).setValue('契約日');
  }

  // 受講生を追加（ID, 契約日, 名前, 各日付を❌で初期化）
  var insertRow = Math.max(postLastRow + 1, 2);
  var totalCols = POST_APP_DATE_START_COL + POST_APP_MONTH_DAYS - 1;
  var newRows = [];
  for (var a = 0; a < toAdd.length; a++) {
    var row = new Array(totalCols).fill('');
    row[POST_APP_ID_COL - 1] = toAdd[a].id;
    row[POST_APP_CONTRACT_COL - 1] = toAdd[a].contract;
    row[POST_APP_NAME_COL - 1] = toAdd[a].name;
    row[POST_APP_TOTAL_COL - 1] = '';
    for (var dd = 0; dd < POST_APP_MONTH_DAYS; dd++) {
      row[POST_APP_DATE_START_COL - 1 + dd] = '❌';
    }
    newRows.push(row);
  }

  postSheet.getRange(insertRow, 1, newRows.length, totalCols).setValues(newRows);
  Logger.log(toAdd.length + '人を追加しました（行 ' + insertRow + '〜' + (insertRow + toAdd.length - 1) + '）');

  // --- 不要行削除 + 契約日順にソート ---
  cleanupPostApp_(postSheet);
  sortPostAppByContract_(postSheet);
}

/**
 * A列（ID）が数字でない行を削除（チーム名・ティア名等の不要行）
 */
function cleanupPostApp_(sheet) {
  if (!sheet) {
    var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
    sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 下から上に削除（行番号がずれないように）
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  var deleted = 0;
  for (var i = ids.length - 1; i >= 0; i--) {
    var id = String(ids[i][0] || '').trim();
    if (!id || !/^\d+$/.test(id)) {
      sheet.deleteRow(i + 2);
      deleted++;
    }
  }
  if (deleted > 0) Logger.log('不要行を ' + deleted + '行 削除しました');
}

/**
 * 【4月】投稿数シートを契約日の古い順にソート（空欄は末尾）
 */
function sortPostAppByContract_(sheet) {
  if (!sheet) {
    var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
    sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // ヘッダー+1行以下ならソート不要

  var POST_APP_CONTRACT_COL = 2;
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  data.sort(function(a, b) {
    var da = String(a[POST_APP_CONTRACT_COL - 1] || '').trim();
    var db = String(b[POST_APP_CONTRACT_COL - 1] || '').trim();
    // 空欄は末尾
    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;
    // 日付比較（yyyy/MM/dd形式）
    var ta = new Date(da).getTime();
    var tb = new Date(db).getTime();
    if (isNaN(ta) && isNaN(tb)) return 0;
    if (isNaN(ta)) return 1;
    if (isNaN(tb)) return -1;
    return ta - tb;
  });

  sheet.getRange(2, 1, data.length, lastCol).setValues(data);
  Logger.log('契約日順にソートしました');
}

// ============================================
// 21時チャットワーク通知（未入力リマインダー）
// ============================================

var POST_APP_CW_ROOM_COL = 3; // C列: チャットワークグルチャID

/**
 * 毎日21時にトリガーで実行
 * 当日未入力の受講生にチャットワークで通知
 */
function postAppDailyReminder() {
  var cwToken = PropertiesService.getScriptProperties().getProperty('CHATWORK_POSTAPP_TOKEN');
  if (!cwToken) {
    Logger.log('CHATWORK_POSTAPP_TOKEN未設定');
    return;
  }

  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 今日の列を特定
  var now = new Date();
  var todayDay = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'dd'));
  var todayCol = POST_APP_DATE_START_COL + todayDay - 1;

  // 全データ取得（ID, 名前, 今日の値, CW room_id）
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  var names = sheet.getRange(2, POST_APP_NAME_COL, lastRow - 1, 1).getValues();
  var todayVals = sheet.getRange(2, todayCol, lastRow - 1, 1).getValues();
  var roomIds = sheet.getRange(2, POST_APP_CW_ROOM_COL, lastRow - 1, 1).getValues();

  var sentCount = 0;

  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i][0]).trim();
    var name = String(names[i][0]).trim();
    var val = String(todayVals[i][0]).trim();
    var roomId = String(roomIds[i][0]).trim();

    // IDなし or room_idなし → スキップ
    if (!id || !roomId) continue;

    // 入力済み（❌以外の値がある）→ スキップ
    if (val && val !== '' && val !== '❌') continue;

    // 連続日数を計算（メッセージに使う）
    var streak = 0;
    for (var j = todayDay - 2; j >= 0; j--) {
      var prevVal = String(sheet.getRange(i + 2, POST_APP_DATE_START_COL + j).getValue()).trim();
      if (prevVal && prevVal !== '' && prevVal !== '❌') streak++;
      else break;
    }

    // メッセージ生成
    var msg = postAppBuildReminderMsg_(name, streak);

    // チャットワーク送信
    postAppSendChatwork_(cwToken, roomId, msg);
    sentCount++;

    // API制限対策（100リクエスト/5分）
    if (sentCount % 10 === 0) Utilities.sleep(1000);
  }

  Logger.log('投稿リマインダー送信完了: ' + sentCount + '件');
}

/**
 * リマインダーメッセージ生成
 */
function postAppBuildReminderMsg_(name, streak) {
  var msg = '[info][title]🔥 投稿記録リマインダー 🔥[/title]';

  if (streak > 0) {
    // 連続記録あり → 途切れる危機感
    var msgs = [
      name + 'さん！今日の投稿記録がまだだよ！\n今 ' + streak + '日連続で記録中なのに、ここで止めたらもったいない！🔥',
      name + 'さん、今日まだ記録してないよ！\n' + streak + '日も連続で頑張ってきたのに途切れちゃう...！今日も記録しよう💪',
      name + 'さん！' + streak + '日連続の記録が途切れそう！😱\nあと少しで今日もクリアできるよ！'
    ];
    msg += msgs[Math.floor(Math.random() * msgs.length)];
  } else {
    // 連続なし → 軽い励まし
    var msgs2 = [
      name + 'さん！今日の投稿記録がまだ入ってないよ📝\n1日1回、記録するだけでOK！今日もやっていこう💪',
      name + 'さん、今日の記録まだだよ！\n今日から連続記録スタートしよう🔥 まずは3日連続を目指そう！',
      name + 'さん！投稿したら記録しよう📝\n記録をつけるだけで意識が変わるよ！'
    ];
    msg += msgs2[Math.floor(Math.random() * msgs2.length)];
  }

  msg += '\n\n▼ 今すぐ記録する\nhttps://giver.work/post-app/[/info]';
  return msg;
}

/**
 * チャットワークにメッセージ送信
 */
function postAppSendChatwork_(token, roomId, message) {
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

// ---- CW通知テスト（Namakaくんに送信） ----
function testCwNotification() {
  var cwToken = PropertiesService.getScriptProperties().getProperty('CHATWORK_POSTAPP_TOKEN');
  if (!cwToken) { Logger.log('CHATWORK_POSTAPP_TOKEN未設定'); return; }

  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME) || ss.getSheets()[0];

  // Namakaくん(ID=0000)の行を探す
  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === '0000') {
      var roomId = String(sheet.getRange(i + 2, POST_APP_CW_ROOM_COL).getValue()).trim();
      var name = String(sheet.getRange(i + 2, POST_APP_NAME_COL).getValue()).trim();
      if (!roomId) { Logger.log('C列にroom_idがありません'); return; }
      var msg = postAppBuildReminderMsg_(name || 'Namakaくん', 0);
      postAppSendChatwork_(cwToken, roomId, msg);
      Logger.log('テスト送信完了: room=' + roomId);
      return;
    }
  }
  Logger.log('Namakaくん(ID=0000)が見つかりません');
}

// ---- 21時リマインダーのトリガー設定 ----
function installPostAppReminderTrigger() {
  // 既存の同名トリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'postAppDailyReminder') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // 毎日21時に実行
  ScriptApp.newTrigger('postAppDailyReminder')
    .timeBased()
    .atHour(21)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();
  Logger.log('トリガー設定完了: postAppDailyReminder 毎日21時');
}

// ---- テストユーザーのデータリセット ----
function resetTestUsers() {
  var ids = ['5229', '5220'];
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var allIds = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();

  for (var i = 0; i < allIds.length; i++) {
    var sid = String(allIds[i][0]).trim();
    if (ids.indexOf(sid) !== -1) {
      var row = i + 2;
      // F列(6)〜AI列(35) = 30日分を❌でリセット（5月以降は31日対応）
      var blank = [];
      for (var d = 0; d < POST_APP_MONTH_DAYS; d++) blank.push('❌');
      sheet.getRange(row, POST_APP_DATE_START_COL, 1, POST_APP_MONTH_DAYS).setValues([blank]);
      Logger.log('Reset row ' + row + ' (ID: ' + sid + ')');
    }
  }

  // 認証シートからも削除
  var authSheet = getAuthSheet_();
  var authLast = authSheet.getLastRow();
  if (authLast >= 2) {
    var authData = authSheet.getRange(2, 1, authLast - 1, 1).getValues();
    for (var a = authData.length - 1; a >= 0; a--) {
      var aid = String(authData[a][0]).trim();
      if (ids.indexOf(aid) !== -1) {
        authSheet.deleteRow(a + 2);
        Logger.log('Deleted auth for ID: ' + aid);
      }
    }
  }

  Logger.log('Done: reset users ' + ids.join(', '));
}

// ---- テストアカウント「Namakaくん」作成（2行目に挿入） ----
function createTestAccount() {
  var id = '0000';
  var name = 'Namakaくん';
  var password = '0403';

  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();

  // 既存行があれば削除
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]).trim() === id) {
      sheet.deleteRow(i + 2);
    }
  }

  // 2行目に挿入
  sheet.insertRowAfter(1);
  // A=ID, B=契約日, C=CW ID, D=名前, E=合計数式, F〜AI=❌
  var f = '=COUNTIF(F2:AI2,"1本")+COUNTIF(F2:AI2,"2本")*2+COUNTIF(F2:AI2,"3本")*3+COUNTIF(F2:AI2,"4本")*4+COUNTIF(F2:AI2,"5本")*5+COUNTIF(F2:AI2,"6本")*6';
  var newRow = [id, '', '', name, f];
  for (var d = 0; d < POST_APP_MONTH_DAYS; d++) newRow.push('❌');
  sheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
  Logger.log('Inserted ' + name + ' at row 2');

  // パスワード登録
  var authSheet = getAuthSheet_();
  var authLast = authSheet.getLastRow();
  if (authLast >= 2) {
    var authData = authSheet.getRange(2, 1, authLast - 1, 1).getValues();
    for (var a = authData.length - 1; a >= 0; a--) {
      if (String(authData[a][0]).trim() === id) authSheet.deleteRow(a + 2);
    }
  }
  authSheet.appendRow([id, hashPassword_(password)]);
  Logger.log('Password set for ' + name + ' (ID: ' + id + ')');
}

// ---- ランキング取得（当月のみ） ----
function postAppRanking_() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ranking: [] };

  var range = getCurrentMonthColRange_();
  // メタ列(A〜E) + 当月の日別列だけ取得
  var metaData = sheet.getRange(2, 1, lastRow - 1, POST_APP_NAME_COL).getValues();
  var monthData = sheet.getRange(2, range.startCol, lastRow - 1, range.monthDays).getValues();
  var now = new Date();
  var todayDay = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'dd'));

  var members = [];
  for (var i = 0; i < metaData.length; i++) {
    var id = String(metaData[i][POST_APP_ID_COL - 1] || '').trim();
    var name = String(metaData[i][POST_APP_NAME_COL - 1] || '').trim();
    if (!id || !name) continue;

    // 投稿本数合計（当月分）
    var total = 0;
    var streak = 0;
    var postDays = 0;
    for (var d = 0; d < range.monthDays; d++) {
      var v = String(monthData[i][d] || '').trim();
      if (v && v !== '❌') {
        var num = parseInt(v) || 0;
        total += num;
        postDays++;
      }
    }
    // 連続日数（今日の前日から遡る）
    for (var j = todayDay - 1; j >= 0; j--) {
      var v2 = String(monthData[i][j] || '').trim();
      if (v2 && v2 !== '❌') streak++;
      else break;
    }

    if (total > 0 || postDays > 0) {
      members.push({ id: id, name: name, total: total, streak: streak, postDays: postDays });
    }
  }

  members.sort(function(a, b) { return b.total - a.total || b.streak - a.streak; });
  return { ranking: members };
}

// ---- ホープ数（リスト数）取得API ----
// 権限判定: 「投稿数」シートD列背景色 ≠ 白 → allowed:true
// データ取得: 「ホープ数」シートから指定月の列範囲を読む
// 合計: days[] の sum を合算（C列の数式に依存しない、月跨ぎでもズレない）
function postAppGetHope_(token, year, month) {
  var id = verifyToken_(token);
  if (!id) return { error: 'セッション切れです。再ログインしてください。' };

  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);

  // ① 投稿数シートD列背景色で権限判定
  var postSheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!postSheet) return { allowed: false };
  var postLastRow = postSheet.getLastRow();
  if (postLastRow < 2) return { allowed: false };

  var postIds = postSheet.getRange(2, POST_APP_ID_COL, postLastRow - 1, 1).getValues();
  var postNameBgs = postSheet.getRange(2, POST_APP_NAME_COL, postLastRow - 1, 1).getBackgrounds();
  var myIdx = -1;
  for (var i = 0; i < postIds.length; i++) {
    if (String(postIds[i][0]).trim() === String(id).trim()) { myIdx = i; break; }
  }
  if (myIdx < 0) return { allowed: false };

  var bg = String(postNameBgs[myIdx][0] || '').toLowerCase();
  if (!bg || bg === '#ffffff' || bg === 'white') return { allowed: false };

  // 指定月のホープ列範囲を計算
  var hRange = getCurrentMonthHopeColRange_(POST_APP_HOPE_SHEET_NAME,
    year ? parseInt(year) : null,
    month ? parseInt(month) : null);

  // ② ホープ数シートから指定月データ取得
  var hopeSheet = ss.getSheetByName(POST_APP_HOPE_SHEET_NAME);
  if (!hopeSheet) {
    return { allowed: true, year: hRange.year, month: hRange.month, total: 0, days: [] };
  }
  var hopeLastRow = hopeSheet.getLastRow();
  if (hopeLastRow < 3) {
    return { allowed: true, year: hRange.year, month: hRange.month, total: 0, days: [] };
  }

  var hopeIds = hopeSheet.getRange(3, 1, hopeLastRow - 2, 1).getValues();
  var hopeRowIdx = -1;
  for (var j = 0; j < hopeIds.length; j++) {
    if (String(hopeIds[j][0]).trim() === String(id).trim()) { hopeRowIdx = j; break; }
  }
  if (hopeRowIdx < 0) {
    return { allowed: true, year: hRange.year, month: hRange.month, total: 0, days: [] };
  }

  var row = hopeRowIdx + 3;
  var values = hopeSheet.getRange(row, hRange.startCol, 1, hRange.totalCols).getValues()[0];

  var days = [];
  var total = 0;
  for (var d = 0; d < hRange.monthDays; d++) {
    var yt = parseInt(values[d * 3] || 0) || 0;
    var ig = parseInt(values[d * 3 + 1] || 0) || 0;
    var tt = parseInt(values[d * 3 + 2] || 0) || 0;
    var sum = yt + ig + tt;
    days.push({ day: d + 1, yt: yt, ig: ig, tt: tt, sum: sum });
    total += sum;
  }

  return { allowed: true, year: hRange.year, month: hRange.month, total: total, days: days };
}

// ---- D列の名前をLPスプシから復元 ----
function restoreNames() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // post-appスプシのID列を読む
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();

  // LPスプシ(gid=655724533)からID→名前マップ作成
  var lpSs = SpreadsheetApp.openById(LP_SS_ID);
  var lpSheets = lpSs.getSheets();
  var lpSheet = null;
  for (var s = 0; s < lpSheets.length; s++) {
    if (lpSheets[s].getSheetId() === LP_POST_SHEET_GID) { lpSheet = lpSheets[s]; break; }
  }
  if (!lpSheet) { Logger.log('LP sheet not found'); return; }

  var lpData = lpSheet.getDataRange().getValues();
  var nameMap = {};
  for (var r = 1; r < lpData.length; r++) {
    var lid = String(lpData[r][LP_POST_ID_COL - 1] || '').trim();
    var lname = String(lpData[r][LP_POST_NAME_COL - 1] || '').trim();
    if (lid && lname) nameMap[lid] = lname;
  }

  // D列に名前を書き戻す
  var names = [];
  var restored = 0;
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i][0]).trim();
    var name = nameMap[id] || '';
    if (id === '0000') name = 'Namakaくん';
    names.push([name]);
    if (name) restored++;
  }
  sheet.getRange(2, POST_APP_NAME_COL, names.length, 1).setValues(names);
  Logger.log('Restored ' + restored + ' names out of ' + names.length + ' rows');
}

// ---- E列(合計)に数式を一括設定 ----
function fixPostAppTotal() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var formulas = [];
  for (var r = 2; r <= lastRow; r++) {
    formulas.push(['=COUNTIF(F' + r + ':AI' + r + ',"1本")+COUNTIF(F' + r + ':AI' + r + ',"2本")*2+COUNTIF(F' + r + ':AI' + r + ',"3本")*3+COUNTIF(F' + r + ':AI' + r + ',"4本")*4+COUNTIF(F' + r + ':AI' + r + ',"5本")*5+COUNTIF(F' + r + ':AI' + r + ',"6本")*6']);
  }
  sheet.getRange(2, POST_APP_TOTAL_COL, formulas.length, 1).setFormulas(formulas);
  Logger.log('Done: ' + formulas.length + ' rows updated');
}

/**
 * 当月シートの存在確認（GASエディタから手動実行用）
 * 月跨ぎデプロイ後に「【○月】投稿数」「【○月】ホープ数」「【○月】プッシュ数」が
 * スプシに作られているか検証する。
 */
function checkCurrentMonthSheets_() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var targets = [POST_APP_SHEET_NAME, POST_APP_HOPE_SHEET_NAME, POST_APP_PUSH_SHEET_NAME];
  for (var i = 0; i < targets.length; i++) {
    var sheet = ss.getSheetByName(targets[i]);
    if (sheet) {
      Logger.log(targets[i] + ': ✅ 存在 (lastRow=' + sheet.getLastRow() + ', lastCol=' + sheet.getLastColumn() + ')');
    } else {
      Logger.log(targets[i] + ': ❌ 無し');
    }
  }
  var range = getCurrentMonthColRange_();
  var hRange = getCurrentMonthHopeColRange_();
  Logger.log('当月: ' + (new Date().getMonth() + 1) + '月 / 投稿数列範囲 ' + postAppColToLetter_(range.startCol) + ':' + postAppColToLetter_(range.endCol) + ' (' + range.monthDays + '日)');
  Logger.log('当月ホープ列範囲 ' + postAppColToLetter_(hRange.startCol) + ':' + postAppColToLetter_(hRange.endCol) + ' (' + hRange.totalCols + '列)');
}

// ===========================================
// 統合シート方式 - 移行・月初列追加・トリガー
// ===========================================

/**
 * 既存の月別シート方式から統合シート方式へ移行する
 * 一回限り、GASエディタから手動実行
 *
 * 前提: スプシ全体のバックアップを「ファイル→コピーを作成」で取得済み
 */
function migrateToUnifiedSheet() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
  Logger.log('=== 統合シート移行開始: ' + ts + ' ===');

  // Step 1: 「【4月】投稿数」 → 「投稿数」
  var aprPost = ss.getSheetByName('【4月】投稿数');
  if (!aprPost) throw new Error('【4月】投稿数シートが見つかりません');
  var existPost = ss.getSheetByName('投稿数');
  if (existPost) {
    existPost.setName('アーカイブ_投稿数_' + ts);
    Logger.log('既存「投稿数」 → アーカイブ_投稿数_' + ts);
  }
  aprPost.setName('投稿数');
  Logger.log('「【4月】投稿数」 → 「投稿数」');

  // Step 2: 「【新】📲 SMC投稿&リスト進捗」 → アーカイブ
  var smcSheet = ss.getSheetByName('【新】📲 SMC投稿&リスト進捗');
  if (smcSheet) {
    smcSheet.setName('アーカイブ_SMC旧_' + ts);
    Logger.log('「【新】📲 SMC投稿&リスト進捗」 → アーカイブ_SMC旧_' + ts);
  }

  // Step 3: 「【4月】ホープ数」 → 「ホープ数」
  var aprHope = ss.getSheetByName('【4月】ホープ数');
  if (aprHope) {
    var existHope = ss.getSheetByName('ホープ数');
    if (existHope) existHope.setName('アーカイブ_ホープ数_' + ts);
    aprHope.setName('ホープ数');
    Logger.log('「【4月】ホープ数」 → 「ホープ数」');
  } else {
    Logger.log('注意: 【4月】ホープ数 シートが無い → リスト数機能はホープ数シート作成後に有効');
  }

  // Step 4: 「【4月】プッシュ数」 → 「プッシュ数」
  var aprPush = ss.getSheetByName('【4月】プッシュ数');
  if (aprPush) {
    var existPush = ss.getSheetByName('プッシュ数');
    if (existPush) existPush.setName('アーカイブ_プッシュ数_' + ts);
    aprPush.setName('プッシュ数');
    Logger.log('「【4月】プッシュ数」 → 「プッシュ数」');
  } else {
    Logger.log('注意: 【4月】プッシュ数 シートが無い');
  }

  // Step 5: 当月の列を追加 + 数式更新
  ensureCurrentMonthColumns_(ss);

  Logger.log('=== 移行完了 ===');
  Logger.log('アーカイブシートは消さずに残しています。スプシで目視確認してください。');
}

/**
 * 投稿数・ホープ数・プッシュ数の当月の列が無ければ追加する
 * トリガーから自動実行 + 移行関数からも呼ばれる
 */
function ensureCurrentMonthColumns_(ss) {
  ss = ss || SpreadsheetApp.openById(POST_APP_SS_ID);
  var month = new Date().getMonth() + 1;

  // 投稿数シート
  var postSheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (postSheet) {
    var range = getCurrentMonthColRange_();
    var lastCol = postSheet.getLastColumn();
    if (lastCol < range.endCol) {
      var addCols = range.endCol - lastCol;
      postSheet.insertColumnsAfter(lastCol, addCols);
      var dates = [];
      for (var d = 1; d <= range.monthDays; d++) dates.push(month + '/' + d);
      postSheet.getRange(1, range.startCol, 1, range.monthDays).setValues([dates]);
      var lastRow = postSheet.getLastRow();
      if (lastRow >= 2) {
        var blank = [];
        for (var r = 0; r < lastRow - 1; r++) {
          var row = [];
          for (var d2 = 0; d2 < range.monthDays; d2++) row.push('❌');
          blank.push(row);
        }
        postSheet.getRange(2, range.startCol, lastRow - 1, range.monthDays).setValues(blank);
      }
      Logger.log('投稿数: ' + addCols + '列追加 (' + month + '月分)');
    }
    updatePostSheetTotalFormula_(postSheet);
  }

  // ホープ数・プッシュ数（2行構造ヘッダー: 日付マージ + YT/IG/TT装飾）
  // シート種別でデータ起点列が違うので getCurrentMonthHopeColRange_(name) で取得
  [POST_APP_HOPE_SHEET_NAME, POST_APP_PUSH_SHEET_NAME].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log('警告: ' + name + ' シートが存在しません');
      return;
    }
    var hRange = getCurrentMonthHopeColRange_(name); // シート名指定で正しい起点列
    var lastCol = sheet.getLastColumn();
    if (lastCol < hRange.endCol) {
      var addCols = hRange.endCol - lastCol;
      sheet.insertColumnsAfter(lastCol, addCols);
      // 2行構造のヘッダー
      writeHope2RowHeader_(sheet, hRange, month);
      // データ行は3行目以降を0で初期化
      var lastRow = sheet.getLastRow();
      if (lastRow >= 3) {
        var zeros = [];
        for (var r = 0; r < lastRow - 2; r++) {
          var row = [];
          for (var c = 0; c < hRange.totalCols; c++) row.push(0);
          zeros.push(row);
        }
        sheet.getRange(3, hRange.startCol, lastRow - 2, hRange.totalCols).setValues(zeros);
      }
      Logger.log(name + ': ' + addCols + '列追加 (' + month + '月分、起点=' + postAppColToLetter_(hRange.startCol) + '列)');
    } else {
      Logger.log(name + ': 当月分は既存。スキップ');
    }
  });
}

/**
 * ホープ数・プッシュ数シートの当月ヘッダーを2行構造で書く
 * 1行目: 日付（3列マージ）"5/1" "5/2" ...
 * 2行目: YT(赤) / IG(紫) / TT(水色) を3列ずつ繰り返し
 */
function writeHope2RowHeader_(sheet, hRange, month) {
  // マージ解除（既にあれば）
  try {
    sheet.getRange(1, hRange.startCol, 1, hRange.totalCols).breakApart();
  } catch (e) {}

  // 1行目: 日付（薄い緑背景・太字・中央揃え。既存4月分と統一）
  var dateRow = [];
  for (var d = 1; d <= hRange.monthDays; d++) {
    dateRow.push(month + '/' + d, '', '');
  }
  sheet.getRange(1, hRange.startCol, 1, hRange.totalCols)
    .setValues([dateRow])
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground('#F1F8E9');
  // 3列ずつマージ
  for (var d2 = 0; d2 < hRange.monthDays; d2++) {
    var col = hRange.startCol + d2 * hRange.colsPerDay;
    sheet.getRange(1, col, 1, hRange.colsPerDay).merge();
  }

  // 2行目: YT/IG/TT
  var categoryRow = [];
  for (var d3 = 1; d3 <= hRange.monthDays; d3++) {
    categoryRow.push('YT', 'IG', 'TT');
  }
  sheet.getRange(2, hRange.startCol, 1, hRange.totalCols)
    .setValues([categoryRow])
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setFontColor('#FFFFFF');

  // 背景色: YT=赤 / IG=紫 / TT=水色
  for (var d4 = 0; d4 < hRange.monthDays; d4++) {
    var baseCol = hRange.startCol + d4 * hRange.colsPerDay;
    sheet.getRange(2, baseCol).setBackground('#E53935');     // YT 赤
    sheet.getRange(2, baseCol + 1).setBackground('#9C27B0'); // IG 紫
    sheet.getRange(2, baseCol + 2).setBackground('#26C6DA'); // TT 水色
  }
}

/**
 * ⚠️⚠️⚠️ 廃止: この関数はプッシュ数を破壊する可能性があります ⚠️⚠️⚠️
 * プッシュ数のデータ起点はL列(12)、ホープ数はD列(4)。シート別に分岐が必要。
 * 当月分の追加は ensureCurrentMonthColumnsTrigger または addPushSheetCurrentMonth を使ってください。
 */
function fixUnifiedSheetHeaders() {
  Logger.log('⚠️⚠️⚠️ 廃止された関数です ⚠️⚠️⚠️');
  Logger.log('代わりに以下を実行:');
  Logger.log('  - 全シート当月分追加: ensureCurrentMonthColumnsTrigger');
  Logger.log('  - プッシュ数のみ:     addPushSheetCurrentMonth');
}

/**
 * プッシュ数シートに当月分の列を最小破壊で追加
 * - 既存4月分（L〜CW列）は絶対触らない
 * - 5月分（CX列〜）に93列追加 + ヘッダー（日付マージ + YT/IG/TT装飾）+ 0初期化
 * GASエディタから手動実行
 */
function addPushSheetCurrentMonth() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_PUSH_SHEET_NAME);
  if (!sheet) {
    Logger.log('❌ プッシュ数シートが見つかりません');
    return;
  }

  var hRange = getCurrentMonthHopeColRange_(POST_APP_PUSH_SHEET_NAME);
  var month = new Date().getMonth() + 1;
  var lastCol = sheet.getLastColumn();

  Logger.log('=== プッシュ数 ' + month + '月分追加（事前チェック）===');
  Logger.log('lastRow: ' + sheet.getLastRow());
  Logger.log('lastCol: ' + lastCol + ' (' + postAppColToLetter_(lastCol) + ')');
  Logger.log('期待される' + month + '月開始列: ' + postAppColToLetter_(hRange.startCol) + ' (' + hRange.startCol + ')');
  Logger.log('期待される' + month + '月終端列: ' + postAppColToLetter_(hRange.endCol) + ' (' + hRange.endCol + ')');

  if (lastCol >= hRange.endCol) {
    Logger.log('✓ ' + month + '月分は既に追加済み。スキップ。');
    return;
  }

  // lastCol < startCol-1 の場合は既存月分が足りない疑い → 安全のため手動確認に回す
  if (lastCol < hRange.startCol - 1) {
    Logger.log('⚠️ lastCol(' + lastCol + ')が期待される前月終端(' + (hRange.startCol - 1) + ')未満。');
    Logger.log('既存月の列が想定通りないため、安全のため処理を中止します。');
    Logger.log('スプシで構造確認 → 必要なら手動で前月までのデータ列を補ってから再実行してください。');
    return;
  }

  var addCols = hRange.endCol - lastCol;
  sheet.insertColumnsAfter(lastCol, addCols);
  Logger.log('insertColumnsAfter: ' + addCols + '列追加');

  writeHope2RowHeader_(sheet, hRange, month);

  // データ行は3行目以降を0で初期化
  var lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    var zeros = [];
    for (var r = 0; r < lastRow - 2; r++) {
      var row = [];
      for (var c = 0; c < hRange.totalCols; c++) row.push(0);
      zeros.push(row);
    }
    sheet.getRange(3, hRange.startCol, lastRow - 2, hRange.totalCols).setValues(zeros);
  }

  Logger.log('✅ プッシュ数 ' + month + '月分追加完了 (' + addCols + '列、起点=' + postAppColToLetter_(hRange.startCol) + ')');
}

/**
 * 指定シートの全月ヘッダー（4月〜当月）を完全リビルドする
 * プッシュ数の merge エラー対策。ヘッダー1〜2行目を完全クリアして書き直す。
 * データ行（3行目以降）は触らない。
 * GASエディタから手動実行。
 */
/**
 * ⚠️⚠️⚠️ 廃止: 2026-05-02 にプッシュ数の4月分を破壊した事故関数 ⚠️⚠️⚠️
 * 当月追加は addPushSheetCurrentMonth を使ってください。
 */
function rebuildPushSheetHeaders() {
  Logger.log('⚠️⚠️⚠️ 廃止された関数です ⚠️⚠️⚠️');
  Logger.log('プッシュ数の起点列(L=12)を考慮してないため、誤った位置に書いてシートを破壊します。');
  Logger.log('代わりに addPushSheetCurrentMonth を実行してください。');
}

/**
 * ⚠️⚠️⚠️ 廃止: 全月リビルドは破壊リスク高 ⚠️⚠️⚠️
 * 当月追加は ensureCurrentMonthColumnsTrigger を使ってください。
 */
function rebuildHopeSheetHeaders() {
  Logger.log('⚠️⚠️⚠️ 廃止された関数です ⚠️⚠️⚠️');
  Logger.log('全月リビルドは既存マージ・装飾を破壊するリスクが高いです。');
  Logger.log('当月追加なら ensureCurrentMonthColumnsTrigger を実行してください。');
}

function rebuildHopeSheetHeadersForName_(name) {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    Logger.log(name + ' シートが見つかりません');
    return;
  }
  var now = new Date();
  var thisYear = now.getFullYear();
  var thisMonth = now.getMonth() + 1;

  // 1〜2行目のD列以降を完全クリア（マージ含む、メタ列A〜Cは保持）
  var maxCol = sheet.getMaxColumns();
  if (maxCol > 3) {
    var rng = sheet.getRange(1, 4, 2, maxCol - 3);
    try { rng.breakApart(); } catch (e) { Logger.log('breakApart skipped: ' + e.message); }
    rng.clearContent();
    rng.clearFormat();
  }

  // 4月〜当月までのヘッダーを順番に書く
  var col = 4;
  for (var m = POST_APP_RUN_START_MONTH; m <= thisMonth; m++) {
    var monthDays = new Date(thisYear, m, 0).getDate();
    writeMonthHopeHeader_(sheet, col, monthDays, m);
    col += monthDays * 3;
  }

  Logger.log(name + ': 全月ヘッダーを再構築 (' + POST_APP_RUN_START_MONTH + '月〜' + thisMonth + '月、' + (col - 4) + '列)');
}

function writeMonthHopeHeader_(sheet, startCol, monthDays, month) {
  var totalCols = monthDays * 3;
  // 1行目: 日付
  var dateRow = [];
  for (var d = 1; d <= monthDays; d++) {
    dateRow.push(month + '/' + d, '', '');
  }
  sheet.getRange(1, startCol, 1, totalCols)
    .setValues([dateRow])
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground('#F1F8E9');
  for (var d2 = 0; d2 < monthDays; d2++) {
    sheet.getRange(1, startCol + d2 * 3, 1, 3).merge();
  }
  // 2行目: YT/IG/TT
  var categoryRow = [];
  for (var d3 = 1; d3 <= monthDays; d3++) {
    categoryRow.push('YT', 'IG', 'TT');
  }
  sheet.getRange(2, startCol, 1, totalCols)
    .setValues([categoryRow])
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setFontColor('#FFFFFF');
  for (var d4 = 0; d4 < monthDays; d4++) {
    var baseCol = startCol + d4 * 3;
    sheet.getRange(2, baseCol).setBackground('#E53935');     // YT 赤
    sheet.getRange(2, baseCol + 1).setBackground('#9C27B0'); // IG 紫
    sheet.getRange(2, baseCol + 2).setBackground('#26C6DA'); // TT 水色
  }
}

/**
 * 投稿数シートE列の合計数式を当月の列範囲ベースで更新
 */
function updatePostSheetTotalFormula_(sheet) {
  var range = getCurrentMonthColRange_();
  var startLetter = postAppColToLetter_(range.startCol);
  var endLetter = postAppColToLetter_(range.endCol);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var formulas = [];
  for (var r = 2; r <= lastRow; r++) {
    var f = '=COUNTIF(' + startLetter + r + ':' + endLetter + r + ',"1本")' +
            '+COUNTIF(' + startLetter + r + ':' + endLetter + r + ',"2本")*2' +
            '+COUNTIF(' + startLetter + r + ':' + endLetter + r + ',"3本")*3' +
            '+COUNTIF(' + startLetter + r + ':' + endLetter + r + ',"4本")*4' +
            '+COUNTIF(' + startLetter + r + ':' + endLetter + r + ',"5本")*5' +
            '+COUNTIF(' + startLetter + r + ':' + endLetter + r + ',"6本")*6' +
            '+COUNTIF(' + startLetter + r + ':' + endLetter + r + ',"7本")*7' +
            '+COUNTIF(' + startLetter + r + ':' + endLetter + r + ',"8本")*8' +
            '+COUNTIF(' + startLetter + r + ':' + endLetter + r + ',"9本")*9' +
            '+COUNTIF(' + startLetter + r + ':' + endLetter + r + ',"10本")*10';
    formulas.push([f]);
  }
  sheet.getRange(2, POST_APP_TOTAL_COL, formulas.length, 1).setFormulas(formulas);
  Logger.log('投稿数 E列の合計数式を更新: ' + formulas.length + '行 (' + startLetter + ':' + endLetter + ')');
}

/**
 * 月初列追加トリガーをインストール（毎日0時起動）
 * GASエディタから1度だけ手動実行
 */
function installMonthlyAutoTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var deleted = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'ensureCurrentMonthColumnsTrigger') {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted++;
    }
  }
  ScriptApp.newTrigger('ensureCurrentMonthColumnsTrigger')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();
  Logger.log('月初列追加トリガーをインストール (毎日0:00, 旧トリガー' + deleted + '個削除)');
}

/** トリガーから呼ばれる本体 */
function ensureCurrentMonthColumnsTrigger() {
  ensureCurrentMonthColumns_();
}

// ===========================================
// 年内の全月一括セットアップ（年初 or 緊急時）
// ===========================================

/**
 * 投稿数シートに指定月の列を右側に追加する
 */
function addPostSheetMonth_(ss, year, month) {
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) {
    Logger.log('投稿数シートが見つかりません');
    return;
  }
  var range = getMonthColRange_(year, month);
  var lastCol = sheet.getLastColumn();
  if (lastCol >= range.endCol) {
    Logger.log('投稿数 ' + year + '/' + month + ': 既存（' + postAppColToLetter_(range.startCol) + '〜' + postAppColToLetter_(range.endCol) + '）→ スキップ');
    return;
  }
  if (lastCol < range.startCol - 1) {
    Logger.log('⚠️ 投稿数 ' + year + '/' + month + ': lastCol(' + lastCol + ')が期待される前月終端(' + (range.startCol - 1) + ')未満。前月までのデータが不足、安全のため中止。');
    return;
  }
  var addCols = range.endCol - lastCol;
  sheet.insertColumnsAfter(lastCol, addCols);
  var dates = [];
  for (var d = 1; d <= range.monthDays; d++) dates.push(month + '/' + d);
  sheet.getRange(1, range.startCol, 1, range.monthDays).setValues([dates]);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var blank = [];
    for (var r = 0; r < lastRow - 1; r++) {
      var row = [];
      for (var d2 = 0; d2 < range.monthDays; d2++) row.push('❌');
      blank.push(row);
    }
    sheet.getRange(2, range.startCol, lastRow - 1, range.monthDays).setValues(blank);
  }
  Logger.log('✅ 投稿数 ' + year + '/' + month + ': ' + addCols + '列追加 (' + postAppColToLetter_(range.startCol) + '〜' + postAppColToLetter_(range.endCol) + ')');
}

/**
 * ホープ数/プッシュ数シートに指定月の列を右側に追加する（2行構造ヘッダー）
 */
function addHopeSheetMonth_(ss, year, month, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(sheetName + ' シートが見つかりません');
    return;
  }
  var hRange = getCurrentMonthHopeColRange_(sheetName, year, month);
  var lastCol = sheet.getLastColumn();
  if (lastCol >= hRange.endCol) {
    Logger.log(sheetName + ' ' + year + '/' + month + ': 既存 → スキップ');
    return;
  }
  if (lastCol < hRange.startCol - 1) {
    Logger.log('⚠️ ' + sheetName + ' ' + year + '/' + month + ': lastCol(' + lastCol + ')が期待される前月終端(' + (hRange.startCol - 1) + ')未満。前月までのデータが不足、安全のため中止。');
    return;
  }
  var addCols = hRange.endCol - lastCol;
  sheet.insertColumnsAfter(lastCol, addCols);
  writeHope2RowHeader_(sheet, hRange, month);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    var zeros = [];
    for (var r = 0; r < lastRow - 2; r++) {
      var row = [];
      for (var c = 0; c < hRange.totalCols; c++) row.push(0);
      zeros.push(row);
    }
    sheet.getRange(3, hRange.startCol, lastRow - 2, hRange.totalCols).setValues(zeros);
  }
  Logger.log('✅ ' + sheetName + ' ' + year + '/' + month + ': ' + addCols + '列追加');
}

/**
 * プッシュ数シートのCX列以降の壊れた範囲を削除する（5/2事故の残骸クリーンアップ）
 * 4月分（L〜CW）は保持。CX列以降を全削除→再構築用にクリーンな状態にする。
 */
function cleanupPushSheetBrokenColumns_(ss) {
  var sheet = ss.getSheetByName(POST_APP_PUSH_SHEET_NAME);
  if (!sheet) {
    Logger.log('プッシュ数シートが見つかりません');
    return;
  }
  // プッシュ数の4月分終端 = L(12) + 30*3 - 1 = 101列目（CW列）
  var aprilEndCol = 12 + 30 * 3 - 1;
  var lastCol = sheet.getLastColumn();
  Logger.log('プッシュ数: lastCol=' + lastCol + ' (' + postAppColToLetter_(lastCol) + '), 4月終端=' + aprilEndCol + ' (' + postAppColToLetter_(aprilEndCol) + ')');
  if (lastCol > aprilEndCol) {
    var deleteCols = lastCol - aprilEndCol;
    Logger.log('プッシュ数: ' + postAppColToLetter_(aprilEndCol + 1) + '列以降の' + deleteCols + '列を削除');
    sheet.deleteColumns(aprilEndCol + 1, deleteCols);
    Logger.log('✅ プッシュ数クリーンアップ完了、新lastCol=' + sheet.getLastColumn() + ' (' + postAppColToLetter_(sheet.getLastColumn()) + ')');
  } else {
    Logger.log('プッシュ数: クリーンアップ不要（lastCol <= 4月終端）');
  }
}

/**
 * 年内の全月（5月〜12月）を3シートに一括追加する
 * 5/2事故対応 + 月跨ぎ対策の年初セットアップ用。
 * GASエディタから手動実行。
 *
 * 実行内容:
 *   Step 1: プッシュ数のCX列以降の壊れた残骸クリーンアップ
 *   Step 2: 5月分を3シートに正しい位置で追加
 *   Step 3: 6月〜12月分を3シートに追加
 */
function setupSheetsThroughDecember() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var year = new Date().getFullYear();

  Logger.log('==========================================');
  Logger.log('=== 年内全月セットアップ開始: ' + year + ' ===');
  Logger.log('==========================================');

  // Step 1: プッシュ数のクリーンアップ
  Logger.log('--- Step 1: プッシュ数の壊れた残骸クリーンアップ ---');
  cleanupPushSheetBrokenColumns_(ss);

  // Step 2: 5月分を3シートに追加（5月分が未追加 or プッシュ数で削除された分の補充）
  Logger.log('--- Step 2: 5月分を3シートに追加 ---');
  addPostSheetMonth_(ss, year, 5);
  addHopeSheetMonth_(ss, year, 5, POST_APP_HOPE_SHEET_NAME);
  addHopeSheetMonth_(ss, year, 5, POST_APP_PUSH_SHEET_NAME);

  // Step 3: 6月〜12月を3シートに追加
  Logger.log('--- Step 3: 6月〜12月を3シートに追加 ---');
  for (var m = 6; m <= 12; m++) {
    Logger.log('-- ' + m + '月 --');
    addPostSheetMonth_(ss, year, m);
    addHopeSheetMonth_(ss, year, m, POST_APP_HOPE_SHEET_NAME);
    addHopeSheetMonth_(ss, year, m, POST_APP_PUSH_SHEET_NAME);
  }

  Logger.log('==========================================');
  Logger.log('=== 完了: 全シートが12月分まで揃った ===');
  Logger.log('==========================================');
  Logger.log('スプシで目視確認推奨');
}

/**
 * ホープ数シート読み取り専用デバッグ（破壊なし）
 * GASエディタから手動実行 → ログを俺(Claude)に貼ってもらう
 *
 * 確認内容:
 * 1. ホープ数シートのlastCol・列範囲計算結果
 * 2. ヘッダー行（1行目）の4月終端〜5月開始付近の値
 * 3. 任意ID(ナマカくん想定: 0000)の4月・5月の生データ + 合計計算結果
 */
function debugHopeSheet() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_HOPE_SHEET_NAME);
  if (!sheet) {
    Logger.log('❌ ホープ数シートが見つからない');
    return;
  }

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  Logger.log('=== ホープ数シート全体 ===');
  Logger.log('lastRow: ' + lastRow + ', lastCol: ' + lastCol + ' (' + postAppColToLetter_(lastCol) + ')');

  // 4月分・5月分の列範囲計算
  var aprRange = getCurrentMonthHopeColRange_(POST_APP_HOPE_SHEET_NAME, 2026, 4);
  var mayRange = getCurrentMonthHopeColRange_(POST_APP_HOPE_SHEET_NAME, 2026, 5);
  Logger.log('--- 列範囲計算結果 ---');
  Logger.log('4月: startCol=' + aprRange.startCol + ' (' + postAppColToLetter_(aprRange.startCol) + '), endCol=' + aprRange.endCol + ' (' + postAppColToLetter_(aprRange.endCol) + '), monthDays=' + aprRange.monthDays);
  Logger.log('5月: startCol=' + mayRange.startCol + ' (' + postAppColToLetter_(mayRange.startCol) + '), endCol=' + mayRange.endCol + ' (' + postAppColToLetter_(mayRange.endCol) + '), monthDays=' + mayRange.monthDays);

  // ヘッダー行の境界付近確認
  Logger.log('--- ヘッダー1行目の境界（4月最終3列+5月最初3列）---');
  var checkStart = aprRange.endCol - 2;
  var checkLen = 6;
  if (checkStart + checkLen - 1 <= lastCol) {
    var headerVals = sheet.getRange(1, checkStart, 1, checkLen).getValues()[0];
    var header2Vals = sheet.getRange(2, checkStart, 1, checkLen).getValues()[0];
    for (var i = 0; i < checkLen; i++) {
      var col = checkStart + i;
      Logger.log('  ' + postAppColToLetter_(col) + ' (col=' + col + ') 1行目="' + headerVals[i] + '" 2行目="' + header2Vals[i] + '"');
    }
  }

  // 全メンバー列挙（先頭5人）
  Logger.log('--- メンバー一覧（先頭10人）---');
  if (lastRow >= 3) {
    var ids = sheet.getRange(3, 1, Math.min(10, lastRow - 2), 3).getValues();
    for (var k = 0; k < ids.length; k++) {
      Logger.log('  row=' + (k + 3) + ' A=' + ids[k][0] + ' B=' + ids[k][1] + ' C(合計)=' + ids[k][2]);
    }
  }

  // 1人だけ詳細チェック（先頭メンバー）
  if (lastRow >= 3) {
    var firstId = String(sheet.getRange(3, 1).getValue()).trim();
    var firstName = sheet.getRange(3, 2).getValue();
    Logger.log('--- 詳細チェック: ' + firstName + ' (ID=' + firstId + ', row=3) ---');

    var aprValues = sheet.getRange(3, aprRange.startCol, 1, aprRange.totalCols).getValues()[0];
    var mayValues = (mayRange.endCol <= lastCol) ? sheet.getRange(3, mayRange.startCol, 1, mayRange.totalCols).getValues()[0] : null;

    var aprTotal = 0, aprSumDays = 0;
    for (var d = 0; d < aprRange.monthDays; d++) {
      var sum = (parseInt(aprValues[d * 3] || 0) || 0) + (parseInt(aprValues[d * 3 + 1] || 0) || 0) + (parseInt(aprValues[d * 3 + 2] || 0) || 0);
      aprTotal += sum;
      if (sum > 0) aprSumDays++;
    }
    Logger.log('  4月合計: ' + aprTotal + ' (データある日数: ' + aprSumDays + ')');

    if (mayValues) {
      var mayTotal = 0, maySumDays = 0;
      for (var d2 = 0; d2 < mayRange.monthDays; d2++) {
        var sum2 = (parseInt(mayValues[d2 * 3] || 0) || 0) + (parseInt(mayValues[d2 * 3 + 1] || 0) || 0) + (parseInt(mayValues[d2 * 3 + 2] || 0) || 0);
        mayTotal += sum2;
        if (sum2 > 0) maySumDays++;
      }
      Logger.log('  5月合計: ' + mayTotal + ' (データある日数: ' + maySumDays + ')');
    } else {
      Logger.log('  5月: 列範囲がlastColを超えてるためデータなし');
    }

    Logger.log('  C列(合計セル)の値: ' + sheet.getRange(3, 3).getValue());
  }

  Logger.log('=== デバッグ完了 ===');
}

// ===========================================
// ホープ数: 月別合計列を C列の右に追加するマイグレーション
// 旧構造: A=ID, B=名前, C=合計, D〜=データ
// 新構造: A=ID, B=名前, C=4月合計, D=5月合計, ..., K=12月合計, L〜=データ
// ===========================================

/**
 * 読み取り専用：ホープ数シートの構造を確認（マイグレーション前）
 * GASエディタから手動実行 → ログを俺(Claude)に貼ってもらう
 */
function inspectHopeStructureBeforeMigration() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_HOPE_SHEET_NAME);
  if (!sheet) {
    Logger.log('❌ ホープ数シート見つからず');
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  Logger.log('=== ホープ数シート構造 ===');
  Logger.log('lastRow: ' + lastRow + ', lastCol: ' + lastCol + ' (' + postAppColToLetter_(lastCol) + ')');
  Logger.log('凍結列: ' + sheet.getFrozenColumns());
  Logger.log('凍結行: ' + sheet.getFrozenRows());

  // メタ列確認（A〜D）
  Logger.log('--- 1〜2行目のメタ列（A〜E）---');
  var metaHeader = sheet.getRange(1, 1, 2, 5).getValues();
  for (var i = 0; i < 2; i++) {
    Logger.log('  行' + (i + 1) + ': A=' + metaHeader[i][0] + ' B=' + metaHeader[i][1] + ' C=' + metaHeader[i][2] + ' D=' + metaHeader[i][3] + ' E=' + metaHeader[i][4]);
  }

  // データ起点（D列）の中身を確認 → これが日付ヘッダーなら旧構造、合計っぽい数字なら既に新構造
  Logger.log('--- D列〜の中身先頭5行 ---');
  if (lastRow >= 3) {
    var dCol = sheet.getRange(3, 4, Math.min(5, lastRow - 2), 1).getValues();
    for (var k = 0; k < dCol.length; k++) {
      Logger.log('  row=' + (k + 3) + ' D列=' + dCol[k][0]);
    }
  }

  // C列に何が入ってるか（合計式 or 値）
  Logger.log('--- C列の数式・値（先頭3メンバー）---');
  if (lastRow >= 3) {
    var cFormulas = sheet.getRange(3, 3, 3, 1).getFormulas();
    var cValues = sheet.getRange(3, 3, 3, 1).getValues();
    for (var n = 0; n < cFormulas.length; n++) {
      Logger.log('  row=' + (n + 3) + ' C列 数式=' + cFormulas[n][0] + ' 値=' + cValues[n][0]);
    }
  }

  Logger.log('=== 判定 ===');
  Logger.log('D列に「日付（4/1, 4/2など）」または「YT/IG/TT」が入ってる → 旧構造、マイグレーション必要');
  Logger.log('D列に「数字（合計値）」が入ってる → 既に新構造（5月合計列）、マイグレーション不要');
  Logger.log('=== inspect 完了 ===');
}

/**
 * ホープ数シートに月別合計列（D〜K）を追加する
 * 既存D〜のデータは右に8列シフトされる
 *
 * 処理内容:
 *   1. C列の右に8列挿入 → 既存4月分データ(D〜CO)が L〜DA にシフト
 *   2. C2セルを「4月合計」とリネーム（C列既存数式の対象列はずれないので維持）
 *   3. D2〜K2 に「5月合計」〜「12月合計」のラベル
 *   4. C3〜K{lastRow} に SUM 数式を入れる
 *
 * GASエディタから手動実行（一回限り）
 */
function migrateHopeAddMonthlyTotals() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_HOPE_SHEET_NAME);
  if (!sheet) {
    Logger.log('❌ ホープ数シート見つからず');
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  Logger.log('=== マイグレーション開始 ===');
  Logger.log('lastRow=' + lastRow + ', lastCol=' + lastCol + ' (' + postAppColToLetter_(lastCol) + ')');

  // 二重実行防止: D列の値が数字（既に合計列になってる）ならスキップ
  if (lastRow >= 3) {
    var d3Value = sheet.getRange(3, 4).getValue();
    var d3Formula = sheet.getRange(3, 4).getFormula();
    Logger.log('D3 値=' + d3Value + ' 数式=' + d3Formula);
    if (d3Formula && d3Formula.indexOf('SUM') >= 0) {
      Logger.log('⚠️ D列が既にSUM式 → 既にマイグレーション済み、中止');
      return;
    }
  }

  // Step 1: C列の右に8列挿入 → 既存D〜が L〜にシフト
  sheet.insertColumnsAfter(3, 8);
  Logger.log('✅ C列の右に8列挿入');

  // Step 2: メタヘッダー設定（C2〜K2 = 4月〜12月合計）
  var monthLabels = ['4月合計', '5月合計', '6月合計', '7月合計', '8月合計', '9月合計', '10月合計', '11月合計', '12月合計'];
  sheet.getRange(2, 3, 1, 9).setValues([monthLabels]).setHorizontalAlignment('center').setFontWeight('bold').setBackground('#E8F0FE');
  // 1行目（マージ）も埋める
  sheet.getRange(1, 3, 1, 9).setValue('月別合計').merge().setHorizontalAlignment('center').setFontWeight('bold').setBackground('#4A90D9').setFontColor('#FFFFFF');
  Logger.log('✅ メタヘッダー設定（C2〜K2）');

  // Step 3: 各行に SUM 数式（4月〜12月）
  if (lastRow >= 3) {
    // 各月のデータ列範囲を計算（4月: L〜DA, 5月: DB〜GS, ...）
    var dataStartCol = 12; // L列
    var monthRanges = [];
    var col = dataStartCol;
    for (var m = 4; m <= 12; m++) {
      var monthDays = new Date(2026, m, 0).getDate();
      var totalCols = monthDays * 3;
      var startLetter = postAppColToLetter_(col);
      var endLetter = postAppColToLetter_(col + totalCols - 1);
      monthRanges.push({ start: startLetter, end: endLetter });
      col += totalCols;
    }

    var formulas = [];
    for (var r = 3; r <= lastRow; r++) {
      var row = [];
      for (var mi = 0; mi < monthRanges.length; mi++) {
        row.push('=SUM(' + monthRanges[mi].start + r + ':' + monthRanges[mi].end + r + ')');
      }
      formulas.push(row);
    }
    sheet.getRange(3, 3, lastRow - 2, 9).setFormulas(formulas);
    Logger.log('✅ ' + (lastRow - 2) + '行に月別合計式を設定');
    Logger.log('  範囲例（4月）: C3 = SUM(' + monthRanges[0].start + '3:' + monthRanges[0].end + '3)');
    Logger.log('  範囲例（5月）: D3 = SUM(' + monthRanges[1].start + '3:' + monthRanges[1].end + '3)');
  }

  Logger.log('=== マイグレーション完了 ===');
  Logger.log('スプシで目視確認: C列〜K列が月別合計、L列以降が4月〜のデータ');
  Logger.log('注意: GAS本番デプロイ更新後でないと、サーバーが古い列位置で読み続けます');
}

// ============================================
// 全期間合計列追加 + 投稿数の月別合計列追加（2026-05-04）
// ============================================

/**
 * 読み取り専用：投稿数シートの構造確認
 */
function inspectPostStructureBeforeMigration() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) { Logger.log('❌ 投稿数シートなし'); return; }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  Logger.log('=== 投稿数シート構造 ===');
  Logger.log('lastRow: ' + lastRow + ', lastCol: ' + lastCol + ' (' + postAppColToLetter_(lastCol) + ')');
  Logger.log('凍結列: ' + sheet.getFrozenColumns() + ', 凍結行: ' + sheet.getFrozenRows());

  Logger.log('--- 1〜2行目のメタ列（A〜H）---');
  var meta = sheet.getRange(1, 1, 2, 8).getValues();
  for (var i = 0; i < 2; i++) {
    Logger.log('  行' + (i + 1) + ': A=' + meta[i][0] + ' B=' + meta[i][1] + ' C=' + meta[i][2] + ' D=' + meta[i][3] + ' E=' + meta[i][4] + ' F=' + meta[i][5] + ' G=' + meta[i][6] + ' H=' + meta[i][7]);
  }

  Logger.log('--- E列（合計）の数式・値（先頭3メンバー）---');
  if (lastRow >= 2) {
    var fs = sheet.getRange(2, 5, 3, 1).getFormulas();
    var vs = sheet.getRange(2, 5, 3, 1).getValues();
    for (var n = 0; n < fs.length; n++) {
      Logger.log('  row=' + (n + 2) + ' E列 数式=' + fs[n][0] + ' 値=' + vs[n][0]);
    }
  }

  Logger.log('--- F列（4/1） の中身先頭5行 ---');
  if (lastRow >= 2) {
    var fcol = sheet.getRange(2, 6, Math.min(5, lastRow - 1), 1).getValues();
    for (var k = 0; k < fcol.length; k++) {
      Logger.log('  row=' + (k + 2) + ' F列=' + fcol[k][0]);
    }
  }

  Logger.log('=== 判定 ===');
  Logger.log('F1が「4月合計」なら既に新構造');
  Logger.log('F1が日付/空欄、E列がCOUNTIF(F:AI...) → 旧構造、migrate必要');
  Logger.log('=== inspect 完了 ===');
}

/**
 * 投稿数 + ホープ数 両方の実体確認
 */
function inspectAllSheetsBeforeMigration() {
  inspectPostStructureBeforeMigration();
  Logger.log('');
  inspectHopeStructureBeforeMigration();
}

/**
 * ホープ数: C列の左に「全期間合計」列を1列追加
 * 既存: A=ID, B=名前, C=4月合計, D=5月合計, ..., K=12月合計, L〜=データ
 * 新規: A=ID, B=名前, C=全期間合計, D=4月合計, ..., L=12月合計, M〜=データ
 */
function migrateAddPeriodTotalToHope() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_HOPE_SHEET_NAME);
  if (!sheet) { Logger.log('❌ ホープ数シートなし'); return; }

  var lastRow = sheet.getLastRow();
  Logger.log('=== ホープ数: 全期間合計列追加 開始 ===');

  if (lastRow >= 2) {
    var c2 = sheet.getRange(2, 3).getValue();
    if (typeof c2 === 'string' && c2.indexOf('全期間') >= 0) {
      Logger.log('⚠️ C2="' + c2 + '" → 既に追加済み、中止');
      return;
    }
  }

  sheet.insertColumnsAfter(2, 1);
  Logger.log('✅ B列の右に1列挿入');

  sheet.getRange(1, 3).setValue('全期間合計').setHorizontalAlignment('center').setFontWeight('bold').setBackground('#FFD600');
  sheet.getRange(2, 3).setValue('全期間合計').setHorizontalAlignment('center').setFontWeight('bold').setBackground('#FFE082');
  Logger.log('✅ C1〜C2 ヘッダー設定');

  if (lastRow >= 3) {
    var formulas = [];
    for (var r = 3; r <= lastRow; r++) {
      formulas.push(['=SUM(D' + r + ':L' + r + ')']);
    }
    sheet.getRange(3, 3, lastRow - 2, 1).setFormulas(formulas);
    Logger.log('✅ C3〜C' + lastRow + ' に SUM(D:L)');
  }

  Logger.log('=== ホープ数 完了 ===');
}

/**
 * 投稿数: E列の右に「月別合計」列を9列追加
 * 既存: A=ID, B=契約日, C=CW_ID, D=名前, E=合計, F〜=データ
 * 新規: A=ID, B=契約日, C=CW_ID, D=名前, E=全期間合計, F=4月合計, ..., N=12月合計, O〜=データ
 */
function migrateAddMonthlyTotalsToPost() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) { Logger.log('❌ 投稿数シートなし'); return; }

  var lastRow = sheet.getLastRow();
  Logger.log('=== 投稿数: 月別合計列追加 開始 ===');

  if (lastRow >= 1) {
    var f1 = sheet.getRange(1, 6).getValue();
    if (typeof f1 === 'string' && f1.indexOf('月合計') >= 0) {
      Logger.log('⚠️ F1="' + f1 + '" → 既に追加済み、中止');
      return;
    }
  }

  sheet.insertColumnsAfter(5, 9);
  Logger.log('✅ E列の右に9列挿入（旧F〜が O〜にシフト）');

  sheet.getRange(1, 5).setValue('全期間合計').setHorizontalAlignment('center').setFontWeight('bold').setBackground('#FFD600');
  var monthLabels = ['4月合計', '5月合計', '6月合計', '7月合計', '8月合計', '9月合計', '10月合計', '11月合計', '12月合計'];
  sheet.getRange(1, 6, 1, 9).setValues([monthLabels]).setHorizontalAlignment('center').setFontWeight('bold').setBackground('#E8F0FE');
  Logger.log('✅ E1=全期間合計, F1〜N1=月別ヘッダー');

  // 各月のデータ列範囲（投稿数: 1日1列、データ起点 O列(15)）
  var dataStartCol = 15;
  var monthRanges = [];
  var col = dataStartCol;
  for (var m = 4; m <= 12; m++) {
    var monthDays = new Date(2026, m, 0).getDate();
    monthRanges.push({ start: postAppColToLetter_(col), end: postAppColToLetter_(col + monthDays - 1) });
    col += monthDays;
  }

  if (lastRow >= 2) {
    // E列（全期間合計）
    var eFormulas = [];
    for (var r = 2; r <= lastRow; r++) {
      eFormulas.push(['=SUM(F' + r + ':N' + r + ')']);
    }
    sheet.getRange(2, 5, lastRow - 1, 1).setFormulas(eFormulas);
    Logger.log('✅ E列（全期間合計）= SUM(F:N)');

    // F〜N列（各月COUNTIF合計）
    var formulas = [];
    for (var r2 = 2; r2 <= lastRow; r2++) {
      var row = [];
      for (var mi = 0; mi < monthRanges.length; mi++) {
        var s = monthRanges[mi].start, e = monthRanges[mi].end;
        var f = '=COUNTIF(' + s + r2 + ':' + e + r2 + ',"1本")' +
                '+COUNTIF(' + s + r2 + ':' + e + r2 + ',"2本")*2' +
                '+COUNTIF(' + s + r2 + ':' + e + r2 + ',"3本")*3' +
                '+COUNTIF(' + s + r2 + ':' + e + r2 + ',"4本")*4' +
                '+COUNTIF(' + s + r2 + ':' + e + r2 + ',"5本")*5' +
                '+COUNTIF(' + s + r2 + ':' + e + r2 + ',"6本")*6' +
                '+COUNTIF(' + s + r2 + ':' + e + r2 + ',"7本")*7' +
                '+COUNTIF(' + s + r2 + ':' + e + r2 + ',"8本")*8' +
                '+COUNTIF(' + s + r2 + ':' + e + r2 + ',"9本")*9' +
                '+COUNTIF(' + s + r2 + ':' + e + r2 + ',"10本")*10';
        row.push(f);
      }
      formulas.push(row);
    }
    sheet.getRange(2, 6, lastRow - 1, 9).setFormulas(formulas);
    Logger.log('✅ F〜N列 月別COUNTIF式設定');
    Logger.log('  範囲例（4月）: F2 = COUNTIF(' + monthRanges[0].start + '2:' + monthRanges[0].end + '2, ...)');
    Logger.log('  範囲例（5月）: G2 = COUNTIF(' + monthRanges[1].start + '2:' + monthRanges[1].end + '2, ...)');
  }

  Logger.log('=== 投稿数 完了 ===');
}

/**
 * 読み取り専用：ホープ数の凍結列・マージ状態確認
 */
function inspectHopeSheetMergeState() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_HOPE_SHEET_NAME);
  if (!sheet) { Logger.log('❌ ホープ数シートなし'); return; }
  Logger.log('=== ホープ数: 凍結・マージ状態 ===');
  Logger.log('凍結列: ' + sheet.getFrozenColumns());
  Logger.log('凍結行: ' + sheet.getFrozenRows());
  var vals = sheet.getRange(1, 1, 2, 13).getValues();
  for (var i = 0; i < 2; i++) {
    var row = '行' + (i + 1) + ':';
    for (var j = 0; j < 13; j++) {
      row += ' ' + postAppColToLetter_(j + 1) + '=' + vals[i][j];
    }
    Logger.log('  ' + row);
  }
  Logger.log('=== 確認完了 ===');
}

/**
 * ホープ数・投稿数のクリーンアップ
 * - ホープ数: 凍結列 → L列(12)、1行目 D〜L マージ「月別合計」
 * - 投稿数: 凍結列 → N列(14)
 */
function adjustHopeAndPostFreeze() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);

  var hopeSheet = ss.getSheetByName(POST_APP_HOPE_SHEET_NAME);
  if (hopeSheet) {
    Logger.log('=== ホープ数 調整 ===');
    Logger.log('凍結列: ' + hopeSheet.getFrozenColumns() + ' → 12 に変更');
    hopeSheet.setFrozenColumns(12);
    try {
      hopeSheet.getRange(1, 4, 1, 9).breakApart();
    } catch (e) { Logger.log('breakApart skipped: ' + e.message); }
    hopeSheet.getRange(1, 4, 1, 9)
      .merge()
      .setValue('月別合計')
      .setHorizontalAlignment('center')
      .setFontWeight('bold')
      .setBackground('#4A90D9')
      .setFontColor('#FFFFFF');
    Logger.log('✅ ホープ数 D1〜L1 を「月別合計」マージ');
  }

  var postSheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (postSheet) {
    Logger.log('=== 投稿数 調整 ===');
    Logger.log('凍結列: ' + postSheet.getFrozenColumns() + ' → 14 に変更');
    postSheet.setFrozenColumns(14);
    Logger.log('✅ 投稿数 凍結列 N列まで');
  }

  Logger.log('=== 完了 ===');
}

// ============================================
// 受講生アカウントURL機能 構造調査（読み取り専用）
// 新シート設計のための事前調査用。書き込み一切なし。
// ============================================
function inspectPostSheetStructureForUrlFeature() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);

  Logger.log('=== スプシ内シート一覧 ===');
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    Logger.log('  ' + sheets[i].getName() + ' (gid=' + sheets[i].getSheetId() + ')');
  }

  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) { Logger.log('!! ' + POST_APP_SHEET_NAME + ' シートが見つかりません'); return; }

  Logger.log('=== 投稿数シート 基本情報 ===');
  Logger.log('最終行: ' + sheet.getLastRow());
  Logger.log('最終列: ' + sheet.getLastColumn());
  Logger.log('最大列: ' + sheet.getMaxColumns());
  Logger.log('固定行(frozen rows): ' + sheet.getFrozenRows());
  Logger.log('固定列(frozen cols): ' + sheet.getFrozenColumns());

  var inspectCols = 15; // A〜O
  Logger.log('=== マージ範囲（行1〜5 × A〜O 内） ===');
  var topMerges = sheet.getRange(1, 1, 5, inspectCols).getMergedRanges();
  if (topMerges.length === 0) {
    Logger.log('  （マージなし）');
  } else {
    for (var i = 0; i < topMerges.length; i++) {
      Logger.log('  ' + topMerges[i].getA1Notation());
    }
  }

  Logger.log('=== ヘッダー行（row1）A〜O ===');
  var hdr = sheet.getRange(1, 1, 1, inspectCols).getValues()[0];
  var hdrBg = sheet.getRange(1, 1, 1, inspectCols).getBackgrounds()[0];
  for (var c = 0; c < inspectCols; c++) {
    Logger.log('  ' + postAppColToLetter_(c + 1) + ': ' + JSON.stringify(hdr[c]) + ' bg=' + hdrBg[c]);
  }

  Logger.log('=== サンプルデータ行 row2〜row6 A〜O（値・背景色） ===');
  var lastRow = Math.min(6, sheet.getLastRow());
  if (lastRow >= 2) {
    var dataVals = sheet.getRange(2, 1, lastRow - 1, inspectCols).getValues();
    var dataBg = sheet.getRange(2, 1, lastRow - 1, inspectCols).getBackgrounds();
    for (var r = 0; r < dataVals.length; r++) {
      Logger.log('--- row ' + (r + 2) + ' ---');
      for (var c = 0; c < inspectCols; c++) {
        Logger.log('  ' + postAppColToLetter_(c + 1) + ': ' + JSON.stringify(dataVals[r][c]) + ' bg=' + dataBg[r][c]);
      }
    }
  } else {
    Logger.log('  （データ行なし）');
  }

  Logger.log('=== 全データ行のID列・名前列 件数確認 ===');
  if (sheet.getLastRow() >= 2) {
    var rows = sheet.getLastRow() - 1;
    var idCol = sheet.getRange(2, POST_APP_ID_COL, rows, 1).getValues();
    var nameCol = sheet.getRange(2, POST_APP_NAME_COL, rows, 1).getValues();
    var nameBg = sheet.getRange(2, POST_APP_NAME_COL, rows, 1).getBackgrounds();
    var idCount = 0, nameCount = 0, bgCount = 0;
    var bgSamples = {};
    for (var i = 0; i < rows; i++) {
      if (idCol[i][0]) idCount++;
      if (nameCol[i][0]) nameCount++;
      var bg = nameBg[i][0];
      if (bg && bg !== '#ffffff') {
        bgCount++;
        bgSamples[bg] = (bgSamples[bg] || 0) + 1;
      }
    }
    Logger.log('  ID入力済み行数: ' + idCount);
    Logger.log('  名前入力済み行数: ' + nameCount);
    Logger.log('  名前列の非白背景色 行数: ' + bgCount);
    Logger.log('  名前列の背景色サンプル: ' + JSON.stringify(bgSamples));
  }
}