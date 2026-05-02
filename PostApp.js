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
var POST_APP_TOTAL_COL = 5;   // E列: 当月合計
var POST_APP_DATE_START_COL = 6;  // F列: 運用開始月の1日

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
 *   ホープ数: D列(4) - メタは A=ID, B=名前, C=合計
 *   プッシュ数: L列(12) - メタは A=ID, B〜J=名前(マージ等), K=合計
 *
 * 引数なしで呼ばれた場合は後方互換のためホープ数として扱う。
 */
function getCurrentMonthHopeColRange_(sheetName) {
  var dataStartCol = 4; // デフォルト: ホープ数(D列)
  if (sheetName === POST_APP_PUSH_SHEET_NAME) {
    dataStartCol = 12; // プッシュ数: L列
  }
  var COLS_PER_DAY = 3;
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth() + 1;
  var startCol = dataStartCol;
  var startYm = POST_APP_RUN_START_YEAR * 12 + (POST_APP_RUN_START_MONTH - 1);
  var nowYm = year * 12 + (month - 1);
  for (var i = startYm; i < nowYm; i++) {
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
// データ取得: 「ホープ数」シートから当月の列範囲を読む
function postAppGetHope_(token) {
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

  // ② ホープ数シートから当月データ取得
  var hopeSheet = ss.getSheetByName(POST_APP_HOPE_SHEET_NAME);
  if (!hopeSheet) {
    return { allowed: true, total: 0, days: [] };
  }
  var hopeLastRow = hopeSheet.getLastRow();
  if (hopeLastRow < 3) {
    return { allowed: true, total: 0, days: [] };
  }

  var hopeIds = hopeSheet.getRange(3, 1, hopeLastRow - 2, 1).getValues();
  var hopeRowIdx = -1;
  for (var j = 0; j < hopeIds.length; j++) {
    if (String(hopeIds[j][0]).trim() === String(id).trim()) { hopeRowIdx = j; break; }
  }
  if (hopeRowIdx < 0) {
    // 権限はあるが、ホープ数シートに自分の行がまだ無い → タブは出すがデータゼロ
    return { allowed: true, total: 0, days: [] };
  }

  var row = hopeRowIdx + 3;
  var hRange = getCurrentMonthHopeColRange_();
  var values = hopeSheet.getRange(row, hRange.startCol, 1, hRange.totalCols).getValues()[0];
  var total = hopeSheet.getRange(row, 3).getValue(); // C列: 当月合計

  var days = [];
  for (var d = 0; d < hRange.monthDays; d++) {
    var yt = parseInt(values[d * 3] || 0) || 0;
    var ig = parseInt(values[d * 3 + 1] || 0) || 0;
    var tt = parseInt(values[d * 3 + 2] || 0) || 0;
    days.push({ day: d + 1, yt: yt, ig: ig, tt: tt, sum: yt + ig + tt });
  }

  return { allowed: true, total: total || 0, days: days };
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