// ============================================
// 投稿本数入力アプリ API
// ============================================

// 4月ホープ数スプレッドシートID
var POST_APP_SS_ID = '1rQSoM2zu38aPXJHHD6ILEgyM0SEDOXHXZXz0qK_nQuk';

// シート構造定数
var POST_APP_SHEET_NAME = '【4月】投稿数';
var POST_APP_AUTH_SHEET = '認証';
var POST_APP_ID_COL = 1;      // A列: ID
var POST_APP_NAME_COL = 4;    // D列: 名前
var POST_APP_TOTAL_COL = 5;   // E列: 合計
var POST_APP_DATE_START_COL = 6;  // F列: 4/1 開始
var POST_APP_MONTH_DAYS = 30;     // 4月は30日

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

function postAppGet_(token) {
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
  var monthData = getMonthData_(sheet, row);

  // ログイン日を記録
  recordLogin_(String(id).trim());
  var loginData = getLoginStreakData_(String(id).trim());

  return { ok: true, id: String(id).trim(), name: String(name), month: monthData, login: loginData };
}

// ---- 月データ一括取得ヘルパー ----

function getMonthData_(sheet, row) {
  var headers = sheet.getRange(1, POST_APP_DATE_START_COL, 1, POST_APP_MONTH_DAYS).getValues()[0];
  var values = sheet.getRange(row, POST_APP_DATE_START_COL, 1, POST_APP_MONTH_DAYS).getValues()[0];
  var total = sheet.getRange(row, POST_APP_TOTAL_COL).getValue();

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
  return { days: days, total: total, contract: contractStr };
}

// ---- 保存（任意の日付を指定可能） ----

function postAppSave_(token, value, col) {
  var id = verifyToken_(token);
  if (!id) return { error: 'セッション切れです。再ログインしてください。' };

  var validValues = ['❌', '1本', '2本', '3本', '4本', '5本', '6本'];
  if (validValues.indexOf(value) < 0) return { error: '無効な値です: ' + value };

  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) return { error: 'シートが見つかりません' };

  // col = 0〜29（日付列のインデックス）
  var colNum = parseInt(col);
  if (isNaN(colNum) || colNum < 0 || colNum >= POST_APP_MONTH_DAYS) return { error: '日付が無効です' };
  var targetCol = POST_APP_DATE_START_COL + colNum;

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
      // F列(6)〜AI列(35) = 30日分を❌でリセット
      var blank = [];
      for (var d = 0; d < 30; d++) blank.push('❌');
      sheet.getRange(row, POST_APP_DATE_START_COL, 1, 30).setValues([blank]);
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
  for (var d = 0; d < 30; d++) newRow.push('❌');
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

// ---- ランキング取得 ----
function postAppRanking_() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ranking: [] };

  var data = sheet.getRange(2, 1, lastRow - 1, POST_APP_DATE_START_COL + POST_APP_MONTH_DAYS - 1).getValues();
  var now = new Date();
  var todayDay = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'dd'));

  var members = [];
  for (var i = 0; i < data.length; i++) {
    var id = String(data[i][POST_APP_ID_COL - 1] || '').trim();
    var name = String(data[i][POST_APP_NAME_COL - 1] || '').trim();
    if (!id || !name) continue;

    // 投稿本数合計
    var total = 0;
    var streak = 0;
    var postDays = 0;
    for (var d = 0; d < POST_APP_MONTH_DAYS; d++) {
      var v = String(data[i][POST_APP_DATE_START_COL - 1 + d] || '').trim();
      if (v && v !== '❌') {
        var num = parseInt(v) || 0;
        total += num;
        postDays++;
      }
    }
    // 連続日数（末尾から）
    for (var j = todayDay - 1; j >= 0; j--) {
      var v2 = String(data[i][POST_APP_DATE_START_COL - 1 + j] || '').trim();
      if (v2 && v2 !== '❌') streak++;
      else break;
    }

    if (total > 0 || postDays > 0) {
      members.push({ id: id, name: name, total: total, streak: streak, postDays: postDays });
    }
  }

  // 投稿本数で降順ソート
  members.sort(function(a, b) { return b.total - a.total || b.streak - a.streak; });

  return { ranking: members };
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