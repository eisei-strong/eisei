// ============================================
// 投稿本数入力アプリ API
// ============================================

// 4月ホープ数スプレッドシートID
var POST_APP_SS_ID = '1rQSoM2zu38aPXJHHD6ILEgyM0SEDOXHXZXz0qK_nQuk';

// シート構造定数
var POST_APP_SHEET_NAME = '【4月】投稿数 ⚠️嬴政と星野のみ編集可能';
var POST_APP_AUTH_SHEET = '認証';
var POST_APP_ID_COL = 1;      // A列: ID
var POST_APP_PW_COL = 2;         // B列: パスワード（平文）
var POST_APP_CONTRACT_COL = 3;   // C列: 契約日
var POST_APP_CW_ROOM_COL = 4;   // D列: チャットワークグルチャID
var POST_APP_NAME_COL = 5;       // E列: 名前
var POST_APP_TOTAL_COL = 6;      // F列: 合計
var POST_APP_DATE_START_COL = 7;  // G列: 4/1 開始
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

// ---- ID比較（先頭0対応） ----
function matchId_(sheetVal, inputId) {
  var a = String(sheetVal).trim();
  var b = String(inputId).trim();
  if (a === b) return true;
  var na = parseInt(a);
  var nb = parseInt(b);
  return !isNaN(na) && !isNaN(nb) && na === nb;
}

// ---- 認証ヘルパー ----

function getAuthSheet_() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_AUTH_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(POST_APP_AUTH_SHEET);
    sheet.getRange(1, 1, 1, 2).setValues([['ID', 'パスワード']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.hideSheet();
  }
  return sheet;
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
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_AUTH_SHEET);
  if (sheet) ss.deleteSheet(sheet);
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
    if (matchId_(ids[i][0], id)) {
      var name = sheet.getRange(i + 2, POST_APP_NAME_COL).getValue();
      // メインシートAJ列の平文パスワードをチェック
      var plainPw = String(sheet.getRange(i + 2, POST_APP_PW_COL).getValue() || '').trim();
      var hasPassword = !!plainPw;
      // なければ認証シートのハッシュもチェック（後方互換）
      if (!hasPassword) {
        var authSheet = getAuthSheet_();
        var authLast = authSheet.getLastRow();
        if (authLast >= 2) {
          var authData = authSheet.getRange(2, 1, authLast - 1, 2).getValues();
          for (var a = 0; a < authData.length; a++) {
            if (matchId_(authData[a][0], id)) { hasPassword = true; break; }
          }
        }
      }
      var contract = sheet.getRange(i + 2, POST_APP_CONTRACT_COL).getValue();
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
  if (!sheet) return { error: 'シートが見つかりません' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: 'データがありません' };
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  var memberRow = -1;
  for (var i = 0; i < ids.length; i++) {
    if (matchId_(ids[i][0], id)) { memberRow = i + 2; break; }
  }
  if (memberRow < 0) return { error: 'IDが見つかりません' };

  // メインシートにパスワード（平文）を保存
  sheet.getRange(memberRow, POST_APP_PW_COL).setValue(password);

  // 認証シートにもハッシュを保存（後方互換）
  var authSheet = getAuthSheet_();
  var authLast = authSheet.getLastRow();
  if (authLast >= 2) {
    var authData = authSheet.getRange(2, 1, authLast - 1, 1).getValues();
    for (var a = 0; a < authData.length; a++) {
      if (matchId_(authData[a][0], id)) {
        authSheet.getRange(a + 2, 2).setValue(hashPassword_(password));
        return { ok: true };
      }
    }
  }

  authSheet.appendRow([String(id).trim(), hashPassword_(password)]);
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
    if (matchId_(sid, id) && semail === email.trim().toLowerCase()) {
      matched = true;
      break;
    }
  }
  if (!matched) return { error: 'メールアドレスが一致しません' };

  // メインシートのパスワード（B列）もクリア
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
      for (var r = 0; r < ids.length; r++) {
        if (matchId_(ids[r][0], id)) {
          sheet.getRange(r + 2, POST_APP_PW_COL).clearContent();
          break;
        }
      }
    }
  }

  // 認証シートからパスワードを削除
  var authSheet = getAuthSheet_();
  var authLast = authSheet.getLastRow();
  if (authLast >= 2) {
    var authData = authSheet.getRange(2, 1, authLast - 1, 1).getValues();
    for (var a = authData.length - 1; a >= 0; a--) {
      if (matchId_(authData[a][0], id)) {
        authSheet.deleteRow(a + 2);
      }
    }
  }

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
    if (matchId_(ids[i][0], id)) { memberRow = i + 2; break; }
  }
  if (memberRow < 0) return { error: 'IDが見つかりません' };

  // まずメインシートの平文パスワード（AJ列）でチェック
  var plainPw = String(sheet.getRange(memberRow, POST_APP_PW_COL).getValue() || '').trim();
  var matched = false;
  if (plainPw && plainPw === password) {
    matched = true;
  }

  // 平文になければ認証シートのハッシュでチェック（後方互換）
  if (!matched) {
    var authSheet = getAuthSheet_();
    var authLast = authSheet.getLastRow();
    if (authLast >= 2) {
      var authData = authSheet.getRange(2, 1, authLast - 1, 2).getValues();
      var hash = hashPassword_(password);
      for (var a = 0; a < authData.length; a++) {
        if (matchId_(authData[a][0], id)) {
          if (authData[a][1] === hash) matched = true;
          break;
        }
      }
    }
  }

  if (!matched) return { error: 'パスワードが正しくありません' };

  // ログイン成功 → AJ列にパスワード平文がなければ書き込む（マイグレーション）
  migratePasswordOnLogin_(sheet, memberRow, password);

  // トークン生成
  var token = hashPassword_(id + '_' + new Date().getTime() + '_' + Math.random());
  var cache = CacheService.getScriptCache();
  cache.put('postapp_token_' + token, id, 21600); // CacheService上限は6時間(21600秒)

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
  if (lastRow < 2) return { error: 'データがありません' };
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  var rowIdx = -1;
  for (var i = 0; i < ids.length; i++) {
    if (matchId_(ids[i][0], id)) { rowIdx = i; break; }
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
    var h = headers[i];
    var label;
    if (h instanceof Date) {
      label = (h.getMonth() + 1) + '/' + h.getDate();
    } else {
      label = String(h);
    }
    days.push({
      col: i,
      label: label,
      value: String(values[i] || '❌')
    });
  }
  var contract = sheet.getRange(row, POST_APP_CONTRACT_COL).getValue();
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

  var validValues = ['❌', '1本', '2本', '3本'];
  if (validValues.indexOf(value) < 0) return { error: '無効な値です: ' + value };

  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) return { error: 'シートが見つかりません' };

  // col = 0〜29（日付列のインデックス）
  var colNum = parseInt(col);
  if (isNaN(colNum) || colNum < 0 || colNum >= POST_APP_MONTH_DAYS) return { error: '日付が無効です' };

  // 当日までしか入力できない（未来日ブロック・JST基準）
  var now = new Date();
  var todayDay = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'dd'));
  if (colNum > todayDay - 1) return { error: '未来の日付は入力できません' };

  var targetCol = POST_APP_DATE_START_COL + colNum;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: 'データがありません' };
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  var rowIdx = -1;
  for (var i = 0; i < ids.length; i++) {
    if (matchId_(ids[i][0], id)) { rowIdx = i; break; }
  }
  if (rowIdx < 0) return { error: 'IDが見つかりません' };

  var row = rowIdx + 2;
  sheet.getRange(row, targetCol).setValue(value);
  SpreadsheetApp.flush();
  var total = sheet.getRange(row, POST_APP_TOTAL_COL).getValue();

  return { ok: true, saved: value, col: colNum, total: total };
}

// ============================================
// 連続ログイン記録
// ============================================

// ログイン日を記録（PropertiesService）
function recordLogin_(id) {
  try {
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
  } catch (e) {
    // PropertiesService障害時もログイン自体は継続させる
    Logger.log('recordLogin_ error for ' + id + ': ' + e.message);
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

  // 現在の連続日数（今日から遡る・JST基準）
  var streak = 0;
  var todayParts = todayStr.split('-');
  var checkY = parseInt(todayParts[0]);
  var checkM = parseInt(todayParts[1]) - 1;
  var checkD = parseInt(todayParts[2]);

  while (true) {
    var cd = new Date(checkY, checkM, checkD);
    var ds = Utilities.formatDate(cd, 'Asia/Tokyo', 'yyyy-MM-dd');
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
// ログイン状況をスプシ背景色に反映
// ============================================

/**
 * ログインした日の❌セルを薄緑にする
 * GASエディタから手動実行 or メニューから実行
 */
function colorLoginStatus() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません'); return; }

  var props = PropertiesService.getScriptProperties();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  var dataRange = sheet.getRange(2, POST_APP_DATE_START_COL, lastRow - 1, POST_APP_MONTH_DAYS);
  var values = dataRange.getValues();
  var backgrounds = dataRange.getBackgrounds();

  var now = new Date();
  var prefix = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM');

  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i][0]).trim();
    if (!id) continue;

    // PropertiesServiceのキーを探す（数値変換対応）
    var key = 'postapp_logins_' + id;
    var raw = props.getProperty(key);
    if (!raw && !isNaN(parseInt(id))) {
      // 0 → 0000 等のケースに対応（逆引き不可なのでスキップ）
    }
    var dates = raw ? JSON.parse(raw) : [];

    for (var d = 0; d < POST_APP_MONTH_DAYS; d++) {
      var dateStr = prefix + '-' + ('0' + (d + 1)).slice(-2);
      var loggedIn = dates.indexOf(dateStr) >= 0;
      var val = String(values[i][d]);

      if (val && val !== '❌' && val !== '') {
        // 投稿済み → 背景なし（白）
        backgrounds[i][d] = '#ffffff';
      } else if (loggedIn) {
        // ログインしたけど未投稿 → 薄緑
        backgrounds[i][d] = '#C6F6D5';
      } else {
        // ログインしていない → 白
        backgrounds[i][d] = '#ffffff';
      }
    }
  }

  dataRange.setBackgrounds(backgrounds);
  Logger.log('ログイン背景色を更新しました（' + ids.length + '人）');
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
  // POST_APP_CONTRACT_COL はグローバル定義（C列=3）を使用
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
    h[POST_APP_PW_COL - 1] = 'パスワード';
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
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // ヘッダー+1行以下ならソート不要

  // Range.sort() を使用（JS .sort() + setValues は禁止）
  // Google Sheets の Range.sort() は空欄を自動的に末尾に配置する
  var lastCol = sheet.getLastColumn();
  sheet.getRange(2, 1, lastRow - 1, lastCol).sort({
    column: POST_APP_CONTRACT_COL,
    ascending: true
  });
  Logger.log('契約日順にソートしました（Range.sort使用）');
}

// ============================================
// 21時チャットワーク通知（未入力リマインダー）
// ============================================

// POST_APP_CW_ROOM_COL は上部で定義済み（D列=4）

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

  // 全データ一括取得（N+1クエリ防止）
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  var names = sheet.getRange(2, POST_APP_NAME_COL, lastRow - 1, 1).getValues();
  var todayVals = sheet.getRange(2, todayCol, lastRow - 1, 1).getValues();
  var roomIds = sheet.getRange(2, POST_APP_CW_ROOM_COL, lastRow - 1, 1).getValues();
  // 連続日数計算用に日付列を一括取得
  var allDayData = sheet.getRange(2, POST_APP_DATE_START_COL, lastRow - 1, todayDay).getValues();

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

    // 連続日数を計算（バッチ取得済みデータから）
    var streak = 0;
    for (var j = todayDay - 2; j >= 0; j--) {
      var prevVal = String(allDayData[i][j] || '').trim();
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
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + POST_APP_SHEET_NAME); return; }

  // Namakaくん(ID=0000)の行を探す
  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (matchId_(ids[i][0], '0000')) {
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
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + POST_APP_SHEET_NAME); return; }
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
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + POST_APP_SHEET_NAME); return; }
  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, POST_APP_ID_COL, lastRow - 1, 1).getValues();

  // 既存行があれば削除
  for (var i = ids.length - 1; i >= 0; i--) {
    if (matchId_(ids[i][0], id)) {
      sheet.deleteRow(i + 2);
    }
  }

  // 2行目に挿入
  sheet.insertRowAfter(1);
  // A列をテキスト形式に（先頭0が消えないように）
  sheet.getRange(2, POST_APP_ID_COL).setNumberFormat('@');
  // A=ID, B=PW, C=契約日, D=CW ID, E=名前, F=合計数式, G〜AJ=❌
  var f = '=COUNTIF(G2:AJ2,"1本")+COUNTIF(G2:AJ2,"2本")*2+COUNTIF(G2:AJ2,"3本")*3';
  var newRow = [id, password, '', '', name, f];
  for (var d = 0; d < 30; d++) newRow.push('❌');
  sheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
  Logger.log('Inserted ' + name + ' at row 2');

  // パスワード登録
  var authSheet = getAuthSheet_();
  var authLast = authSheet.getLastRow();
  if (authLast >= 2) {
    var authData = authSheet.getRange(2, 1, authLast - 1, 1).getValues();
    for (var a = authData.length - 1; a >= 0; a--) {
      if (matchId_(authData[a][0], id)) authSheet.deleteRow(a + 2);
    }
  }
  authSheet.appendRow([id, hashPassword_(password)]);
  Logger.log('Password set for ' + name + ' (ID: ' + id + ')');
}

// ---- ランキング取得 ----
function postAppRanking_() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) return { ranking: [], error: 'シートが見つかりません' };
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

  return { ranking: members.slice(0, 50) };
}

// ---- D列の名前をLPスプシから復元 ----
function restoreNames() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + POST_APP_SHEET_NAME); return; }
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

// ---- B列にパスワード列を挿入（1回だけ実行） ----
function insertPasswordColumnB() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + POST_APP_SHEET_NAME); return; }

  // 再実行ガード: B1が既に「パスワード」なら中断
  var b1 = String(sheet.getRange(1, 2).getValue() || '').trim();
  if (b1 === 'パスワード') {
    Logger.log('B列は既にパスワード列です。重複挿入を防止しました。');
    return;
  }

  // B列（2列目）に列を挿入
  sheet.insertColumnBefore(2);
  sheet.getRange(1, 2).setValue('パスワード');

  // B列を保護（編集者以外変更不可）
  var protection = sheet.getRange(1, 2, sheet.getMaxRows(), 1).protect();
  protection.setDescription('パスワード列（自動管理）');
  protection.setWarningOnly(true);

  Logger.log('B列にパスワード列を挿入しました');
  Logger.log('※ 列が1つ右にずれています。全定数は更新済みです');
  Logger.log('次に fixPostAppTotal を実行してF列の数式を更新してください');
}

// ---- ログイン成功時にAJ列が空なら平文を書き込む（マイグレーション） ----
function migratePasswordOnLogin_(sheet, memberRow, password) {
  var existing = String(sheet.getRange(memberRow, POST_APP_PW_COL).getValue() || '').trim();
  if (!existing) {
    sheet.getRange(memberRow, POST_APP_PW_COL).setValue(password);
  }
}

// ---- E列(合計)に数式を一括設定 ----
function fixPostAppTotal() {
  var ss = SpreadsheetApp.openById(POST_APP_SS_ID);
  var sheet = ss.getSheetByName(POST_APP_SHEET_NAME);
  if (!sheet) { Logger.log('シートが見つかりません: ' + POST_APP_SHEET_NAME); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var formulas = [];
  for (var r = 2; r <= lastRow; r++) {
    formulas.push(['=COUNTIF(G' + r + ':AJ' + r + ',"1本")+COUNTIF(G' + r + ':AJ' + r + ',"2本")*2+COUNTIF(G' + r + ':AJ' + r + ',"3本")*3']);
  }
  sheet.getRange(2, POST_APP_TOTAL_COL, formulas.length, 1).setFormulas(formulas);
  Logger.log('Done: ' + formulas.length + ' rows updated');
}
