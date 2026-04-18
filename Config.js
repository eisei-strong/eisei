// ============================================
// Config.js — 基盤定数とヘルパー (v2)
// ============================================

// --- Chatwork APIトークン取得 ---
function getChatworkToken_() {
  return PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN') || '';
}

// --- スプレッドシートID ---
var SPREADSHEET_ID = '1k_x3aNRTbojmhJZGMS6JGNiTNJLQR4sD5zyJCBh1YqY';

// --- シート名定数 ---
var SHEET_SETTINGS      = '設定';
var SHEET_SUMMARY       = 'サマリー';
var SHEET_ARCHIVE       = '月次アーカイブ';
var SHEET_CO_MANAGE     = 'CO残管理';
var SHEET_DAILY_PREFIX  = '日別入力_';
var SHEET_SHINPAN       = '⚔️信販会社割合';

// --- 旧シート検索パターン ---
var SHEET_NAME_PATTERN  = 'ウォーリアーズ数値';

// --- チーム目標（デフォルト値、設定シートで上書き可能） ---
var TEAM_GOAL = 15000; // 1.5億円 = 15,000万円

// ============================================
// 日別入力シート: 列定義 (0-based index)
// ============================================
var COL_DATE             = 0;   // A: 日付
var COL_FUNDED_DEALS     = 1;   // B: 資金有 商談数
var COL_FUNDED_CLOSED    = 2;   // C: 資金有 成約数
var COL_REVENUE          = 3;   // D: 着金額（万円）
var COL_SALES            = 4;   // E: 売上（万円）
var COL_CO_COUNT         = 5;   // F: CO数
var COL_CO_AMOUNT        = 6;   // G: CO金額（万円）
var COL_FUNDED_PASSED    = 7;   // H: 資金有で見送り
var COL_UNFUNDED_PASSED  = 8;   // I: 資金なし 見送り数
var COL_UNFUNDED_CLOSED  = 9;   // J: 資金なし 成約数
var COL_UNFUNDED_DEALS   = 10;  // K: 資金なし 商談数

var DAILY_COL_COUNT  = 11;  // A〜K = 11列

// --- 日別入力シート: ヘッダーラベル ---
var DAILY_HEADERS = [
  '日付', '資金有 商談数', '資金有 成約数', '着金額（万円）', '売上（万円）',
  'CO数', 'CO金額（万円）', '資金有で見送り', '資金なし 見送り数', '資金なし 成約数', '資金なし 商談数'
];

// ============================================
// 日別入力シート: 行定義 (1-based, スプレッドシート行番号)
// ============================================
var DAILY_ROW_HEADER       = 1;
var DAILY_ROW_CARRYOVER    = 2;   // 前月繰越
var DAILY_ROW_DATA_START   = 3;   // 日別データ開始 (1日)
var DAILY_ROW_DATA_END     = 33;  // 日別データ終了 (31日)
var DAILY_ROW_TOTAL        = 35;  // 合計行
var DAILY_ROW_CBS_APPROVED = 37;  // CBS承認数
var DAILY_ROW_CBS_APPLIED  = 38;  // CBS申請数
var DAILY_ROW_LF_APPROVED  = 39;  // ライフティ承認数
var DAILY_ROW_LF_APPLIED   = 40;  // ライフティ申請数
var DAILY_ROW_CREDIT_TOTAL = 42;  // クレカ着金合計
var DAILY_ROW_SHINPAN_TOTAL= 43;  // 信販着金合計

// ============================================
// サマリーシート: セクション行定義 (1-based)
// ============================================
var SM_ROW_TITLE            = 1;
var SM_ROW_UPDATED          = 2;
var SM_ROW_KPI_REVENUE      = 3;   // 合計着金額（繰越含）
var SM_ROW_KPI_DEALS        = 4;   // 合計商談数（当月）
var SM_ROW_KPI_CLOSED       = 5;   // 合計成約数
var SM_ROW_KPI_SALES        = 6;   // 合計売上
var SM_ROW_KPI_CO           = 7;   // 合計CO
var SM_ROW_KPI_PROGRESS     = 8;   // 年間目標進捗率

var SM_ROW_MEMBER_HEADER    = 10;
var SM_ROW_MEMBER_START     = 11;  // メンバーデータ開始
var SM_ROW_MEMBER_END       = 26;  // 16人分 (11-26) ※10→16人に拡張
var SM_ROW_MEMBER_TOTAL     = 27;  // 合計行

var SM_ROW_PERIOD_HEADER    = 30;
var SM_ROW_PERIOD_1_10      = 31;
var SM_ROW_PERIOD_11_20     = 32;
var SM_ROW_PERIOD_21_END    = 33;

var SM_ROW_BREAKDOWN_HEADER = 36;
var SM_ROW_CREDIT_TOTAL     = 37;
var SM_ROW_SHINPAN_TOTAL    = 38;
var SM_ROW_CURRENT_REVENUE  = 39;
var SM_ROW_CARRYOVER_REVENUE= 40;
var SM_ROW_CUMULATIVE_CLOSED= 41;  // 8月からの累計成約数

// 退職者CO残セクション
var SM_ROW_CO_HEADER        = 43;
var SM_ROW_CO_START          = 44;

// --- サマリーシート: メンバーテーブル列定義 (1-based) ---
var SM_COL_NAME          = 1;   // A: メンバー
var SM_COL_RANK          = 2;   // B: ランク
var SM_COL_FUNDED_DEALS  = 3;   // C: 資金有商談
var SM_COL_UNFUNDED_DEALS= 4;   // D: 資金なし商談
var SM_COL_TOTAL_DEALS   = 5;   // E: 合計商談
var SM_COL_CLOSED        = 6;   // F: 成約数
var SM_COL_CLOSE_RATE    = 7;   // G: 成約率
var SM_COL_SALES         = 8;   // H: 売上
var SM_COL_REVENUE       = 9;   // I: 着金額(繰越含)
var SM_COL_CO            = 10;  // J: CO
var SM_COL_LIFETY_RATE   = 11;  // K: ライフティ承認率
var SM_COL_PREV_REVENUE  = 12;  // L: 前月着金額
var SM_COL_DIFF_REVENUE  = 13;  // M: 前月比（着金）
var SM_COL_GAP_TO_TOP    = 14;  // N: トップとの差
var SM_COL_DIFF_DEALS    = 15;  // O: 前月比（商談）
var SM_COL_DIFF_CLOSED   = 16;  // P: 前月比（成約）
var SM_COL_DIFF_RATE     = 17;  // Q: 前月比（成約率）
var SM_MEMBER_COL_COUNT  = 17;

// ============================================
// 設定シート: 定義
// ============================================
var SETTINGS_ROW_HEADER      = 1;
var SETTINGS_ROW_DATA_START  = 2;
var SETTINGS_COL_NAME        = 1;  // A: メンバー名
var SETTINGS_COL_DISPLAY_NAME= 2;  // B: 表示名
var SETTINGS_COL_RANK        = 3;  // C: 着金ランキング順位
var SETTINGS_COL_STATUS      = 4;  // D: ステータス
var SETTINGS_COL_COLOR       = 5;  // E: 配色コード
var SETTINGS_MEMBER_COL_COUNT= 5;

// グローバル設定（行22以降）※メンバー増加対応: 行2〜21で最大20人分
var SETTINGS_ROW_GLOBAL_START = 22;
// A列=ラベル, B列=値

// ============================================
// 月次アーカイブ: 列定義 (1-based)
// ============================================
var ARC_COL_YEAR           = 1;
var ARC_COL_MONTH          = 2;
var ARC_COL_NAME           = 3;
var ARC_COL_FUNDED_DEALS   = 4;
var ARC_COL_FUNDED_CLOSED  = 5;
var ARC_COL_UNFUNDED_DEALS = 6;
var ARC_COL_UNFUNDED_CLOSED= 7;
var ARC_COL_TOTAL_DEALS    = 8;
var ARC_COL_CLOSE_RATE     = 9;
var ARC_COL_SALES          = 10;
var ARC_COL_REVENUE        = 11;
var ARC_COL_CO_COUNT       = 12;
var ARC_COL_CO_AMOUNT      = 13;
var ARC_COL_CREDIT_CARD    = 14;
var ARC_COL_SHINPAN        = 15;
var ARC_COL_CBS            = 16;
var ARC_COL_LIFETY         = 17;

var ARCHIVE_HEADERS = [
  '年', '月', 'メンバー名', '資金有商談数', '資金有成約数', '資金なし商談数', '資金なし成約数',
  '合計商談数', '成約率', '売上', '着金額', 'CO数', 'CO金額', 'クレカ', '信販', 'CBS', 'ライフティ'
];

// ============================================
// CO残管理シート: 列定義 (1-based)
// ============================================
var CO_COL_MEMBER      = 1;   // A: メンバー名
var CO_COL_DEAL_DATE   = 2;   // B: 成約日
var CO_COL_CLIENT      = 3;   // C: 顧客名/案件ID
var CO_COL_DEAL_AMOUNT = 4;   // D: 成約金額（万円）
var CO_COL_CO_DATE     = 5;   // E: CO発生日
var CO_COL_CO_AMOUNT   = 6;   // F: CO金額（万円）
var CO_COL_STATUS      = 7;   // G: 回収ステータス
var CO_COL_CLAIM_DATE  = 8;   // H: 請求日
var CO_COL_COLLECT_DATE= 9;   // I: 回収日
var CO_COL_COLLECTED   = 10;  // J: 回収金額（万円）
var CO_COL_REMAINING   = 11;  // K: 未回収残高（万円）= F-J
var CO_COL_NOTE        = 12;  // L: 備考
var CO_COL_COUNT       = 12;

var CO_ROW_SUMMARY_TITLE  = 1;
var CO_ROW_SUMMARY_HEADER = 2;
var CO_ROW_SUMMARY_START  = 3;  // 退職者サマリー開始
var CO_ROW_SUMMARY_TOTAL  = 5;
var CO_ROW_DATA_HEADER    = 7;
var CO_ROW_DATA_START     = 8;

var CO_STATUSES = ['未請求', '請求済', '回収済', '回収不能'];

// ============================================
// 旧シート: メンバー列・行定義（フォールバック用）
// ============================================
var OLD_MEMBER_COLS = [2, 5, 11, 17, 20, 23, 26, 29];

var OLD_MEMBER_NAME_MAP = {
  2: '李信', 5: '勝友美', 11: '流川', 17: '首斬り桓騎',
  20: 'ヒカル', 23: '本田圭佑', 26: 'セナ', 29: '大飛'
};

var OLD_ROW_NAME         = 5;
var OLD_ROW_FUNDED_DEALS = 6;
var OLD_ROW_DEALS        = 7;
var OLD_ROW_CLOSED       = 8;
var OLD_ROW_RATE         = 9;
var OLD_ROW_SALES        = 10;
var OLD_ROW_AVG_PRICE    = 11;
var OLD_ROW_CO_REVENUE   = 12;
var OLD_ROW_CO_AMOUNT    = 13;
var OLD_ROW_REVENUE      = 14;
var OLD_ROW_CREDIT_CARD  = 15;
var OLD_ROW_SHINPAN      = 16;
var OLD_ROW_CBS          = 17;
var OLD_ROW_LIFETY       = 19;

// ============================================
// 3月以降 旧シート: メンバー別セクション定義
// summaryCol: サマリー行での列番号(1-indexed)
// dataStart/dataEnd: セクション内データ行(1-indexed)
// ============================================
var MEMBER_SECTIONS = [
  { name: '意思決定',              summaryCol: 3,  dataStart: 32,  dataEnd: 62 },
  { name: 'ポジティブ',                  summaryCol: 6,  dataStart: 68,  dataEnd: 98 },
  { name: 'トニー',                  summaryCol: 9,  dataStart: 325, dataEnd: 355 },
  { name: 'ヒトコト',                summaryCol: 12, dataStart: 104, dataEnd: 134 },
  { name: 'ゴン',                    summaryCol: 15, dataStart: 362, dataEnd: 392 },
  { name: 'ありのまま',            summaryCol: 18, dataStart: 140, dataEnd: 170 },
  { name: 'けつだん',                summaryCol: 21, dataStart: 288, dataEnd: 318 },
  { name: 'ぜんぶり',               summaryCol: 24, dataStart: 177, dataEnd: 207 },
  { name: 'スクリプト通りに営業するくん', summaryCol: 27, dataStart: 214, dataEnd: 244 },
  { name: 'スマイル',              summaryCol: 30, dataStart: 251, dataEnd: 281 }
];

// ============================================
// 名前マッピング
// ============================================

// 本名（苗字）→ v2メンバー名（⚔️信販会社割合シート用）
var REAL_NAME_TO_V2 = {
  '辻阪': 'ありのまま',
  '久保田': 'ヒトコト',
  '阿部': '意思決定',
  '伊東': 'ポジティブ',
  '川合': 'リヴァイ',
  '大内': 'ガロウ',
  '福島': 'けつだん',
  '五十嵐': 'ぜんぶり',
  '新居': 'スクリプトくん',
  '佐々木': 'スマイル',
  '佐々木心雪': 'スマイル',
  '長谷部': '長谷部',
  '吉崎': 'ゴジータ',
  '荒木': '悟空',
  'こうつさ': 'やまと'
};

// v1内部名 → v2メンバー名（移行用）
var LEGACY_TO_V2_NAME = {
  '李信': '意思決定',
  '童信': '意思決定',
  '信': '意思決定',
  'AをAでやる': '意思決定',
  'AをAで': '意思決定',
  'ビッグマウス': 'ありのまま',
  'ワントーン': 'スマイル',
  '勝友美': 'ポジティブ',
  'ドライ': 'ポジティブ',
  '流川': 'ヒトコト',
  '首斬り桓騎': 'ありのまま',
  '桓騎': 'ありのまま',
  'ヒカル': 'けつだん',
  '本田圭佑': 'ぜんぶり',
  'セナ': 'スクリプト通りに営業するくん',
  '大飛': 'スマイル'
};

// 旧表示名→v2メンバー名（フォールバック用）
var DISPLAY_NAME_MAP = {
  '流川': 'ヒトコト',
  '大飛': 'スマイル',
  '李信': '意思決定',
  '本田圭佑': 'ぜんぶり',
  '勝友美': 'ポジティブ',
  '首斬り桓騎': 'ありのまま',
  'セナ': 'スクリプトくん',
  'ヒカル': 'けつだん'
};

var NAME_MAP = {
  '童信': '李信'
};

// 旧シート表示名→v2メンバー名の追加マッピング
var OLD_DISPLAY_TO_V2 = {
  'ゴン': 'ゴン',  // 新メンバー（旧シートに追加）
  'スクリプト通りに営業': 'スクリプト通りに営業するくん'
};

// アイコンマップ（v2メンバー名キー）
var ICON_MAP = {
  'ヒトコト':     'https://lh3.googleusercontent.com/d/14TcuxzbVRRVNSjhlaOFDXdCXke_jV7m3',
  'スクリプト通りに営業するくん': 'https://lh3.googleusercontent.com/d/1BSBMs3h5BgC1z0Tx8jyprPmDs11LTBPn',
  'ポジティブ':       'https://lh3.googleusercontent.com/d/1N75FOIOJnh2Qun8fEUXmxwv3A3tLxlWR',
  '意思決定':   'https://giver.work/sales-dashboard/icons/ishikettei.png',
  'ぜんぶり':     'https://lh3.googleusercontent.com/d/11_mTOKu5m2MFoufn36NjUQyLjOdXrpa5',
  'けつだん':     'https://lh3.googleusercontent.com/d/1wnoxiF7PRZKSFPnjn0WXQb16hm-68Jlk',
  'スマイル':   'https://giver.work/sales-dashboard/icons/smile.png',
  'ありのまま': 'https://giver.work/sales-dashboard/icons/arinomama.png',
  'ゴン':         'https://lh3.googleusercontent.com/d/1iwBxoCgXfmfOoUhTv4OUy7mir9XmvjJV',
  'トニー':       'https://lh3.googleusercontent.com/d/1sHZ_zFFAitl7iVPEcIzQzpTD9cwL9FHv',
  '長谷部':       'https://appdata.chatwork.com/avatar/R769mN4PAr.rsz.jpg',
  'ゴジータ':     'https://appdata.chatwork.com/avatar/372J8vnz75.png',
  'L':            'https://appdata.chatwork.com/avatar/w7zBRgGD7l.png',
  '悟空':         'https://appdata.chatwork.com/avatar/Vq3WYmk4ql.png',
  'やまと':       'https://appdata.chatwork.com/avatar/Vq3WYnr8ql.rsz.png',
  '夜神月':       'https://appdata.chatwork.com/avatar/zMEPJERa73.png'
};

// Chatwork account_id → v2メンバー名（アバター自動取得用）
var CHATWORK_TO_V2 = {
  '4415237':  'ありのまま',
  '10258043': 'ヒトコト',
  '10751140': '意思決定',
  '10751530': 'ポジティブ',
  '9418659':  'けつだん',
  '10751652': 'スクリプト通りに営業するくん',
  '10750441': 'ぜんぶり',
  '9398311':  'スマイル',
  '11109913': 'ゴン',
  '11105287': 'トニー',
  '11159019': '長谷部',
  '10841091': 'ゴジータ',
  '11232346': 'L',
  '11205416': '悟空',
  '10471342': 'やまと',
  '11237452': '夜神月'
};

// ============================================
// ヘルパー関数
// ============================================

/** 数値を安全にパース */
function parseNum_(val) {
  if (typeof val === 'number') return val;
  if (!val || val === '-') return 0;
  var cleaned = String(val).replace(/[¥￥,、円万\s%％]/g, '');
  return Number(cleaned) || 0;
}

/** 名前を正規化（旧名→v1内部名） */
function normalizeName_(name) {
  return NAME_MAP[name] || name;
}

/** 表示名を取得（v1内部名→v2メンバー名） */
function displayName_(name) {
  return DISPLAY_NAME_MAP[name] || name;
}

/** アイコンURLを取得（v2メンバー名で検索） */
function iconUrl_(name) {
  return ICON_MAP[name] || '';
}

/**
 * ChatworkルームメンバーのアバターURLを取得してICON_MAPを更新
 * ICON_MAPに未登録のメンバーのみ追加（既存は上書きしない）
 * @param {string} roomId ルームID（デフォルト: ウォーリアーズ全体FF）
 * @returns {Object} {updated: [], skipped: [], unknown: []}
 */
function syncChatworkAvatars(roomId) {
  roomId = roomId || '412557550';
  var token = getChatworkToken_();
  if (!token) return { error: 'CHATWORK_API_TOKEN未設定' };

  var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/members';
  var response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-ChatWorkToken': token },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) return { error: 'API失敗: ' + response.getResponseCode() };

  var members = JSON.parse(response.getContentText());
  var result = { updated: [], skipped: [], unknown: [] };

  for (var i = 0; i < members.length; i++) {
    var m = members[i];
    var v2Name = CHATWORK_TO_V2[String(m.account_id)];
    if (!v2Name) {
      result.unknown.push({ name: m.name, id: m.account_id, avatar: m.avatar_image_url });
      continue;
    }
    if (ICON_MAP[v2Name]) {
      result.skipped.push(v2Name);
      continue;
    }
    ICON_MAP[v2Name] = m.avatar_image_url;
    result.updated.push({ name: v2Name, avatar: m.avatar_image_url });
  }

  return result;
}

/**
 * 全メンバーのアバターURLをChatworkから最新に更新（既存も上書き）
 * @returns {Object} 更新結果
 */
function refreshAllAvatars(roomId) {
  roomId = roomId || '412557550';
  var token = getChatworkToken_();
  if (!token) return { error: 'CHATWORK_API_TOKEN未設定' };

  var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/members';
  var response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-ChatWorkToken': token },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) return { error: 'API失敗: ' + response.getResponseCode() };

  var members = JSON.parse(response.getContentText());
  var updated = [];

  for (var i = 0; i < members.length; i++) {
    var m = members[i];
    var v2Name = CHATWORK_TO_V2[String(m.account_id)];
    if (!v2Name) continue;
    ICON_MAP[v2Name] = m.avatar_image_url;
    updated.push({ name: v2Name, avatar: m.avatar_image_url });
  }

  return { updated: updated };
}

/** 小数第1位まで丸める */
function round1_(val) {
  return Math.round(val * 10) / 10;
}

// ============================================
// シートアクセサー
// ============================================

/** スプレッドシートを開く */
function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/** 設定シートを取得 */
function getSettingsSheet_(ss) {
  return (ss || getSpreadsheet_()).getSheetByName(SHEET_SETTINGS);
}

/** サマリーシートを取得 */
function getSummarySheet_(ss) {
  return (ss || getSpreadsheet_()).getSheetByName(SHEET_SUMMARY);
}

/** 日別入力シートを取得（名前の揺れに対応するフォールバック付き） */
function getDailySheet_(ss, memberName) {
  ss = ss || getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_DAILY_PREFIX + memberName);
  if (sheet) return sheet;

  // フォールバック: LEGACY_TO_V2_NAME の逆引きで別名シートを探す
  for (var legacyName in LEGACY_TO_V2_NAME) {
    if (LEGACY_TO_V2_NAME[legacyName] === memberName) {
      sheet = ss.getSheetByName(SHEET_DAILY_PREFIX + legacyName);
      if (sheet) {
        // シート名を正規名にリネーム
        sheet.setName(SHEET_DAILY_PREFIX + memberName);
        return sheet;
      }
    }
  }

  // フォールバック2: 設定シートのメンバー名で探す
  var members = getMembersFromSettings_(ss);
  for (var i = 0; i < members.length; i++) {
    var m = members[i];
    var resolved = LEGACY_TO_V2_NAME[m.name] || DISPLAY_NAME_MAP[m.name] || m.name;
    if (resolved === memberName && m.name !== memberName) {
      sheet = ss.getSheetByName(SHEET_DAILY_PREFIX + m.name);
      if (sheet) {
        // シート名を正規名にリネーム
        sheet.setName(SHEET_DAILY_PREFIX + memberName);
        return sheet;
      }
    }
  }

  return null;
}

/** 月次アーカイブシートを取得 */
function getArchiveSheet_(ss) {
  return (ss || getSpreadsheet_()).getSheetByName(SHEET_ARCHIVE);
}

/** CO残管理シートを取得 */
function getCOManageSheet_(ss) {
  return (ss || getSpreadsheet_()).getSheetByName(SHEET_CO_MANAGE);
}

/** 旧シートを月番号から検索（"2月"が"12月"にマッチしないよう前方の文字もチェック） */
function getSheetByMonth_(ss, month) {
  var sheets = ss.getSheets();
  var monthStr = month + '月';
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf(SHEET_NAME_PATTERN) === -1) continue;
    var idx = name.indexOf(monthStr);
    if (idx < 0) continue;
    // 前の文字が数字でないことを確認（"12月"で"2月"にマッチしない）
    if (idx > 0) {
      var prevChar = name.charAt(idx - 1);
      if (prevChar >= '0' && prevChar <= '9') continue;
    }
    return sheets[i];
  }
  return null;
}

// ============================================
// 設定シート読み取り
// ============================================

/**
 * 設定シートからメンバー一覧を取得
 * @returns {Array<{name, displayName, rank, status, color, iconUrl}>}
 */
function getMembersFromSettings_(ss) {
  var sheet = getSettingsSheet_(ss);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < SETTINGS_ROW_DATA_START) return [];

  // 行14以降はグローバル設定なのでメンバーデータは行2〜13（最大）
  var maxMemberRow = Math.min(lastRow, SETTINGS_ROW_GLOBAL_START - 2);
  if (maxMemberRow < SETTINGS_ROW_DATA_START) return [];

  var data = sheet.getRange(SETTINGS_ROW_DATA_START, 1, maxMemberRow - SETTINGS_ROW_DATA_START + 1, SETTINGS_MEMBER_COL_COUNT).getValues();
  var members = [];

  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    if (!name) continue;

    // 旧名→v2正規名の自動修正（例: ポジティブ→ポジティブ）
    var v2Name = LEGACY_TO_V2_NAME[name];
    if (v2Name && ICON_MAP[v2Name]) {
      sheet.getRange(SETTINGS_ROW_DATA_START + i, SETTINGS_COL_NAME).setValue(v2Name);
      sheet.getRange(SETTINGS_ROW_DATA_START + i, SETTINGS_COL_DISPLAY_NAME).setValue(v2Name);
      name = v2Name;
    }

    var displayName = String(data[i][1] || '').trim() || name;
    // displayNameも旧名の場合修正
    if (LEGACY_TO_V2_NAME[displayName] && ICON_MAP[LEGACY_TO_V2_NAME[displayName]]) {
      displayName = LEGACY_TO_V2_NAME[displayName];
    }

    var status = String(data[i][3] || 'アクティブ').trim();

    members.push({
      name: name,
      displayName: displayName,
      rank: data[i][2],
      status: status,
      color: String(data[i][4] || '').trim(),
      iconUrl: iconUrl_(name) || iconUrl_(displayName)
    });
  }

  return members;
}

/** アクティブメンバーのみ取得 */
function getActiveMembers_(ss) {
  return getMembersFromSettings_(ss).filter(function(m) {
    return m.status === 'アクティブ' || m.status === 'active';
  });
}

/** 設定シートからグローバル設定を取得 */
function getGlobalSettings_(ss) {
  var sheet = getSettingsSheet_(ss);
  var now = new Date();
  var defaults = { year: now.getFullYear(), month: now.getMonth() + 1, teamGoal: TEAM_GOAL };
  if (!sheet) return defaults;

  var lastRow = sheet.getLastRow();
  if (lastRow < SETTINGS_ROW_GLOBAL_START) return defaults;

  var rowCount = lastRow - SETTINGS_ROW_GLOBAL_START + 1;
  var data = sheet.getRange(SETTINGS_ROW_GLOBAL_START, 1, rowCount, 2).getValues();
  var settings = { year: defaults.year, month: defaults.month, teamGoal: defaults.teamGoal };

  for (var i = 0; i < data.length; i++) {
    var label = String(data[i][0] || '').trim();
    var value = data[i][1];
    if (label === '年間着金目標' || label === 'チーム目標') settings.teamGoal = parseNum_(value) || settings.teamGoal;
    if (label === '対象年度' || label === '対象年') settings.year = parseNum_(value) || settings.year;
    if (label === '対象月') settings.month = parseNum_(value) || settings.month;
  }

  return settings;
}

/**
 * 新構造が有効かチェック
 */
function isNewStructureReady_(ss) {
  ss = ss || getSpreadsheet_();
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) return false;

  var summarySheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!summarySheet) return false;

  var val = summarySheet.getRange(SM_ROW_KPI_REVENUE, 2).getValue();
  return val !== '' && val !== null && val !== undefined;
}

/**
 * 日別入力シートから全データを読み取る (v2)
 * @returns {Object} {totals, daily[], cbs, lifety, carryover}
 */
function readDailySheetData_(sheet) {
  if (!sheet) return null;

  var lastRow = Math.max(sheet.getLastRow(), DAILY_ROW_SHINPAN_TOTAL);
  var lastCol = Math.max(sheet.getLastColumn(), DAILY_COL_COUNT);
  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // 合計行 (row 35, index 34)
  var totalIdx = DAILY_ROW_TOTAL - 1;
  var t = data[totalIdx] || [];
  var totals = {
    fundedDeals:    parseNum_(t[COL_FUNDED_DEALS]),
    fundedClosed:   parseNum_(t[COL_FUNDED_CLOSED]),
    revenue:        round1_(parseNum_(t[COL_REVENUE])),
    sales:          round1_(parseNum_(t[COL_SALES])),
    coCount:        parseNum_(t[COL_CO_COUNT]),
    coAmount:       round1_(parseNum_(t[COL_CO_AMOUNT])),
    fundedPassed:   parseNum_(t[COL_FUNDED_PASSED]),
    unfundedPassed: parseNum_(t[COL_UNFUNDED_PASSED]),
    unfundedClosed: parseNum_(t[COL_UNFUNDED_CLOSED]),
    unfundedDeals:  parseNum_(t[COL_UNFUNDED_DEALS])
  };

  // 計算フィールド
  totals.totalDeals = totals.fundedDeals + totals.unfundedDeals;
  totals.closeRate = totals.totalDeals > 0
    ? round1_((totals.fundedClosed / totals.totalDeals) * 100)
    : 0;
  totals.avgPrice = totals.fundedClosed > 0
    ? round1_(totals.sales / totals.fundedClosed)
    : 0;

  // クレカ・信販の合計行（行42/43から取得）
  var creditIdx = DAILY_ROW_CREDIT_TOTAL - 1;
  var shinpanIdx = DAILY_ROW_SHINPAN_TOTAL - 1;
  totals.creditCard = 0;
  totals.shinpan = 0;
  if (data.length > creditIdx) {
    totals.creditCard = round1_(parseNum_(data[creditIdx][1]));
  }
  if (data.length > shinpanIdx) {
    totals.shinpan = round1_(parseNum_(data[shinpanIdx][1]));
  }

  // CBS / ライフティ
  var cbsApproved = data[DAILY_ROW_CBS_APPROVED - 1] ? data[DAILY_ROW_CBS_APPROVED - 1][1] : '';
  var cbsApplied  = data[DAILY_ROW_CBS_APPLIED - 1]  ? data[DAILY_ROW_CBS_APPLIED - 1][1]  : '';
  var lfApproved  = data[DAILY_ROW_LF_APPROVED - 1]  ? data[DAILY_ROW_LF_APPROVED - 1][1]  : '';
  var lfApplied   = data[DAILY_ROW_LF_APPLIED - 1]   ? data[DAILY_ROW_LF_APPLIED - 1][1]   : '';

  var cbs = '-';
  if (cbsApproved !== '' || cbsApplied !== '') {
    cbs = String(cbsApproved || 0) + '/' + String(cbsApplied || 0);
  }
  var lifety = '-';
  if (lfApproved !== '' || lfApplied !== '') {
    lifety = String(lfApproved || 0) + '/' + String(lfApplied || 0);
  }

  // ライフティ承認率
  var lifetyRate = '-';
  var lfApprovedNum = parseNum_(lfApproved);
  var lfAppliedNum = parseNum_(lfApplied);
  if (lfAppliedNum > 0) {
    lifetyRate = round1_((lfApprovedNum / lfAppliedNum) * 100);
  }

  // 日別データ（着金速報用）
  var daily = [];
  for (var r = DAILY_ROW_DATA_START - 1; r < DAILY_ROW_DATA_END; r++) {
    if (r >= data.length) break;
    var row = data[r];
    var dateVal = row[COL_DATE];
    if (!(dateVal instanceof Date)) continue;

    var revVal = parseNum_(row[COL_REVENUE]);
    if (revVal > 0) {
      daily.push({
        date: Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy-MM-dd'),
        dateShort: (dateVal.getMonth() + 1) + '/' + dateVal.getDate(),
        revenue: round1_(revVal),
        day: dateVal.getDate()
      });
    }
  }

  // 前月繰越（row 2, index 1）
  var carryoverRow = data[DAILY_ROW_CARRYOVER - 1] || [];
  var carryover = {
    revenue: round1_(parseNum_(carryoverRow[COL_REVENUE]))
  };

  return {
    totals: totals,
    daily: daily,
    cbs: cbs,
    lifety: lifety,
    lifetyRate: lifetyRate,
    carryover: carryover
  };
}
