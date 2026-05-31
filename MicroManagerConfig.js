// ============================================
// MicroManagerConfig.js — 営業マイクロマネジメント 設定
// ============================================

// --- 投稿先ルーム ---
// ScriptProperties の 'MICRO_ROOM_ID' から取得（コード変更不要）
var MICRO_ROOM_ID = (function () {
  try {
    return PropertiesService.getScriptProperties().getProperty('MICRO_ROOM_ID') || '';
  } catch (e) {
    return '';
  }
})();

// --- 監視対象メンバー ---
// account_id → 表示名
var MICRO_MEMBERS = {
  '4415237':  'ありのまま',      // 辻阪
  '10258043': 'ヒトコト',        // 久保田
  '10751140': '意思決定',        // 阿部
  '10751530': 'ポジティブ',      // 伊東
  '9418659':  'けつだん',        // 福島
  '10751652': '1日1more',        // 新居
  '10750441': 'ぜんぶり',        // 五十嵐
  '11109913': '言い切り',        // 大久保
  '11105287': '週1休みくん',     // 矢吹
  '10841091': 'ゴジータ',        // 吉崎
  '11237452': '夜神月',          // 中市
  '11205416': '悟空'             // 荒木
  // TODO: account_id 確認次第追加（サンウォン/司波/信）
};

// --- 提出種別 ---
var MICRO_TYPE_MORNING = 'MORNING_PLAN';
var MICRO_TYPE_DAILY   = 'DAILY_REPORT';

// --- 提出物のキーワード判定（含まれてれば提出扱い） ---
var MICRO_KEYWORDS_MORNING = ['翌朝計画', '明朝計画', '明日のアポ', '明日の計画', '明日やる'];
var MICRO_KEYWORDS_DAILY   = ['日報', '今日の結果', '本日の報告', '本日の振り返り', '今日の報告'];

// --- 締切: 翌朝計画・日報ともに 深夜1:59 / チェック発火 2:00 ---
//   営業日キー(workDate) = (タイムスタンプ - 2h) の日付
//   例: 5/30 23:00 投稿 → workDate=2026-05-30
//       5/31 01:30 投稿 → workDate=2026-05-30
//       5/31 02:01 投稿 → workDate=2026-05-31
var MICRO_DEADLINE_HOUR = 1;
var MICRO_DEADLINE_MIN  = 59;

// --- ペナルティ金額（1ペナ = 1万円） ---
var MICRO_PENALTY_AMOUNT = 10000;

// --- ボットラベル ---
var MICRO_BOT_LABEL = '[マイクロマネジメント]';

// --- ログシート名 ---
var MICRO_SHEET_BOTLOG = 'マイクロマネジメントBot';    // bot動作ログ
var MICRO_SHEET_STATE  = 'マイクロマネジメント処理済'; // 重複処理防止

// --- レート制御 ---
var MICRO_POLL_MIN = 5;  // 5分ポーリング

// ============================================
// ヘルパー
// ============================================
function microMonthKey_(d) {
  return Utilities.formatDate(d || new Date(), 'Asia/Tokyo', 'yyyy-MM');
}

function microDateKey_(d) {
  return Utilities.formatDate(d || new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
}

// 営業日キー = (timestamp - 2h) の yyyy-MM-dd
function microWorkDateKey_(d) {
  var t = new Date((d || new Date()).getTime() - 2 * 60 * 60 * 1000);
  return Utilities.formatDate(t, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function getMicroSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function microRoomReady_() {
  return MICRO_ROOM_ID && MICRO_ROOM_ID.length > 0;
}
