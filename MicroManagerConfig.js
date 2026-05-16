// ============================================
// MicroManagerConfig.js — 営業マイクロマネジメント 設定
// ============================================

// --- 投稿先ルーム ---
// ScriptProperties の 'MICRO_ROOM_ID' から取得（コード変更不要）
// 設定方法: microSetupAll() を1回実行 → 自動検出してScriptPropertiesに保存
// 手動で設定する場合: microSetRoomId('xxxxx')
var MICRO_ROOM_ID = (function () {
  try {
    return PropertiesService.getScriptProperties().getProperty('MICRO_ROOM_ID') || '';
  } catch (e) {
    return '';
  }
})();

// --- 招待リンク（参考、自動解決のヒント） ---
var MICRO_ROOM_INVITE_URL = 'https://www.chatwork.com/g/j35f6ai00udio7';

// --- 監視対象メンバー（5月以降の現役営業15人） ---
// account_id → 表示名
// ※ 新規メンバーの account_id は要追加（サンウォン/司波/信）
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
  // TODO: account_id 確認次第追加
  // 'XXXXXXXX': 'サンウォン',     // 笹山楓太
  // 'XXXXXXXX': '司波',           // 坂野宙輝
  // 'XXXXXXXX': '信'              // 吉田羚虹
};

// --- 提出種別 ---
var MICRO_TYPE_MORNING = 'MORNING_PLAN';
var MICRO_TYPE_DAILY   = 'DAILY_REPORT';
var MICRO_TYPE_WEEKLY  = 'WEEKLY_REVIEW';

// --- 期限定義（Asia/Tokyo） ---
var MICRO_DEADLINES = {
  MORNING_PLAN_HOUR: 1,    // 翌朝計画: 深夜1:59
  MORNING_PLAN_MIN:  59,
  DAILY_REPORT_HOUR: 21,   // 日報: 21:00
  DAILY_REPORT_MIN:  0,
  WEEKLY_REVIEW_DOW: 5,    // 週次振り返り: 金曜 (1=月 ... 5=金)
  WEEKLY_REVIEW_HOUR: 18,
  WEEKLY_REVIEW_MIN: 0
};

// --- ペナルティ金額（1ペナ = 1万円） ---
var MICRO_PENALTY_AMOUNT = 10000;

// --- 提出物の合格基準（AI判定の必須項目） ---
var MICRO_REQUIREMENTS = {
  MORNING_PLAN: [
    '翌朝の最優先タスク（1件、具体的・固有名詞ベース）',
    '翌日のアポ件数 or 架電件数の数値目標',
    '一番気合入れる商談 or 行動（だれに/なにを）'
  ],
  DAILY_REPORT: [
    '今日の実数（アポ数・契約数・失注数のうち2つ以上の数字）',
    '失注/未達の具体的な原因（精神論NG・行動レベルで）',
    '明日の打ち手（具体的な行動、抽象NG）',
    '朝の翌朝計画で宣言した内容への結果言及'
  ],
  WEEKLY_REVIEW: [
    '今週KPI実績 vs 目標（数字3つ以上）',
    '上手くいったこと1つ + 再現方法',
    '失敗1つ + 来週の対策',
    '来週の数字目標'
  ]
};

// --- ボットラベル（自分の投稿スキップ用） ---
var MICRO_BOT_LABEL = '[マイクロマネジメント]';

// --- ログシート名（既存スクリプトSPREADSHEET_ID内） ---
var MICRO_SHEET_LOG    = 'マイクロマネジメントログ';   // 各提出の判定結果
var MICRO_SHEET_PENA   = 'マイクロマネジメントペナ';   // 月次ペナ集計
var MICRO_SHEET_BOTLOG = 'マイクロマネジメントBot';    // bot動作ログ
var MICRO_SHEET_STATE  = 'マイクロマネジメント処理済'; // 重複処理防止

// --- AIモデル（ShachoFFBot同等） ---
var MICRO_AI_MODEL      = 'claude-opus-4-7';
var MICRO_AI_MAX_TOKENS = 1024;

// --- レート制御 ---
var MICRO_POLL_MIN = 5;  // 5分ポーリング

// --- KPI警告（ペナなし、詰めDMのみ） ---
var MICRO_KPI_ENABLED = true;
// 17:00 当日アポ0警告: 当日アポ数がこの数値以下で発火（0=完全ゼロ）
var MICRO_KPI_DAILY_THRESHOLD = 0;
// 水曜12:00 週次アポ警告: 月〜火のアポ合計がこの数値以下で発火
var MICRO_KPI_WEEKLY_THRESHOLD = 1;
// 月の21日以降 月次着金警告: 月初〜現在の着金がこの数値以下で発火（万円）
var MICRO_KPI_MONTHLY_REVENUE_THRESHOLD = 0;

// --- 表示名 → 内部v2名 マッピング（API突合用） ---
// MICRO_MEMBERS の表示名と api-proxy.php / Config の v2名のズレを吸収
var MICRO_DISPLAY_TO_V2 = {
  '1日1more':     'スクリプト通りに営業するくん',
  '言い切り':     'ゴン',
  '週1休みくん':  'トニー'
  // 一致する名前はマップ不要（意思決定/ヒトコト/ありのまま/ポジティブ/けつだん/ぜんぶり/ゴジータ/夜神月/悟空）
};

// ============================================
// ヘルパー
// ============================================

function microMonthKey_(d) {
  return Utilities.formatDate(d || new Date(), 'Asia/Tokyo', 'yyyy-MM');
}

function microDateKey_(d) {
  return Utilities.formatDate(d || new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
}

function microNowJst_() {
  return new Date();
}

// 翌朝計画 締切日時（その日の 1:59:59）を返す
function microMorningPlanDeadline_(d) {
  d = d || new Date();
  var t = new Date(d.getTime());
  t.setHours(MICRO_DEADLINES.MORNING_PLAN_HOUR, MICRO_DEADLINES.MORNING_PLAN_MIN, 59, 0);
  return t;
}

// 日報 締切日時（その日の 21:00:00）
function microDailyReportDeadline_(d) {
  d = d || new Date();
  var t = new Date(d.getTime());
  t.setHours(MICRO_DEADLINES.DAILY_REPORT_HOUR, MICRO_DEADLINES.DAILY_REPORT_MIN, 0, 0);
  return t;
}

// 週次 締切日時（その週の金曜 18:00）
function microWeeklyReviewDeadline_(d) {
  d = d || new Date();
  var t = new Date(d.getTime());
  // getDay(): 0=日 1=月 ... 5=金 6=土
  var diff = MICRO_DEADLINES.WEEKLY_REVIEW_DOW - t.getDay();
  t.setDate(t.getDate() + diff);
  t.setHours(MICRO_DEADLINES.WEEKLY_REVIEW_HOUR, MICRO_DEADLINES.WEEKLY_REVIEW_MIN, 0, 0);
  return t;
}

// スクリプト所有のスプシ（マスターとは別、書込みOK）
function getMicroSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// MICRO_ROOM_ID 未設定チェック
function microRoomReady_() {
  return MICRO_ROOM_ID && MICRO_ROOM_ID.length > 0;
}
