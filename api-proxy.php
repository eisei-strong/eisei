<?php
// 営業ダッシュボード API プロキシ v6
// 全データソース: マスターCSV（旧シート debugSheetByGid 廃止）
// paymentNews/prev月: GAS action=api
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Cache-Control: public, max-age=60');

// ===== 設定 =====
// 注意: GAS exec URLは今後使わない（トリガー＝clasp pushのみ、デプロイ不要）
// フォールバック用に残すが、主要データはMaster CSVから直接取得
$GAS_URL = 'https://script.google.com/macros/s/AKfycby6qaaiUoadCBnxHlUNKd-RkHxarE0WBGiitkdV0IbzL6ninM-df0FFx4SYRYVfdwcxqg/exec';
$MASTER_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1KxHeLmrpdaw1IUhBaQ46UWSHu-8SCRZqcrHOE2hMwDo/export?format=csv&gid=326094286';
$CACHE_DIR = __DIR__ . '/cache';
$CACHE_TTL_LIVE = 180;
$CACHE_TTL_ARCHIVE = 3600;
$CACHE_TTL_OTHER = 600;

if (!is_dir($CACHE_DIR)) mkdir($CACHE_DIR, 0755, true);

// ===== 共創pt（Chatwork投稿回数カウント） =====
$CW_TOKEN = '561f22f75377bfa3c9a5c1ba18d38342';
$KYOSO_ROOM_ID = '419910408';
$KYOSO_FILE = $CACHE_DIR . '/kyoso_counts.json';

// Chatwork表示名 → ダッシュボード名マッピング
$CW_NAME_MAP = [
    // ありのまま（旧ビッグマウス）
    'ありのままを捨てる' => 'ありのまま',
    '桓齮' => 'ありのまま',
    '桓騎' => 'ありのまま',
    '首斬り桓騎' => 'ありのまま',
    '辻阪' => 'ありのまま',
    '辻坂' => 'ありのまま',
    'ビッグマウス' => 'ありのまま',
    // 意思決定（旧AをAでやる）
    'AをAでやる' => '意思決定',
    '阿部' => '意思決定',
    '李信' => '意思決定',
    '信' => '意思決定',
    // ポジティブ
    'ドライ' => 'ポジティブ',
    '勝友美' => 'ポジティブ',
    '伊東' => 'ポジティブ',
    '勝' => 'ポジティブ',
    // ぜんぶり
    '五十嵐' => 'ぜんぶり',
    '本田圭佑' => 'ぜんぶり',
    // ヒトコト
    '流川' => 'ヒトコト',
    '久保田' => 'ヒトコト',
    // セナ（旧スクリプトくん）
    'スクリプト通りに営業' => '1日1more',
    'スクリプト通りに営業するくん' => '1日1more',
    'スクリプトくん' => '1日1more',
    '新居' => '1日1more',
    // スマイル（旧ワントーン）
    'ワントーン' => 'スマイル',
    '佐々木' => 'スマイル',
    '佐々木心雪' => 'スマイル',
    // ゴン
    '大久保' => '言い切り',
    '大久保友佑悟' => '言い切り',
    // 週1休みくん（旧トニー）
    '矢吹' => '週1休みくん',
    '矢吹友一' => '週1休みくん',
    'トニー' => '週1休みくん',
    // ゴジータ
    '吉崎' => 'ゴジータ',
    '吉崎息吹' => 'ゴジータ',
    // 悟空
    '荒木' => '悟空',
    // やまと
    'こうつさ' => 'やまと',
    // ダイレクトマッチ
    '意思決定' => '意思決定',
    'ポジティブ' => 'ポジティブ',
    'ヒトコト' => 'ヒトコト',
    'ありのまま' => 'ありのまま',
    'ぜんぶり' => 'ぜんぶり',
    '1日1more' => '1日1more',
    'スマイル' => 'スマイル',
    '言い切り' => '言い切り',
    '週1休みくん' => '週1休みくん',
    'ゴジータ' => 'ゴジータ',
    'L' => 'L',
    '悟空' => '悟空',
    'やまと' => 'やまと',
    '夜神月' => '夜神月',
    '福島' => 'けつだん',
    'けつだん' => 'けつだん',
];

// 属性列インデックス（マスターシート列追加後に実際の値に更新）
$COL_SEIKATSU_HOGO = 89;      // 生活保護
$COL_SEISHIN_SHIKKAN = 90;    // 精神疾患
$COL_SHOUGAISHA_TECHOU = 91;  // 障害者手帳

// 本名 → v2メンバー名（マスターCSV用）
$REAL_NAME_MAP = [
    '阿部' => '意思決定',
    '伊東' => 'ポジティブ',
    '久保田' => 'ヒトコト',
    '辻阪' => 'ありのまま',
    '五十嵐' => 'ぜんぶり',
    '新居' => '1日1more',
    '佐々木' => 'スマイル',
    '佐々木心雪' => 'スマイル',
    '大久保' => '言い切り',
    '大久保友佑悟' => '言い切り',
    '矢吹' => '週1休みくん',
    '矢吹友一' => '週1休みくん',
    '勝友美' => 'ポジティブ',
    '勝' => 'ポジティブ',
    'ドライ' => 'ポジティブ',
    '吉崎' => 'ゴジータ',
    '吉崎息吹' => 'ゴジータ',
    'ゴジータ' => 'ゴジータ',
    'L' => 'L',
    '鍋嶋' => 'L',
    '中市' => '夜神月',
    '荒木' => '悟空',
    '悟空' => '悟空',
    'こうつさ' => 'やまと',
    'やまと' => 'やまと',
    '夜神月' => '夜神月',
    '福島' => 'けつだん',
    'けつだん' => 'けつだん',
];

// レガシー名 → 現在の名前（アーカイブの旧名マッピング）
$LEGACY_NAME_MAP = [
    'ドライ' => 'ポジティブ',
    '勝友美' => 'ポジティブ',
    'スクリプト通りに営業' => '1日1more',
    'スクリプト通りに営業するくん' => '1日1more',
    'スクリプトくん' => '1日1more',
    '李信' => '意思決定',
    'AをAでやる' => '意思決定',
    '流川' => 'ヒトコト',
    '首斬り桓騎' => 'ありのまま',
    'ビッグマウス' => 'ありのまま',
    'ワントーン' => 'スマイル',
    'トニー' => '週1休みくん',
];

// アイコンマップ（4月以前用 = 本番現状の13人）
// アイコンURLは自前ホスト形式（giver.work/sales-dashboard/icons/）に統一
$ICON_MAP = [
    '意思決定' => 'https://giver.work/sales-dashboard/icons/ishikettei.png',
    'ポジティブ' => 'https://giver.work/sales-dashboard/icons/positive.png',
    '週1休みくん' => 'https://giver.work/sales-dashboard/icons/shu1yasumi.png',
    'ヒトコト' => 'https://giver.work/sales-dashboard/icons/hitokoto.png',
    '言い切り' => 'https://giver.work/sales-dashboard/icons/iikiri.png',
    'ありのまま' => 'https://giver.work/sales-dashboard/icons/arinomama.png',
    'ぜんぶり' => 'https://giver.work/sales-dashboard/icons/zenburi.png',
    '1日1more' => 'https://giver.work/sales-dashboard/icons/ichinichi1more.png',
    'スマイル' => 'https://giver.work/sales-dashboard/icons/smile.png',
    'ゴジータ' => 'https://giver.work/sales-dashboard/icons/gojita.png',
    'L' => 'https://giver.work/sales-dashboard/icons/l.png',
    '夜神月' => 'https://giver.work/sales-dashboard/icons/yagami.png',
    'けつだん' => 'https://giver.work/sales-dashboard/icons/ketsudan.png',
];

// チームマップ（4月以前用）
$TEAM_MAP = [
    '意思決定' => 1, '言い切り' => 1, 'ありのまま' => 1, '1日1more' => 1,
    'ぜんぶり' => 2, 'スマイル' => 2, 'ヒトコト' => 2, 'ポジティブ' => 2,
    'L' => 3, '夜神月' => 3, '週1休みくん' => 3, 'ゴジータ' => 3, 'けつだん' => 3,
];
$TEAM_NAMES = [1 => 'チーム1億', 2 => 'シリウス', 3 => 'ジャイアントキリング'];

// ===== 5月以降用マップ（2026/05〜）=====
// 変更点: L・スマイル除外、サンウォン・司波・信・悟空 追加、チーム制度は廃止だが既存所属は残す（新規4人は無所属）
$REAL_NAME_MAP_MAY2026 = [
    '阿部' => '意思決定',
    '伊東' => 'ポジティブ',
    // '久保田' => 'ヒトコト',  // 5月から休止（データはマスターに保存、復活時に戻す）
    '辻阪' => 'ありのまま',
    '五十嵐' => 'ぜんぶり',
    '新居' => '1日1more',
    '大久保' => '言い切り',
    '大久保友佑悟' => '言い切り',
    '矢吹' => '週1休みくん',
    '矢吹友一' => '週1休みくん',
    '勝友美' => 'ポジティブ',
    '勝' => 'ポジティブ',
    'ドライ' => 'ポジティブ',
    '吉崎' => 'ゴジータ',
    '吉崎息吹' => 'ゴジータ',
    'ゴジータ' => 'ゴジータ',
    '中市' => '夜神月',
    '夜神月' => '夜神月',
    '福島' => 'けつだん',
    'けつだん' => 'けつだん',
    // 5月から復活
    '荒木' => '悟空',
    '荒木泰人' => '悟空',
    '悟空' => '悟空',
    // 5月から新規
    '笹山' => 'サンウォン',
    '笹山楓太' => 'サンウォン',
    'サンウォン' => 'サンウォン',
    '坂野' => '司波',
    '坂野宙輝' => '司波',
    '司波' => '司波',
    '吉田' => '信',
    '吉田羚虹' => '信',
    '信' => '信',
    // 押切は意思決定のメンター専用表示（メンバー一覧には出さない）→ ICON_MAPに含めない
    // 「鍋嶋(L)」「佐々木/佐々木心雪(スマイル)」は意図的に削除 → 集計対象外
];

$ICON_MAP_MAY2026 = [
    '意思決定' => 'https://giver.work/sales-dashboard/icons/ishikettei.png',
    'ポジティブ' => 'https://giver.work/sales-dashboard/icons/positive.png',
    '週1休みくん' => 'https://giver.work/sales-dashboard/icons/shu1yasumi.png',
    // 'ヒトコト' => 'https://giver.work/sales-dashboard/icons/hitokoto.png',  // 5月から休止
    '言い切り' => 'https://giver.work/sales-dashboard/icons/iikiri.png',
    'ありのまま' => 'https://giver.work/sales-dashboard/icons/arinomama.png',
    'ぜんぶり' => 'https://giver.work/sales-dashboard/icons/zenburi.png',
    '1日1more' => 'https://giver.work/sales-dashboard/icons/ichinichi1more.png',
    'ゴジータ' => 'https://giver.work/sales-dashboard/icons/gojita.png',
    '夜神月' => 'https://giver.work/sales-dashboard/icons/yagami.png',
    'けつだん' => 'https://giver.work/sales-dashboard/icons/ketsudan.png',
    // 5月から
    'サンウォン' => 'https://giver.work/sales-dashboard/icons/sungwon.png',
    '司波' => 'https://giver.work/sales-dashboard/icons/shiba.png',
    '信' => 'https://giver.work/sales-dashboard/icons/shin.png',
    '悟空' => 'https://giver.work/sales-dashboard/icons/goku.png',
    // '押切' は意思決定のメンター表示専用 → メンバー一覧から除外 (Dashboard-wp.html PULLING_PAIRS の pullerIconUrl で直接参照)
    // 'L', 'スマイル' は意図的に削除
];

$TEAM_MAP_MAY2026 = [
    // 5月からチーム制度廃止だが、既存メンバーの所属は維持。新規4人は無所属(=未指定)
    '意思決定' => 1, '言い切り' => 1, 'ありのまま' => 1, '1日1more' => 1,
    'ぜんぶり' => 2, 'ポジティブ' => 2,
    '夜神月' => 3, 'ゴジータ' => 3,
    // 'L', 'スマイル', '週1休みくん', 'けつだん', 'ヒトコト' は5月メンバー一覧外、新規4人は無所属
];

/**
 * 当該月が2026年5月以降かを判定
 */
function isMay2026OrLater($month, $year) {
    return ($year > 2026) || ($year == 2026 && $month >= 5);
}

/**
 * 当該月の名前マップ・アイコンマップ・チームマップを返す
 */
function getMapsForMonth($month, $year) {
    global $REAL_NAME_MAP, $ICON_MAP, $TEAM_MAP,
           $REAL_NAME_MAP_MAY2026, $ICON_MAP_MAY2026, $TEAM_MAP_MAY2026;
    if (isMay2026OrLater($month, $year)) {
        return [$REAL_NAME_MAP_MAY2026, $ICON_MAP_MAY2026, $TEAM_MAP_MAY2026];
    }
    return [$REAL_NAME_MAP, $ICON_MAP, $TEAM_MAP];
}

// ===== ヘルパー =====

function gasRequest($url) {
    $ch = curl_init();
    curl_setopt_array($ch, [
        CURLOPT_URL => $url,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_FOLLOWLOCATION => true,
        CURLOPT_MAXREDIRS => 5,
        CURLOPT_TIMEOUT => 30,
        CURLOPT_SSL_VERIFYPEER => true,
    ]);
    $resp = curl_exec($ch);
    $code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);
    return ($resp !== false && $code === 200) ? $resp : false;
}

function parseCsv($csv) {
    $rows = [];
    $handle = fopen('php://temp', 'r+');
    fwrite($handle, $csv);
    rewind($handle);
    while (($row = fgetcsv($handle)) !== false) {
        $rows[] = $row;
    }
    fclose($handle);
    return $rows;
}

// マスターCSVをリクエスト内でキャッシュして返す
function getMasterCsvRows() {
    global $MASTER_SHEET_URL;
    static $cached = null;
    if ($cached !== null) return $cached;

    $csv = gasRequest($MASTER_SHEET_URL);
    if (!$csv) return null;

    $cached = parseCsv($csv);
    return empty($cached) ? null : $cached;
}

// ===== 共創pt関数 =====

function resolveKyosoName($cwName) {
    global $CW_NAME_MAP;
    if (isset($CW_NAME_MAP[$cwName])) return $CW_NAME_MAP[$cwName];
    // 部分一致（CWマップのキーがCW名に含まれているか）
    foreach ($CW_NAME_MAP as $key => $val) {
        if (mb_strlen($key) >= 2 && mb_strpos($cwName, $key) !== false) return $val;
    }
    return $cwName;
}

function updateKyosoCounts() {
    global $CW_TOKEN, $KYOSO_ROOM_ID, $KYOSO_FILE;

    // 既存データ読み込み
    $data = file_exists($KYOSO_FILE) ? json_decode(file_get_contents($KYOSO_FILE), true) : [];
    if (!is_array($data)) $data = [];

    $now = new DateTime('now', new DateTimeZone('Asia/Tokyo'));
    $month = $now->format('Y-m');
    if (!isset($data[$month])) $data[$month] = [];

    // Chatwork APIから新着メッセージ取得（force=0: 未読のみ）
    $url = "https://api.chatwork.com/v2/rooms/$KYOSO_ROOM_ID/messages?force=0";
    $ch = curl_init();
    curl_setopt_array($ch, [
        CURLOPT_URL => $url,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER => ["X-ChatWorkToken: $CW_TOKEN"],
        CURLOPT_TIMEOUT => 10,
    ]);
    $resp = curl_exec($ch);
    $code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($code === 200 && $resp) {
        $messages = json_decode($resp, true);
        if (is_array($messages)) {
            foreach ($messages as $msg) {
                $cwName = $msg['account']['name'] ?? '';
                if (!$cwName) continue;
                $name = resolveKyosoName($cwName);
                if (!isset($data[$month][$name])) $data[$month][$name] = 0;
                $data[$month][$name]++;
            }
        }
    }
    // 204 = 新着なし（正常）

    // 保存
    file_put_contents($KYOSO_FILE, json_encode($data, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT), LOCK_EX);

    return $data[$month] ?? [];
}

function getKyosoData() {
    global $KYOSO_FILE;
    $now = new DateTime('now', new DateTimeZone('Asia/Tokyo'));
    $month = $now->format('Y-m');
    $data = file_exists($KYOSO_FILE) ? json_decode(file_get_contents($KYOSO_FILE), true) : [];
    return is_array($data) && isset($data[$month]) ? $data[$month] : [];
}

function resolveV2Name($rawField, $nameMap = null) {
    global $REAL_NAME_MAP;
    if ($nameMap === null) $nameMap = $REAL_NAME_MAP;
    $rawPerson = explode('：', str_replace(':', '：', $rawField))[0];
    $rawPerson = trim($rawPerson);
    if (isset($nameMap[$rawPerson])) return $nameMap[$rawPerson];
    foreach ($nameMap as $key => $val) {
        if (mb_strpos($rawPerson, $key) === 0) return $val;
    }
    return null;
}

function serveCache($cacheFile) {
    if (file_exists($cacheFile)) { readfile($cacheFile); exit; }
    http_response_code(502);
    echo '{"error":"GAS API unavailable"}';
    exit;
}

function writeCache($cacheFile, $json) {
    file_put_contents($cacheFile, $json, LOCK_EX);
}

function parseAmount($s) {
    if (!$s || !trim($s)) return 0.0;
    return floatval(str_replace(',', '', trim($s)));
}

/** 支払日付文字列をyyyy-MM-dd形式にパース（着金速報用） */
function parsePayDate($dateStr, $defaultYear) {
    $d = trim($dateStr);
    if (!$d) return null;
    // "2026/03/10" or "2026/3/10" 形式
    if (preg_match('/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/', $d, $m)) {
        return sprintf('%04d-%02d-%02d', $m[1], $m[2], $m[3]);
    }
    // "3/10" or "03/10" 形式（年なし → defaultYear）
    if (preg_match('/^(\d{1,2})[\/\-](\d{1,2})/', $d, $m)) {
        return sprintf('%04d-%02d-%02d', $defaultYear, $m[1], $m[2]);
    }
    return null;
}

// ===== マスターCSVから全データ構築 =====

function fetchFromMasterCSV($month, $year) {
    global $MASTER_SHEET_URL, $TEAM_NAMES, $COL_SEIKATSU_HOGO, $COL_SEISHIN_SHIKKAN, $COL_SHOUGAISHA_TECHOU;

    // 月別マップ取得（5月以降: 新メンバー4人追加 + L/スマイル除外、4月以前: 既存）
    list($nameMap, $iconMap, $teamMap) = getMapsForMonth($month, $year);

    $rows = getMasterCsvRows();
    if (!$rows) return null;

    // アポ日時の月マッチ用
    $monthPrefix = $year . '/' . $month;
    $monthPrefix2 = $year . '/' . str_pad($month, 2, '0', STR_PAD_LEFT);

    // 前月プレフィックス（前月成約の当月着金判定用）
    $prevMonth = $month - 1;
    $prevYear = $year;
    if ($prevMonth < 1) { $prevMonth = 12; $prevYear--; }
    $prevMonthPrefix = $prevYear . '/' . $prevMonth;
    $prevMonthPrefix2 = $prevYear . '/' . str_pad($prevMonth, 2, '0', STR_PAD_LEFT);

    // 支払日付の月マッチ用（過去成約の当月着金判定）
    //
    // 仕様:
    //   ① 年あり日付（例: "2026/04/03"）→ その年が $year と一致するときだけ採用
    //   ② 年なし日付（例: "4/3"）→ アポ年と同年の着金とみなす → アポ年 == $year のときだけ採用
    //   ③ アポ年が取れない場合 → 安全のため不採用（false）
    //
    $targetMonth = intval($month);
    $isPayMonth = function($dateStr, $appoTs = '') use ($year, $targetMonth) {
        if (!$dateStr || !trim($dateStr)) return false;
        $d = trim($dateStr);

        // ① 年あり日付: "YYYY/M/D" or "YYYY/MM/DD"
        if (preg_match('/^(\d{4})\/(\d{1,2})\//', $d, $m)) {
            $payYear = intval($m[1]);
            $payMonth = intval($m[2]);
            return ($payYear === $year && $payMonth === $targetMonth);
        }

        // ② 年なし日付: "M/D" or "MM/DD" → アポ年と同年の着金とみなす
        if (preg_match('/^(\d{1,2})\/(\d{1,2})/', $d, $m)) {
            $payMonth = intval($m[1]);
            if ($payMonth !== $targetMonth) return false;
            // アポ年を取得: アポ年 == target year のときだけ採用
            if ($appoTs && preg_match('/^(\d{4})\//', $appoTs, $am)) {
                $appoYear = intval($am[1]);
                return ($appoYear === $year);
            }
            // アポ年不明 → 安全のため不採用
            return false;
        }

        return false;
    };

    // 支払スロット: [日付列, 着金額列, 支払手段列]
    $paySlots = [
        [38, 41, 40],  // ①
        [45, 48, 47],  // ②
        [52, 55, 54],  // ③
        [63, 66, 65],  // ④
        [70, 73, 72],  // ⑤
        [77, 80, 79],  // ⑥
        [84, 87, 86],  // ⑦
    ];

    // 重複ブロック検出（全行スキャン）
    $lastDataRow = 0;
    $duplicateStart = PHP_INT_MAX;
    foreach ($rows as $idx => $row) {
        if ($idx === 0 || count($row) < 12) continue;
        if (!resolveV2Name($row[3] ?? '', $nameMap)) continue;
        if ($lastDataRow > 0 && $idx - $lastDataRow > 500) {
            $duplicateStart = $idx;
            break;
        }
        $lastDataRow = $idx;
    }

    // 着金速報用
    $paymentNews = [];

    // dailyPushes用（日付別プッシュ数）
    $dailyPushTotals = [];   // dateKey => count
    $dailyPushByMember = []; // dateKey => [ memberName => count ]

    // 全メンバーを初期化（データがなくても表示するため）
    $memberData = [];
    foreach ($iconMap as $name => $icon) {
        $memberData[$name] = [
            'revenue' => 0, 'pastRevenue' => 0, 'prevMonthRevenue' => 0, 'deals' => 0, 'closed' => 0, 'coCount' => 0,
            'closedOnclass' => 0, 'closedConsul' => 0,
            'lost' => 0, 'lostByLfCbs' => 0, 'lostContinuing' => 0, 'continuing' => 0,
            'sales' => 0, 'coAmount' => 0, 'coRevenue' => 0,
            'creditCard' => 0, 'shinpan' => 0, 'fundedDeals' => 0,
            'cbsApproved' => 0, 'cbsApplied' => 0,
            'lfApproved' => 0, 'lfApplied' => 0,
            'seikatsuHogo' => 0, 'seishinShikkan' => 0, 'shougaishaTechou' => 0,
        ];
    }

    // メインスキャン
    foreach ($rows as $idx => $row) {
        if ($idx === 0 || count($row) < 12) continue;
        if ($idx >= $duplicateStart) break;

        $v2Name = resolveV2Name($row[3] ?? '', $nameMap);
        if (!$v2Name) continue;
        if (!isset($memberData[$v2Name])) {
            $memberData[$v2Name] = [
                'revenue' => 0, 'deals' => 0, 'closed' => 0, 'coCount' => 0,
                'closedOnclass' => 0, 'closedConsul' => 0,
                'lost' => 0, 'lostByLfCbs' => 0, 'lostContinuing' => 0, 'continuing' => 0,
                'sales' => 0, 'coAmount' => 0, 'coRevenue' => 0,
                'creditCard' => 0, 'shinpan' => 0, 'fundedDeals' => 0,
                'cbsApproved' => 0, 'cbsApplied' => 0,
                'lfApproved' => 0, 'lfApplied' => 0,
                'seikatsuHogo' => 0, 'seishinShikkan' => 0, 'shougaishaTechou' => 0,
            ];
        }

        $ts = isset($row[2]) ? $row[2] : $row[0];
        $isCurrentMonth = (strpos($ts, $monthPrefix) === 0 || strpos($ts, $monthPrefix2) === 0);

        // 成約➔CO / 成約➔キャンセル / 成約➔失注 は成約・着金から除外（当月・過去月とも）
        $status = trim($row[11]);
        $isExcludedStatus = (strpos($status, '成約') !== false && (
            strpos($status, 'CO') !== false ||
            strpos($status, 'キャンセル') !== false ||
            strpos($status, '失注') !== false
        ));

        // --- dailyPushes: 当月アポの日付別カウント ---
        if ($isCurrentMonth) {
            $pushDateKey = parsePayDate($ts, $year);
            if ($pushDateKey) {
                $dailyPushTotals[$pushDateKey] = ($dailyPushTotals[$pushDateKey] ?? 0) + 1;
                if (!isset($dailyPushByMember[$pushDateKey])) $dailyPushByMember[$pushDateKey] = [];
                $dailyPushByMember[$pushDateKey][$v2Name] = ($dailyPushByMember[$pushDateKey][$v2Name] ?? 0) + 1;
            }
        }

        // --- 商談集計（当月アポのみ） ---
        // ※ deals は個別カウントせず、最終出力時に closed + 全失注 で再計算
        if ($isCurrentMonth) {
            $product = isset($row[12]) ? trim($row[12]) : '';
            $payMethod1 = isset($row[19]) ? trim($row[19]) : '';
            $payMethod2 = isset($row[20]) ? trim($row[20]) : '';
            $lfVal = isset($row[24]) ? str_replace('✅', '', trim($row[24])) : '';
            $cbsVal = isset($row[25]) ? str_replace('✅', '', trim($row[25])) : '';
            $contractAmount = parseAmount($row[16] ?? '');

            if (strpos($status, '成約') !== false && strpos($status, 'CO') !== false) {
                // 成約➔CO: COカウント + 失注扱い
                $memberData[$v2Name]['coCount']++;
                $memberData[$v2Name]['coAmount'] += $contractAmount;
            } elseif ($isExcludedStatus) {
                // 成約➔キャンセル / 成約➔失注: 失注としてカウント
                $memberData[$v2Name]['lost']++;
            } elseif (strpos($status, '成約') !== false) {
                $memberData[$v2Name]['closed']++;
                $memberData[$v2Name]['sales'] += $contractAmount;
                if (mb_strpos($product, 'オンクラス') !== false) {
                    $memberData[$v2Name]['closedOnclass']++;
                } else {
                    $memberData[$v2Name]['closedConsul']++;
                }
            } elseif (strpos($status, '顧客情報に記入') !== false) {
                // 継続中: アポ日から10日以上経過 → 失注（継続）扱い
                $appoDate = strtotime($ts);
                $nowTs = time();
                if ($appoDate && ($nowTs - $appoDate) / 86400 >= 10) {
                    $memberData[$v2Name]['lostContinuing']++;
                } else {
                    $memberData[$v2Name]['continuing']++;
                }
            } elseif ($status === '失注') {
                if ($lfVal === '否決' || $cbsVal === '否決') {
                    $memberData[$v2Name]['lostByLfCbs']++;
                } else {
                    $memberData[$v2Name]['lost']++;
                }
            }

            // 属性カウント
            $shVal = isset($row[$COL_SEIKATSU_HOGO]) ? trim($row[$COL_SEIKATSU_HOGO]) : '';
            $ssVal = isset($row[$COL_SEISHIN_SHIKKAN]) ? trim($row[$COL_SEISHIN_SHIKKAN]) : '';
            $stVal = isset($row[$COL_SHOUGAISHA_TECHOU]) ? trim($row[$COL_SHOUGAISHA_TECHOU]) : '';
            if ($shVal !== '') $memberData[$v2Name]['seikatsuHogo']++;
            if ($ssVal !== '') $memberData[$v2Name]['seishinShikkan']++;
            if ($stVal !== '') $memberData[$v2Name]['shougaishaTechou']++;

            // ライフティ: Y列(24)に値がある or 支払方法②にライフがある場合
            if ($lfVal !== '' && $lfVal !== 'キャンセル') {
                $memberData[$v2Name]['lfApplied']++;
                if ($lfVal === '承認') $memberData[$v2Name]['lfApproved']++;
            } elseif (mb_strpos($payMethod2, 'ライフ') !== false) {
                $memberData[$v2Name]['lfApplied']++;
            }
            // CBS: Z列(25)に値がある場合
            if ($cbsVal !== '' && $cbsVal !== 'キャンセル') {
                $memberData[$v2Name]['cbsApplied']++;
                if ($cbsVal === '承認') $memberData[$v2Name]['cbsApproved']++;
            }
        }

        // --- 着金集計（当月アポの案件に紐づく全着金額を合計） ---
        // 成約➔CO / 成約➔キャンセル / 成約➔失注 は着金から除外
        if ($isCurrentMonth && !$isExcludedStatus) {
            $dealRevenue = 0;
            foreach ($paySlots as $slot) {
                list($dateCol, $amountCol, $methodCol) = $slot;
                if (count($row) <= $amountCol) continue;

                $payAmount = parseAmount(isset($row[$amountCol]) ? $row[$amountCol] : '');
                if ($payAmount > 0) {
                    $memberData[$v2Name]['revenue'] += $payAmount;
                    $dealRevenue += $payAmount;

                    $payMethod = isset($row[$methodCol]) ? trim($row[$methodCol]) : '';
                    if (mb_strpos($payMethod, 'ライフ') !== false || mb_strpos($payMethod, 'CBS') !== false) {
                        $memberData[$v2Name]['shinpan'] += $payAmount;
                    } else {
                        $memberData[$v2Name]['creditCard'] += $payAmount;
                    }

                    // 着金速報: 支払日付をパースして追加（日付なしの場合はアポ日にフォールバック）
                    $rawPayDate = isset($row[$dateCol]) ? trim($row[$dateCol]) : '';
                    $parsedDate = parsePayDate($rawPayDate, $year);
                    if (!$parsedDate) {
                        $parsedDate = parsePayDate($ts, $year);
                    }
                    if ($parsedDate && $payAmount > 0) {
                        $pd = explode('-', $parsedDate);
                        $paymentNews[] = [
                            'date' => $parsedDate,
                            'dateShort' => intval($pd[1]) . '/' . intval($pd[2]),
                            'name' => $v2Name,
                            'icon' => $iconMap[$v2Name] ?? '',
                            'amount' => round($payAmount, 1),
                        ];
                    }
                }
            }

            if ($dealRevenue > 0) {
                $memberData[$v2Name]['fundedDeals']++;
            }

            // CO案件の着金額
            $status2 = trim($row[11]);
            if (strpos($status2, '成約') !== false && strpos($status2, 'CO') !== false && $dealRevenue > 0) {
                $memberData[$v2Name]['coRevenue'] += $dealRevenue;
            }
        } elseif (!$isExcludedStatus) {
            // --- 過去成約の当月着金（前月以前のアポで、支払日付が当月のもの） ---
            $isPrevMonth = (strpos($ts, $prevMonthPrefix) === 0 || strpos($ts, $prevMonthPrefix2) === 0);
            foreach ($paySlots as $slot) {
                list($dateCol, $amountCol, $methodCol) = $slot;
                if (count($row) <= $amountCol) continue;

                $payDate = isset($row[$dateCol]) ? $row[$dateCol] : '';
                $payAmount = parseAmount(isset($row[$amountCol]) ? $row[$amountCol] : '');

                if ($payAmount > 0 && $isPayMonth($payDate, $ts)) {
                    // 着金速報: 過去アポの当月着金も追加
                    $parsedDate = parsePayDate($payDate, $year);
                    if ($parsedDate) {
                        $pd = explode('-', $parsedDate);
                        $paymentNews[] = [
                            'date' => $parsedDate,
                            'dateShort' => intval($pd[1]) . '/' . intval($pd[2]),
                            'name' => $v2Name,
                            'icon' => $iconMap[$v2Name] ?? '',
                            'amount' => round($payAmount, 1),
                        ];
                    }

                    if ($isPrevMonth) {
                        $memberData[$v2Name]['prevMonthRevenue'] += $payAmount;
                    } else {
                        $memberData[$v2Name]['pastRevenue'] += $payAmount;
                    }
                }
            }
        }
    }

    // メンバー配列に変換
    $members = [];
    $totalRevenue = 0;

    foreach ($memberData as $name => $d) {
        if (!isset($iconMap[$name])) continue;
        $closed = $d['closed'];
        // 商談数 = 成約 + 全失注（CO・否決・継続失注を含む）
        $totalLost = $d['lost'] + $d['lostByLfCbs'] + $d['lostContinuing'] + $d['coCount'];
        $deals = $closed + $totalLost;
        $sales = round($d['sales'], 1);
        $revenue = round($d['revenue'], 1);

        $members[] = [
            'name' => $name,
            'icon' => $iconMap[$name],
            'team' => $teamMap[$name] ?? 0,
            'revenue' => $revenue,
            'pastRevenue' => round($d['pastRevenue'], 1),
            'prevMonthRevenue' => round($d['prevMonthRevenue'], 1),
            'deals' => $deals,
            'closed' => $closed,
            'closeRate' => $deals > 0 ? round($closed / $deals * 100, 1) : 0,
            'coAmount' => round($d['coAmount'], 1),
            'coRevenue' => round($d['coRevenue'], 1),
            'creditCard' => round($d['creditCard'], 1),
            'shinpan' => round($d['shinpan'], 1),
            'avgPrice' => $closed > 0 ? round($sales / $closed, 1) : 0,
            'sales' => $sales,
            'fundedDeals' => $d['fundedDeals'],
            'cbs' => $d['cbsApproved'] . '/' . $d['cbsApplied'],
            'lifety' => $d['lfApproved'] . '/' . $d['lfApplied'],
            'closedOnclass' => $d['closedOnclass'],
            'closedConsul' => $d['closedConsul'],
            'coCount' => $d['coCount'],
            'continuing' => $d['continuing'],
            'lost' => $d['lost'],
            'lostByLfCbs' => $d['lostByLfCbs'],
            'lostContinuing' => $d['lostContinuing'],
            'seikatsuHogo' => $d['seikatsuHogo'],
            'seishinShikkan' => $d['seishinShikkan'],
            'shougaishaTechou' => $d['shougaishaTechou'],
            'prevRevenue' => 0, 'diffRevenue' => 0,
            'prevDeals' => 0, 'diffDeals' => 0,
            'prevClosed' => 0, 'diffClosed' => 0,
            'prevCloseRate' => 0, 'diffCloseRate' => 0,
        ];

        $totalRevenue += $revenue;
    }

    // ランキング（revenue降順）
    usort($members, function($a, $b) { return $b['revenue'] <=> $a['revenue']; });
    $lastRev = -1; $lastRank = 0;
    $topRev = !empty($members) ? $members[0]['revenue'] : 0;
    foreach ($members as $i => &$m) {
        $m['rank'] = ($m['revenue'] == $lastRev) ? $lastRank : $i + 1;
        $lastRev = $m['revenue'];
        $lastRank = $m['rank'];
        $m['gapToTop'] = round($topRev - $m['revenue'], 1);
    }
    unset($m);

    $now = new DateTime('now', new DateTimeZone('Asia/Tokyo'));
    $today = $now->format('Y-m-d');

    // 着金速報: 未来日付を除外してから日付降順・金額降順でソート
    $paymentNews = array_values(array_filter($paymentNews, function($n) use ($today) {
        return $n['date'] <= $today;
    }));
    usort($paymentNews, function($a, $b) {
        if ($a['date'] !== $b['date']) return strcmp($b['date'], $a['date']);
        return $b['amount'] <=> $a['amount'];
    });

    // dailyPushes: byMember を [{name, count}, ...] 形式に変換
    $dailyPushByMemberFormatted = [];
    foreach ($dailyPushByMember as $dateKey => $members_) {
        $arr = [];
        foreach ($members_ as $mName => $cnt) {
            $arr[] = ['name' => $mName, 'count' => $cnt];
        }
        usort($arr, function($a, $b) { return $b['count'] - $a['count']; });
        $dailyPushByMemberFormatted[$dateKey] = $arr;
    }

    return [
        'members' => $members,
        'teamNames' => $TEAM_NAMES,
        'totalRevenue' => round($totalRevenue, 1),
        'teamGoal' => 15000,
        'remaining' => 0,
        'progressRate' => 0,
        'dailyTarget' => 0,
        'daysLeft' => max(1, intval($now->format('t')) - intval($now->format('j'))),
        'currentMonth' => intval($now->format('n')),
        'paymentNews' => $paymentNews,
        'dailyPushes' => [
            'totals' => empty($dailyPushTotals) ? new \stdClass() : $dailyPushTotals,
            'byMember' => empty($dailyPushByMemberFormatted) ? new \stdClass() : $dailyPushByMemberFormatted,
        ],
        'updatedAt' => $now->format('Y/m/d H:i:s'),
    ];
}

// ===== 着金速報取得（action=apiからpaymentNewsのみ抽出） =====

function fetchPaymentNews() {
    global $GAS_URL, $LEGACY_NAME_MAP;
    $resp = gasRequest($GAS_URL . '?action=api');
    if (!$resp) return [];

    $data = json_decode($resp, true);
    if (!$data || empty($data['paymentNews'])) return [];

    $news = $data['paymentNews'];
    $today = date('Y-m-d', strtotime('now', strtotime('+9 hours'))); // JST
    $filtered = [];
    foreach ($news as &$n) {
        if (isset($n['date']) && $n['date'] > $today) continue;
        if (isset($LEGACY_NAME_MAP[$n['name']])) {
            $n['name'] = $LEGACY_NAME_MAP[$n['name']];
        }
        $filtered[] = $n;
    }
    unset($n);

    return $filtered;
}

// ===== 前月データ補完（GASアーカイブから） =====

function fillPrevMonthData(&$data) {
    if (empty($data['members'])) return;

    $curMonth = $data['currentMonth'];
    $curYear = intval(date('Y'));
    $prevMonth = $curMonth - 1;
    $prevYear = $curYear;
    if ($prevMonth < 1) { $prevMonth = 12; $prevYear--; }

    // マスターCSVから前月データ取得（GAS API廃止）
    $prevData = fetchFromMasterCSV($prevMonth, $prevYear);
    if (!$prevData || empty($prevData['members'])) return;

    $prevMap = [];
    foreach ($prevData['members'] as $pm) {
        $prevMap[$pm['name']] = $pm;
    }

    foreach ($data['members'] as &$m) {
        $prev = $prevMap[$m['name']] ?? null;
        if (!$prev) continue;

        $m['prevRevenue'] = round(floatval($prev['revenue'] ?? 0), 1);
        $m['diffRevenue'] = round($m['revenue'] - $m['prevRevenue'], 1);
        $m['prevDeals'] = intval($prev['deals'] ?? 0);
        $m['diffDeals'] = $m['deals'] - $m['prevDeals'];
        $m['prevClosed'] = intval($prev['closed'] ?? 0);
        $m['diffClosed'] = $m['closed'] - $m['prevClosed'];
        $m['prevCloseRate'] = round(floatval($prev['closeRate'] ?? 0), 1);
        $m['diffCloseRate'] = round($m['closeRate'] - $m['prevCloseRate'], 1);
        $m['prevFundedDeals'] = intval($prev['fundedDeals'] ?? 0);
        $m['prevSales'] = round(floatval($prev['sales'] ?? 0), 1);
        $m['prevAvgPrice'] = round(floatval($prev['avgPrice'] ?? 0), 1);
        $m['prevCoAmount'] = round(floatval($prev['coAmount'] ?? 0), 1);
        $m['prevCoRevenue'] = round(floatval($prev['coRevenue'] ?? 0), 1);
        $m['prevCreditCard'] = round(floatval($prev['creditCard'] ?? 0), 1);
        $m['prevShinpan'] = round(floatval($prev['shinpan'] ?? 0), 1);
    }
    unset($m);
}

// ===== ゴール設定 =====

function applyGoalSettings(&$data, $month, $year) {
    $settingsFile = __DIR__ . '/goal-settings-data.json';
    if (!file_exists($settingsFile)) return;

    $settings = json_decode(file_get_contents($settingsFile), true);
    if (!$settings) return;

    $m = $month ?: ($data['currentMonth'] ?? intval(date('n')));
    $y = $year ?: intval(date('Y'));
    $key = $y . '-' . $m;
    $monthGoal = $settings[$key] ?? null;

    if ($monthGoal && !empty($monthGoal['teamGoal']) && $monthGoal['teamGoal'] > 0) {
        $data['teamGoal'] = floatval($monthGoal['teamGoal']);
    }
    $data['memberKgi'] = $monthGoal['memberKgi'] ?? [];

    // 全月のゴール設定をフロントに返す（キー形式を統一: Y_M）
    $allGoals = [];
    foreach ($settings as $k => $v) {
        // "2026-4" → "2026_4" に変換
        $normalized = str_replace('-', '_', $k);
        $allGoals[$normalized] = $v;
    }
    $data['goalSettings'] = $allGoals;
}

// ===== 派生値再計算 =====

function recalculate(&$data) {
    $teamGoal = $data['teamGoal'] ?? 15000;
    $totalRevenue = $data['totalRevenue'] ?? 0;
    $remaining = max(0, round($teamGoal - $totalRevenue, 1));
    $daysLeft = $data['daysLeft'] ?? 1;

    $data['remaining'] = $remaining;
    $data['progressRate'] = $teamGoal > 0 ? round($totalRevenue / $teamGoal * 100) : 0;
    $data['dailyTarget'] = ($daysLeft > 0 && $remaining > 0) ? round($remaining / $daysLeft, 1) : 0;
}

// ===== 休日データ取得 =====

function fetchHolidayData($month, $year) {
    // 月別マップ取得（5月以降は新マップ、4月以前は既存）
    list($_, $iconMap, $__) = getMapsForMonth($month, $year);

    $holidayFile = __DIR__ . '/holiday-data.json';
    $empty = ['byDate' => new \stdClass(), 'counts' => []];

    if (!file_exists($holidayFile)) return $empty;

    $raw = json_decode(file_get_contents($holidayFile), true);
    if (!$raw) return $empty;

    // 月が一致するか確認（異なる月のデータは返さない）
    if (intval($raw['month'] ?? 0) !== $month || intval($raw['year'] ?? 0) !== $year) {
        return $empty;
    }

    // byDateにPHP側のアイコンURLを付与
    $byDate = [];
    foreach (($raw['byDate'] ?? []) as $dateKey => $entries) {
        $byDate[$dateKey] = [];
        foreach ($entries as $entry) {
            $name = $entry['name'] ?? '';
            $byDate[$dateKey][] = [
                'name' => $name,
                'type' => $entry['type'] ?? 'full',
                'icon' => $iconMap[$name] ?? '',
            ];
        }
    }

    return [
        'byDate' => empty($byDate) ? new \stdClass() : $byDate,
        'counts' => $raw['counts'] ?? [],
    ];
}

// ===== メインダッシュボードデータ取得 =====

function fetchDashboardData() {
    $now = new DateTime('now', new DateTimeZone('Asia/Tokyo'));
    $currentMonth = intval($now->format('n'));
    $currentYear = intval($now->format('Y'));

    // 1. マスターCSVから全データ取得
    $data = fetchFromMasterCSV($currentMonth, $currentYear);
    if (!$data) return null;

    // 2. 着金速報はマスターCSVから生成済み（fetchFromMasterCSV内で収集）

    // 3. 前月データ補完（GASアーカイブから）
    fillPrevMonthData($data);

    // 4. 共創pt（Chatwork投稿回数）
    $kyosoCounts = updateKyosoCounts();
    if (!empty($data['members'])) {
        foreach ($data['members'] as &$m) {
            $m['kyosoPt'] = $kyosoCounts[$m['name']] ?? 0;
        }
        unset($m);
    }

    // 5. ゴール設定 & 派生値
    applyGoalSettings($data, null, null);
    recalculate($data);

    // 6. 休日データ（GASトリガーから受信したJSON）
    $holidayData = fetchHolidayData($currentMonth, $currentYear);
    $data['holidayByDate'] = $holidayData['byDate'];
    if (!empty($data['members'])) {
        foreach ($data['members'] as &$m) {
            $m['holidays'] = $holidayData['counts'][$m['name']] ?? 0;
        }
        unset($m);
    }

    return $data;
}

// ===== アーカイブデータ取得（過去月） =====

function fetchArchiveData($month, $year) {
    // マスターCSVから過去月データを取得（GAS API廃止）
    $data = fetchFromMasterCSV(intval($month), intval($year));
    if (!$data) return null;

    applyGoalSettings($data, $month, $year);
    recalculate($data);

    return $data;
}

// ===== ルーティング =====

// POST処理
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $input = json_decode(file_get_contents('php://input'), true);
    $postAction = $input['action'] ?? '';

    // POST: ゴール設定保存
    if ($postAction === 'saveGoals') {
        $settingsFile = __DIR__ . '/goal-settings-data.json';
        $settings = file_exists($settingsFile) ? json_decode(file_get_contents($settingsFile), true) : [];
        if (!is_array($settings)) $settings = [];

        $key = $input['key'] ?? '';
        // フロント "2026_4" → サーバー "2026-4" に統一
        $key = str_replace('_', '-', $key);
        if ($key) {
            $settings[$key] = [
                'teamGoal' => floatval($input['teamGoal'] ?? 0),
                'memberKgi' => $input['memberKgi'] ?? [],
            ];
            file_put_contents($settingsFile, json_encode($settings, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT), LOCK_EX);
            // ゴール変更時はキャッシュクリア
            array_map('unlink', glob($CACHE_DIR . '/*.json'));
        }
        echo json_encode(['ok' => true], JSON_UNESCAPED_UNICODE);
        exit;
    }

    // POST: 休日データ更新（GASトリガーから1時間ごとに送信される）
    if ($postAction === 'updateHoliday') {
        if (($input['secret'] ?? '') !== 'gas_holiday_push_2026') {
            http_response_code(403);
            echo json_encode(['error' => 'unauthorized']);
            exit;
        }
        $holidayFile = __DIR__ . '/holiday-data.json';
        $holidayPayload = [
            'month' => intval($input['month'] ?? 0),
            'year' => intval($input['year'] ?? 0),
            'byDate' => $input['byDate'] ?? [],
            'counts' => $input['counts'] ?? [],
            'updatedAt' => date('Y-m-d H:i:s'),
        ];
        file_put_contents($holidayFile, json_encode($holidayPayload, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT), LOCK_EX);
        // キャッシュクリア
        array_map('unlink', glob($CACHE_DIR . '/*.json'));
        echo json_encode(['ok' => true, 'rows' => count($holidayPayload['byDate'])]);
        exit;
    }
}

$query = $_SERVER['QUERY_STRING'] ?? '';
parse_str($query, $params);
$action = $params['action'] ?? '';
$type = $params['type'] ?? '';

if ($action === 'api' && $type === '') {
    $month = $params['month'] ?? null;
    $year = $params['year'] ?? null;

    $now = new DateTime('now', new DateTimeZone('Asia/Tokyo'));
    $isArchive = ($month && $year)
        && !(intval($month) === intval($now->format('n')) && intval($year) === intval($now->format('Y')));
    $ttl = $isArchive ? $CACHE_TTL_ARCHIVE : $CACHE_TTL_LIVE;
    $cacheKey = 'v6_dashboard_' . ($isArchive ? $month . '_' . $year : 'live');
    $cacheFile = $CACHE_DIR . '/' . md5($cacheKey) . '.json';

    if (file_exists($cacheFile) && (time() - filemtime($cacheFile)) < $ttl) {
        readfile($cacheFile);
        exit;
    }

    $data = $isArchive ? fetchArchiveData($month, $year) : fetchDashboardData();
    if ($data === null) serveCache($cacheFile);

    $json = json_encode($data, JSON_UNESCAPED_UNICODE);
    writeCache($cacheFile, $json);
    echo $json;
    exit;
}

// type=months: マスターCSVからPHPで直接生成（GAS API廃止）
if ($action === 'api' && $type === 'months') {
    $cacheKey = 'v6_months';
    $cacheFile = $CACHE_DIR . '/' . md5($cacheKey) . '.json';

    if (file_exists($cacheFile) && (time() - filemtime($cacheFile)) < $CACHE_TTL_OTHER) {
        readfile($cacheFile);
        exit;
    }

    $rows = getMasterCsvRows();
    $now = new DateTime('now', new DateTimeZone('Asia/Tokyo'));
    $curMonth = intval($now->format('n'));
    $curYear = intval($now->format('Y'));

    if (!$rows) {
        $json = json_encode(['months' => [], 'current' => ['month' => $curMonth, 'year' => $curYear]], JSON_UNESCAPED_UNICODE);
        writeCache($cacheFile, $json);
        echo $json;
        exit;
    }

    $seen = [];
    $months = [];
    foreach ($rows as $idx => $row) {
        if ($idx === 0 || count($row) < 5) continue;
        $ts = $row[2] ?? '';
        if (preg_match('/^(\d{4})\/(\d{1,2})\//', $ts, $m)) {
            $key = intval($m[1]) . '-' . intval($m[2]);
            if (!isset($seen[$key])) {
                $seen[$key] = true;
                $months[] = ['month' => intval($m[2]), 'year' => intval($m[1])];
            }
        }
    }
    usort($months, function($a, $b) {
        if ($a['year'] !== $b['year']) return $a['year'] - $b['year'];
        return $a['month'] - $b['month'];
    });

    $json = json_encode(['months' => $months, 'current' => ['month' => $curMonth, 'year' => $curYear]], JSON_UNESCAPED_UNICODE);
    writeCache($cacheFile, $json);
    echo $json;
    exit;
}

// その他はGASにそのまま転送（フォールバック）
$cacheKey = 'v6_passthrough_' . md5($query);
$cacheFile = $CACHE_DIR . '/' . $cacheKey . '.json';

if (file_exists($cacheFile) && (time() - filemtime($cacheFile)) < $CACHE_TTL_OTHER) {
    readfile($cacheFile);
    exit;
}

$response = gasRequest($GAS_URL . ($query ? '?' . $query : ''));
if ($response === false) serveCache($cacheFile);

writeCache($cacheFile, $response);
echo $response;
