<?php
// 営業ダッシュボード API プロキシ v5
// revenue: 旧シート(debugSheetByGid) = スプシの数式が正
// deals/closed: マスターCSV(フォーム回答) = GAS日別より正確
// paymentNews/prev月: GAS action=api
// syncFromOldSheetは走らせない
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Cache-Control: public, max-age=60');

// ===== 設定 =====
$GAS_URL = 'https://script.google.com/macros/s/AKfycbwojGHuvzycc07FJKwBdbBJJQZpssF6lYz0DbNJlu6zsVuXkAj8V8w3XNBPieo2wsYbFg/exec';
$MASTER_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1KxHeLmrpdaw1IUhBaQ46UWSHu-8SCRZqcrHOE2hMwDo/export?format=csv&gid=326094286';
$OLD_SHEET_GID = '1235299010';
$CACHE_DIR = __DIR__ . '/cache';
$CACHE_TTL_LIVE = 180;
$CACHE_TTL_ARCHIVE = 3600;
$CACHE_TTL_OTHER = 600;

if (!is_dir($CACHE_DIR)) mkdir($CACHE_DIR, 0755, true);

// 旧シートの列（メンバー位置）
$COLS = ['C','F','I','L','O','R','U','X','AA','AD'];

// 旧シート表示名の補正
$NAME_MAP = [
    'スクリプト通りに営業' => 'スクリプトくん',
    'スクリプト通りに営業するくん' => 'スクリプトくん',
    'ドライ' => 'ポジティブ',
    '勝友美' => 'ポジティブ',
];

// 本名 → v2メンバー名（マスターCSV用）
$REAL_NAME_MAP = [
    '阿部' => 'AをAでやる',
    '伊東' => 'ポジティブ',
    '久保田' => 'ヒトコト',
    '辻阪' => 'ビッグマウス',
    '五十嵐' => 'ぜんぶり',
    '新居' => 'スクリプトくん',
    '佐々木心雪' => 'ワントーン',
    '佐々木' => 'ワントーン',
    '福島' => 'けつだん',
    '大久保友佑悟' => 'ゴン',
    '大久保' => 'ゴン',
    '矢吹友一' => 'トニー',
    '矢吹' => 'トニー',
    'トニー' => 'トニー',
    '勝友美' => 'ポジティブ',
    '勝' => 'ポジティブ',
    'ドライ' => 'ポジティブ',
    '川合' => 'リヴァイ',
    '大内' => 'ガロウ',
];

// レガシー名 → 現在の名前（アーカイブの旧名マッピング）
$LEGACY_NAME_MAP = [
    'ドライ' => 'ポジティブ',
    '勝友美' => 'ポジティブ',
    'スクリプト通りに営業' => 'スクリプトくん',
    'スクリプト通りに営業するくん' => 'スクリプトくん',
    '李信' => 'AをAでやる',
    '流川' => 'ヒトコト',
    '首斬り桓騎' => 'ビッグマウス',
];

// アイコンマップ
$ICON_MAP = [
    'AをAでやる' => 'https://lh3.googleusercontent.com/d/10gj3l2D7PqGqZQZ1mmwyu4ZEZZPdem--',
    'ポジティブ' => 'https://lh3.googleusercontent.com/d/1AdF_IRXMi_uGG7ctCjO7CkaJw7XUOb6y',
    'トニー' => 'https://lh3.googleusercontent.com/d/1sHZ_zFFAitl7iVPEcIzQzpTD9cwL9FHv',
    'ヒトコト' => 'https://lh3.googleusercontent.com/d/14TcuxzbVRRVNSjhlaOFDXdCXke_jV7m3',
    'ゴン' => 'https://lh3.googleusercontent.com/d/1iwBxoCgXfmfOoUhTv4OUy7mir9XmvjJV',
    'ビッグマウス' => 'https://lh3.googleusercontent.com/d/13EV9ouH2X5tD7GzqfSTA2osSptFuzrqZ',
    'けつだん' => 'https://lh3.googleusercontent.com/d/1wnoxiF7PRZKSFPnjn0WXQb16hm-68Jlk',
    'ぜんぶり' => 'https://lh3.googleusercontent.com/d/11_mTOKu5m2MFoufn36NjUQyLjOdXrpa5',
    'スクリプトくん' => 'https://lh3.googleusercontent.com/d/1BSBMs3h5BgC1z0Tx8jyprPmDs11LTBPn',
    'ワントーン' => 'https://lh3.googleusercontent.com/d/1tTXYHdXlPELox3hwXpAUnBCG_w2shjIQ',
    'リヴァイ' => '',
    'ガロウ' => '',
];

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

function resolveV2Name($rawField) {
    global $REAL_NAME_MAP;
    $rawPerson = explode('：', str_replace(':', '：', $rawField))[0];
    $rawPerson = trim($rawPerson);
    if (isset($REAL_NAME_MAP[$rawPerson])) return $REAL_NAME_MAP[$rawPerson];
    foreach ($REAL_NAME_MAP as $key => $val) {
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

// ===== 旧シートからメンバーデータ構築（debugSheetByGid、sync不要） =====

function fetchFromOldSheet() {
    global $GAS_URL, $OLD_SHEET_GID, $COLS, $NAME_MAP, $ICON_MAP;

    $debugUrl = $GAS_URL . '?action=run&fn=debugSheetByGid&gid=' . $OLD_SHEET_GID
        . '&rows=4,5,6,7,8,9,10,11,12,13,15,17,18,20,21,25,29,30&cols=C,F,I,L,O,R,U,X,AA,AD,AG';
    $resp = gasRequest($debugUrl);
    if ($resp === false) return null;

    $debugData = json_decode($resp, true);
    $rows = $debugData['result']['data'] ?? [];
    if (empty($rows)) return null;

    $row4 = $rows['row4'] ?? [];
    $members = [];

    // AG6 = チーム合計着金額
    $totalRev = round(floatval($rows['row6']['AG'] ?? 0), 1);

    foreach ($COLS as $c) {
        $raw = str_replace(['【','】'], '', $row4[$c] ?? '');
        $name = $NAME_MAP[$raw] ?? $raw;
        if (empty($name)) continue;

        $revenue = round(floatval($rows['row6'][$c] ?? 0), 1);

        $members[] = [
            'name' => $name,
            'icon' => $ICON_MAP[$name] ?? '',
            'revenue' => $revenue,
            'deals' => 0,     // マスターCSVで上書き
            'closed' => 0,    // マスターCSVで上書き
            'closeRate' => 0, // マスターCSVで再計算
            'coAmount' => round(floatval($rows['row30'][$c] ?? 0), 1),
            'coRevenue' => intval($rows['row29'][$c] ?? 0),
            'creditCard' => round(floatval($rows['row10'][$c] ?? 0), 1),
            'shinpan' => round(floatval($rows['row11'][$c] ?? 0), 1),
            'avgPrice' => round(floatval($rows['row12'][$c] ?? 0), 1),
            'sales' => round(floatval($rows['row13'][$c] ?? 0), 1),
            'fundedDeals' => intval($rows['row25'][$c] ?? 0),
            'cbs' => '-',
            'lifety' => '-',
            'closedOnclass' => 0,
            'closedConsul' => 0,
            'lost' => 0,
            'lostByLfCbs' => 0,
            'prevRevenue' => 0, 'diffRevenue' => 0,
            'prevDeals' => 0, 'diffDeals' => 0,
            'prevClosed' => 0, 'diffClosed' => 0,
            'prevCloseRate' => 0, 'diffCloseRate' => 0,
        ];
    }

    // ランキング（revenue降順）
    usort($members, function($a, $b) { return $b['revenue'] <=> $a['revenue']; });
    $lastRev = -1; $lastRank = 0;
    $topRev = $members[0]['revenue'] ?? 0;
    foreach ($members as $i => &$m) {
        $m['rank'] = ($m['revenue'] == $lastRev) ? $lastRank : $i + 1;
        $lastRev = $m['revenue'];
        $lastRank = $m['rank'];
        $m['gapToTop'] = round($topRev - $m['revenue'], 1);
    }
    unset($m);

    $now = new DateTime('now', new DateTimeZone('Asia/Tokyo'));

    return [
        'members' => $members,
        'totalRevenue' => $totalRev,
        'teamGoal' => 15000,
        'remaining' => 0,
        'progressRate' => 0,
        'dailyTarget' => 0,
        'daysLeft' => max(1, intval($now->format('t')) - intval($now->format('j'))),
        'currentMonth' => intval($now->format('n')),
        'paymentNews' => [],
        'updatedAt' => $now->format('Y/m/d H:i:s'),
    ];
}

// ===== マスターCSVから商談数・成約内訳・失注・CBS/LFを取得 =====

function fetchMasterSupplement($currentMonth, $currentYear) {
    global $MASTER_SHEET_URL;
    $csv = gasRequest($MASTER_SHEET_URL);
    if (!$csv) return [];

    $rows = parseCsv($csv);
    $monthPrefix = $currentYear . '/' . $currentMonth;
    $monthPrefix2 = $currentYear . '/' . str_pad($currentMonth, 2, '0', STR_PAD_LEFT);
    $supplement = [];

    foreach ($rows as $idx => $row) {
        if ($idx === 0 || count($row) < 12) continue;
        $ts = isset($row[2]) ? $row[2] : $row[0];
        if (strpos($ts, $monthPrefix) !== 0 && strpos($ts, $monthPrefix2) !== 0) continue;

        $v2Name = resolveV2Name($row[3] ?? '');
        if (!$v2Name) continue;

        if (!isset($supplement[$v2Name])) {
            $supplement[$v2Name] = [
                'deals' => 0, 'closed' => 0,
                'closedOnclass' => 0, 'closedConsul' => 0,
                'lost' => 0, 'lostByLfCbs' => 0,
                'cbsApproved' => 0, 'cbsApplied' => 0,
                'lfApproved' => 0, 'lfApplied' => 0,
            ];
        }

        $status = trim($row[11]);
        $product = isset($row[12]) ? trim($row[12]) : '';
        $lfVal = isset($row[24]) ? str_replace('✅', '', trim($row[24])) : '';
        $cbsVal = isset($row[25]) ? str_replace('✅', '', trim($row[25])) : '';
        $payMethod2 = isset($row[20]) ? trim($row[20]) : '';

        // 商談数カウント（statusが空でなければ全てカウント）
        if ($status !== '') {
            $supplement[$v2Name]['deals']++;
        }

        // 成約内訳
        if ($status === '成約' || (strpos($status, '成約') !== false)) {
            $supplement[$v2Name]['closed']++;
            if (mb_strpos($product, 'オンクラス') !== false) {
                $supplement[$v2Name]['closedOnclass']++;
            } else {
                $supplement[$v2Name]['closedConsul']++;
            }
        } elseif ($status === '失注') {
            if ($lfVal === '否決' || $cbsVal === '否決') {
                $supplement[$v2Name]['lostByLfCbs']++;
            } else {
                $supplement[$v2Name]['lost']++;
            }
        }

        // ライフティ（母数はidx20の支払方法2で判定）
        if (mb_strpos($payMethod2, 'ライフ') !== false) {
            $supplement[$v2Name]['lfApplied']++;
            if ($lfVal === '承認') $supplement[$v2Name]['lfApproved']++;
        }

        // CBS
        if ($cbsVal !== '' && $cbsVal !== 'キャンセル') {
            $supplement[$v2Name]['cbsApplied']++;
            if ($cbsVal === '承認') $supplement[$v2Name]['cbsApproved']++;
        }
    }

    return $supplement;
}

// ===== 着金速報取得（action=apiからpaymentNewsのみ抽出） =====

function fetchPaymentNews() {
    global $GAS_URL, $LEGACY_NAME_MAP;
    $resp = gasRequest($GAS_URL . '?action=api');
    if (!$resp) return [];

    $data = json_decode($resp, true);
    if (!$data || empty($data['paymentNews'])) return [];

    $news = $data['paymentNews'];
    foreach ($news as &$n) {
        if (isset($LEGACY_NAME_MAP[$n['name']])) {
            $n['name'] = $LEGACY_NAME_MAP[$n['name']];
        }
    }
    unset($n);

    return $news;
}

// ===== 前月データ補完（GASアーカイブから） =====

function fillPrevMonthData(&$data) {
    global $GAS_URL, $LEGACY_NAME_MAP;
    if (empty($data['members'])) return;

    $curMonth = $data['currentMonth'];
    $curYear = intval(date('Y'));
    $prevMonth = $curMonth - 1;
    $prevYear = $curYear;
    if ($prevMonth < 1) { $prevMonth = 12; $prevYear--; }

    $prevResp = gasRequest($GAS_URL . '?action=api&month=' . $prevMonth . '&year=' . $prevYear);
    if (!$prevResp) return;

    $prevData = json_decode($prevResp, true);
    if (!$prevData || empty($prevData['members'])) return;

    $prevMap = [];
    foreach ($prevData['members'] as $pm) {
        $name = $LEGACY_NAME_MAP[$pm['name']] ?? $pm['name'];
        $prevMap[$name] = $pm;
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

// ===== 補足データマージ =====

function mergeSupplement(&$data, $supplement) {
    if (empty($supplement) || empty($data['members'])) return;

    foreach ($data['members'] as &$m) {
        $s = $supplement[$m['name']] ?? null;
        if ($s) {
            if ($s['deals'] > 0) {
                $m['deals'] = $s['deals'];
                $m['closed'] = $s['closed'];
                $m['closeRate'] = $m['deals'] > 0 ? round($m['closed'] / $m['deals'] * 100, 1) : 0;
            }
            $m['closedOnclass'] = $s['closedOnclass'];
            $m['closedConsul'] = $s['closedConsul'];
            $m['lost'] = $s['lost'];
            $m['lostByLfCbs'] = $s['lostByLfCbs'];
            $m['cbs'] = $s['cbsApproved'] . '/' . $s['cbsApplied'];
            $m['lifety'] = $s['lfApproved'] . '/' . $s['lfApplied'];
        }
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

// ===== メインダッシュボードデータ取得 =====

function fetchDashboardData() {
    // 1. 旧シートからrevenue等を取得（debugSheetByGid、syncなし）
    $data = fetchFromOldSheet();
    if (!$data) return null;

    $currentMonth = $data['currentMonth'];
    $currentYear = intval(date('Y'));

    // 2. マスターCSVでdeals/closed/成約内訳/失注/CBS/LFを上書き
    $supplement = fetchMasterSupplement($currentMonth, $currentYear);
    mergeSupplement($data, $supplement);

    // 3. 着金速報を取得（action=apiからpaymentNewsのみ）
    $data['paymentNews'] = fetchPaymentNews();

    // 4. 前月データ補完（GASアーカイブから）
    fillPrevMonthData($data);

    // 5. ゴール設定 & 派生値
    applyGoalSettings($data, null, null);
    recalculate($data);

    return $data;
}

// ===== アーカイブデータ取得（過去月） =====

function fetchArchiveData($month, $year) {
    global $GAS_URL, $LEGACY_NAME_MAP;

    $resp = gasRequest($GAS_URL . '?action=api&month=' . intval($month) . '&year=' . intval($year));
    if (!$resp) return null;

    $data = json_decode($resp, true);
    if (!$data || isset($data['error'])) return null;

    if (!empty($data['members'])) {
        foreach ($data['members'] as &$m) {
            if (isset($LEGACY_NAME_MAP[$m['name']])) {
                $m['name'] = $LEGACY_NAME_MAP[$m['name']];
            }
        }
        unset($m);
    }

    applyGoalSettings($data, $month, $year);
    recalculate($data);

    return $data;
}

// ===== ルーティング =====

$query = $_SERVER['QUERY_STRING'] ?? '';
parse_str($query, $params);
$action = $params['action'] ?? '';
$type = $params['type'] ?? '';

if ($action === 'api' && $type === '') {
    $month = $params['month'] ?? null;
    $year = $params['year'] ?? null;

    // 当月リクエストはlive扱い（旧シートから取得）
    $now = new DateTime('now', new DateTimeZone('Asia/Tokyo'));
    $isArchive = ($month && $year)
        && !(intval($month) === intval($now->format('n')) && intval($year) === intval($now->format('Y')));
    $ttl = $isArchive ? $CACHE_TTL_ARCHIVE : $CACHE_TTL_LIVE;
    $cacheKey = 'v5_dashboard_' . ($isArchive ? $month . '_' . $year : 'live');
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

// その他（type=months, action=run等）はGASにそのまま転送
$cacheKey = 'v5_passthrough_' . md5($query);
$cacheFile = $CACHE_DIR . '/' . $cacheKey . '.json';

if (file_exists($cacheFile) && (time() - filemtime($cacheFile)) < $CACHE_TTL_OTHER) {
    readfile($cacheFile);
    exit;
}

$response = gasRequest($GAS_URL . ($query ? '?' . $query : ''));
if ($response === false) serveCache($cacheFile);

writeCache($cacheFile, $response);
echo $response;
