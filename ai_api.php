<?php
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') { http_response_code(200); exit; }
if ($_SERVER['REQUEST_METHOD'] !== 'POST') { http_response_code(405); echo json_encode(['error'=>'POST only']); exit; }

$input = json_decode(file_get_contents('php://input'), true);
if (!$input) { http_response_code(400); echo json_encode(['error'=>'Invalid JSON']); exit; }

$apiKeyFile = '/home/kodaidai/.claude_api_key';
if (!file_exists($apiKeyFile)) { http_response_code(500); echo json_encode(['error'=>'API key not configured']); exit; }
$apiKey = trim(file_get_contents($apiKeyFile));

$type = $input['type'] ?? '';       // 'grade' or 'chat'
$userName = $input['userName'] ?? '匿名';
$message = $input['message'] ?? '';
$stepTitle = $input['stepTitle'] ?? '';
$stepQuestion = $input['question'] ?? '';
$history = $input['history'] ?? [];  // chat history

if (!$message) { http_response_code(400); echo json_encode(['error'=>'message required']); exit; }

// Load course config for context
$configFile = __DIR__ . '/course_config.json';
$courseContext = '';
if (file_exists($configFile)) {
    $cfg = json_decode(file_get_contents($configFile), true);
    if ($cfg && !empty($cfg['steps'])) {
        $stepList = [];
        foreach ($cfg['steps'] as $s) {
            $stepList[] = '- ' . ($s['title'] ?? '');
        }
        $courseContext = "【講座カリキュラム】\n" . implode("\n", $stepList);
    }
}

// Build system prompt
if ($type === 'grade') {
    $systemPrompt = <<<PROMPT
あなたは「30日間で300万円を着金する講座」の営業コーチです。
受講生のテスト回答を評価し、具体的なフィードバックを返してください。

【評価基準】
- S（完璧）: 台本の内容を正確に再現でき、自分の言葉で説明もできている
- A（だいたい）: 大筋は合っているが、細部の抜けや曖昧な点がある
- B（うっすら）: 方向性は理解しているが、具体性に欠ける
- C（覚えてない）: 内容の理解が不十分、または的外れ

【営業の基本方針】
- 失敗を恐れずにまず行動する（プールに飛び込む）
- プッシュ構成: 自己紹介→合意形成→理想ヒアリング→現在地ヒアリング→TikTok稼ぎ方→契約・決済
- 逆算思考: 300万÷単価50万=6件成約、成約率30%→20プッシュ最低、余裕持って30プッシュ
- Day 5までに本番プッシュデビュー

{$courseContext}

【回答フォーマット】
まず評価（S/A/B/C）を1行目に書き、その後に具体的なフィードバックを簡潔に書いてください。
良い点と改善点を明確にしてください。厳しくも温かいコーチングを心がけてください。
PROMPT;

    $userMessage = "【ステップ】{$stepTitle}\n【質問】{$stepQuestion}\n\n【受講生の回答】\n{$message}";

} elseif ($type === 'chat' && ($input['category'] ?? '') === 'kirigaeshi') {
    // Load existing kirigaeshi scripts for context
    $kiriContext = '';
    if (file_exists($configFile)) {
        $cfg = json_decode(file_get_contents($configFile), true);
        if ($cfg && !empty($cfg['steps'])) {
            $kiriList = [];
            foreach ($cfg['steps'] as $s) {
                if (($s['phase'] ?? '') === 'kirigaeshi') {
                    $title = str_replace('の切り返しを覚えた', '', $s['title'] ?? '');
                    $kiriList[] = '- ' . $title . ($s['desc'] ? "\n  スクリプト: " . mb_substr($s['desc'], 0, 200) : '');
                }
            }
            if ($kiriList) {
                $kiriContext = "【既存の切り返しスクリプト一覧】\n" . implode("\n", $kiriList);
            }
        }
    }

    // Load existing reports for duplicate check
    $existingReports = [];
    $historyFileCheck = __DIR__ . '/qa_history.json';
    if (file_exists($historyFileCheck)) {
        $existingQA = json_decode(file_get_contents($historyFileCheck), true) ?: [];
        foreach ($existingQA as $qa) {
            if (($qa['category'] ?? '') === 'kirigaeshi' && ($qa['type'] ?? '') === 'chat') {
                $existingReports[] = '- [' . ($qa['kiriGenre'] ?? '未分類') . '] ' . mb_substr($qa['question'] ?? '', 0, 100);
            }
        }
    }
    $existingReportContext = '';
    if ($existingReports) {
        $existingReportContext = "\n\n【過去に報告された質問一覧（重複チェック用）】\n" . implode("\n", array_slice($existingReports, -50));
    }

    $systemPrompt = <<<PROMPT
あなたは営業の切り返しスクリプトの「素案」を作成するAIアシスタントです。
受講生が商談中に遭遇した反論・断り文句に対する切り返しトークの素案を書いてください。

【重要】
- これは素案（ドラフト）です。最終版はコーチ「ほし」が書き直します。
- 既存の切り返しスクリプトと重複する内容であれば、その旨を伝えてください。
- ニュアンスが違うだけで本質的に同じ質問は「重複」として扱ってください。

【切り返しの基本方針】
- 相手の不安や反論を否定しない。まず受け止める
- 共感してから、別の視点を提示する
- 具体的な事例や数字を使って説得力を出す
- 最終的に「行動しない理由」を「行動する理由」に変換する
- 短く、実際の会話で使える自然な言い回しにする

{$kiriContext}
{$existingReportContext}

【回答フォーマット（必ずこの形式で回答）】
1行目: 【ジャンル】（以下から1つ選択: 料金・費用, 時間・忙しい, 信頼性・実績, 副業・会社, 家族・相談, 自信・不安, 競合・比較, 契約・決済, その他）
2行目: 【重複】なし or 【重複】あり（類似: 「〇〇」）

3行目以降:
【反論】（相手の言葉）
【切り返し素案】
（実際に話すトークスクリプト。会話調で書く）

※これはAIの素案です。コーチが最終確認・書き直しを行います。
PROMPT;

    $userMessage = $message;

} else {
    $systemPrompt = <<<PROMPT
あなたは「30日間で300万円を着金する講座」の営業コーチ「ほし」です。
受講生からの質問に対して、講座の内容に基づいて的確に回答してください。

【営業の基本方針】
- 失敗を恐れずにまず行動する（プールに飛び込む）
- 不安=挑戦=成長。不安な状態は正常
- プッシュ構成: 自己紹介→合意形成→理想ヒアリング→現在地ヒアリング→TikTok稼ぎ方→契約・決済
- 記憶術: 見ないで話す→詰まったところだけ見てまた話す、の繰り返し
- 練習プッシュ→資金なしプッシュ3件でデビュー→本番プッシュ
- 逆算: 300万÷50万=6件成約、成約率30%→20-30プッシュ必要
- Day 5までにデビュー

{$courseContext}

【コーチングスタイル】
- 簡潔かつ具体的に回答
- 受講生のモチベーションを高める
- 「まずやってみろ」のスタンス
- 分からないことは正直に伝える
PROMPT;

    // Search existing Q&A for similar questions
    $similarContext = '';
    $historyFile = __DIR__ . '/qa_history.json';
    if (file_exists($historyFile)) {
        $existingQA = json_decode(file_get_contents($historyFile), true) ?: [];
        $similar = [];
        $keywords = mb_strtolower($message);
        foreach ($existingQA as $qa) {
            if ($qa['type'] !== 'chat' || empty($qa['answer'])) continue;
            $qLower = mb_strtolower($qa['question'] ?? '');
            // Simple keyword matching
            $words = preg_split('/[\s　、。？！]+/u', $keywords);
            $matchCount = 0;
            foreach ($words as $w) {
                if (mb_strlen($w) >= 2 && mb_strpos($qLower, $w) !== false) $matchCount++;
            }
            if ($matchCount >= 2 || (mb_strlen($keywords) > 5 && mb_strpos($qLower, $keywords) !== false)) {
                $similar[] = $qa;
            }
        }
        if (!empty($similar)) {
            $similarContext = "\n\n【過去の類似Q&A（参考にしてください。的確な回答があればそのまま活用してOK）】\n";
            foreach (array_slice($similar, -5) as $sq) {
                $similarContext .= "Q: " . ($sq['question'] ?? '') . "\n";
                $similarContext .= "A (" . ($sq['answeredBy'] ?? 'AI') . "): " . mb_substr($sq['answer'] ?? '', 0, 300) . "\n\n";
            }
        }
    }

    $systemPrompt .= $similarContext;
    $userMessage = $message;
}

// Build messages array
$messages = [];
if ($type === 'chat' && !empty($history)) {
    foreach (array_slice($history, -10) as $h) {
        $messages[] = ['role' => $h['role'], 'content' => $h['content']];
    }
}
$messages[] = ['role' => 'user', 'content' => $userMessage];

// Call Claude API
$payload = [
    'model' => 'claude-sonnet-4-20250514',
    'max_tokens' => 1024,
    'system' => $systemPrompt,
    'messages' => $messages,
];

$ch = curl_init('https://api.anthropic.com/v1/messages');
curl_setopt_array($ch, [
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_POST => true,
    CURLOPT_POSTFIELDS => json_encode($payload),
    CURLOPT_HTTPHEADER => [
        'Content-Type: application/json',
        'x-api-key: ' . $apiKey,
        'anthropic-version: 2023-06-01',
    ],
    CURLOPT_TIMEOUT => 30,
]);

$response = curl_exec($ch);
$httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
curl_close($ch);

if ($httpCode !== 200) {
    http_response_code(502);
    echo json_encode(['error' => 'API error', 'status' => $httpCode, 'detail' => $response]);
    exit;
}

$result = json_decode($response, true);
$aiAnswer = '';
if (!empty($result['content'])) {
    foreach ($result['content'] as $block) {
        if ($block['type'] === 'text') {
            $aiAnswer .= $block['text'];
        }
    }
}

// Extract grade if grading
$grade = null;
if ($type === 'grade') {
    if (preg_match('/\*{0,2}[【\[]?\s*([SABC])\s*[（(]?/m', $aiAnswer, $m)) {
        $grade = $m[1];
    }
}

// Extract kiriGenre and duplicate status
$kiriGenre = null;
$kiriDuplicate = false;
if ($type === 'chat' && ($input['category'] ?? '') === 'kirigaeshi') {
    if (preg_match('/【ジャンル】\s*(.+)/u', $aiAnswer, $gm)) {
        $kiriGenre = trim($gm[1]);
    }
    if (preg_match('/【重複】\s*あり/u', $aiAnswer)) {
        $kiriDuplicate = true;
    }
}

// Save to history
$historyFile = __DIR__ . '/qa_history.json';
$qaHistory = [];
if (file_exists($historyFile)) {
    $qaHistory = json_decode(file_get_contents($historyFile), true) ?: [];
}

// Detect category
$categoryMap = [
    'mindset'  => ['不安','失敗','挑戦','成長','メンタル','モチベ','やる気','怖い','自信'],
    'overview' => ['構成','全体','マインドマップ','流れ','順番','ステップ'],
    'intro'    => ['自己紹介','合意形成','アイスブレイク','挨拶','最初'],
    'hearing'  => ['ヒアリング','理想','現在地','質問','聞き'],
    'tiktok'   => ['TikTok','tiktok','稼ぎ方','プレゼン','提案','説明'],
    'closing'  => ['契約','決済','クロージング','サイン','支払','クレジット'],
    'push'     => ['プッシュ','練習','資金なし','本番','デビュー','商談','アポ','録画'],
    'script'   => ['台本','スクリプト','記憶','覚え','暗記'],
    'kirigaeshi' => ['切り返し','反論','断り','断られ','言い返','対処','こう言われ'],
    'jimu'       => ['事務','契約書','テンプレ','手続き','書類','送付','振込','請求','領収'],
];
$detectedCat = $input['category'] ?? '';
if (!$detectedCat || !isset($categoryMap[$detectedCat])) {
    $detectedCat = 'other';
    foreach ($categoryMap as $catId => $keywords) {
        foreach ($keywords as $kw) {
            if (mb_stripos($message, $kw) !== false) { $detectedCat = $catId; break 2; }
        }
    }
}

// Count frequency of similar questions
$freq = 1;
foreach ($qaHistory as $qa) {
    if ($qa['type'] === 'chat' && isset($qa['category']) && $qa['category'] === $detectedCat) {
        $qLower = mb_strtolower($qa['question'] ?? '');
        $mLower = mb_strtolower($message);
        $words = preg_split('/[\s　、。？！]+/u', $mLower);
        $matchCount = 0;
        foreach ($words as $w) {
            if (mb_strlen($w) >= 2 && mb_strpos($qLower, $w) !== false) $matchCount++;
        }
        if ($matchCount >= 2) $freq++;
    }
}

$entry = [
    'id' => uniqid(),
    'type' => $type,
    'userName' => $userName,
    'question' => $type === 'grade' ? "[{$stepTitle}] {$message}" : $message,
    'answer' => $aiAnswer,
    'answeredBy' => 'AI',
    'grade' => $grade,
    'category' => $detectedCat,
    'freq' => $freq,
    'kiriGenre' => $kiriGenre,
    'kiriDuplicate' => $kiriDuplicate ?? false,
    'timestamp' => date('Y-m-d H:i:s'),
];
$qaHistory[] = $entry;

// Keep last 1000 entries
if (count($qaHistory) > 1000) {
    $qaHistory = array_slice($qaHistory, -1000);
}

file_put_contents($historyFile, json_encode($qaHistory, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT));

echo json_encode([
    'success' => true,
    'answer' => $aiAnswer,
    'grade' => $grade,
    'answeredBy' => 'AI',
    'id' => $entry['id'],
    'kiriGenre' => $kiriGenre,
    'kiriDuplicate' => $kiriDuplicate ?? false,
], JSON_UNESCAPED_UNICODE);
