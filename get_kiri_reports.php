<?php
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');

$historyFile = __DIR__ . '/qa_history.json';
if (!file_exists($historyFile)) {
    echo json_encode(['reports' => [], 'grouped' => []]);
    exit;
}

$qaHistory = json_decode(file_get_contents($historyFile), true) ?: [];

$filterType = $_GET['type'] ?? 'kirigaeshi';

$reports = [];
$grouped = [];
foreach ($qaHistory as $entry) {
    if (($entry['category'] ?? '') === $filterType && ($entry['type'] ?? '') === 'chat') {
        // Skip duplicates flagged by AI
        if ($filterType === 'kirigaeshi' && !empty($entry['kiriDuplicate'])) continue;

        $genre = $entry['kiriGenre'] ?? '未分類';
        $r = [
            'id' => $entry['id'] ?? '',
            'userName' => $entry['userName'] ?? '匿名',
            'question' => $entry['question'] ?? '',
            'answer' => $entry['answer'] ?? '',
            'answeredBy' => $entry['answeredBy'] ?? 'AI',
            'genre' => $genre,
            'timestamp' => $entry['timestamp'] ?? '',
        ];
        $reports[] = $r;
        if (!isset($grouped[$genre])) $grouped[$genre] = [];
        $grouped[$genre][] = $r;
    }
}

echo json_encode([
    'reports' => array_reverse($reports),
    'grouped' => $grouped,
], JSON_UNESCAPED_UNICODE);
