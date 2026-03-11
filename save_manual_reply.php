<?php
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') { http_response_code(200); exit; }
if ($_SERVER['REQUEST_METHOD'] !== 'POST') { http_response_code(405); echo json_encode(['error'=>'POST only']); exit; }

$input = json_decode(file_get_contents('php://input'), true);

if (!$input || empty($input['password']) || $input['password'] !== 'hoshi2026') {
    http_response_code(403);
    echo json_encode(['error' => 'Unauthorized']);
    exit;
}

$qaId = $input['qaId'] ?? '';
$answer = $input['answer'] ?? '';
$answeredBy = $input['answeredBy'] ?? 'ほし';

if (!$qaId || !$answer) {
    http_response_code(400);
    echo json_encode(['error' => 'qaId and answer required']);
    exit;
}

$historyFile = __DIR__ . '/qa_history.json';
if (!file_exists($historyFile)) {
    http_response_code(404);
    echo json_encode(['error' => 'No history file']);
    exit;
}

$qaHistory = json_decode(file_get_contents($historyFile), true) ?: [];

$found = false;
foreach ($qaHistory as &$entry) {
    if ($entry['id'] === $qaId) {
        $entry['answer'] = $answer;
        $entry['answeredBy'] = $answeredBy;
        $entry['manualAt'] = date('Y-m-d H:i:s');
        $found = true;
        break;
    }
}
unset($entry);

if (!$found) {
    http_response_code(404);
    echo json_encode(['error' => 'Q&A entry not found']);
    exit;
}

file_put_contents($historyFile, json_encode($qaHistory, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT));

echo json_encode(['success' => true]);
