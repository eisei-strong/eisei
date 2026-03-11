<?php
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(200);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['error' => 'POST only']);
    exit;
}

$input = json_decode(file_get_contents('php://input'), true);

if (!$input || empty($input['password']) || $input['password'] !== 'hoshi2026') {
    http_response_code(403);
    echo json_encode(['error' => 'Unauthorized']);
    exit;
}

if (empty($input['data']) || !isset($input['data']['steps']) || !isset($input['data']['phases'])) {
    http_response_code(400);
    echo json_encode(['error' => 'Invalid data']);
    exit;
}

$configFile = __DIR__ . '/course_config.json';

// Backup
if (file_exists($configFile)) {
    $backupDir = __DIR__ . '/backups';
    if (!is_dir($backupDir)) mkdir($backupDir, 0755, true);
    copy($configFile, $backupDir . '/course_config_' . date('Ymd_His') . '.json');
}

$data = $input['data'];
$data['updatedAt'] = date('Y-m-d H:i:s');

if (file_put_contents($configFile, json_encode($data, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT)) === false) {
    http_response_code(500);
    echo json_encode(['error' => 'Write failed']);
    exit;
}

echo json_encode(['success' => true, 'updatedAt' => $data['updatedAt']]);
