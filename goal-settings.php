<?php
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');
header('Cache-Control: no-store, no-cache, must-revalidate, max-age=0');
header('Pragma: no-cache');
header('Expires: 0');
header('X-Accel-Expires: 0');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(204);
    exit;
}

$file = __DIR__ . '/goal-settings-data.json';

if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    if (file_exists($file)) {
        echo file_get_contents($file);
    } else {
        echo '{}';
    }
    exit;
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $input = file_get_contents('php://input');
    $data = json_decode($input, true);
    if ($data === null && json_last_error() !== JSON_ERROR_NONE) {
        http_response_code(400);
        echo json_encode(['error' => 'Invalid JSON: ' . json_last_error_msg()]);
        exit;
    }
    $json = json_encode($data, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT | JSON_FORCE_OBJECT);
    $result = file_put_contents($file, $json, LOCK_EX);
    if ($result === false) {
        http_response_code(500);
        echo json_encode(['error' => 'Write failed', 'file' => $file, 'writable' => is_writable($file)]);
        exit;
    }
    echo json_encode(['ok' => true, 'bytes' => $result]);
    exit;
}

http_response_code(405);
echo json_encode(['error' => 'Method not allowed']);
