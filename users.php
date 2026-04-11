<?php
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') { http_response_code(200); exit; }
if ($_SERVER['REQUEST_METHOD'] !== 'POST') { http_response_code(405); echo json_encode(['error'=>'POST only']); exit; }

$input = json_decode(file_get_contents('php://input'), true);
if (!$input) { http_response_code(400); echo json_encode(['error'=>'Invalid JSON']); exit; }

$usersFile = __DIR__ . '/users.json';
$users = [];
if (file_exists($usersFile)) {
    $users = json_decode(file_get_contents($usersFile), true) ?: [];
}

// Seed initial users if file doesn't exist
if (!file_exists($usersFile) || empty($users)) {
    $seed = [
        ['userId'=>'hoshi','name'=>'ほし','pw'=>'hoshi2025'],
        ['userId'=>'dry','name'=>'ドライ','pw'=>'dry2025'],
        ['userId'=>'aaa','name'=>'AをAでやる','pw'=>'a2025'],
        ['userId'=>'hitokoto','name'=>'ヒトコト','pw'=>'hitokoto2025'],
        ['userId'=>'big','name'=>'ビッグマウス','pw'=>'big2025'],
        ['userId'=>'zen','name'=>'ぜんぶり','pw'=>'zen2025'],
        ['userId'=>'script','name'=>'スクリプト','pw'=>'script2025'],
        ['userId'=>'one','name'=>'ワントーン','pw'=>'one2025'],
        ['userId'=>'ketsu','name'=>'けつだん','pw'=>'ketsu2025'],
        ['userId'=>'tony','name'=>'トニー','pw'=>'tony2025'],
        ['userId'=>'gon','name'=>'ゴン','pw'=>'gon2025'],
        ['userId'=>'kodai','name'=>'こだい','pw'=>'kodai2025'],
        ['userId'=>'test','name'=>'テスト','pw'=>'test2025'],
    ];
    $users = [];
    foreach ($seed as $s) {
        $users[] = [
            'userId' => $s['userId'],
            'name' => $s['name'],
            'password' => password_hash($s['pw'], PASSWORD_DEFAULT),
            'createdAt' => '2025-01-01 00:00:00',
        ];
    }
    file_put_contents($usersFile, json_encode($users, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT));
}

$action = $input['action'] ?? '';

if ($action === 'register') {
    $userId = trim($input['userId'] ?? '');
    $realName = trim($input['realName'] ?? '');
    $name = trim($input['name'] ?? '');
    $password = trim($input['password'] ?? '');

    if (!$userId) { echo json_encode(['success'=>false, 'error'=>'IDを入力してください']); exit; }
    if (!$realName) { echo json_encode(['success'=>false, 'error'=>'本名を入力してください']); exit; }
    if (!$name) { echo json_encode(['success'=>false, 'error'=>'ユーザー名を入力してください']); exit; }
    if (!$password) { echo json_encode(['success'=>false, 'error'=>'パスワードを入力してください']); exit; }
    if (mb_strlen($password) < 4) { echo json_encode(['success'=>false, 'error'=>'パスワードは4文字以上にしてください']); exit; }

    // Check duplicate userId
    foreach ($users as $u) {
        if ($u['userId'] === $userId) {
            echo json_encode(['success'=>false, 'error'=>'このIDは既に使われています']);
            exit;
        }
    }

    $users[] = [
        'userId' => $userId,
        'realName' => $realName,
        'name' => $name,
        'password' => password_hash($password, PASSWORD_DEFAULT),
        'createdAt' => date('Y-m-d H:i:s'),
    ];

    file_put_contents($usersFile, json_encode($users, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT));
    echo json_encode(['success'=>true, 'userId'=>$userId, 'realName'=>$realName, 'name'=>$name]);
    exit;
}

if ($action === 'login') {
    $userId = trim($input['userId'] ?? '');
    $password = trim($input['password'] ?? '');

    if (!$userId) { echo json_encode(['success'=>false, 'error'=>'IDを入力してください']); exit; }
    if (!$password) { echo json_encode(['success'=>false, 'error'=>'パスワードを入力してください']); exit; }

    foreach ($users as $u) {
        if ($u['userId'] === $userId) {
            if (password_hash_verify($u['password'], $password)) {
                echo json_encode(['success'=>true, 'userId'=>$userId, 'name'=>$u['name']]);
                exit;
            } else {
                echo json_encode(['success'=>false, 'error'=>'パスワードが正しくありません']);
                exit;
            }
        }
    }

    echo json_encode(['success'=>false, 'error'=>'このIDは登録されていません']);
    exit;
}

echo json_encode(['success'=>false, 'error'=>'Invalid action']);

function password_hash_verify($hash, $password) {
    return password_verify($password, $hash);
}
