<?php
// GAS API キャッシュプロキシ
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Cache-Control: public, max-age=60');

$GAS_URL = 'https://script.google.com/macros/s/AKfycbwojGHuvzycc07FJKwBdbBJJQZpssF6lYz0DbNJlu6zsVuXkAj8V8w3XNBPieo2wsYbFg/exec';
$CACHE_DIR = __DIR__ . '/cache';
$CACHE_TTL = 180;

if (!is_dir($CACHE_DIR)) mkdir($CACHE_DIR, 0755, true);

$query = $_SERVER['QUERY_STRING'] ?? '';
$cacheKey = md5($query);
$cacheFile = $CACHE_DIR . '/' . $cacheKey . '.json';

// キャッシュが有効ならそれを返して即終了
if (file_exists($cacheFile) && (time() - filemtime($cacheFile)) < $CACHE_TTL) {
    readfile($cacheFile);
    exit;
}

// cURLでGAS APIにリクエスト（file_get_contentsよりリダイレクト処理が確実）
$ch = curl_init();
curl_setopt_array($ch, [
    CURLOPT_URL => $GAS_URL . ($query ? '?' . $query : ''),
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_FOLLOWLOCATION => true,
    CURLOPT_MAXREDIRS => 5,
    CURLOPT_TIMEOUT => 30,
    CURLOPT_SSL_VERIFYPEER => true,
]);
$response = curl_exec($ch);
$httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
curl_close($ch);

if ($response !== false && $httpCode === 200) {
    file_put_contents($cacheFile, $response, LOCK_EX);
    echo $response;
} elseif (file_exists($cacheFile)) {
    // GAS失敗 → 古いキャッシュ返す
    readfile($cacheFile);
} else {
    http_response_code(502);
    echo '{"error":"GAS API unavailable"}';
}
