<?php
/**
 * data.php — ดึงข้อมูลจาก Google Sheet มาให้ dashboard.html
 * วางไว้โฟลเดอร์เดียวกับ dashboard.html บน XAMPP
 *
 * ต้องตั้งค่าชีตให้ "ผู้ที่มีลิงก์" อ่านได้ก่อน
 * เปิดชีต -> แชร์ -> การเข้าถึงทั่วไป -> ทุกคนที่มีลิงก์ -> ผู้อ่าน
 */
declare(strict_types=1);

/* ================= ตั้งค่าตรงนี้ ================= */
const SHEET_ID      = '1epiIfBG3FE29f7RwsOZF8rozIpXWTMMqczl5Xvc57Y4';
const SHEET_GID     = '2041319209';       // รหัสแท็บ (ดูได้จาก ...#gid=xxxxx ที่ท้าย URL) ใช้ตัวนี้ก่อน
const SHEET_NAME    = 'บันทึกมิเตอร์';   // ชื่อแท็บ ใช้สำรองเมื่อ gid ใช้ไม่ได้
const CACHE_MINUTES = 10;                 // เก็บข้อมูลไว้กี่นาทีก่อนไปดึงใหม่
/* ================================================ */

const CACHE_FILE = 'cache_meters.csv';
$cache = __DIR__ . DIRECTORY_SEPARATOR . CACHE_FILE;

function fail(int $code, string $msg, string $hint = ''): never {
    http_response_code($code);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(['ok' => false, 'error' => $msg, 'hint' => $hint], JSON_UNESCAPED_UNICODE);
    exit;
}
function sendCsv(string $csv, string $src, int $age): never {
    header('Content-Type: text/csv; charset=utf-8');
    header('Cache-Control: no-store');
    header('X-Data-Source: ' . $src);
    header('X-Data-Age: ' . $age);
    echo $csv;
    exit;
}

$force   = isset($_GET['refresh']);
$hasCache = is_file($cache) && filesize($cache) > 0;
$age      = $hasCache ? (time() - (int)filemtime($cache)) : PHP_INT_MAX;

/* ยังไม่หมดอายุ ใช้ของเดิม */
if (!$force && $hasCache && $age < CACHE_MINUTES * 60) {
    sendCsv((string)file_get_contents($cache), 'cache', $age);
}

$base = 'https://docs.google.com/spreadsheets/d/' . SHEET_ID . '/gviz/tq?tqx=out:csv';
$urls = [];
if (SHEET_GID !== '') $urls[] = $base . '&gid=' . rawurlencode(SHEET_GID);
$urls[] = $base . '&sheet=' . rawurlencode(SHEET_NAME);

$body = null;
foreach ($urls as $url) {
  $body = fetchUrl($url);
  if ($body !== null && $body !== '' && stripos(ltrim(substr($body, 0, 200)), '<') !== 0) break;
  $body = null;
}

function fetchUrl(string $url): ?string {
if (function_exists('curl_init')) {
    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_FOLLOWLOCATION => true,
        CURLOPT_MAXREDIRS      => 5,
        CURLOPT_TIMEOUT        => 20,
        CURLOPT_USERAGENT      => 'Mozilla/5.0 (compatible; CopierDashboard/1.0)',
    ]);
    $res  = curl_exec($ch);
    $code = (int)curl_getinfo($ch, CURLINFO_RESPONSE_CODE);
    curl_close($ch);
    return (is_string($res) && $code === 200) ? $res : null;
} elseif (ini_get('allow_url_fopen')) {
    $ctx = stream_context_create(['http' => ['timeout' => 20, 'follow_location' => 1,
        'user_agent' => 'Mozilla/5.0 (compatible; CopierDashboard/1.0)']]);
    $res = @file_get_contents($url, false, $ctx);
    return (is_string($res) && $res !== '') ? $res : null;
}
return null;
}

/* ดึงไม่สำเร็จ -> ใช้ของเก่าถ้ามี */
if ($body === null || $body === '') {
    if ($hasCache) sendCsv((string)file_get_contents($cache), 'cache-stale', $age);
    fail(502, 'ดึงข้อมูลจาก Google Sheet ไม่สำเร็จ',
        'ตรวจว่าเครื่องเซิร์ฟเวอร์ต่ออินเทอร์เน็ตได้ และเปิด extension php_curl ใน XAMPP แล้ว');
}

/* ได้หน้า HTML แปลว่ายังไม่ได้เปิดสิทธิ์ให้อ่าน */
$head = ltrim(substr($body, 0, 400));
if (stripos($head, '<html') !== false || stripos($head, '<!doctype') !== false) {
    if ($hasCache) sendCsv((string)file_get_contents($cache), 'cache-stale', $age);
    fail(403, 'Google ไม่ยอมให้อ่านชีตนี้แบบไม่ล็อกอิน',
        'เปิดชีต -> แชร์ -> การเข้าถึงทั่วไป -> เปลี่ยนเป็น "ทุกคนที่มีลิงก์" สิทธิ์ "ผู้อ่าน" แล้วลองใหม่');
}
if (strpos($body, ',') === false && strlen($body) < 50) {
    if ($hasCache) sendCsv((string)file_get_contents($cache), 'cache-stale', $age);
    fail(502, 'ข้อมูลที่ได้จาก Google Sheet ว่างเปล่า', 'ตรวจว่าชื่อแท็บในไฟล์ตรงกับค่า SHEET_NAME ใน data.php');
}

@file_put_contents($cache, $body, LOCK_EX);
sendCsv($body, 'live', 0);
