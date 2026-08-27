<?php
/**
 * save.php — บันทึกข้อมูลเครื่องถ่ายเอกสารลงไฟล์ machines.json
 * วางไว้โฟลเดอร์เดียวกับ copierentryform.html บน XAMPP
 */
declare(strict_types=1);
header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store');

const FILE_NAME = 'machines.json';
const MAX_BYTES = 300000;
const MAX_MACHINES = 200;
$path = __DIR__ . DIRECTORY_SEPARATOR . FILE_NAME;

function out(int $code, array $body): never {
    http_response_code($code);
    echo json_encode($body, JSON_UNESCAPED_UNICODE);
    exit;
}
function s(mixed $v, int $max = 120): string {
    if (!is_string($v)) return '';
    $v = preg_replace('/[\x00-\x1F\x7F]/u', '', $v) ?? '';
    return mb_substr(trim($v), 0, $max, 'UTF-8');
}
function i(mixed $v): int {
    return (is_numeric($v) && abs((float)$v) < 1e12) ? (int)round((float)$v) : 0;
}

if (($_SERVER['REQUEST_METHOD'] ?? '') !== 'POST') {
    out(405, ['ok' => false, 'error' => 'ต้องส่งข้อมูลด้วยวิธี POST']);
}
$raw = file_get_contents('php://input');
if ($raw === false || strlen($raw) > MAX_BYTES) {
    out(413, ['ok' => false, 'error' => 'ข้อมูลใหญ่เกินกำหนด']);
}
$in = json_decode($raw, true);
if (!is_array($in) || !isset($in['machines']) || !is_array($in['machines'])) {
    out(400, ['ok' => false, 'error' => 'รูปแบบข้อมูลไม่ถูกต้อง']);
}
if (count($in['machines']) > MAX_MACHINES) {
    out(400, ['ok' => false, 'error' => 'จำนวนเครื่องเกิน ' . MAX_MACHINES]);
}

/* ---- ตรวจและคัดเฉพาะฟิลด์ที่รู้จัก ---- */
$KINDS = ['ขาว-ดำ', 'สี'];
$PKS   = ['bw', 'bwc', 'col'];
$clean = [];
$uids  = [];
foreach ($in['machines'] as $m) {
    if (!is_array($m)) continue;
    $uid = s($m['uid'] ?? '', 40);
    if ($uid === '' || isset($uids[$uid])) continue;
    $uids[$uid] = true;
    $kind = in_array($m['kind'] ?? '', $KINDS, true) ? $m['kind'] : 'ขาว-ดำ';
    $lines = [];
    $seen  = [];
    foreach ((is_array($m['lines'] ?? null) ? $m['lines'] : []) as $ln) {
        if (!is_array($ln)) continue;
        $pk = in_array($ln['pk'] ?? '', $PKS, true) ? $ln['pk'] : null;
        if ($pk === null || isset($seen[$pk])) continue;
        $seen[$pk] = true;
        $hist = [];
        foreach ((is_array($ln['hist'] ?? null) ? $ln['hist'] : []) as $h) {
            $hist[] = i($h);
            if (count($hist) >= 24) break;
        }
        $lines[] = ['pk' => $pk, 'lastMeter' => i($ln['lastMeter'] ?? 0),
                    'avg' => i($ln['avg'] ?? 0), 'hist' => $hist];
    }
    if (!$lines) continue;
    $clean[] = [
        'uid'    => $uid,
        'serial' => s($m['serial'] ?? '', 60),
        'kind'   => $kind,
        'model'  => s($m['model'] ?? '', 60),
        'place'  => s($m['place'] ?? '', 120),
        'dept'   => s($m['dept'] ?? '', 120),
        'lines'  => $lines,
    ];
}
if (!$clean) out(400, ['ok' => false, 'error' => 'ไม่พบข้อมูลเครื่องที่ใช้ได้']);

/* ---- เขียนไฟล์แบบล็อก กันเขียนชนกัน ---- */
$fp = @fopen($path, 'c+');
if (!$fp) out(500, ['ok' => false, 'error' => 'เขียนไฟล์ไม่ได้ ตรวจสิทธิ์ของโฟลเดอร์']);
if (!flock($fp, LOCK_EX)) { fclose($fp); out(500, ['ok' => false, 'error' => 'ล็อกไฟล์ไม่สำเร็จ']); }

$curRev = 0;
$cur = stream_get_contents($fp);
if (is_string($cur) && $cur !== '') {
    $j = json_decode($cur, true);
    if (is_array($j) && isset($j['rev'])) $curRev = (int)$j['rev'];
}
$base = isset($in['base']) ? (int)$in['base'] : -1;
if ($base >= 0 && $base !== $curRev) {
    flock($fp, LOCK_UN); fclose($fp);
    out(409, ['ok' => false, 'conflict' => true, 'rev' => $curRev,
              'error' => 'มีคนอื่นแก้ข้อมูลเครื่องไปแล้ว กรุณารีเฟรชหน้าเพื่อดึงข้อมูลล่าสุดก่อนแก้ต่อ']);
}

$payload = [
    'rev'      => $curRev + 1,
    'updated'  => date('Y-m-d H:i:s'),
    'machines' => $clean,
];
$json = json_encode($payload, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
rewind($fp); ftruncate($fp, 0); fwrite($fp, $json); fflush($fp);
flock($fp, LOCK_UN); fclose($fp);

out(200, ['ok' => true, 'rev' => $payload['rev'], 'updated' => $payload['updated'], 'count' => count($clean)]);
