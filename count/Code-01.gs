/**
 * ส่งผลตรวจนับสดจากชีต asset_count_list เป็น JSON ให้หน้า pivot (index.html)
 * ส่งออกเฉพาะ asset_id กับ count_status ไม่ส่งเลขบัตรผู้ตรวจนับ (user_count) และวันที่ออกไป
 * ชีตจึงเป็นส่วนตัวได้ ไม่ต้องแชร์ลิงก์สาธารณะ
 */
const SHEET_ID = '1sEhA-A-k1e5fdb2RLz2txirNCAatX-XBk4dyz6-Jbhs';
const TAB_NAME = 'ชีต1'; // ถ้าไม่พบแท็บชื่อนี้ จะใช้แท็บแรก

function doGet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(TAB_NAME) || ss.getSheets()[0];
  const v = sh.getDataRange().getValues();
  // ค่าเริ่มต้น: ไม่มีแถวหัว คอลัมน์ C = asset_id, E = count_status
  let ia = 2, ic = 4, start = 0;
  const head = v[0].map(x => String(x).trim().toLowerCase());
  if (head.indexOf('asset_id') >= 0 && head.indexOf('count_status') >= 0) {
    ia = head.indexOf('asset_id'); ic = head.indexOf('count_status'); start = 1;
  }
  const rows = [];
  for (let i = start; i < v.length; i++) {
    const id = Number(v[i][ia]);
    if (!id || !isFinite(id)) continue;
    rows.push([id, String(v[i][ic]).trim().toUpperCase()]);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ updated: new Date().toISOString(), rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}
