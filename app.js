/**
 * ============================================================
 *  ระบบจัดการครุภัณฑ์ CICM — Web Edition
 *  เวอร์ชัน : 5.0  (Client-side / SheetJS port of GAS v4.1)
 *  Runtime  : Browser (ES2020+), ไม่มี dependency นอกจาก SheetJS
 * ============================================================
 *
 *  Pipeline (เหมือนเดิมทุกขั้น เปลี่ยนเฉพาะ I/O layer)
 *  ─────────────────────────────────────────────────────────
 *    [file: data]     ──processRow──▶  information (33 cols)
 *    [file: RP023/15] ──ErpBuilder──▶  3.ทะเบียนสินทรัพย์ (44 cols)
 *    [file: รูปภาพ]   ──ImageCache─▶  url index
 *                                │
 *                       information + ERP + images
 *                                │
 *                            mapRow
 *                                ▼
 *              3.ทะเบียนรายการครุภัณฑ์ CICM (74 cols)
 *
 *  สิ่งที่เปลี่ยนจาก GAS
 *  ─────────────────────────────────────────────────────────
 *  ❌ SpreadsheetApp / Utilities / PropertiesService / ScriptApp / MailApp
 *  ✅ SheetJS (XLSX.read / XLSX.write) + FileReader
 *  ✅ _mapFormulas  →  FormulaEngine  (DATEDIF / XLOOKUP / VLOOKUP เป็น JS)
 *  ✅ BatchRunner (time-based trigger) → async chunk loop + progress bar
 *  ✅ Formatter.setRichTextValues → cell.l (hyperlink ของ xlsx)
 *
 *  สิ่งที่คงไว้ตามเดิม 100%
 *  ─────────────────────────────────────────────────────────
 *  CONFIG · HEADERS · TRANSLATION_MAPS · TRANSLATION_PAD · Translators
 *  StringHelper · DateFormatter · DataProcessor.processRow
 *  KeyNormalizer · UrlNormalizer · ImageCache · Validators · Mapper
 * ============================================================
 */

'use strict';

/* ============================================================
 *  PART A — FOUNDATION
 * ============================================================ */

/**
 * รูปแบบรหัสสินทรัพย์ ERP — ใช้ร่วมกันทุกที่
 *   ตัวเลข 21 หลัก  =  2 + 2 + 4 + 3 + 4 + 6
 *   ขีดกลาง 5 ตัว   →  รวม 26 อักขระ
 *   ตัวอย่าง: 27-69-0950-001-0012-000002
 */
/**
 * index ของคอลัมน์ใน target sheet ที่โค้ดอ้างถึงบ่อย
 * (อ้าง "แถวของ target 74 คอลัมน์" ไม่ใช่ information 33 คอลัมน์ — คนละชุด index)
 *   TGT.detail = E = 'รายละเอียดเพิ่มเติม'      ← information[8]  (asset_detail)
 *   TGT.assetNo= F = 'หมายเลขครุภัณฑ์'          ← information[9]  (asset_no)
 *   TGT.oldNo  = G = 'หมายเลขครุภัณฑ์เดิม'      ← information[11] (asset_no_old)
 *   TGT.erpNo  = BV= 'หมายเลขครุภัณฑ์ ERP'      ← information[10] (asset_no_erp)
 */
const TGT = Object.freeze({
  detail:  4,   // E
  assetNo: 5,   // F
  oldNo:   6,   // G
  image:  25,   // Z
  endDate:26,   // AA
  age:    27,   // AB
  erpCode:28,   // AC  ← ผลลัพธ์ของ acLookupChain
  pullStart:29, // AD  ← เริ่มดึงจาก ERP
  erpNo:  73,   // BV
  project:74     // BW ← โครงการจาก asset!AA
});

/** index ใน information (33 คอลัมน์) ที่ Mapper อ้างถึง */
const INFO = Object.freeze({
  detail: 8, assetNo: 9, erpNo: 10, oldNo: 11, project: 33
});

/**
 * โครงสร้าง input ของชีต asset (0-based)
 *
 * มีการเพิ่ม project_id ที่ AA (คอลัมน์ 27) ดังนั้น section_id เดิมและฟิลด์
 * ถัดไปจะเลื่อนไป AB–AH ทั้งหมด ห้ามอ้าง index เหล่านี้
 * โดยตรงนอก schema นี้ เพื่อให้การอ่านไฟล์ที่มีหัวตารางยังทนต่อการสลับคอลัมน์ได้
 */
const ASSET_INPUT = Object.freeze({
  columns: Object.freeze({
    id: 0, receive_date: 1, type_asset_code: 2, asset_name: 3, asset_brand: 4,
    asset_model: 5, asset_color: 6, asset_donate: 7, asset_detail: 8, asset_no: 9,
    asset_no_erp: 10, asset_no_old: 11, serial_no: 12, group_code: 13, class_code: 14,
    type_code: 15, type_code_id: 16, unit_code: 17, amount: 18, price: 19,
    seller_name: 20, warranty_start: 21, warranty_end: 22, build_code: 23, floor_code: 24,
    room_code: 25, project_id: 26, section_id: 27, budget_code: 28, status_code: 29,
    budget_year: 30, remark_detail: 31, remark_section: 32, istatus: 33
  }),
  headerAliases: Object.freeze({
    id: ['id', 'ลำดับ', 'no', 'no.', '#'],
    receive_date: ['receive_date', 'วันที่ตรวจรับ'],
    type_asset_code: ['type_asset_code'], asset_name: ['asset_name', 'ชื่อครุภัณฑ์'],
    asset_brand: ['asset_brand', 'ยี่ห้อ'], asset_model: ['asset_model', 'รุ่น'],
    asset_color: ['asset_color', 'สี'], asset_donate: ['asset_donate'],
    asset_detail: ['asset_detail', 'รายละเอียด'], asset_no: ['asset_no', 'หมายเลขครุภัณฑ์'],
    asset_no_erp: ['asset_no_erp', 'หมายเลขครุภัณฑ์erp'],
    asset_no_old: ['asset_no_old', 'หมายเลขครุภัณฑ์เดิม'], serial_no: ['serial_no', 'serialno'],
    group_code: ['group_code'], class_code: ['class_code'], type_code: ['type_code'],
    type_code_id: ['type_code_id'], unit_code: ['unit_code'], amount: ['amount', 'จำนวน'],
    price: ['price', 'ราคา'], seller_name: ['seller_name', 'ผู้ขาย'],
    warranty_start: ['warranty_start', 'เริ่มรับประกัน'], warranty_end: ['warranty_end', 'สิ้นสุดรับประกัน'],
    build_code: ['build_code'], floor_code: ['floor_code', 'ชั้น'], room_code: ['room_code', 'ห้อง'],
    project_id: ['project_id', 'projectid', 'project', 'โครงการ'],
    section_id: ['section_id', 'sectionid', 'รหัสหน่วยงาน', 'หน่วยงาน', 'department'],
    budget_code: ['budget_code', 'budgetcode', 'รหัสงบประมาณ'],
    status_code: ['status_code', 'statuscode', 'รหัสสถานะ'], budget_year: ['budget_year', 'ปีงบ'],
    remark_detail: ['remark_detail', 'หมายเหตุ'], remark_section: ['remark_section', 'หมายเหตุส่วนงาน'],
    istatus: ['istatus', 'สถานะระบบงาน']
  })
});

const ASSET_CODE_GUARD = Object.freeze({
  prefix: '27-',
  length: 26,
  digits: 21,
  regex:  /^\d{2}-\d{2}-\d{4}-\d{3}-\d{4}-\d{6}$/
});

const CONFIG = Object.freeze({

  /** ชื่อชีตในไฟล์ผลลัพธ์ (เดิมคือชื่อชีตใน spreadsheet) */
  sheets: {
    data:        'data',
    source:      'information',
    target:      '3.ทะเบียนรายการครุภัณฑ์ CICM',
    image:       'รูปภาพ',
    assetLookup: '3.ทะเบียนสินทรัพย์',
    diagnostic:  '_diagnose_images',
    validation:  '_validation_log'
  },

  area: { startRow: 2, startCol: 1, maxCols: 75 },

  /** information sheet schema (output ของ processRow; รวมโครงการจาก asset!AA) */
  source: {
    totalCols:  34,
    batchSize:  2000,
    assetNoCol: 0        // index ที่ใช้เป็น key matching รูปภาพ = "ลำดับ"
  },

  dataProcessing: { batchSize: 5000, progressEvery: 500 },

  /** layout ของชีต "รูปภาพ" (0-based array index) */
  image: {
    codeCol:   1,        // B = ลำดับ (key)
    urlCol:    6,        // G = URL หรือชื่อไฟล์
    totalCols: 7,
    targetCol: 26,       // Z (1-based) ใน target
    baseUrl:   'https://eptumed.com/main/manageAsset/uploads/asset_picture/'
  },

  /** layout ของชีต ERP "3.ทะเบียนสินทรัพย์" (0-based) */
  erp: {
    totalCols:    44,
    codeCol:       2,    // C  = รหัสสินทรัพย์      (คีย์ของ VLOOKUP)
    refNo1Col:    15,    // P  = Ref No. 1          (คีย์ของ XLOOKUP)
    firstPullCol:  3,    // เริ่มดึงเข้า target ที่ index 3
    pullCount:    41,    // ดึง 41 คอลัมน์ → target[29..69]

    /**
     * ลำดับการค้นหาของคอลัมน์ AC (target index 28)
     * ลำดับ:
     *   1) XLOOKUP(G, ERP!P, ERP!C)          — G ↔ Ref No. 1
     *   2) XLOOKUP(G, ERP!C, ERP!C)          — G เป็นรหัสสินทรัพย์อยู่แล้ว
     *   3) ถ้า F ตรงรูปแบบรหัสสินทรัพย์ → ใช้ F เป็น AC ตรงๆ (ไม่ต้องมีใน ERP)
     *   4) XLOOKUP(E, ERP!C, ERP!C)          — E เป็นรหัสสินทรัพย์
     *   ไม่เข้าเงื่อนไขไหนเลย → ''
     *
     * `by`    : 'ref1' = ค้นด้วย Ref No.1 | 'code' = ค้นด้วยรหัสสินทรัพย์
     *           'self' = ใช้ค่าในคอลัมน์นั้นเป็นคำตอบเลย (ไม่ค้น ERP)
     * `guard` : เงื่อนไขที่ต้องผ่านก่อนจึงจะทำชั้นนั้น
     *           ไม่ผ่าน → ข้ามไปชั้นถัดไปทันที
     * แก้ลำดับ / เพิ่มชั้น / ปรับ guard ได้ที่นี่จุดเดียว
     */
    acLookupChain: [
      // 1) คอลัมน์ใหม่ BV 'หมายเลขครุภัณฑ์ ERP' (← data คอลัมน์ K) — แม่นที่สุด
      { col: TGT.erpNo,   by: 'code', label: 'BV→C', guard: ASSET_CODE_GUARD },
      { col: TGT.oldNo,   by: 'ref1', label: 'G→P' },   // G → Ref No. 1
      { col: TGT.oldNo,   by: 'code', label: 'G→C' },   // G เป็นรหัสสินทรัพย์อยู่แล้ว
      { col: TGT.assetNo, by: 'self', label: 'F=รหัส',  // F ตรงรูปแบบ → ใช้เป็น AC เลย
        guard: ASSET_CODE_GUARD },
      { col: TGT.detail,  by: 'code', label: 'E→C' },   // E → รหัสสินทรัพย์
      // 6) กวาดทุกคอลัมน์ในแถวเดียวกัน หาค่าที่ตรงรูปแบบรหัสสินทรัพย์ แล้ว XLOOKUP
      { by: 'scan', label: 'SCAN', guard: ASSET_CODE_GUARD }
    ],

    /** index ของคอลัมน์ใหม่ 'หมายเลขครุภัณฑ์ ERP' ใน target */
    erpNoCol: TGT.erpNo,
    /** index ที่ไม่ต้องกวาดตอน by:'scan' (AC คือช่องคำตอบเอง) */
    scanSkip: [TGT.erpCode],

    /** จำนวนแถวที่ยอมสแกนหาหัวตารางของรายงาน ERP */
    headerScanRows: 40,
    /** คำที่ใช้ระบุว่าแถวนั้นเป็นหัวตาราง (เจอคำใดคำหนึ่งก็พอ) */
    headerHints: ['รหัสสินทรัพย์', 'รหัสหน่วยงาน', 'ชื่อสินทรัพย์'],
    /** index ของคอลัมน์ค่าเสื่อม (1)…(7) */
    dep: { cost: 37, acc2: 38, cur3: 39, per4: 40, sum5: 41, acc6: 42, nbv7: 43 }
  },

  /**
   * คอลัมน์วันที่ใน target (1-based) — ใช้กำหนดความกว้าง
   */
  dateColumns: [2, 27, 44, 58, 59, 60, 61],

  /**
   * ⚠️ คอลัมน์วันที่ในไฟล์นี้ "ปนกัน 2 ระบบปี" — ห้ามแปลงรวดเดียวทั้งหมด
   *
   *   ค.ศ. (มาจากชีต data)      : วันที่ตรวจรับ · วันที่สิ้นสุด · วันที่ตัดจำหน่าย
   *   พ.ศ. (มาจากรายงาน ERP)    : วันที่จัดทำ · ขึ้นทะเบียน · เริ่มคำนวณค่าเสื่อม · คำนวณล่าสุด
   *
   * ถ้าเขียนคอลัมน์ พ.ศ. เป็น native Excel date จะกลายเป็น ค.ศ. 2565
   * = เพี้ยนไป 543 ปี  →  จึงเก็บเป็นข้อความไว้เหมือนเดิม
   * ส่วนคอลัมน์ ค.ศ. แปลงเป็นวันที่จริง เพื่อให้ Sort/Filter ตามเวลาได้
   */
  dateColumnsCE: [1, 26, 43],        // 0-based — แปลงเป็น native date
  dateColumnsBE: [57, 58, 59, 60],   // 0-based — คงเป็นข้อความ (ปี พ.ศ.)
  dateNumberFormat: 'dd/mm/yyyy',
  /** เพดานปีที่ยอมให้แปลง — เกินนี้ถือว่าเป็น พ.ศ. ไม่แปลง */
  ceYearMax: 2400,

  /** ตัวเลือกที่ผู้ใช้ปรับได้จาก UI */
  options: {
    skipDataHeaderRow:  true,
    dropErpTotalRows:   true,
    computeDepreciation:true,
    stripLeadingQuote:  true,
    /** แปลงคอลัมน์วันที่ ค.ศ. เป็น native Excel date (Sort/Filter ได้) */
    nativeExcelDates:   true
  },

  /** ขนาด chunk ของ async loop (แทน trigger-based batch) */
  ui: { chunkSize: 750 },

  debug: true
});

/* ── HEADERS — 75 คอลัมน์ ของ target ───────────────────── */
const HEADERS = Object.freeze([
  'ลำดับ', 'วันที่ตรวจรับ', 'ประเภท', 'ชื่อครุภัณฑ์', 'รายละเอียดเพิ่มเติม',
  'หมายเลขครุภัณฑ์', 'หมายเลขครุภัณฑ์เดิม', 'ยี่ห้อ', 'รุ่น', 'สี',
  'ครุภัณฑ์บริจาค', 'วิธีจำหน่าย', 'ราคา', 'ชื่อผู้ขาย', 'สถานที่ตั้ง-ตึก/อาคาร',
  'ชั้น', 'ห้อง', 'หน่วยงาน', 'งบประมาณ', 'สถานะ',
  'ปีงบ', 'หมายเหตุ', 'หมายเหตุส่วนงาน', 'สถานะระบบงาน', 'รายละเอียดครุภัณฑ์',
  'รูปภาพ', 'วันที่สิ้นสุด', 'อายุการใช้งาน', 'หมายเลขครุภัณฑ์ ERP (ระบบสินทรัพย์)',
  'สินทรัพย์ย่อย', 'รหัสสินทรัพย์อ้างอิง', 'ชื่อสินทรัพย์',
  'รหัสและชื่อผู้ขาย/ผู้ให้บริการ', 'เลขที่ (เอกสาร) ตรวจรับ', 'เลขที่ (เอกสาร) ขึ้นทะเบียน',
  'รายละเอียดสินทรัพย์', 'ยี่ห้อERP', 'รุ่นERP', 'สีERP', 'Serial No. ERP', 'สภาพ',
  'Ref No. 1', 'Ref No. 2', 'วันที่ตัดจำหน่าย', 'ศูนย์ศึกษา', 'แหล่งเงิน',
  'ตึก/อาคาร', 'ชั้นERP', 'ห้องERP', 'ศูนย์กำไรขาดทุน (ค่าเสื่อมราคา)',
  'ประเภทสินทรัพย์ตามผังบัญชี', 'ของบริจาค', 'ของแถม',
  'กลุ่ม (FSN)', 'ประเภท (FSN)', 'ชนิด (FSN)', 'รายละเอียด (FSN)',
  'วันที่จัดทำ', 'วันที่ขึ้นทะเบียน', 'วันที่เริ่มคำนวณค่าเสื่อม', 'วันที่คำนวณล่าสุด',
  'อายุใช้งาน (ปี)', 'อายุใช้งาน (เดือน)',
  'มูลค่าต้นทุนที่ได้มา (1)', 'ค่าเสื่อมราคาสะสมยกมา (2)',
  'ค่าเสื่อมราคาสะสมระหว่างปียกมา (3)', 'ค่าเสื่อมราคางวดนี้ (4)',
  'ค่าเสื่อมราคาสะสมระหว่างปียกไป (5) = (3)+(4)',
  'ค่าเสื่อมราคาสะสมยกไป (6) = (2)+(5)', 'มูลค่าปัจจุบัน (7) = (1)-(6)',
  'ประเภทพัสดุ', 'สภาพครุภัณฑ์', 'เอกสารอ้างอิง',
  /* index 73 (คอลัมน์ BV) รับค่าจาก information คอลัมน์ K (asset_no_erp) */
  'หมายเลขครุภัณฑ์ ERP',
  /* index 74 (คอลัมน์ BW) รับค่าจาก asset คอลัมน์ AA (project_id) */
  'โครงการ'
]);

/* ── ERP HEADERS — 44 คอลัมน์ ของชีต "3.ทะเบียนสินทรัพย์" ── */
const ERP_HEADERS = Object.freeze([
  'รหัสหน่วยงาน - ชื่อหน่วยงาน', 'ผู้รับผิดชอบ', 'รหัสสินทรัพย์', 'สินทรัพย์ย่อย',
  'รหัสสินทรัพย์อ้างอิง', 'ชื่อสินทรัพย์', 'รหัส และชื่อผู้ขาย/ผู้ให้บริการ',
  'เลขที่ (เอกสาร) ตรวจรับ', 'เลขที่ (เอกสาร) ขึ้นทะเบียน', 'รายละเอียดสินทรัพย์',
  'ยี่ห้อ', 'รุ่น', 'สี', 'Serial No.', 'สภาพ', 'Ref No. 1', 'Ref No. 2',
  'สถานะ/วันที่ตัดจำหน่าย', 'ศูนย์ศึกษา', 'แหล่งเงิน', 'ตึก/อาคาร', 'ชั้น', 'ห้อง',
  'ศูนย์กำไรขาดทุน (ค่าเสื่อมราคา)', 'ประเภทสินทรัพย์ตามผังบัญชี', 'ของบริจาค', 'ของแถม',
  'กลุ่ม (FSN)', 'ประเภท (FSN)', 'ชนิด (FSN)', 'รายละเอียด (FSN)',
  'วันที่จัดทำ', 'วันที่ขึ้นทะเบียน', 'วันที่เริ่มคำนวณค่าเสื่อม', 'วันที่คำนวณล่าสุด',
  'อายุใช้งาน (ปี)', 'อายุใช้งาน (เดือน)',
  'มูลค่าต้นทุนที่ได้มา (1)', 'ค่าเสื่อมราคาสะสมยกมา (2)',
  'ค่าเสื่อมราคาสะสมระหว่างปียกมา (3)', 'ค่าเสื่อมราคางวดนี้ (4)',
  'ค่าเสื่อมราคาสะสมระหว่างปียกไป (5)', 'ค่าเสื่อมราคาสะสมยกไป (6)',
  'มูลค่าปัจจุบัน (7)'
]);

/* ── Invariants — ตรวจตอนโหลดไฟล์ ถ้าไม่ผ่านให้ล้มทันที ─── *
 *  กันเคสเพิ่ม/ลบคอลัมน์แล้วลืมอัปเดตที่อื่น ซึ่งจะทำให้ข้อมูล
 *  เลื่อนคอลัมน์แบบเงียบๆ (หาเจอยากกว่า error ตอนโหลดมาก)
 * ============================================================ */
(function assertInvariants() {
  const fail = [];
  if (HEADERS.length !== CONFIG.area.maxCols)
    fail.push(`HEADERS มี ${HEADERS.length} คอลัมน์ แต่ CONFIG.area.maxCols = ${CONFIG.area.maxCols}`);
  if (ERP_HEADERS.length !== CONFIG.erp.totalCols)
    fail.push(`ERP_HEADERS มี ${ERP_HEADERS.length} คอลัมน์ แต่ CONFIG.erp.totalCols = ${CONFIG.erp.totalCols}`);
  if (HEADERS[TGT.erpNo] !== 'หมายเลขครุภัณฑ์ ERP')
    fail.push(`HEADERS[${TGT.erpNo}] ควรเป็น 'หมายเลขครุภัณฑ์ ERP' แต่ได้ '${HEADERS[TGT.erpNo]}'`);
  if (TGT.pullStart + CONFIG.erp.pullCount - 1 !== TGT.erpNo - 4)
    fail.push(`ช่วงดึงจาก ERP (AD…) ไม่ลงท้ายที่คอลัมน์ 70 ตามที่ควรเป็น`);
  if (CONFIG.erp.firstPullCol + CONFIG.erp.pullCount - 1 !== CONFIG.erp.totalCols - 1)
    fail.push(`firstPullCol + pullCount ไม่ครอบคลุมถึงคอลัมน์สุดท้ายของ ERP`);

  if (fail.length) {
    const msg = 'CICM config ไม่สอดคล้องกัน:\n  • ' + fail.join('\n  • ');
    console.error(msg);
    throw new Error(msg);
  }
})();

/* ── Logger — เขียนทั้ง console และ log panel ────────────── */
const Logger_ = {
  _sink: null,
  bind(fn) { this._sink = fn; },
  _emit(kind, args) {
    const msg = args.map(a =>
      typeof a === 'object' && a !== null ? JSON.stringify(a) : String(a)
    ).join(' ');
    if (this._sink) this._sink(kind, msg);
  },
  info(...a)  { if (CONFIG.debug) console.log(...a); this._emit('i', a); },
  ok(...a)    { console.log(...a);                   this._emit('s', a); },
  warn(...a)  { console.warn(...a);                  this._emit('w', a); },
  error(...a) { console.error(...a);                 this._emit('e', a); }
};

/* ── ErrorHandler — เดิมเรียก SpreadsheetApp.getUi().alert ─ */
const ErrorHandler = {
  handle(fnName, err) {
    const message = err?.message ?? String(err);
    const stack   = err?.stack ?? '(no stack)';
    Logger_.error(`[${fnName}] ${message}`);
    if (CONFIG.debug) console.error(stack);
    UI.alert(`❌ เกิดข้อผิดพลาด\n\n[${fnName}] ${message}`);
  }
};


/* ============================================================
 *  PART B — data → information
 * ============================================================ */

const TRANSLATION_MAPS = Object.freeze({

  category: {
    '09': 'ที่ดิน อาคาร สิ่งก่อสร้าง',
    '23': 'ยานพาหนะพื้นดิน ยานยนต์ รถพ่วง และจักรยาน',
    '35': 'เครื่องมือและอุปกรณ์ห้องปฏิบัติการ(ทดลอง)',
    '36': 'เครื่องจักรกล เฉพาะงานอุตสาหกรรม',
    '39': 'อุปกรณ์การยกขนพัสดุ',
    '41': 'อุปกรณ์ตู้เย็น เครื่องปรับอากาศ และเครื่องถ่ายเทอากาศ',
    '42': 'อุปกรณ์ผจญเพลิง เครื่องช่วยชีวิตและความปลอดภัย',
    '44': 'เตาหลอมโรงไอน้ำและอุปกรณ์ทำให้แห้งและปฏิกรณ์นิวเคลียร์',
    '46': 'เครื่องกรองน้ำและเครื่องขจัดสิ่งโสโครก',
    '51': 'เครื่องมือ',
    '58': 'อุปกรณ์คมนาคม การค้นหา และการกระจายคลื่น',
    '59': 'อุปกรณ์และส่วนประกอบเครื่องไฟฟ้าและเครื่องอีเลคโทรนิค',
    '61': 'สายไฟ เครื่องกำเนิดไฟฟ้าและอุปกรณ์จ่ายไฟฟ้า',
    '65': 'อุปกรณ์ และเครื่องใช้ทางการแพทย์ ทันตแพทย์ และสัตวแพทย์',
    '66': 'เครื่องมือและอุปกรณ์ห้องปฏิบัติการ(ทดลอง)',
    '67': 'อุปกรณ์การภาพ',
    '69': 'เครื่องช่วยฝึกและอุปกรณ์',
    '71': 'เครื่องตกแต่ง',
    '72': 'เครื่องใช้และเครื่องตกแต่งบ้านและร้านค้า',
    '73': 'อุปกรณ์ประกอบอาหารและเลี้ยงดู',
    '74': 'เครื่องกลสำนักงานและอุปกรณ์ กรรมวิธีบันทึกและลงข้อมูล',
    '75': 'พัสดุและเครื่องใช้สำนักงาน',
    '77': 'เครื่องดนตรี เครื่องเล่นแผ่นเสียง และวิทยุใช้ในบ้าน',
    '81': 'ภาชนะบรรจุหีบห่อ',
    '99': 'เบ็ดเตล็ด'
  },

  unit: {
    '001': 'กรง',      '002': 'กล้อง',   '003': 'ขา',      '004': 'คัน',
    '005': 'เครื่อง',  '006': 'งาน',     '007': 'จอ',      '008': 'ชิ้น',
    '009': 'ชุด',      '010': 'ดอก',     '011': 'ตัว',     '012': 'ตู้',
    '013': 'เตียง',    '014': 'ถัง',     '015': 'แถว',     '016': 'แท่น',
    '017': 'ใบ',       '018': 'แผง',     '019': 'แผ่น',    '020': 'ระบบ',
    '021': 'รูป',      '022': 'ไลเซนส์', '023': 'หัว',     '024': 'องค์',
    '025': 'อัน'
  },

  building: {
    '001': 'อาคารเรียนและปฏิบัติการรวม',
    '002': 'อาคารปิยชาติ',
    '003': 'บร.5',
    '004': 'บร.4',
    '005': 'ราชสุดา',
    '006': 'รพ.ธรรมศาสตร์',
    '007': 'มูลนิธิแพนราชเทวี',
    '008': 'รพ.บำรุงราษฎร์ (อาคาร BIC)',
    '009': 'รพ.บำรุงราษฎร์ (อาคาร BIT)',
    '010': 'รพ.ยาสูบ',
    '011': 'รพ.อยุธยา',
    '012': 'หอพักนักศึกษาแพทย์',
    '013': 'สำนักงานวิทยาศาสตร์และเทคโนโลยีชั้นสูง',
    '014': 'ศูนย์เครื่องมือวิทยาศาสตร์เพื่อการวิจัยขั้นสูง',
    '015': 'อาคารศูนย์วิจัยชั้นสูง',
    '016': 'อุทยานการเรียนรู้ป๋วย 100 ปี'
  },

  room: {
    '001': 'ห้องปฏิบัติการ',             '002': 'ห้องเรียน',
    '003': 'ห้อง Server',                 '004': 'ห้อง Shaft',
    '005': 'ห้องพักอาจารย์',             '006': 'ห้องประชุม',
    '007': 'ห้องเก็บของ',                '008': 'สำนักงาน',
    '009': 'สำนักงานเลขานุการ',          '010': 'สำนักงานคณบดี',
    '011': 'คลินิกแพทย์แผนจีน',         '012': 'คลินิกแพทย์ผสมผสาน',
    '013': 'ห้องกายวิภาค',               '014': 'สำนักงานกิจการนักศึกษา',
    '015': 'ห้อง 230',                   '016': 'ห้อง 231',
    '017': 'สำนักงานศูนย์สอบ',          '018': 'สำนักงานการแพทย์บูรณาการ',
    '019': 'คลังพัสดุ',                  '020': 'สโมสรนักศึกษา',
    '021': 'Zone HA (หญิง)',             '022': 'Zone HB (ชาย)',
    '023': 'A-3004',                     '024': 'A-3005',
    '025': 'A-3011',                     '026': 'A-3014',
    '027': 'A-3025',                     '028': 'A-3026',
    '029': 'A-3027',                     '030': 'A-3028',
    '031': 'A-3029',                     '032': 'A-3030',
    '033': 'A-3031',                     '034': 'A-3032',
    '035': 'A-3033',                     '036': 'A-3034',
    '037': 'A-3035',                     '038': 'A-4018',
    '039': 'Skill Lab',                  '040': 'Skill Lab Teambase',
    '041': 'Skill Lab ทัศนมาศ',         '042': 'ห้องทันตกรรม',
    '043': 'ห้องเรียน GI',              '044': 'คลินิกการแพทย์ผสมผสาน',
    '045': 'CVS 1',                      '046': 'CVS 2',
    '047': 'Skill Lab CVS',              '048': 'คลินิกผิวหนัง',
    '049': 'อายุรกรรม',                 '050': 'คลินิกแพทย์แผนไทย',
    '051': 'ลานจอดรถ',                  '052': 'A-4003',
    '053': 'ห้องปฏิบัติการวิจัย',       '054': 'โถงกลาง',
    '055': 'ห้อง 241',                  '056': 'A-3010',
    '057': 'ห้องพักอาจารย์แพทย์แผนจีน', '058': 'A-1016',
    '059': 'ห้องชมรม CICM Band',        '060': 'TBL',
    '061': 'ห้อง 1004',                 '062': 'Control & Service room',
    '063': 'GI',                         '064': 'ทางเข้าออกด้านหน้า',
    '065': 'Traning laboratory',         '066': 'ประตูห้อง Research and Innovation laboratory',
    '067': 'Central laboratory',         '068': 'Cell Culture room',
    '069': 'Server Room',                '070': 'Staff Room',
    '071': 'Co-Working Space',           '072': 'Training laboratory',
    '073': 'Research Innovation 2',      '074': 'Central Lab',
    '075': 'Research Innovation 4',      '076': 'Research Innovation 5',
    '077': 'ห้องกิจการนักศึกษาและวิเทศสัมพันธ์',
    '078': 'ห้องเรียนทัศนมาตร',        '079': 'Control room',
    '080': '6001',                       '081': 'ห้อง 2025',
    '082': 'ห้องเรียน CVS',             '083': 'ทางเดิน',
    '084': 'ห้อง ECHO',                 '085': 'ห้องพักอาจารย์บูรณาการ',
    '086': 'ทางเข้าออกด้านหลัง',       '087': 'International Office',
    '088': 'ทางเดินเข้าห้องเรียนกลุ่มย่อย A3025-3030',
    '089': 'ทางเดินเข้าห้องเรียนกลุ่มย่อย A3031-3034',
    '090': 'ทางเดินห้องเรียน A3004-A3005',
    '091': 'สำนักงานงานบริการการศึกษา','092': 'ห้องพักผู้ทรงคุณวุฒิ',
    '093': 'Operating Room',             '094': 'Dialogue Room 1',
    '095': 'Dialogue Room 2',            '096': 'Dialogue Room 3',
    '097': 'Dialogue Room 4',            '098': 'Dialogue Room 5',
    '099': 'ห้อง Control',
    '100': 'ห้องงานประกันคุณภาพการศึกษา',
    '101': 'RA (Echo Lab)',              '102': 'RV (Cath Sim)',
    '103': 'LA (OR Sim)',                '104': 'LV (ICU Sim)',
    '105': 'CVS 3',                      '106': 'ห้องเปลี่ยนเสื้อผ้า',
    '107': 'Ao',                         '108': 'stock I',
    '109': 'research and innovation laboratory 5',
    '110': 'Washing Room',               '111': 'lounge',
    '112': 'Dialogue & Meeting Space office',
    '113': 'ห้อง 257',                  '114': 'ห้องประชุมกายวิภาค',
    '454': 'ห้องเก็บตัวอย่าง/อุปกรณ์ กายวิภาค',
    '455': 'ห้องอเนกประสงค์กายวิภาค',
    '456': 'ห้องเจ้าหน้าที่วิภาค',
    '457': 'ห้องปฏิบัติการ 1 กายวิภาค',
    '458': 'ห้องปฏิบัติการ 2 กายวิภาค',
    '459': 'ห้องปฏิบัติการ 3 กายวิภาค',
    '460': 'ห้องเรียนกายวิภาค',
    '461': 'research and innovation laboratory 4',
    '462': 'research and innovation laboratory 1',
    '463': 'research and innovation laboratory 2',
    '464': 'research and innovation laboratory 3',
    '465': 'ห้องประชุม 10-1',
    '466': 'ห้องประชุม 10-2'
  },

  /** โครงการจากชีต asset คอลัมน์ AA — ไม่ pad zero */
  project: {
    '1':    'ส่วนกลาง',
    '2':    'แพทย์บูรณาการ',
    '3':    'ตจวิทยา',
    '4':    'ทันตแพทยศาสตร์',
    '5':    'ทัศนมาตรศาสตร์',
    '6':    'เทคโนโลยีหัวใจและทรวงอก',
    '7':    'แพททย์แผนจีน'
  },

  /** หน่วยงานเดิมจาก asset คอลัมน์ AB — ไม่ pad zero */
  department: {
    '68':   'รพ.ธรรมศาสตร์เฉลิมพระเกียรติ',
    '1001': 'ผู้มีความรู้ความสามารถพิเศษ',
    '1014': 'คณะแพทยศาสตร์',
    '50':   'งานแผนและงบประมาณ',
    '42':   'งานบริหารทรัพยากรมนุษย์',
    '43':   'งานบริหารทั่วไป',
    '1016': 'สำนักวิชาแพทยศาสตร์',
    '35':   'สำนักงานเลขานุการ',
    '44':   'งานคลังและพัสดุ',
    '51':   'งานบริการการศึกษา',
    '1013': 'งานประกันคุณภาพการศึกษา',
    '54':   'งานเทคโนโลยีสารสนเทศ',
    '1017': 'วิทยาลัยแพทยศาสตร์นานาชาติจุฬาภรณ์',
    '1015': 'สำนักวิชาการแพทย์บูรณาการ',
    '1020': 'งานวิเทศสัมพันธ์',
    '1021': 'งานกิจการนักศึกษา',
    '1022': 'งานวิจัย',
    '1023': 'โครงการบริการวิชาการฯ(แผนจีน)',
    '1024': 'งานบัณฑิตศึกษา',
    '1025': 'งานยุทธศาสตร์และงบประมาณ',
    '1026': 'งานส่งเสริมและสนับสนุนการศึกษา',
    '1027': 'งานบริหาร',
    '1028': 'งานวิจัย ส่งเสริมและพัฒนาวิชาการ',
    '1029': 'งานวิจัยอื่นๆ',
    '1030': 'ศูนย์สอบ',
    '1031': 'Skill Lab',
    '1032': 'สถาบันนวัตกรรมวิทยาศาสตร์การแพทย์และวิศวกรรม',
    '1033': 'คณะวิทยาศาสตร์และเทคโนโลยี',
    '1034': 'งานประชาสัมพันธ์'
  },

  funding: {
    '01': 'กองทุนค่าธรรมเนียมการศึกษา',
    '02': 'งบคลัง',
    '03': 'งบวิทยาลัยแพทยศาสตร์นานาชาติจุฬาภรณ์',
    '04': 'งบบริจาค',
    '05': 'งบกองทุนวิจัย'
  },

  status: {
    '01': 'ปกติ',
    '02': 'ชำรุด',
    '03': 'จำหน่ายพัสดุ ตามระเบียบ ฯ ข้อ 215 วรรคหนึ่ง (1) (ก) โดยวิธีเฉพาะเจาะจง',
    '04': 'รอจำหน่าย',
    '05': 'หาไม่เจอ',
    '06': 'ครุภัณฑ์โอน',
    '07': 'จำหน่ายพัสดุ ตามระเบียบ ฯ ข้อ 215 วรรคหนึ่ง (1) (ค) การขายอุปกรณ์อิเล็กทรอนิกส์',
    '08': 'จำหน่ายพัสดุ ตามระเบียบ ฯ ข้อ 215 วรรคหนึ่ง (3) วิธีโอนแก่หน่วยงานรัฐ หรือองค์การสถานสาธารณกุศล ตามมาตรา 47 (7) แห่งประมวลรัษฎากร',
    '09': 'จำหน่ายพัสดุ ตามระเบียบ ฯ ข้อ 215 วรรคหนึ่ง (4) วิธีแปรสภาพหรือทำลาย',
    '10': 'จำหน่ายพัสดุ ตามระเบียบ ฯ ข้อ 217 วิธีจำหน่ายเป็นสูญ',
    '11': 'ส่งคืนพัสดุ',
    '12': 'ตัดจำหน่าย เนื่องจากเกิดข้อผิดพลาดจากการบันทึกข้อมูล'
  }
});

/** ความยาว pad zero ของแต่ละ map (0 = ไม่ pad) */
const TRANSLATION_PAD = Object.freeze({
  category: 2, unit: 3, building: 3, room: 3,
  project: 0, department: 0, funding: 2, status: 2
});

/* ── Translators ─────────────────────────────────────────── */
const Translators = {
  translate(code, cache, mapName) {
    if (code === null || code === undefined || code === '') return '';

    const num = parseInt(code, 10);
    if (isNaN(num)) {
      return ['department', 'project'].includes(mapName)
        ? `ไม่พบข้อมูล (${code})` : `ไม่พบรหัส (${code})`;
    }

    const padLen  = TRANSLATION_PAD[mapName];
    const codeStr = padLen > 0 ? String(num).padStart(padLen, '0') : String(num);

    const cacheBucket = cache[mapName] ?? (cache[mapName] = {});
    if (cacheBucket[codeStr] !== undefined) return cacheBucket[codeStr];

    const map   = TRANSLATION_MAPS[mapName];
    const found = map?.[codeStr];

    let result;
    if (found !== undefined) {
      result = found;
    } else {
      Stats.bump('unknownCode');
      result = ['department', 'project'].includes(mapName)
        ? `ไม่พบข้อมูล (${codeStr})`
        : `ไม่พบรหัส (${codeStr})`;
    }

    cacheBucket[codeStr] = result;
    return result;
  },

  category:   (c, k) => Translators.translate(c, k, 'category'),
  unit:       (c, k) => Translators.translate(c, k, 'unit'),
  building:   (c, k) => Translators.translate(c, k, 'building'),
  room:       (c, k) => Translators.translate(c, k, 'room'),
  project:    (c, k) => Translators.translate(c, k, 'project'),
  department: (c, k) => Translators.translate(c, k, 'department'),
  funding:    (c, k) => Translators.translate(c, k, 'funding'),
  status:     (c, k) => Translators.translate(c, k, 'status')
};

/** ตัวนับ warning แบบไม่ท่วม log (เดิมใช้ Logger_.warn ทุกครั้ง) */
const Stats = {
  _c: {},
  reset() { this._c = {}; },
  bump(k, n = 1) { this._c[k] = (this._c[k] || 0) + n; },
  get(k) { return this._c[k] || 0; },
  all() { return { ...this._c }; }
};

/* ── StringHelper ────────────────────────────────────────── */
const StringHelper = {
  safe(val) {
    if (val === null || val === undefined || val === '') return '';
    const s = String(val).trim();
    if (/^null$/i.test(s)) return '';
    return s;
  },

  /** รวมข้อความ 4 ส่วน + pad ให้ครบ 6 บรรทัด */
  formatCombined(d, e, x, y) {
    const TARGET_LINES = 6;
    const parts = [d, e, x, y]
      .filter(v => v !== null && v !== undefined && v !== '')
      .map(String);

    if (parts.length === 0) return '\n'.repeat(TARGET_LINES - 1);

    const combined = parts.join('\n');
    const needed   = Math.max(0, TARGET_LINES - parts.length);
    return combined + (needed > 0 ? '\n'.repeat(needed) : '');
  },

  /** ตัด ' นำหน้า (Google Sheets text-forcing) — xlsx ไม่ต้องใช้ */
  stripQuote(val) {
    if (typeof val !== 'string') return val;
    return val.charAt(0) === "'" ? val.slice(1) : val;
  }
};

/* ── DateFormatter ───────────────────────────────────────── *
 *  แทน Utilities.formatDate(date, tz, 'dd/MM/yyyy') ด้วย JS ล้วน
 *  ⚠️ ไม่แปลง พ.ศ. → ค.ศ. โดยพลการ: ปีอะไรเข้ามาก็คงปีนั้นไว้
 *     (รายงาน ERP ใช้ พ.ศ. เช่น 01/10/2565 — ถ้าแปลงจะผิด)
 * ============================================================ */
const DateFormatter = {
  _pad2(n) { return String(n).padStart(2, '0'); },

  /** dd/MM/yyyy จาก Date object (ใช้ local time เหมือน tz ของ spreadsheet) */
  _fmt(d) {
    return `${this._pad2(d.getDate())}/${this._pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
  },

  format(val, cache, _timeZone, addPrefix = false) {
    if (val === null || val === undefined || val === '') return '';

    let dateObj  = null;
    let cacheKey = null;

    if (val instanceof Date) {
      if (isNaN(val.getTime())) return '';
      cacheKey = `${val.getFullYear()}-${this._pad2(val.getMonth() + 1)}-${this._pad2(val.getDate())}`;
      dateObj  = val;

    } else if (typeof val === 'number') {
      // SheetJS serial date (กรณีอ่านด้วย cellDates:false)
      const parsed = XLSX.SSF ? XLSX.SSF.parse_date_code(val) : null;
      if (!parsed) return String(val);
      dateObj  = new Date(parsed.y, parsed.m - 1, parsed.d);
      cacheKey = `n${val}`;

    } else if (typeof val === 'string') {
      const s = val.trim();
      if (!s || /^null$/i.test(s)) return '';
      cacheKey = s;

      const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
      const m4 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

      if (m2) {
        dateObj = this._mk(FormulaEngine.expandYear2(m2[3]), Number(m2[2]), Number(m2[1]));
      } else if (m4) {
        dateObj = this._mk(Number(m4[3]), Number(m4[2]), Number(m4[1]));
      } else {
        const attempt = new Date(s);
        if (!isNaN(attempt.getTime())) {
          dateObj = attempt;
        } else {
          Stats.bump('dateParseFail');
          return addPrefix ? `'${s}` : s;
        }
      }
    } else {
      Stats.bump('dateUnknownType');
      return '';
    }

    if (!dateObj || isNaN(dateObj.getTime())) return '';

    const dateCache = cache.date ?? (cache.date = {});
    if (cacheKey && dateCache[cacheKey] !== undefined) {
      const cached = dateCache[cacheKey];
      return addPrefix ? `'${cached}` : cached;
    }

    const formatted = this._fmt(dateObj);
    if (cacheKey) dateCache[cacheKey] = formatted;
    return addPrefix ? `'${formatted}` : formatted;
  },

  /** new Date(y,…) จะแปลงปี 0–99 เป็น 19xx — ต้องกันไว้ */
  _mk(y, m, d) {
    const dt = new Date(y, m - 1, d);
    if (y >= 0 && y < 100) dt.setFullYear(y);
    return dt;
  }
};

/* ── AssetInputSchema — ป้องกันผลกระทบเมื่อเพิ่มคอลัมน์ AA ── */
const AssetInputSchema = {
  _normaliseHeader(value) {
    return String(value ?? '').trim().toLowerCase().replace(/[\s_.()\-]/g, '');
  },

  /**
 * ใช้ index มาตรฐาน A–AH เป็นค่าเริ่มต้น และแทนที่ด้วยตำแหน่งจริงเมื่อ
 * ไฟล์มีหัวตารางที่รู้จัก จึงรองรับทั้งไฟล์ใหม่ (มี AA = project_id)
 * และไฟล์เก่าที่ยังไม่มีคอลัมน์นี้ได้โดยไม่ทำให้ AB–AH อ่านเหลื่อมกัน
   */
  resolve(headerRow) {
    const columns = { ...ASSET_INPUT.columns };
    if (!Array.isArray(headerRow)) return { columns, matched: 0, projectFound: false };

    const headerIndex = new Map();
    headerRow.forEach((cell, index) => {
      const key = this._normaliseHeader(cell);
      if (key && !headerIndex.has(key)) headerIndex.set(key, index);
    });

    let matched = 0;
    let projectFound = false;
    for (const [field, aliases] of Object.entries(ASSET_INPUT.headerAliases)) {
      const index = aliases
        .map(alias => headerIndex.get(this._normaliseHeader(alias)))
        .find(index => index !== undefined);
      if (index === undefined) continue;
      columns[field] = index;
      matched++;
      if (field === 'project_id') projectFound = true;
    }
    // หากหัวตารางระบุฟิลด์อื่นแล้ว แต่ไม่มี project_id แสดงว่าเป็นไฟล์เก่า
    // AA ของไฟล์นั้นคือ section_id จึงต้องไม่ fallback ไปอ่านตำแหน่ง 26 เป็นโครงการ
    if (matched > 0 && !projectFound) columns.project_id = -1;
    return { columns, matched, projectFound };
  },

  read(row, schema) {
    const source = Array.isArray(row) ? row : [];
    const columns = schema?.columns ?? ASSET_INPUT.columns;
    return Object.fromEntries(Object.keys(ASSET_INPUT.columns)
      .map(field => [field, source[columns[field]]]));
  }
};

/* ── DataProcessor: data → information (34 คอลัมน์) ──────── */
const DataProcessor = {

  /**
   * @param {Array<Array>} dataValues  AoA จากไฟล์ data
   * @param {function}     onProgress
   * @returns {Promise<{rows:Array<Array>, errorRows:number[], execSec:string}>}
   */
  async run(dataValues, onProgress) {
    const t0 = performance.now();

    if (!dataValues || dataValues.length === 0) {
      throw new Error(`ไม่พบข้อมูลในไฟล์ "${CONFIG.sheets.data}"`);
    }

    let values = dataValues;
    let schema = AssetInputSchema.resolve(null);
    if (CONFIG.options.skipDataHeaderRow && this._looksLikeHeader(values[0])) {
      schema = AssetInputSchema.resolve(values[0]);
      Logger_.info(`   ↳ ตรวจพบแถวหัวตาราง — ข้ามแถวที่ 1 ("${String(values[0][0]).slice(0, 20)}…")`);
      if (schema.matched) {
        Logger_.info(`   ↳ จับคู่หัวคอลัมน์ asset ได้ ${schema.matched} ฟิลด์` +
                     (schema.projectFound ? ' · พบ project_id สำหรับ AA' :
                      ' · ไม่พบ project_id: ข้ามคอลัมน์โครงการ'));
      }
      values = values.slice(1);
    }

    const cache = {
      category: {}, unit: {}, building: {}, room: {}, project: {},
      department: {}, funding: {}, status: {}, date: {}
    };
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;

    const rows      = [];
    const errorRows = [];
    const total     = values.length;

    await Async.eachChunk(values, CONFIG.ui.chunkSize, (row, i) => {
      try {
        rows.push(this.processRow(row, cache, tz, schema));
      } catch (e) {
        Logger_.error(`[Row ${i + 1}] ${e.message}`);
        errorRows.push(i + 1);
      }
    }, done => onProgress?.(done, total));

    const execSec = ((performance.now() - t0) / 1000).toFixed(1);
    return { rows, errorRows, execSec };
  },

  /** @private ตรวจว่าแถวแรกเป็นหัวตารางหรือไม่ */
  _looksLikeHeader(row) {
    if (!row) return false;
    const first = String(row[0] ?? '').trim().toLowerCase();
    if (first === '' ) return false;
    if (/^\d+(\.0+)?$/.test(first)) return false;         // เป็นตัวเลข = ข้อมูลจริง
    return ['id', 'ลำดับ', 'no', 'no.', '#'].includes(first) || isNaN(Number(first));
  },

  /**
   * ตรรกะเหมือน GAS ทุกบรรทัด
   * @param {Array} row
   * @returns {Array} 34 elements
   */
  processRow(row, cache, timeZone, schema) {
    const {
      id, receive_date, type_asset_code, asset_name, asset_brand,
      asset_model, asset_color, asset_donate, asset_detail, asset_no,
      asset_no_erp, asset_no_old, serial_no, group_code, class_code,
      type_code, unit_code, amount, price, seller_name, warranty_start,
      warranty_end, build_code, floor_code, room_code, project_id, section_id,
      budget_code, status_code, budget_year, remark_detail, remark_section, istatus
    } = AssetInputSchema.read(row, schema);

    const colB = DateFormatter.format(receive_date, cache, timeZone);

    const typeNum = Number(type_asset_code);
    const colC = typeNum === 1 ? 'ครุภัณฑ์'
               : typeNum === 2 ? 'วัสดุคงทน' : '';

    const colD = StringHelper.safe(asset_name);
    const colE = StringHelper.safe(asset_brand);
    const colF = StringHelper.safe(asset_model);
    const colG = StringHelper.safe(asset_color);

    const colH = (asset_donate === 'Y') ? 'เป็นครุภัณฑ์บริจาค' : '';

    const rawDetail = StringHelper.safe(asset_detail);
    const colI = rawDetail ? `'${rawDetail}` : '';

    const colJ = asset_no ? `'${asset_no}` : '';

    const rawErp = StringHelper.safe(asset_no_erp);
    const colK = rawErp ? `'${rawErp}` : '';

    const rawOld = StringHelper.safe(asset_no_old);
    const colL = rawOld ? `'${rawOld}` : '';

    const rawSerial = StringHelper.safe(serial_no);
    const colM = rawSerial ? `'${rawSerial}` : '';

    const colN = StringHelper.safe(group_code);
    const colO = StringHelper.safe(class_code);
    const colP = StringHelper.safe(type_code);

    const colQ = Translators.category(group_code, cache);
    const colR = Translators.unit(unit_code, cache);

    const colS = (amount !== null && amount !== undefined && amount !== '') ? amount : '';
    const colT = (price  !== null && price  !== undefined && price  !== '') ? price  : '';

    const colU = StringHelper.safe(seller_name);

    const colV = DateFormatter.format(warranty_start, cache, timeZone, true);
    const colW = DateFormatter.format(warranty_end,   cache, timeZone, true);

    const colX = Translators.building(build_code, cache);
    const colY = StringHelper.safe(floor_code);
    const colZ = Translators.room(room_code, cache);

    const colAA = Translators.department(section_id, cache);
    const colAB = Translators.funding(budget_code, cache);
    const colAC = Translators.status(status_code, cache);
    const colAD = StringHelper.safe(budget_year);
    const colAE = StringHelper.safe(remark_detail);
    const colAF = StringHelper.safe(remark_section);
    const colAH = Translators.project(project_id, cache);

    let colAG;
    switch (String(istatus).trim()) {
      case 'Y':   colAG = 'ใช้งานอยู่';   break;
      case 'N':   colAG = 'ไม่ได้ใช้งาน'; break;
      case 'Del': colAG = 'ลบ';           break;
      default:    colAG = 'ไม่พบสถานะ';   break;
    }

    return [
      id, colB, colC, colD, colE, colF, colG, colH, colI, colJ,
      colK, colL, colM, colN, colO, colP, colQ, colR, colS, colT,
      colU, colV, colW, colX, colY, colZ, colAA, colAB, colAC, colAD,
      colAE, colAF, colAG, colAH
    ];
  }
};


/* ============================================================
 *  PART C — information → target
 * ============================================================ */

/* ── KeyNormalizer ───────────────────────────────────────── */
const KeyNormalizer = {
  normalize(val) {
    if (val === null || val === undefined || val === '') return '';
    let s = String(val);
    s = s.replace(/\.0+$/, '');
    s = s.trim().replace(/\s+/g, ' ');
    return s;
  },

  makeKeys(val) {
    const primary = this.normalize(val).toLowerCase();
    const compact = primary.replace(/[\s\-_.]/g, '');
    const noZeros = primary
      .split(/([\-_/])/)
      .map(part => /^\d+$/.test(part) ? String(parseInt(part, 10) || 0) : part)
      .join('');
    return { primary, compact, noZeros };
  }
};

/* ── UrlNormalizer (v4.1 logic) ──────────────────────────── */
const UrlNormalizer = {
  normalize(raw) {
    if (raw === null || raw === undefined) return '';
    const s = String(raw).trim();
    if (!s) return '';

    if (/^https?:\/\//i.test(s)) return this._convertDriveUrl(s);
    if (/^www\./i.test(s))       return `https://${s}`;
    if (/\.(jpe?g|png|gif|webp|bmp|svg)$/i.test(s)) return this._prependBaseUrl(s);
    if (/^[A-Za-z0-9_-]{20,}$/.test(s)) return `https://drive.google.com/file/d/${s}/view`;

    return '';
  },

  isValid(url) { return /^https?:\/\//i.test(String(url || '').trim()); },

  _prependBaseUrl(fileName) {
    const base = CONFIG.image.baseUrl || '';
    if (!base) { Stats.bump('noBaseUrl'); return ''; }
    const cleanName = fileName.replace(/^\/+/, '');
    const sep = base.endsWith('/') ? '' : '/';
    return base + sep + cleanName;
  },

  _convertDriveUrl(url) {
    if (!/drive\.google\.com/i.test(url)) return url;
    const m = url.match(/\/file\/d\/([A-Za-z0-9_-]+)/)
            ?? url.match(/[?&]id=([A-Za-z0-9_-]+)/);
    return m ? `https://drive.google.com/file/d/${m[1]}/view` : url;
  }
};

/* ── ImageCache ──────────────────────────────────────────── */
class ImageCache {
  constructor() {
    this._primary = new Map();
    this._compact = new Map();
    this._noZeros = new Map();
    this._stats = {
      totalRows: 0, indexedRows: 0, emptyAssetRows: 0,
      invalidUrlRows: 0, uniqueAssets: 0
    };
  }

  clear() {
    this._primary.clear(); this._compact.clear(); this._noZeros.clear();
    this._stats = { totalRows:0, indexedRows:0, emptyAssetRows:0, invalidUrlRows:0, uniqueAssets:0 };
  }

  add(assetNoRaw, url) {
    const { primary, compact, noZeros } = KeyNormalizer.makeKeys(assetNoRaw);
    if (!primary) return;
    this._addTo(this._primary, primary, url);
    this._addTo(this._compact, compact, url);
    this._addTo(this._noZeros, noZeros, url);
  }

  lookup(assetNoRaw) {
    const { primary, compact, noZeros } = KeyNormalizer.makeKeys(assetNoRaw);
    if (!primary) return { urls: [], level: 'NONE' };

    const p = this._primary.get(primary);
    if (p?.length) return { urls: p, level: 'PRIMARY' };
    const c = this._compact.get(compact);
    if (c?.length) return { urls: c, level: 'COMPACT' };
    const n = this._noZeros.get(noZeros);
    if (n?.length) return { urls: n, level: 'NO-ZEROS' };

    return { urls: [], level: 'NONE' };
  }

  get(assetNoRaw) { return this.lookup(assetNoRaw).urls.join(' , '); }

  getStats()          { return { ...this._stats, uniqueAssets: this._primary.size }; }
  recordStats(partial){ Object.assign(this._stats, partial); }

  _addTo(map, key, url) {
    if (!map.has(key)) map.set(key, []);
    const arr = map.get(key);
    if (!arr.includes(url)) arr.push(url);
  }
}

/* ── ImageMapBuilder — เดิมรับ ss, ตอนนี้รับ AoA ─────────── */
const ImageMapBuilder = {
  /**
   * @param {Array<Array>} aoa  แถวข้อมูลของชีต "รูปภาพ" (ไม่รวม header)
   * @returns {ImageCache}
   */
  build(aoa) {
    const cache = new ImageCache();
    if (!aoa || aoa.length === 0) {
      Logger_.warn(`ImageMapBuilder: ไฟล์ "${CONFIG.sheets.image}" ว่าง`);
      return cache;
    }

    const { codeCol, urlCol } = CONFIG.image;
    let emptyAsset = 0, invalidUrl = 0, indexed = 0;
    const samples = [];

    aoa.forEach((row, i) => {
      const rawCode = row[codeCol];
      const rawUrl  = row[urlCol];

      const key = KeyNormalizer.normalize(rawCode);
      if (!key) { emptyAsset++; return; }

      const url = UrlNormalizer.normalize(rawUrl);
      if (!url) {
        invalidUrl++;
        if (samples.length < 5 && rawUrl) samples.push(`แถว ${i + 2}: "${rawUrl}"`);
        return;
      }

      cache.add(rawCode, url);
      indexed++;
    });

    cache.recordStats({
      totalRows: aoa.length, indexedRows: indexed,
      emptyAssetRows: emptyAsset, invalidUrlRows: invalidUrl
    });

    if (samples.length) Logger_.warn(`   ↳ URL ไม่ valid (ตัวอย่าง): ${samples.join(' · ')}`);
    Logger_.info(`   ↳ ImageCache: ${cache.getStats().uniqueAssets} keys · ${indexed} URLs · ` +
                 `ว่าง ${emptyAsset} · ไม่ valid ${invalidUrl}`);
    return cache;
  }
};

/* ── Validators ──────────────────────────────────────────── */
const Validators = {
  isValidDate(val) {
    if (val instanceof Date) return !isNaN(val.getTime());
    if (typeof val === 'string' && val.trim()) return !isNaN(new Date(val).getTime());
    return false;
  },

  hasValue(val) {
    return val !== null && val !== undefined && String(val).trim() !== '';
  },

  validateAssetRow(srcRow) {
    const errors = [], warnings = [];

    if (!this.hasValue(srcRow[CONFIG.source.assetNoCol])) errors.push('ลำดับว่าง');
    if (!this.hasValue(srcRow[9])) warnings.push('หมายเลขครุภัณฑ์ว่าง');

    if (this.hasValue(srcRow[1]) && !this.isValidDate(srcRow[1])
        && !/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(String(srcRow[1]))) {
      warnings.push(`วันที่ตรวจรับไม่ valid: "${srcRow[1]}"`);
    }
    if (this.hasValue(srcRow[22]) && !this.isValidDate(srcRow[22])
        && !/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(String(srcRow[22]))) {
      warnings.push(`วันที่สิ้นสุดไม่ valid: "${srcRow[22]}"`);
    }

    return { valid: errors.length === 0, errors, warnings };
  }
};

/* ── FormulaEngine — แทนสูตร Excel ที่เคยฝังลงเซลล์ ─────── *
 *
 *  เดิม _mapFormulas เขียนสูตรลงไป 43 เซลล์ต่อแถว:
 *    AB : DATEDIF(B, AA, "Y"/"YM"/"MD")
 *    AC : XLOOKUP ซ้อน 4 ชั้น G→P, G→C, F→C, E→C (ดู CONFIG.erp.acLookupChain)
 *    AD…: IFNA(VLOOKUP(AC, ERP!C:AR, n, FALSE), "")  × 41
 *
 *  ทั้งหมดถูกคำนวณเป็น "ค่า" ด้วย JS แทน ทำให้
 *  ไฟล์ผลลัพธ์เปิดได้ทุกโปรแกรม ไม่ต้อง recalc และไม่พึ่ง XLOOKUP
 * ============================================================ */
const FormulaEngine = {

  /** parse "dd/MM/yyyy" | "dd/MM/yy" | Date → {y,m,d} (คงปีตามที่ให้มา) */
  parseDMY(val) {
    if (val === null || val === undefined || val === '') return null;
    if (val instanceof Date && !isNaN(val.getTime())) {
      return { y: val.getFullYear(), m: val.getMonth() + 1, d: val.getDate() };
    }
    const s = StringHelper.stripQuote(String(val).trim());
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (!m) return null;
    const y = String(m[3]).length === 2 ? this.expandYear2(m[3]) : Number(m[3]);
    return { y, m: Number(m[2]), d: Number(m[1]) };
  },

  /**
   * ปี 2 หลัก → 4 หลัก ด้วย pivot (ค่าเริ่มต้น 50)
   *   '00'–'50' → 2000–2050   ·   '51'–'99' → 1951–1999
   * เดิมบวก 2000 เสมอ ทำให้ครุภัณฑ์เก่าที่กรอก '98 กลายเป็นปี 2098
   * และ DATEDIF คำนวณอายุผิดทั้งแถว
   * ⚠️ ข้อมูลชุดปัจจุบันไม่มีปี 2 หลักเลย (0 ค่า) — นี่คือกันไว้ล่วงหน้า
   *    และ bump สถิติทุกครั้งที่เจอ เพื่อให้เห็นถ้ามีข้อมูลแบบนี้เข้ามา
   */
  YEAR_PIVOT: 50,
  expandYear2(yy) {
    const n = Number(yy);
    Stats.bump('year2Digit');
    return n <= this.YEAR_PIVOT ? 2000 + n : 1900 + n;
  },

  /** จำนวนวันจาก 1899-12-30 ถึง 1970-01-01 (ฐานของ Excel serial) */
  EXCEL_EPOCH_OFFSET: 25569,

  /**
   * แปลง 'dd/MM/yyyy' → Excel serial number (สำหรับเขียนเป็นวันที่จริง)
   * คืน null ถ้าไม่ใช่รูปแบบนี้ หรือปี ≥ ceYearMax (= เป็น พ.ศ. ห้ามแปลง)
   *
   * ⚠️ ทำไมไม่ส่ง Date object ให้ SheetJS แปลงเอง
   *    SheetJS อิงฐาน new Date(1899,11,30) ตาม timezone ของเครื่อง
   *    ซึ่งกรุงเทพฯ ปี 1899 ยังใช้ LMT +06:42:04 (ไม่ใช่ +07:00)
   *    ผลคือ serial คลาดไป ~4 วินาที (45306.0000462963 แทน 45306)
   *    การคำนวณจาก Date.UTC เองได้เลขจำนวนเต็มเป๊ะทุก timezone
   */
  toExcelSerial(val) {
    const p = this.parseDMY(val);
    if (!p) return null;
    if (p.y >= CONFIG.ceYearMax) return null;      // พ.ศ. — ปล่อยเป็นข้อความ

    const ms = Date.UTC(p.y, p.m - 1, p.d);
    if (isNaN(ms)) return null;
    const back = new Date(ms);                      // ตรวจว่าวันที่มีจริง (31/02 ต้องตก)
    if (back.getUTCFullYear() !== p.y || back.getUTCMonth() !== p.m - 1 ||
        back.getUTCDate() !== p.d) return null;

    return ms / 86400000 + this.EXCEL_EPOCH_OFFSET;
  },

  _daysInMonth(y, m) { const d = new Date(2000, m, 0); d.setFullYear(y); return d.getDate(); },

  /** เทียบว่า a < b หรือไม่ (ไม่ผ่าน Date object เพื่อรองรับปี พ.ศ.) */
  _lt(a, b) {
    if (a.y !== b.y) return a.y < b.y;
    if (a.m !== b.m) return a.m < b.m;
    return a.d < b.d;
  },

  /** เลขวันแบบต่อเนื่อง (ใช้ UTC เพื่อกัน DST) */
  _serial(p) { return Date.UTC(p.y, p.m - 1, p.d) / 86400000; },

  /** บวกเดือนแบบหนีบวันสิ้นเดือน: 31/01 + 1 เดือน = 29/02 */
  _addMonths(p, n) {
    const total = (p.y * 12) + (p.m - 1) + n;
    const y = Math.floor(total / 12);
    const m = (total % 12) + 1;
    return { y, m, d: Math.min(p.d, this._daysInMonth(y, m)) };
  },

  /**
   * เทียบเท่า DATEDIF(start, end, "Y") / ("YM") / ("MD")
   *
   * ⚠️ หมายเหตุการย้ายระบบ
   * DATEDIF ตัวจริงของ Excel/Sheets มีบั๊กที่รู้จักกันดีในโหมด "MD":
   * DATEDIF("31/01/2024","01/03/2024","MD") คืน -1 เพราะยืมวันจาก
   * "เดือนก่อนหน้าวันสิ้นสุด" (ก.พ. = 29 วัน) แล้วยังติดลบ
   *
   * เวอร์ชันนี้ใช้นิยามที่สอดคล้องกันเสมอแทน:
   *   ปี/เดือน = จำนวนเดือนเต็มที่ผ่านไป
   *   วัน      = ระยะจาก (start + ปี + เดือน) ถึง end  →  ไม่มีทางติดลบ
   * ผลลัพธ์ตรงกับ DATEDIF ทุกกรณีปกติ ต่างเฉพาะเคสบั๊กข้างต้น
   *
   * @returns {{y:number,m:number,d:number}|null}
   */
  datedif(start, end) {
    if (!start || !end) return null;
    if (this._lt(end, start)) return null;   // DATEDIF จะ error ถ้า end < start

    let y = end.y - start.y;
    let m = end.m - start.m;
    if (end.d < start.d) m -= 1;
    if (m < 0) { y -= 1; m += 12; }

    const anchor = this._addMonths(start, y * 12 + m);
    const d = this._serial(end) - this._serial(anchor);

    return { y, m, d: Math.max(0, d) };
  },

  /** AB: "X ปี Y เดือน Z วัน" — คืน '' ถ้าช่องใดว่าง (ตาม IF(OR(...))) */
  ageText(bVal, aaVal) {
    if (!Validators.hasValue(bVal) || !Validators.hasValue(aaVal)) return '';
    const r = this.datedif(this.parseDMY(bVal), this.parseDMY(aaVal));
    if (!r) return '';
    return `${r.y} ปี ${r.m} เดือน ${r.d} วัน`;
  }
};

/* ── ErpIndex — แทน XLOOKUP / VLOOKUP บนชีต ERP ─────────── */
class ErpIndex {
  /**
   * @param {Array<Array>} rows  แถวข้อมูล ERP 44 คอลัมน์ (ไม่รวม header)
   */
  constructor(rows) {
    this.rows    = rows || [];
    this._byRef1 = new Map();   // P (idx 15) → รหัสสินทรัพย์ (idx 2)
    this._byCode = new Map();   // C (idx 2)  → ทั้งแถว
    this._build();
  }

  /** normalize + ตัด ' นำหน้า (ค่าจาก processRow ถูกบังคับเป็น text ด้วย ') */
  static key(v) {
    return KeyNormalizer.normalize(StringHelper.stripQuote(
      v === null || v === undefined ? '' : String(v)
    )).toLowerCase();
  }

  _build() {
    const { codeCol, refNo1Col } = CONFIG.erp;
    for (const row of this.rows) {
      const ref  = ErpIndex.key(row[refNo1Col]);
      const code = KeyNormalizer.normalize(StringHelper.stripQuote(String(row[codeCol] ?? '')));

      // XLOOKUP / VLOOKUP คืน "แถวแรกที่เจอ" → ไม่เขียนทับของเดิม
      if (ref  && !this._byRef1.has(ref))              this._byRef1.set(ref, code);
      if (code && !this._byCode.has(code.toLowerCase())) this._byCode.set(code.toLowerCase(), row);
    }
  }

  /** XLOOKUP(x, ERP!P:P, ERP!C:C, "") — ค้นด้วย Ref No. 1 */
  lookupCodeByRef1(x) {
    const k = ErpIndex.key(x);
    if (!k) return '';
    return this._byRef1.get(k) ?? '';
  }

  /**
   * XLOOKUP(x, ERP!C:C, ERP!C:C, "") — ค้นด้วยรหัสสินทรัพย์เอง
   * ใช้ยืนยันว่าค่าที่ถืออยู่มีจริงใน ERP แล้วคืน "รูปแบบตามต้นฉบับ ERP"
   */
  lookupCodeByCode(x) {
    const k = ErpIndex.key(x);
    if (!k) return '';
    const row = this._byCode.get(k);
    if (!row) return '';
    return KeyNormalizer.normalize(StringHelper.stripQuote(String(row[CONFIG.erp.codeCol] ?? '')));
  }

  /** IFNA(VLOOKUP(AC, ERP!C:AR, n, FALSE), "") → คืนทั้งแถว */
  rowByCode(code) {
    const k = ErpIndex.key(code);
    if (!k) return null;
    return this._byCode.get(k) ?? null;
  }

  get size() { return this._byCode.size; }
  get refSize() { return this._byRef1.size; }
}

/* ── Mapper: information (34) → target (75) ─────────────── */
const Mapper = {
  /**
   * @param {Array}      src
   * @param {number}     rowNum     แถวจริงในชีตปลายทาง (ใช้อ้างอิงตอน log)
   * @param {ImageCache} cache
   * @param {ErpIndex}   erp
   * @returns {Array} 75 elements
   */
  mapRow(src, rowNum, cache, erp) {
    const row = new Array(75).fill('');
    this._mapBasicFields(row, src);
    this._mapImageField(row, src, cache);
    this._mapDateField(row, src);
    this._mapComputed(row, erp);      // ← เดิมคือ _mapFormulas
    return row;
  },

  /** @private — mapping เดิมทุกช่อง */
  _mapBasicFields(row, src) {
    row[0]  = src[0];   row[1]  = src[1];   row[2]  = src[2];   row[3]  = src[3];
    row[4]  = src[8];   row[5]  = src[9];   row[6]  = src[11];  row[7]  = src[4];
    row[8]  = src[5];   row[9]  = src[6];   row[10] = src[7];   row[11] = '';
    row[12] = src[19];  row[13] = src[20];  row[14] = src[23];  row[15] = src[24];
    row[16] = src[25];  row[17] = src[26];  row[18] = src[27];  row[19] = src[28];
    row[20] = src[29];  row[21] = src[30];  row[22] = src[31];  row[23] = src[32];
    row[24] = '';
    // BV (index 73) — คอลัมน์ใหม่ต่อจาก 'เอกสารอ้างอิง'
    // src[10] = information คอลัมน์ K = data.asset_no_erp
    row[TGT.erpNo] = src[INFO.erpNo];
    // BW (index 74) = project ที่แปลจาก data asset คอลัมน์ AA
    row[TGT.project] = src[INFO.project];
  },

  /** @private */
  _mapImageField(row, src, cache) {
    row[TGT.image] = cache.get(src[CONFIG.source.assetNoCol]);
  },

  /** @private */
  _mapDateField(row, src) {
    row[TGT.endDate] = src[22] || '';
  },

  /**
   * @private
   * ตรวจว่าค่าตรงรูปแบบรหัสสินทรัพย์หรือไม่
   *   prefix "27-" · ตัวเลข 21 หลัก (2+2+4+3+4+6) · ขีดกลาง 5 · รวม 26 อักขระ
   *
   * ⚠️ ตัด ' นำหน้าออกก่อนวัดความยาว — ถ้าไม่ตัด LEN จะเกินไป 1 เสมอ
   *    และ trim ช่องว่างหัวท้าย (ต่างจาก LEN ของ Excel เล็กน้อย แต่สอดคล้อง
   *    กับการ normalize ที่ใช้ตอน lookup)
   *
   * @param {*} value
   * @param {{prefix?:string, length?:number, digits?:number, regex?:RegExp}} [guard]
   * @returns {boolean} true = ผ่าน ให้ทำชั้นนี้
   */
  _passGuard(value, guard) {
    if (!guard) return true;
    const v = StringHelper.stripQuote(
      value === null || value === undefined ? '' : String(value)
    ).trim();
    if (guard.prefix && !v.startsWith(guard.prefix)) return false;
    if (guard.length !== undefined && v.length !== guard.length) return false;
    if (guard.digits !== undefined &&
        (v.match(/\d/g) || []).length !== guard.digits) return false;
    if (guard.regex && !guard.regex.test(v)) return false;
    return true;
  },

  /**
   * @private
   * by:'scan' — กวาดทุกคอลัมน์ในแถวเดียวกันจากซ้ายไปขวา
   * หาค่าที่ตรงรูปแบบรหัสสินทรัพย์ แล้ว XLOOKUP ใน ERP!C
   * คืนค่าแรกที่ "ตรงรูปแบบ และมีอยู่จริงใน ERP"
   * @returns {string}
   */
  _scanRow(row, guard, erp) {
    if (!erp) return '';
    const skip = CONFIG.erp.scanSkip;
    for (let i = 0; i < row.length; i++) {
      if (skip.includes(i)) continue;
      const v = row[i];
      if (v === '' || v === null || v === undefined) continue;
      if (!this._passGuard(v, guard)) continue;
      const found = erp.lookupCodeByCode(v);
      if (found !== '') { Stats.bump(`scanHit:col${i}`); return found; }
    }
    return '';
  },

  /** @private ค่าที่ผ่าน normalize + ตัด ' แล้ว (ใช้ตอน by:'self') */
  _clean(v) {
    return KeyNormalizer.normalize(StringHelper.stripQuote(
      v === null || v === undefined ? '' : String(v)
    ));
  },

  /**
   * @private
   * แทน _mapFormulas ของ GAS ด้วยการคำนวณค่าจริง
   */
  _mapComputed(row, erp) {
    // ── ตรวจรูปแบบคอลัมน์ใหม่ BV 'หมายเลขครุภัณฑ์ ERP' (21 หลัก / 26 อักขระ)
    const bv = this._clean(row[TGT.erpNo]);
    if (bv === '')                                Stats.bump('erpNo:ว่าง');
    else if (this._passGuard(bv, ASSET_CODE_GUARD)) Stats.bump('erpNo:ถูกรูปแบบ');
    else                                          Stats.bump('erpNo:ผิดรูปแบบ');

    // AB — DATEDIF(B, AA)
    row[TGT.age] = FormulaEngine.ageText(row[1], row[TGT.endDate]);

    // AC — ไล่ตาม acLookupChain: G→P, G→C, F=รหัส (มี guard), E→C, สุดท้าย ''
    let code = '', via = 'NONE';
    for (const step of CONFIG.erp.acLookupChain) {
      let found = '';

      if (step.by === 'scan') {
        found = this._scanRow(row, step.guard, erp);
      } else {
        const raw = row[step.col];
        if (!this._passGuard(raw, step.guard)) {
          Stats.bump(`acSkip:${step.label}`);
          continue;                                 // = ELSE branch ของ IF()
        }
        if (step.by === 'self')      found = this._clean(raw); // ใช้ค่านั้นเป็นคำตอบเลย
        else if (!erp)               found = '';
        else if (step.by === 'ref1') found = erp.lookupCodeByRef1(raw);
        else                         found = erp.lookupCodeByCode(raw);
      }

      if (found !== '') { code = found; via = step.label; break; }
    }
    row[TGT.erpCode] = code;
    Stats.bump(`ac:${via}`);

    // AD…BR — VLOOKUP(AC, ERP!C:AR, i+2) × 41  → ERP index 3..43
    const erpRow = erp ? erp.rowByCode(code) : null;
    if (erpRow) Stats.bump('erpRowMatched');
    else if (code !== '') Stats.bump('acNoErpRow');   // AC มีค่าแต่ไม่มีแถวใน ERP

    const { firstPullCol, pullCount } = CONFIG.erp;
    for (let i = 0; i < pullCount; i++) {
      const v = erpRow ? erpRow[firstPullCol + i] : '';
      row[TGT.pullStart + i] = (v === null || v === undefined) ? '' : v;
    }
  }
};


/* ============================================================
 *  PART D — copyAssetData() port : RP023 + RP015 → ERP 44 cols
 * ============================================================ */

const ErpBuilder = {

  /** ตรวจว่า AoA ที่ได้เป็นรายงานแบบไหน */
  detect(aoa) {
    const { headerScanRows, headerHints } = CONFIG.erp;
    const norm = c => String(c ?? '').replace(/\s+/g, ' ').trim();

    /**
     * หาแถวหัวตาราง — ยืดหยุ่นกว่าเดิม 3 ทาง:
     *   1) สแกนได้ลึกถึง headerScanRows แถว (เดิม 12)
     *   2) รับคำใบ้ได้หลายคำ ไม่ใช่ 'รหัสสินทรัพย์' คำเดียว
     *   3) เทียบแบบ "มีคำนี้อยู่ในเซลล์" ไม่ต้องตรงเป๊ะทั้งช่อง
     *      (กันกรณี ERP เติมวงเล็บ/หน่วยต่อท้ายชื่อคอลัมน์)
     * ถ้ายังไม่เจอ ใช้ fallback = แถวที่มีเซลล์ไม่ว่างมากที่สุดใน 10 แถวแรก
     */
    const scan = Math.min(headerScanRows, aoa.length);
    let best = null;
    for (let i = 0; i < scan; i++) {
      const cells = (aoa[i] || []).map(norm);
      const hit = headerHints.some(h => cells.some(c => c === h || c.includes(h)));
      if (hit) { best = { idx: i, width: cells.filter(c => c !== '').length, cells, exact: true }; break; }
    }
    if (!best) {
      for (let i = 0; i < Math.min(10, aoa.length); i++) {
        const cells = (aoa[i] || []).map(norm);
        const w = cells.filter(c => c !== '').length;
        if (!best || w > best.width) best = { idx: i, width: w, cells, exact: false };
      }
      if (best && best.width < 20) best = null;    // กว้างไม่พอ ไม่น่าใช่หัวตาราง
    }
    if (!best) return { kind: 'UNKNOWN', headerIdx: -1, width: 0, exact: false };

    const has = t => best.cells.some(c => c === t || c.startsWith(t));
    const hasColor = best.cells.includes('สี');
    const hasNbv   = has('มูลค่าปัจจุบัน');
    const w        = best.width;

    // จำนวนคอลัมน์เป็นสัญญาณหลัก ชื่อคอลัมน์เป็นสัญญาณรอง
    let kind;
    if (w >= 44 || (hasColor && hasNbv))            kind = 'MERGED';   // 44 คอลัมน์ พร้อมใช้
    else if (w >= 40 || (hasNbv && !hasColor))      kind = 'RP023';    // 43 คอลัมน์ ไม่มี "สี"
    else if (w >= 34 && w <= 39)                    kind = 'RP015';    // 36 คอลัมน์
    else if (hasColor && !hasNbv)                   kind = 'RP015';
    else                                            kind = 'UNKNOWN';

    return { kind, headerIdx: best.idx, width: w, exact: best.exact };
  },

  /** 2.1 (RP023, 43 คอล.) → 44 คอล. : แทรกช่องว่าง "สี" ที่ index 12 */
  transformRP023(dataRows) {
    return dataRows.map(row => {
      const newRow = new Array(44).fill('');
      for (let i = 0; i < 12; i++) newRow[i] = row[i] ?? '';
      newRow[12] = '';                                   // สี — RP023 ไม่มี
      const right = row.slice(12, 43);                   // 31 ค่า → index 13..43
      for (let i = 0; i < right.length; i++) newRow[13 + i] = right[i] ?? '';
      newRow[43] = row[42] ?? '';                        // (คงบรรทัดเดิมของ GAS ไว้)
      return newRow;
    });
  },

  /** 2.2 (RP015, 36 คอล.) → 44 คอล. */
  transformRP015(dataRows) {
    return dataRows.map(row => {
      const newRow = new Array(44).fill('');
      for (let i = 0; i < 33; i++) newRow[i] = row[i] ?? '';
      newRow[35] = row[33] ?? '';   // อายุใช้งาน (ปี)
      newRow[36] = row[34] ?? '';   // อายุใช้งาน (เดือน)
      newRow[37] = row[35] ?? '';   // มูลค่าต้นทุนที่ได้มา
      return newRow;
    });
  },

  /**
   * @param {Array<{name:string, aoa:Array<Array>}>} sources
   * @returns {{rows:Array<Array>, report:object}}
   */
  build(sources) {
    const out    = [];
    const report = { files: [], droppedTotalRows: 0, computedDep: 0 };

    for (const src of sources) {
      const det = this.detect(src.aoa);
      const dataRows = det.headerIdx >= 0 ? src.aoa.slice(det.headerIdx + 1) : src.aoa;

      let mapped;
      switch (det.kind) {
        case 'RP023': mapped = this.transformRP023(dataRows); break;
        case 'RP015': mapped = this.transformRP015(dataRows); break;
        case 'MERGED':
          mapped = dataRows.map(r => {
            const n = new Array(44).fill('');
            for (let i = 0; i < 44; i++) n[i] = r[i] ?? '';
            return n;
          });
          break;
        default:
          Logger_.warn(`   ⚠️ "${src.name}" — ไม่รู้จักรูปแบบ ข้ามไฟล์นี้`);
          report.files.push({ name: src.name, kind: 'UNKNOWN', rows: 0 });
          continue;
      }

      // ตัดแถวว่างสนิท
      mapped = mapped.filter(r => r.some(c => c !== '' && c !== null && c !== undefined));

      // แถวยอดรวมของรายงาน = ไม่มีรหัสสินทรัพย์ แต่มีตัวเลข
      let dropped = 0;
      if (CONFIG.options.dropErpTotalRows) {
        const before = mapped.length;
        mapped = mapped.filter(r => KeyNormalizer.normalize(r[CONFIG.erp.codeCol]) !== '');
        dropped = before - mapped.length;
        report.droppedTotalRows += dropped;
      }

      if (!det.exact) {
        Logger_.warn(`   ⚠️ "${src.name}" — ไม่พบคำใบ้หัวตาราง เดาว่าเป็นแถวที่ ${det.headerIdx + 1} ` +
                     `(${det.width} คอลัมน์) โปรดตรวจผลลัพธ์`);
      }
      report.files.push({ name: src.name, kind: det.kind, rows: mapped.length, dropped, exact: det.exact });
      Logger_.info(`   ↳ "${src.name}" → ${det.kind} · ${mapped.length} แถว` +
                   (dropped ? ` (คัดแถวยอดรวม ${dropped})` : ''));
      out.push(...mapped);
    }

    if (CONFIG.options.computeDepreciation) {
      report.computedDep = this._computeDepreciation(out);
    }

    return { rows: out, report };
  },

  /** เติมคอลัมน์ (5)(6)(7) ที่ต้นฉบับเว้นว่างไว้ (เป็นสูตรที่ไม่ถูกคำนวณ) */
  _computeDepreciation(rows) {
    const { cost, acc2, cur3, per4, sum5, acc6, nbv7 } = CONFIG.erp.dep;
    const num = v => {
      if (v === '' || v === null || v === undefined) return null;
      const n = Number(String(v).replace(/,/g, ''));
      return isNaN(n) ? null : n;
    };
    let filled = 0;
    for (const r of rows) {
      const c = num(r[cost]), a2 = num(r[acc2]), c3 = num(r[cur3]), p4 = num(r[per4]);
      if (a2 === null && c3 === null && p4 === null) continue;   // RP015 ไม่มีค่าเสื่อม
      const s5 = (c3 ?? 0) + (p4 ?? 0);
      const a6 = (a2 ?? 0) + s5;
      r[sum5] = s5;
      r[acc6] = a6;
      r[nbv7] = (c ?? 0) - a6;
      filled++;
    }
    return filled;
  }
};


/* ============================================================
 *  PART E — WORKBOOK I/O  (SheetJS)
 * ============================================================ */

const WorkbookIO = {

  /**
   * อ่านไฟล์เป็น AoA
   * @param {File} file
   * @param {string} [preferSheet]  ชื่อชีตที่ต้องการ (ถ้าไม่มีใช้ชีตแรก)
   * @returns {Promise<{aoa:Array<Array>, sheetName:string, sheetNames:string[]}>}
   */
  async readAoA(file, preferSheet) {
    const buf = await file.arrayBuffer();
    /* dense:true — SheetJS เก็บชีตเป็น array-of-arrays แทน object ที่ key
       ด้วย A1/B1/… ทำให้ sheet_to_json ไม่ต้องถอดรหัสที่อยู่เซลล์ทีละช่อง
       วัดกับ asset.xlsm (5,581 แถว): 1,732ms → 514ms และ AoA ตรงกันทุกค่า */
    const wb  = XLSX.read(buf, {
      type: 'array', cellDates: true, cellNF: false, cellText: false, dense: true
    });

    const names = wb.SheetNames;
    const pick  = (preferSheet && names.includes(preferSheet)) ? preferSheet : names[0];
    const ws    = wb.Sheets[pick];

    const aoa = XLSX.utils.sheet_to_json(ws, {
      header: 1, defval: '', blankrows: false, raw: true
    });

    return { aoa, sheetName: pick, sheetNames: names };
  },

  /**
   * สร้าง worksheet จาก header + rows พร้อม column width / autofilter
   */
  makeSheet(header, rows, opts = {}) {
    const strip    = CONFIG.options.stripLeadingQuote;
    const dateCols = (CONFIG.options.nativeExcelDates && opts.dateCols) ? new Set(opts.dateCols) : null;
    let converted  = 0;

    const body = rows.map(r => r.map((v, i) => {
      const x = strip ? StringHelper.stripQuote(v) : v;
      if (dateCols && dateCols.has(i)) {
        const serial = FormulaEngine.toExcelSerial(x);
        if (serial !== null) { converted++; return serial; }   // → วันที่จริงของ Excel
      }
      return x;                                     // พ.ศ. / รูปแบบอื่น → คงเป็นข้อความ
    }));

    const ws = XLSX.utils.aoa_to_sheet([header, ...body], { cellDates: false });

    // ใส่ number format ให้เซลล์ที่แปลงเป็น serial สำเร็จ → Excel แสดงเป็นวันที่
    if (dateCols) {
      for (const c of dateCols) {
        for (let r = 0; r < body.length; r++) {
          const cell = ws[XLSX.utils.encode_cell({ r: r + 1, c })];
          if (cell && cell.t === 'n') cell.z = CONFIG.dateNumberFormat;
        }
      }
      if (converted) Logger_.info(`   ↳ แปลงเป็นวันที่จริงของ Excel ${converted.toLocaleString()} เซลล์`);
    }

    // ความกว้างคอลัมน์
    ws['!cols'] = header.map((h, i) => ({
      wch: opts.widths?.[i] ?? Math.min(Math.max(String(h).length + 4, 12), 42)
    }));

    // autofilter (freeze panes ต้องใช้ SheetJS Pro — ใส่ไม่ได้ในรุ่น community)
    if (rows.length) {
      ws['!autofilter'] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: rows.length, c: header.length - 1 }
        })
      };
    }

    // hyperlink คอลัมน์รูปภาพ (แทน RichTextValue ของ Google Sheets)
    if (opts.linkCol !== undefined) {
      for (let r = 0; r < body.length; r++) {
        const addr = XLSX.utils.encode_cell({ r: r + 1, c: opts.linkCol });
        const cell = ws[addr];
        if (!cell || typeof cell.v !== 'string') continue;
        const first = cell.v.split(/[\s,]+/).find(u => /^https?:\/\//i.test(u));
        if (first) cell.l = { Target: first, Tooltip: 'เปิดรูปภาพครุภัณฑ์' };
      }
    }

    return ws;
  },

  /** ประกอบ workbook แล้วสั่งดาวน์โหลด */
  download(sheets, fileName) {
    const wb = XLSX.utils.book_new();
    for (const { name, ws } of sheets) {
      XLSX.utils.book_append_sheet(wb, ws, name.substring(0, 31));
    }
    XLSX.writeFile(wb, fileName, { bookType: 'xlsx', compression: true });
  }
};

/* ── StagedFile ──────────────────────────────────────────────
 *  เก็บ "เนื้อไฟล์" ไว้ตั้งแต่ตอนผู้ใช้เลือก ไม่ถือ File object ค้างไว้
 *
 *  ทำไม: File ที่ได้จาก <input> เป็นเพียง "การอ้างอิง" ไปยังไฟล์บนดิสก์
 *  ถ้าไฟล์ถูกแก้ไข ย้าย ลบ หรือถูก Excel ล็อกหลังจากเลือกไว้
 *  การเรียก .arrayBuffer() ทีหลังจะโยน NotReadableError
 *  (Excel ตอนกด Save จะเขียนไฟล์ใหม่แล้วลบไฟล์เดิม → handle เดิมตายทันที)
 *
 *  เดิมอ่านไฟล์ตอนกด "รัน" ซึ่งอาจห่างจากตอนเลือกเป็นนาที — ช่วงนั้นคือ
 *  ช่องโหว่ทั้งหมด ย้ายมาอ่านทันทีที่เลือกจึงปิดปัญหานี้ได้
 * ============================================================ */
class StagedFile {
  constructor(name, size, buf) { this.name = name; this.size = size; this._buf = buf; }
  async arrayBuffer() { return this._buf; }
}

const FileIO = {
  /** อ่าน File → ArrayBuffer พร้อมแปลง error เป็นข้อความที่ทำตามได้ */
  async stage(file) {
    try {
      const buf = await file.arrayBuffer();
      if (!buf || buf.byteLength === 0) {
        throw new Error(`ไฟล์ "${file.name}" ว่างเปล่า (0 ไบต์)`);
      }
      return new StagedFile(file.name, file.size, buf);
    } catch (err) {
      throw new Error(this.describe(err, file.name));
    }
  },

  /** @private แปลข้อความ error ของเบราว์เซอร์เป็นภาษาที่แก้ปัญหาได้ */
  describe(err, name) {
    const raw = err?.message ?? String(err);
    if (err?.name === 'NotReadableError' || /could not be read|permission problems/i.test(raw)) {
      return `อ่านไฟล์ "${name}" ไม่ได้ — ไฟล์อาจถูกแก้ไข ย้าย หรือยังเปิดค้างอยู่ใน Excel ` +
             `หลังจากที่เลือกไว้ · วิธีแก้: ปิดไฟล์ใน Excel ให้เรียบร้อย แล้วเลือกไฟล์ใหม่อีกครั้ง`;
    }
    if (err?.name === 'NotFoundError') {
      return `ไม่พบไฟล์ "${name}" แล้ว — อาจถูกย้ายหรือลบไป กรุณาเลือกไฟล์ใหม่`;
    }
    if (err?.name === 'SecurityError') {
      return `เบราว์เซอร์ไม่อนุญาตให้อ่านไฟล์ "${name}" · ถ้าเปิดหน้านี้จาก file:// ` +
             `ให้เปิดผ่านเว็บเซิร์ฟเวอร์แทน`;
    }
    return raw;
  }
};

/* ── Async helper — แทน BatchRunner / trigger ─────────────── */
const Async = {
  /** วน array เป็น chunk แล้วคืนคิวให้ UI หายใจระหว่างกลาง */
  async eachChunk(arr, size, fn, onProgress) {
    for (let i = 0; i < arr.length; i += size) {
      const end = Math.min(i + size, arr.length);
      for (let j = i; j < end; j++) fn(arr[j], j);
      onProgress?.(end);
      await new Promise(r => setTimeout(r, 0));
    }
  }
};


/* ============================================================
 *  PART F — PIPELINE ORCHESTRATOR
 * ============================================================ */

const Pipeline = {

  /**
   * @param {{data:File, image:File, erp:File[]}} files
   * @returns {Promise<object>} result bundle
   */
  async runFull(files) {
    Stats.reset();
    const t0 = performance.now();

    /* ── STEP 0: อ่านไฟล์ ─────────────────────────────── */
    Logger_.info('═══ STEP 0: แกะไฟล์ Excel ═══');
    const total = 1 + (files.image ? 1 : 0) + files.erp.length;
    let done = 0;
    const tick = async label => {
      UI.progress('แกะไฟล์ Excel', done, total, label);
      await new Promise(r => setTimeout(r, 0));        // ให้ progress ได้วาดก่อน parse
    };

    await tick(`${files.data.name} …`);
    const dataFile  = await WorkbookIO.readAoA(files.data,  CONFIG.sheets.data);
    done++;
    Logger_.info(`   ↳ data: ชีต "${dataFile.sheetName}" · ${dataFile.aoa.length} แถว`);

    let imageFile = { aoa: [], sheetName: '(ไม่ได้แนบ)' };
    if (files.image) {
      await tick(`${files.image.name} …`);
      imageFile = await WorkbookIO.readAoA(files.image, CONFIG.sheets.image);
      done++;
      Logger_.info(`   ↳ รูปภาพ: ชีต "${imageFile.sheetName}" · ${imageFile.aoa.length} แถว`);
    } else {
      Logger_.warn('   ⚠️ ไม่ได้แนบไฟล์รูปภาพ — คอลัมน์ Z (รูปภาพ) จะว่างทั้งหมด');
    }

    const erpSources = [];
    for (const f of files.erp) {
      await tick(`${f.name} …`);
      const r = await WorkbookIO.readAoA(f, CONFIG.sheets.assetLookup);
      erpSources.push({ name: f.name, aoa: r.aoa });
      done++;
    }
    UI.progress('แกะไฟล์ Excel', total, total, 'อ่านครบทุกไฟล์');

    /* ── STEP 1: ERP (copyAssetData) ──────────────────── */
    Logger_.info('═══ STEP 1: สร้างทะเบียนสินทรัพย์ (ERP) ═══');
    const { rows: erpRows, report: erpReport } = ErpBuilder.build(erpSources);
    const erpIndex = new ErpIndex(erpRows);
    Logger_.ok(`   ✓ ERP รวม ${erpRows.length} แถว · ` +
               `index รหัสสินทรัพย์ ${erpIndex.size} · Ref No.1 ${erpIndex.refSize}`);
    if (erpReport.droppedTotalRows) {
      Logger_.warn(`   ⚠️ คัดแถวยอดรวมออก ${erpReport.droppedTotalRows} แถว`);
    }
    if (erpReport.computedDep) {
      Logger_.info(`   ↳ คำนวณค่าเสื่อม (5)(6)(7) ให้ ${erpReport.computedDep} แถว`);
    }

    /* ── STEP 2: image cache ──────────────────────────── */
    Logger_.info('═══ STEP 2: สร้างดัชนีรูปภาพ ═══');
    // ชีตรูปภาพมี header 1 แถว → ตัดออกก่อน (เดิม getRange(2,…))
    const imageBody = imageFile.aoa.length > 1 ? imageFile.aoa.slice(1) : [];
    const imgCache  = ImageMapBuilder.build(imageBody);

    /* ── STEP 3: data → information ───────────────────── */
    Logger_.info('═══ STEP 3: data → information ═══');
    const step1 = await DataProcessor.run(dataFile.aoa, (done, total) =>
      UI.progress('Step 1 · data → information', done, total, `${done.toLocaleString()} / ${total.toLocaleString()} แถว`)
    );
    Logger_.ok(`   ✓ ${step1.rows.length.toLocaleString()} แถว (${step1.execSec}s)` +
               (step1.errorRows.length ? ` · error ${step1.errorRows.length} แถว` : ''));

    /* ── STEP 4: information → target ─────────────────── */
    Logger_.info('═══ STEP 4: information → ทะเบียน CICM ═══');
    const target = [];
    let imageHits = 0;

    await Async.eachChunk(step1.rows, CONFIG.ui.chunkSize, (srcRow, i) => {
      const r = Mapper.mapRow(srcRow, CONFIG.area.startRow + i, imgCache, erpIndex);
      if (r[25]) imageHits++;
      target.push(r);
    }, done => UI.progress('Step 2 · information → target', done, step1.rows.length,
                           `${done.toLocaleString()} / ${step1.rows.length.toLocaleString()} แถว`));

    Logger_.ok(`   ✓ ${target.length.toLocaleString()} แถว · พบรูป ${imageHits.toLocaleString()} แถว`);

    /* ── STEP 5: diagnostic + validation ──────────────── */
    Logger_.info('═══ STEP 5: Diagnostic + Validation ═══');
    const diag = Diagnostic.build(step1.rows, imgCache, imageBody);
    const vlog = ValidationLog.build(step1.rows);
    Logger_.info(`   ↳ แมตช์รูป ${diag.matched.toLocaleString()} / ไม่แมตช์ ${diag.unmatched.toLocaleString()}`);
    Logger_.info(`   ↳ validation: error ${vlog.errors} · warning ${vlog.warnings}`);

    const execSec = ((performance.now() - t0) / 1000).toFixed(1);
    Logger_.ok(`✅ Pipeline เสร็จสมบูรณ์ใน ${execSec} วินาที`);

    return {
      target, information: step1.rows, erpRows, erpReport,
      diag, vlog, imageHits, imgStats: imgCache.getStats(),
      errorRows: step1.errorRows, execSec, stats: Stats.all()
    };
  },

  /** สร้าง workbook ผลลัพธ์แล้วดาวน์โหลด */
  export(result) {
    const wsTarget = WorkbookIO.makeSheet(HEADERS, result.target, {
      linkCol:  CONFIG.image.targetCol - 1,         // Z (0-based = 25)
      dateCols: CONFIG.dateColumnsCE,               // เฉพาะคอลัมน์ ค.ศ. เท่านั้น
      widths:  HEADERS.map((h, i) =>
                 i === 25 ? 46 :
                 i === CONFIG.erp.erpNoCol ? 28 :
                 CONFIG.dateColumns.includes(i + 1) ? 15 :
                 Math.min(Math.max(h.length + 4, 12), 40))
    });

    const INFO_HEADERS = [
      'ลำดับ','วันที่ตรวจรับ','ประเภท','ชื่อครุภัณฑ์','ยี่ห้อ','รุ่น','สี','ครุภัณฑ์บริจาค',
      'รายละเอียด','หมายเลขครุภัณฑ์','หมายเลขครุภัณฑ์ ERP','หมายเลขครุภัณฑ์เดิม','Serial No.',
      'group_code','class_code','type_code','ประเภท (แปล)','หน่วยนับ','จำนวน','ราคา','ผู้ขาย',
      'เริ่มรับประกัน','สิ้นสุดรับประกัน','ตึก/อาคาร','ชั้น','ห้อง','หน่วยงาน','งบประมาณ',
      'สถานะ','ปีงบ','หมายเหตุ','หมายเหตุส่วนงาน','สถานะระบบงาน','โครงการ'
    ];

    const sheets = [
      { name: CONFIG.sheets.target,      ws: wsTarget },
      { name: CONFIG.sheets.source,      ws: WorkbookIO.makeSheet(INFO_HEADERS, result.information) },
      { name: CONFIG.sheets.assetLookup, ws: WorkbookIO.makeSheet(ERP_HEADERS,  result.erpRows) },
      { name: CONFIG.sheets.diagnostic,  ws: WorkbookIO.makeSheet(Diagnostic.HEADERS,   result.diag.rows) },
      { name: CONFIG.sheets.validation,  ws: WorkbookIO.makeSheet(ValidationLog.HEADERS, result.vlog.rows) }
    ];

    const d  = new Date();
    const ts = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}` +
               `_${String(d.getHours()).padStart(2,'0')}${String(d.getMinutes()).padStart(2,'0')}`;

    WorkbookIO.download(sheets, `CICM_ทะเบียนครุภัณฑ์_${ts}.xlsx`);
    Logger_.ok(`⬇️ ดาวน์โหลด CICM_ทะเบียนครุภัณฑ์_${ts}.xlsx (${sheets.length} ชีต)`);
  }
};

/* ── Diagnostic — เดิมเขียนลงชีต _diagnose_images ────────── */
const Diagnostic = {
  HEADERS: ['Source Row','ลำดับ (raw)','ลำดับ (normalized)','Matched?','Match Level',
            'URLs Found','Validation','URLs (preview)'],

  build(infoRows, cache, imageBody) {
    const rows = [];
    let matched = 0, unmatched = 0;
    const levels = {};

    infoRows.forEach((row, i) => {
      const raw = row[CONFIG.source.assetNoCol];
      const key = KeyNormalizer.normalize(raw);
      if (!key) return;

      const { urls, level } = cache.lookup(raw);
      const found = urls.length > 0;
      found ? matched++ : unmatched++;
      levels[level] = (levels[level] || 0) + 1;

      const v = Validators.validateAssetRow(row);
      const valStr = v.errors.length   ? `❌ ${v.errors.join(', ')}`
                   : v.warnings.length ? `⚠️ ${v.warnings.join(', ')}`
                   : '✓';

      rows.push([
        i + 2, String(raw), key, found ? '✅ YES' : '❌ NO',
        level, urls.length, valStr, urls.slice(0, 3).join(' | ')
      ]);
    });

    // ต่อท้ายด้วยตัวอย่างจากชีต "รูปภาพ" (เหมือน _writeReport เดิม)
    rows.push([], ['📋 ตัวอย่างจากชีต "รูปภาพ" (10 แถวแรก)']);
    rows.push(['Row','Code (raw)','Code (normalized)','URL (raw)','URL (normalized)']);
    imageBody.slice(0, 10).forEach((r, i) => rows.push([
      i + 2,
      String(r[CONFIG.image.codeCol] ?? ''),
      KeyNormalizer.normalize(r[CONFIG.image.codeCol]),
      String(r[CONFIG.image.urlCol] ?? ''),
      UrlNormalizer.normalize(r[CONFIG.image.urlCol])
    ]));

    return { rows, matched, unmatched, levels };
  }
};

/* ── ValidationLog — เดิมคือ validateSourceData() ─────────── */
const ValidationLog = {
  HEADERS: ['Row','ลำดับ','หมายเลขครุภัณฑ์','สถานะ','Errors','Warnings'],

  build(infoRows) {
    const rows = [];
    let errors = 0, warnings = 0;

    infoRows.forEach((row, i) => {
      const v = Validators.validateAssetRow(row);
      if (v.valid && v.warnings.length === 0) return;   // เก็บเฉพาะที่มีปัญหา
      if (!v.valid) errors++;
      if (v.warnings.length) warnings++;

      rows.push([
        i + 2,
        String(row[0] ?? ''),
        StringHelper.stripQuote(String(row[9] ?? '')),
        !v.valid ? '❌ ERROR' : '⚠️ WARNING',
        v.errors.join(' · '),
        v.warnings.join(' · ')
      ]);
    });

    if (rows.length === 0) rows.push(['—','—','—','✅ ไม่พบปัญหา','','']);
    return { rows, errors, warnings };
  }
};


/* ============================================================
 *  PART G — UI CONTROLLER
 * ============================================================ */

const UI = {
  el: {},
  files: { data: null, image: null, erp: [] },
  result: null,

  init() {
    const $ = id => document.getElementById(id);
    this.el = {
      log: $('log'),
      progressCard: $('progressCard'), progressTitle: $('progressTitle'),
      barFill: $('barFill'), barText: $('barText'), barPct: $('barPct'),
      resultCard: $('resultCard'), stats: $('stats'),
      btnRun: $('btn-run'), btnDiag: $('btn-diag'),
      btnReset: $('btn-reset'), btnDownload: $('btn-download')
    };

    Logger_.bind((kind, msg) => this._log(kind, msg));

    this._wireZone('data',  false);
    this._wireZone('image', false);
    this._wireZone('erp',   true);

    ['skip-header','drop-total','calc-dep','strip-quote','native-dates'].forEach(k => {
      const map = { 'skip-header':'skipDataHeaderRow', 'drop-total':'dropErpTotalRows',
                    'calc-dep':'computeDepreciation', 'strip-quote':'stripLeadingQuote',
                    'native-dates':'nativeExcelDates' };
      const box = $(`opt-${k}`);
      if (!box) return;
      box.addEventListener('change', () => {
        // CONFIG ถูก freeze ที่ระดับบนสุด แต่ CONFIG.options ยังเขียนได้
        CONFIG.options[map[k]] = box.checked;
      });
    });

    this.el.btnRun.addEventListener('click', () => this.run(false));
    this.el.btnDiag.addEventListener('click', () => this.run(true));
    this.el.btnReset.addEventListener('click', () => location.reload());
    this.el.btnDownload.addEventListener('click', () => {
      if (this.result) Pipeline.export(this.result);
    });

    Logger_.info('CICM Asset Manager · Web Edition 5.0 พร้อมทำงาน');
    Logger_.info(`SheetJS ${typeof XLSX !== 'undefined' ? XLSX.version : '(ไม่พบ!)'}`);
  },

  /** @private ผูก dropzone + input */
  _wireZone(key, multiple) {
    const zone  = document.getElementById(`zone-${key}`);
    const input = document.getElementById(`file-${key}`);
    const nameEl = document.getElementById(`name-${key}`);
    const metaEl = document.getElementById(`meta-${key}`);

    const accept = async list => {
      const arr = Array.from(list).filter(f => /\.(xlsx|xlsm|xls|csv)$/i.test(f.name));
      if (!arr.length) { Logger_.warn('ไฟล์ที่เลือกไม่ใช่ Excel/CSV'); return; }

      // อ่านเนื้อไฟล์ทันที ไม่รอจนกดรัน — ดูเหตุผลที่ StagedFile
      zone.classList.add('filled');
      nameEl.textContent = arr.map(f => `✓ ${f.name}`).join('\n');
      metaEl.textContent = 'กำลังอ่านไฟล์…';
      this.files[key] = null;
      this._refreshButtons();

      const staged = [];
      try {
        for (const f of arr) {
          staged.push(await FileIO.stage(f));
          await new Promise(r => setTimeout(r, 0));   // ให้ UI ได้วาด
        }
      } catch (err) {
        zone.classList.remove('filled');
        nameEl.textContent = '';
        metaEl.textContent = '';
        this.files[key] = multiple ? [] : null;
        this._refreshButtons();
        Logger_.error(err.message);
        this.alert(`❌ ${err.message}`);
        input.value = '';                              // ให้เลือกไฟล์เดิมซ้ำได้
        return;
      }

      this.files[key] = multiple ? staged : staged[0];
      const kb = staged.reduce((n, f) => n + f.size, 0) / 1024;
      metaEl.textContent = staged.map(f => `${(f.size / 1024).toFixed(0)} KB`).join(' · ')
                         + `  ·  อ่านแล้ว ${kb.toFixed(0)} KB`;
      this._refreshButtons();
      Logger_.ok(`อ่านไฟล์ [${key}] แล้ว: ${staged.map(f => f.name).join(', ')}`);
    };

    input.addEventListener('change', e => accept(e.target.files));

    ['dragenter','dragover'].forEach(ev =>
      zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.add('dragover'); }));
    ['dragleave','drop'].forEach(ev =>
      zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.remove('dragover'); }));
    zone.addEventListener('drop', e => {
      if (e.dataTransfer?.files?.length) accept(e.dataTransfer.files);
    });
  },

  /** @private */
  _refreshButtons() {
    // ไฟล์รูปภาพเป็น optional — ไม่มีก็รันได้ คอลัมน์ Z จะว่าง
    const ready = !!(this.files.data && this.files.erp?.length);
    this.el.btnRun.disabled  = !ready;
    this.el.btnDiag.disabled = !(this.files.data && this.files.image);   // diagnostic ต้องมีรูป
  },

  async run(diagOnly) {
    this.el.btnRun.disabled = this.el.btnDiag.disabled = true;
    this.el.progressCard.style.display = 'block';
    this.el.resultCard.style.display   = 'none';
    this.el.log.innerHTML = '';

    try {
      const files = {
        data:  this.files.data,
        image: this.files.image,
        erp:   diagOnly ? [] : (this.files.erp || [])
      };
      const result = await Pipeline.runFull(files);
      this.result = result;

      this._renderStats(result, diagOnly);
      this.el.resultCard.style.display = 'block';
      this.progress('เสร็จสิ้น', 1, 1, 'พร้อมดาวน์โหลด');

    } catch (err) {
      // แปลง error ของ File API ให้เป็นคำแนะนำที่ทำตามได้ ก่อนส่งต่อ
      err.message = FileIO.describe(err, 'ที่เลือกไว้');
      ErrorHandler.handle('Pipeline.runFull', err);
      this.el.progressCard.style.display = 'none';
    } finally {
      this._refreshButtons();
    }
  },

  /** @private */
  _renderStats(r, diagOnly) {
    const n = v => Number(v || 0).toLocaleString();
    const cards = [
      { k: 'แถวใน target',        v: n(r.target.length) },
      { k: 'พบรูปภาพ',            v: n(r.imageHits),   cls: r.imageHits ? 'ok' : 'warn' },
      { k: 'BV ถูกรูปแบบ (21/26)',   v: n(r.stats['erpNo:ถูกรูปแบบ']), cls: 'ok' },
      { k: 'BV ผิดรูปแบบ',           v: n(r.stats['erpNo:ผิดรูปแบบ']),
        cls: r.stats['erpNo:ผิดรูปแบบ'] ? 'err' : 'ok' },
      { k: 'BV ว่าง',                v: n(r.stats['erpNo:ว่าง']),
        cls: r.stats['erpNo:ว่าง'] ? 'warn' : 'ok' },
      { k: 'AC ชั้น 1 (BV→รหัส)',    v: n(r.stats['ac:BV→C']), cls: 'ok' },
      { k: 'AC ชั้น 2 (G→Ref No.1)', v: n(r.stats['ac:G→P']) },
      { k: 'AC ชั้น 3–5 (G/F/E)',
        v: n((r.stats['ac:G→C'] || 0) + (r.stats['ac:F=รหัส'] || 0) + (r.stats['ac:E→C'] || 0)) },
      { k: 'AC ชั้น 6 (กวาดทั้งแถว)',  v: n(r.stats['ac:SCAN']),
        cls: r.stats['ac:SCAN'] ? 'warn' : 'ok' },
      { k: 'AC หาไม่เจอ (ว่าง)',      v: n(r.stats['ac:NONE']), cls: r.stats['ac:NONE'] ? 'warn' : 'ok' },
      { k: 'AC มีค่าแต่ไม่มีแถวใน ERP', v: n(r.stats.acNoErpRow),
        cls: r.stats.acNoErpRow ? 'warn' : 'ok' },
      { k: 'Validation error',    v: n(r.vlog.errors), cls: r.vlog.errors ? 'err' : 'ok' },
      { k: 'Validation warning',  v: n(r.vlog.warnings), cls: r.vlog.warnings ? 'warn' : 'ok' },
      { k: 'ERP ทั้งหมด',         v: n(r.erpRows.length) },
      { k: 'เวลาที่ใช้',          v: `${r.execSec}s` }
    ];
    this.el.stats.innerHTML = cards.map(c =>
      `<div class="stat ${c.cls || ''}"><div class="k">${c.k}</div><div class="v">${c.v}</div></div>`
    ).join('');
    if (diagOnly) Logger_.warn('โหมด Diagnostic: ไม่ได้โหลด ERP — คอลัมน์ AC…BR จะว่าง');
  },

  _lastPct: -1,
  _lastTitle: '',
  progress(title, done, total, text) {
    const pct = total > 0 ? Math.round((done / total) * 100) : 0;
    // เขียน DOM เฉพาะตอนตัวเลขเปลี่ยนจริง — การ์ดใช้ backdrop-filter
    // ทุกครั้งที่ layout เปลี่ยนเครื่องที่ไม่มี GPU แยกต้อง re-composite ใหม่
    if (pct === this._lastPct && title === this._lastTitle) return;
    this._lastPct = pct; this._lastTitle = title;

    this.el.progressTitle.textContent = title;
    this.el.barFill.style.width = `${pct}%`;
    this.el.barPct.textContent  = `${pct}%`;
    this.el.barText.textContent = text || '';
  },

  alert(msg) { window.alert(msg); },

  /** @private */
  _log(kind, msg) {
    const div = document.createElement('div');
    div.className = kind;
    div.textContent = msg;
    this.el.log.appendChild(div);
    this.el.log.scrollTop = this.el.log.scrollHeight;
  }
};

document.addEventListener('DOMContentLoaded', () => UI.init());
