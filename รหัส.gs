/**
 * ========================================================================
 *  CMMBA 34 — Aloha Party Registration Web App
 *  Backend: Google Apps Script (with OCR slip + Top Spenders)
 * ========================================================================
 *  ⚠️ ก่อนใช้: ต้อง enable Drive API ใน Apps Script
 *  → Services (+ icon ในแถบซ้าย) → Drive API → Add
 * ========================================================================
 */

const CONFIG = {
  SHEET_ID:   '1grJlq5HKqH8C6lo806igLHy5hpod4U4PXzGHQMkswUg',
  FOLDER_ID:  '1_Y5b1hdfNahaetMZCvdsbE-bcxSmDUcu',  // assets: bg, logos, QR
  SLIP_FOLDER_ID: '1MoY2U_UtU_4HpZz7HRVirytJtHgv4_IK',  // สลิปการโอนเงิน
  DEADLINE:   '2026-07-01T00:00:00+07:00',  // ปิดรับลงทะเบียนเที่ยงคืน — เปิดถึงสิ้นวัน 30 มิ.ย. 2569
  DONATION_DEADLINE: '2026-08-03T00:00:00+07:00',  // ปิดรับบริจาคหลังงานจบ — เที่ยงคืนวันที่ 2 ส.ค. 2569
  EVENT_NAME: 'CMMBA 34 Aloha Party',
  TZ:         'Asia/Bangkok',
  QR_CACHE_TTL: 21600,
  STAFF_PASSWORD: '3334',   // password เจ้าหน้าที่บันทึกบริจาคเงินสด
};

const GROUP_TABS = {
  student34: 'น้องรุ่น 34',
  senior33:  'พี่รุ่น 33',
  alumni:    'ศิษย์เก่า',
  faculty:   'อาจารย์/บุคลากร',
};
const ALL_REG_TABS = Object.values(GROUP_TABS);
const DON_TAB = 'บริจาค';

// reverse map: tab name → group
const TAB_TO_GROUP = {};
Object.keys(GROUP_TABS).forEach(g => { TAB_TO_GROUP[GROUP_TABS[g]] = g; });

function getSheetByGroup(ss, group) {
  const name = GROUP_TABS[group];
  if (!name) throw new Error('กลุ่มไม่ถูกต้อง: ' + group);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initRegSheet(sheet, group);
  }
  return sheet;
}

function getGroupOfSheet(sheet) {
  return TAB_TO_GROUP[sheet.getName()];
}

function getHeadersOfSheet(sheet) {
  const group = getGroupOfSheet(sheet);
  return HEADERS_BY_GROUP[group] || HEADERS;
}

function initRegSheet(sheet, group) {
  const headers = HEADERS_BY_GROUP[group] || HEADERS;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0EA5E9')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    // ไม่ตรึงคอลัมน์ตามคำขอ
  }
}

// ===== HEADERS per group (ลบคอลัมน์ที่ไม่เกี่ยวข้องของแต่ละกลุ่ม) =====
const HEADERS_COMMON_TOP = [
  'Registration ID', 'วันที่ลงทะเบียน', 'กลุ่ม',
  'คำนำหน้า', 'ชื่อ', 'นามสกุล', 'ชื่อ-นามสกุล', 'ชื่อเล่น',
];

const HEADERS_COMMON_HEALTH = [
  'ข้อจำกัดอาหาร', 'แพ้อาหาร', 'โรคประจำตัว', 'ยาประจำตัว',
];

const HEADERS_COMMON_BOTTOM = [
  'ยินยอมใช้ภาพ', 'หมายเหตุ',
  'ค่าใช้จ่ายที่คำนวณ (บาท)',
  'สลิปโอน', 'QR Ref', 'ยอดโอนจริง OCR (บาท)', 'สถานะ OCR',
  'สถานะ',
  'จำนวนครั้งที่แก้ไข', 'เวลาแก้ไขล่าสุด',
];

const HEADERS_BY_GROUP = {
  // น้อง 34: ไม่มี roomType/roommate (ทีมจัดให้), ไม่มี companion, มี emergency contact, มี nationalId+PDPA
  student34: [
    ...HEADERS_COMMON_TOP,
    'รหัสนักศึกษา', 'เลขบัตรประชาชน',
    'เพศ', 'วันเกิด', 'เบอร์โทร', 'อีเมล',
    ...HEADERS_COMMON_HEALTH,
    'ผู้ติดต่อฉุกเฉิน - ชื่อ', 'ผู้ติดต่อฉุกเฉิน - ความสัมพันธ์', 'ผู้ติดต่อฉุกเฉิน - เบอร์',
    'ยินยอม PDPA (เลขบัตร)',
    ...HEADERS_COMMON_BOTTOM,
  ],
  // พี่ 33: ไม่มี roomType/roommate (ทีมจัดให้), ไม่มี companion, มี nationalId+PDPA
  senior33: [
    ...HEADERS_COMMON_TOP,
    'รหัสนักศึกษา', 'เลขบัตรประชาชน',
    'เพศ', 'เบอร์โทร', 'อีเมล',
    ...HEADERS_COMMON_HEALTH,
    'ยินยอม PDPA (เลขบัตร)',
    ...HEADERS_COMMON_BOTTOM,
  ],
  // ศิษย์เก่า: มี roomType/roommate/companion ครบ, ไม่มี nationalId
  alumni: [
    ...HEADERS_COMMON_TOP,
    'CMMBA รุ่นที่',
    'เพศ', 'เบอร์โทร', 'อีเมล',
    ...HEADERS_COMMON_HEALTH,
    'ประเภทห้องพัก', 'ประเภทเตียง',
    'รูมเมท - ชื่อ', 'รูมเมท - นามสกุล', 'รูมเมท - ชื่อ-นามสกุล', 'รูมเมท - ชื่อเล่น', 'รูมเมท - รหัส/รุ่น',
    'รูมเมท - เพศ', 'รูมเมท - เบอร์', 'รูมเมท - อีเมล',
    'รูมเมท - ข้อจำกัดอาหาร', 'รูมเมท - แพ้อาหาร', 'รูมเมท - โรคประจำตัว', 'รูมเมท - ยา',
    'มีผู้ติดตาม', 'ผู้ติดตาม - ผู้ใหญ่', 'ผู้ติดตาม - เด็ก ≤12', 'ผู้ติดตาม - รายละเอียด',
    'ผู้ติดตาม - ข้อจำกัดอาหาร', 'ผู้ติดตาม - แพ้อาหาร', 'ผู้ติดตาม - โรค/ยา',
    ...HEADERS_COMMON_BOTTOM,
  ],
  // อาจารย์: มี roomType/roommate/companion, ไม่มี nationalId
  faculty: [
    ...HEADERS_COMMON_TOP,
    'ตำแหน่ง/สังกัด',
    'เบอร์โทร', 'อีเมล',
    ...HEADERS_COMMON_HEALTH,
    'ประเภทห้องพัก', 'ประเภทเตียง',
    'รูมเมท - ชื่อ', 'รูมเมท - นามสกุล', 'รูมเมท - ชื่อ-นามสกุล', 'รูมเมท - ชื่อเล่น', 'รูมเมท - รหัส/รุ่น',
    'รูมเมท - เพศ', 'รูมเมท - เบอร์', 'รูมเมท - อีเมล',
    'รูมเมท - ข้อจำกัดอาหาร', 'รูมเมท - แพ้อาหาร', 'รูมเมท - โรคประจำตัว', 'รูมเมท - ยา',
    'มีผู้ติดตาม', 'ผู้ติดตาม - ผู้ใหญ่', 'ผู้ติดตาม - เด็ก ≤12', 'ผู้ติดตาม - รายละเอียด',
    'ผู้ติดตาม - ข้อจำกัดอาหาร', 'ผู้ติดตาม - แพ้อาหาร', 'ผู้ติดตาม - โรค/ยา',
    ...HEADERS_COMMON_BOTTOM,
  ],
};

// HEADERS เดิม = union เพื่อ backward compat ใน code ที่ยังใช้ indexOf
const HEADERS = [
  ...HEADERS_COMMON_TOP,
  'รหัสนักศึกษา', 'เลขบัตรประชาชน', 'CMMBA รุ่นที่', 'ตำแหน่ง/สังกัด',
  'เพศ', 'วันเกิด',
  'เบอร์โทร', 'อีเมล',
  ...HEADERS_COMMON_HEALTH,
  'ประเภทห้องพัก', 'ประเภทเตียง',
  'รูมเมท - ชื่อ', 'รูมเมท - นามสกุล', 'รูมเมท - ชื่อ-นามสกุล', 'รูมเมท - ชื่อเล่น', 'รูมเมท - รหัส/รุ่น',
  'รูมเมท - เพศ', 'รูมเมท - เบอร์', 'รูมเมท - อีเมล',
  'รูมเมท - ข้อจำกัดอาหาร', 'รูมเมท - แพ้อาหาร', 'รูมเมท - โรคประจำตัว', 'รูมเมท - ยา',
  'มีผู้ติดตาม', 'ผู้ติดตาม - ผู้ใหญ่', 'ผู้ติดตาม - เด็ก ≤12', 'ผู้ติดตาม - รายละเอียด',
  'ผู้ติดตาม - ข้อจำกัดอาหาร', 'ผู้ติดตาม - แพ้อาหาร', 'ผู้ติดตาม - โรค/ยา',
  'ผู้ติดต่อฉุกเฉิน - ชื่อ', 'ผู้ติดต่อฉุกเฉิน - ความสัมพันธ์', 'ผู้ติดต่อฉุกเฉิน - เบอร์',
  'ยินยอมใช้ภาพ', 'ยินยอม PDPA (เลขบัตร)', 'หมายเหตุ',
  'ค่าใช้จ่ายที่คำนวณ (บาท)',
  'สลิปโอน', 'QR Ref', 'ยอดโอนจริง OCR (บาท)', 'สถานะ OCR',
  'สถานะ',
  'จำนวนครั้งที่แก้ไข', 'เวลาแก้ไขล่าสุด',
];

const DON_HEADERS = [
  'Donation ID', 'วันที่บริจาค', 'Registration ID',
  'ชื่อ-นามสกุล', 'ชื่อเล่น', 'กลุ่ม', 'รุ่น',
  'เบอร์โทร', 'อีเมล',
  'บริจาคในนาม', 'ยอดบริจาค (บาท)',
  'สลิปโอน', 'QR Ref', 'ยอดโอน OCR', 'สถานะ OCR',
  'แหล่งที่มา', // "พร้อมลงทะเบียน" / "บริจาคเพิ่มเติม"
];

// ====== ENTRY POINTS ======

function doGet(e) {
  const page = (e && e.parameter && e.parameter.p) || '';
  if (page === 'display') {
    return HtmlService.createTemplateFromFile('Display')
      .evaluate()
      .setTitle('Aloha Party CMMBA 34 — Live Donations')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Aloha Party CMMBA 34 — ลงทะเบียน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * คืน URL ของหน้า Display (live projector) — สำหรับให้แอปเปิด tab ใหม่
 */
function getDisplayUrl() {
  try {
    return ScriptApp.getService().getUrl() + '?p=display';
  } catch (e) {
    return '';
  }
}

function getConfig() {
  const now = new Date();
  return {
    eventName: CONFIG.EVENT_NAME,
    deadline: CONFIG.DEADLINE,
    isOpen: now <= new Date(CONFIG.DEADLINE),
    deadlineLabel: '30 มิถุนายน 2569',
    donationDeadline: CONFIG.DONATION_DEADLINE,
    donationOpen: now <= new Date(CONFIG.DONATION_DEADLINE),
    donationDeadlineLabel: '2 สิงหาคม 2569',
  };
}

/**
 * เช็คว่ายังเปิดรับบริจาคอยู่ไหม (หลังงานจบ = ปิด)
 */
function isDonationOpen() {
  return new Date() <= new Date(CONFIG.DONATION_DEADLINE);
}

/**
 * Top 10 ผู้โอนเงินมากที่สุด (จากยอด OCR ที่อ่านได้จากสลิปจริง)
 */
function getTopSpenders() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const donSheet = ss.getSheetByName(DON_TAB);
    if (!donSheet || donSheet.getLastRow() < 2) return [];

    const data = donSheet.getRange(2, 1, donSheet.getLastRow() - 1, DON_HEADERS.length).getValues();

    const idxName    = DON_HEADERS.indexOf('ชื่อ-นามสกุล');
    const idxNick    = DON_HEADERS.indexOf('ชื่อเล่น');
    const idxGroup   = DON_HEADERS.indexOf('กลุ่ม');
    const idxBatch   = DON_HEADERS.indexOf('รุ่น');
    const idxInName  = DON_HEADERS.indexOf('บริจาคในนาม');
    const idxAmount  = DON_HEADERS.indexOf('ยอดบริจาค (บาท)');
    const idxOcr     = DON_HEADERS.indexOf('ยอดโอน OCR');
    if (idxName < 0 || idxInName < 0 || idxAmount < 0) return [];

    const map = {};
    data.forEach(function(row) {
      const inName = String(row[idxInName] || '');
      if (inName.indexOf('ส่วนตัว') !== 0) return;
      const amount = parseFloat(row[idxAmount]) || 0;
      const finalAmount = parseFloat(row[idxOcr]) || amount;
      if (finalAmount <= 0) return;
      const key = String(row[idxName] || '').trim();
      if (!key) return;
      if (!map[key]) {
        map[key] = {
          name: key,
          nickname: String(row[idxNick] || ''),
          group: String(row[idxGroup] || ''),
          batch: String(row[idxBatch] || ''),
          total: 0,
        };
      }
      map[key].total += amount;
    });

    const arr = [];
    for (const k in map) arr.push(map[k]);
    arr.sort(function(a, b) { return b.total - a.total; });
    return arr.slice(0, 10);
  } catch (e) {
    console.error('getTopSpenders error:', e);
    return [];
  }
}

function getTopBatchSpenders() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const donSheet = ss.getSheetByName(DON_TAB);
    if (!donSheet || donSheet.getLastRow() < 2) return [];

    const data = donSheet.getRange(2, 1, donSheet.getLastRow() - 1, DON_HEADERS.length).getValues();

    const idxBatch  = DON_HEADERS.indexOf('รุ่น');
    const idxInName = DON_HEADERS.indexOf('บริจาคในนาม');
    const idxAmount = DON_HEADERS.indexOf('ยอดบริจาค (บาท)');
    if (idxBatch < 0 || idxInName < 0 || idxAmount < 0) return [];

    const map = {};
    data.forEach(function(row) {
      const inName = String(row[idxInName] || '');
      if (inName.indexOf('รุ่น') !== 0) return;
      const amount = parseFloat(row[idxAmount]) || 0;
      if (amount <= 0) return;
      const batchName = String(row[idxBatch] || '').trim();
      if (!batchName) return;
      if (!map[batchName]) {
        map[batchName] = { name: batchName, batch: batchName, total: 0, contributors: 0 };
      }
      map[batchName].total += amount;
      map[batchName].contributors += 1;
    });

    const arr = [];
    for (const k in map) arr.push(map[k]);
    arr.sort(function(a, b) { return b.total - a.total; });
    return arr.slice(0, 10);
  } catch (e) {
    console.error('getTopBatchSpenders error:', e);
    return [];
  }
}

/**
 * รวมยอดบริจาคทั้งหมด — สำหรับ live display
 */
function getDonationTotal() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(DON_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { total: 0, count: 0 };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idxAmount = headers.indexOf('ยอดบริจาค (บาท)');
    const idxOcr    = headers.indexOf('ยอด OCR (บาท)');
    if (idxAmount < 0) return { total: 0, count: 0 };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    let total = 0;
    let count = 0;
    for (let i = 0; i < data.length; i++) {
      const amt = parseFloat(data[i][idxAmount]) || 0;
      const ocr = idxOcr >= 0 ? (parseFloat(data[i][idxOcr]) || 0) : 0;
      // ใช้ amount หลัก fallback OCR ถ้า amount เป็น 0
      const final = amt > 0 ? amt : ocr;
      if (final > 0) {
        total += final;
        count++;
      }
    }
    return { total: total, count: count };
  } catch (e) {
    console.error('getDonationTotal error:', e);
    return { total: 0, count: 0 };
  }
}

function getQRImage() {
  try {
    const cache = CacheService.getScriptCache();
    // Cache fileId แทน base64 (cache จำกัด 100KB)
    let fileId = cache.get('qr_file_id');

    if (!fileId) {
      const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
      const files = folder.getFiles();
      const allFiles = [];
      let qrFile = null;

      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();
        const mime = file.getMimeType();
        allFiles.push(`${name} (${mime})`);
        if (!mime.startsWith('image/')) continue;

        if (/qr[-_\s]?payment/i.test(name)) {
          qrFile = file;
          break;
        }
      }

      console.log('[getQRImage] Files in folder: ' + allFiles.join(', '));

      if (!qrFile) {
        console.log('[getQRImage] ไม่พบไฟล์ Qr-payment');
        return '';
      }

      fileId = qrFile.getId();
      console.log('[getQRImage] เลือกไฟล์: ' + qrFile.getName() + ' (id: ' + fileId + ')');
      cache.put('qr_file_id', fileId, CONFIG.QR_CACHE_TTL);
    }

    // Encode base64 ทุกครั้ง — ไม่เก็บ cache (ใหญ่เกิน 100KB)
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const dataUrl = 'data:' + file.getMimeType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    return dataUrl;
  } catch (e) {
    console.error('getQRImage error:', e);
    return '';
  }
}

/**
 * ดึงรูป background + logo จาก Drive folder
 * ส่ง URL (lh3.googleusercontent.com) แทน base64 — เพื่อหลีกเลี่ยงปัญหา cache size
 * - bg-*.{png,jpg,jpeg,webp}  → background ของหน้าแรก
 * - logo-ku.* → logo KU
 * - logo-cmmba.* → logo CMMBA
 */
function getAssets() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('assets_v6');
    if (cached) {
      const result = JSON.parse(cached);
      // เติม logoAloha base64 เพิ่ม (ไม่ cache ตัวนี้เพราะใหญ่)
      result.logoAloha = getLogoAlohaBase64();
      return result;
    }

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const files = folder.getFiles();
    const result = { bg: '', bgMobile: '', logoKu: '', logoCmmba: '', logoAloha: '', logoFlower: '' };

    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName().toLowerCase();
      const mime = file.getMimeType();
      if (!mime.startsWith('image/')) continue;

      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) { /* ignore */ }

      const fileId = file.getId();
      const url = `https://lh3.googleusercontent.com/d/${fileId}=s2000`;

      // bg-mobile ต้องตรวจก่อน bg ทั่วไป — ไม่งั้นจะ match กับ regex bg- ทันที
      if (/^bg[-_]mobile/i.test(name) && !result.bgMobile) {
        result.bgMobile = url;
      } else if (/^bg[-_]/i.test(name) && !result.bg) {
        result.bg = url;
      } else if (/^logo[-_]aloha/i.test(name)) {
        // logo aloha → ส่งเป็น base64 (โหลดเร็วกว่า) — เติมท้าย
      } else if (/^logo[-_]flower/i.test(name) && !result.logoFlower) {
        result.logoFlower = url;
      } else if (/^logo[-_]ku\b/i.test(name) && !result.logoKu) {
        result.logoKu = url;
      } else if (/^logo[-_]cmmba\b/i.test(name) && !result.logoCmmba) {
        result.logoCmmba = url;
      }
    }
    cache.put('assets_v6', JSON.stringify(result), CONFIG.QR_CACHE_TTL);

    // เติม logo aloha base64
    result.logoAloha = getLogoAlohaBase64();
    return result;
  } catch (e) {
    console.error('getAssets error:', e);
    return { bg: '', bgMobile: '', logoKu: '', logoCmmba: '', logoAloha: '', logoFlower: '' };
  }
}

/**
 * โหลด logo-aloha-* เป็น base64 (cache แยก เผื่อมีปัญหา size)
 */
function getLogoAlohaBase64() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('logo_aloha_b64');
    if (cached) return cached;

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName().toLowerCase();
      const mime = file.getMimeType();
      if (!mime.startsWith('image/')) continue;
      if (!/^logo[-_]aloha/i.test(name)) continue;

      const blob = file.getBlob();
      const bytes = blob.getBytes();
      // ถ้าใหญ่เกิน 80KB → fallback เป็น URL (cache limit ~100KB หลัง base64)
      if (bytes.length > 80 * 1024) {
        try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (e) { /* ignore */ }
        return `https://lh3.googleusercontent.com/d/${file.getId()}=s2000`;
      }

      const dataUrl = 'data:' + mime + ';base64,' + Utilities.base64Encode(bytes);
      try {
        cache.put('logo_aloha_b64', dataUrl, CONFIG.QR_CACHE_TTL);
      } catch (e) { /* cache อาจ full — ไม่เป็นไร */ }
      return dataUrl;
    }
    return '';
  } catch (e) {
    console.error('getLogoAlohaBase64 error:', e);
    return '';
  }
}

// ====== MAIN SUBMIT ======

function submitRegistration(payload) {
  try {
    if (new Date() > new Date(CONFIG.DEADLINE)) {
      return { ok: false, message: 'ปิดรับลงทะเบียนแล้ว (หมดเขต 30 มิ.ย. 2569)' };
    }

    const validation = validatePayload(payload);
    if (!validation.ok) return validation;

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureSheetReady(ss);
    const sheet = getSheetByGroup(ss, payload.group);

    const regId = generateRegistrationId(ss, sheet, payload.group);
    const now = new Date();
    const timestamp = Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');

    const cost = calculateCost(payload);
    const expectedTotal = totalAmount(payload); // cost + supportPersonal + supportBatch

    let slipUrl = '';
    let slipFileId = '';
    let ocrAmount = 0;
    let ocrStatus = 'ไม่มีสลิป';

    if (payload.slip && payload.slip.data) {
      const result = uploadSlip(payload.slip, regId, payload);
      slipUrl = result.url;
      slipFileId = result.fileId;

      const qrAmount = parseFloat(payload.slipQrAmount);
      if (!isNaN(qrAmount) && qrAmount > 0) {
        ocrAmount = qrAmount;
        ocrStatus = 'QR ✓';
      } else if (slipFileId) {
        const ocr = extractSlipAmount(slipFileId);
        ocrAmount = ocr.amount;
        ocrStatus = ocr.status;
      }
      console.log(`[submit] regId=${regId} cost=${cost} expected=${expectedTotal} ocr=${ocrAmount} status=${ocrStatus}`);

      // ตรวจ amount vs expected (รวม donation)
      const TOLERANCE = 5;
      if (expectedTotal > 0) {
        if (ocrAmount === 0) {
          ocrStatus = ocrStatus + ' (รอตรวจสอบมือ)';
        } else {
          const diff = ocrAmount - expectedTotal;
          // โอนเกิน → ไม่ block (frontend คุมแล้ว) แค่ log
          if (diff > TOLERANCE) {
            ocrStatus = `${ocrStatus} (โอนเกิน +${diff.toLocaleString()})`;
          }
          // โอนน้อย → block เว้นแต่ผู้ใช้ติด flag adminReview (ครั้งที่ 3+)
          else if (diff < -TOLERANCE) {
            if (payload.adminReviewFlag) {
              ocrStatus = `${ocrStatus} (รอแอดมินตรวจ — ขาด ${Math.abs(diff).toLocaleString()})`;
            } else {
              try { DriveApp.getFileById(slipFileId).setTrashed(true); } catch (e) {}
              return {
                ok: false,
                message: `❌ ยอดในสลิป (${ocrAmount.toLocaleString()} บาท) น้อยกว่ายอดที่ต้องโอน (${expectedTotal.toLocaleString()} บาท)\nกรุณาโอนเงินใหม่ให้ครบ แล้วแนบสลิปอีกครั้ง`,
                ocrAmount: ocrAmount,
                expectedCost: expectedTotal,
              };
            }
          } else {
            ocrStatus = `${ocrStatus} ตรง`;
          }
        }
      }

      // Anti-fraud
      const transRef = (payload.slipQrTransRef || '').toString().trim();
      if (transRef) {
        const existingRef = checkDuplicateTransRef(sheet, transRef);
        if (existingRef) {
          try { DriveApp.getFileById(slipFileId).setTrashed(true); } catch (e) {}
          return {
            ok: false,
            message: `❌ สลิปนี้ถูกใช้ไปแล้ว (${existingRef})\nกรุณาใช้สลิปการโอนครั้งใหม่`,
          };
        }
      }
    } else if (expectedTotal > 0) {
      return {
        ok: false,
        message: '❌ กรุณาแนบสลิปการโอนเงินก่อนกดลงทะเบียน',
      };
    }

    const row = buildRow(payload, regId, timestamp, cost, slipUrl, ocrAmount, ocrStatus, 0, '');
    sheet.appendRow(row);

    // เขียน donation rows (ถ้ามี supportPersonal/supportBatch)
    writeDonationRows(payload, regId, timestamp, slipUrl, payload.slipQrTransRef || '', ocrAmount, ocrStatus, 'พร้อมลงทะเบียน');

    return {
      ok: true,
      registrationId: regId,
      cost: cost,
      ocrAmount: ocrAmount,
      ocrStatus: ocrStatus,
      message: 'ลงทะเบียนสำเร็จ! แล้วเจอกันที่ศรีราชา 🌺',
    };

  } catch (err) {
    console.error('submitRegistration error:', err, err.stack);
    return { ok: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }
}

// ====== OCR ======

/**
 * OCR slip image → ดึงยอดเงิน
 * ต้อง enable Advanced Drive Service: Services → Drive API → Add
 */
function extractSlipAmount(fileId) {
  try {
    // ตรวจว่า Drive API ถูก enable หรือยัง
    if (typeof Drive === 'undefined') {
      return { amount: 0, status: 'Drive API ไม่ได้ enable' };
    }

    // ใช้ Drive.Files.copy with OCR เพื่อแปลง image → Google Doc
    const tempName = 'ocr_temp_' + fileId + '_' + Date.now();
    const copy = Drive.Files.copy({
      name: tempName,
      mimeType: 'application/vnd.google-apps.document'
    }, fileId, {
      ocr: true,
      ocrLanguage: 'th'
    });

    let text = '';
    try {
      const doc = DocumentApp.openById(copy.id);
      text = doc.getBody().getText();
    } catch (e) {
      console.error('Read OCR doc failed:', e);
    }

    // ลบไฟล์ temp
    try { Drive.Files.remove(copy.id); } catch (e) { console.warn('cleanup failed:', e); }

    if (!text) return { amount: 0, status: 'OCR ไม่พบข้อความ' };

    const amount = parseSlipAmount(text);
    if (amount > 0) {
      return { amount: amount, status: 'สำเร็จ' };
    }
    return { amount: 0, status: 'ไม่พบยอดเงิน' };

  } catch (e) {
    console.error('extractSlipAmount error:', e);
    return { amount: 0, status: 'ผิดพลาด: ' + e.message };
  }
}

/**
 * วิเคราะห์ text จาก slip → หายอดเงิน
 * Strategy:
 *  1. หา pattern "X,XXX.XX บาท" หรือ "บาท XXX.XX"
 *  2. หา "จำนวน:", "Amount:", "Total:" + เลข
 *  3. fallback: หาเลขใหญ่สุดในข้อความ
 */
function parseSlipAmount(text) {
  if (!text) return 0;
  // Normalize: ตัดช่องว่างซ้ำ + แปลงเลขไทย/อาหรับเป็น ASCII
  let normalized = text.replace(/\s+/g, ' ').trim()
    .replace(/[\u0E50-\u0E59]/g, d => d.charCodeAt(0) - 0x0E50)
    .replace(/[\u0660-\u0669]/g, d => d.charCodeAt(0) - 0x0660);

  const candidates = [];
  let m;

  // ===== Tier 5: keyword + เลข — ชัดเจนที่สุด =====
  const patternKeyword = /(?:จำนวน(?:เงิน)?|จํานวน(?:เงิน)?|amount|total|ยอด(?:เงิน|รวม|โอน|ที่โอน|ทั้งสิ้น)?|รวม(?:ทั้งสิ้น|ทั้งหมด)?|net amount|grand total)[:\s]*(?:THB|฿|บาท)?[:\s]*([\d,]+(?:\.\d{1,2})?)/gi;
  while ((m = patternKeyword.exec(normalized)) !== null) {
    const num = parseFloat(m[1].replace(/,/g, ''));
    if (num > 0 && num < 10000000) candidates.push({ value: num, score: 5 });
  }

  // ===== Tier 4: เลข + บาท / THB =====
  const patternBaht = /([\d,]+(?:\.\d{1,2})?)\s*(?:บาท|baht|THB)\b/gi;
  while ((m = patternBaht.exec(normalized)) !== null) {
    const num = parseFloat(m[1].replace(/,/g, ''));
    if (num > 0 && num < 10000000) candidates.push({ value: num, score: 4 });
  }

  // ===== Tier 4: ฿ / THB ก่อนเลข =====
  const patternThb = /(?:THB|฿)\s*([\d,]+(?:\.\d{1,2})?)/gi;
  while ((m = patternThb.exec(normalized)) !== null) {
    const num = parseFloat(m[1].replace(/,/g, ''));
    if (num > 0 && num < 10000000) candidates.push({ value: num, score: 4 });
  }

  // ===== Tier 3: keyword โอน/รับ/paid/sent =====
  const patternTransfer = /(?:โอน(?:เงิน|ออก|เข้า|ไป|ให้)?|รับ(?:เงิน|โอน)?|paid|sent|transfer(?:red)?)[:\s]*(?:THB|฿)?[:\s]*([\d,]+(?:\.\d{1,2})?)/gi;
  while ((m = patternTransfer.exec(normalized)) !== null) {
    const num = parseFloat(m[1].replace(/,/g, ''));
    if (num > 0 && num < 10000000) candidates.push({ value: num, score: 3 });
  }

  if (candidates.length > 0) {
    candidates.sort((a, b) => b.score - a.score || b.value - a.value);
    return candidates[0].value;
  }

  // ===== Fallback: เลขที่ดูเหมือนยอดเงิน (มี comma หรือ decimal) =====
  const moneyPattern = /([\d]{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+\.\d{2})/g;
  const numbers = [];
  while ((m = moneyPattern.exec(normalized)) !== null) {
    const num = parseFloat(m[1].replace(/,/g, ''));
    if (num >= 1 && num < 10000000) numbers.push(num);
  }
  if (numbers.length > 0) {
    // กรองที่อยู่ในช่วงสมเหตุสมผล (10-100,000 บาท)
    const reasonable = numbers.filter(n => n >= 10 && n <= 100000);
    if (reasonable.length > 0) return Math.max(...reasonable);
    return Math.max(...numbers);
  }

  return 0;
}

// ====== EDIT REGISTRATION ======

/**
 * ค้นหา registration จาก email + phone
 * อีเมลกับเบอร์ต้องเป็นคู่ของบุคคลเดียวกัน (เจ้าของใบ หรือ รูมเมท)
 */
function findRegistration(email, phone) {
  try {
    if (!email || !phone) {
      return { ok: false, message: 'กรุณากรอกอีเมลและเบอร์โทร' };
    }
    const emailLower = String(email).trim().toLowerCase();
    const phoneClean = String(phone).replace(/[\s-]/g, '').replace(/^0+/, '');

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

    // วน 4 tabs
    for (const tabName of ALL_REG_TABS) {
      const sheet = ss.getSheetByName(tabName);
      if (!sheet || sheet.getLastRow() < 2) continue;
      const headers = getHeadersOfSheet(sheet);
      const idxEmail    = headers.indexOf('อีเมล');
      const idxPhone    = headers.indexOf('เบอร์โทร');
      const idxRmEmail  = headers.indexOf('รูมเมท - อีเมล');  // อาจ -1 สำหรับน้อง 34/พี่ 33
      const idxRmPhone  = headers.indexOf('รูมเมท - เบอร์');
      const idxRegId    = headers.indexOf('Registration ID');
      const idxEditCnt  = headers.indexOf('จำนวนครั้งที่แก้ไข');

      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const ownerEmail = String(row[idxEmail] || '').toLowerCase();
        const ownerPhone = String(row[idxPhone] || '').replace(/[\s-]/g, '').replace(/^0+/, '');
        const rmEmail    = idxRmEmail >= 0 ? String(row[idxRmEmail] || '').toLowerCase() : '';
        const rmPhone    = idxRmPhone >= 0 ? String(row[idxRmPhone] || '').replace(/[\s-]/g, '').replace(/^0+/, '') : '';

        const isOwner = ownerEmail === emailLower && ownerPhone === phoneClean;
        const isRoommate = rmEmail && rmPhone && rmEmail === emailLower && rmPhone === phoneClean;

        if (isOwner || isRoommate) {
          const editCount = parseInt(row[idxEditCnt], 10) || 0;
          if (editCount >= 1) {
            return {
              ok: false,
              message: 'คุณได้แก้ไขข้อมูลแล้ว 1 ครั้ง · หากต้องการแก้ไขเพิ่มเติม กรุณาติดต่อทีมงาน',
            };
          }
          return {
            ok: true,
            registrationId: row[idxRegId],
            tabName: tabName,
            rowNumber: i + 2,
            data: rowToPayload(row, headers),
          };
        }
      }
    }
    return { ok: false, message: 'ไม่พบข้อมูล กรุณาตรวจสอบอีเมลและเบอร์โทร' };
  } catch (e) {
    console.error('findRegistration error:', e);
    return { ok: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

/**
 * อัปเดต registration — limit 1 ครั้ง
 */
function updateRegistration(regId, payload) {
  try {
    if (!regId) return { ok: false, message: 'ไม่พบ Registration ID' };

    const validation = validatePayload(payload);
    if (!validation.ok) return validation;

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

    // ค้นทุก tab หา regId
    let sheet = null;
    let headers = null;
    let rowIndex = -1;
    let originalRow = null;
    for (const tabName of ALL_REG_TABS) {
      const s = ss.getSheetByName(tabName);
      if (!s || s.getLastRow() < 2) continue;
      const h = getHeadersOfSheet(s);
      const idxId = h.indexOf('Registration ID');
      const data = s.getRange(2, 1, s.getLastRow() - 1, h.length).getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][idxId] === regId) {
          sheet = s;
          headers = h;
          rowIndex = i + 2;
          originalRow = data[i];
          break;
        }
      }
      if (sheet) break;
    }
    if (!sheet) return { ok: false, message: 'ไม่พบข้อมูล' };

    const idxEditCnt = headers.indexOf('จำนวนครั้งที่แก้ไข');
    const idxOrigTs  = headers.indexOf('วันที่ลงทะเบียน');
    const idxSlipUrl = headers.indexOf('สลิปโอน');
    const idxOcr     = headers.indexOf('ยอดโอนจริง OCR (บาท)');
    const idxOcrStat = headers.indexOf('สถานะ OCR');

    const editCount = parseInt(originalRow[idxEditCnt], 10) || 0;
    if (editCount >= 1) {
      return {
        ok: false,
        message: 'คุณได้แก้ไขข้อมูลแล้ว 1 ครั้ง · หากต้องการแก้ไขเพิ่มเติม กรุณาติดต่อทีมงาน',
      };
    }

    const cost = calculateCost(payload);

    // Slip handling: ถ้ามี slip ใหม่ → upload + OCR ใหม่; ไม่ใส่ → เก็บค่าเดิม
    let slipUrl = originalRow[idxSlipUrl] || '';
    let ocrAmount = parseFloat(originalRow[idxOcr]) || 0;
    let ocrStatus = originalRow[idxOcrStat] || 'ไม่มีสลิป';

    if (payload.slip && payload.slip.data) {
      const result = uploadSlip(payload.slip, regId, payload);
      const newSlipUrl = result.url;
      const newSlipFileId = result.fileId;
      let newOcrAmount = 0;
      let newOcrStatus = 'ไม่มีสลิป';

      // QR ก่อน
      const qrAmount = parseFloat(payload.slipQrAmount);
      if (!isNaN(qrAmount) && qrAmount > 0) {
        newOcrAmount = qrAmount;
        newOcrStatus = 'QR ✓';
      } else if (newSlipFileId) {
        const ocr = extractSlipAmount(newSlipFileId);
        newOcrAmount = ocr.amount;
        newOcrStatus = ocr.status;
      }

      const TOLERANCE = 5;
      if (cost > 0) {
        if (newOcrAmount === 0) {
          newOcrStatus = newOcrStatus + ' (รอตรวจสอบมือ)';
        } else {
          const diff = newOcrAmount - cost;
          if (diff < -TOLERANCE) {
            try { DriveApp.getFileById(newSlipFileId).setTrashed(true); } catch (e) {}
            return {
              ok: false,
              message: `❌ ยอดในสลิป (${newOcrAmount.toLocaleString()} บาท) น้อยกว่ายอดที่ต้องโอน (${cost.toLocaleString()} บาท)\nกรุณาโอนเงินใหม่ให้ครบ แล้วแนบสลิปอีกครั้ง`,
              ocrAmount: newOcrAmount,
              expectedCost: cost,
            };
          } else if (diff > TOLERANCE) {
            newOcrStatus = `${newOcrStatus} (โอนเกิน +${diff.toLocaleString()})`;
          } else {
            newOcrStatus = `${newOcrStatus} ตรง`;
          }
        }
      }

      slipUrl = newSlipUrl;
      ocrAmount = newOcrAmount;
      ocrStatus = newOcrStatus;
    } else if (cost > 0 && !slipUrl) {
      return {
        ok: false,
        message: '❌ กรุณาแนบสลิปการโอนเงินก่อนบันทึก',
      };
    }

    const now = new Date();
    const editTime = Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
    const origTs = originalRow[idxOrigTs];

    const newRow = buildRow(payload, regId, origTs, cost, slipUrl, ocrAmount, ocrStatus, editCount + 1, editTime);
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([newRow]);

    return {
      ok: true,
      registrationId: regId,
      cost: cost,
      ocrAmount: ocrAmount,
      ocrStatus: ocrStatus,
      message: 'แก้ไขข้อมูลสำเร็จ! 🌺',
    };
  } catch (err) {
    console.error('updateRegistration error:', err, err.stack);
    return { ok: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }
}

/**
 * แปลงแถวข้อมูลกลับเป็น payload object — เพื่อ pre-fill ในฟอร์มแก้ไข
 */
function rowToPayload(row, headers) {
  headers = headers || HEADERS;
  const get = name => {
    const idx = headers.indexOf(name);
    return idx >= 0 ? (row[idx] || '') : '';
  };
  const groupReverse = {
    'น้อง CMMBA รุ่น 34': 'student34',
    'พี่ CMMBA รุ่น 33': 'senior33',
    'ศิษย์เก่า CMMBA รุ่น 1–32': 'alumni',
    'อาจารย์และบุคลากร': 'faculty',
  };
  const roomReverse = {
    'พักคู่ (มาตรฐาน)': 'twin',
    'พักเดี่ยว/มีผู้ติดตาม (+1,000)': 'single',
    'ไม่ค้างคืน': 'nostay',
  };
  const bedReverse = { 'เตียงใหญ่': 'king', 'เตียงแยก': 'twin' };

  return {
    group: groupReverse[get('กลุ่ม')] || '',
    title: get('คำนำหน้า'),
    firstName: get('ชื่อ'),
    lastName: get('นามสกุล'),
    fullName: get('ชื่อ-นามสกุล'),
    nickname: get('ชื่อเล่น'),
    studentId: get('รหัสนักศึกษา'),
    nationalId: get('เลขบัตรประชาชน'),
    alumniBatch: get('CMMBA รุ่นที่'),
    position: get('ตำแหน่ง/สังกัด'),
    gender: get('เพศ'),
    birthDate: get('วันเกิด'),
    phone: get('เบอร์โทร'),
    email: get('อีเมล'),
    dietRestrictions: parseChoices(get('ข้อจำกัดอาหาร')).list,
    dietOther: parseChoices(get('ข้อจำกัดอาหาร')).other,
    allergies: parseChoices(get('แพ้อาหาร')).list,
    allergyOther: parseChoices(get('แพ้อาหาร')).other,
    medicalConditions: parseChoices(get('โรคประจำตัว')).list,
    medicalOther: parseChoices(get('โรคประจำตัว')).other,
    medications: parseChoices(get('ยาประจำตัว')).list,
    medicationOther: parseChoices(get('ยาประจำตัว')).other,
    roomType: roomReverse[get('ประเภทห้องพัก')] || '',
    bedType: bedReverse[get('ประเภทเตียง')] || '',
    roommate: {
      firstName: get('รูมเมท - ชื่อ'),
      lastName: get('รูมเมท - นามสกุล'),
      fullName: get('รูมเมท - ชื่อ-นามสกุล'),
      nickname: get('รูมเมท - ชื่อเล่น'),
      idOrBatch: get('รูมเมท - รหัส/รุ่น'),
      gender: get('รูมเมท - เพศ'),
      phone: get('รูมเมท - เบอร์'),
      email: get('รูมเมท - อีเมล'),
      dietRestrictions: parseChoices(get('รูมเมท - ข้อจำกัดอาหาร')).list,
      dietOther: parseChoices(get('รูมเมท - ข้อจำกัดอาหาร')).other,
      allergies: parseChoices(get('รูมเมท - แพ้อาหาร')).list,
      allergyOther: parseChoices(get('รูมเมท - แพ้อาหาร')).other,
      medicalConditions: parseChoices(get('รูมเมท - โรคประจำตัว')).list,
      medicalOther: parseChoices(get('รูมเมท - โรคประจำตัว')).other,
      medications: parseChoices(get('รูมเมท - ยา')).list,
      medicationOther: parseChoices(get('รูมเมท - ยา')).other,
    },
    hasCompanion: get('มีผู้ติดตาม') === 'มี',
    companion: {
      adultCount: get('ผู้ติดตาม - ผู้ใหญ่'),
      childCount: get('ผู้ติดตาม - เด็ก ≤12'),
      details: get('ผู้ติดตาม - รายละเอียด'),
      dietNotes: get('ผู้ติดตาม - ข้อจำกัดอาหาร'),
      allergyNotes: get('ผู้ติดตาม - แพ้อาหาร'),
      medicalNotes: get('ผู้ติดตาม - โรค/ยา'),
    },
    emergencyName: get('ผู้ติดต่อฉุกเฉิน - ชื่อ'),
    emergencyRelation: get('ผู้ติดต่อฉุกเฉิน - ความสัมพันธ์'),
    emergencyPhone: get('ผู้ติดต่อฉุกเฉิน - เบอร์'),
    consentPhoto: get('ยินยอมใช้ภาพ') === 'ยินยอม',
    consentPdpa: get('ยินยอม PDPA (เลขบัตร)') === 'ยินยอม',
    notes: get('หมายเหตุ'),
  };
}

/**
 * Parse "ตัวเลือก1, ตัวเลือก2; อื่นๆ: text" กลับเป็น { list, other }
 */
function parseChoices(str) {
  if (!str) return { list: [], other: '' };
  const s = String(str);
  const otherMatch = s.match(/อื่นๆ:\s*(.+)$/);
  const other = otherMatch ? otherMatch[1].trim() : '';
  const listPart = s.replace(/;?\s*อื่นๆ:.+$/, '').trim();
  const list = listPart ? listPart.split(',').map(x => x.trim()).filter(Boolean) : [];
  return { list, other };
}

// ====== HELPERS ======

function ensureSheetReady(ss) {
  // สร้างทุก group tab พร้อม HEADERS ของกลุ่มนั้นๆ
  Object.keys(GROUP_TABS).forEach(group => {
    const name = GROUP_TABS[group];
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    initRegSheet(sheet, group);
  });

  let donSheet = ss.getSheetByName(DON_TAB);
  if (!donSheet) donSheet = ss.insertSheet(DON_TAB);
  if (donSheet.getLastRow() === 0) {
    donSheet.getRange(1, 1, 1, DON_HEADERS.length).setValues([DON_HEADERS]);
    donSheet.getRange(1, 1, 1, DON_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#D97706')
      .setFontColor('#FFFFFF');
    donSheet.setFrozenRows(1);
  }

  // return registration sheet (default = student34) — caller จะ override ตาม group
  return ss.getSheetByName(GROUP_TABS.student34);
}

const GROUP_PREFIX = {
  student34: 'S34',
  senior33:  'S33',
  alumni:    'AL',
  faculty:   'FAC',
};

function generateRegistrationId(ss, sheet, group) {
  // นับ row ทั้งหมดใน 4 tabs รวมกัน เพื่อ unique number ต่อเนื่อง
  let total = 0;
  ALL_REG_TABS.forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && s.getLastRow() > 1) total += (s.getLastRow() - 1);
  });
  const prefix = GROUP_PREFIX[group] || 'REG';
  return `${prefix}-${String(total + 1).padStart(4, '0')}`;
}

function checkDuplicateTransRef(_unused, transRef) {
  if (!transRef) return null;
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  // เช็คทุก registration tab
  for (const name of ALL_REG_TABS) {
    const s = ss.getSheetByName(name);
    if (!s || s.getLastRow() < 2) continue;
    const headers = getHeadersOfSheet(s);
    const idxRef = headers.indexOf('QR Ref');
    const idxRegId = headers.indexOf('Registration ID');
    if (idxRef < 0) continue;
    const data = s.getRange(2, 1, s.getLastRow() - 1, headers.length).getValues();
    for (let i = 0; i < data.length; i++) {
      const existing = (data[i][idxRef] || '').toString().trim();
      if (existing && existing === transRef) {
        return data[i][idxRegId] || 'ลงทะเบียนเดิม';
      }
    }
  }
  // เช็ค donation
  const donSheet = ss.getSheetByName(DON_TAB);
  if (donSheet && donSheet.getLastRow() >= 2) {
    const idxRef = DON_HEADERS.indexOf('QR Ref');
    const idxDonId = DON_HEADERS.indexOf('Donation ID');
    if (idxRef >= 0) {
      const data = donSheet.getRange(2, 1, donSheet.getLastRow() - 1, DON_HEADERS.length).getValues();
      for (let i = 0; i < data.length; i++) {
        const existing = (data[i][idxRef] || '').toString().trim();
        if (existing && existing === transRef) {
          return data[i][idxDonId] || 'รายการบริจาคเดิม';
        }
      }
    }
  }
  return null;
}

function buildDonationRow(p, donId, regId, ts, inNameOf, amount, slipUrl, qrRef, ocrAmount, ocrStatus, source) {
  return [
    donId, ts, regId || '',
    p.fullName || '', p.nickname || '',
    labelGroup(p.group), deriveBatchName(p),
    p.phone || '', p.email || '',
    inNameOf,
    amount,
    slipUrl || '',
    qrRef || '',
    ocrAmount || 0,
    ocrStatus || '',
    source,
  ];
}

function writeDonationRows(p, regId, ts, slipUrl, qrRef, ocrAmount, ocrStatus, source) {
  // เขียน donation rows ตาม supportPersonal + supportBatch
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let donSheet = ss.getSheetByName(DON_TAB);
  if (!donSheet) {
    donSheet = ss.insertSheet(DON_TAB);
    donSheet.getRange(1, 1, 1, DON_HEADERS.length).setValues([DON_HEADERS]).setFontWeight('bold');
    donSheet.setFrozenRows(1);
  }

  const sp = parseFloat(p.supportPersonal) || 0;
  const sb = parseFloat(p.supportBatch) || 0;
  const total = sp + sb;
  if (total === 0) return [];

  const created = [];
  // หา running id
  const lastRow = donSheet.getLastRow();
  let nextNum = lastRow; // 1-based, header = row 1

  if (sp > 0) {
    nextNum++;
    const id = 'DON-' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMM') + '-' + String(nextNum).padStart(4, '0');
    donSheet.appendRow(buildDonationRow(p, id, regId, ts, 'ส่วนตัว', sp, slipUrl, qrRef, ocrAmount, ocrStatus, source));
    created.push(id);
  }
  if (sb > 0) {
    nextNum++;
    const id = 'DON-' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMM') + '-' + String(nextNum).padStart(4, '0');
    donSheet.appendRow(buildDonationRow(p, id, regId, ts, 'รุ่น (' + deriveBatchName(p) + ')', sb, slipUrl, qrRef, ocrAmount, ocrStatus, source));
    created.push(id);
  }
  return created;
}

/**
 * Endpoint สำหรับบริจาคเพิ่มเติม (หลังลงทะเบียนแล้ว)
 * payload: { email, phone, supportPersonal, supportBatch, slip, slipQrAmount, slipQrTransRef }
 */
function submitDonation(payload) {
  try {
    if (!isDonationOpen()) {
      return { ok: false, message: 'ปิดรับบริจาคแล้ว — งานจบเมื่อวันที่ 2 สิงหาคม 2569 🙏 ขอบคุณทุกท่าน' };
    }
    if (!payload.email && !payload.phone) {
      return { ok: false, message: 'กรุณากรอกอีเมลหรือเบอร์โทรอย่างน้อย 1 อย่าง' };
    }
    const found = findRegistrationForDonation(payload.email, payload.phone);
    if (!found.ok || !found.data) {
      return {
        ok: false,
        message: found.message || 'ไม่พบข้อมูลลงทะเบียนของท่าน — กรุณาลงทะเบียนก่อน',
      };
    }
    const reg = found.data;
    const regId = found.registrationId;

    const sp = parseFloat(payload.supportPersonal) || 0;
    const sb = parseFloat(payload.supportBatch) || 0;
    const total = sp + sb;
    if (total <= 0) {
      return { ok: false, message: 'กรุณากรอกยอดบริจาคอย่างน้อย 1 ช่อง' };
    }
    if (!payload.slip || !payload.slip.data) {
      return { ok: false, message: 'กรุณาแนบสลิปการโอนเงิน' };
    }

    const mergedPayload = Object.assign({}, reg, {
      supportPersonal: sp,
      supportBatch: sb,
    });
    const result = uploadSlip(payload.slip, regId + '-DON', mergedPayload);
    const slipUrl = result.url;
    const slipFileId = result.fileId;

    let ocrAmount = 0;
    let ocrStatus = 'ไม่มีสลิป';
    const qrAmount = parseFloat(payload.slipQrAmount);
    if (!isNaN(qrAmount) && qrAmount > 0) {
      ocrAmount = qrAmount;
      ocrStatus = 'QR ✓';
    } else if (slipFileId) {
      const ocr = extractSlipAmount(slipFileId);
      ocrAmount = ocr.amount;
      ocrStatus = ocr.status;
    }

    const TOLERANCE = 5;
    if (ocrAmount > 0) {
      const diff = ocrAmount - total;
      if (diff < -TOLERANCE) {
        try { DriveApp.getFileById(slipFileId).setTrashed(true); } catch (e) {}
        return {
          ok: false,
          message: `❌ ยอดในสลิป (${ocrAmount.toLocaleString()}) น้อยกว่าที่ต้องโอน (${total.toLocaleString()})`,
          ocrAmount: ocrAmount,
          expectedCost: total,
        };
      } else if (diff > TOLERANCE) {
        ocrStatus = `${ocrStatus} (โอนเกิน +${diff.toLocaleString()})`;
      } else {
        ocrStatus = `${ocrStatus} ตรง`;
      }
    } else {
      ocrStatus = ocrStatus + ' (รอตรวจสอบมือ)';
    }

    // Anti-fraud
    const transRef = (payload.slipQrTransRef || '').toString().trim();
    if (transRef) {
      const existing = checkDuplicateTransRef(null, transRef);
      if (existing) {
        try { DriveApp.getFileById(slipFileId).setTrashed(true); } catch (e) {}
        return {
          ok: false,
          message: `❌ สลิปนี้ถูกใช้ไปแล้ว (${existing})\nกรุณาใช้สลิปการโอนครั้งใหม่`,
        };
      }
    }

    const now = new Date();
    const ts = Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
    const created = writeDonationRows(mergedPayload, regId, ts, slipUrl, transRef, ocrAmount, ocrStatus, 'บริจาคเพิ่มเติม');

    return {
      ok: true,
      donationIds: created,
      total: total,
      ocrAmount: ocrAmount,
      ocrStatus: ocrStatus,
      registeredName: reg.fullName,
      message: 'ขอบคุณสำหรับการสนับสนุน 🌺',
    };
  } catch (err) {
    console.error('submitDonation error:', err, err.stack);
    return { ok: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }
}

/**
 * Flexible lookup สำหรับหน้าบริจาคเพิ่มเติม (เงินโอน) — รับแค่ email หรือ phone อย่างใดอย่างหนึ่ง
 * Phone normalization: strip leading 0 ทั้ง input และ sheet (Google Sheets อาจตัด 0 หน้าออก)
 * ไม่ตรวจ editCount (ไม่เกี่ยวกับการแก้ไข)
 */
function findRegistrationForDonation(email, phone) {
  try {
    if (!email && !phone) {
      return { ok: false, message: 'กรุณากรอกอีเมลหรือเบอร์โทรอย่างน้อย 1 อย่าง' };
    }
    var emailLower = email ? String(email).trim().toLowerCase() : '';
    var phoneClean = phone ? String(phone).replace(/[\s-]/g, '').replace(/^0+/, '') : '';

    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

    for (var t = 0; t < ALL_REG_TABS.length; t++) {
      var sheet = ss.getSheetByName(ALL_REG_TABS[t]);
      if (!sheet || sheet.getLastRow() < 2) continue;
      var headers = getHeadersOfSheet(sheet);
      var idxEmail    = headers.indexOf('อีเมล');
      var idxPhone    = headers.indexOf('เบอร์โทร');
      var idxRmEmail  = headers.indexOf('รูมเมท - อีเมล');
      var idxRmPhone  = headers.indexOf('รูมเมท - เบอร์');
      var idxRegId    = headers.indexOf('Registration ID');

      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var ownerEmail = String(row[idxEmail] || '').toLowerCase();
        var ownerPhone = String(row[idxPhone] || '').replace(/[\s-]/g, '').replace(/^0+/, '');
        var rmEmail    = idxRmEmail >= 0 ? String(row[idxRmEmail] || '').toLowerCase() : '';
        var rmPhone    = idxRmPhone >= 0 ? String(row[idxRmPhone] || '').replace(/[\s-]/g, '').replace(/^0+/, '') : '';

        // ทุก field ที่ user กรอก ต้องตรง — กรอกแค่ 1 → match แค่ 1
        var ownerMatch = (emailLower || phoneClean) &&
                         (!emailLower || ownerEmail === emailLower) &&
                         (!phoneClean || ownerPhone === phoneClean);
        var rmMatch = (rmEmail || rmPhone) && (emailLower || phoneClean) &&
                      (!emailLower || rmEmail === emailLower) &&
                      (!phoneClean || rmPhone === phoneClean);

        if (ownerMatch || rmMatch) {
          return {
            ok: true,
            registrationId: row[idxRegId],
            tabName: ALL_REG_TABS[t],
            rowNumber: i + 2,
            data: rowToPayload(row, headers),
          };
        }
      }
    }
    return { ok: false, message: 'ไม่พบข้อมูล · ตรวจสอบอีเมลหรือเบอร์โทรอีกครั้ง' };
  } catch (e) {
    console.error('findRegistrationForDonation error:', e);
    return { ok: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

/**
 * บันทึกการบริจาคเงินสด — สำหรับเจ้าหน้าที่
 * payload = { password, batchLabel, amount }
 *   batchLabel: ค่าที่ aggregate ใน Top batch leaderboard เช่น "CMMBA รุ่น 33", "อาจารย์ CMMBA"
 */
/**
 * ตรวจรหัสผ่านเจ้าหน้าที่ (เรียกตอน login เงินสด — กันผ่านหน้าฟอร์มถ้ารหัสผิด)
 */
function verifyStaffPassword(password) {
  return { ok: String(password || '') === String(CONFIG.STAFF_PASSWORD) };
}

function submitCashDonation(payload) {
  try {
    payload = payload || {};
    if (String(payload.password || '') !== String(CONFIG.STAFF_PASSWORD)) {
      return { ok: false, message: 'รหัสผ่านไม่ถูกต้อง' };
    }
    var batchLabel = String(payload.batchLabel || '').trim();
    var amount = parseFloat(payload.amount);
    if (!batchLabel) return { ok: false, message: 'กรุณาเลือกรุ่น' };
    if (!amount || amount <= 0) return { ok: false, message: 'กรุณากรอกจำนวนเงินที่ถูกต้อง' };

    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var donSheet = ss.getSheetByName(DON_TAB);
    if (!donSheet) {
      donSheet = ss.insertSheet(DON_TAB);
      donSheet.getRange(1, 1, 1, DON_HEADERS.length).setValues([DON_HEADERS]).setFontWeight('bold');
      donSheet.setFrozenRows(1);
    }

    var ts = new Date();
    var nextNum = donSheet.getLastRow() + 1;
    var donId = 'DON-' + Utilities.formatDate(ts, CONFIG.TZ, 'yyyyMM') + '-' + String(nextNum - 1).padStart(4, '0');

    // map header → value (เขียน row ที่จะนับใน Top batch leaderboard ได้)
    var rowMap = {};
    rowMap['Donation ID']         = donId;
    rowMap['Timestamp']           = ts;
    rowMap['Registration ID']     = '';
    rowMap['ชื่อ-นามสกุล']         = '';
    rowMap['ชื่อเล่น']              = '';
    rowMap['กลุ่ม']                = 'cash';
    rowMap['รุ่น']                 = batchLabel;          // ← key ที่ getTopBatchSpenders ใช้ aggregate
    rowMap['เบอร์โทร']             = '';
    rowMap['อีเมล']                = '';
    rowMap['บริจาคในนาม']         = 'รุ่น (' + batchLabel + ')';   // ← ต้องขึ้นต้น "รุ่น" จึงนับ
    rowMap['ยอดบริจาค (บาท)']      = amount;
    rowMap['สลิป (URL)']           = '';
    rowMap['QR transRef']         = 'CASH-' + donId;
    rowMap['ยอด OCR (บาท)']       = amount;             // ← นับใน getDonationTotal()
    rowMap['สถานะตรวจสอบ']         = 'cash-verified';
    rowMap['ที่มา']                = 'cash';

    // build row ตามลำดับ DON_HEADERS
    var rowArr = DON_HEADERS.map(function(h) { return rowMap.hasOwnProperty(h) ? rowMap[h] : ''; });
    donSheet.appendRow(rowArr);

    return { ok: true, donationId: donId, message: 'บันทึกการบริจาคเงินสดสำเร็จ' };
  } catch (e) {
    console.error('submitCashDonation error:', e);
    return { ok: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

/**
 * บันทึกการบริจาคแบบ "โอนด่วน" — ไม่ต้องลงทะเบียนก่อน
 * payload = {
 *   slip: { name, mimeType, data(base64) },
 *   slipQrAmount?: number,
 *   slipQrTransRef?: string,
 *   donor?: { firstName, lastName, nickname, batch }  // ถ้าไม่มี = ไม่ประสงค์ออกนาม
 * }
 */
function submitQuickDonation(payload) {
  try {
    payload = payload || {};
    if (!isDonationOpen()) {
      return { ok: false, message: 'ปิดรับบริจาคแล้ว — งานจบเมื่อวันที่ 2 สิงหาคม 2569 🙏 ขอบคุณทุกท่าน' };
    }
    if (!payload.slip || !payload.slip.data) {
      return { ok: false, message: 'กรุณาแนบสลิปการโอนเงิน' };
    }

    var donor = payload.donor || null;
    var firstName = donor ? String(donor.firstName || '').trim() : '';
    var lastName  = donor ? String(donor.lastName  || '').trim() : '';
    var nickname  = donor ? String(donor.nickname  || '').trim() : '';
    var batch     = donor ? String(donor.batch     || '').trim() : '';
    var fullName  = (firstName + ' ' + lastName).trim();
    var isAnon = !donor || (!fullName && !nickname && !batch);

    var displayName = isAnon ? 'ไม่ประสงค์ออกนาม' : (fullName || nickname || '-');

    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var donSheet = ss.getSheetByName(DON_TAB);
    if (!donSheet) {
      donSheet = ss.insertSheet(DON_TAB);
      donSheet.getRange(1, 1, 1, DON_HEADERS.length).setValues([DON_HEADERS]).setFontWeight('bold');
      donSheet.setFrozenRows(1);
    }
    var ts = new Date();
    var nextNum = donSheet.getLastRow() + 1;
    var donId = 'DON-' + Utilities.formatDate(ts, CONFIG.TZ, 'yyyyMM') + '-' + String(nextNum - 1).padStart(4, '0') + '-Q';

    var pseudoP = {
      fullName: isAnon ? 'anonymous' : (fullName || nickname || 'donor'),
      supportPersonal: 0,
      supportBatch: 0,
      group: 'quick',
      alumniBatch: batch,
    };
    var slipResult = uploadSlip(payload.slip, donId, pseudoP);
    var slipUrl = slipResult.url;
    var slipFileId = slipResult.fileId;

    var ocrAmount = 0;
    var ocrStatus = 'ไม่มีสลิป';
    var qrAmount = parseFloat(payload.slipQrAmount);
    if (!isNaN(qrAmount) && qrAmount > 0) {
      ocrAmount = qrAmount;
      ocrStatus = 'QR ✓';
    } else if (slipFileId) {
      try {
        var ocr = extractSlipAmount(slipFileId);
        ocrAmount = ocr.amount || 0;
        ocrStatus = ocr.status || 'OCR fail';
      } catch (e) {
        ocrStatus = 'OCR error';
      }
    }

    var transRef = String(payload.slipQrTransRef || '').trim();
    if (transRef) {
      try {
        var existing = checkDuplicateTransRef(null, transRef);
        if (existing) {
          try { DriveApp.getFileById(slipFileId).setTrashed(true); } catch (e) {}
          return { ok: false, message: 'สลิปนี้ถูกใช้ไปแล้ว (' + existing + ')' };
        }
      } catch (e) {}
    }

    var rowMap = {};
    rowMap['Donation ID']        = donId;
    rowMap['วันที่บริจาค']        = ts;
    rowMap['Registration ID']    = '';
    rowMap['ชื่อ-นามสกุล']        = isAnon ? '' : fullName;
    rowMap['ชื่อเล่น']             = isAnon ? '' : nickname;
    rowMap['กลุ่ม']               = 'quick';
    rowMap['รุ่น']                = isAnon ? '' : batch;
    rowMap['เบอร์โทร']            = '';
    rowMap['อีเมล']               = '';
    rowMap['บริจาคในนาม']        = isAnon ? 'ไม่ประสงค์ออกนาม' : (displayName + (batch ? ' (' + batch + ')' : ''));
    rowMap['ยอดบริจาค (บาท)']     = ocrAmount;
    rowMap['สลิปโอน']            = slipUrl;
    rowMap['QR Ref']             = transRef;
    rowMap['ยอดโอน OCR']         = ocrAmount;
    rowMap['สถานะ OCR']          = ocrStatus;
    rowMap['แหล่งที่มา']         = 'โอนด่วน';

    var rowArr = DON_HEADERS.map(function(h) { return rowMap.hasOwnProperty(h) ? rowMap[h] : ''; });
    donSheet.appendRow(rowArr);

    return {
      ok: true,
      donationId: donId,
      ocrAmount: ocrAmount,
      ocrStatus: ocrStatus,
      message: 'ขอบคุณสำหรับการสนับสนุน 🌺',
    };
  } catch (err) {
    console.error('submitQuickDonation error:', err, err.stack);
    return { ok: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }
}

function generateDonationId_(ss) {
  var sheet = ss.getSheetByName(DON_TAB);
  var n = sheet.getLastRow();
  return 'D' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyMMddHHmmss') + n;
}

function lookupForDonation(email, phone) {
  const found = findRegistrationForDonation(email, phone);
  if (!found.ok || !found.data) return { ok: false, message: found.message || 'ไม่พบข้อมูล' };
  const reg = found.data;
  return {
    ok: true,
    registrationId: found.registrationId,
    fullName: reg.fullName,
    nickname: reg.nickname,
    group: reg.group,
    alumniBatch: reg.alumniBatch || '',
    batchName: deriveBatchName(reg),
  };
}

function deriveBatchName(p) {
  const g = p.group;
  if (g === 'student34') return 'CMMBA รุ่น 34';
  if (g === 'senior33')  return 'CMMBA รุ่น 33';
  if (g === 'alumni') {
    const b = (p.alumniBatch || '').toString().trim();
    return b ? `CMMBA ${b}` : 'CMMBA ศิษย์เก่า';
  }
  if (g === 'faculty')   return 'อาจารย์ CMMBA';
  return 'CMMBA';
}

function calculateCost(p) {
  let cost = 0;
  const group = p.group;
  const roomType = p.roomType;

  if (group === 'faculty' || group === 'senior33' || group === 'student34') {
    return 0;
  }

  if (group === 'alumni') {
    if (roomType === 'single') {
      if (p.companion) {
        const adults = parseInt(p.companion.adultCount, 10) || 0;
        const children = parseInt(p.companion.childCount, 10) || 0;
        cost += adults * 1000 + children * 500;
      } else {
        cost += 1000;
      }
    }
  }

  return cost;
}

// ยอดที่ user ต้องโอน = cost (ลงทะเบียน) + supportPersonal + supportBatch
function totalAmount(p) {
  const supportP = parseFloat(p.supportPersonal) || 0;
  const supportB = parseFloat(p.supportBatch) || 0;
  return calculateCost(p) + supportP + supportB;
}

function uploadSlip(slip, regId, p) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.SLIP_FOLDER_ID);
    const contentType = slip.mimeType || 'image/jpeg';
    const ext = (slip.name && slip.name.includes('.'))
      ? slip.name.split('.').pop()
      : (contentType.split('/')[1] || 'jpg');

    const sanitize = (s) => String(s || '')
      .replace(/[^\w\u0E00-\u0E7F-]/g, '-')
      .replace(/-+/g, '-')
      .replace(/^-|-$/g, '')
      .slice(0, 60);

    const userName = sanitize(p.fullName) || 'unknown';
    const batchName = sanitize(deriveBatchName(p));
    const personalAmount = parseFloat(p.supportPersonal) || 0;
    const batchAmount = parseFloat(p.supportBatch) || 0;

    let filename;
    if (batchAmount > 0 && personalAmount === 0) {
      // รุ่น 100%
      filename = `รุ่น-${batchName}_${userName}_${regId}.${ext}`;
    } else if (batchAmount > 0 && personalAmount > 0) {
      // ผสม
      filename = `ผสม-${batchName}_${userName}_${regId}.${ext}`;
    } else {
      // ส่วนตัว / ไม่มี support
      filename = `${userName}_${regId}.${ext}`;
    }

    const blob = Utilities.newBlob(Utilities.base64Decode(slip.data), contentType, filename);
    const file = folder.createFile(blob);
    return { url: file.getUrl(), fileId: file.getId() };
  } catch (e) {
    console.error('uploadSlip failed:', e);
    return { url: 'อัปโหลดไม่สำเร็จ: ' + e.message, fileId: '' };
  }
}

function joinChoices(arr, otherText) {
  const list = Array.isArray(arr) ? arr.filter(x => x) : [];
  let result = list.join(', ');
  if (otherText && otherText.trim()) {
    result = result ? result + '; อื่นๆ: ' + otherText.trim() : 'อื่นๆ: ' + otherText.trim();
  }
  return result;
}

/**
 * แปลง companion.list เป็น text สำหรับเก็บใน sheet
 * เช่น "1) นาย ก. (สามี, 35) | 2) ด.ญ. ข. (บุตรสาว, 8)"
 */
function formatCompanionList(cp) {
  if (!cp) return '';
  if (Array.isArray(cp.list) && cp.list.length > 0) {
    return cp.list.map((p, i) => {
      const name = p.fullName || (p.firstName + ' ' + p.lastName).trim();
      const parts = [p.relation || '-'];
      if (p.age) parts.push(`${p.age} ปี`);
      return `${i + 1}) ${name} (${parts.join(', ')})`;
    }).join(' | ');
  }
  return cp.details || '';
}

/**
 * รวมข้อจำกัดอาหาร/แพ้/โรค ของผู้ติดตามทุกคน
 */
function extractCompanionField(cp, listKey, otherKey) {
  if (!cp) return '';
  if (Array.isArray(cp.list) && cp.list.length > 0) {
    return cp.list.map((p, i) => {
      const items = (p[listKey] || []).filter(Boolean);
      const other = p[otherKey] || '';
      const text = [items.join(', '), other ? `อื่นๆ: ${other}` : ''].filter(Boolean).join('; ');
      return `${i + 1}) ${text || '-'}`;
    }).join(' | ');
  }
  // fallback
  if (listKey === 'dietRestrictions') return cp.dietNotes || '';
  if (listKey === 'allergies') return cp.allergyNotes || '';
  if (listKey === 'medicalConditions') return cp.medicalNotes || '';
  return '';
}

function buildRow(p, regId, ts, cost, slipUrl, ocrAmount, ocrStatus, editCount, lastEditTime) {
  const rm = p.roommate || {};
  const cp = p.companion || {};
  // สร้าง map ของทุก field ที่อาจมี — แต่ละ tab จะดึงเฉพาะ HEADERS ของตัวเอง
  const map = {
    'Registration ID': regId,
    'วันที่ลงทะเบียน': ts,
    'กลุ่ม': labelGroup(p.group),
    'คำนำหน้า': p.title || '',
    'ชื่อ': p.firstName || '',
    'นามสกุล': p.lastName || '',
    'ชื่อ-นามสกุล': p.fullName || '',
    'ชื่อเล่น': p.nickname || '',
    'รหัสนักศึกษา': p.studentId || '',
    'เลขบัตรประชาชน': p.nationalId || '',
    'CMMBA รุ่นที่': p.alumniBatch || '',
    'ตำแหน่ง/สังกัด': p.position || '',
    'เพศ': p.gender || '',
    'วันเกิด': p.birthDate || '',
    'เบอร์โทร': p.phone || '',
    'อีเมล': p.email || '',
    'ข้อจำกัดอาหาร': joinChoices(p.dietRestrictions, p.dietOther),
    'แพ้อาหาร': joinChoices(p.allergies, p.allergyOther),
    'โรคประจำตัว': joinChoices(p.medicalConditions, p.medicalOther),
    'ยาประจำตัว': joinChoices(p.medications, p.medicationOther),
    'ประเภทห้องพัก': labelRoomType(p.roomType),
    'ประเภทเตียง': labelBedType(p.bedType),
    'รูมเมท - ชื่อ': rm.firstName || '',
    'รูมเมท - นามสกุล': rm.lastName || '',
    'รูมเมท - ชื่อ-นามสกุล': rm.fullName || '',
    'รูมเมท - ชื่อเล่น': rm.nickname || '',
    'รูมเมท - รหัส/รุ่น': rm.idOrBatch || '',
    'รูมเมท - เพศ': rm.gender || '',
    'รูมเมท - เบอร์': rm.phone || '',
    'รูมเมท - อีเมล': rm.email || '',
    'รูมเมท - ข้อจำกัดอาหาร': joinChoices(rm.dietRestrictions, rm.dietOther),
    'รูมเมท - แพ้อาหาร': joinChoices(rm.allergies, rm.allergyOther),
    'รูมเมท - โรคประจำตัว': joinChoices(rm.medicalConditions, rm.medicalOther),
    'รูมเมท - ยา': joinChoices(rm.medications, rm.medicationOther),
    'มีผู้ติดตาม': p.hasCompanion ? 'มี' : 'ไม่มี',
    'ผู้ติดตาม - ผู้ใหญ่': cp.adultCount || 0,
    'ผู้ติดตาม - เด็ก ≤12': cp.childCount || 0,
    'ผู้ติดตาม - รายละเอียด': formatCompanionList(cp),
    'ผู้ติดตาม - ข้อจำกัดอาหาร': extractCompanionField(cp, 'dietRestrictions', 'dietOther'),
    'ผู้ติดตาม - แพ้อาหาร': extractCompanionField(cp, 'allergies', 'allergyOther'),
    'ผู้ติดตาม - โรค/ยา': extractCompanionField(cp, 'medicalConditions', 'medicalOther'),
    'ผู้ติดต่อฉุกเฉิน - ชื่อ': p.emergencyName || '',
    'ผู้ติดต่อฉุกเฉิน - ความสัมพันธ์': p.emergencyRelation || '',
    'ผู้ติดต่อฉุกเฉิน - เบอร์': p.emergencyPhone || '',
    'ยินยอมใช้ภาพ': p.consentPhoto ? 'ยินยอม' : 'ไม่ยินยอม',
    'ยินยอม PDPA (เลขบัตร)': p.consentPdpa ? 'ยินยอม' : '-',
    'หมายเหตุ': p.notes || '',
    'ค่าใช้จ่ายที่คำนวณ (บาท)': cost,
    'สลิปโอน': slipUrl || '',
    'QR Ref': p.slipQrTransRef || '',
    'ยอดโอนจริง OCR (บาท)': ocrAmount || 0,
    'สถานะ OCR': ocrStatus || '',
    'สถานะ': cost > 0 ? (slipUrl ? 'รอตรวจสลิป' : 'รอโอนเงิน') : 'ยืนยันแล้ว',
    'จำนวนครั้งที่แก้ไข': editCount || 0,
    'เวลาแก้ไขล่าสุด': lastEditTime || '',
  };

  // ดึงเฉพาะ headers ของกลุ่มนี้
  const headers = HEADERS_BY_GROUP[p.group] || HEADERS;
  return headers.map(h => map[h] !== undefined ? map[h] : '');
}

function labelGroup(g) {
  return ({
    student34: 'น้อง CMMBA รุ่น 34',
    senior33:  'พี่ CMMBA รุ่น 33',
    alumni:    'ศิษย์เก่า CMMBA รุ่น 1–32',
    faculty:   'อาจารย์และบุคลากร',
  })[g] || g || '';
}

function labelRoomType(r) {
  return ({
    twin:   'พักคู่ (มาตรฐาน)',
    single: 'พักเดี่ยว/มีผู้ติดตาม (+1,000)',
    nostay: 'ไม่ค้างคืน',
  })[r] || r || '';
}

function labelBedType(b) {
  return ({ king: 'เตียงใหญ่', twin: 'เตียงแยก' })[b] || b || '';
}

function validatePayload(p) {
  if (!p) return { ok: false, message: 'ไม่ได้รับข้อมูล' };
  if (!p.group) return { ok: false, message: 'กรุณาเลือกกลุ่ม' };
  if (!p.fullName) return { ok: false, message: 'กรุณากรอกชื่อ-นามสกุล' };
  if (!p.phone) return { ok: false, message: 'กรุณากรอกเบอร์โทรศัพท์' };
  if (!/^0\d{9}$/.test(String(p.phone).replace(/[\s-]/g, ''))) {
    return { ok: false, message: 'รูปแบบเบอร์โทรไม่ถูกต้อง' };
  }
  if (p.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p.email)) {
    return { ok: false, message: 'รูปแบบอีเมลไม่ถูกต้อง' };
  }
  return { ok: true };
}

// ====== SETUP ======

function setupSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  // ลบ tab เก่าทั้งหมดที่ไม่ต้องการ
  ['รูมเมท', 'ผู้ติดตาม', 'ลงทะเบียน'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });
  ensureSheetReady(ss);
  console.log('✅ Setup เสร็จ — สร้าง 4 tabs: ' + ALL_REG_TABS.join(', ') + ' + ' + DON_TAB);
}

function clearQRCache() {
  CacheService.getScriptCache().remove('qr_image');
  CacheService.getScriptCache().remove('qr_file_id');
  CacheService.getScriptCache().remove('assets_v1');
  CacheService.getScriptCache().remove('assets_v2');
  CacheService.getScriptCache().remove('assets_v3');
  CacheService.getScriptCache().remove('assets_v4');
  CacheService.getScriptCache().remove('assets_v5');
  CacheService.getScriptCache().remove('logo_aloha_b64');
  console.log('✅ ล้าง cache รูปภาพแล้ว');
}

/**
 * Debug — รันเพื่อดูว่า getQRImage หาเจอไหม
 * Apps Script Editor → เลือก function "debugQR" → Run
 */
function debugQR() {
  console.log('=== DEBUG QR ===');
  console.log('FOLDER_ID: ' + CONFIG.FOLDER_ID);

  // ล้าง cache ก่อน
  CacheService.getScriptCache().remove('qr_image');
  console.log('✓ Cleared cache');

  try {
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    console.log('✓ เปิด folder ได้: ' + folder.getName());

    const files = folder.getFiles();
    let count = 0;
    while (files.hasNext()) {
      const f = files.next();
      count++;
      console.log(`  [${count}] ${f.getName()} | mime: ${f.getMimeType()} | size: ${f.getSize()} bytes`);
    }
    if (count === 0) console.log('⚠️ ไม่มีไฟล์ใน folder');

    console.log('--- เรียก getQRImage() ---');
    const result = getQRImage();
    if (result) {
      console.log('✅ ได้ผลลัพธ์ ความยาว: ' + result.length + ' chars');
      console.log('Prefix: ' + result.substring(0, 80));
    } else {
      console.log('❌ getQRImage คืน empty string');
    }
  } catch (e) {
    console.error('❌ Error: ' + e.message);
    console.error(e.stack);
  }
}

/**
 * ทดสอบ OCR ด้วย file ID เฉพาะ (สำหรับ debug)
 */
function testOcr(fileId) {
  const result = extractSlipAmount(fileId);
  console.log('OCR result:', JSON.stringify(result));
  return result;
}