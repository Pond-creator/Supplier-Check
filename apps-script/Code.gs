/**
 * Supplier Match Checker — ฝั่งเจ้าของข้อมูล (Google Apps Script)
 * --------------------------------------------------------------
 * วางโค้ดนี้ใน Google Sheet ของเรา: เมนู Extensions > Apps Script
 * แล้ว Deploy เป็น Web app (Execute as: Me, Who has access: Anyone)
 *
 * หลักความลับ:
 *  - หน้าเว็บของผู้ตรวจ "ไม่เคยได้รับ" รายชื่อในชีทของเรา
 *    สคริปต์นี้ตอบกลับแค่ว่า บรรทัดที่เขาส่งมา ตรง / ใกล้เคียง / ไม่พบ เท่านั้น
 *  - สคริปต์ "ไม่บันทึก" ข้อมูลที่ผู้ตรวจส่งมา (ไม่เขียนลงชีท ไม่ log)
 *  - ต้องใส่รหัสเข้าใช้ (ACCESS_CODE ใน Script Properties) ถึงจะตรวจได้
 *    ใส่รหัสผิดเกิน 10 ครั้ง ระบบล็อก 15 นาที — กันคนนอกสุ่มเดาชื่อ
 */

// ===== ตั้งค่า =====
const SHEET_NAME = '';        // ชื่อแท็บชีท ('' = ใช้แท็บแรก)
const HEADER_ROWS = 1;        // จำนวนแถวหัวตาราง
const COL_TH = 2;             // คอลัมน์ชื่อซัพภาษาไทย (B)
const COL_EN = 3;             // คอลัมน์ชื่อซัพภาษาอังกฤษ (C)
const MATCH_SCORE = 0.90;     // คะแนน >= นี้ = "ตรงกัน" (สีเขียว)
const NEAR_SCORE = 0.80;      // คะแนน >= นี้ = "ใกล้เคียง" (สีเหลือง)
const MAX_LINES = 500;        // จำนวนบรรทัดสูงสุดต่อการส่ง 1 ครั้ง
const MAX_LINE_LEN = 200;     // ความยาวสูงสุดต่อบรรทัด (ตัวอักษร)
const MAX_FAILS = 10;         // ใส่รหัสผิดได้กี่ครั้ง ก่อนล็อกระบบชั่วคราว
const LOCK_SECONDS = 900;     // ล็อกนานกี่วินาที (900 = 15 นาที)

// ===== รหัสเข้าใช้ =====
// รหัสเก็บใน Google เท่านั้น (ไม่อยู่ในโค้ด ไม่ขึ้น GitHub)
// ตั้งที่: ⚙️ การตั้งค่าโปรเจ็กต์ > พร็อพเพอร์ตีของสคริปต์ > เพิ่ม
//   พร็อพเพอร์ตี = ACCESS_CODE   ค่า = รหัสที่จะให้ผู้ตรวจ
// ถ้ายังไม่ตั้งรหัส ระบบจะปฏิเสธทุกคำขอ (ปลอดภัยไว้ก่อน)
function checkAccess_(code) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) {
    return 'ใส่รหัสผิดหลายครั้ง ระบบล็อกชั่วคราว 15 นาที';
  }
  const real = PropertiesService.getScriptProperties().getProperty('ACCESS_CODE');
  if (!real) return 'ระบบยังไม่ได้ตั้งรหัสเข้าใช้ — โปรดติดต่อผู้ดูแล';
  if (String(code || '').trim() !== String(real).trim()) {
    cache.put('fails', String(fails + 1), LOCK_SECONDS);
    return 'รหัสเข้าใช้ไม่ถูกต้อง';
  }
  return '';
}

// ===== Web app =====
function doGet() {
  return json_({ ok: true, service: 'supplier-match-checker' });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const denied = checkAccess_(body.code);
    if (denied) return json_({ ok: false, auth: true, error: denied });
    const lines = Array.isArray(body.lines) ? body.lines.slice(0, MAX_LINES) : [];
    const index = buildIndex_(readNames_());
    const results = lines.map(function (line) {
      return matchOne_(String(line || '').slice(0, MAX_LINE_LEN), index);
    });
    return json_({ ok: true, results: results });
  } catch (err) {
    return json_({ ok: false, error: 'ตรวจสอบไม่สำเร็จ กรุณาลองใหม่' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readNames_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
  const last = sh.getLastRow();
  if (last <= HEADER_ROWS) return [];
  const width = Math.max(COL_TH, COL_EN);
  const values = sh.getRange(HEADER_ROWS + 1, 1, last - HEADER_ROWS, width).getValues();
  const names = [];
  values.forEach(function (r) {
    if (r[COL_TH - 1]) names.push(String(r[COL_TH - 1]));
    if (r[COL_EN - 1]) names.push(String(r[COL_EN - 1]));
  });
  return names;
}

// ===== ตรรกะการจับคู่ (ไม่ขึ้นกับ Google — ทดสอบนอก Apps Script ได้) =====

// คำที่ไม่ใช่ "ชื่อ" จริง ๆ ตัดทิ้งก่อนเทียบ
const TH_NOISE = ['ห้างหุ้นส่วนจำกัด', 'ห้างหุ้นส่วนสามัญ', 'ห้างหุ้นส่วน', 'บริษัทฯ', 'บริษัท', 'ฯ',
  'จำกัด', 'มหาชน', 'บจก', 'หจก', 'บมจ', 'ร้าน', 'สำนักงานใหญ่', 'สาขา'];
const EN_NOISE = ['co', 'company', 'ltd', 'limited', 'inc', 'incorporated', 'corp', 'corporation',
  'plc', 'llc', 'part', 'partnership', 'ltdpart', 'the', 'pcl', 'public', 'hq', 'head', 'office', 'thailand', 'th'];

/** ทำให้ชื่อเป็นรูปมาตรฐาน: ตัวเล็ก ตัดวรรค/สัญลักษณ์ ตัดคำว่า บริษัท/จำกัด/Co.,Ltd. ฯลฯ */
function normalize_(s) {
  let t = String(s).normalize('NFC').toLowerCase();
  t = t.replace(/[​ ]/g, ' ');
  t = t.replace(/(^|[^฀-๿])บ\s*\.\s*/g, '$1 ');   // "บ." = บริษัท (ย่อ)
  TH_NOISE.forEach(function (w) { t = t.split(w).join(' '); });
  t = t.replace(/[^a-z0-9฀-๿]+/g, ' ');
  const tokens = t.split(' ').filter(function (w) { return w && EN_NOISE.indexOf(w) === -1; });
  return tokens.join('');
}

function buildIndex_(names) {
  const exact = {};
  const list = [];
  names.forEach(function (n) {
    const k = normalize_(n);
    if (k && !exact[k]) { exact[k] = true; list.push(k); }
  });
  return { exact: exact, list: list };
}

/** ระยะแก้ไข (Levenshtein) แบบหยุดก่อนถ้าเกิน max */
function editDistance_(a, b, max) {
  const la = a.length, lb = b.length;
  if (Math.abs(la - lb) > max) return max + 1;
  let prev = new Array(lb + 1), cur = new Array(lb + 1);
  for (let j = 0; j <= lb; j++) prev[j] = j;
  for (let i = 1; i <= la; i++) {
    cur[0] = i;
    let rowMin = cur[0];
    const ca = a.charCodeAt(i - 1);
    for (let j = 1; j <= lb; j++) {
      const cost = ca === b.charCodeAt(j - 1) ? 0 : 1;
      let v = prev[j - 1] + cost;
      if (prev[j] + 1 < v) v = prev[j] + 1;
      if (cur[j - 1] + 1 < v) v = cur[j - 1] + 1;
      cur[j] = v;
      if (v < rowMin) rowMin = v;
    }
    if (rowMin > max) return max + 1;
    const tmp = prev; prev = cur; cur = tmp;
  }
  return prev[lb];
}

function matchOne_(line, index) {
  const key = normalize_(line);
  if (!key) return { status: 'empty', score: 0 };
  if (index.exact[key]) return { status: 'match', score: 1 };

  let best = 0;
  for (let i = 0; i < index.list.length; i++) {
    const cand = index.list[i];
    const maxLen = Math.max(key.length, cand.length);
    const allowed = Math.floor(maxLen * (1 - NEAR_SCORE));
    const d = editDistance_(key, cand, allowed);
    if (d <= allowed) {
      const s = 1 - d / maxLen;
      if (s > best) best = s;
    }
    // ชื่อหนึ่งอยู่ในอีกชื่อ เช่น "เอบีซี เทรดดิ้ง" vs "เอบีซี เทรดดิ้ง (ประเทศไทย)"
    const shorter = key.length < cand.length ? key : cand;
    const longer = key.length < cand.length ? cand : key;
    if (shorter.length >= 5 && longer.indexOf(shorter) !== -1) {
      // ชื่อไม่ครบ = "ใกล้เคียง" เสมอ (0.80–0.89) ให้คนตรวจดูอีกที
      const s = 0.8 + 0.09 * (shorter.length / longer.length);
      if (s > best) best = s;
    }
    if (best >= 0.999) break;
  }
  best = Math.round(best * 100) / 100;
  const status = best >= MATCH_SCORE ? 'match' : best >= NEAR_SCORE ? 'near' : 'none';
  return { status: status, score: best };
}

// ===== ทดสอบใน Apps Script editor: เลือก testMatch แล้วกด Run =====
function testMatch() {
  const index = buildIndex_(readNames_());
  ['ตัวอย่าง ชื่อซัพ', 'Sample Supplier Co., Ltd.'].forEach(function (l) {
    Logger.log(l + ' => ' + JSON.stringify(matchOne_(l, index)));
  });
  Logger.log('จำนวนชื่อในดัชนี: ' + index.list.length);
}
