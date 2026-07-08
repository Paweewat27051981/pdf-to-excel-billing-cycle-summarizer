// ============================================================
// NEOSIAM — Auto Backup to Google Drive + Sheets (ฟรี)
// สำรองข้อมูลจากแอปทุกวัน: ไฟล์ JSON เต็ม -> Google Drive (กู้กลับได้)
//                        + สรุป -> ชีตในไฟล์นี้ (ดูง่าย)
// วิธีใช้: สร้าง Google Sheet ใหม่ -> เมนู Extensions > Apps Script
//         -> วางโค้ดนี้ทั้งหมด -> เลือกฟังก์ชัน setupDailyTrigger แล้วกด Run 1 ครั้ง
//         -> กดอนุญาต (Authorize) -> เสร็จ! ระบบจะสำรองเองทุกวันตี 1
// ============================================================

const API_URL   = 'https://neosiam.dscloud.biz:8444/neosiam/api/state';  // NAS (ตัวหลัก) — Render suspend แล้ว
const FOLDER    = 'NEOSIAM-Backup';   // ชื่อโฟลเดอร์ใน Google Drive (สร้างให้อัตโนมัติ)
const LOG_SHEET = 'สำรอง-log';         // ชื่อชีตบันทึกการสำรอง
const KEEP      = 30;                  // เก็บไฟล์ล่าสุดกี่ไฟล์ (ลบที่เก่ากว่านี้ กันเปลือง Drive)
const HOUR      = 1;                   // สำรองตอนกี่โมง (0-23) — ตี 1

// ---------- ฟังก์ชันหลัก: รันทุกวันโดย trigger ----------
function backupNow() {
  const data = fetchState_();
  const now  = new Date();
  const stamp = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd_HH-mm');
  const json = JSON.stringify(data);

  // 1) เซฟไฟล์ JSON เต็มเข้า Drive (ไฟล์นี้ใช้ "กู้กลับเข้าระบบ" ได้)
  const folder = getFolder_(FOLDER);
  folder.createFile('backup-' + stamp + '.json', json, 'application/json');
  pruneOld_(folder, KEEP);

  // 2) บันทึกสรุป 1 แถวลงชีต (ไว้ดูว่าสำรองเมื่อไหร่ มีข้อมูลเท่าไร)
  const trips = data.tripDocuments || [];
  const totalTrip = trips.reduce((s, t) => s + (t.tripAmount || 0), 0);
  logRow_([
    Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm'),
    (data.cycles || []).length,
    trips.length,
    (data.vehicles || []).length,
    (data.rateMasters || []).length,
    Math.round(json.length / 1024) + ' KB',
    Math.round(totalTrip),
  ]);
}

// ---------- ดึงข้อมูลจากแอป (retry เผื่อ Render กำลังหลับ/cold-start) ----------
function fetchState_() {
  for (let i = 0; i < 5; i++) {
    try {
      const res = UrlFetchApp.fetch(API_URL, { muteHttpExceptions: true, followRedirects: true });
      if (res.getResponseCode() === 200) return JSON.parse(res.getContentText());
    } catch (e) { /* ลองใหม่ */ }
    Utilities.sleep(15000); // รอ 15 วิ ให้ Render ตื่น
  }
  throw new Error('ดึงข้อมูลไม่สำเร็จ (server อาจหลับอยู่) — ลองใหม่รอบหน้า');
}

function getFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

// ลบไฟล์เก่า เก็บแค่ KEEP ไฟล์ล่าสุด (ย้ายเข้าถังขยะ Drive)
function pruneOld_(folder, keep) {
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  files.slice(keep).forEach((f) => f.setTrashed(true));
}

function logRow_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.appendRow(['เวลาสำรอง', 'จำนวนรอบ', 'จำนวนใบกระจาย', 'จำนวนรถ', 'จำนวนราคา', 'ขนาดไฟล์', 'รวมค่าเที่ยว(บาท)']);
    sh.getRange('1:1').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.appendRow(row);
}

// ---------- ตั้งเวลาให้รันอัตโนมัติทุกวัน (กด Run ฟังก์ชันนี้ครั้งเดียวตอนติดตั้ง) ----------
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t)); // ล้าง trigger เก่า
  ScriptApp.newTrigger('backupNow').timeBased().everyDays(1).atHour(HOUR).create();
  backupNow(); // ทดสอบสำรองทันที 1 ครั้ง (จะเห็นไฟล์ใน Drive + แถวในชีต)
}
