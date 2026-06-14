/**
 * ====================================================================
 * ArrivalCheck — Google Apps Script Backend
 * ระบบเช็คชื่อ (สำหรับรถบัส / งานสัมมนา / เข้าคลาส ฯลฯ)
 * คณะวิศวกรรมศาสตร์ มหาวิทยาลัยมหาสารคาม
 * ====================================================================
 *
 * วิธีติดตั้ง:
 * 1) สร้าง Google Sheet ใหม่ ตั้งชื่อ "ArrivalCheck"
 * 2) เมนู Extensions → Apps Script → วางโค้ดนี้ทั้งหมด
 * 3) รันฟังก์ชัน setupSheets() หนึ่งครั้ง เพื่อสร้างชีต Roster + CheckIns + mock data
 * 4) Deploy → New deployment → Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    - คัดลอก URL ที่ได้ ไปใส่ใน config.js
 * 5) ถ้าแก้ไขโค้ด ต้อง Deploy → Manage deployments → Edit → New version ทุกครั้ง
 */

// ----- ค่าคงที่ -----
const SHEET_ROSTER   = 'Roster';
const SHEET_CHECKINS = 'CheckIns';

// ====================================================================
// SETUP — รันครั้งเดียวเพื่อสร้างชีตและ mock data
// ====================================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // -- Roster --
  let roster = ss.getSheetByName(SHEET_ROSTER);
  if (!roster) roster = ss.insertSheet(SHEET_ROSTER);
  roster.clear();
  roster.getRange(1, 1, 1, 4).setValues([['id', 'ชื่อ-สกุล', 'สาขาวิชา', 'เบอร์โทร']]);
  roster.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#5C0011').setFontColor('#FFFFFF');

  const mockData = [
    ['01', 'นายสมชาย ใจดี',        'เมคาทรอนิกส์',    '0815557777'],
    ['02', 'นางสาวมาลี แสนดี',     'ออโตเมชัน',       '0863334444'],
    ['03', 'นางสาวสุดา รักดี',     'EV',              '0812345678'],
    ['04', 'นายก้อง ขยัน',         'เมคาทรอนิกส์',    '0891112222'],
    ['05', 'นายธนกร พากเพียร',     'EV',              '0823456789'],
    ['06', 'นางสาวพิมพ์ใจ สดใส',   'ออโตเมชัน',       '0876543210'],
    ['07', 'นายอนุชา ตั้งใจ',      'เมคาทรอนิกส์',    '0998877665'],
    ['08', 'นางสาวกานต์ มุ่งมั่น', 'EV',              '0834567890'],
    ['09', 'นายปิติ ยินดี',        'ออโตเมชัน',       '0867891234'],
    ['10', 'นางสาวอาทิตยา ส่องแสง','เมคาทรอนิกส์',    '0845678901']
  ];
  roster.getRange(2, 1, mockData.length, 4).setValues(mockData);
  roster.getRange(2, 1, mockData.length, 1).setHorizontalAlignment('center');
  roster.getRange(2, 4, mockData.length, 1).setNumberFormat('@'); // เก็บเบอร์เป็น text
  roster.setColumnWidth(1, 60);
  roster.setColumnWidth(2, 200);
  roster.setColumnWidth(3, 150);
  roster.setColumnWidth(4, 120);
  roster.setFrozenRows(1);

  // -- CheckIns --
  let checkins = ss.getSheetByName(SHEET_CHECKINS);
  if (!checkins) checkins = ss.insertSheet(SHEET_CHECKINS);
  checkins.clear();
  checkins.getRange(1, 1, 1, 6).setValues([['timestamp', 'round', 'id', 'ชื่อ', 'สาขา', 'เบอร์']]);
  checkins.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#5C0011').setFontColor('#FFFFFF');
  checkins.setColumnWidth(1, 160);
  checkins.setColumnWidth(2, 60);
  checkins.setColumnWidth(3, 60);
  checkins.setColumnWidth(4, 200);
  checkins.setColumnWidth(5, 150);
  checkins.setColumnWidth(6, 120);
  checkins.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('สร้างชีตและ mock data เรียบร้อยแล้ว');
}

// ====================================================================
// GET — สำหรับ index.html (roster) และ dashboard.html (รายชื่อ + สถานะ)
// ====================================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'roster';

  if (action === 'roster') {
    return jsonOut({ ok: true, roster: getRosterPublic() });
  }
  if (action === 'dashboard') {
    const round = parseInt(e.parameter.round || '1', 10);
    return jsonOut({ ok: true, round: round, rows: getDashboard(round) });
  }
  return jsonOut({ ok: false, error: 'unknown action' });
}

// ====================================================================
// POST — บันทึกการเช็คชื่อ
// ====================================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const id    = String(body.id || '').trim();
    const pin4  = String(body.pin4 || '').trim();
    const round = parseInt(body.round, 10);

    if (!id || !pin4 || !round) {
      return jsonOut({ ok: false, error: 'ข้อมูลไม่ครบ' });
    }

    // หา record ใน roster
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const roster = ss.getSheetByName(SHEET_ROSTER);
    const data = roster.getRange(2, 1, roster.getLastRow() - 1, 4).getValues();
    const found = data.find(r => String(r[0]).trim() === id);
    if (!found) return jsonOut({ ok: false, error: 'ไม่พบรหัสนิสิตนี้' });

    const phone = String(found[3]).trim();
    const realPin4 = phone.slice(-4);
    if (pin4 !== realPin4) {
      return jsonOut({ ok: false, error: 'รหัส 4 หลักท้ายเบอร์โทรไม่ถูกต้อง' });
    }

    // ตรวจกดซ้ำ
    const checkins = ss.getSheetByName(SHEET_CHECKINS);
    const last = checkins.getLastRow();
    if (last >= 2) {
      const existing = checkins.getRange(2, 1, last - 1, 6).getValues();
      const dup = existing.find(r => String(r[2]).trim() === id && parseInt(r[1], 10) === round);
      if (dup) {
        const dupTime = Utilities.formatDate(new Date(dup[0]), 'Asia/Bangkok', 'HH:mm:ss');
        return jsonOut({
          ok: false,
          duplicate: true,
          error: 'คุณเช็คชื่อครั้งที่ ' + round + ' ไปแล้วเมื่อ ' + dupTime
        });
      }
    }

    // append
    const now = new Date();
    checkins.appendRow([now, round, id, found[1], found[2], phone]);
    const timeStr = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');

    return jsonOut({
      ok: true,
      message: 'บันทึกสำเร็จ',
      name: found[1],
      round: round,
      time: timeStr
    });
  } catch (err) {
    return jsonOut({ ok: false, error: 'เกิดข้อผิดพลาด: ' + err.message });
  }
}

// ====================================================================
// HELPERS
// ====================================================================
function getRosterPublic() {
  // ส่งเฉพาะ id + ชื่อ + สาขา (ไม่ส่งเบอร์ ป้องกันการเดารหัส)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_ROSTER);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 3).getValues().map(r => ({
    id:     String(r[0]).trim(),
    name:   String(r[1]).trim(),
    branch: String(r[2]).trim()
  }));
}

function getDashboard(round) {
  // dashboard ส่งเบอร์เต็ม (ใช้คนเดียว)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roster = ss.getSheetByName(SHEET_ROSTER);
  const last = roster.getLastRow();
  if (last < 2) return [];
  const rosterData = roster.getRange(2, 1, last - 1, 4).getValues();

  // map การ check-in ของรอบนี้
  const checkins = ss.getSheetByName(SHEET_CHECKINS);
  const cLast = checkins.getLastRow();
  const checkMap = {};
  if (cLast >= 2) {
    const cData = checkins.getRange(2, 1, cLast - 1, 6).getValues();
    cData.forEach(r => {
      if (parseInt(r[1], 10) === round) {
        checkMap[String(r[2]).trim()] = r[0];
      }
    });
  }

  return rosterData.map(r => {
    const id = String(r[0]).trim();
    const checkedAt = checkMap[id];
    return {
      id:        id,
      name:      String(r[1]).trim(),
      branch:    String(r[2]).trim(),
      phone:     String(r[3]).trim(),
      checked:   !!checkedAt,
      time:      checkedAt ? Utilities.formatDate(new Date(checkedAt), 'Asia/Bangkok', 'HH:mm:ss') : ''
    };
  });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
