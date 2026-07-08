/*************************************************************
 * ระบบจองสนามฟุตบอล Phitsanulok United — หลังบ้าน (เวอร์ชันเลือกช่องว่าง)
 * ไฟล์ Google Sheet แยกใหม่ ชื่อชีต "FieldBooking" 18 คอลัมน์
 *
 * ติดตั้ง: วางโค้ด > ใส่ค่า 3 ตัวด้านล่าง > รัน setupSheet() ครั้งเดียว
 *         > Deploy Web app (Execute as: Me, Who has access: Anyone)
 *         > เอา URL ไปใส่ GAS_URL ใน field.html และ field-staff.html
 * แก้โค้ดทุกครั้งต้อง Deploy เวอร์ชันใหม่
 *************************************************************/

// ====== ค่าที่ต้องตั้ง ======
var SHEET_ID       = '175bu0IaRFDBW-JiUVbn5pf3dX32foma0jcLoUXG2tzo';
var SHEET_NAME     = 'FieldBooking';
var STAFF_PASSWORD = 'PU@2521';
var ADMIN_EMAIL    = 'ann.apichaya@gmail.com';

// ตารางราคา/มัดจำ (ระบบยึดค่านี้เป็นหลัก กันราคาปลอมจากฝั่งลูกค้า)
var FIELD = {
  '1': { name: 'สนาม 1 (สนามใหญ่)', first: 1200, add: 1000, dep: 400 },
  '2': { name: 'สนาม 2 (สนามเล็ก)', first: 700,  add: 600,  dep: 200 },
  '3': { name: 'สนาม 3 (สนามเล็ก)', first: 700,  add: 600,  dep: 200 }
};
var OPEN_H = 8, CLOSE_H = 24, MAX_HOURS = 6;
var ACTIVE = { 'รอยืนยัน': 1, 'ยืนยันแล้ว': 1 }; // สถานะที่ถือว่า "ยึดช่องไว้แล้ว"

var HEADERS = [
  'รหัสจอง','เวลาที่ส่งคำขอ','LINE User ID','ชื่อ LINE','สนาม',
  'วันที่จอง','เวลาเริ่ม','เวลาจบ','จำนวนชั่วโมง','ค่าเช่าสนาม',
  'มัดจำ','คงเหลือ','ชื่อผู้จอง','ชื่อทีม','เบอร์โทร',
  'หมายเหตุลูกค้า','สถานะ','บันทึกเจ้าหน้าที่'
];

function setupSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#3C3489').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.getRange(1, 15, sh.getMaxRows(), 1).setNumberFormat('@'); // เบอร์โทร
  sh.getRange(1, 6, sh.getMaxRows(), 3).setNumberFormat('@');  // วันที่/เวลาเริ่ม/เวลาจบ
  SpreadsheetApp.flush();
}

function doGet(e) {
  var cb = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : 'callback';
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  var out;
  try {
    if      (action === 'availability') out = getAvailability(e.parameter);
    else if (action === 'submit')       out = submitBooking(e.parameter);
    else if (action === 'list')         out = listBookings(e.parameter);
    else if (action === 'update')       out = updateBooking(e.parameter);
    else if (action === 'login')        out = checkLogin(e.parameter);
    else                                out = { ok: false, error: 'ไม่รู้จักคำสั่ง' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ====== ดึงชั่วโมงที่ถูกจองไปแล้ว (สำหรับระบายสีแดง) ======
function getAvailability(p) {
  var f = FIELD[String(p.field)];
  if (!f) return { ok: false, error: 'ไม่พบสนามนี้' };
  var date = String(p.date || '');
  var booked = occupiedHours(f.name, date);
  return { ok: true, booked: booked };
}

// คืน array ของชั่วโมงเริ่มที่ถูกยึด เช่น [10,11] = 10:00-11:00 และ 11:00-12:00 เต็ม
function occupiedHours(fieldName, date, skipRow) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var last = sh.getLastRow();
  var occ = [];
  if (last < 2) return occ;
  var v = sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  for (var i = 0; i < v.length; i++) {
    var r = v[i];
    if (skipRow && (i + 2) === skipRow) continue;
    if (r[4] !== fieldName) continue;   // สนาม
    if (r[5] !== date) continue;        // วันที่
    if (!ACTIVE[r[16]]) continue;       // สถานะ active เท่านั้น
    var s = parseHour(r[6]), en = parseHour(r[7]);
    for (var h = s; h < en; h++) occ.push(h);
  }
  return occ;
}

function parseHour(t) {
  var m = String(t).match(/(\d+):/);
  return m ? parseInt(m[1], 10) : -1;
}

// ====== ลูกค้าส่งคำขอจอง ======
function submitBooking(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var f = FIELD[String(p.field)];
    if (!f) return { ok: false, error: 'ไม่พบสนามนี้' };

    var date = String(p.date || '').trim();
    var startH = parseInt(p.startHour, 10);
    var hours = parseInt(p.hours, 10) || 1;
    if (!date) return { ok: false, error: 'ไม่ได้เลือกวันที่' };
    if (isNaN(startH) || startH < OPEN_H) return { ok: false, error: 'เวลาเริ่มไม่ถูกต้อง' };
    if (hours < 1 || hours > MAX_HOURS)   return { ok: false, error: 'จำนวนชั่วโมงไม่ถูกต้อง' };
    var endH = startH + hours;
    if (endH > CLOSE_H) return { ok: false, error: 'ช่วงเวลาเกินเวลาปิดสนาม' };

    // เช็คซ้ำว่าช่องยังว่าง เผื่อมีคนจองตัดหน้าในจังหวะเดียวกัน
    var occ = occupiedHours(f.name, date);
    for (var h = startH; h < endH; h++) {
      if (occ.indexOf(h) > -1) return { ok: false, error: 'ช่วงเวลานี้เพิ่งถูกจองไปแล้ว กรุณาเลือกใหม่' };
    }

    var rent = f.first + (hours - 1) * f.add;
    var deposit = f.dep * hours;
    var remaining = rent - deposit;
    var phone = String(p.phone || '').trim();

    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    var bookingId = nextBookingId(sh);
    var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'd/M/yyyy HH:mm');

    var row = [
      bookingId, now, String(p.userId || ''), String(p.lineName || ''), f.name,
      "'" + date, "'" + fmt(startH), "'" + fmt(endH), hours, rent,
      deposit, remaining, String(p.guestName || '').trim(), String(p.teamName || '').trim(),
      "'" + phone, String(p.note || '').trim(), 'รอยืนยัน', ''
    ];
    sh.appendRow(row);

    if (ADMIN_EMAIL) {
      try {
        MailApp.sendEmail(ADMIN_EMAIL, 'คำขอจองสนามใหม่ ' + bookingId,
          bookingId + '\n' + f.name + '\n' + date + ' ' + fmt(startH) + '-' + fmt(endH) +
          ' (' + hours + ' ชม.)\nค่าเช่า ' + rent + ' มัดจำ ' + deposit + ' คงเหลือ ' + remaining +
          '\n' + (p.guestName || '') + (p.teamName ? ' (' + p.teamName + ')' : '') + ' ' + phone);
      } catch (mailErr) {}
    }

    var summary = 'คำขอจองสนาม ' + bookingId + '\n' + f.name + '\n' +
      date + ' ' + fmt(startH) + '-' + fmt(endH) + ' (' + hours + ' ชม.)\n' +
      'ค่าเช่า ' + rent + ' บาท · มัดจำ ' + deposit + ' บาท · คงเหลือ ' + remaining + ' บาท\n' +
      (p.guestName || '') + (p.teamName ? ' (' + p.teamName + ')' : '') + ' ' + phone;

    return { ok: true, bookingId: bookingId, summary: summary };
  } finally {
    lock.releaseLock();
  }
}

function listBookings(p) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, rows: [] };
  var v = sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  var rows = [];
  for (var i = 0; i < v.length; i++) {
    var r = v[i];
    rows.push({
      bookingId: r[0], sentAt: r[1], field: r[4], date: r[5], start: r[6], end: r[7],
      hours: r[8], rent: r[9], deposit: r[10], remaining: r[11],
      guestName: r[12], teamName: r[13], phone: r[14], note: r[15],
      status: r[16], staffNote: r[17]
    });
  }
  var filter = String(p.status || '').trim();
  if (filter && filter !== 'ทั้งหมด') rows = rows.filter(function (x) { return x.status === filter; });
  rows.reverse();
  return { ok: true, rows: rows };
}

function updateBooking(p) {
  if (String(p.pw || '') !== STAFF_PASSWORD) return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    var targetRow = -1;
    for (var i = 1; i < ids.length; i++) {
      if (String(ids[i][0]) === String(p.bookingId)) { targetRow = i + 1; break; }
    }
    if (targetRow === -1) return { ok: false, error: 'ไม่พบรหัสจองนี้' };

    var status = String(p.status || '').trim();
    var staff = String(p.staff || '').trim();
    var note = String(p.note || '').trim();
    var stamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'd/M HH:mm');
    var combined = staff;
    if (note) combined += ' · ' + note;
    combined += ' · ' + stamp;

    sh.getRange(targetRow, 17).setValue(status);
    sh.getRange(targetRow, 18).setValue(combined);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function checkLogin(p) { return { ok: String(p.pw || '') === STAFF_PASSWORD }; }

function nextBookingId(sh) {
  var last = sh.getLastRow(), max = 0;
  if (last >= 2) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      var m = String(ids[i][0]).match(/FB-(\d+)/);
      if (m) { var n = parseInt(m[1], 10); if (n > max) max = n; }
    }
  }
  return 'FB-' + ('0000' + (max + 1)).slice(-4);
}

function fmt(h) { return ('0' + h).slice(-2) + ':00'; }
