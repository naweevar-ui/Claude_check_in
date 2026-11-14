/**
 * Backend สำหรับเว็บเช็คชื่อคลาส Claude AI — เวอร์ชัน 1.3.2
 * วางไฟล์นี้ใน Google Sheets > ส่วนขยาย > Apps Script แล้ว Deploy เป็น Web app (ดู README.md)
 *
 * สร้างชีตให้อัตโนมัติ 3 แผ่น: รายชื่อ / เช็คอิน / ตั้งค่า
 */

// รหัสผู้ดูแล (ใช้เปิดเมนูนำเข้า แก้ไขชื่อ ส่งออก ตั้งค่า) — เว้นว่าง = ไม่ล็อก
const ADMIN_PIN = '';

// true = บันทึกรูปลายเซ็นเป็นไฟล์ใน Google Drive แล้วแสดงเป็นรูปในชีต "เช็คอิน"
// (รูปจะถูกแชร์แบบ "ทุกคนที่มีลิงก์ดูได้" เพื่อให้ =IMAGE() แสดงผลได้)
const SAVE_SIGNATURE_IMAGES = false;
const SIGNATURE_FOLDER = 'ลายเซ็นเช็คอิน Claude AI';

const PEOPLE_HEAD = ['id', 'ลำดับ', 'ชื่อ-นามสกุล', 'ชื่อเล่น', 'รหัส', 'บริษัท/หน่วยงาน', 'อีเมล', 'เบอร์โทร'];
const CHECK_HEAD = ['key', 'วันที่', 'id', 'ชื่อ-นามสกุล', 'บริษัท/หน่วยงาน', 'เวลาเช็คอิน', 'ตรงเวลา/สาย',
  'ขอแก้ชื่อเป็น', 'timestamp', 'ลายเซ็น (data URL)', 'รูปลายเซ็น'];

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(head);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#d3e8ff');
  }
  return s;
}
const people_ = () => sheet_('รายชื่อ', PEOPLE_HEAD);
const checks_ = () => sheet_('เช็คอิน', CHECK_HEAD);
const settings_ = () => sheet_('ตั้งค่า', ['key', 'value']);

function rows_(s) { const v = s.getDataRange().getValues(); v.shift(); return v.filter(r => r[0] !== ''); }
// บังคับให้วันที่ เวลา เบอร์โทร เก็บเป็นข้อความ ไม่ให้ Sheets แปลงเอง
const txt_ = v => (typeof v === 'string' && v !== '' && /^[\d+\-:. ()]+$/.test(v)) ? "'" + v : (v == null ? '' : v);
const str_ = v => v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v);
const time_ = v => v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm') : String(v);
const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

const VERSION = '1.3.2';

function doGet() {
  const out = { ok: true, version: VERSION, settings: {} };
  out.people = rows_(people_()).map(r => ({
    id: String(r[0]), order: Number(r[1]) || 0, name: String(r[2]), nick: String(r[3]),
    code: String(r[4]), org: String(r[5]), email: String(r[6]), phone: String(r[7])
  }));
  out.checkins = rows_(checks_()).map(r => ({
    date: str_(r[1]), pid: String(r[2]), name: String(r[3]), fix: String(r[7] || ''),
    at: Number(r[8]), sig: String(r[9] || '')
  }));
  rows_(settings_()).forEach(r => {
    const k = String(r[0]);
    out.settings[k] = k === 'requireSig' ? (r[1] === true || String(r[1]).toLowerCase() === 'true')
      : k === 'start' ? time_(r[1]) : String(r[1]);
  });
  return json_(out);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    const isAdmin = !ADMIN_PIN || String(b.pin || '') === ADMIN_PIN;
    const adminOnly = ['auth', 'putPeople', 'delPerson', 'clearPeople', 'uncheck', 'settings'];
    if (adminOnly.indexOf(b.action) >= 0 && !isAdmin) return json_({ ok: false, error: 'pin' });

    switch (b.action) {
      case 'auth': break;
      case 'checkin': {
        const c = b.checkin;
        upsert_(checks_(), c.date + '__' + c.pid, [c.date + '__' + c.pid, c.date, c.pid, c.name, c.org || '',
          c.time || '', c.status || '', c.fix || '', c.at, String(c.sig || '').slice(0, 49000), sigImage_(c)]);
        break;
      }
      case 'uncheck': remove_(checks_(), b.date + '__' + b.pid); break;
      case 'seed': if (rows_(people_()).length === 0) putPeople_(b.people || []); break;
      case 'putPeople': putPeople_(b.people || []); break;
      case 'delPerson': remove_(people_(), b.id); break;
      case 'clearPeople': { const s = people_(); if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1); break; }
      case 'settings': {
        const s = settings_(), v = b.settings || {};
        if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
        Object.keys(v).forEach(k => s.appendRow([k, txt_(String(v[k]))]));
        break;
      }
      default: return json_({ ok: false, error: 'unknown action' });
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function sigImage_(c) {
  if (!SAVE_SIGNATURE_IMAGES || !c.sig) return '';
  try {
    const m = String(c.sig).match(/^data:(image\/[a-z]+);base64,(.+)$/);
    if (!m) return '';
    const it = DriveApp.getFoldersByName(SIGNATURE_FOLDER);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(SIGNATURE_FOLDER);
    const ext = m[1] === 'image/png' ? 'png' : 'jpg';
    const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], c.date + '_' + c.name + '.' + ext));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return '=IMAGE("https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w400")';
  } catch (err) {
    return '';
  }
}

function putPeople_(list) {
  const s = people_();
  const ids = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 1).getValues().map(r => String(r[0]));
  const add = [];
  list.forEach(p => {
    const row = [p.id, p.order || 0, p.name || '', p.nick || '', txt_(p.code || ''), p.org || '', p.email || '', txt_(p.phone || '')];
    const i = ids.indexOf(String(p.id));
    if (i > 0) s.getRange(i + 1, 1, 1, row.length).setValues([row]);
    else add.push(row);
  });
  if (add.length) s.getRange(s.getLastRow() + 1, 1, add.length, add[0].length).setValues(add);
}

function upsert_(s, key, row) {
  row = row.map(txt_);
  const keys = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 1).getValues().map(r => String(r[0]));
  const i = keys.indexOf(key);
  if (i > 0) s.getRange(i + 1, 1, 1, row.length).setValues([row]);
  else s.appendRow(row);
}

function remove_(s, key) {
  const keys = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 1).getValues().map(r => String(r[0]));
  const i = keys.indexOf(String(key));
  if (i > 0) s.deleteRow(i + 1);
}
