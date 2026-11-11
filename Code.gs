/************************************************************************
 * ระบบข้อเสนอโครงการบริการวิชาการรับใช้สังคม 2570 — Apps Script Backend
 * ใช้คู่กับ Google Sheets (ผูกกับชีตนี้) + โฟลเดอร์ Google Drive สำหรับไฟล์แนบ
 * ขั้นตอน: แก้ 3 ค่าด้านล่าง -> รันฟังก์ชัน setup() หนึ่งครั้ง -> Deploy เป็น Web app
 ************************************************************************/
var ADMIN_EMAIL = 'anuyachai@gmail.com';   // บัญชีผู้ดูแลระบบหลัก
var ADMIN_NAME  = 'ผู้ดูแลระบบ';
var ADMIN_PIN   = 'CHANGE_THIS_PIN';     // ตั้ง PIN ตัวเลขใหม่อย่างน้อย 6 หลักก่อนรัน setup() ครั้งแรก

var CH = 45000, NCH = 6, TOKEN_TTL = 21600, MAX_FILE = 2000000;
var PH = [
  ['ร่าง / บันทึกฉบับร่าง', 'ส่งข้อเสนอแล้ว', 'รอตรวจสอบเอกสาร', 'รอเอกสารเพิ่มเติม', 'อยู่ระหว่างแก้ไข'],
  ['รอผลการพิจารณา', 'ผ่าน', 'ผ่านแบบมีเงื่อนไข', 'ไม่ผ่าน', 'รออนุมัติทุน', 'อนุมัติทุน'],
  ['รอทำสัญญา', 'ทำสัญญาแล้ว', 'รอเริ่มดำเนินโครงการ'],
  ['อยู่ระหว่างดำเนินโครงการ', 'เบิกเงินงวด 1', 'เบิกเงินงวด 2', 'รอเบิกเงินงวดถัดไป', 'รอรายงานความก้าวหน้า', 'รอตรวจสอบรายงาน'],
  ['รอส่งรายงานฉบับสมบูรณ์', 'รอตรวจสอบรายงานฉบับสมบูรณ์', 'รอเคลียร์ค่าใช้จ่าย', 'รอแก้ไขเอกสารปิดโครงการ'],
  ['ปิดโครงการสมบูรณ์', 'ยกเลิกโครงการ']
];
var STAT = [].concat.apply([], PH);
var ALLOW = [
  STAT,
  ['ร่าง / บันทึกฉบับร่าง', 'ส่งข้อเสนอแล้ว', 'อยู่ระหว่างแก้ไข', 'รอตรวจสอบรายงาน', 'รอตรวจสอบรายงานฉบับสมบูรณ์'],
  ['รอตรวจสอบเอกสาร', 'รอเอกสารเพิ่มเติม', 'รอผลการพิจารณา'].concat(PH[2], PH[3], PH[4], PH[5]),
  ['รอผลการพิจารณา', 'ผ่าน', 'ผ่านแบบมีเงื่อนไข', 'ไม่ผ่าน', 'รออนุมัติทุน'],
  ['อนุมัติทุน', 'ไม่ผ่าน', 'ยกเลิกโครงการ']
];
var EDIT = ['ร่าง / บันทึกฉบับร่าง', 'รอเอกสารเพิ่มเติม', 'อยู่ระหว่างแก้ไข'];
var NOTE = ['รอเอกสารเพิ่มเติม', 'อยู่ระหว่างแก้ไข', 'ผ่านแบบมีเงื่อนไข', 'ไม่ผ่าน', 'ยกเลิกโครงการ', 'รอแก้ไขเอกสารปิดโครงการ'];
var SHEETS = {
  Users: ['email', 'name', 'role', 'salt', 'hash', 'active', 'mustChange', 'hashVersion'],
  Projects: ['id', 'code', 'owner', 'status', 'title', 'dept', 'ver', 'updated', 'j1', 'j2', 'j3', 'j4', 'j5', 'j6'],
  Log: ['time', 'projectId', 'by', 'from', 'to', 'note'],
  Files: ['fileId', 'projectId', 'name', 'mime', 'size', 'by'],
  Config: ['key', 'value']
};

/* ---------- ติดตั้งครั้งแรก ---------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pr = PropertiesService.getScriptProperties();
  pr.setProperty('ss', ss.getId());
  Object.keys(SHEETS).forEach(function (n) {
    var s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (s.getLastRow() === 0) { s.appendRow(SHEETS[n]); s.setFrozenRows(1); }
    else if (n === 'Users') {
      var heads = s.getRange(1, 1, 1, Math.max(s.getLastColumn(), 1)).getValues()[0];
      ['mustChange', 'hashVersion'].forEach(function (h) { if (heads.indexOf(h) < 0) { s.getRange(1, s.getLastColumn() + 1).setValue(h); heads.push(h); } });
    }
  });
  var cfg = ss.getSheetByName('Config');
  if (cfg.getLastRow() < 2) [['cap', 50000], ['inst1', 70], ['pass', 70]].forEach(function (r) { cfg.appendRow(r); });
  if (!pr.getProperty('folder')) pr.setProperty('folder', DriveApp.createFolder('ASV70-files').getId());
  var primary = getUsers().filter(function (u) { return u.email === ADMIN_EMAIL.toLowerCase(); })[0];
  if (!primary) {
    if (!/^\d{6,}$/.test(String(ADMIN_PIN)) || ADMIN_PIN === '123456') throw new Error('กรุณากำหนด ADMIN_PIN เป็นตัวเลขอย่างน้อย 6 หลักใน Code.gs ก่อนสร้างบัญชีผู้ดูแลระบบ');
    putUser({ email: ADMIN_EMAIL, name: ADMIN_NAME, role: 0, pin: ADMIN_PIN, active: true });
  } else if (primary.role !== 0 || !primary.active) {
    putUser({ email: ADMIN_EMAIL, name: ADMIN_NAME, role: 0, active: true });
  }
  Logger.log('ติดตั้ง/ตรวจสอบบัญชีผู้ดูแลระบบแล้ว: ' + ADMIN_EMAIL);
}

/* ---------- ตัวช่วย ---------- */
var _ss;
function ss() { return _ss || (_ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ss'))); }
function sh(n) { return ss().getSheetByName(n); }
function table(n) { var v = sh(n).getDataRange().getValues(); v.shift(); return v; }
function safe(s) { s = String(s == null ? '' : s); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function hex(b) { return b.map(function (x) { x = x < 0 ? x + 256 : x; return ('0' + x.toString(16)).slice(-2); }).join(''); }
function hashPin(pin, salt) { return hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pin + '|asv70')); }
// Iterated salted SHA-256 for newly set passwords; legacy PIN hashes remain verifiable.
function hashPassword(password, salt) {
  var v = salt + '|asv70-password-v2|' + password;
  for (var i = 0; i < 10000; i++) v = hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, v));
  return v;
}
function verifyUserPassword(u, password) {
  return u.hashVersion >= 2 ? hashPassword(password, u.salt) === u.hash : hashPin(password, u.salt) === u.hash;
}
function isAp(s) { return STAT.indexOf(s) >= STAT.indexOf('อนุมัติทุน') && s !== 'ยกเลิกโครงการ'; }
function withLock(f) { var l = LockService.getScriptLock(); l.waitLock(25000); try { return f(); } finally { l.releaseLock(); } }

/* ---------- ผู้ใช้ ---------- */
function getUsers() {
  return table('Users').map(function (r, i) {
    return { row: i + 2, email: String(r[0]).toLowerCase(), name: r[1], role: +r[2], salt: r[3], hash: r[4], active: r[5] === true || r[5] === 'TRUE', mustChange: r[6] === true || r[6] === 'TRUE', hashVersion: +r[7] || 1 };
  });
}
function putUser(u) {
  var email = String(u.email).trim().toLowerCase();
  var f = getUsers().filter(function (x) { return x.email === email; })[0], salt, hash, mustChange, hashVersion;
  if (u.pin) { salt = Utilities.getUuid(); hash = hashPassword(String(u.pin), salt); hashVersion = 2; mustChange = u.mustChange !== false; }
  else if (f) { salt = f.salt; hash = f.hash; hashVersion = f.hashVersion; mustChange = f.mustChange; }
  else throw 'ต้องกำหนดรหัสผ่านสำหรับผู้ใช้ใหม่';
  var row = [email, safe(u.name), +u.role, salt, hash, !!u.active, !!mustChange, hashVersion || 2];
  if (f) sh('Users').getRange(f.row, 1, 1, 8).setValues([row]); else sh('Users').appendRow(row);
}
function listUsers() { return getUsers().map(function (u) { return { email: u.email, name: u.name, role: u.role, active: u.active }; }); }
function login(d) {
  var email = String(d.email || '').trim().toLowerCase(), pin = String(d.password != null ? d.password : (d.pin || ''));
  var c = CacheService.getScriptCache(), fk = 'f:' + email, n = +(c.get(fk) || 0);
  if (n >= 5) throw 'พยายามเข้าสู่ระบบผิดหลายครั้ง กรุณารอ 10 นาที';
  var u = getUsers().filter(function (x) { return x.email === email; })[0];
  if (!u || !u.active || !verifyUserPassword(u, pin)) { c.put(fk, String(n + 1), 600); throw 'อีเมลหรือรหัสผ่านไม่ถูกต้อง'; }
  // Upgrade an existing legacy PIN hash after a successful login.
  if (u.hashVersion < 2) { var ns = Utilities.getUuid(); sh('Users').getRange(u.row, 4, 1, 5).setValues([[ns, hashPassword(pin, ns), true, u.mustChange, 2]]); u.hashVersion = 2; }
  c.remove(fk);
  var t = Utilities.getUuid() + Utilities.getUuid();
  c.put('t:' + t, email, TOKEN_TTL);
  return { t: t, mustChange: !!u.mustChange };
}
function auth(t) {
  var email = t && CacheService.getScriptCache().get('t:' + t);
  if (!email) throw 'AUTH';
  var u = getUsers().filter(function (x) { return x.email === email && x.active; })[0];
  if (!u) throw 'AUTH';
  return u;
}

/* ---------- ตั้งค่า ---------- */
function getCfg() {
  var o = {}; table('Config').forEach(function (r) { o[r[0]] = +r[1]; });
  return { cap: o.cap || 50000, inst1: o.inst1 == null ? 70 : o.inst1, pass: o.pass || 70 };
}

/* ---------- โครงการ ---------- */
function allLogs() {
  var m = {};
  table('Log').forEach(function (r) { var k = String(r[1]); (m[k] = m[k] || []).push({ t: r[0], by: r[2], f: r[3], to: r[4], n: r[5] }); });
  return m;
}
function parseRow(r, logs) {
  var js = ''; for (var i = 8; i < 14; i++) js += r[i] || '';
  var p = js ? JSON.parse(js) : {};
  p.id = isNaN(+r[0]) ? r[0] : +r[0]; p.code = r[1]; p.owner = r[2]; p.status = r[3]; p.ver = +r[6] || 1; p.sv = 1;
  p.log = (logs && logs[String(r[0])]) || [];
  return p;
}
function packRow(p) {
  var o = JSON.parse(JSON.stringify(p));
  ['id', 'code', 'owner', 'status', 'ver', 'sv', 'log'].forEach(function (k) { delete o[k]; });
  var js = JSON.stringify(o);
  if (js.length > CH * NCH) throw 'ข้อมูลโครงการใหญ่เกินไป (ลดขนาดข้อความหรือลายเซ็น)';
  var row = [p.id, p.code, p.owner, p.status, safe(p.th), safe((p.pi || {}).dept), p.ver, Date.now()];
  for (var i = 0; i < NCH; i++) row.push(js.substr(i * CH, CH));
  return row;
}
function findProj(id) {
  var rows = table('Projects');
  for (var i = 0; i < rows.length; i++) if (String(rows[i][0]) === String(id)) return { row: i + 2, r: rows[i] };
  return null;
}
function canView(u, p) { return u.role === 0 || (u.role === 1 ? p.owner === u.email : p.status !== EDIT[0]); }
function pub(u, p) { if (u.role === 1) delete p.rev; return p; }
function withLog(p) { p.log = (allLogs()[String(p.id)]) || []; return p; }
function nextCode() {
  var pr = PropertiesService.getScriptProperties(), n = +(pr.getProperty('seq') || 0) + 1;
  pr.setProperty('seq', String(n));
  return 'ASV70-' + ('00' + n).slice(-3);
}
var KEEP = { id: 1, code: 1, owner: 1, status: 1, log: 1, rev: 1, ver: 1, sv: 1 };
function merge(cur, inc, mode) {
  var o = JSON.parse(JSON.stringify(cur));
  if (mode === 'full') { for (var k in inc) if (!KEEP[k]) o[k] = inc[k]; }
  else {
    (inc.acts || []).forEach(function (a, i) { if (o.acts && o.acts[i]) { o.acts[i].spend = a.spend; o.acts[i].files = a.files; } });
    if (mode === 'part' && inc.files) o.files = inc.files;
    if (inc.ck) o.ck = inc.ck;
    if (mode === 'staff' && inc.sroi != null) o.sroi = inc.sroi;
  }
  return o;
}
function actBoot(u) {
  var logs = allLogs(), ps = [];
  table('Projects').forEach(function (r) { var p = parseRow(r, logs); if (canView(u, p)) ps.push(pub(u, p)); });
  return { user: { email: u.email, name: u.name, role: u.role }, cfg: getCfg(), users: (u.role === 0 || u.role === 2) ? listUsers() : [], projects: ps };
}
function actSave(u, d) {
  var inc = d.p;
  if (!inc || !inc.id) throw 'ข้อมูลไม่ครบ';
  return withLock(function () {
    var f = findProj(inc.id), p;
    if (!f) {
      if (u.role !== 0 && u.role !== 1) throw 'FORBIDDEN';
      p = merge({ acts: [] }, inc, 'full');
      p.id = inc.id; p.code = nextCode(); p.owner = u.email; p.status = EDIT[0]; p.ver = 1; p.rev = {};
      sh('Projects').appendRow(packRow(p));
      return { p: pub(u, withLog(p)) };
    }
    var cur = parseRow(f.r, null), mode = '';
    if (u.role === 0) mode = 'full';
    else if (u.role === 1 && cur.owner === u.email) mode = EDIT.indexOf(cur.status) > -1 ? 'full' : (isAp(cur.status) ? 'part' : '');
    else if (u.role === 2 && isAp(cur.status)) mode = 'staff';
    if (!mode) throw 'FORBIDDEN';
    if (mode === 'full' && d.base != null && +d.base !== cur.ver) throw 'CONFLICT';
    p = merge(cur, inc, mode); p.ver = cur.ver + 1;
    sh('Projects').getRange(f.row, 1, 1, 14).setValues([packRow(p)]);
    return { p: pub(u, withLog(p)) };
  });
}
function actStatus(u, d) {
  var to = d.to, note = String(d.note || '').trim();
  if (STAT.indexOf(to) < 0) throw 'สถานะไม่ถูกต้อง';
  if (ALLOW[u.role].indexOf(to) < 0) throw 'FORBIDDEN';
  if (NOTE.indexOf(to) > -1 && !note) throw 'กรุณาระบุหมายเหตุ';
  return withLock(function () {
    var f = findProj(d.id); if (!f) throw 'ไม่พบโครงการ (กรุณากรอกชื่อโครงการและบันทึกก่อน)';
    var p = parseRow(f.r, null); if (!canView(u, p)) throw 'FORBIDDEN';
    if (u.role === 1) {
      var inEdit = EDIT.indexOf(p.status) > -1, editSide = EDIT.indexOf(to) > -1 || to === 'ส่งข้อเสนอแล้ว';
      if (p.owner !== u.email || (editSide && !inEdit) || (!editSide && !isAp(p.status))) throw 'FORBIDDEN';
    }
    if (to === 'ส่งข้อเสนอแล้ว' && !String(p.th || '').trim()) throw 'ยังไม่ได้ระบุชื่อโครงการ';
    var old = p.status; p.status = to; p.ver += 1;
    sh('Projects').getRange(f.row, 1, 1, 14).setValues([packRow(p)]);
    sh('Log').appendRow([Date.now(), p.id, safe(u.name), old, to, safe(note)]);
    return { p: pub(u, withLog(p)) };
  });
}
function actReview(u, d) {
  if ([0, 2, 3].indexOf(u.role) < 0) throw 'FORBIDDEN';
  return withLock(function () {
    var f = findProj(d.id); if (!f) throw 'ไม่พบโครงการ';
    var p = parseRow(f.r, null); if (!canView(u, p)) throw 'FORBIDDEN';
    p.rev = p.rev || {};
    p.rev[u.name] = { s: (d.s || []).slice(0, 5).map(function (x) { return Math.max(0, Math.min(20, +x || 0)); }), n: String(d.n || '').substr(0, 1000) };
    sh('Projects').getRange(f.row, 1, 1, 14).setValues([packRow(p)]);
    return { ok: 1 };
  });
}
function actUpload(u, d) {
  var f = findProj(d.id); if (!f) throw 'ไม่พบโครงการ (กรุณากรอกชื่อโครงการและบันทึกก่อนแนบไฟล์)';
  var p = parseRow(f.r, null);
  if (!(u.role === 0 || u.role === 2 || (u.role === 1 && p.owner === u.email))) throw 'FORBIDDEN';
  var bytes = Utilities.base64Decode(String(d.b || ''));
  if (bytes.length > MAX_FILE) throw 'ไฟล์ใหญ่เกินกำหนด';
  var name = String(d.n || 'file').replace(/[\\\/:*?"<>|]/g, '_');
  var file = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('folder')).createFile(Utilities.newBlob(bytes, d.t || 'application/octet-stream', name));
  sh('Files').appendRow([file.getId(), p.id, safe(name), d.t || '', bytes.length, u.email]);
  return { f: { id: file.getId(), n: name, s: bytes.length, t: d.t || '' } };
}
function actFile(u, d) {
  var fr = table('Files').filter(function (r) { return r[0] === d.id; })[0];
  if (!fr) throw 'ไม่พบไฟล์';
  var f = findProj(fr[1]); if (!f || !canView(u, parseRow(f.r, null))) throw 'FORBIDDEN';
  var blob = DriveApp.getFileById(d.id).getBlob();
  return { n: fr[2], t: fr[3] || blob.getContentType(), b: Utilities.base64Encode(blob.getBytes()) };
}
function actCfg(u, d) {
  if (u.role !== 0) throw 'FORBIDDEN';
  var vals = { cap: Math.max(1, +d.cap || 50000), inst1: Math.max(0, Math.min(100, +d.inst1 || 0)), pass: Math.max(0, Math.min(100, +d.pass || 70)) };
  return withLock(function () {
    var s = sh('Config'), rows = table('Config');
    Object.keys(vals).forEach(function (k) {
      var i = -1; rows.forEach(function (r, j) { if (r[0] === k) i = j; });
      if (i >= 0) s.getRange(i + 2, 2).setValues([[vals[k]]]); else s.appendRow([k, vals[k]]);
    });
    return { cfg: getCfg() };
  });
}
function actChangePassword(u, d) {
  var current = String(d.currentPassword || ''), next = String(d.newPassword || ''), confirm = String(d.confirmPassword || '');
  if (!verifyUserPassword(u, current)) throw 'รหัสผ่านเดิมไม่ถูกต้อง';
  if (next.length < 8) throw 'รหัสผ่านใหม่ต้องมีอย่างน้อย 8 ตัวอักษร';
  if (!/[A-Za-zก-๙]/.test(next) || !/[0-9]/.test(next)) throw 'รหัสผ่านใหม่ต้องมีทั้งตัวอักษรและตัวเลข';
  if (next !== confirm) throw 'ยืนยันรหัสผ่านใหม่ไม่ตรงกัน';
  if (next === current) throw 'รหัสผ่านใหม่ต้องไม่ซ้ำกับรหัสผ่านเดิม';
  return withLock(function () {
    var f = getUsers().filter(function (x) { return x.email === u.email; })[0];
    if (!f) throw 'ไม่พบบัญชีผู้ใช้';
    var salt = Utilities.getUuid(), hash = hashPassword(next, salt);
    sh('Users').getRange(f.row, 4, 1, 5).setValues([[salt, hash, f.active, false, 2]]);
    return { changed: true };
  });
}

function actUserSave(u, d) {
  if (u.role !== 0) throw 'FORBIDDEN';
  var email = String(d.email || '').trim().toLowerCase(), role = +d.role;
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw 'รูปแบบอีเมลไม่ถูกต้อง';
  if (!(role >= 0 && role <= 4)) throw 'บทบาทไม่ถูกต้อง';
  if (!String(d.name || '').trim()) throw 'กรุณาระบุชื่อ';
  if (d.pin && (String(d.pin).length < 8 || !/[A-Za-zก-๙]/.test(String(d.pin)) || !/[0-9]/.test(String(d.pin)))) throw 'รหัสผ่านต้องมีอย่างน้อย 8 ตัว และมีทั้งตัวอักษรกับตัวเลข';
  return withLock(function () {
    var us = getUsers().map(function (x) { return x.email === email ? { role: role, active: !!d.active } : x; });
    if (!us.some(function (x) { return x.role === 0 && x.active; })) throw 'ต้องมีผู้ดูแลระบบที่ใช้งานอยู่อย่างน้อย 1 คน';
    putUser({ email: email, name: d.name, role: role, pin: d.pin, active: !!d.active, mustChange: !!d.pin });
    return { users: listUsers() };
  });
}

/* ---------- Web app ---------- */
function doGet() { return ContentService.createTextOutput('ASV70 API พร้อมใช้งาน'); }
function doPost(e) {
  var out;
  try {
    var q = JSON.parse(e.postData.contents), a = q.a, d = q.d || {};
    if (a === 'login') out = login(d);
    else {
      var u = auth(q.t);
      if (u.mustChange && a !== 'password') throw 'PASSWORD_CHANGE_REQUIRED';
      var R = { password: actChangePassword, boot: actBoot, save: actSave, status: actStatus, review: actReview, upload: actUpload, file: actFile, cfg: actCfg, usave: actUserSave };
      if (a === 'users') { if (u.role !== 0 && u.role !== 2) throw 'FORBIDDEN'; out = { users: listUsers() }; }
      else if (R[a]) out = R[a](u, d);
      else throw 'ไม่รู้จักคำสั่ง';
    }
    out = { ok: true, d: out };
  } catch (err) {
    out = { ok: false, e: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
