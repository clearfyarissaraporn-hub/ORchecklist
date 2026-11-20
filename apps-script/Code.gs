/**
 * ใบเช็กลิสต์ยืมของ ห้องผ่าตัด — Backend (Google Apps Script)
 * วิธีใช้: วางไฟล์นี้ใน Apps Script ของ Google Sheet → เลือกฟังก์ชัน setup → กด Run ครั้งแรก
 * แล้ว Deploy เป็น Web app (Execute as: Me, Who has access: Anyone)
 */

// รหัสทีม: ใส่คำที่รู้กันในทีม (เช่น 'or2569') เพื่อกันคนนอกเขียนข้อมูล — ปล่อยว่าง = ไม่ใช้รหัส
const TEAM_KEY = '';

// เวอร์ชัน (GitHub Actions ใส่ให้อัตโนมัติทุกครั้งที่อัปเดต)
const VERSION = '__VERSION__';

const COLS = {
  cases:     ['id','cat','name','ins','sup'],
  doctors:   ['name'],
  staff:     ['name'],
  templates: ['key','doctor','caseId','caseName','ins','sup','updatedAt'],
  records:   ['id','savedAt','date','caseId','caseName','doctor','by','by2','borrowIns','borrowSup','returned','ins','sup','editedAt']
};
const TAB = {cases:'เคส', doctors:'แพทย์', staff:'รายชื่อ', templates:'รายการตามแพทย์', records:'บันทึก'};
const JSON_FIELDS = ['ins','sup'];

function sheet_(k){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(TAB[k]);
  if(!s){
    s = ss.insertSheet(TAB[k]);
    s.getRange(1,1,1000,COLS[k].length).setNumberFormat('@');
    s.getRange(1,1,1,COLS[k].length).setValues([COLS[k]]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function setup(){
  Object.keys(COLS).forEach(sheet_);
  const s = sheet_('cases');
  if(s.getLastRow() < 2){
    const rows = SEED.map(c => [c.id, c.cat, c.name, JSON.stringify(c.ins), JSON.stringify(c.sup)]);
    s.getRange(2,1,rows.length,5).setNumberFormat('@').setValues(rows);
  }
  const d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1') || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ชีต1');
  if(d && d.getLastRow() === 0 && SpreadsheetApp.getActiveSpreadsheet().getSheets().length > 1) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(d);
}

function rows_(k){
  const s = sheet_(k), n = s.getLastRow();
  if(n < 2) return [];
  return s.getRange(2,1,n-1,COLS[k].length).getDisplayValues().map((r,i) => {
    const o = {_row: i+2};
    COLS[k].forEach((h,j) => {
      let v = r[j];
      if(JSON_FIELDS.indexOf(h) >= 0){ try{ v = v ? JSON.parse(v) : []; }catch(e){ v = []; } }
      o[h] = v;
    });
    return o;
  });
}

function toRow_(k, o){
  return COLS[k].map(h => {
    const v = o[h];
    if(JSON_FIELDS.indexOf(h) >= 0) return JSON.stringify(v || []);
    return v == null ? '' : String(v);
  });
}

function upsert_(k, keyField, o){
  const s = sheet_(k);
  const hit = rows_(k).find(r => r[keyField] === o[keyField]);
  const row = toRow_(k, o);
  if(hit) s.getRange(hit._row,1,1,row.length).setNumberFormat('@').setValues([row]);
  else { const n = s.getLastRow()+1; s.getRange(n,1,1,row.length).setNumberFormat('@').setValues([row]); }
}

function remove_(k, keyField, val){
  const hit = rows_(k).find(r => r[keyField] === val);
  if(hit) sheet_(k).deleteRow(hit._row);
}

function replaceList_(k, names){
  const s = sheet_(k), n = s.getLastRow();
  if(n > 1) s.getRange(2,1,n-1,1).clearContent();
  const clean = names.map(x => String(x).trim()).filter(Boolean);
  if(clean.length) s.getRange(2,1,clean.length,1).setNumberFormat('@').setValues(clean.map(x => [x]));
}

function now_(){ return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm'); }
function strip_(o){ delete o._row; return o; }
function borrowText_(list){
  return (list || []).filter(x => x.s === 'B').map(x => x.n + (x.q ? ' ('+x.q+')' : '') + (x.from ? ' ← ' + x.from : '')).join('; ');
}

function getAll_(){
  const tpl = {};
  rows_('templates').forEach(t => { tpl[t.key] = {ins: t.ins, sup: t.sup, updatedAt: t.updatedAt}; });
  const recs = rows_('records').map(strip_).sort((a,b) => (b.savedAt > a.savedAt ? 1 : -1)).slice(0, 300);
  return {
    cases: rows_('cases').map(strip_),
    doctors: rows_('doctors').map(r => r.name).filter(Boolean),
    staff: rows_('staff').map(r => r.name).filter(Boolean),
    templates: tpl,
    records: recs,
    serverTime: now_(),
    version: VERSION
  };
}

const ACTIONS = {
  saveRecord(b){
    const r = b.record;
    if(!r || !r.caseId || !r.by) throw new Error('ข้อมูลไม่ครบ (เคส / ผู้ลงชื่อ)');
    const isEdit = !!r.id;
    r.id = r.id || ('r' + new Date().getTime() + Math.floor(Math.random()*1000));
    if(isEdit){ const old = rows_('records').find(x => x.id === r.id); r.savedAt = old ? old.savedAt : now_(); r.editedAt = now_(); }
    else { r.savedAt = now_(); r.editedAt = ''; }
    r.borrowIns = borrowText_(r.ins);
    r.borrowSup = borrowText_(r.sup);
    r.returned = r.returned || '';
    upsert_('records', 'id', r);
    if(r.doctor){
      const clean = list => (list || []).map(x => ({n: x.n, q: x.q || '', s: x.s || '', from: x.from || ''}));
      upsert_('templates', 'key', {key: r.doctor + '|' + r.caseId, doctor: r.doctor, caseId: r.caseId, caseName: r.caseName, ins: clean(r.ins), sup: clean(r.sup), updatedAt: now_()});
      if(rows_('doctors').map(x => x.name).indexOf(r.doctor) < 0) upsert_('doctors', 'name', {name: r.doctor});
    }
  },
  deleteRecord(b){ remove_('records', 'id', b.id); },
  setReturned(b){
    const hit = rows_('records').find(x => x.id === b.id);
    if(!hit) throw new Error('ไม่พบใบนี้');
    hit.returned = b.returned ? now_() + (b.who ? ' (' + b.who + ')' : '') : '';
    upsert_('records', 'id', strip_(hit));
  },
  saveCase(b){
    const c = b.case;
    if(!c || !c.name) throw new Error('ใส่ชื่อเคส');
    c.id = c.id || ('n' + new Date().getTime());
    c.cat = c.cat || 'เคสที่เพิ่มเอง';
    upsert_('cases', 'id', c);
  },
  deleteCase(b){ remove_('cases', 'id', b.id); },
  addDoctor(b){ const n = String(b.name || '').trim(); if(n && rows_('doctors').map(x => x.name).indexOf(n) < 0) upsert_('doctors', 'name', {name: n}); },
  savePeople(b){ replaceList_('doctors', b.doctors || []); replaceList_('staff', b.staff || []); },
  deleteTemplate(b){ remove_('templates', 'key', b.key); }
};

function check_(k){ if(TEAM_KEY && k !== TEAM_KEY) throw new Error('รหัสทีมไม่ถูกต้อง'); }
function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e){
  const p = (e && e.parameter) || {};
  try{
    check_(p.key || '');
    if(p.mode === 'export') return out_(export_());
    return out_({ok: true, data: getAll_()});
  }catch(err){ return out_({ok: false, error: String(err.message || err)}); }
}

// สำรองข้อมูล: ส่งทั้งไฟล์ชีตเป็น Excel (.xlsx) ให้ GitHub Actions นำไปเก็บใน OneDrive
// ต้องตั้ง TEAM_KEY ก่อน (กันคนนอกดึงข้อมูล)
// DriveApp.getFiles() — บรรทัดนี้ทำให้ Google ขอสิทธิ์อ่านไฟล์ เพื่อส่งออกเป็น Excel
function export_(){
  if(!TEAM_KEY) throw new Error('ต้องตั้ง TEAM_KEY ก่อนใช้การสำรองข้อมูล');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx', {
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true
  });
  if(res.getResponseCode() !== 200) throw new Error('ส่งออก Excel ไม่สำเร็จ (' + res.getResponseCode() + ')');
  return {ok: true, version: VERSION, name: ss.getName(), records: Math.max(0, sheet_('records').getLastRow() - 1), xlsx: Utilities.base64Encode(res.getBlob().getBytes())};
}

// ทดสอบการสำรองจากหน้า Apps Script: เลือกฟังก์ชันนี้แล้วกด Run (ครั้งแรกจะขออนุญาตเพิ่ม)
function testExport(){ const r = export_(); Logger.log('OK: ' + r.records + ' ใบ, ขนาด ' + r.xlsx.length + ' ตัวอักษร'); }

function doPost(e){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(25000);
    const b = JSON.parse(e.postData.contents);
    check_(b.key || '');
    const fn = ACTIONS[b.action];
    if(!fn) throw new Error('ไม่รู้จักคำสั่ง ' + b.action);
    fn(b);
    return out_({ok: true, data: getAll_()});
  }catch(err){
    return out_({ok: false, error: String(err.message || err)});
  }finally{
    try{ lock.releaseLock(); }catch(e){}
  }
}

// ข้อมูลเคสเริ่มต้น (มาจากไฟล์ของห้องผ่าตัด) — ใช้ครั้งแรกตอน setup เท่านั้น แก้ต่อในแอปหรือในชีต "เคส" ได้
const SEED = [
{
"id": "c1",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "ORIF Philos plate",
"ins": [
[
"Small fragment",
""
],
[
"LCP 3.5 set",
""
],
[
"Set General",
""
],
[
"Set Reduction",
""
]
],
"sup": [
[
"Plate (ระบุชนิด: LCP / Philos / Hook / DMT / Narrow / L plate / Radius)",
"",
""
],
[
"Screw (ระบุชนิด/ขนาด)",
"",
""
],
[
"Screw cortex Extra long",
"",
""
],
[
"ลวด no.23 / K-wire (TWB)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
],
[
"Radivac drain no.8 / 200 ml / 600 ml",
"",
"I"
],
[
"Image intensifier cover",
"",
"I"
],
[
"Ioban 6650",
"",
"OI"
],
[
"Leukostrip 13x102",
"",
""
],
[
"Esmarch / Webril / Conform / EB (ระบุ)",
"",
""
]
]
},
{
"id": "c2",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "ORIF LCP distal humerus",
"ins": [
[
"Set General",
""
],
[
"Set Reduction",
""
],
[
"LCP 3.5 set",
""
],
[
"Small fragment",
""
],
[
"Bone clamp กลาง",
"O"
]
],
"sup": [
[
"Plate (ระบุชนิด: LCP / Philos / Hook / DMT / Narrow / L plate / Radius)",
"",
""
],
[
"Screw (ระบุชนิด/ขนาด)",
"",
""
],
[
"Screw cortex Extra long",
"",
""
],
[
"ลวด no.23 / K-wire (TWB)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
],
[
"Radivac drain no.8 / 200 ml / 600 ml",
"",
"I"
],
[
"Image intensifier cover",
"",
"I"
],
[
"Ioban 6650",
"",
"OI"
],
[
"Leukostrip 13x102",
"",
""
],
[
"Esmarch / Webril / Conform / EB (ระบุ)",
"",
""
]
]
},
{
"id": "c3",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "ORIF Hook plate",
"ins": [
[
"Small fragment",
""
]
],
"sup": [
[
"Plate (ระบุชนิด: LCP / Philos / Hook / DMT / Narrow / L plate / Radius)",
"",
""
],
[
"Screw (ระบุชนิด/ขนาด)",
"",
""
],
[
"Screw cortex Extra long",
"",
""
],
[
"ลวด no.23 / K-wire (TWB)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
],
[
"Radivac drain no.8 / 200 ml / 600 ml",
"",
"I"
],
[
"Image intensifier cover",
"",
"I"
],
[
"Ioban 6650",
"",
"OI"
],
[
"Leukostrip 13x102",
"",
""
],
[
"Esmarch / Webril / Conform / EB (ระบุ)",
"",
""
]
]
},
{
"id": "c4",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "ORIF medial proximal tibia / L plate",
"ins": [
[
"Set plate medial proximal tibia",
""
],
[
"Set L plate proximal tibia",
""
],
[
"Small fragment",
""
],
[
"LCP 3.5 set",
""
],
[
"Set General",
""
],
[
"Set Reduction",
""
],
[
"Set chisel",
""
],
[
"Depth gauge long",
""
],
[
"Pelvic reduction forceps",
""
],
[
"Drill long 2.5",
""
]
],
"sup": [
[
"Plate (ระบุชนิด: LCP / Philos / Hook / DMT / Narrow / L plate / Radius)",
"",
""
],
[
"Screw (ระบุชนิด/ขนาด)",
"",
""
],
[
"Screw cortex Extra long",
"",
""
],
[
"ลวด no.23 / K-wire (TWB)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
],
[
"Radivac drain no.8 / 200 ml / 600 ml",
"",
"I"
],
[
"Image intensifier cover",
"",
"I"
],
[
"Ioban 6650",
"",
"OI"
],
[
"Leukostrip 13x102",
"",
""
],
[
"Esmarch / Webril / Conform / EB (ระบุ)",
"",
""
]
]
},
{
"id": "c5",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "ORIF distal radius",
"ins": [
[
"Set radius (instrument)",
""
],
[
"Sleeve two column",
""
]
],
"sup": [
[
"Plate (ระบุชนิด: LCP / Philos / Hook / DMT / Narrow / L plate / Radius)",
"",
""
],
[
"Screw (ระบุชนิด/ขนาด)",
"",
""
],
[
"Screw cortex Extra long",
"",
""
],
[
"ลวด no.23 / K-wire (TWB)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
],
[
"Radivac drain no.8 / 200 ml / 600 ml",
"",
"I"
],
[
"Image intensifier cover",
"",
"I"
],
[
"Ioban 6650",
"",
"OI"
],
[
"Leukostrip 13x102",
"",
""
],
[
"Esmarch / Webril / Conform / EB (ระบุ)",
"",
""
]
]
},
{
"id": "c6",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "ORIF TWB",
"ins": [
[
"Set wire",
""
],
[
"Set chisel",
""
],
[
"ลวดเล็ก (ถ้าเป็น ulnar ต้องยืม)",
""
]
],
"sup": [
[
"Plate (ระบุชนิด: LCP / Philos / Hook / DMT / Narrow / L plate / Radius)",
"",
""
],
[
"Screw (ระบุชนิด/ขนาด)",
"",
""
],
[
"Screw cortex Extra long",
"",
""
],
[
"ลวด no.23 / K-wire (TWB)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
],
[
"Radivac drain no.8 / 200 ml / 600 ml",
"",
"I"
],
[
"Image intensifier cover",
"",
"I"
],
[
"Ioban 6650",
"",
"OI"
],
[
"Leukostrip 13x102",
"",
""
],
[
"Esmarch / Webril / Conform / EB (ระบุ)",
"",
""
]
]
},
{
"id": "c7",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "ORIF distal medial tibia LCP (DMT)",
"ins": [
[
"Set General",
""
],
[
"Set Reduction",
""
],
[
"Instrument 4.5",
""
],
[
"Basic set",
""
],
[
"Screw set",
""
],
[
"LCP 4.5 set",
""
],
[
"LCP 3.5 set",
""
],
[
"Small fragment",
""
]
],
"sup": [
[
"Plate (ระบุชนิด: LCP / Philos / Hook / DMT / Narrow / L plate / Radius)",
"",
""
],
[
"Screw (ระบุชนิด/ขนาด)",
"",
""
],
[
"Screw cortex Extra long",
"",
""
],
[
"ลวด no.23 / K-wire (TWB)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
],
[
"Radivac drain no.8 / 200 ml / 600 ml",
"",
"I"
],
[
"Image intensifier cover",
"",
"I"
],
[
"Ioban 6650",
"",
"OI"
],
[
"Leukostrip 13x102",
"",
""
],
[
"Esmarch / Webril / Conform / EB (ระบุ)",
"",
""
]
]
},
{
"id": "c8",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "ORIF distal fibula / distal clavicle LCP",
"ins": [
[
"Small fragment",
""
],
[
"Instrument 2.7",
""
],
[
"Depth gauge เล็ก",
""
],
[
"Set wire pusher (clavicle)",
""
]
],
"sup": [
[
"Plate (ระบุชนิด: LCP / Philos / Hook / DMT / Narrow / L plate / Radius)",
"",
""
],
[
"Screw (ระบุชนิด/ขนาด)",
"",
""
],
[
"Screw cortex Extra long",
"",
""
],
[
"ลวด no.23 / K-wire (TWB)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
],
[
"Radivac drain no.8 / 200 ml / 600 ml",
"",
"I"
],
[
"Image intensifier cover",
"",
"I"
],
[
"Ioban 6650",
"",
"OI"
],
[
"Leukostrip 13x102",
"",
""
],
[
"Esmarch / Webril / Conform / EB (ระบุ)",
"",
""
]
]
},
{
"id": "c9",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "External fix",
"ins": [
[
"Set External fix",
""
],
[
"T handle",
""
],
[
"T universal",
""
],
[
"Set General",
""
],
[
"Set Reduction",
""
],
[
"ประแจ",
""
]
],
"sup": [
[
"Pin / Rod ชุด External fix (ระบุ)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c10",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "ORIF Narrow LCP / Small LCP",
"ins": [
[
"LCP 4.5 set",
""
],
[
"LCP 3.5 set",
""
],
[
"Set screw",
""
],
[
"Set basic",
"O"
],
[
"Set General",
""
],
[
"Set Reduction",
""
],
[
"ที่ดัด Plate ใหญ่",
"O"
],
[
"Small fragment",
""
]
],
"sup": [
[
"Plate (ระบุชนิด: LCP / Philos / Hook / DMT / Narrow / L plate / Radius)",
"",
""
],
[
"Screw (ระบุชนิด/ขนาด)",
"",
""
],
[
"Screw cortex Extra long",
"",
""
],
[
"ลวด no.23 / K-wire (TWB)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
],
[
"Radivac drain no.8 / 200 ml / 600 ml",
"",
"I"
],
[
"Image intensifier cover",
"",
"I"
],
[
"Ioban 6650",
"",
"OI"
],
[
"Leukostrip 13x102",
"",
""
],
[
"Esmarch / Webril / Conform / EB (ระบุ)",
"",
""
]
]
},
{
"id": "c11",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "Repair AC & CC ligament",
"ins": [
[
"Set wire pusher",
""
],
[
"Drill 2.0",
""
]
],
"sup": [
[
"Suture button",
"",
""
],
[
"Anchor suture",
"",
""
],
[
"Fiber wire 5 (ไม่ติดเข็ม)",
"",
""
],
[
"Ethibond 5 (ไม่ติดเข็ม)",
"",
""
],
[
"Prolene 1",
"",
""
],
[
"ลวด no.23",
"",
""
]
]
},
{
"id": "c12",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "PFNA",
"ins": [
[
"Life Plus",
""
],
[
"Set General",
""
],
[
"Set Reduction",
""
],
[
"PFNA set",
""
],
[
"Synream set",
""
],
[
"Ruler",
""
],
[
"Guide nail",
""
]
],
"sup": [
[
"PFNA Blade",
"",
""
],
[
"PFNA small / long",
"",
""
],
[
"PFNA End cap / End cap IM nail (แจ้ง Inven)",
"",
""
],
[
"IM nail (Tibia / Femur ระบุขนาด)",
"",
""
],
[
"Radivac drain",
"",
"I"
]
]
},
{
"id": "c13",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "IM nail",
"ins": [
[
"Nail instrument set",
""
],
[
"Set expert",
""
],
[
"Life plus",
""
],
[
"Set General",
""
],
[
"Set Reduction",
""
],
[
"Synream set",
""
],
[
"Ruler",
""
],
[
"Depth gauge",
""
],
[
"Guide nail",
""
],
[
"Bone hook",
""
],
[
"T universal",
""
],
[
"Hammer",
""
],
[
"F-ret (ถ้าทำ Femur)",
""
]
],
"sup": [
[
"PFNA Blade",
"",
""
],
[
"PFNA small / long",
"",
""
],
[
"PFNA End cap / End cap IM nail (แจ้ง Inven)",
"",
""
],
[
"IM nail (Tibia / Femur ระบุขนาด)",
"",
""
],
[
"Radivac drain",
"",
"I"
]
]
},
{
"id": "c14",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "ORIF proximal lateral tibia",
"ins": [
[
"Set proximal tibia (plate อยู่ในชุด)",
""
],
[
"Set basic",
"O"
],
[
"Set General",
""
],
[
"Set Reduction",
""
],
[
"Pelvic clamp",
""
],
[
"Instrument 4.5",
""
],
[
"Depth gauge long",
""
],
[
"Bone hook",
""
],
[
"Bone curette",
"O"
]
],
"sup": [
[
"Plate (ระบุชนิด: LCP / Philos / Hook / DMT / Narrow / L plate / Radius)",
"",
""
],
[
"Screw (ระบุชนิด/ขนาด)",
"",
""
],
[
"Screw cortex Extra long",
"",
""
],
[
"ลวด no.23 / K-wire (TWB)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
],
[
"Radivac drain no.8 / 200 ml / 600 ml",
"",
"I"
],
[
"Image intensifier cover",
"",
"I"
],
[
"Ioban 6650",
"",
"OI"
],
[
"Leukostrip 13x102",
"",
""
],
[
"Esmarch / Webril / Conform / EB (ระบุ)",
"",
""
]
]
},
{
"id": "c15",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "TKA",
"ins": [
[
"Set General",
""
],
[
"Set Osteotome ใหญ่",
""
],
[
"Mallet ใหญ่",
"O"
],
[
"Bone rongeur",
"O"
],
[
"Volkman",
""
],
[
"Bone curette ตรง",
"O"
],
[
"Bone hook / T handle",
""
],
[
"Hohmann retractor large wide",
""
],
[
"Hohmann retractor small narrow",
""
],
[
"ตัวดันขาข้าง",
""
]
],
"sup": [
[
"Implant TKA (ระบุยี่ห้อ/ขนาด)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c16",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "ACL",
"ins": [],
"sup": [
[
"Suture button",
"",
""
],
[
"Anchor suture",
"",
""
],
[
"Fiber wire 5 (ไม่ติดเข็ม)",
"",
""
],
[
"Ethibond 5 (ไม่ติดเข็ม)",
"",
""
],
[
"Prolene 1",
"",
""
],
[
"ลวด no.23",
"",
""
]
]
},
{
"id": "c17",
"cat": "B. Spine / พลาสติก / ENT / Uro",
"name": "Open discectomy",
"ins": [
[
"Set Laminectomy",
""
],
[
"Set special spine",
""
],
[
"Metric instrument",
""
],
[
"Anterior cervical ret.",
""
],
[
"Set Midas + เครื่อง (หัวกรอยืม Inven)",
""
],
[
"Frame",
""
],
[
"Kerrison 2 mm / 3 mm",
""
],
[
"ข้อต่อเตียงท่อนสีดำ",
""
]
],
"sup": [
[
"Spongostan",
"",
"I"
],
[
"Surgicel",
"",
"OI"
],
[
"เวชภัณฑ์ตามใบแพทย์ (ระบุ)",
"",
""
]
]
},
{
"id": "c18",
"cat": "B. Spine / พลาสติก / ENT / Uro",
"name": "STSG with mesh",
"ins": [
[
"Set Dermatome",
""
],
[
"Set mesh graft",
""
],
[
"ข้อต่อลม",
""
]
],
"sup": [
[
"Dermatome blade",
"",
""
],
[
"Dermacarrier 1.5:1",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c19",
"cat": "B. Spine / พลาสติก / ENT / Uro",
"name": "ORIF Zygoma / orbital",
"ins": [
[
"Set plastic small",
"O"
],
[
"Molt periosteal",
""
],
[
"Zygoma elevator",
""
],
[
"Drill 1.6",
""
],
[
"Set Leibinger",
""
],
[
"Midas instrument + เครื่อง + drill",
""
],
[
"Brain retractor",
""
],
[
"Vein retractor",
""
],
[
"Shape hook",
""
],
[
"ตัวงัด orbital",
""
],
[
"Wire scissor (ถ้าวาง mesh)",
""
]
],
"sup": [
[
"Mesh / plate / screw (ระบุ)",
"",
""
],
[
"ลวด Extra",
"",
""
],
[
"ยางมัดฟัน",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c20",
"cat": "B. Spine / พลาสติก / ENT / Uro",
"name": "IMF",
"ins": [
[
"Wire pusher",
""
],
[
"Set arch bar",
""
],
[
"Wire scissor",
"O"
],
[
"Tonsil suction",
""
],
[
"E pen drive",
""
],
[
"Mouth gag",
""
]
],
"sup": [
[
"Mesh / plate / screw (ระบุ)",
"",
""
],
[
"ลวด Extra",
"",
""
],
[
"ยางมัดฟัน",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c21",
"cat": "B. Spine / พลาสติก / ENT / Uro",
"name": "Vasectomy",
"ins": [
[
"Set TR",
"O"
]
],
"sup": [
[
"เวชภัณฑ์ตามใบแพทย์ (ระบุ)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c22",
"cat": "B. Spine / พลาสติก / ENT / Uro",
"name": "Thyroidectomy",
"ins": [
[
"สาย Harmonic",
""
],
[
"เครื่อง Harmonic",
""
]
],
"sup": [
[
"เวชภัณฑ์ตามใบแพทย์ (ระบุ)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c23",
"cat": "B. Spine / พลาสติก / ENT / Uro",
"name": "Tonsillectomy",
"ins": [
[
"Set tonsil",
""
]
],
"sup": [
[
"เวชภัณฑ์ตามใบแพทย์ (ระบุ)",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c24",
"cat": "C. นรีเวช / ส่องกล้อง / Uro scope",
"name": "TAH",
"ins": [
[
"Explore set",
""
],
[
"Cath set",
"O"
],
[
"Scrub set",
"O"
],
[
"Valley lab",
"O"
],
[
"Myoma screw",
""
]
],
"sup": [
[
"Myoma screw",
"",
""
],
[
"Interceed",
"",
""
],
[
"Fibrillar",
"",
""
],
[
"Surgicel",
"",
"OI"
],
[
"Spongostan",
"",
"I"
],
[
"หัวจี้ยาว E1450-6",
"",
""
],
[
"Megadyne M/Blade 6 1/2\" Ez (0014M)",
"",
"I"
],
[
"Dermabond advanced / Mepilex post-op 10x20, 10x25",
"",
"I"
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c25",
"cat": "C. นรีเวช / ส่องกล้อง / Uro scope",
"name": "LC / Lap appendectomy",
"ins": [
[
"Lens 5 mm 0°",
""
],
[
"Lens 10 mm 0°",
""
],
[
"Lens 10 mm 30°",
""
],
[
"Set camera",
""
],
[
"สายน้ำฟ้าขาว",
""
],
[
"Suction and irrigate",
""
],
[
"กระป๋องน้ำ LC",
""
],
[
"Cautery cable LC",
""
],
[
"LC set",
""
],
[
"ด้าม Hemoloc 5 / 10 / 12 mm",
""
],
[
"Pressure bag",
""
],
[
"Harmonic (พ.ณัฐพงศ์) / Ligasure (พ.บัณฑิต, พ.ปราการ)",
""
],
[
"(ใช้ของบริษัท) จุกน้ำ + จุกแก๊ส / สายน้ำ + สายแก๊ส",
""
]
],
"sup": [
[
"Surgical trocar 12 mm (D12LT)",
"2",
""
],
[
"Dilating tip trocar 5 mm (D5LT)",
"3–4",
""
],
[
"Ligamex 5 mm / 12 mm",
"",
""
],
[
"Endoloop (PDS)",
"3",
""
],
[
"Hemoloc ทุก size",
"",
""
],
[
"Polymer clip 5 / 10 / 12 mm",
"",
""
],
[
"Endobag 3x6 / 5x8 / Lina endobag",
"",
""
],
[
"Endo scissor 5 mm",
"",
""
],
[
"Ligasure Maryland 37 cm LF1937 / Harmonic Ace 36 cm",
"",
""
],
[
"Jackson Pratt 400 / Jackson round drain 15 Fr",
"",
"I"
],
[
"Glove port",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c26",
"cat": "C. นรีเวช / ส่องกล้อง / Uro scope",
"name": "Lap ovarian cyst / TLH",
"ins": [
[
"Laparoscope set",
""
],
[
"Lens 5 mm / 10 mm",
""
],
[
"Cautery cable micro jaw",
""
],
[
"Cautery cable LC",
""
],
[
"Uterine elevator",
""
],
[
"Endo Allis",
""
],
[
"Vaginal speculum L",
""
],
[
"สายน้ำฟ้าขาว",
""
],
[
"Suction and irrigate",
""
],
[
"กระป๋องน้ำ LC",
""
],
[
"ขาหยั่งดำ",
""
],
[
"เครื่อง Ligasure",
""
],
[
"Gynaecology set",
""
],
[
"Vagina tube / Uterine sound / Needle handle LC (TLH)",
""
]
],
"sup": [
[
"Dilating tip trocar 5 mm with stability sleeve (D5LT)",
"4",
""
],
[
"Surgical trocar 11 mm bladeless (B11LT)",
"2",
""
],
[
"Surgical trocar 12 mm (D12LT)",
"1",
""
],
[
"Bipolar macro jaw (Ebfo1)",
"",
""
],
[
"Ligasure blunt / Maryland 37 cm",
"1 / 1",
""
],
[
"Endo scissor 5 mm",
"",
""
],
[
"Endobag 3x6 / 5x8",
"",
""
],
[
"Interceed 3x4",
"",
""
],
[
"V-loc 180 (No.0316)  [TLH]",
"2",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c27",
"cat": "C. นรีเวช / ส่องกล้อง / Uro scope",
"name": "TUERP",
"ins": [
[
"Lens 4 mm 30°",
""
],
[
"Set cystoscope",
""
],
[
"หม้อ TUR",
""
],
[
"สายน้ำ LC",
""
],
[
"Stop cock",
""
],
[
"Sound dilator",
""
],
[
"Guide foley",
""
]
],
"sup": [
[
"Sterile water irrigate 1000 ml (ใช้ 2–3 ขวด)",
"",
"I"
],
[
"สายน้ำ + Stop cock",
"",
""
],
[
"เวชภัณฑ์ LEEP (ระบุ)",
"",
""
]
]
},
{
"id": "c28",
"cat": "C. นรีเวช / ส่องกล้อง / Uro scope",
"name": "Hysteroscope",
"ins": [
[
"Set hysteroscope",
""
],
[
"Cutting loop hystero",
""
],
[
"Lens 4 mm 30°",
""
],
[
"สายจี้",
""
],
[
"Pressure bag",
""
],
[
"Camera",
""
],
[
"จุกเสียบสายน้ำ",
""
]
],
"sup": [
[
"Sterile water irrigate 1000 ml (ใช้ 2–3 ขวด)",
"",
"I"
],
[
"สายน้ำ + Stop cock",
"",
""
],
[
"เวชภัณฑ์ LEEP (ระบุ)",
"",
""
]
]
},
{
"id": "c29",
"cat": "C. นรีเวช / ส่องกล้อง / Uro scope",
"name": "LEEP",
"ins": [
[
"เครื่องจี้ LEEP",
""
],
[
"สายจี้ conization",
""
],
[
"Set conization",
""
],
[
"หัวจี้ Loop + หัว ball (ยืม Inven)",
""
]
],
"sup": [
[
"Sterile water irrigate 1000 ml (ใช้ 2–3 ขวด)",
"",
"I"
],
[
"สายน้ำ + Stop cock",
"",
""
],
[
"เวชภัณฑ์ LEEP (ระบุ)",
"",
""
]
]
},
{
"id": "c30",
"cat": "B. ศัลยกรรมทั่วไป / ทวารหนัก / Spine / พลาสติก / ENT / Uro",
"name": "Excision / Hemorrhoid / I+D / Bartholin",
"ins": [
[
"Minor set",
"O"
],
[
"Plastic small",
"O"
],
[
"Coronado",
""
],
[
"Valley lab",
"O"
],
[
"Surgical pack",
""
]
],
"sup": [
[
"Anal spongostan / Spongostan",
"",
"I"
],
[
"Penrose drain ทุก size",
"",
"O"
],
[
"Dermabond 0.5",
"",
"O"
],
[
"Megadyne micro needle 0118",
"",
"I"
],
[
"Horizon titanium clip (small)",
"",
""
],
[
"Marcaine 0.5 %  /  Chalocaine 2 % jelly",
"",
""
],
[
"Hydrogen peroxide / Tube culture",
"",
"I"
],
[
"Hand control pencil / Urine catheter red rubber no.14",
"",
"I"
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c31",
"cat": "B. ศัลยกรรมทั่วไป / ทวารหนัก / Spine / พลาสติก / ENT / Uro",
"name": "Explor lap / Open colectomy",
"ins": [
[
"Explor lap set",
"O"
],
[
"Intestinal set",
""
],
[
"Omnitract ret. set",
""
],
[
"Scrub set",
"O"
],
[
"Cath set",
"O"
],
[
"Valley lab",
"O"
],
[
"Harmonic scapler",
""
],
[
"Drape x5",
"O"
],
[
"Colonoscope (ถ้าต้องใช้)",
""
]
],
"sup": [
[
"Endo GIA stapler 80-3.8 / 75 mm NTLC75 + reload SR75",
"",
""
],
[
"Contour curved cutter 40 mm Green CS40G + reload CR40G",
"",
""
],
[
"CDH 25 / 29 / 33  /  Pursting device",
"",
""
],
[
"Harmonic focus 17 cm / Harmonic 1100 20 cm / Ace",
"",
""
],
[
"Ligasure Impact LF4318 / Atlas LS1020",
"",
""
],
[
"Alexis wound retractor (ทุก size)",
"",
""
],
[
"Hemoloc",
"",
""
],
[
"Skin staple",
"",
"OI"
],
[
"Surgicel 5x7 W1913T",
"",
"OI"
],
[
"Redivac 200 ml / สาย no.8",
"",
"I"
],
[
"Colostomy / Stomahesive",
"",
""
],
[
"Endo ICG",
"",
""
],
[
"ไหมตามใบเตรียมแพทย์",
"",
"O"
]
]
},
{
"id": "c32",
"cat": "A. ออร์โธปิดิกส์ / กระดูก",
"name": "อื่นๆ (ระบุเอง)",
"ins": [],
"sup": []
}
];
