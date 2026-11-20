/**
 * ប្រព័ន្ធ គគថ — ទទួល/ផ្ញើទិន្នន័យជាមួយ Google Sheet
 * ដាក់ក្នុង Extensions → Apps Script ហើយ Deploy ជា Web app
 * (Execute as: Me · Who has access: Anyone)
 */
const SHEET_ID = '1PwnCl_EXaIKbP3k6KD90WKxzfrJeIbhX23RPX8YIGao';
const DATA = '_Data';          // ទិន្នន័យដើម (JSON) — ប្រព័ន្ធអានពីទីនេះ
const CHUNK = 40000;           // 1 cell ≤ 50,000 តួអក្សរ
const SKIP = {photo: 1, logo: 1, cover: 1};

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(name) {
  const ss = ss_();
  return ss.getSheetByName(name) || ss.insertSheet(name);   // បង្កើត Sheet ស្វ័យប្រវត្តិ
}

function doGet(e) {
  try {
    const a = (e && e.parameter && e.parameter.action) || 'ping';
    if (a === 'load') return out_({ ok: true, state: read_() });
    return out_({ ok: true, pong: true, name: ss_().getName() });
  } catch (x) { return out_({ ok: false, error: String(x) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const b = JSON.parse(e.postData.contents);
    if (b.action !== 'save' || !b.state) return out_({ ok: false, error: 'bad request' });
    write_(b.state);
    return out_({ ok: true, saved: new Date().toISOString() });
  } catch (x) { return out_({ ok: false, error: String(x) }); }
  finally { try { lock.releaseLock(); } catch (y) {} }
}

function read_() {
  const sh = ss_().getSheetByName(DATA);
  if (!sh || sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  const txt = rows.map(r => r[0]).join('');
  return txt ? JSON.parse(txt) : null;
}

function write_(S) {
  // 1) ទិន្នន័យដើម
  const sh = sheet_(DATA);
  const json = JSON.stringify(S), rows = [];
  for (let i = 0; i < json.length; i += CHUNK) rows.push([rows.length + 1, json.substr(i, CHUNK)]);
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([['#', 'JSON (កុំកែ)']]);
  sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  sh.getRange(1, 3).setValue('Updated: ' + new Date().toISOString());
  sh.setColumnWidth(2, 300);

  // 2) សន្លឹកអានងាយ (បង្កើតឡើងវិញរាល់ពេលរក្សាទុក)
  const info = S.info || {}, cls = S.classes || [];
  put_('ព័ត៌មានទូទៅ', ['ចំណងជើង', 'តម្លៃ'],
    Object.keys(info).filter(k => !SKIP[k]).map(k => [k, info[k]]));
  put_('ថ្នាក់', ['លេខ', 'ថ្នាក់', 'ចំនួនសមាជិក', 'ចំនួនកិច្ចប្រជុំ'],
    cls.map(c => [c.id, c.grade, (c.members || []).length, (c.meetings || []).length]));
  flat_('សមាជិក', cls, 'members');
  flat_('កិច្ចប្រជុំ', cls, 'meetings');

  const months = S.months || [], head = ['ថ្នាក់', 'សកម្មភាព'].concat(months.map(m => 'ខែ ' + m));
  const plan = [];
  cls.forEach(c => (c.acts || []).forEach(a =>
    plan.push([c.grade, a.t].concat(months.map((m, i) => (a.c && a.c[i]) ? '✓' : '')))));
  put_('ផែនការប្រចាំឆ្នាំ', head, plan);
}

function flat_(name, cls, key) {
  const keys = [];
  cls.forEach(c => (c[key] || []).forEach(o => Object.keys(o).forEach(k => {
    if (!SKIP[k] && keys.indexOf(k) < 0) keys.push(k);
  })));
  const rows = [];
  cls.forEach(c => (c[key] || []).forEach(o =>
    rows.push([c.grade].concat(keys.map(k => o[k] == null ? '' : String(o[k]))))));
  put_(name, ['ថ្នាក់'].concat(keys), rows);
}

function put_(name, head, rows) {
  const sh = sheet_(name);
  sh.clear();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8eefc');
  if (rows.length) {
    const w = head.length;
    rows = rows.map(r => { r = r.slice(0, w); while (r.length < w) r.push(''); return r; });
    sh.getRange(2, 1, rows.length, w).setNumberFormat('@').setValues(rows);
  }
  sh.setFrozenRows(1);
}
