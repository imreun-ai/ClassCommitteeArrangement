/**
 * ប្រព័ន្ធ គគថ — Google Sheet ជាមូលដ្ឋានទិន្នន័យ (ទាញ និងផ្ញើ ទាំងពីរទិស)
 * ដាក់ក្នុង Extensions → Apps Script ហើយ Deploy ជា Web app
 * (Execute as: Me · Who has access: Anyone)
 * កែ/លុបជួរក្នុង Sheet → ប្រព័ន្ធប្តូរតាម។ កែក្នុងប្រព័ន្ធ → Sheet ប្តូរតាម។
 */
const SHEET_ID = '1PwnCl_EXaIKbP3k6KD90WKxzfrJeIbhX23RPX8YIGao';
const S_INFO = 'ព័ត៌មានទូទៅ', S_CLS = 'បញ្ជីថ្នាក់', S_MEM = 'សមាជិក',
      S_PLAN = 'ផែនការប្រចាំឆ្នាំ', S_MEET = 'កិច្ចប្រជុំ', S_PHOTO = '_រូបថត';
const CHUNK = 40000;
const LBL = {name:'ឈ្មោះ', sex:'ភេទ', age:'អាយុ', cat:'ប្រភេទ', role:'តួនាទី', job:'មុខរបរ', task:'ភារកិច្ច',
             date:'កាលបរិច្ឆេទ', place:'ទីកន្លែង', topic:'ប្រធានបទ', notes:'កំណត់ចំណាំ', decision:'សេចក្តីសម្រេច'};
const REV = {}; Object.keys(LBL).forEach(k => REV[LBL[k]] = k);
const ORDER = ['name','sex','age','cat','role','job','task','date','place','topic','notes','decision'];

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_(n) { const s = ss_(); return s.getSheetByName(n) || s.insertSheet(n); }   // បង្កើត Sheet ស្វ័យប្រវត្តិ

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
    write_(b.state, b.photos || null);
    return out_({ ok: true, saved: new Date().toISOString() });
  } catch (x) { return out_({ ok: false, error: String(x) }); }
  finally { try { lock.releaseLock(); } catch (y) {} }
}

/* ---------- សរសេរ ---------- */
function write_(S, photos) {
  const cls = S.classes || [], info = S.info || {};
  put_(S_INFO, ['ចំណងជើង', 'តម្លៃ'], Object.keys(info).map(k => [k, info[k]]));
  put_(S_CLS, ['ថ្នាក់', 'លេខថ្នាក់ (កុំកែ)', 'ខែក្នុងផែនការ (ញែកដោយ ,)'],
       cls.map(c => [c.grade, c.id, (c.months || []).join(',')]));
  flat_(S_MEM, cls, 'members');
  flat_(S_MEET, cls, 'meetings');

  let maxM = 0; cls.forEach(c => maxM = Math.max(maxM, (c.months || []).length));
  const head = ['ថ្នាក់', 'លេខថ្នាក់', 'លេខកូដ', 'សកម្មភាព'];
  for (let i = 1; i <= maxM; i++) head.push('ខែទី' + i);
  const plan = [];
  cls.forEach(c => (c.acts || []).forEach(a => {
    const r = [c.grade, c.id, a.id, a.t];
    for (let i = 0; i < maxM; i++) r.push(i >= (c.months || []).length ? '' : a.c && a.c[i] == 1 ? '✓' : a.c && a.c[i] == 2 ? '●' : '');
    plan.push(r);
  }));
  put_(S_PLAN, head, plan);

  if (photos) {                                     // សរសេររូបថតតែពេលមានការប្ដូរ
    const rows = [];
    Object.keys(photos).forEach(k => {
      const v = String(photos[k] || '');
      for (let i = 0, n = 0; i < v.length; i += CHUNK, n++) rows.push([k, n, v.substr(i, CHUNK)]);
    });
    put_(S_PHOTO, ['គន្លឹះ', 'ផ្នែក', 'ទិន្នន័យ (កុំកែ)'], rows);
  }
  ['ថ្នាក់', '_Data'].forEach(n => { const s = ss_().getSheetByName(n); if (s && ss_().getSheets().length > 1) ss_().deleteSheet(s); });
}

function flat_(name, cls, key) {
  const keys = [];
  cls.forEach(c => (c[key] || []).forEach(o => Object.keys(o).forEach(k => {
    if (k !== 'id' && k !== 'photo' && keys.indexOf(k) < 0) keys.push(k);
  })));
  keys.sort((a, b) => (ORDER.indexOf(a) < 0 ? 99 : ORDER.indexOf(a)) - (ORDER.indexOf(b) < 0 ? 99 : ORDER.indexOf(b)));
  const rows = [];
  cls.forEach(c => (c[key] || []).forEach(o =>
    rows.push([c.grade, c.id, o.id].concat(keys.map(k => o[k] == null ? '' : o[k])))));
  put_(name, ['ថ្នាក់', 'លេខថ្នាក់', 'លេខកូដ'].concat(keys.map(k => LBL[k] || k)), rows);
}

function put_(name, head, rows) {
  const sh = sheet_(name), w = head.length;
  sh.clear();
  sh.getRange(1, 1, 1, w).setValues([head]).setFontWeight('bold').setBackground('#e8eefc');
  if (rows.length) {
    rows = rows.map(r => { r = r.slice(0, w).map(v => v == null ? '' : String(v)); while (r.length < w) r.push(''); return r; });
    sh.getRange(2, 1, rows.length, w).setNumberFormat('@').setValues(rows);
  }
  sh.setFrozenRows(1);
}

/* ---------- អាន (Sheet ជាប្រភពពិត) ---------- */
function tbl_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 1) return null;
  const v = sh.getDataRange().getValues().map(r => r.map(x => x == null ? '' : String(x).trim()));
  return { head: v[0], rows: v.slice(1).filter(r => r.some(x => x !== '')) };
}
const num_ = x => (x !== '' && !isNaN(x)) ? Number(x) : x;

function read_() {
  const ct = tbl_(S_CLS);
  if (!ct || !ct.rows.length) return null;            // Sheet ទទេ → ប្រព័ន្ធផ្ញើទិន្នន័យឡើង

  const photos = {}, pt = tbl_(S_PHOTO);
  if (pt) {
    const parts = {};
    pt.rows.forEach(r => { (parts[r[0]] = parts[r[0]] || []).push([Number(r[1]), r[2]]); });
    Object.keys(parts).forEach(k => photos[k] = parts[k].sort((a, b) => a[0] - b[0]).map(p => p[1]).join(''));
  }
  const info = {}, it = tbl_(S_INFO);
  if (it) it.rows.forEach(r => { if (r[0] !== '') info[r[0]] = r[1]; });
  info.logo = photos.logo || '';

  const classes = [], byId = {}, byGrade = {};
  let maxC = 0; ct.rows.forEach(r => { const n = Number(r[1]); if (n > maxC) maxC = n; });
  ct.rows.forEach(r => {
    if (r[0] === '') return;
    const id = r[1] === '' ? ++maxC : Number(r[1]);
    const c = { id: id, grade: r[0], cover: photos['cover:' + id] || '',
                members: [], months: r[2] === '' ? [] : r[2].split(',').map(s => s.trim()), acts: [], meetings: [] };
    classes.push(c); byId[id] = c; byGrade[r[0]] = c;
  });
  if (!classes.length) return null;
  const find = r => byId[Number(r[1])] || byGrade[r[0]];

  [[S_MEM, 'members', true], [S_MEET, 'meetings', false]].forEach(([name, key, withPhoto]) => {
    const t = tbl_(name); if (!t) return;
    const keys = t.head.slice(3).map(h => REV[h] || h);
    t.rows.forEach(r => {
      const c = find(r); if (!c) return;               // ជួរដែលគ្មានថ្នាក់ត្រូវគ្នា ត្រូវបានរំលង
      const o = {}; keys.forEach((k, i) => { if (k) o[k] = r[3 + i] == null ? '' : r[3 + i]; });
      if (!Object.keys(o).some(k => o[k] !== '')) return;
      let max = 0; c[key].forEach(x => { if (x.id > max) max = x.id; });
      o.id = r[2] === '' ? max + 1 : num_(r[2]);
      if (key === 'members') o.photo = photos['m:' + c.id + ':' + o.id] || '';
      c[key].push(o);
    });
  });

  const pl = tbl_(S_PLAN);
  if (pl) pl.rows.forEach(r => {
    const c = find(r); if (!c || r[3] === '') return;
    const a = { id: r[2] === '' ? c.acts.length + 1 : num_(r[2]), t: r[3], c: {} };
    for (let i = 4; i < r.length && i - 4 < c.months.length; i++) {
      const v = r[i]; if (v === '') continue;
      a.c[i - 4] = (v === '●' || v === '2') ? 2 : 1;
    }
    c.acts.push(a);
  });
  return { v: 3, info: info, cur: classes[0].id, classes: classes };
}
