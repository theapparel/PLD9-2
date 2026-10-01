/**
 * พอแล้วดีสโมสร — follow-up board backend
 * ------------------------------------------------------------------
 * Paste into the Apps Script editor of the Google Sheet that holds the
 * board's data (Extensions ▸ Apps Script), run `setup` once, then
 * Deploy ▸ New deployment ▸ Web app, "Execute as: Me",
 * "Who has access: Anyone". Put the /exec URL into index.html (API_URL).
 *
 * One tab, "Data": one row per record.
 *   collection | id | json | updatedAt
 * collections used by the page: tasks, comments, agenda, links, meta
 *
 * Optional edit PIN: Project Settings ▸ Script properties ▸ add
 * EDIT_PIN = 1234. Then anyone can still view, but saving asks for the PIN.
 */

var TAB = 'Data';
var HEAD = ['collection', 'id', 'json', 'updatedAt'];
var COLLECTIONS = ['tasks', 'comments', 'agenda', 'links', 'meta'];

function setup() {
  sheet_();
  SpreadsheetApp.getActiveSpreadsheet().toast('พร้อมใช้งาน: แท็บ Data ถูกสร้างแล้ว');
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.insertSheet(TAB);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(3, 600);
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- reads ---------- */

function readAll_() {
  var sh = sheet_();
  var n = sh.getLastRow();
  var rows = [];
  if (n < 2) return rows;
  var vals = sh.getRange(2, 1, n - 1, 4).getValues();
  for (var r = 0; r < vals.length; r++) {
    if (!vals[r][0] || !vals[r][1]) continue;
    var data = {};
    try { data = JSON.parse(vals[r][2] || '{}'); } catch (e) { data = {}; }
    rows.push({ c: String(vals[r][0]), id: String(vals[r][1]), data: data, row: r + 2 });
  }
  return rows;
}

function list_() {
  var rows = readAll_().map(function (x) { return { c: x.c, id: x.id, data: x.data }; });
  return { ok: true, rows: rows, ts: Date.now(), pin: !!pin_() };
}

/* ---------- writes ---------- */

function pin_() {
  return PropertiesService.getScriptProperties().getProperty('EDIT_PIN') || '';
}

function isObj_(v) { return v && typeof v === 'object' && !Array.isArray(v); }

/* nested objects merge; arrays and everything else replace; null is kept as null */
function merge_(base, patch) {
  var outv = isObj_(base) ? JSON.parse(JSON.stringify(base)) : {};
  Object.keys(patch).forEach(function (k) {
    var v = patch[k];
    if (isObj_(v) && isObj_(outv[k])) outv[k] = merge_(outv[k], v);
    else outv[k] = v;
  });
  return outv;
}

function apply_(sh, index, op) {
  var c = String(op.c || ''), id = String(op.id || '');
  if (COLLECTIONS.indexOf(c) < 0) throw new Error('unknown collection: ' + c);
  if (!id || id.length > 200) throw new Error('bad id');
  var key = c + '/' + id;
  var hit = index[key];
  if (op.op === 'delete') {
    if (hit) { sh.deleteRow(hit.row); rebuild_(index); }
    return;
  }
  var data;
  if (op.op === 'update') {
    if (!hit) throw new Error('not found: ' + key);
    data = merge_(hit.data, op.data || {});
  } else if (op.op === 'set') {
    data = op.data || {};
  } else throw new Error('bad op');
  var json = JSON.stringify(data);
  if (json.length > 45000) throw new Error('record too large');
  if (hit) {
    sh.getRange(hit.row, 3, 1, 2).setValues([[json, new Date()]]);
    hit.data = data;
  } else {
    sh.appendRow([c, id, json, new Date()]);
    index[key] = { c: c, id: id, data: data, row: sh.getLastRow() };
  }
}

function rebuild_(index) {
  Object.keys(index).forEach(function (k) { delete index[k]; });
  readAll_().forEach(function (x) { index[x.c + '/' + x.id] = x; });
}

function write_(body) {
  var p = pin_();
  if (p && String(body.pin || '') !== p) return { ok: false, error: 'pin' };
  var ops = body.action === 'batch' ? (body.ops || []) : [body];
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_();
    var index = {};
    rebuild_(index);
    ops.forEach(function (op) { apply_(sh, index, { op: op.op || op.action, c: op.c, id: op.id, data: op.data }); });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { ok: true, ts: Date.now() };
}

/* ---------- entry points ---------- */

function doGet(e) {
  try {
    var a = (e && e.parameter && e.parameter.action) || 'list';
    if (a === 'list') return out_(list_());
    if (a === 'checkpin') return out_({ ok: !pin_() || String(e.parameter.pin || '') === pin_() });
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var a = body.action;
    if (a === 'list') return out_(list_());
    if (a === 'set' || a === 'update' || a === 'delete') { body.op = a; return out_(write_(body)); }
    if (a === 'batch') return out_(write_(body));
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}
