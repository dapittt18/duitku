// ===== Duitku — backend (Google Apps Script, bound ke Google Sheet) =====
const SH = { tx: 'Transactions', acc: 'Accounts', cat: 'Categories' };
const TX_HEAD = ['id', 'date', 'type', 'amount', 'account', 'to_account', 'category', 'note', 'created_at', 'subcategory'];
const CAT_HEAD = ['name', 'type', 'icon', 'parent'];
const DEFAULT_SUBS = [
  ['Wifi', 'expense', '📶', 'Tagihan'], ['Pulsa & Data', 'expense', '📱', 'Tagihan'],
  ['Listrik', 'expense', '⚡', 'Tagihan'], ['Air', 'expense', '🚰', 'Tagihan'],
  ['Bensin', 'expense', '⛽', 'Transport'], ['Ojol', 'expense', '🛵', 'Transport'], ['Parkir', 'expense', '🅿️', 'Transport'],
];

// Jalanin SEKALI dari editor Apps Script buat bikin sheet + data default
function setup() {
  const ss = ss_();
  const mk = (name, head, rows) => {
    let s = ss.getSheetByName(name) || ss.insertSheet(name);
    if (s.getLastRow() === 0) {
      s.appendRow(head);
      s.setFrozenRows(1);
      if (rows && rows.length) s.getRange(2, 1, rows.length, head.length).setValues(rows);
    }
    return s;
  };
  mk(SH.acc, ['name', 'initial_balance'], [['Cash', 0], ['BCA', 0], ['GoPay', 0]]);
  mk(SH.cat, CAT_HEAD, [
    ['Makan', 'expense', '🍜'], ['Transport', 'expense', '🛵'], ['Belanja', 'expense', '🛒'],
    ['Tagihan', 'expense', '💡'], ['Hiburan', 'expense', '🎮'], ['Kesehatan', 'expense', '💊'],
    ['Lainnya', 'expense', '📦'],
    ['Gaji', 'income', '💼'], ['Bonus', 'income', '🎁'], ['Usaha', 'income', '🧶'], ['Lainnya', 'income', '💰'],
  ].map(r => r.concat('')).concat(DEFAULT_SUBS));
  const tx = mk(SH.tx, TX_HEAD);
  tx.getRange('B:B').setNumberFormat('@'); // simpan tanggal sbg teks yyyy-MM-dd
  tx.getRange('I:I').setNumberFormat('@');
}

// Jalanin SEKALI dari editor: bikin token API random, simpan di Script Properties, token muncul di log
function setToken() {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('API_TOKEN', token);
  Logger.log('API_TOKEN: ' + token);
  return token;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Duitku')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// ---------- JSON API (dipanggil dari PWA di GitHub Pages) ----------
// Body: {token, fn, args}. Dikirim sbg text/plain biar browser gak kirim CORS preflight.
const API_FNS_ = ['getData', 'saveTransaction', 'deleteTransaction', 'addAccount', 'addCategory'];

function doPost(e) {
  let res;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const token = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
    if (!token || req.token !== token) throw new Error('unauthorized');
    if (API_FNS_.indexOf(req.fn) < 0) throw new Error('unknown fn: ' + req.fn);
    const fns = { getData, saveTransaction, deleteTransaction, addAccount, addCategory };
    res = { ok: true, data: fns[req.fn].apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    res = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- helpers ----------
// Pake Sheet tempat script nempel; kalau script-nya standalone, bikin/buka "Duitku Data" otomatis
function ss_() {
  const active = SpreadsheetApp.getActive();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create('Duitku Data');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

// Upgrade sheet lama: tambah kolom yang belum ada (subcategory, parent) + sub kategori default
let migrated_ = false;
function migrate_() {
  if (migrated_) return;
  migrated_ = true;
  const ss = ss_();
  const addCols = (name, head) => {
    const s = ss.getSheetByName(name);
    if (!s || s.getLastRow() === 0) return [];
    const lc = s.getLastColumn();
    const cur = s.getRange(1, 1, 1, lc).getValues()[0];
    const add = head.filter(h => cur.indexOf(h) < 0);
    if (add.length) s.getRange(1, lc + 1, 1, add.length).setValues([add]);
    return add;
  };
  addCols(SH.tx, TX_HEAD);
  if (addCols(SH.cat, CAT_HEAD).indexOf('parent') >= 0) {
    const s = ss.getSheetByName(SH.cat);
    const names = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
    const rows = DEFAULT_SUBS.filter(r => names.indexOf(r[3]) >= 0);
    if (rows.length) s.getRange(s.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }
}

// Kalau sheet belum ada (setup belum dijalanin), bikin otomatis
function sheet_(name) {
  migrate_();
  let s = ss_().getSheetByName(name);
  if (!s) { setup(); s = ss_().getSheetByName(name); }
  return s;
}

function rows_(name) {
  const v = sheet_(name).getDataRange().getValues();
  const h = v.shift();
  const tz = Session.getScriptTimeZone();
  return v.filter(r => r[0] !== '').map(r => {
    const o = {};
    h.forEach((k, i) => {
      o[k] = r[i] instanceof Date ? Utilities.formatDate(r[i], tz, 'yyyy-MM-dd') : r[i];
    });
    return o;
  });
}

function findRow_(s, id) {
  if (!id) return 0;
  const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (ids[i][0] === id) return i + 1;
  return 0;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { fn(); } finally { lock.releaseLock(); }
  return getData();
}

// ---------- API (dipanggil dari frontend) ----------
function getData() {
  return { accounts: rows_(SH.acc), categories: rows_(SH.cat), transactions: rows_(SH.tx),
           lastDate: PropertiesService.getUserProperties().getProperty('lastDate') || '' };
}

function saveTransaction(t) {
  if (!t.id && t.date) PropertiesService.getUserProperties().setProperty('lastDate', String(t.date));
  return withLock_(() => {
    const s = sheet_(SH.tx);
    const row = [
      t.id || Utilities.getUuid(), String(t.date), t.type, Number(t.amount),
      t.account, t.type === 'transfer' ? t.to_account : '',
      t.type === 'transfer' ? '' : t.category, t.note || '', new Date().toISOString(),
      t.type === 'transfer' ? '' : (t.subcategory || ''),
    ];
    const r = findRow_(s, t.id);
    if (r) s.getRange(r, 1, 1, row.length).setValues([row]);
    else s.appendRow(row);
  });
}

function deleteTransaction(id) {
  return withLock_(() => {
    const s = sheet_(SH.tx);
    const r = findRow_(s, id);
    if (r) s.deleteRow(r);
  });
}

function addAccount(name, initial) {
  return withLock_(() => sheet_(SH.acc).appendRow([String(name).trim(), Number(initial) || 0]));
}

function addCategory(name, type, icon, parent) {
  return withLock_(() => sheet_(SH.cat).appendRow([String(name).trim(), type, icon || '🏷️', parent || '']));
}
