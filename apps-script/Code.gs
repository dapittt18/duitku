// ===== Duitku — backend (Google Apps Script, bound ke Google Sheet) =====
const SH = { tx: 'Transactions', acc: 'Accounts', cat: 'Categories' };
const TX_HEAD = ['id', 'date', 'type', 'amount', 'account', 'to_account', 'category', 'note', 'created_at', 'subcategory', 'source'];
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

// ---------- Auto-catat dari email BCA (trigger tiap 5 menit) ----------
// Sekali jalan dari editor: testBcaEmails() buat cek hasil baca (gak nulis apa-apa), lalu installBcaTrigger().
const BCA_ACCOUNT = 'BCA';
const BCA_SINCE = '2026-10-26T00:00:00';   // email sebelum ini gak dicatat (udah masuk dari import)
const NEED_CHECK = 'Perlu dicek';           // kategori buat merchant yang belum dikenal
// [pola merchant, kategori, sub kategori, catatan]
const MERCHANT_RULES = [
  [/^grab/i, 'Transportation', 'Ojek Online', 'Grab'],
  [/gojek|goride|gocar/i, 'Transportation', 'Ojek Online', 'Gojek'],
  [/indomaret|indomart/i, 'Belanja', 'Indomart', 'Indomaret'],
  [/shopee/i, 'Belanja', 'Shopee', 'Shopee'],
  [/blibli/i, 'Belanja', 'Blibli', 'Blibli'],
  [/spotify/i, 'Tagihan', 'Spotify', 'Spotify'],
  [/netflix/i, 'Tagihan', 'Netflix', 'Netflix'],
  [/vidio/i, 'Tagihan', 'Vidio', 'Vidio'],
];

// Email "Informasi Transaksi Online - Debit BCA [BERHASIL]" → {date, created_at, amount, category, subcategory, note}
function parseBcaDebit_(text) {
  const f = k => { const m = text.match(new RegExp(k + '\\s*:\\s*([^\\n\\r]+)', 'i')); return m ? m[1].trim() : ''; };
  const merchant = f('Merchant'), w = f('Waktu').match(/(\d{2})-(\d{2})-(\d{4})\s+(\d{2}:\d{2}:\d{2})/);
  const amount = Number(f('Nilai Transaksi').replace(/^[A-Z]{3}\s*/i, '').replace(/,\d{1,2}$/, '').replace(/\D/g, ''));
  if (!merchant || !w || !amount) return null;
  const rule = MERCHANT_RULES.find(r => r[0].test(merchant));
  const date = w[3] + '-' + w[2] + '-' + w[1];
  return { date, created_at: date + 'T' + w[4], amount, merchant,
           category: rule ? rule[1] : NEED_CHECK, subcategory: rule ? rule[2] : '', note: rule ? rule[3] : merchant };
}

function htmlText_(html) {
  return html.replace(/<br\s*\/?>|<\/(p|div|tr|td|th|li)>/gi, '\n').replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/[ \t]+/g, ' ');
}

// Logika murni (tanpa Gmail/Sheet). items: [{id, refund, t}] urut waktu; existing: [{row, source, amount}] dari sheet.
// Debit → transaksi baru. Pengembalian dana → debit dengan merchant + nominal sama gak dicatat (atau dihapus kalau
// udah terlanjur). Dihitung ulang tiap jalan, jadi aman walau debit & refund kebaca di run yang beda.
// Refund yang debitnya gak ketemu sama sekali dicatat jadi pemasukan biar saldo tetap bener.
function planBcaSync_(items, existing, refundsDone) {
  const idOf = src => String(src).split('|')[0].replace(/^gmail:/, '');
  const merchantOf = src => { src = String(src); return src.indexOf('|') < 0 ? '' : src.slice(src.indexOf('|') + 1); };
  const seen = {}; existing.forEach(r => { seen[idOf(r.source)] = true; });
  const debits = items.filter(x => !x.refund), cancel = {};
  const add = [], del = [], refunds = [], log = [];
  const tag = t => t.created_at + '  Rp' + t.amount + '  ' + t.merchant;
  items.filter(x => x.refund).forEach(({ id, t }) => {
    const d = debits.filter(x => !cancel[x.id] && x.t.merchant === t.merchant && x.t.amount === t.amount && x.t.created_at <= t.created_at).pop();
    if (d) { cancel[d.id] = true; return; }
    if (seen[id] || refundsDone.indexOf(id) >= 0) return;
    const ex = existing.find(r => merchantOf(r.source) === t.merchant && Number(r.amount) === t.amount && del.indexOf(r.row) < 0);
    if (ex) { del.push(ex.row); refunds.push(id); log.push(tag(t) + '  ↩︎ refund  →  debitnya dihapus'); return; }
    if (t.created_at < BCA_SINCE) return;
    add.push([Utilities.getUuid(), t.date, 'income', t.amount, BCA_ACCOUNT, '', 'Lainnya', 'Refund ' + t.note, t.created_at, '', 'gmail:' + id + '|' + t.merchant]);
    log.push(tag(t) + '  ↩︎ refund tanpa pasangan  →  pemasukan Lainnya');
  });
  debits.forEach(({ id, t }) => {
    if (cancel[id]) {
      const ex = existing.find(r => idOf(r.source) === id);
      if (ex && del.indexOf(ex.row) < 0) { del.push(ex.row); log.push(tag(t) + '  ✗ di-refund  →  dihapus dari catatan'); }
      else if (!ex) log.push(tag(t) + '  ✗ di-refund  →  gak dicatat');
      return;
    }
    if (seen[id] || t.created_at < BCA_SINCE) return;
    add.push([Utilities.getUuid(), t.date, 'expense', t.amount, BCA_ACCOUNT, '', t.category, t.note, t.created_at, t.subcategory, 'gmail:' + id + '|' + t.merchant]);
    log.push(tag(t) + '  →  ' + t.category + (t.subcategory ? ' › ' + t.subcategory : ''));
  });
  return { add, del, refunds, log };
}

function syncBcaEmails(dryRun) {
  dryRun = dryRun === true; // dari trigger, argumennya event object
  // 7 hari terakhir: cukup buat nangkep refund yang telat, dan debit pasangannya ikut kebaca
  const q = 'from:PasporBCA@klikbca.com (subject:BERHASIL OR subject:"Pengembalian Dana") newer_than:7d';
  const items = [], log = [];
  GmailApp.search(q, 0, 50).forEach(th => th.getMessages().forEach(m => {
    const t = parseBcaDebit_(m.getPlainBody()) || parseBcaDebit_(htmlText_(m.getBody()));
    if (!t) return log.push('⚠️ gak kebaca: ' + m.getSubject() + ' (' + m.getDate() + ')');
    items.push({ id: m.getId(), refund: /Pengembalian Dana/i.test(m.getSubject()), t });
  }));
  items.sort((a, b) => a.t.created_at.localeCompare(b.t.created_at));

  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_(SH.tx), v = s.getDataRange().getValues(), h = v[0];
    const iS = h.indexOf('source'), iA = h.indexOf('amount');
    const existing = v.slice(1).map((r, k) => ({ row: k + 2, source: r[iS], amount: r[iA] })).filter(r => r.source);
    const done = JSON.parse(props.getProperty('BCA_REFUNDS_DONE') || '[]');
    const p = planBcaSync_(items, existing, done);
    log.push.apply(log, p.log);
    if (!dryRun) {
      if (p.add.some(r => r[6] === NEED_CHECK) && !rows_(SH.cat).some(c => c.name === NEED_CHECK && c.type === 'expense'))
        sheet_(SH.cat).appendRow([NEED_CHECK, 'expense', '❓', '']);
      p.del.sort((a, b) => b - a).forEach(r => s.deleteRow(r));
      if (p.add.length) {
        const start = s.getLastRow() + 1;
        s.getRange(start, 2, p.add.length, 1).setNumberFormat('@');
        s.getRange(start, 9, p.add.length, 1).setNumberFormat('@');
        s.getRange(start, 1, p.add.length, p.add[0].length).setValues(p.add);
      }
      if (p.refunds.length) props.setProperty('BCA_REFUNDS_DONE', JSON.stringify(done.concat(p.refunds).slice(-300)));
    }
  } finally { lock.releaseLock(); }
  Logger.log((dryRun ? '[TES, gak ada yang dicatat]\n' : '') + (log.join('\n') || 'Gak ada perubahan.'));
  return log;
}

function testBcaEmails() { return syncBcaEmails(true); }

function installBcaTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncBcaEmails').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncBcaEmails').timeBased().everyMinutes(5).create();
  sheet_(SH.tx).getRange('B:B').setNumberFormat('@'); // tanggal tetap teks yyyy-MM-dd
  sheet_(SH.tx).getRange('I:I').setNumberFormat('@');
  syncBcaEmails();
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
