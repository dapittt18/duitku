// Mock backend buat test lokal: serve folder ini + tiruan doPost Apps Script (data di memori).
// Jalanin: node dev/mock-server.js  → buka http://localhost:8787  (token: test123)
const http = require('http'), fs = require('fs'), path = require('path'), crypto = require('crypto');
const ROOT = path.join(__dirname, '..'), PORT = process.env.PORT || 8787, TOKEN = 'test123';
const TYPES = {'.html': 'text/html', '.js': 'text/javascript', '.json': 'application/json', '.png': 'image/png'};

const d = n => { const x = new Date(); x.setDate(x.getDate() - n); return x.toISOString().slice(0, 10); };
const DB = {
  accounts: [{name: 'Cash', initial_balance: 500000}, {name: 'BCA', initial_balance: 5000000}, {name: 'GoPay', initial_balance: 150000}],
  categories: [['Makan', 'expense', '🍜'], ['Transport', 'expense', '🛵'], ['Belanja', 'expense', '🛒'], ['Tagihan', 'expense', '💡'],
    ['Hiburan', 'expense', '🎮'], ['Lainnya', 'expense', '📦'], ['Gaji', 'income', '💼'], ['Bonus', 'income', '🎁'],
    ['Wifi', 'expense', '📶', 'Tagihan'], ['Listrik', 'expense', '⚡', 'Tagihan'], ['Ojol', 'expense', '🛵', 'Transport']]
    .map(([name, type, icon, parent]) => ({name, type, icon, parent: parent || ''})),
  transactions: [
    {id: 't1', date: d(0), type: 'expense', amount: 35000, account: 'Cash', to_account: '', category: 'Makan', subcategory: '', note: 'Bakmi', created_at: new Date().toISOString()},
    {id: 't2', date: d(1), type: 'expense', amount: 350000, account: 'BCA', to_account: '', category: 'Tagihan', subcategory: 'Wifi', note: '', created_at: new Date().toISOString()},
    {id: 't3', date: d(2), type: 'income', amount: 8000000, account: 'BCA', to_account: '', category: 'Gaji', subcategory: '', note: '', created_at: new Date().toISOString()},
    {id: 't4', date: d(2), type: 'transfer', amount: 200000, account: 'BCA', to_account: 'GoPay', category: '', subcategory: '', note: '', created_at: new Date().toISOString()},
  ],
};
let lastDate = '';
const getData = () => ({...DB, lastDate});
const API = {
  getData,
  saveTransaction(t) {
    if (!t.id && t.date) lastDate = String(t.date);
    const tr = t.type === 'transfer';
    const row = {id: t.id || crypto.randomUUID(), date: String(t.date), type: t.type, amount: Number(t.amount), account: t.account,
      to_account: tr ? t.to_account : '', category: tr ? '' : t.category, note: t.note || '', created_at: new Date().toISOString(), subcategory: tr ? '' : (t.subcategory || '')};
    const i = DB.transactions.findIndex(x => x.id === t.id);
    if (i >= 0) DB.transactions[i] = row; else DB.transactions.push(row);
    return getData();
  },
  deleteTransaction(id) { DB.transactions = DB.transactions.filter(x => x.id !== id); return getData(); },
  addAccount(name, initial) { DB.accounts.push({name: String(name).trim(), initial_balance: Number(initial) || 0}); return getData(); },
  addCategory(name, type, icon, parent) { DB.categories.push({name: String(name).trim(), type, icon: icon || '🏷️', parent: parent || ''}); return getData(); },
};

http.createServer((req, res) => {
  const url = new URL(req.url, 'http://x');
  if (req.method === 'POST' && url.pathname === '/mock-api') {
    let body = ''; req.on('data', c => body += c); req.on('end', () => {
      let out;
      try {
        const r = JSON.parse(body || '{}');
        if (r.token !== TOKEN) throw new Error('unauthorized');
        if (!API[r.fn]) throw new Error('unknown fn: ' + r.fn);
        out = {ok: true, data: API[r.fn](...(r.args || []))};
      } catch (e) { out = {ok: false, error: e.message}; }
      console.log('API', req.headers['content-type'], body.slice(0, 80), '→', out.ok ? 'ok' : out.error);
      res.writeHead(200, {'Content-Type': 'application/json'}); res.end(JSON.stringify(out));
    });
    return;
  }
  if (url.pathname === '/config.js') { res.writeHead(200, {'Content-Type': 'text/javascript'}); return res.end("const API_URL = '/mock-api';"); }
  const f = path.join(ROOT, url.pathname === '/' ? 'index.html' : path.normalize(url.pathname));
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.writeHead(404); return res.end('404'); }
  res.writeHead(200, {'Content-Type': TYPES[path.extname(f)] || 'application/octet-stream'}); fs.createReadStream(f).pipe(res);
}).listen(PORT, () => console.log('Duitku mock → http://localhost:' + PORT + '  (token: test123)'));
