# Duitku

Money manager pribadi. PWA di GitHub Pages, data di Google Sheet lewat Apps Script (`apps-script/Code.gs`).

- `index.html` — UI (sama persis dengan versi Apps Script, `call()` pakai `fetch`)
- `config.js` — URL Web App Apps Script (`API_URL`)
- `manifest.json`, `sw.js`, `icons/` — PWA
- `apps-script/Code.gs` — backend; paste ke editor Apps Script
- `dev/mock-server.js` — test lokal tanpa Google: `node dev/mock-server.js` → http://localhost:8787 (token `test123`)
- `dev/make-icons.js` — generate ulang icon

Setelah ubah file app shell, naikin `VERSION` di `sw.js` biar HP ambil versi baru.
