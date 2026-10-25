// Generate icon PNG (dompet + koin, gradient #5b4bff→#ec4899) tanpa dependency. Jalanin: node dev/make-icons.js
const fs = require('fs'), zlib = require('zlib'), path = require('path');
const OUT = path.join(__dirname, '..', 'icons');

const crcT = new Uint32Array(256).map((_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
const crc = b => { let c = ~0; for (const x of b) c = crcT[(c ^ x) & 255] ^ (c >>> 8); return ~c >>> 0; };
const chunk = (t, d) => { const l = Buffer.alloc(4); l.writeUInt32BE(d.length); const td = Buffer.concat([Buffer.from(t), d]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([l, td, c]); };
function png(w, h, rgba) {
  const raw = Buffer.alloc((w * 4 + 1) * h);
  for (let y = 0; y < h; y++) { raw[y * (w * 4 + 1)] = 0; rgba.copy(raw, y * (w * 4 + 1) + 1, y * w * 4, (y + 1) * w * 4); }
  const ih = Buffer.alloc(13); ih.writeUInt32BE(w, 0); ih.writeUInt32BE(h, 4); ih[8] = 8; ih[9] = 6;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ih), chunk('IDAT', zlib.deflateSync(raw, {level: 9})), chunk('IEND', Buffer.alloc(0))]);
}
const hex = h => [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16));
const A = hex('#5b4bff'), B = hex('#ec4899');
const rrect = (x, y, x0, y0, x1, y1, r) => { const cx = Math.max(x0 + r, Math.min(x, x1 - r)), cy = Math.max(y0 + r, Math.min(y, y1 - r)); return Math.hypot(x - cx, y - cy) <= r && x >= x0 && x <= x1 && y >= y0 && y <= y1; };
const circ = (x, y, cx, cy, r) => Math.hypot(x - cx, y - cy) <= r;

// warna 1 titik (koordinat 0..1). Konten masuk safe zone maskable (radius 0.4 dari tengah).
function sample(x, y, rounded) {
  if (rounded && !rrect(x, y, 0, 0, 1, 1, 0.225)) return null;
  const t = Math.min(1, Math.max(0, (x + y) / 2));
  let c = A.map((a, i) => a + (B[i] - a) * t);
  const lerp = (p, q, k) => p.map((v, i) => v + (q[i] - v) * k);
  // sinar lembut kanan atas
  if (circ(x, y, 0.92, 0.08, 0.42)) c = lerp(c, [255, 255, 255], 0.10);
  // koin di belakang dompet
  if (circ(x, y, 0.60, 0.34, 0.135)) {
    c = circ(x, y, 0.60, 0.34, 0.095) ? [255, 200, 87] : [255, 222, 140];
    if (rrect(x, y, 0.585, 0.29, 0.615, 0.39, 0.012)) c = [240, 170, 50];
  }
  // dompet
  if (rrect(x, y, 0.21, 0.40, 0.79, 0.76, 0.07)) {
    c = [255, 255, 255];
    if (rrect(x, y, 0.21, 0.40, 0.79, 0.47, 0.035) && y < 0.47) c = [236, 233, 255];
    if (rrect(x, y, 0.56, 0.52, 0.84, 0.66, 0.045)) {
      c = lerp(A, B, 0.45);
      if (circ(x, y, 0.645, 0.59, 0.026)) c = [255, 255, 255];
    }
  }
  return c;
}
function draw(size, rounded) {
  const S = 4, buf = Buffer.alloc(size * size * 4);
  for (let py = 0; py < size; py++) for (let px = 0; px < size; px++) {
    let r = 0, g = 0, b = 0, a = 0;
    for (let sy = 0; sy < S; sy++) for (let sx = 0; sx < S; sx++) {
      const c = sample((px + (sx + .5) / S) / size, (py + (sy + .5) / S) / size, rounded);
      if (c) { r += c[0]; g += c[1]; b += c[2]; a++; }
    }
    const o = (py * size + px) * 4;
    if (a) { buf[o] = r / a; buf[o + 1] = g / a; buf[o + 2] = b / a; buf[o + 3] = 255 * a / (S * S); }
  }
  return png(size, size, buf);
}
fs.mkdirSync(OUT, {recursive: true});
for (const [f, s, r] of [['icon-192.png', 192, 1], ['icon-512.png', 512, 1], ['maskable-192.png', 192, 0], ['maskable-512.png', 512, 0], ['apple-touch-icon.png', 180, 0]])
  fs.writeFileSync(path.join(OUT, f), draw(s, r));
console.log('icons ok');
