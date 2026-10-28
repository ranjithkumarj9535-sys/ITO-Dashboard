// Packs public/data/orders.json into a compact columnar file (repeated text stored once) for the static site.
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const rows = JSON.parse(fs.readFileSync(path.join(root, 'public', 'data', 'orders.json'), 'utf8'));
const dicts = { status: [], cart: [], cart_type: [] };
const idx = (k, v) => { let i = dicts[k].indexOf(v); if (i < 0) { i = dicts[k].length; dicts[k].push(v); } return i; };
const packed = {
  columns: ['txn', 'student', 'date', 'status', 'paid', 'plans_bought', 'cart', 'cart_type'],
  dicts,
  rows: rows.map((r) => [+r.txn, +r.student, r.date.slice(5).replace('-', ''), idx('status', r.status), r.paid, r.plans_bought, idx('cart', r.cart), idx('cart_type', r.cart_type)]),
};
// gzip so the deploy stays small; the site unzips it in the browser (DecompressionStream).
const out = path.join(root, 'public', 'data', 'orders.bin');
fs.writeFileSync(out, require('zlib').gzipSync(JSON.stringify(packed), { level: 9 }));
console.log('packed', fs.statSync(out).size, 'bytes');
