// Runs the dashboard's own calcs (from dashboard/page.html) over the order rows and writes the results compactly,
// so the static site can ship summaries only.
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const page = fs.readFileSync(path.join(root, 'dashboard', 'page.html'), 'utf8');
const calcScript = [...page.matchAll(/<script>([\s\S]*?)<\/script>/g)][0][1];
const calcs = {};
new Function('dash', calcScript)({ calc: (id, def) => { calcs[id] = def; } });
const orders = JSON.parse(fs.readFileSync(path.join(root, 'public', 'data', 'orders.json'), 'utf8'));
const round = (v) => (typeof v === 'number' && !Number.isInteger(v) ? Math.round(v * 1e4) / 1e4 : v);
const out = {};
for (const [id, def] of Object.entries(calcs)) {
  // Rows with no orders carry no information; the page treats a missing row as empty.
  const rows = def.fn(orders).filter((r) => r.attempts > 0);
  const cols = Object.keys(rows[0]).filter((c) => c !== 'k');
  out[id] = { c: cols, r: rows.map((r) => cols.map((c) => round(r[c]))) };
}
const file = path.join(root, 'public', 'data', 'summary.json');
fs.writeFileSync(file, JSON.stringify(out));
console.log('wrote', fs.statSync(file).size, 'bytes');
