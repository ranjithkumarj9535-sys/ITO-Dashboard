// Runs the dashboard's own calcs (from dashboard/page.html) over the order rows and writes the results compactly,
// so the static site ships summaries only.
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const page = fs.readFileSync(path.join(root, 'dashboard', 'page.html'), 'utf8');
const calcScript = [...page.matchAll(/<script>([\s\S]*?)<\/script>/g)][0][1];
const calcs = [];
new Function('dash', calcScript)({ calc: (id, def) => calcs.push({ id, def }) });
const results = { orders: JSON.parse(fs.readFileSync(path.join(root, 'public', 'data', 'orders.json'), 'utf8')) };
const round = (v) => (typeof v === 'number' && !Number.isInteger(v) ? Math.round(v * 1e4) / 1e4 : v);
const out = {};
for (const { id, def } of calcs) {
  // Calcs can read other calcs, so run them in page order on the full-precision results.
  results[id] = def.fn(...def.inputs.map((i) => results[i]));
  const cols = Object.keys(results[id][0]);
  out[id] = { c: cols, r: results[id].map((r) => cols.map((c) => round(r[c]))) };
}
const file = path.join(root, 'public', 'data', 'summary.json');
fs.writeFileSync(file, JSON.stringify(out));
console.log('wrote', Object.keys(out).join(', '), fs.statSync(file).size, 'bytes');
