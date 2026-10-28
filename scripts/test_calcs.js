// Runs every dash.calc in dashboard/page.html over public/data/orders.json and checks the totals against the Excel report.
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const page = fs.readFileSync(path.join(root, 'dashboard', 'page.html'), 'utf8');
const scripts = [...page.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]);
scripts.forEach((s, i) => { try { new Function(s); } catch (e) { console.log('SYNTAX script', i, e.message); process.exitCode = 1; } });
new Function(fs.readFileSync(path.join(root, 'dashboard', 'reports.js'), 'utf8'));

const calcs = [];
new Function('dash', scripts[0])({ calc: (id, def) => calcs.push({ id, def }) });
const results = { orders: JSON.parse(fs.readFileSync(path.join(root, 'public', 'data', 'orders.json'), 'utf8')) };
for (const { id, def } of calcs) {
  const t = Date.now();
  results[id] = def.fn(...def.inputs.map((i) => results[i]));
  console.log(id.padEnd(11), String(results[id].length).padStart(5), 'rows', Date.now() - t, 'ms');
}
const a = results.cube.find((r) => r.k === 'all|all|all');
const expect = { attempts: 3458, paid: 1915, students: 2286, payers: 1732, never_paid: 554, plans_sold: 3109, recovered: 771, lost: 772, repeat_clicks: 455, first_try: 1378 };
for (const [k, v] of Object.entries(expect)) if (a[k] !== v) { console.log('MISMATCH', k, a[k], 'expected', v); process.exitCode = 1; }
const dailyPaid = results.daily.filter((r) => r.subject === 'all' && r.tier === 'all').reduce((s, r) => s + r.paid, 0);
const weeklyPaid = results.weekly.filter((r) => r.subject === 'all' && r.tier === 'all').reduce((s, r) => s + r.captured, 0);
const wdPaid = results.weekday.filter((r) => r.k.startsWith('all|all|all|')).reduce((s, r) => s + r.paid, 0);
const retryStudents = results.retry.filter((r) => r.k.startsWith('all|all|all|')).reduce((s, r) => s + r.students, 0);
const comboPaid = results.combos.filter((r) => r.month === 'all').reduce((s, r) => s + r.paid, 0);
console.log('daily paid', dailyPaid, 'weekly paid', weeklyPaid, 'weekday paid', wdPaid, 'retry students', retryStudents, 'combo paid', comboPaid);
[dailyPaid, weeklyPaid, wdPaid, comboPaid].forEach((v) => { if (v !== 1915) { console.log('MISMATCH total paid', v); process.exitCode = 1; } });
if (retryStudents !== 2286) { console.log('MISMATCH retry students'); process.exitCode = 1; }
console.log('highlights all', JSON.stringify(results.highlights.find((r) => r.k === 'all|all')));
console.log('addons all', results.addons.filter((r) => r.k.startsWith('all|all|')).map((r) => r.addon + ' ' + (r.attach * 100).toFixed(1) + '%').join(', '));
console.log(process.exitCode ? 'FAILED' : 'ALL CHECKS PASSED');
module.exports = results;
