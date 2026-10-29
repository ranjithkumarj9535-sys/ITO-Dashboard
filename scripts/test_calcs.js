// Runs every dash.calc in dashboard/page.html over public/data/orders.json and checks the totals against the Excel report.
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const page = fs.readFileSync(path.join(root, 'dashboard', 'page.html'), 'utf8');
const scripts = [...page.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]);
scripts.forEach((s, i) => { try { new Function(s); } catch (e) { console.log('SYNTAX script', i, e.message); process.exitCode = 1; } });
new Function(fs.readFileSync(path.join(root, 'dashboard', 'reports.js'), 'utf8'));
new Function(fs.readFileSync(path.join(root, 'dashboard', 'departments.js'), 'utf8'));

const calcs = [];
new Function('dash', scripts[0])({ calc: (id, def) => calcs.push({ id, def }) });
const read = (name) => JSON.parse(fs.readFileSync(path.join(root, 'public', 'data', name + '.json'), 'utf8'));
const results = { orders: read('orders'), dispatch: read('dispatch'), dispatch_material: read('dispatch_material') };
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
// Dispatch: totals must match the workbook's own "Summary Dispatch" sheet (shipped, delivered, not delivered per courier).
const dc = (k) => results.dsp_cube.find((r) => r.k === k);
const sheet = { all: [2487, 2369, 118], Speed: [789, 755, 34], DTDC: [1322, 1276, 46], PRO: [146, 117, 29], Maruti: [230, 221, 9] };
for (const [c, [s, d, t]] of Object.entries(sheet)) {
  const r = dc('all|' + c + '|all');
  if (!r || r.shipped !== s || r.delivered !== d || r.in_transit !== t) { console.log('MISMATCH dispatch', c, r && [r.shipped, r.delivered, r.in_transit], 'expected', [s, d, t]); process.exitCode = 1; }
}
const da = dc('all|all|all');
console.log('dispatch all', JSON.stringify({ orders: da.orders, shipped: da.shipped, delivered: da.delivered, in_transit: da.in_transit, late: da.transit_late, on_hold: da.on_hold, not_shipped: da.not_shipped, avg_days: da.avg_days, within_7: da.within_7, issues: da.issues, as_of: da.as_of }));
const sumBy = (rows, f) => rows.reduce((a, r) => a + r[f], 0);
const monthsTotal = sumBy(results.dsp_cube.filter((r) => r.exam === 'all' && r.courier === 'all' && r.month !== 'all'), 'orders');
const regionsTotal = sumBy(results.dsp_regions.filter((r) => r.k.startsWith('all|all|all|')), 'orders');
const examTotal = dc('ITO|all|all').orders + dc('LC|all|all').orders;
const daysTotal = sumBy(results.dsp_days.filter((r) => r.k.startsWith('all|all|all|')), 'delivered');
const shippedDaily = sumBy(results.dsp_daily.filter((r) => r.exam === 'all' && r.courier === 'all'), 'shipped');
const openTotal = results.dsp_open.filter((r) => r.stage !== 'Delivered').length;
console.log('dispatch months', monthsTotal, 'regions', regionsTotal, 'ITO+LC', examTotal, 'days bands', daysTotal, 'shipped per day', shippedDaily, 'open', openTotal);
[monthsTotal, regionsTotal, examTotal].forEach((v) => { if (v !== da.orders) { console.log('MISMATCH dispatch orders', v); process.exitCode = 1; } });
if (openTotal !== da.orders - da.delivered) { console.log('MISMATCH open list'); process.exitCode = 1; }
console.log(process.exitCode ? 'FAILED' : 'ALL CHECKS PASSED');
module.exports = results;
