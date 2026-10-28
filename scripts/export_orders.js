// Writes the cleaned orders (no names, emails or phone numbers) as JSON for the dashboard.
const fs = require('fs');
const { load, compute } = require('./analyze');
const [dir, out] = process.argv.slice(2);
const d = compute(load(dir));
const rows = d.clean
  .slice().sort((a, b) => a.txn - b.txn)
  .map(r => ({ txn: r.txn, student: r.userId, date: r.date, month: r.month, status: r.status, paid: r.paid ? 1 : 0, plans_bought: r.plansBought, cart: r.cart, cart_type: r.tierMix, addons: r.flexi.join(' ; ') }));
fs.writeFileSync(out, JSON.stringify(rows));
console.log(rows.length, 'rows', fs.statSync(out).size, 'bytes');
