// Loads the monthly CSVs, cleans them and computes every metric used in the report.
const fs = require('fs'), path = require('path');

function parseCSV(text) {
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true; else if (c === ',') { row.push(f); f = ''; }
    else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(f); f = ''; rows.push(row); row = []; }
    else f += c;
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return rows.filter(r => r.length > 1 || r[0] !== '');
}

const SUBJECTS = ['Maths', 'Science', 'English', 'GK', 'Computer', 'Social', 'JEE', 'NEET'];
const TIERS = ['Flexi', 'Smart', 'Champion'];
const PLANS = [];
SUBJECTS.forEach(s => TIERS.forEach(t => PLANS.push({ name: `${s} ${t} Plan`, subject: s, tier: t })));

const normPlan = p => p.replace(/\s+/g, ' ').trim().replace(/ plan$/i, ' Plan');
// Flexi product names contain commas (e.g. "JEE Olympiad Exam – Physics, Chemistry & Mathematics"), so re-join those fragments.
function splitFlexi(s) {
  if (!s) return [];
  const out = [];
  s.split(', ').forEach(p => {
    if (/^(Chemistry|Mathematics|Biology)/.test(p) && out.length) out[out.length - 1] += ', ' + p;
    else out.push(p);
  });
  return out.map(p => p.replace(/\s+/g, ' ').replace(/^Social studies/, 'Social Studies').trim());
}

function testReason(r) {
  const e = r.email.toLowerCase(), n = r.name;
  if (/@itoexam\.com$/.test(e)) return 'Internal ITOUSER account (@itoexam.com)';
  if (/indiantalent/.test(e)) return 'Internal staff email (indiantalent)';
  if (/\btest(ing)?\b/i.test(n)) return 'Student name contains "test"';
  if (/^testing\d|accemail\.ito/.test(e)) return 'Test email address';
  if (/^[12]/.test(r.mobile)) return 'Invalid mobile (starts with 1/2) – dummy entry';
  return null;
}

function load(dir) {
  const raw = [];
  for (const file of fs.readdirSync(dir).filter(n => n.toLowerCase().endsWith('.csv'))) {
    const rows = parseCSV(fs.readFileSync(path.join(dir, file), 'utf8').replace(/^﻿/, ''));
    const h = rows[0];
    rows.slice(1).forEach(x => {
      const o = Object.fromEntries(h.map((k, i) => [k, (x[i] || '').trim()]));
      raw.push({
        userId: o['User ID'], name: o['Student Name'], email: o['Parent Email'], mobile: o.Mobile,
        txn: o['Transaction ID'], status: o['Payment Status'].toLowerCase(), date: o['Order Date'],
        plansBought: +o['Plans Bought'], cartRaw: o['Exact Cart Combination'], flexiRaw: o['Flexi Products Bought'], file,
      });
    });
  }
  raw.forEach(r => {
    r.month = r.date.slice(0, 7);
    r.plans = r.cartRaw.split(' + ').map(normPlan);
    r.cart = r.plans.join(' + ');
    r.flexi = splitFlexi(r.flexiRaw);
    const tiers = new Set(r.plans.map(p => p.split(' ')[1]));
    r.tierMix = tiers.size > 1 ? 'Mixed tiers' : [...tiers][0] + ' only';
    r.paid = r.status === 'captured';
    r.test = testReason(r);
  });
  return raw;
}

function compute(raw) {
  const clean = raw.filter(r => !r.test);
  const tests = raw.filter(r => r.test);
  const months = [...new Set(clean.map(r => r.month))].sort();
  // Did this student pay on or after the non-paid attempt (within the data)?
  const byUser = {};
  clean.forEach(r => (byUser[r.userId] ||= []).push(r));
  clean.forEach(r => {
    r.paidLater = r.paid ? '' : (byUser[r.userId].some(x => x.paid && x.date >= r.date) ? 'Y' : 'N');
  });
  // Same student, same cart, same day, not paid, already a previous identical row → repeat click
  const seen = new Set();
  [...clean].sort((a, b) => a.txn - b.txn).forEach(r => {
    const k = r.userId + '|' + r.cart + '|' + r.date;
    r.repeatSameDay = !r.paid && seen.has(k) ? 'Y' : '';
    seen.add(k);
  });

  const monthly = months.map(m => {
    const rs = clean.filter(r => r.month === m);
    const st = s => rs.filter(r => r.status === s).length;
    const users = new Set(rs.map(r => r.userId));
    const payers = new Set(rs.filter(r => r.paid).map(r => r.userId));
    const nonPayers = [...users].filter(u => !payers.has(u));
    const neverPaid = nonPayers.filter(u => !byUser[u].some(x => x.paid && x.month >= m));
    return {
      month: m, days: new Set(rs.map(r => r.date)).size, attempts: rs.length, paid: st('captured'), pending: st('pending'),
      attempted: st('attempted'), created: st('created'), failed: st('failed'),
      students: users.size, payers: payers.size, unpaidStudents: nonPayers.length, neverPaid: neverPaid.length,
      plansSold: rs.filter(r => r.paid).reduce((a, r) => a + r.plansBought, 0),
      multiPlanPaid: rs.filter(r => r.paid && r.plansBought > 1).length,
    };
  });

  // student-level funnel over whole period
  const users = Object.keys(byUser);
  const firstAttemptPaid = users.filter(u => byUser[u].slice().sort((a, b) => a.txn - b.txn)[0].paid);
  const payers = users.filter(u => byUser[u].some(r => r.paid));
  const funnel = {
    students: users.length, payers: payers.length, firstAttemptPaid: firstAttemptPaid.length,
    paidAfterRetry: payers.length - firstAttemptPaid.length, neverPaid: users.length - payers.length,
    neverPaidMultiAttempt: users.filter(u => !byUser[u].some(r => r.paid) && byUser[u].length > 1).length,
    repeatPayersMultiMonth: users.filter(u => new Set(byUser[u].filter(r => r.paid).map(r => r.month)).size > 1).length,
    multiPaidOrders: users.filter(u => byUser[u].filter(r => r.paid).length > 1).length,
  };

  const plans = PLANS.map(p => {
    const has = r => r.plans.includes(p.name);
    const o = { ...p, attempts: clean.filter(has).length, paid: clean.filter(r => r.paid && has(r)).length };
    months.forEach(m => o[m] = clean.filter(r => r.paid && r.month === m && has(r)).length);
    return o;
  }).filter(p => p.attempts > 0);

  const carts = {};
  clean.forEach(r => { carts[r.cart] ||= { cart: r.cart, attempts: 0, paid: 0 }; carts[r.cart].attempts++; if (r.paid) carts[r.cart].paid++; });
  const topCarts = Object.values(carts).sort((a, b) => b.paid - a.paid).slice(0, 15);

  const addons = {};
  clean.filter(r => r.paid).forEach(r => r.flexi.forEach(f => {
    const type = /Exam Round 1|Olympiad Exam –/.test(f) ? 'Olympiad Exam' : /Certificate/.test(f) ? 'Certificate'
      : /Previous Year/.test(f) ? 'ePrevious Year Papers' : /eWorkbook/.test(f) ? 'eWorkbook'
      : /Test Bank|Practice Test/.test(f) ? 'Online Test Bank / Practice' : /Monthly Olympiad/.test(f) ? 'Monthly Olympiad (1 yr)' : 'Other';
    addons[type] ||= { type, units: 0 }; addons[type].units++;
  }));
  const flexiPaid = clean.filter(r => r.paid && r.plans.some(p => p.includes('Flexi')));
  const flexiWithAddon = flexiPaid.filter(r => r.flexi.some(f => !/Exam Round 1|Olympiad Exam –/.test(f)));

  const tierMix = {};
  clean.forEach(r => { tierMix[r.tierMix] ||= { mix: r.tierMix, attempts: 0, paid: 0 }; tierMix[r.tierMix].attempts++; if (r.paid) tierMix[r.tierMix].paid++; });

  const daily = {};
  clean.forEach(r => { daily[r.date] ||= { date: r.date, attempts: 0, paid: 0 }; daily[r.date].attempts++; if (r.paid) daily[r.date].paid++; });

  return { raw, clean, tests, months, monthly, funnel, plans, topCarts, addons: Object.values(addons).sort((a, b) => b.units - a.units),
    flexiPaid: flexiPaid.length, flexiWithAddon: flexiWithAddon.length, tierMix: Object.values(tierMix), daily: Object.values(daily).sort((a, b) => a.date < b.date ? -1 : 1) };
}

module.exports = { load, compute, PLANS };

if (require.main === module) {
  const d = compute(load(process.argv[2]));
  console.log('raw', d.raw.length, 'clean', d.clean.length, 'tests', d.tests.length, 'test users', new Set(d.tests.map(r => r.userId)).size);
  const tr = {}; d.tests.forEach(r => tr[r.test] = (tr[r.test] || 0) + 1); console.log(tr);
  console.log('test paid rows', d.tests.filter(r => r.paid).length);
  console.table(d.monthly);
  console.log(d.funnel);
  console.table(d.plans.map(p => ({ name: p.name, att: p.attempts, paid: p.paid, sr: (p.paid / p.attempts * 100).toFixed(1), ...Object.fromEntries(d.months.map(m => [m, p[m]])) })));
  console.table(d.topCarts);
  console.table(d.addons); console.log('flexiPaid', d.flexiPaid, 'withAddon', d.flexiWithAddon);
  console.table(d.tierMix);
  const ds = d.daily.slice().sort((a, b) => b.attempts - a.attempts);
  console.log('top days', ds.slice(0, 8)); console.log('low days', ds.slice(-5));
  console.log('repeat same-day clicks', d.clean.filter(r => r.repeatSameDay).length);
  console.log('nonpaid paidLater Y', d.clean.filter(r => r.paidLater === 'Y').length, 'N', d.clean.filter(r => r.paidLater === 'N').length);
  // weekday pattern
  const wd = {}; d.clean.forEach(r => { const k = new Date(r.date + 'T00:00:00Z').getUTCDay(); wd[k] ||= [0, 0]; wd[k][0]++; if (r.paid) wd[k][1]++; }); console.log('weekday', wd);
  // october by week
  const ow = {}; d.clean.filter(r => r.month === '2026-10').forEach(r => { const k = Math.ceil(+r.date.slice(8) / 7); ow[k] ||= {}; ow[k][r.status] = (ow[k][r.status] || 0) + 1; }); console.log('oct weeks', ow);
}
