// Builds the Excel report from the cleaned data. Formula cells carry a pre-computed result so they display before Excel recalculates.
const ExcelJS = require('exceljs');
const { load, compute } = require('./analyze');
const [dir, out] = process.argv.slice(2);
const d = compute(load(dir));
const { clean, months, monthly, funnel } = d;

const MONTH_NAME = { '2026-06': 'Jun 2026', '2026-07': 'Jul 2026', '2026-08': 'Aug 2026', '2026-09': 'Sep 2026', '2026-10': 'Oct 2026 (1–27)' };
const maskMobile = m => m.length >= 6 ? m.slice(0, 2) + 'X'.repeat(m.length - 4) + m.slice(-2) : 'XXXX';
const maskEmail = e => { if (!e) return ''; const [u, dom] = e.split('@'); return u.slice(0, 2) + '***@' + (dom || ''); };
const initials = n => n.split(/\s+/).filter(Boolean).map(w => w[0].toUpperCase() + '.').join('');

const wb = new ExcelJS.Workbook();
wb.creator = 'Data analyst report';
wb.calcProperties.fullCalcOnLoad = true;
const NAVY = 'FF1F3864', LIGHT = 'FFDCE6F1', GREY = 'FFF2F2F2';
const PCT = '0.0%', INT = '#,##0', DEC = '0.00';

function header(ws, rowNum, cols) {
  const row = ws.getRow(rowNum);
  cols.forEach((c, i) => { row.getCell(i + 1).value = c; });
  row.eachCell(c => { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; c.alignment = { wrapText: true, vertical: 'middle' }; });
  row.height = 32;
}
function title(ws, text, sub) {
  ws.getCell('A1').value = text; ws.getCell('A1').font = { bold: true, size: 14, color: { argb: NAVY } };
  if (sub) { ws.getCell('A2').value = sub; ws.getCell('A2').font = { italic: true, color: { argb: 'FF595959' } }; }
}
function totalStyle(row) { row.eachCell(c => { c.font = { bold: true }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT } }; }); }
const F = (formula, result) => ({ formula, result });
const su = wb.addWorksheet('Summary'); // created first so it is the first tab; filled at the end

// ---------- Clean_Data ----------
const cd = wb.addWorksheet('Clean_Data', { views: [{ state: 'frozen', ySplit: 1 }] });
const cdCols = ['Transaction ID', 'User ID', 'Student (initials)', 'Parent Email (masked)', 'Mobile (masked)', 'Order Date', 'Month', 'Payment Status',
  'Plans Bought', 'Cart (normalised)', 'Tier Mix', 'Flexi Products (normalised)', 'Unpaid → same student paid on/after this date', 'Repeat same-day unpaid click', 'Source file'];
header(cd, 1, cdCols);
const sorted = clean.slice().sort((a, b) => a.date === b.date ? a.txn - b.txn : a.date < b.date ? -1 : 1);
sorted.forEach((r, i) => {
  cd.getRow(i + 2).values = [+r.txn, +r.userId, initials(r.name), maskEmail(r.email), maskMobile(r.mobile), new Date(r.date + 'T00:00:00Z'), r.month, r.status,
    r.plansBought, r.cart, r.tierMix, r.flexi.join(', '), r.paidLater, r.repeatSameDay, r.file];
});
cd.getColumn(6).numFmt = 'dd-mmm-yyyy';
[12, 10, 12, 24, 14, 13, 10, 12, 8, 60, 13, 60, 16, 14, 18].forEach((w, i) => cd.getColumn(i + 1).width = w);
cd.autoFilter = { from: 'A1', to: 'O1' };
const N = clean.length + 1;
const R = col => `Clean_Data!$${col}$2:$${col}$${N}`;
const MON = R('G'), ST = R('H'), PB = R('I'), CART = R('J'), MIX = R('K'), FLX = R('L'), LATER = R('M'), REP = R('N');

// ---------- Monthly ----------
const mo = wb.addWorksheet('Monthly', { views: [{ state: 'frozen', xSplit: 1, ySplit: 4 }] });
title(mo, 'Monthly performance – orders, payments and students', 'Counts are live COUNTIFS on Clean_Data (test accounts excluded). Unique-student columns are pre-computed values (marked *).');
header(mo, 4, ['Month', 'Month key', 'Days with orders', 'Order attempts', 'Paid orders (captured)', 'Pending', 'Attempted', 'Created', 'Failed', 'Not paid (total)',
  'Payment success %', 'Paid orders / day', 'MoM change in paid orders', 'Students starting checkout *', 'Paying students *', 'Student conversion %',
  'Students who never paid *', 'Plans sold (paid)', 'Plans per paid order', 'Multi-plan paid orders', 'Multi-plan share']);
monthly.forEach((m, i) => {
  const r = 5 + i, row = mo.getRow(r);
  const prev = i ? monthly[i - 1] : null;
  row.values = [MONTH_NAME[m.month], m.month, m.days,
    F(`COUNTIFS(${MON},B${r})`, m.attempts),
    F(`COUNTIFS(${MON},B${r},${ST},"captured")`, m.paid),
    F(`COUNTIFS(${MON},B${r},${ST},"pending")`, m.pending),
    F(`COUNTIFS(${MON},B${r},${ST},"attempted")`, m.attempted),
    F(`COUNTIFS(${MON},B${r},${ST},"created")`, m.created),
    F(`COUNTIFS(${MON},B${r},${ST},"failed")`, m.failed),
    F(`D${r}-E${r}`, m.attempts - m.paid),
    F(`E${r}/D${r}`, m.paid / m.attempts),
    F(`E${r}/C${r}`, m.paid / m.days),
    prev ? F(`E${r}/E${r - 1}-1`, m.paid / prev.paid - 1) : 'n/a',
    m.students, m.payers, F(`O${r}/N${r}`, m.payers / m.students), m.neverPaid,
    F(`SUMIFS(${PB},${MON},B${r},${ST},"captured")`, m.plansSold),
    F(`R${r}/E${r}`, m.plansSold / m.paid),
    F(`COUNTIFS(${MON},B${r},${ST},"captured",${PB},">1")`, m.multiPlanPaid),
    F(`T${r}/E${r}`, m.multiPlanPaid / m.paid)];
});
const tr = 5 + monthly.length, last = tr - 1;
const sum = k => monthly.reduce((a, m) => a + m[k], 0);
const tot = { attempts: sum('attempts'), paid: sum('paid'), days: sum('days'), plansSold: sum('plansSold'), multi: sum('multiPlanPaid') };
mo.getRow(tr).values = ['Total (Jun 1 – Oct 27)', '', F(`SUM(C5:C${last})`, tot.days),
  ...['D', 'E', 'F', 'G', 'H', 'I', 'J'].map(c => F(`SUM(${c}5:${c}${last})`, { D: tot.attempts, E: tot.paid, F: sum('pending'), G: sum('attempted'), H: sum('created'), I: sum('failed'), J: tot.attempts - tot.paid }[c])),
  F(`E${tr}/D${tr}`, tot.paid / tot.attempts), F(`E${tr}/C${tr}`, tot.paid / tot.days), '',
  funnel.students, funnel.payers, F(`O${tr}/N${tr}`, funnel.payers / funnel.students), funnel.neverPaid,
  F(`SUM(R5:R${last})`, tot.plansSold), F(`R${tr}/E${tr}`, tot.plansSold / tot.paid), F(`SUM(T5:T${last})`, tot.multi), F(`T${tr}/E${tr}`, tot.multi / tot.paid)];
totalStyle(mo.getRow(tr));
mo.getCell(`A${tr + 1}`).value = 'Check: total paid orders in Clean_Data'; mo.getCell(`E${tr + 1}`).value = F(`COUNTIFS(${ST},"captured")`, tot.paid);
mo.getCell(`A${tr + 2}`).value = 'Check: monthly total = Clean_Data total?'; mo.getCell(`E${tr + 2}`).value = F(`IF(E${tr}=E${tr + 1},"OK","MISMATCH")`, 'OK');
mo.getCell(`A${tr + 4}`).value = '* Unique-student totals are de-duplicated across the whole period, so they are smaller than the sum of the months (some students ordered in several months).';
mo.getCell(`A${tr + 5}`).value = 'October covers 1–27 Oct only (27 days). Compare months on "Paid orders / day", not on totals.';
mo.getCell(`A${tr + 6}`).value = 'From 23 Oct the payment system started using new statuses (attempted / created / failed); before that every unpaid order was "pending".';
[18, 9, 9, 10, 11, 9, 9, 9, 8, 10, 10, 10, 11, 12, 10, 11, 11, 10, 10, 10, 10].forEach((w, i) => mo.getColumn(i + 1).width = w);
for (let r = 5; r <= tr; r++) {
  ['K', 'M', 'P', 'U'].forEach(c => mo.getCell(`${c}${r}`).numFmt = PCT);
  ['L', 'S'].forEach(c => mo.getCell(`${c}${r}`).numFmt = DEC);
}

// ---------- Product_Plan ----------
const pp = wb.addWorksheet('Product_Plan', { views: [{ state: 'frozen', ySplit: 4 }] });
title(pp, 'Product and plan-wise performance (no price data – volumes only)', 'A plan is counted once per order that contains it. Revenue and AOV cannot be calculated because the export has no amount column.');
header(pp, 4, ['Plan', 'Subject', 'Tier', 'Orders containing plan', 'Paid orders containing plan', 'Plan payment success %', 'Share of paid plan units', ...months.map(m => 'Paid ' + MONTH_NAME[m])]);
const plans = d.plans.slice().sort((a, b) => b.paid - a.paid);
const pTotalPaid = plans.reduce((a, p) => a + p.paid, 0);
const pFirst = 5, pLast = pFirst + plans.length - 1;
const mCol = i => String.fromCharCode(72 + i); // H..L
plans.forEach((p, i) => {
  const r = pFirst + i;
  pp.getRow(r).values = [p.name, p.subject, p.tier,
    F(`COUNTIFS(${CART},"*"&A${r}&"*")`, p.attempts),
    F(`COUNTIFS(${CART},"*"&A${r}&"*",${ST},"captured")`, p.paid),
    F(`E${r}/D${r}`, p.paid / p.attempts),
    F(`E${r}/$E$${pLast + 1}`, p.paid / pTotalPaid),
    ...months.map(m => F(`COUNTIFS(${CART},"*"&$A${r}&"*",${ST},"captured",${MON},"${m}")`, p[m]))];
});
const ptr = pLast + 1;
pp.getRow(ptr).values = ['Total plan units', '', '', F(`SUM(D${pFirst}:D${pLast})`, plans.reduce((a, p) => a + p.attempts, 0)), F(`SUM(E${pFirst}:E${pLast})`, pTotalPaid),
  F(`E${ptr}/D${ptr}`, pTotalPaid / plans.reduce((a, p) => a + p.attempts, 0)), F(`SUM(G${pFirst}:G${pLast})`, 1),
  ...months.map((m, i) => F(`SUM(${mCol(i)}${pFirst}:${mCol(i)}${pLast})`, plans.reduce((a, p) => a + p[m], 0)))];
totalStyle(pp.getRow(ptr));
pp.getCell(`A${ptr + 1}`).value = 'Check: equals "Plans sold" on Monthly sheet?';
pp.getCell(`E${ptr + 1}`).value = F(`IF(E${ptr}=Monthly!R${tr},"OK","MISMATCH")`, pTotalPaid === tot.plansSold ? 'OK' : 'MISMATCH');
for (let r = pFirst; r <= ptr; r++) { pp.getCell(`F${r}`).numFmt = PCT; pp.getCell(`G${r}`).numFmt = PCT; }

// Subject x tier matrix
let r0 = ptr + 3;
pp.getCell(`A${r0}`).value = 'Paid plan units by subject × tier'; pp.getCell(`A${r0}`).font = { bold: true, size: 12 };
header(pp, r0 + 1, ['Subject', 'Flexi', 'Smart', 'Champion', 'Total', 'Share']);
const subjects = ['Maths', 'Science', 'English', 'GK', 'Computer', 'Social', 'JEE', 'NEET'];
subjects.forEach((s, i) => {
  const r = r0 + 2 + i;
  const val = t => (plans.find(p => p.subject === s && p.tier === t) || { paid: 0 }).paid;
  pp.getRow(r).values = [s, ...['Flexi', 'Smart', 'Champion'].map(t => F(`SUMIFS($E$${pFirst}:$E$${pLast},$B$${pFirst}:$B$${pLast},$A${r},$C$${pFirst}:$C$${pLast},"${t}")`, val(t))),
    F(`SUM(B${r}:D${r})`, val('Flexi') + val('Smart') + val('Champion')), F(`E${r}/$E$${r0 + 2 + subjects.length}`, (val('Flexi') + val('Smart') + val('Champion')) / pTotalPaid)];
  pp.getCell(`F${r}`).numFmt = PCT;
});
const mt = r0 + 2 + subjects.length;
const tierPaid = t => plans.filter(p => p.tier === t).reduce((a, p) => a + p.paid, 0);
pp.getRow(mt).values = ['Total', ...['B', 'C', 'D', 'E'].map((c, i) => F(`SUM(${c}${r0 + 2}:${c}${mt - 1})`, i < 3 ? tierPaid(['Flexi', 'Smart', 'Champion'][i]) : pTotalPaid)), F(`E${mt}/E${mt}`, 1)];
pp.getRow(mt + 1).values = ['Tier share', ...['B', 'C', 'D', 'E'].map((c, i) => F(`${c}${mt}/$E$${mt}`, (i < 3 ? tierPaid(['Flexi', 'Smart', 'Champion'][i]) : pTotalPaid) / pTotalPaid))];
totalStyle(pp.getRow(mt)); ['B', 'C', 'D', 'E', 'F'].forEach(c => pp.getCell(`${c}${mt + 1}`).numFmt = PCT); pp.getCell(`F${mt}`).numFmt = PCT;

// Tier mix of carts
r0 = mt + 4;
pp.getCell(`A${r0}`).value = 'Payment success by cart type'; pp.getCell(`A${r0}`).font = { bold: true, size: 12 };
header(pp, r0 + 1, ['Cart type', 'Order attempts', 'Paid orders', 'Payment success %']);
['Flexi only', 'Smart only', 'Champion only', 'Mixed tiers'].forEach((t, i) => {
  const r = r0 + 2 + i, x = d.tierMix.find(z => z.mix === t);
  pp.getRow(r).values = [t, F(`COUNTIFS(${MIX},A${r})`, x.attempts), F(`COUNTIFS(${MIX},A${r},${ST},"captured")`, x.paid), F(`C${r}/B${r}`, x.paid / x.attempts)];
  pp.getCell(`D${r}`).numFmt = PCT;
});

// Top carts
r0 = r0 + 8;
pp.getCell(`A${r0}`).value = 'Top 15 exact cart combinations (by paid orders)'; pp.getCell(`A${r0}`).font = { bold: true, size: 12 };
header(pp, r0 + 1, ['Cart combination', 'Order attempts', 'Paid orders', 'Payment success %', 'Share of all paid orders']);
d.topCarts.forEach((c, i) => {
  const r = r0 + 2 + i;
  pp.getRow(r).values = [c.cart, F(`COUNTIFS(${CART},A${r})`, c.attempts), F(`COUNTIFS(${CART},A${r},${ST},"captured")`, c.paid), F(`C${r}/B${r}`, c.paid / c.attempts), F(`C${r}/Monthly!$E$${tr}`, c.paid / tot.paid)];
  pp.getCell(`D${r}`).numFmt = PCT; pp.getCell(`E${r}`).numFmt = PCT;
});

// Add-ons
r0 = r0 + 2 + d.topCarts.length + 2;
pp.getCell(`A${r0}`).value = 'Flexi add-on products (paid orders)'; pp.getCell(`A${r0}`).font = { bold: true, size: 12 };
header(pp, r0 + 1, ['Add-on type', 'Search text', 'Paid orders containing it', 'Units sold (value)', '% of paid Flexi orders']);
const addonKeys = [['Olympiad Exam', '*Olympiad Exam*'], ['Certificate', '*Certificate*'], ['ePrevious Year Papers', '*Previous Year*'], ['eWorkbook', '*eWorkbook*'],
  ['Online Test Bank / Practice', '*Test*'], ['Monthly Olympiad (1 yr)', '*Monthly Olympiad*']];
const flexiRow = r0 + 2 + addonKeys.length;
addonKeys.forEach(([t, pat], i) => {
  const r = r0 + 2 + i;
  const ordersWith = clean.filter(x => x.paid && x.flexi.some(f => new RegExp(pat.replace(/\*/g, '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).test(f))).length;
  pp.getRow(r).values = [t, pat, F(`COUNTIFS(${FLX},B${r},${ST},"captured")`, ordersWith), (d.addons.find(a => a.type === t) || { units: 0 }).units, F(`C${r}/$C$${flexiRow}`, ordersWith / d.flexiPaid)];
  pp.getCell(`E${r}`).numFmt = PCT;
});
pp.getRow(flexiRow).values = ['Paid orders with a Flexi plan', '*Flexi*', F(`COUNTIFS(${CART},B${flexiRow},${ST},"captured")`, d.flexiPaid)];
pp.getRow(flexiRow + 1).values = ['…of which include at least one add-on beyond the exam (value)', '', d.flexiWithAddon, '', F(`C${flexiRow + 1}/C${flexiRow}`, d.flexiWithAddon / d.flexiPaid)];
pp.getCell(`E${flexiRow + 1}`).numFmt = PCT;
totalStyle(pp.getRow(flexiRow));
[44, 11, 11, 13, 13, 13, 13, 11, 11, 11, 11, 13].forEach((w, i) => pp.getColumn(i + 1).width = w);
pp.getColumn(1).alignment = { wrapText: true };

// ---------- Funnel ----------
const fu = wb.addWorksheet('Funnel');
title(fu, 'Payment and conversion funnel', 'Lead and checkout-started data are not in the export: the funnel starts at "order created at checkout".');
header(fu, 4, ['Order-level stage', 'Orders', '% of order attempts', 'Note']);
const orderStages = [
  ['Order attempts (checkout reached, order created)', F(`Monthly!D${tr}`, tot.attempts), 'Every row in the export = one order sent to the payment gateway'],
  ['Paid (captured)', F(`Monthly!E${tr}`, tot.paid), 'Only these count as successful payments'],
  ['Not paid – pending', F(`Monthly!F${tr}`, sum('pending')), 'NOT counted as paid'],
  ['Not paid – attempted (from 23 Oct)', F(`Monthly!G${tr}`, sum('attempted')), 'New gateway status'],
  ['Not paid – created (from 23 Oct)', F(`Monthly!H${tr}`, sum('created')), 'New gateway status'],
  ['Not paid – failed (from 23 Oct)', F(`Monthly!I${tr}`, sum('failed')), 'New gateway status'],
  ['Refunded / cancelled', 'n/a', 'Not present in the export'],
];
orderStages.forEach((s, i) => { const r = 5 + i; fu.getRow(r).values = [s[0], s[1], typeof s[1] === 'object' ? F(`B${r}/$B$5`, s[1].result / tot.attempts) : '', s[2]]; fu.getCell(`C${r}`).numFmt = PCT; });
fu.getRow(5).font = { bold: true }; fu.getRow(6).font = { bold: true };

const unpaid = clean.filter(r => !r.paid);
const laterY = unpaid.filter(r => r.paidLater === 'Y').length, repY = unpaid.filter(r => r.repeatSameDay === 'Y').length;
header(fu, 14, ['What happened to the unpaid orders', 'Orders', '% of unpaid orders', 'Note']);
fu.getRow(15).values = ['Unpaid orders', F(`Monthly!J${tr}`, unpaid.length), F('B15/B15', 1), ''];
fu.getRow(16).values = ['…student paid on or after that date (recovered retry)', F(`COUNTIFS(${LATER},"Y")`, laterY), F('B16/B15', laterY / unpaid.length), 'Abandoned attempt, but the same student paid later'];
fu.getRow(17).values = ['…student never paid afterwards (lost)', F(`COUNTIFS(${LATER},"N")`, unpaid.length - laterY), F('B17/B15', (unpaid.length - laterY) / unpaid.length), 'Real revenue leakage'];
fu.getRow(18).values = ['…repeat same-day clicks (same student, same cart, same day)', F(`COUNTIFS(${REP},"Y")`, repY), F('B18/B15', repY / unpaid.length), 'Overlaps the two rows above; inflates attempt counts'];
for (let r = 15; r <= 18; r++) fu.getCell(`C${r}`).numFmt = PCT;

header(fu, 21, ['Student-level funnel (unique User IDs, Jun 1 – Oct 27)', 'Students', '% of students starting checkout', 'Definition']);
const sf = [
  ['1. Students who created at least one order', funnel.students, 'Unique User IDs in Clean_Data'],
  ['2. Students who paid at least once', funnel.payers, 'Student conversion = row 2 ÷ row 1'],
  ['3. …paid on their very first order attempt', funnel.firstAttemptPaid, 'First transaction ID for the student was captured'],
  ['4. …paid only after one or more failed/pending attempts', funnel.paidAfterRetry, 'Row 2 − row 3'],
  ['5. Students who never paid', funnel.neverPaid, 'Row 1 − row 2 → abandoned-checkout audience'],
  ['6. …of whom tried more than once', funnel.neverPaidMultiAttempt, 'High-intent students who hit a problem'],
  ['Students paying in more than one month', funnel.repeatPayersMultiMonth, 'Repeat buyers'],
];
sf.forEach((s, i) => { const r = 22 + i; fu.getRow(r).values = [s[0], s[1], F(`B${r}/$B$22`, s[1] / funnel.students), s[2]]; fu.getCell(`C${r}`).numFmt = PCT; });
fu.getCell('A30').value = 'Check: row 3 + row 4 = row 2'; fu.getCell('B30').value = F('IF(B24+B25=B23,"OK","MISMATCH")', 'OK');
fu.getCell('A31').value = 'Check: row 2 + row 5 = row 1'; fu.getCell('B31').value = F('IF(B23+B26=B22,"OK","MISMATCH")', 'OK');

header(fu, 34, ['Monthly funnel', 'Students starting checkout', 'Paying students', 'Student conversion %', 'Order attempts', 'Paid orders', 'Payment success %']);
monthly.forEach((m, i) => {
  const r = 35 + i, mr = 5 + i;
  fu.getRow(r).values = [MONTH_NAME[m.month], F(`Monthly!N${mr}`, m.students), F(`Monthly!O${mr}`, m.payers), F(`C${r}/B${r}`, m.payers / m.students),
    F(`Monthly!D${mr}`, m.attempts), F(`Monthly!E${mr}`, m.paid), F(`F${r}/E${r}`, m.paid / m.attempts)];
  fu.getCell(`D${r}`).numFmt = PCT; fu.getCell(`G${r}`).numFmt = PCT;
});
// October by week – the late-October dip
const ow = [[1, 7], [8, 14], [15, 21], [22, 27]].map(([a, b]) => {
  const rs = clean.filter(r => r.month === '2026-10' && +r.date.slice(8) >= a && +r.date.slice(8) <= b);
  return { label: `Oct ${a}–${b}`, a, b, att: rs.length, paid: rs.filter(r => r.paid).length };
});
const owr = 35 + monthly.length + 2;
header(fu, owr, ['October by week', 'Order attempts', 'Paid orders', 'Payment success %']);
ow.forEach((w, i) => {
  const r = owr + 1 + i;
  fu.getRow(r).values = [w.label, F(`COUNTIFS(Clean_Data!$F$2:$F$${N},">="&DATE(2026,10,${w.a}),Clean_Data!$F$2:$F$${N},"<="&DATE(2026,10,${w.b}))`, w.att),
    F(`COUNTIFS(Clean_Data!$F$2:$F$${N},">="&DATE(2026,10,${w.a}),Clean_Data!$F$2:$F$${N},"<="&DATE(2026,10,${w.b}),${ST},"captured")`, w.paid), F(`C${r}/B${r}`, w.paid / w.att)];
  fu.getCell(`D${r}`).numFmt = PCT;
});
[58, 14, 16, 52, 12, 12, 12].forEach((w, i) => fu.getColumn(i + 1).width = w);

// ---------- Channels_Campaigns ----------
const ch = wb.addWorksheet('Channels_Campaigns');
title(ch, 'Channel and campaign comparison – cannot be calculated');
[
  'The export has no channel, campaign, UTM source/medium/campaign, coupon code or campaign-cost columns.',
  'So WhatsApp, email, Google Ads, Meta Ads, push and organic cannot be compared, and CPL, CPA, ROAS, CTR and open rates cannot be calculated.',
  'No attribution model (first- or last-touch) is possible with this file.',
  '',
  'To unlock this section, add these columns to the monthly order export:',
  '  • utm_source, utm_medium, utm_campaign (captured at sign-up and at checkout)',
  '  • coupon / offer code and discount amount',
  '  • order amount, GST, currency, refund amount',
  '  • lead / registration date (to measure lead → paid)',
  'And share a campaign sheet per channel with: sent, delivered, opened/read, clicked, spend.',
].forEach((t, i) => ch.getCell(`A${3 + i}`).value = t);
ch.getColumn(1).width = 120;

// ---------- Data_Quality ----------
const dq = wb.addWorksheet('Data_Quality');
title(dq, 'Data quality checks and cleaning log', 'Original CSV files were not modified. All cleaning was done on a copy.');
header(dq, 4, ['Check', 'Rows affected', 'Action taken']);
const raw = d.raw;
const reasonCount = {}; d.tests.forEach(r => reasonCount[r.test] = (reasonCount[r.test] || 0) + 1);
const flexiMissing = raw.filter(r => !r.flexiRaw);
const flexiMissingButFlexi = flexiMissing.filter(r => /Flexi/.test(r.cartRaw)).length;
const typoDomains = clean.filter(r => /@(gamil|gmial|gmal|gnail|gmai)\.|@gmail\.co$|@gmail\.com\d/i.test(r.email)).length;
const intlMob = clean.filter(r => !/^[6-9]\d{9}$/.test(r.mobile)).length;
const checks = [
  ['Rows loaded (5 files: Jun, Jul, Aug, Sep, Oct)', raw.length, 'All rows parsed; no malformed rows'],
  ['Columns available', 10, 'User ID, Student Name, Parent Email, Mobile, Transaction ID, Payment Status, Order Date, Plans Bought, Exact Cart Combination, Flexi Products Bought'],
  ['Missing columns: amount, currency, discount/coupon, GST, refund, channel/UTM/campaign, lead date', 'all', 'Revenue, AOV, refunds, channel and campaign sections cannot be calculated'],
  ['Duplicate Transaction IDs', 0, 'None found'],
  ['Invalid / out-of-range dates', 0, 'All dates fall inside their file month. October file ends 27 Oct (27 Oct has only 5 orders – likely a partial day).'],
  ['Test / internal accounts', d.tests.length, `Excluded: ${new Set(d.tests.map(r => r.userId)).size} User IDs, of which ${d.tests.filter(r => r.paid).length} rows were "captured" (would have inflated paid orders). See list below.`],
  ...Object.entries(reasonCount).map(([k, v]) => ['   • ' + k, v, 'Excluded']),
  ['Inconsistent plan labels ("JEE  Champion plan", "NEET Champion plan")', raw.filter(r => /JEE  Champion plan|NEET Champion plan/.test(r.cartRaw)).length, 'Normalised to "JEE Champion Plan" / "NEET Champion Plan"'],
  ['Inconsistent product labels ("Social studies…", double spaces)', raw.filter(r => /Social studies|  /.test(r.flexiRaw)).length, 'Normalised to "Social Studies …" with single spaces'],
  ['Product names containing commas (JEE/NEET exam "Physics, Chemistry & …")', raw.filter(r => /, (Chemistry|Mathematics|Biology)/.test(r.flexiRaw)).length, 'Re-joined so they count as one product'],
  ['Payment-status labels changed on 23 Oct (attempted / created / failed added; "pending" stops)', raw.filter(r => ['attempted', 'created', 'failed'].includes(r.status)).length, 'Kept as separate statuses; all treated as NOT paid'],
  ['Missing Flexi Products Bought', flexiMissing.length, `Expected: ${flexiMissing.length - flexiMissingButFlexi} of these are Smart/Champion-only carts with no Flexi products. ${flexiMissingButFlexi} contain a Flexi plan but no products listed.`],
  ['Missing Parent Email', raw.filter(r => !r.email).length, 'Kept (1 student, 3 pending orders)'],
  ['Mobile not a 10-digit Indian number (after test exclusion)', intlMob, 'Kept – look like overseas numbers (start with 0)'],
  ['Likely typo in email domain (e.g. @gamil.com)', typoDomains, 'Kept; fix before using for email campaigns'],
  ['Suspicious but kept: email contains "testing" with a real-looking student name', clean.filter(r => /testing/i.test(r.email)).length, 'Kept as a real order – please confirm'],
  ['Repeat same-day unpaid clicks (same student, cart and day)', repY, 'Kept and flagged in Clean_Data column N; they inflate order attempts'],
  ['Same mobile used by more than one User ID', 0, 'None found'],
  ['Pending vs paid', '', 'Pending/attempted/created/failed are never counted as paid'],
];
checks.forEach((c, i) => dq.getRow(5 + i).values = c);
const lr = 5 + checks.length + 2;
dq.getCell(`A${lr}`).value = 'Excluded test / internal rows (masked)'; dq.getCell(`A${lr}`).font = { bold: true, size: 12 };
header(dq, lr + 1, ['Transaction ID', 'User ID', 'Student (initials)', 'Email (masked)', 'Mobile (masked)', 'Order Date', 'Status', 'Reason']);
d.tests.sort((a, b) => a.txn - b.txn).forEach((r, i) => dq.getRow(lr + 2 + i).values = [+r.txn, +r.userId, initials(r.name), maskEmail(r.email), maskMobile(r.mobile), r.date, r.status, r.test]);
[70, 13, 110, 22, 14, 12, 10, 40].forEach((w, i) => dq.getColumn(i + 1).width = w);
dq.getColumn(3).alignment = { wrapText: true };

// ---------- Summary (first sheet) ----------
title(su, 'ITO Olympiad orders & payments – Jun 1 to Oct 27, 2026', 'Source: 5 monthly order exports (3,517 rows). Test accounts excluded. Revenue cannot be reported: the export has no amount column.');
header(su, 4, ['Key metric', 'Value', 'How it is calculated']);
const flexiOnly = d.tierMix.find(x => x.mix === 'Flexi only'), smartOnly = d.tierMix.find(x => x.mix === 'Smart only');
const kpis = [
  ['Order attempts', F(`Monthly!D${tr}`, tot.attempts), INT, 'Rows in Clean_Data'],
  ['Paid orders (captured)', F(`Monthly!E${tr}`, tot.paid), INT, 'Status = captured only'],
  ['Payment success rate', F(`Monthly!K${tr}`, tot.paid / tot.attempts), PCT, 'Paid orders ÷ order attempts'],
  ['Students who started checkout', F(`Monthly!N${tr}`, funnel.students), INT, 'Unique User IDs'],
  ['Paying students', F(`Monthly!O${tr}`, funnel.payers), INT, 'Unique User IDs with ≥1 captured order'],
  ['Student conversion (checkout → paid)', F(`Monthly!P${tr}`, funnel.payers / funnel.students), PCT, 'Paying students ÷ students who started checkout'],
  ['Students who never paid', F(`Funnel!B26`, funnel.neverPaid), INT, 'Abandoned-checkout audience'],
  ['Plans sold (paid)', F(`Monthly!R${tr}`, tot.plansSold), INT, 'Sum of "Plans Bought" on paid orders'],
  ['Plans per paid order', F(`Monthly!S${tr}`, tot.plansSold / tot.paid), DEC, 'Plans sold ÷ paid orders'],
  ['Gross revenue / refunds / net revenue / AOV', 'Not available', null, 'No amount or refund column in the export'],
];
kpis.forEach((k, i) => { const r = 5 + i; su.getRow(r).values = [k[0], k[1], k[3]]; if (k[2]) su.getCell(`B${r}`).numFmt = k[2]; });

const pctS = x => (x * 100).toFixed(1) + '%';
const m = Object.fromEntries(d.monthly.map(x => [x.month, x]));
const fmt = n => n.toLocaleString('en-IN');
const lines = [
  ['1. EXECUTIVE SUMMARY'],
  [`• ${fmt(tot.paid)} paid orders from ${fmt(tot.attempts)} order attempts (${pctS(tot.paid / tot.attempts)} payment success). ${fmt(funnel.payers)} of ${fmt(funnel.students)} students (${pctS(funnel.payers / funnel.students)}) paid at least once.`],
  [`• Peak was Jul–Aug (${m['2026-07'].paid} and ${m['2026-08'].paid} paid orders). September fell ${pctS(1 - m['2026-09'].paid / m['2026-08'].paid)} to ${m['2026-09'].paid} because fewer students reached checkout (${m['2026-08'].students} → ${m['2026-09'].students}), not because conversion dropped (${pctS(m['2026-08'].paid / m['2026-08'].attempts)} → ${pctS(m['2026-09'].paid / m['2026-09'].attempts)} payment success).`],
  [`• October (1–27): ${m['2026-10'].paid} paid orders = ${(m['2026-10'].paid / m['2026-10'].days).toFixed(1)}/day, the same pace as September (${(m['2026-09'].paid / m['2026-09'].days).toFixed(1)}/day). Payment success dipped to ${pctS(ow[2].paid / ow[2].att)} in Oct 15–21.`],
  [`• Biggest leak: ${fmt(funnel.neverPaid)} students (${pctS(funnel.neverPaid / funnel.students)}) created an order but never paid; another ${funnel.paidAfterRetry} paid only after a failed/pending attempt.`],
  [`• Flexi-only carts convert worst: ${pctS(flexiOnly.paid / flexiOnly.attempts)} payment success vs ${pctS(smartOnly.paid / smartOnly.attempts)} for Smart-only. Maths Flexi Plan alone is ${d.plans.find(p => p.name === 'Maths Flexi Plan').paid} paid plans (${pctS(d.plans.find(p => p.name === 'Maths Flexi Plan').paid / tot.plansSold)} of all plans sold).`],
  [`• Revenue, AOV, refunds and channel/campaign performance cannot be reported – the export has no amount or marketing columns.`],
  [''],
  ['7. KEY INSIGHTS AND RECOMMENDED ACTIONS (ranked by expected impact)'],
  ['HIGH', `Win back abandoned checkouts. ${funnel.neverPaid} students never paid, and ${funnel.neverPaidMultiAttempt} of them tried more than once. Send an automatic WhatsApp + email payment link at ~1 hour and ~24 hours after an unpaid order. Recovering 20% would add ~${Math.round(funnel.neverPaid * 0.2)} paying students (+${pctS(funnel.neverPaid * 0.2 / funnel.payers)}).`],
  ['HIGH', `Fix Flexi checkout drop-off. Flexi-only carts: ${fmt(flexiOnly.attempts - flexiOnly.paid)} unpaid of ${fmt(flexiOnly.attempts)} (${pctS(1 - flexiOnly.paid / flexiOnly.attempts)}). Smart-only carts lose only ${pctS(1 - smartOnly.paid / smartOnly.attempts)}. Test the Flexi payment page (UPI options, price clarity, add-on step) – this is the most-bought product.`],
  ['HIGH', `Add amount, coupon, GST, refund and UTM/channel columns to the export. Without them revenue, AOV, ROAS and channel ROI cannot be measured, so marketing spend cannot be judged.`],
  ['MEDIUM', `Investigate the late-October payment issue. Oct 15–21 success was ${pctS(ow[2].paid / ow[2].att)} (vs ${pctS(ow[0].paid / ow[0].att)} in Oct 1–7), and on 23 Oct the gateway started using new statuses (attempted / created / failed). Confirm with the payment provider that nothing broke.`],
  ['MEDIUM', `Rebuild demand after August. Students starting checkout fell from ${m['2026-08'].students} (Aug) to ${m['2026-09'].students} (Sep) and ${m['2026-10'].students} (Oct 1–27). Plan reminder campaigns before exam registration deadlines.`],
  ['MEDIUM', `Stop duplicate clicks. ${repY} unpaid orders (${pctS(repY / unpaid.length)} of unpaid) are the same student re-ordering the same cart on the same day. Show a "payment in progress / retry this order" screen instead of creating new orders.`],
  ['MEDIUM', `Check JEE/NEET checkout. ${d.plans.filter(p => /JEE|NEET/.test(p.subject)).reduce((a, p) => a + p.attempts, 0)} orders contained a JEE/NEET plan but only ${d.plans.filter(p => /JEE|NEET/.test(p.subject)).reduce((a, p) => a + p.paid, 0)} were paid. Small volume but a likely broken or confusing flow.`],
  ['LOW', `Champion plans are fading: Maths Champion paid ${d.plans.find(p => p.name === 'Maths Champion Plan')['2026-08']} in Aug vs ${d.plans.find(p => p.name === 'Maths Champion Plan')['2026-10']} in Oct. Review its pricing/positioning.`],
  ['LOW', `Grow add-on sales. ${pctS(d.flexiWithAddon / d.flexiPaid)} of paid Flexi orders include an add-on beyond the exam; Certificates lead. Bundle certificate + previous-year papers as a default suggestion.`],
  ['LOW', `Keep test accounts out of production data: ${d.tests.length} test/internal rows were in the exports, ${d.tests.filter(r => r.paid).length} of them marked "captured".`],
  [''],
  ['8. FORMULAS, ASSUMPTIONS AND LIMITATIONS'],
  ['• Paid order = Payment Status "captured". Pending, attempted, created and failed are never counted as paid.'],
  ['• Payment success rate = paid orders ÷ all order attempts (each row = one order sent to the gateway).'],
  ['• Student conversion = unique students with ≥1 paid order ÷ unique students with ≥1 order (User ID).'],
  ['• MoM change = this month paid orders ÷ previous month − 1. October covers only 27 days, so use paid orders/day for comparison.'],
  ['• Plan units: an order containing 3 plans counts 3 plan units. "Plans Bought" matched the cart in every row.'],
  ['• Dates are taken as given (assumed IST). Orders are assigned to the month of Order Date.'],
  ['• "Paid later" only looks inside Jun–Oct data; students who paid after 27 Oct are counted as never paid.'],
  ['• Test accounts were identified by internal email domains, "test" in the name, test emails and dummy mobile numbers. See Data_Quality.'],
  ['• Not possible from this file: revenue, refunds, AOV, discounts, GST, lead → paid conversion, channel/campaign attribution, CPL/CPA/ROAS.'],
  ['• Personal data is masked in every sheet (initials, partial email and mobile).'],
  [''],
  ['Report sections by sheet: 2 → Monthly · 3 → Product_Plan · 4 → Funnel · 5 → Channels_Campaigns · 6 → Data_Quality · Clean data → Clean_Data'],
];
let sr = 5 + kpis.length + 2;
lines.forEach(l => {
  const row = su.getRow(sr++);
  if (l.length === 2) {
    row.getCell(1).value = l[0]; row.getCell(2).value = l[1];
    row.getCell(1).font = { bold: true, color: { argb: l[0] === 'HIGH' ? 'FFC00000' : l[0] === 'MEDIUM' ? 'FFBF8F00' : 'FF548235' } };
    su.mergeCells(`B${row.number}:C${row.number}`);
    row.getCell(2).alignment = { wrapText: true, vertical: 'top' }; row.height = 46;
  } else {
    row.getCell(1).value = l[0];
    if (/^\d\./.test(l[0])) row.getCell(1).font = { bold: true, size: 12, color: { argb: NAVY } };
    else if (l[0]) { su.mergeCells(`A${row.number}:C${row.number}`); row.getCell(1).alignment = { wrapText: true, vertical: 'top' }; row.height = l[0].length > 130 ? 32 : 18; }
  }
});
su.getColumn(1).width = 42; su.getColumn(2).width = 16; su.getColumn(3).width = 90;

wb.xlsx.writeFile(out).then(() => console.log('written', out));
