(() => {
  const root = typeof window !== 'undefined' ? window : globalThis;

  // ---------- reading files ----------

  const decodeText = (buf) => {
    try { return new TextDecoder('utf-8', { fatal: true }).decode(buf).replace(/^﻿/, ''); }
    catch (e) { return new TextDecoder('windows-1252').decode(buf); }
  };

  const parseCSV = (text) => {
    const first = text.split(/\r?\n/, 1)[0] || '';
    const count = (ch) => first.split('"').filter((_, i) => i % 2 === 0).join('').split(ch).length - 1;
    const sep = [',', ';', '\t'].reduce((a, b) => (count(b) > count(a) ? b : a), ',');
    const rows = []; let row = []; let f = ''; let q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c;
      } else if (c === '"') q = true;
      else if (c === sep) { row.push(f); f = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(f); rows.push(row); row = []; f = '';
      } else f += c;
    }
    if (f || row.length) { row.push(f); rows.push(row); }
    return rows;
  };

  const unXml = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => {
    const k = e.toLowerCase();
    if (k === 'amp') return '&'; if (k === 'lt') return '<'; if (k === 'gt') return '>';
    if (k === 'quot') return '"'; if (k === 'apos') return "'";
    return String.fromCodePoint(k[1] === 'x' ? parseInt(k.slice(2), 16) : parseInt(k.slice(1), 10));
  });
  const attr = (tag, name) => { const m = tag.match(new RegExp('\\s' + name + '="([^"]*)"')); return m ? unXml(m[1]) : null; };

  const readZip = async (buf) => {
    const u8 = new Uint8Array(buf);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) { if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
    if (eocd < 0) throw new Error('This file is not a valid .xlsx workbook.');
    const n = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const entries = new Map();
    for (let k = 0; k < n; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nl));
      entries.set(name.replace(/^\//, ''), { method, csize, off });
      p += 46 + nl + el + cl;
    }
    const read = async (name) => {
      const e = entries.get(name);
      if (!e) return null;
      const start = e.off + 30 + dv.getUint16(e.off + 26, true) + dv.getUint16(e.off + 28, true);
      const raw = u8.subarray(start, start + e.csize);
      if (e.method === 0) return new TextDecoder().decode(raw);
      if (e.method !== 8) throw new Error('This workbook uses a compression this page cannot read.');
      const stream = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      return new TextDecoder().decode(await new Response(stream).arrayBuffer());
    };
    return { has: (name) => entries.has(name), read };
  };

  const DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
  const colIndex = (ref) => { let n = 0; for (const ch of ref.replace(/\d+$/, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const serialToISO = (v) => {
    const d = new Date(Math.round((v - 25569) * 86400000));
    if (Number.isNaN(d.getTime())) return String(v);
    const iso = d.toISOString();
    return Number.isInteger(v) ? iso.slice(0, 10) : iso.slice(0, 19).replace('T', ' ');
  };

  const readXlsx = async (buf) => {
    const z = await readZip(buf);
    const wb = await z.read('xl/workbook.xml');
    if (!wb) throw new Error('This file is not a valid .xlsx workbook.');
    const rels = new Map();
    ((await z.read('xl/_rels/workbook.xml.rels')) || '').replace(/<Relationship\b[^>]*>/g, (t) => { rels.set(attr(t, 'Id'), attr(t, 'Target')); return t; });
    const shared = [];
    const ss = await z.read('xl/sharedStrings.xml');
    if (ss) ss.replace(/<si>([\s\S]*?)<\/si>|<si\/>/g, (m, inner) => {
      const body = (inner || '').replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
      let s = ''; body.replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, (mm, t) => { s += unXml(t); return mm; });
      shared.push(s); return m;
    });
    const custom = new Map();
    const dateXf = [];
    const st = await z.read('xl/styles.xml');
    if (st) {
      st.replace(/<numFmt\b[^>]*>/g, (t) => { custom.set(Number(attr(t, 'numFmtId')), attr(t, 'formatCode') || ''); return t; });
      const xfs = (st.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/) || [])[1] || '';
      xfs.replace(/<xf\b[^>]*>/g, (t) => {
        const id = Number(attr(t, 'numFmtId') || 0);
        const code = custom.has(id) ? custom.get(id).replace(/"[^"]*"|\[[^\]]*\]|\\./g, '') : '';
        dateXf.push(DATE_IDS.has(id) || (custom.has(id) && /[dmy]/i.test(code) && !/^general$/i.test(code)));
        return t;
      });
    }
    const sheets = [];
    const tags = wb.match(/<sheet\b[^>]*>/g) || [];
    for (let i = 0; i < tags.length; i++) {
      const name = attr(tags[i], 'name') || 'Sheet ' + (i + 1);
      const rid = attr(tags[i], 'r:id');
      let target = rels.get(rid) || 'worksheets/sheet' + (i + 1) + '.xml';
      target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      const xml = await z.read(target);
      if (!xml) continue;
      const aoa = [];
      let rowNo = 0;
      xml.replace(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, (m, ra, body) => {
        const r = attr(' ' + ra, 'r');
        rowNo = r ? Number(r) - 1 : rowNo + 1 - (aoa.length ? 0 : 1);
        if (!body) return m;
        const row = [];
        let col = 0;
        body.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (mc, ca, inner) => {
          const ref = attr(' ' + ca, 'r');
          if (ref) col = colIndex(ref);
          const t = attr(' ' + ca, 't');
          const s = Number(attr(' ' + ca, 's') || 0);
          let v = null;
          if (inner) {
            const vm = inner.match(/<v>([\s\S]*?)<\/v>/);
            if (t === 'inlineStr') { let s2 = ''; inner.replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, (q, x) => { s2 += unXml(x); return q; }); v = s2; }
            else if (vm) {
              const raw = unXml(vm[1]);
              if (t === 's') v = shared[Number(raw)] ?? '';
              else if (t === 'b') v = raw === '1' ? 'TRUE' : 'FALSE';
              else if (t === 'str' || t === 'e' || t === 'd') v = raw;
              else { const num = Number(raw); v = dateXf[s] && Number.isFinite(num) ? serialToISO(num) : (Number.isFinite(num) ? num : raw); }
            }
          }
          if (v !== null) row[col] = v;
          col++;
          return mc;
        });
        aoa[rowNo] = row;
        return m;
      });
      sheets.push({ name, aoa: Array.from(aoa, (r) => r || []) });
    }
    return sheets;
  };

  const toTable = (aoa) => {
    const filled = (r) => (r || []).filter((c) => c !== undefined && c !== null && String(c).trim() !== '').length;
    let h = aoa.findIndex((r) => filled(r) >= 2);
    if (h < 0) h = aoa.findIndex((r) => filled(r) >= 1);
    if (h < 0) return { columns: [], rows: [] };
    const width = Math.max(...aoa.slice(h).map((r) => (r || []).length));
    const seen = new Map();
    const columns = Array.from({ length: width }, (_, i) => {
      let c = String(aoa[h][i] ?? '').trim().replace(/\s+/g, ' ') || 'Column ' + (i + 1);
      const k = c.toLowerCase();
      if (seen.has(k)) { seen.set(k, seen.get(k) + 1); c += ' (' + seen.get(k) + ')'; } else seen.set(k, 1);
      return c;
    });
    const rows = [];
    for (const r of aoa.slice(h + 1)) {
      if (!filled(r)) continue;
      const o = {};
      columns.forEach((c, i) => { const v = r[i]; o[c] = v === undefined || v === null ? '' : typeof v === 'number' ? v : String(v).trim(); });
      rows.push(o);
    }
    return { columns, rows };
  };

  const readFile = async (file) => {
    const name = file.name;
    const buf = await file.arrayBuffer();
    const u8 = new Uint8Array(buf);
    if (u8[0] === 0x50 && u8[1] === 0x4b) {
      const sheets = (await readXlsx(buf)).map((s) => ({ name: s.name, ...toTable(s.aoa) })).filter((s) => s.columns.length);
      if (!sheets.length) throw new Error(name + ' has no rows to read.');
      return { name, sheets };
    }
    if (u8[0] === 0xd0 && u8[1] === 0xcf) { const e = new Error('xls'); e.code = 'xls'; throw e; }
    const t = toTable(parseCSV(decodeText(buf)));
    if (!t.columns.length) throw new Error(name + ' has no rows to read.');
    return { name, sheets: [{ name, ...t }] };
  };

  // ---------- ITO order exports ----------

  const norm = (s) => String(s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  const ITO_COLS = {
    uid: ['user id', 'student id', 'userid'], name: ['student name', 'name'], email: ['parent email', 'email'],
    mobile: ['mobile', 'phone', 'mobile number'], txn: ['transaction id', 'order id', 'txn id'], status: ['payment status', 'status'],
    date: ['order date', 'date'], plans: ['plans bought'], cart: ['exact cart combination', 'cart normalised', 'cart normalized', 'cart', 'plans'],
  };
  const findCols = (columns) => {
    const by = new Map(columns.map((c) => [norm(c), c]));
    const out = {};
    for (const [k, names] of Object.entries(ITO_COLS)) out[k] = names.map((n) => by.get(n)).find(Boolean) || null;
    return out;
  };
  const isIto = (columns) => { const c = findCols(columns); return !!(c.status && c.cart && c.txn && c.uid); };

  // Test and internal accounts, as removed from the dashboard's own data.
  const TEST_IDS = new Set(['421956', '420782']);
  const isTestRow = (o) => TEST_IDS.has(o.student) || /test/i.test(o.name) || /^\s*test/i.test(o.email) || /indiantalent|@itoexam\.com/i.test(o.email);

  const SUBJECTS = { maths: 'Maths', math: 'Maths', mathematics: 'Maths', science: 'Science', english: 'English', gk: 'GK', 'general knowledge': 'GK', computer: 'Computer', computers: 'Computer', social: 'Social', 'social studies': 'Social', 'social science': 'Social', jee: 'JEE', neet: 'NEET' };
  const TIERS = { flexi: 'Flexi', smart: 'Smart', champion: 'Champion' };
  const normPlan = (label) => {
    const w = String(label).trim().split(/\s+/).filter(Boolean);
    const ti = w.findIndex((x) => TIERS[x.toLowerCase()]);
    if (ti < 1) return { label: w.join(' '), subject: null, tier: null };
    const sKey = w.slice(0, ti).join(' ').toLowerCase();
    const subject = SUBJECTS[sKey] || w.slice(0, ti).join(' ');
    const tier = TIERS[w[ti].toLowerCase()];
    return { label: subject + ' ' + tier + ' Plan', subject, tier };
  };
  const toISODate = (v) => {
    if (typeof v === 'number' && v > 20000 && v < 80000) return serialToISO(Math.floor(v));
    const s = String(v ?? '').trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
    m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})/);
    if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
    if (/^\d{5}(\.\d+)?$/.test(s)) return serialToISO(Math.floor(Number(s)));
    return '';
  };

  const cleanIto = (tables, keepTest) => {
    const seen = new Set();
    const orders = [];
    let removed = 0, dupes = 0;
    for (const t of tables) {
      const c = findCols(t.columns);
      for (const r of t.rows) {
        const items = String(r[c.cart] ?? '').split('+').map((x) => x.trim()).filter(Boolean).map(normPlan);
        const date = c.date ? toISODate(r[c.date]) : '';
        const status = String(r[c.status] ?? '').trim().toLowerCase();
        const o = {
          txn: String(r[c.txn] ?? '').trim(), student: String(r[c.uid] ?? '').trim(),
          name: c.name ? String(r[c.name] ?? '') : '', email: c.email ? String(r[c.email] ?? '') : '', mobile: c.mobile ? String(r[c.mobile] ?? '') : '',
          date, month: date ? date.slice(0, 7) : 'unknown', status, paid: status === 'captured',
          items, cart: items.map((x) => x.label).join(' + '),
        };
        if (!o.txn && !o.student) continue;
        const key = o.txn || JSON.stringify(r);
        if (seen.has(key)) { dupes++; continue; }
        seen.add(key);
        if (!keepTest && isTestRow(o)) { removed++; continue; }
        const tiers = new Set(items.map((x) => x.tier));
        o.cart_type = tiers.size === 1 && !tiers.has(null) ? [...tiers][0] + ' only' : tiers.has(null) || !tiers.size ? 'Other' : 'Mixed tiers';
        orders.push(o);
      }
    }
    return { orders, removed, dupes };
  };

  const txnNum = (o) => { const n = Number(o.txn); return Number.isFinite(n) ? n : Infinity; };
  const statsFor = (rows) => {
    const first = new Map();
    rows.forEach((o) => { const f = first.get(o.student); if (!f || txnNum(o) < txnNum(f) || (txnNum(o) === txnNum(f) && o.date < f.date)) first.set(o.student, o); });
    const paidRows = rows.filter((o) => o.paid);
    const attempts = rows.length, paid = paidRows.length;
    const pending = rows.filter((o) => o.status === 'pending').length;
    const students = first.size;
    const payers = new Set(paidRows.map((o) => o.student)).size;
    const firstTry = [...first.values()].filter((o) => o.paid).length;
    return {
      attempts, paid, success: attempts ? paid / attempts : null, pending, other_unpaid: attempts - paid - pending,
      students, payers, conversion: students ? payers / students : null, first_try: firstTry, after_retry: payers - firstTry,
      never_paid: students - payers, plans_sold: paidRows.reduce((a, o) => a + o.items.length, 0),
    };
  };

  const itoReport = (orders) => {
    const months = [...new Set(orders.map((o) => o.month))].sort();
    const byMonth = months.map((m) => ({ month: m, ...statsFor(orders.filter((o) => o.month === m)) }));
    const summary = statsFor(orders);
    const plansMap = new Map();
    let units = 0;
    orders.forEach((o) => o.items.forEach((it) => {
      const r = plansMap.get(it.label) || { plan: it.label, subject: it.subject || '', tier: it.tier || '', attempts: 0, paid: 0 };
      r.attempts++; if (o.paid) { r.paid++; units++; }
      plansMap.set(it.label, r);
    }));
    const plans = [...plansMap.values()].map((r) => ({ ...r, success: r.paid / r.attempts, share: units ? r.paid / units : 0 }))
      .sort((a, b) => b.paid - a.paid || b.attempts - a.attempts);
    const carts = ['Flexi only', 'Smart only', 'Champion only', 'Mixed tiers', 'Other'].map((c) => {
      const rows = orders.filter((o) => o.cart_type === c);
      const paid = rows.filter((o) => o.paid).length;
      return { cart_type: c, attempts: rows.length, paid, success: rows.length ? paid / rows.length : null };
    }).filter((r) => r.attempts > 0);
    const dates = orders.map((o) => o.date).filter(Boolean).sort();
    return { summary, byMonth, plans, carts, from: dates[0] || '', to: dates[dates.length - 1] || '' };
  };

  const followUp = (orders) => {
    const by = new Map();
    orders.forEach((o) => { const a = by.get(o.student) || []; a.push(o); by.set(o.student, a); });
    const out = [];
    by.forEach((list, student) => {
      if (list.some((o) => o.paid)) return;
      list.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : txnNum(b) - txnNum(a)));
      const last = list[0];
      out.push({
        student, name: last.name, email: last.email, mobile: last.mobile, last_order: last.date, last_status: last.status,
        orders_tried: list.length, plans_tried: [...new Set(list.flatMap((o) => o.items.map((i) => i.label)))].join(' + '),
      });
    });
    return out.sort((a, b) => (a.last_order < b.last_order ? 1 : a.last_order > b.last_order ? -1 : 0));
  };

  // ---------- any other spreadsheet ----------

  const asNumber = (v) => {
    if (typeof v === 'number') return v;
    const s = String(v).replace(/[,\s₹$€£]/g, '').replace(/%$/, '');
    if (!/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s)) return null;
    return Number(s);
  };
  const profile = (table) => table.columns.map((c) => {
    const vals = table.rows.map((r) => r[c]).filter((v) => v !== '' && v !== null && v !== undefined);
    const nums = vals.map(asNumber).filter((n) => n !== null);
    const numeric = vals.length > 0 && nums.length / vals.length >= 0.9;
    const dates = !numeric && vals.length > 0 && vals.filter((v) => /^\d{4}-\d{2}-\d{2}/.test(String(v))).length / vals.length >= 0.9;
    const distinct = new Set(vals.map(String)).size;
    const sorted = numeric ? nums.slice().sort((a, b) => a - b) : dates ? vals.map(String).sort() : [];
    return {
      column: c, type: numeric ? 'Number' : dates ? 'Date' : 'Text', filled: vals.length, empty: table.rows.length - vals.length, distinct,
      total: numeric ? nums.reduce((a, b) => a + b, 0) : null,
      min: sorted.length ? sorted[0] : null,
      max: sorted.length ? sorted[sorted.length - 1] : null,
    };
  });
  const groupBy = (table, by, sum) => {
    const m = new Map();
    table.rows.forEach((r) => {
      const k = String(r[by] ?? '').trim() || '(blank)';
      const g = m.get(k) || { group: k, rows: 0, total: 0 };
      g.rows++;
      if (sum) { const n = asNumber(r[sum]); if (n !== null) g.total += n; }
      m.set(k, g);
    });
    const all = table.rows.length || 1;
    return [...m.values()].map((g) => ({ ...g, share: g.rows / all, total: sum ? g.total : null }))
      .sort((a, b) => (sum ? b.total - a.total : 0) || b.rows - a.rows);
  };

  // ---------- writing CSV and Excel ----------

  const csvCell = (v) => {
    if (v === null || v === undefined) return '';
    let s = typeof v === 'number' ? String(v) : String(v);
    if (typeof v !== 'number' && /^[=+@]/.test(s)) s = "'" + s;
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const toCSV = (columns, rows) => '﻿' + [columns.map((c) => csvCell(c.label)).join(','),
    ...rows.map((r) => columns.map((c) => csvCell(r[c.key])).join(','))].join('\r\n') + '\r\n';

  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (u8) => { let c = 0xffffffff; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const zip = (files) => {
    const enc = new TextEncoder();
    const parts = [], central = [];
    let offset = 0;
    for (const f of files) {
      const name = enc.encode(f.name);
      const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
      const crc = crc32(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(12, 0x21, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, name.length, true);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(14, 0x21, true);
      ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, name.length, true); ch.setUint32(42, offset, true);
      parts.push(new Uint8Array(lh.buffer), name, data);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + data.length;
    }
    const cdSize = central.reduce((a, b) => a + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true); end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    const all = [...parts, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((a, b) => a + b.length, 0));
    let p = 0; all.forEach((b) => { out.set(b, p); p += b.length; });
    return out;
  };
  const esc = (s) => String(s).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '').slice(0, 32767).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colName = (i) => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const sheetXml = (sheet) => {
    const widths = sheet.columns.map((c) => Math.min(60, Math.max(8, String(c.label).length + 2, ...sheet.rows.slice(0, 200).map((r) => String(r[c.key] ?? '').length + 2))));
    const cell = (ref, v, type) => {
      if (v === null || v === undefined || v === '') return '';
      if (typeof v === 'number' && Number.isFinite(v)) return '<c r="' + ref + '"' + (type === 'pct' ? ' s="2"' : type === 'int' ? ' s="3"' : '') + '><v>' + v + '</v></c>';
      return '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
    };
    const head = '<row r="1">' + sheet.columns.map((c, i) => '<c r="' + colName(i) + '1" t="inlineStr" s="1"><is><t>' + esc(c.label) + '</t></is></c>').join('') + '</row>';
    const body = sheet.rows.map((r, j) => '<row r="' + (j + 2) + '">' + sheet.columns.map((c, i) => cell(colName(i) + (j + 2), r[c.key], c.type)).join('') + '</row>').join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
      + '<cols>' + widths.map((w, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols>'
      + '<sheetData>' + head + body + '</sheetData></worksheet>';
  };
  const toXlsx = (sheets) => {
    const used = new Set();
    const names = sheets.map((s, i) => {
      let n = String(s.name).replace(/[[\]:*?/\\]/g, ' ').slice(0, 31).trim() || 'Sheet' + (i + 1);
      while (used.has(n.toLowerCase())) n = n.slice(0, 28) + ' ' + (i + 1);
      used.add(n.toLowerCase()); return n;
    });
    const ns = 'http://schemas.openxmlformats.org/';
    const files = [
      { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="' + ns + 'package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + sheets.map((_, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') + '</Types>' },
      { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + ns + 'package/2006/relationships"><Relationship Id="rId1" Type="' + ns + 'officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="' + ns + 'spreadsheetml/2006/main" xmlns:r="' + ns + 'officeDocument/2006/relationships"><sheets>' + names.map((n, i) => '<sheet name="' + esc(n) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + ns + 'package/2006/relationships">' + sheets.map((_, i) => '<Relationship Id="rId' + (i + 1) + '" Type="' + ns + 'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') + '<Relationship Id="rId' + (sheets.length + 1) + '" Type="' + ns + 'officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/styles.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="' + ns + 'spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>' },
      ...sheets.map((s, i) => ({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: sheetXml(s) })),
    ];
    return zip(files);
  };

  // ---------- report layouts (shared by uploads, saved reports and the dashboard) ----------

  const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const monthLabel = (m) => (/^\d{4}-\d{2}$/.test(m) ? MONTH_NAMES[Number(m.slice(5)) - 1] + ' ' + m.slice(0, 4) : m === 'all' ? 'All months' : m === 'unknown' ? 'No date' : m);
  const COLS = {
    month: [
      { key: 'month_label', label: 'Month', type: 'text' }, { key: 'attempts', label: 'Order attempts', type: 'int' }, { key: 'paid', label: 'Paid orders', type: 'int' },
      { key: 'success', label: 'Payment success', type: 'pct' }, { key: 'pending', label: 'Pending', type: 'int' }, { key: 'other_unpaid', label: 'Attempted / created / failed', type: 'int' },
      { key: 'students', label: 'Students at checkout', type: 'int' }, { key: 'payers', label: 'Students who paid', type: 'int' }, { key: 'conversion', label: 'Student conversion', type: 'pct' },
      { key: 'never_paid', label: "Students who didn't pay", type: 'int' }, { key: 'first_try', label: 'Paid on first try', type: 'int' }, { key: 'after_retry', label: 'Paid after a retry', type: 'int' },
      { key: 'plans_sold', label: 'Plans sold', type: 'int' },
    ],
    plans: [{ key: 'plan', label: 'Plan', type: 'text' }, { key: 'attempts', label: 'Orders', type: 'int' }, { key: 'paid', label: 'Paid', type: 'int' }, { key: 'success', label: 'Payment success', type: 'pct' }, { key: 'share', label: 'Share of paid plans', type: 'pct' }],
    carts: [{ key: 'cart_type', label: 'Cart type', type: 'text' }, { key: 'attempts', label: 'Order attempts', type: 'int' }, { key: 'paid', label: 'Paid orders', type: 'int' }, { key: 'success', label: 'Payment success', type: 'pct' }],
    orders: [{ key: 'txn', label: 'Transaction ID', type: 'text' }, { key: 'student', label: 'User ID', type: 'text' }, { key: 'date', label: 'Order date', type: 'text' }, { key: 'status', label: 'Payment status', type: 'text' }, { key: 'paid_n', label: 'Paid (1 = yes)', type: 'int' }, { key: 'plans_n', label: 'Plans in order', type: 'int' }, { key: 'cart', label: 'Cart', type: 'text' }, { key: 'cart_type', label: 'Cart type', type: 'text' }],
    follow: [{ key: 'student', label: 'User ID', type: 'text' }, { key: 'name', label: 'Student name', type: 'text' }, { key: 'email', label: 'Parent email', type: 'text' }, { key: 'mobile', label: 'Mobile', type: 'text' }, { key: 'last_order', label: 'Last order date', type: 'text' }, { key: 'last_status', label: 'Last status', type: 'text' }, { key: 'orders_tried', label: 'Orders tried', type: 'int' }, { key: 'plans_tried', label: 'Plans tried', type: 'text' }],
    profile: [{ key: 'column', label: 'Column', type: 'text' }, { key: 'type', label: 'Type', type: 'text' }, { key: 'filled', label: 'Filled', type: 'int' }, { key: 'empty', label: 'Empty', type: 'int' }, { key: 'distinct', label: 'Distinct values', type: 'int' }, { key: 'total', label: 'Total', type: 'num' }, { key: 'min', label: 'Lowest', type: 'num' }, { key: 'max', label: 'Highest', type: 'num' }],
  };
  const summaryRows = (s, extra) => [
    ...extra,
    { metric: 'Order attempts', value: s.attempts, type: 'int' }, { metric: 'Paid orders', value: s.paid, type: 'int' },
    { metric: 'Payment success', value: s.success, type: 'pct' }, { metric: 'Students at checkout', value: s.students, type: 'int' },
    { metric: 'Students who paid', value: s.payers, type: 'int' }, { metric: 'Student conversion', value: s.conversion, type: 'pct' },
    { metric: "Students who didn't pay", value: s.never_paid, type: 'int' }, { metric: 'Plans sold', value: s.plans_sold, type: 'int' },
  ];
  const summarySheet = (rows) => ({
    name: 'Summary',
    columns: [{ key: 'metric', label: 'Measure', type: 'text' }, { key: 'v_int', label: 'Value', type: 'int' }, { key: 'v_pct', label: 'Rate', type: 'pct' }, { key: 'v_text', label: 'Detail', type: 'text' }],
    rows: rows.map((r) => ({ metric: r.metric, v_int: r.type === 'int' ? r.value : null, v_pct: r.type === 'pct' ? r.value : null, v_text: r.type === 'text' ? r.value : null })),
  });
  const withMonthLabel = (rows) => rows.map((r) => ({ ...r, month_label: monthLabel(r.month) }));

  root.itoReports = { parseCSV, readXlsx, readZip, toTable, readFile, isIto, cleanIto, itoReport, followUp, profile, groupBy, toCSV, toXlsx, COLS, monthLabel, summaryRows, summarySheet, withMonthLabel, isTestRow };
  if (typeof document === 'undefined' || typeof dash === 'undefined') return;

  // ---------- the page ----------

  const { fmtInt, fmtPct, fmt } = window.ito;
  const fmtDate = (iso) => { if (!iso) return ''; const d = new Date(iso.length === 10 ? iso + 'T00:00:00Z' : iso); return Number.isNaN(d.getTime()) ? iso : d3.utcFormat('%-d %b %Y')(d); };
  const fmtWhen = (iso) => { const d = new Date(iso); return Number.isNaN(d.getTime()) ? '' : d3.timeFormat('%-d %b %Y, %H:%M')(d); };
  const fmtCell = (v, type) => {
    if (v === null || v === undefined || v === '') return '';
    if (type === 'pct') return fmtPct(v);
    if (type === 'int') return fmtInt(v);
    if (type === 'num') return typeof v === 'number' ? d3.format(',.4~f')(v) : String(v);
    return String(v);
  };
  const all = (sel) => document.querySelectorAll(sel);
  const fileStem = (s) => String(s).replace(/\.[^.]+$/, '').replace(/[^\w\-. ]+/g, ' ').trim().replace(/\s+/g, '_').slice(0, 60) || 'report';

  const S = {
    busy: false, error: '', errorKind: '', files: [], sheetIdx: 0, keepTest: false,
    kind: null, name: '', meta: '', report: null, orders: null, removed: 0, dupes: 0, savedId: null,
    table: null, prof: null, by: '', sum: '', group: null,
    saved: [], savedState: 'loading', names: {}, confirmDel: null, saveState: '',
    db: null, dl: null, canEdit: false, uid: null, userApi: null,
  };

  const sync = () => {
    all('[data-up-msg]').forEach((el) => { const k = el.dataset.upMsg; el.hidden = !(k === 'reading' ? S.busy : k === 'error' ? !!S.error && S.errorKind !== 'xls' : k === S.errorKind && !!S.error); });
    all('[data-up-view]').forEach((el) => { el.hidden = el.dataset.upView !== S.kind; });
    all('[data-up-sheets]').forEach((el) => { el.hidden = !(S.files.length === 1 && S.files[0].sheets.length > 1 && !S.savedId); });
    all('[data-up-needs="rows"]').forEach((el) => { el.hidden = !S.orders && !(S.kind === 'generic' && S.table); });
    all('[data-up-needs="save"]').forEach((el) => { el.hidden = !S.db || !S.canEdit || !!S.savedId; });
    all('[data-up-needs="noedit"]').forEach((el) => { el.hidden = !S.db || S.canEdit || !!S.savedId; });
    all('[data-up-needs="savedview"]').forEach((el) => { el.hidden = !S.savedId; });
    all('[data-up-save]').forEach((el) => { el.hidden = S.saveState !== el.dataset.upSave; });
    all('.ito-save').forEach((b) => { b.disabled = S.saveState === 'saving' || S.saveState === 'saved'; });
    all('.ito-keeptest').forEach((c) => { c.checked = S.keepTest; });
    all('[data-saved-state]').forEach((el) => { el.hidden = el.dataset.savedState !== S.savedState; });
    all('.ito-gen-group-box').forEach((el) => { el.hidden = !S.by; });
  };
  const DRAWS = ['up-error', 'up-sheet', 'up-name', 'up-meta', 'up-kpi', 'up-month-rows', 'up-plan-rows', 'up-cart-rows', 'gen-by', 'gen-sum', 'gen-cols', 'gen-group', 'gen-preview', 'saved-list'];
  const render = () => { sync(); DRAWS.forEach((n) => dash.draw(n)); };

  const build = () => {
    S.kind = null; S.report = null; S.orders = null; S.table = null; S.prof = null; S.group = null; S.savedId = null; S.saveState = '';
    if (!S.files.length) return;
    const pick = (f) => (f.sheets.length === 1 ? f.sheets[0] : f.sheets[S.files.length === 1 ? Math.min(S.sheetIdx, f.sheets.length - 1) : Math.max(0, f.sheets.findIndex((s) => isIto(s.columns)))]);
    const tables = S.files.map(pick);
    const names = S.files.map((f) => f.name);
    S.name = names.length === 1 ? names[0] : names.length + ' files';
    if (tables.every((t) => isIto(t.columns))) {
      const { orders, removed, dupes } = cleanIto(tables, S.keepTest);
      S.kind = 'ito'; S.orders = orders; S.removed = removed; S.dupes = dupes; S.report = itoReport(orders);
      S.fileNames = names;
      return;
    }
    const cols = [...new Set(tables.flatMap((t) => t.columns))];
    const many = tables.length > 1;
    const rows = tables.flatMap((t, i) => t.rows.map((r) => (many ? { 'Source file': names[i], ...r } : r)));
    S.kind = 'generic';
    S.fileNames = names;
    S.table = { columns: many ? ['Source file', ...cols] : cols, rows };
    S.prof = profile(S.table);
    const firstText = S.prof.find((c) => c.type === 'Text' && c.distinct > 1 && c.distinct <= Math.max(50, rows.length / 4));
    S.by = firstText ? firstText.column : '';
    S.sum = '';
    S.group = S.by ? groupBy(S.table, S.by, S.sum) : null;
  };

  const metaText = () => {
    if (S.kind === 'ito') {
      const r = S.report;
      const parts = [fmtInt(r.summary.attempts) + ' orders'];
      if (r.from) parts.push(fmtDate(r.from) + ' – ' + fmtDate(r.to));
      if (S.removed) parts.push(fmtInt(S.removed) + ' test or internal rows removed');
      else if (S.keepTest) parts.push('test and internal accounts kept');
      if (S.dupes) parts.push(fmtInt(S.dupes) + ' repeated orders counted once');
      if (S.fileNames && S.fileNames.length > 1) parts.push(S.fileNames.length + ' files combined');
      return parts.join(' · ');
    }
    if (S.kind === 'generic') {
      const n = S.table ? S.table.rows.length : S.savedRows || 0;
      const c = S.prof ? S.prof.length : 0;
      return fmtInt(n) + ' rows · ' + fmtInt(c) + ' columns' + (S.fileNames && S.fileNames.length > 1 ? ' · ' + S.fileNames.length + ' files combined' : '');
    }
    return '';
  };

  const handleFiles = async (list) => {
    const files = [...list];
    if (!files.length) return;
    S.busy = true; S.error = ''; S.errorKind = ''; render();
    try {
      const read = [];
      for (const f of files) read.push(await readFile(f));
      S.files = read; S.sheetIdx = 0;
      if (read.length === 1 && read[0].sheets.length > 1) { const i = read[0].sheets.findIndex((s) => isIto(s.columns)); S.sheetIdx = i >= 0 ? i : 0; }
      build();
    } catch (e) {
      S.error = e && e.code === 'xls' ? 'xls' : (e && e.message) || 'This file could not be read.';
      S.errorKind = e && e.code === 'xls' ? 'xls' : 'error';
    }
    S.busy = false;
    render();
  };

  // downloads
  const save = async (filename, data) => {
    const blob = data instanceof Blob ? data : new Blob([data]);
    if (S.dl) {
      try { await S.dl.save({ filename, data: blob }); return; }
      catch (e) { if (e && (e.code === 'declined' || e.code === 'rate_limited')) return; }
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.rel = 'noopener';
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  };
  const ordersOut = () => S.orders.map((o) => ({ ...o, paid_n: o.paid ? 1 : 0, plans_n: o.items.length }));
  const itoSheets = () => {
    const r = S.report;
    const extra = [{ metric: 'Source', value: (S.fileNames || [S.name]).join(', '), type: 'text' }];
    if (r.from) extra.push({ metric: 'Period', value: fmtDate(r.from) + ' – ' + fmtDate(r.to), type: 'text' });
    extra.push({ metric: 'Test or internal rows removed', value: S.removed, type: 'int' });
    const sheets = [
      summarySheet(summaryRows(r.summary, extra)),
      { name: 'By month', columns: COLS.month, rows: withMonthLabel(r.byMonth) },
      { name: 'Plans', columns: COLS.plans, rows: r.plans },
      { name: 'Cart types', columns: COLS.carts, rows: r.carts },
    ];
    if (S.orders) sheets.push({ name: 'Orders (cleaned)', columns: COLS.orders, rows: ordersOut() });
    return sheets;
  };
  const genSheets = () => {
    const sheets = [];
    if (S.table) sheets.push({ name: 'Data', columns: S.table.columns.map((c) => ({ key: c, label: c, type: 'auto' })), rows: S.table.rows });
    if (S.prof) sheets.push({ name: 'Columns', columns: COLS.profile, rows: S.prof });
    if (S.group) sheets.push({ name: 'By ' + S.by, columns: groupCols(), rows: S.group });
    return sheets;
  };
  const groupCols = () => [{ key: 'group', label: S.by, type: 'text' }, { key: 'rows', label: 'Rows', type: 'int' }, { key: 'share', label: 'Share of rows', type: 'pct' }, ...(S.sum ? [{ key: 'total', label: 'Total ' + S.sum, type: 'num' }] : [])];
  const DOWNLOADS = {
    'ito-xlsx': () => save(fileStem(S.name) + '_report.xlsx', new Blob([toXlsx(itoSheets())])),
    'ito-month': () => save(fileStem(S.name) + '_by_month.csv', toCSV(COLS.month, withMonthLabel(S.report.byMonth))),
    'ito-plans': () => save(fileStem(S.name) + '_plans.csv', toCSV(COLS.plans, S.report.plans)),
    'ito-carts': () => save(fileStem(S.name) + '_cart_types.csv', toCSV(COLS.carts, S.report.carts)),
    'ito-orders': () => S.orders && save(fileStem(S.name) + '_orders_cleaned.csv', toCSV(COLS.orders, ordersOut())),
    'ito-follow': () => S.orders && save(fileStem(S.name) + '_unpaid_follow_up.csv', toCSV(COLS.follow, followUp(S.orders))),
    'gen-xlsx': () => save(fileStem(S.name) + '_report.xlsx', new Blob([toXlsx(genSheets())])),
    'gen-csv': () => S.table && save(fileStem(S.name) + '.csv', toCSV(S.table.columns.map((c) => ({ key: c, label: c })), S.table.rows)),
    'gen-group-csv': () => S.group && save(fileStem(S.name) + '_by_' + fileStem(S.by) + '.csv', toCSV(groupCols(), S.group)),
    'gen-cols-csv': () => S.prof && save(fileStem(S.name) + '_columns.csv', toCSV(COLS.profile, S.prof)),
  };

  // shared "Saved reports"
  const savedDoc = () => {
    const base = { name: S.name, files: S.fileNames || [S.name], kind: S.kind, at: new Date().toISOString(), by: S.uid || null };
    if (S.kind === 'ito') return { ...base, keepTest: S.keepTest, removed: S.removed, dupes: S.dupes, ...S.report };
    return { ...base, rowsCount: S.table.rows.length, columns: S.prof.slice(0, 300), group: S.group ? { by: S.by, sum: S.sum || null, rows: S.group.slice(0, 1000) } : null };
  };
  const saveShared = async () => {
    if (!S.db || !S.kind) return;
    S.saveState = 'saving'; sync();
    try {
      const id = 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
      await S.db.collection('reports').doc(id).set(savedDoc());
      S.saveState = 'saved';
    } catch (e) {
      S.saveState = 'failed';
      if (e && e.code === 'invalid_argument') { S.canEdit = false; }
    }
    sync();
  };
  const openSaved = (id) => {
    const d = S.saved.find((x) => x.id === id);
    if (!d) return;
    S.files = []; S.error = ''; S.errorKind = ''; S.savedId = id; S.saveState = '';
    S.name = d.name || 'Saved report'; S.fileNames = d.files || [S.name];
    S.orders = null; S.table = null;
    if (d.kind === 'ito') {
      S.kind = 'ito'; S.removed = d.removed || 0; S.dupes = d.dupes || 0; S.keepTest = !!d.keepTest;
      S.report = { summary: d.summary, byMonth: d.byMonth || [], plans: d.plans || [], carts: d.carts || [], from: d.from || '', to: d.to || '' };
    } else {
      S.kind = 'generic'; S.prof = d.columns || []; S.savedRows = d.rowsCount || 0;
      S.by = d.group ? d.group.by : ''; S.sum = d.group ? d.group.sum || '' : ''; S.group = d.group ? d.group.rows : null;
    }
    render();
    all('#upload-result').forEach((el) => el.scrollIntoView({ block: 'start' }));
  };
  const downloadSaved = (id) => {
    const d = S.saved.find((x) => x.id === id);
    if (!d) return;
    openSaved(id);
    if (S.kind === 'ito') DOWNLOADS['ito-xlsx'](); else DOWNLOADS['gen-xlsx']();
  };
  const deleteSaved = async (id) => {
    if (S.confirmDel !== id) { S.confirmDel = id; dash.draw('saved-list'); return; }
    S.confirmDel = null;
    try { await S.db.collection('reports').doc(id).delete(); } catch (e) { /* the list stays as the store has it */ }
    if (S.savedId === id) { S.savedId = null; S.kind = null; }
    render();
  };

  // ---------- drawings ----------

  dash.draw('up-error', (el) => { el.textContent = S.errorKind === 'error' ? S.error : '—'; });
  dash.draw('up-sheet', (sel) => {
    const f = S.files[0];
    sel.replaceChildren(...(f ? f.sheets : []).map((s, i) => new Option(s.name, String(i))));
    sel.value = String(S.sheetIdx);
    sel.onchange = () => { S.sheetIdx = Number(sel.value); build(); render(); };
  });
  dash.draw('up-name', (el) => { el.textContent = S.kind ? S.name : '—'; });
  dash.draw('up-meta', (el) => { el.textContent = S.kind ? metaText() : '—'; });
  dash.draw('up-kpi', (el) => {
    const s = S.report && S.report.summary;
    const f = el.dataset.field;
    const v = f === 'removed' ? S.removed : s ? s[f] : null;
    el.textContent = s ? fmt(v, el.dataset.fmt) : '—';
  });
  const fillRows = (tbody, rows, cols) => {
    tbody.replaceChildren();
    rows.forEach((r) => {
      const tr = document.createElement('tr');
      cols.forEach((c) => { const td = document.createElement('td'); td.textContent = fmtCell(r[c.key], c.type); tr.append(td); });
      tbody.append(tr);
    });
  };
  dash.draw('up-month-rows', (tbody) => {
    if (!S.report) { tbody.replaceChildren(); return; }
    const rows = withMonthLabel(S.report.byMonth);
    if (rows.length > 1) rows.push({ month_label: 'All months', ...S.report.summary });
    fillRows(tbody, rows, COLS.month.slice(0, 10));
  });
  dash.draw('up-plan-rows', (tbody) => fillRows(tbody, S.report ? S.report.plans : [], COLS.plans));
  dash.draw('up-cart-rows', (tbody) => fillRows(tbody, S.report ? S.report.carts : [], COLS.carts));

  const colOptions = (sel, first, filter, value, onPick) => {
    const cols = (S.prof || []).filter(filter).map((c) => c.column);
    sel.replaceChildren(new Option(first, ''), ...cols.map((c) => new Option(c, c)));
    sel.value = cols.includes(value) ? value : '';
    sel.disabled = !S.table;
    sel.onchange = () => { onPick(sel.value); S.group = S.by && S.table ? groupBy(S.table, S.by, S.sum) : S.group; render(); };
  };
  dash.draw('gen-by', (sel) => colOptions(sel, 'Pick a column', (c) => c.type !== 'Number' || c.distinct <= 50, S.by, (v) => { S.by = v; }));
  dash.draw('gen-sum', (sel) => colOptions(sel, 'Count rows only', (c) => c.type === 'Number', S.sum, (v) => { S.sum = v; }));
  dash.draw('gen-cols', (tbody) => fillRows(tbody, S.prof || [], COLS.profile));
  const fillTable = (table, cols, rows) => {
    table.replaceChildren();
    const thead = document.createElement('thead');
    const hr = document.createElement('tr');
    cols.forEach((c) => { const th = document.createElement('th'); th.textContent = c.label; hr.append(th); });
    thead.append(hr);
    const tbody = document.createElement('tbody');
    fillRows(tbody, rows, cols);
    table.append(thead, tbody);
  };
  dash.draw('gen-group', (table) => { if (!S.group || !S.by) { table.replaceChildren(); return; } fillTable(table, groupCols(), S.group.slice(0, 500)); });
  dash.draw('gen-preview', (table) => {
    if (!S.table) { table.replaceChildren(); return; }
    fillTable(table, S.table.columns.map((c) => ({ key: c, label: c, type: 'text' })), S.table.rows.slice(0, 50));
  });

  dash.draw('saved-list', (ul) => {
    ul.replaceChildren();
    S.saved.forEach((d) => {
      const li = document.createElement('li');
      li.className = 'ito-saved-item' + (S.savedId === d.id ? ' is-open' : '');
      const info = document.createElement('div');
      const t = document.createElement('div'); t.className = 'ito-saved-name'; t.textContent = d.name || 'Saved report';
      const m = document.createElement('div'); m.className = 'ito-sub';
      const who = d.by ? S.names[d.by] || 'Someone' : 'Someone';
      const what = d.kind === 'ito' ? fmtInt((d.summary && d.summary.attempts) || 0) + ' orders' + (d.from ? ', ' + fmtDate(d.from) + ' – ' + fmtDate(d.to) : '') : fmtInt(d.rowsCount || 0) + ' rows';
      m.textContent = what + ' · saved by ' + who + (d.at ? ', ' + fmtWhen(d.at) : '');
      info.append(t, m);
      const acts = document.createElement('div'); acts.className = 'ito-saved-acts';
      const btn = (label, act, cls) => { const b = document.createElement('button'); b.type = 'button'; b.className = 'ito-btn' + (cls ? ' ' + cls : ''); b.dataset.act = act; b.dataset.id = d.id; b.textContent = label; acts.append(b); };
      btn('Open', 'open');
      btn('Download Excel', 'download');
      if (S.canEdit) btn(S.confirmDel === d.id ? 'Confirm delete' : 'Delete', 'delete', 'is-quiet');
      li.append(info, acts);
      ul.append(li);
    });
  });

  // ---------- controls ----------

  all('.ito-file').forEach((input) => input.addEventListener('change', () => { handleFiles(input.files); input.value = ''; }));
  all('.ito-drop').forEach((zone) => {
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('is-over'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('is-over'));
    zone.addEventListener('drop', (e) => { e.preventDefault(); zone.classList.remove('is-over'); if (e.dataTransfer) handleFiles(e.dataTransfer.files); });
  });
  all('.ito-keeptest').forEach((c) => c.addEventListener('change', () => { S.keepTest = c.checked; if (S.files.length) build(); render(); }));
  all('[data-dl]').forEach((b) => b.addEventListener('click', () => { const f = DOWNLOADS[b.dataset.dl]; if (f && S.kind) f(); }));
  all('.ito-save').forEach((b) => b.addEventListener('click', saveShared));
  all('.ito-clear').forEach((b) => b.addEventListener('click', () => { S.files = []; S.kind = null; S.savedId = null; S.error = ''; S.errorKind = ''; render(); }));
  all('[data-draw="saved-list"]').forEach((ul) => ul.addEventListener('click', (e) => {
    const b = e.target.closest('button[data-act]');
    if (!b) return;
    if (b.dataset.act === 'open') openSaved(b.dataset.id);
    else if (b.dataset.act === 'download') downloadSaved(b.dataset.id);
    else if (b.dataset.act === 'delete') deleteSaved(b.dataset.id);
  }));

  // dashboard tab: download what the filters show
  const overviewXlsx = () => {
    const q = window.ito.p();
    const cube = dash.data('cube'), plans = dash.data('plans'), carts = dash.data('carttype');
    if (cube.status !== 'ok' || plans.status !== 'ok' || carts.status !== 'ok') return;
    const now = cube.data.find((r) => r.k === q.month + '|' + q.subject + '|' + q.tier);
    if (!now) return;
    const view = [
      { metric: 'Month', value: monthLabel(q.month), type: 'text' },
      { metric: 'Subject', value: q.subject === 'all' ? 'All subjects' : q.subject === 'Social' ? 'Social Studies' : q.subject, type: 'text' },
      { metric: 'Plan tier', value: q.tier === 'all' ? 'All tiers' : q.tier, type: 'text' },
    ];
    const byMonth = cube.data.filter((r) => r.subject === q.subject && r.tier === q.tier).sort((a, b) => (a.month === 'all' ? 1 : b.month === 'all' ? -1 : a.month < b.month ? -1 : 1));
    const planRows = plans.data.filter((r) => r.month === q.month && (q.subject === 'all' || r.subject === q.subject) && (q.tier === 'all' || r.tier === q.tier)).sort((a, b) => b.paid - a.paid || b.attempts - a.attempts);
    const cartRows = carts.data.filter((r) => r.month === q.month && r.attempts > 0);
    const book = toXlsx([
      summarySheet(summaryRows(now, view)),
      { name: 'By month', columns: COLS.month, rows: withMonthLabel(byMonth) },
      { name: 'Plans', columns: COLS.plans, rows: planRows },
      { name: 'Cart types', columns: COLS.carts, rows: cartRows },
    ]);
    const tag = [q.month === 'all' ? 'Jun-Oct_2026' : q.month, q.subject !== 'all' ? q.subject : '', q.tier !== 'all' ? q.tier : ''].filter(Boolean).join('_');
    save('ITO_orders_report_' + tag + '.xlsx', new Blob([book]));
  };
  const overviewPlansCsv = () => {
    const q = window.ito.p();
    const plans = dash.data('plans');
    if (plans.status !== 'ok') return;
    const rows = plans.data.filter((r) => r.month === q.month && (q.subject === 'all' || r.subject === q.subject) && (q.tier === 'all' || r.tier === q.tier)).sort((a, b) => b.paid - a.paid || b.attempts - a.attempts);
    save('ITO_plans_' + (q.month === 'all' ? 'Jun-Oct_2026' : q.month) + '.csv', toCSV(COLS.plans, rows));
  };
  all('.ito-dl-overview').forEach((b) => b.addEventListener('click', overviewXlsx));
  all('.ito-dl-plans').forEach((b) => b.addEventListener('click', overviewPlansCsv));
  dash.onData(() => {
    const ok = ['cube', 'plans', 'carttype'].every((id) => dash.data(id).status === 'ok');
    all('.ito-dl-overview, .ito-dl-plans').forEach((b) => { b.disabled = !ok; });
  });

  // tabs
  const showTab = (name) => {
    all('[data-panel]').forEach((p) => { p.hidden = p.dataset.panel !== name; });
    all('[data-tab]').forEach((t) => { t.setAttribute('aria-selected', String(t.dataset.tab === name)); t.tabIndex = t.dataset.tab === name ? 0 : -1; });
  };
  all('[data-tab]').forEach((t) => t.addEventListener('click', () => { showTab(t.dataset.tab); dash.setLink(t.dataset.tab); }));
  all('[role="tablist"]').forEach((list) => list.addEventListener('keydown', (e) => {
    if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
    const tabs = [...list.querySelectorAll('[data-tab]')];
    const i = tabs.indexOf(document.activeElement);
    if (i < 0) return;
    const next = tabs[(i + (e.key === 'ArrowRight' ? 1 : tabs.length - 1)) % tabs.length];
    next.focus(); next.click();
  }));
  const linked = location.hash.slice(1);
  if ([...all('[data-panel]')].some((p) => p.dataset.panel === linked)) showTab(linked);

  sync();

  // capabilities light up when the viewer's page grants them
  (async () => {
    const use = window.claude && typeof window.claude.use === 'function' ? (n) => window.claude.use(n).catch(() => null) : async () => null;
    const [dl, db, user] = await Promise.all([use('downloads'), use('db'), use('user')]);
    S.dl = dl; S.db = db; S.userApi = user;
    if (user) {
      S.canEdit = await user.canEdit();
      S.uid = await user.id();
    }
    if (!db) { S.savedState = 'off'; render(); return; }
    db.collection('reports').orderBy('at', 'desc').limit(100).onSnapshot(async (snap) => {
      S.saved = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      S.savedState = S.saved.length ? 'list' : 'empty';
      render();
      const ids = [...new Set(S.saved.map((d) => d.by).filter(Boolean))];
      if (user && ids.length) {
        const ps = await user.profiles(ids);
        ids.forEach((id) => { S.names[id] = (ps[id] && ps[id].name) || 'Someone'; });
        dash.draw('saved-list');
      }
    }, () => { S.savedState = 'off'; render(); });
    render();
  })();
})();
