// Department pages beside Orders & payments: Dispatch (the study-material dispatch sheet), MIS and Support Team for Schools
// (any sheet they upload). The top half reads and analyses sheets and also runs in Node (scripts/export_dispatch.js);
// the rest draws the pages.
(() => {
  const root = typeof window !== 'undefined' ? window : globalThis;
  const R = root.itoReports;

  const key = (s) => String(s ?? '').toLowerCase().replace(/[^a-z0-9]+/g, '');
  const txt = (v) => (v === null || v === undefined ? '' : String(v).trim());
  const num = (v) => { if (typeof v === 'number') return Number.isFinite(v) ? v : 0; const n = Number(txt(v).replace(/,/g, '')); return Number.isFinite(n) ? n : 0; };

  // ---------- dispatch sheets ----------

  // Header names as the dispatch team writes them ("COURIERName", "ORDER DESPATCH DATE", "NO-OFPYQP'S"…), compared without spaces or punctuation.
  const DSP_COLS = {
    sno: (k) => /^(sno|srno|slno|serialno|sno\w{0,3})$/.test(k),
    order_date: (k) => k === 'date' || k === 'orderdate' || k === 'orderreceiveddate' || k === 'receiveddate',
    courier: (k) => k.includes('courier'),
    pod: (k) => k === 'pod' || k.startsWith('podno') || k.includes('awb') || k.includes('tracking') || k.includes('docket') || k.includes('consignment'),
    school: (k) => k.includes('schoolcode') || k === 'school' || k === 'schoolid',
    pyqps: (k) => k.includes('pyq') || k.includes('questionpaper') || /noofqps?$/.test(k),
    guides: (k) => k.includes('guide'),
    textbooks: (k) => k.includes('textbook'),
    order_status: (k) => k === 'orderstatus',
    dispatch_date: (k) => (k.includes('despatch') || k.includes('dispatch')) && k.includes('date'),
    delivery_date: (k) => k.includes('delivery') && k.includes('date'),
    days: (k) => k === 'days' || k === 'transitdays' || k === 'deliverydays',
    status: (k) => k === 'status' || k === 'deliverystatus' || k === 'currentstatus',
  };
  const findDispatchCols = (columns) => {
    const ks = columns.map(key);
    const c = {};
    for (const [name, test] of Object.entries(DSP_COLS)) { const i = ks.findIndex(test); c[name] = i < 0 ? null : columns[i]; }
    c.workbooks = columns.filter((col, i) => ks[i].includes('workbook'));
    const used = new Set(Object.values(c).flat());
    const si = c.status ? columns.indexOf(c.status) : -1;
    // Notes sit in unnamed columns right of the status ("RTO", "dtdc to speed", a new AWB number).
    c.notes = columns.filter((col, i) => !used.has(col) && (/remark|note|comment|rto|reason/.test(ks[i]) || (si >= 0 && i > si && /^column\d+$/.test(ks[i]))));
    return c;
  };
  const isDispatch = (columns) => { const c = findDispatchCols(columns); return !!(c.courier && c.school && (c.pod || c.dispatch_date)); };
  const courierName = (v) => { const t = txt(v).replace(/\s+/g, ' '); if (!t) return ''; const u = t.toUpperCase(); return u.length <= 4 ? u : u[0] + u.slice(1).toLowerCase(); };
  const issueOf = (s) => (!s ? '' : /lost/i.test(s) ? 'Lost' : /rt[io]?o|return/i.test(s) ? 'Returned (RTO)' : /speed|re-?dispatch|\bsen[dt]\b|no ?ser/i.test(s) ? 'Re-routed' : '');

  const normalizeDispatch = (table) => {
    const c = findDispatchCols(table.columns);
    const seen = new Set();
    const out = [];
    table.rows.forEach((r, i) => {
      const school = txt(r[c.school]).toUpperCase().replace(/\s+/g, '');
      if (!school) return;
      const exam = /^LC/.test(school) ? 'LC' : 'ITO';
      const courier = courierName(r[c.courier]);
      const pod = txt(r[c.pod]);
      const date = (col) => (col ? R.toISODate(r[col]) : '');
      const order_date = date(c.order_date), dispatch_date = date(c.dispatch_date), delivery_date = date(c.delivery_date);
      const status = txt(r[c.status]);
      const os = txt(r[c.order_status]);
      const wb = c.workbooks.map((w) => num(r[w]));
      let itoWb = 0, lcWb = 0;
      if (wb.length >= 2) { itoWb = wb[0]; lcWb = wb.slice(1).reduce((a, b) => a + b, 0); }
      else if (wb.length === 1) { if (/lc|little/.test(key(c.workbooks[0])) || exam === 'LC') lcWb = wb[0]; else itoWb = wb[0]; }
      const pyqps = num(r[c.pyqps]), guides = num(r[c.guides]), textbooks = num(r[c.textbooks]);
      const shipped = !!(courier || pod);
      const delivered = shipped && (!!delivery_date || (/deliver/i.test(status) && !/not|undeliver|pending|fail/i.test(status)));
      const remark = c.notes.map((n) => txt(r[n])).filter((v) => v && !/^deliver/i.test(v)).join(' · ');
      const daysRaw = c.days && txt(r[c.days]) !== '' ? Number(r[c.days]) : NaN;
      let sno = txt(r[c.sno]) || String(i + 1);
      while (seen.has(sno)) sno += '·';
      seen.add(sno);
      out.push({
        sno, school, exam, region: school.replace(/^(LC|WF)/, '').slice(0, 2), order_date, month: (order_date || dispatch_date).slice(0, 7) || 'unknown',
        courier, pod, ito_workbooks: itoWb, pyqps, guides, lc_workbooks: lcWb, textbooks, units: itoWb + lcWb + pyqps + guides + textbooks,
        order_status: /hold/i.test(os) ? 'Hold' : os ? 'Dispatch' : '', dispatch_date, delivery_date, days: Number.isFinite(daysRaw) ? Math.round(daysRaw) : null,
        stage: delivered ? 'Delivered' : shipped ? 'In transit' : /hold/i.test(os) ? 'On hold' : 'Not shipped yet',
        issue: issueOf(remark), remark,
      });
    });
    return out;
  };

  // The "Total Order Summary" sheet: one block of rows per programme (ITO, then LC - Little Champ).
  const isMaterial = (columns) => { const ks = columns.map(key); return ks.includes('description') && ks.some((k) => k === 'count' || k === 'total') && ks.includes('packed'); };
  const ITEM = [[/order/i, 'School orders'], [/workbook/i, 'Workbooks'], [/qp|question/i, 'Question papers'], [/guide/i, 'Guides'], [/text ?book/i, 'Textbooks'], [/combo/i, 'Combo packs']];
  const normalizeMaterial = (table) => {
    const col = (re) => table.columns.find((c) => re.test(key(c)));
    const d = col(/^description$/), cnt = col(/^(count|total)$/), pk = col(/^packed$/), today = col(/today/), pend = col(/^pending$/), first = table.columns[0];
    let exam = 'ITO';
    const out = [];
    table.rows.forEach((r) => {
      const label = txt(r[d]);
      if (!label) { const s = txt(r[first]); if (s && !/^\d+$/.test(s)) exam = /lc|little/i.test(s) ? 'LC' : 'ITO'; return; }
      if (txt(r[cnt]) === '') return;
      const item = (ITEM.find(([re]) => re.test(label)) || [null, label])[1];
      out.push({ k: exam + '|' + item, exam, item, count: num(r[cnt]), packed: num(r[pk]), dispatch_today: num(r[today]), pending: num(r[pend]) });
    });
    return out;
  };

  // A whole workbook (or CSV): the dispatch rows from its dispatch sheet, and the material summary if it has one.
  const dispatchBook = (file) => {
    const det = file.sheets.find((s) => isDispatch(s.columns));
    if (!det) return null;
    const mat = file.sheets.find((s) => isMaterial(s.columns));
    return { rows: normalizeDispatch(det), material: mat ? normalizeMaterial(mat) : [], sheet: det.name };
  };

  // ---------- any other sheet (MIS, Support) ----------

  const PERSONAL = /e-?mail|phone|mobile|whats ?app|contact|aadha?r|address|password|student ?name|parent|father|mother|guardian|^name$/i;
  const IDLIKE = /^(s\.?\s?no\.?|sr\.?\s?no\.?|sl\.?\s?no\.?|serial.*|id|.*\bid|.*\bcode|.*phone|.*mobile|pin ?code|zip|roll ?no\.?|.*number|year)$/i;
  const DONE = /done|closed|resolv|complet|deliver|success|captur|\bpaid\b|solved|approv|finish|convert|\bwon\b|^yes$|dispatched/i;
  const NOTDONE = /\bnot\b|\bun|pending|progress|\bopen|re-?open|fail|reject|cancel|hold/i;
  const START = /creat|open|raise|receiv|request|regist|order|log|start|submit|^date$/i;
  const END = /resol|close|complet|deliver|done|end|finish|dispatch|despatch|reply|respon/i;
  const CAT = /type|categor|issue|priority|channel|source|team|agent|assign|owner|depart|city|state|region|zone|board|class|grade|subject|product|plan|courier|mode|reason|level/i;
  const isDoneValue = (v) => { const s = txt(v); return !!s && DONE.test(s) && !NOTDONE.test(s); };
  const median = (a) => { if (!a.length) return null; const s = [...a].sort((x, y) => x - y); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
  const mean = (a) => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : null);

  const analyze = (table) => {
    const rows = table.rows, n = rows.length;
    const prof = R.profile(table);
    const cols = prof.map((p) => {
      const nameDate = /date|time|\bday\b|created|opened|closed|resolved|\bon$|\bat$/i.test(p.column);
      const dateOf = (v) => {
        if (v === '' || v === null || v === undefined) return '';
        if (!nameDate && (typeof v === 'number' || /^\d+(\.\d+)?$/.test(String(v).trim()))) return '';
        return R.toISODate(v);
      };
      const vals = rows.map((r) => r[p.column]).filter((v) => v !== '' && v !== null && v !== undefined);
      const isDate = vals.length > 0 && vals.filter((v) => dateOf(v)).length / vals.length >= 0.8;
      const nums = p.type === 'Number' ? vals.map(R.asNumber).filter((x) => x !== null) : [];
      const serial = nums.length > 20 && p.distinct === p.filled && nums.every(Number.isInteger) && Math.min(...nums) <= 1 && Math.max(...nums) - Math.min(...nums) + 1 === nums.length;
      const bigId = nums.length > 0 && nums.every((x) => Math.abs(x) >= 1e9);
      return { ...p, isDate, dateOf, personal: PERSONAL.test(p.column), idLike: IDLIKE.test(p.column.trim()) || serial || bigId };
    });
    const dateCols = cols.filter((c) => c.isDate);
    const dateCol = dateCols.find((c) => START.test(c.column) && !END.test(c.column)) || dateCols.find((c) => !END.test(c.column)) || dateCols[0] || null;
    const endCol = dateCols.find((c) => c !== dateCol && END.test(c.column)) || null;
    const textCols = cols.filter((c) => c.type !== 'Number' && !c.isDate && !c.personal);
    const statusCol = textCols.find((c) => /status|stage|resolution|result|outcome|progress/i.test(c.column) && c.distinct <= 25) || null;
    const catCols = textCols.filter((c) => c !== statusCol && c.distinct >= 2 && c.distinct <= 40 && c.distinct < n && c.filled >= n * 0.3)
      .sort((a, b) => (CAT.test(b.column) - CAT.test(a.column)) || a.distinct - b.distinct).slice(0, 6);
    const numCols = cols.filter((c) => c.type === 'Number' && !c.idLike && !c.personal && !c.isDate).slice(0, 8);

    const starts = dateCol ? rows.map((r) => dateCol.dateOf(r[dateCol.column])) : [];
    const ends = endCol ? rows.map((r) => endCol.dateOf(r[endCol.column])) : [];
    const doneRow = (r) => (statusCol ? isDoneValue(r[statusCol.column]) : false);
    const byMonth = new Map();
    starts.forEach((d, i) => {
      if (!d) return;
      const g = byMonth.get(d.slice(0, 7)) || { month: d.slice(0, 7), rows: 0, done: 0 };
      g.rows++; if (doneRow(rows[i])) g.done++;
      byMonth.set(g.month, g);
    });
    const sortedDates = starts.filter(Boolean).sort();
    const durations = endCol ? starts.map((d, i) => (d && ends[i] ? Math.round((Date.parse(ends[i]) - Date.parse(d)) / 864e5) : null)).filter((v) => v !== null && v >= 0) : [];
    const groupTop = (col, limit) => {
      const m = new Map();
      rows.forEach((r) => {
        const v = txt(r[col]) || '(blank)';
        const g = m.get(v) || { value: v, rows: 0, done: 0 };
        g.rows++; if (doneRow(r)) g.done++;
        m.set(v, g);
      });
      const list = [...m.values()].sort((a, b) => b.rows - a.rows);
      const top = list.slice(0, limit);
      if (list.length > limit) top.push(list.slice(limit).reduce((a, g) => ({ value: 'Other (' + (list.length - limit) + ' more)', rows: a.rows + g.rows, done: a.done + g.done }), { rows: 0, done: 0 }));
      return top.map((g) => ({ ...g, share: n ? g.rows / n : 0, done_share: g.rows ? g.done / g.rows : null }));
    };
    const statusFilled = statusCol ? rows.filter((r) => txt(r[statusCol.column])).length : 0;
    const done = statusCol ? rows.filter(doneRow).length : null;
    return {
      rows: n, columns: table.columns.length,
      from: sortedDates[0] || '', to: sortedDates[sortedDates.length - 1] || '',
      dateCol: dateCol ? dateCol.column : null, endCol: endCol ? endCol.column : null, statusCol: statusCol ? statusCol.column : null,
      done, statusFilled, doneRate: statusFilled ? done / statusFilled : null, open: statusCol ? statusFilled - done : null,
      avgClose: mean(durations), medianClose: median(durations), closedCount: durations.length,
      byMonth: [...byMonth.values()].sort((a, b) => (a.month < b.month ? -1 : 1)),
      byStatus: statusCol ? groupTop(statusCol.column, 20).map((g) => ({ ...g, is_done: isDoneValue(g.value) })) : [],
      categories: catCols.map((c) => ({ column: c.column, distinct: c.distinct, values: groupTop(c.column, 10) })),
      numbers: numCols.map((c) => {
        const vals = rows.map((r) => R.asNumber(r[c.column])).filter((x) => x !== null);
        return { column: c.column, filled: vals.length, total: vals.reduce((a, b) => a + b, 0), avg: mean(vals), min: vals.length ? Math.min(...vals) : null, max: vals.length ? Math.max(...vals) : null };
      }),
      profile: prof.slice(0, 300).map((p) => ({ ...p, personal: PERSONAL.test(p.column) })),
    };
  };

  root.itoDepts = { findDispatchCols, isDispatch, normalizeDispatch, isMaterial, normalizeMaterial, dispatchBook, analyze, isDoneValue };
  if (typeof document === 'undefined' || typeof dash === 'undefined') return;

  // ---------- shared page helpers ----------

  const { fmtInt, fmtPct, tip, textWidth, enc } = window.ito;
  const C = dash.colors;
  const all = (sel) => document.querySelectorAll(sel);
  const f1 = d3.format('.1f');
  const fmtDate = (iso) => { if (!iso) return ''; const d = new Date(String(iso).slice(0, 10) + 'T00:00:00Z'); return Number.isNaN(d.getTime()) ? String(iso) : d3.utcFormat('%-d %b %Y')(d); };
  const fmtWhen = (iso) => { const d = new Date(iso); return Number.isNaN(d.getTime()) ? '' : d3.timeFormat('%-d %b %Y, %H:%M')(d); };
  const fmtv = (v, kind) => {
    if (v === null || v === undefined || v === '' || (typeof v === 'number' && Number.isNaN(v))) return '—';
    if (kind === 'pct') return fmtPct(v);
    if (kind === 'days') return f1(v) + (v === 1 ? ' day' : ' days');
    if (kind === 'num') return typeof v === 'number' ? d3.format(',.4~f')(v) : String(v);
    if (kind === 'date') return fmtDate(v);
    if (kind === 'text') return String(v);
    return typeof v === 'number' ? fmtInt(v) : String(v);
  };
  const fileStem = (s) => String(s).replace(/\.[^.]+$/, '').replace(/[^\w\-. ]+/g, ' ').trim().replace(/\s+/g, '_').slice(0, 60) || 'report';
  const CAP = { dl: null, db: null, canEdit: false, uid: null, names: {} };
  const save = async (filename, data) => {
    const blob = data instanceof Blob ? data : new Blob([data]);
    if (CAP.dl) {
      try { await CAP.dl.save({ filename, data: blob }); return; }
      catch (e) { if (e && (e.code === 'declined' || e.code === 'rate_limited')) return; }
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.rel = 'noopener';
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  };
  const readAll = async (list) => { const out = []; for (const f of list) out.push(await R.readFile(f)); return out; };
  const errorOf = (e) => (e && e.code === 'xls' ? ['xls', 'xls'] : ['error', (e && e.message) || 'This file could not be read.']);
  const fillBody = (tbody, rows, cols, keyOf) => {
    tbody.replaceChildren();
    rows.forEach((r) => {
      const tr = document.createElement('tr');
      if (keyOf) tr.dataset.key = keyOf(r);
      if (r.__cls) tr.className = r.__cls;
      cols.forEach((c) => { const td = document.createElement('td'); td.textContent = c.show ? c.show(r) : fmtv(r[c.key], c.type); if (c.cls) td.className = c.cls; tr.append(td); });
      tbody.append(tr);
    });
  };
  const fillTable = (table, cols, rows) => {
    table.replaceChildren();
    const thead = document.createElement('thead');
    const hr = document.createElement('tr');
    cols.forEach((c) => { const th = document.createElement('th'); th.textContent = c.label; hr.append(th); });
    thead.append(hr);
    const tbody = document.createElement('tbody');
    fillBody(tbody, rows, cols);
    table.append(thead, tbody);
  };
  // Horizontal bars, 0–1 each; a row is clickable when onClick is given.
  const hbars = (el, rows, o) => {
    const W = Math.max(260, el.clientWidth);
    const probe = d3.select(el).append('svg');
    const labelW = (d3.max(rows, (r) => textWidth(probe, o.label(r))) || 0) + 12;
    const right = (d3.max(rows, (r) => textWidth(probe, o.text(r))) || 0) + 10;
    probe.remove();
    const stacked = W - labelW - right < 120;
    const rowH = stacked ? 44 : 34, H = rows.length * rowH + 8;
    const left = stacked ? 0 : labelW;
    const svg = d3.select(el).append('svg').attr('width', '100%').attr('viewBox', `0 0 ${W} ${H}`).attr('role', 'img').attr('aria-label', o.aria);
    const x = d3.scaleLinear([0, 1], [left, W - right]);
    rows.forEach((r, i) => {
      const yy = i * rowH + 4 + (stacked ? 16 : 0);
      const g = svg.append('g');
      if (o.onClick) g.style('cursor', 'pointer').on('click', () => o.onClick(r));
      if (o.picked && o.picked(r)) g.append('rect').attr('x', 0).attr('y', yy - (stacked ? 18 : 2)).attr('width', W).attr('height', rowH - 4).attr('rx', 4).attr('fill', 'var(--ito-tint)');
      if (stacked) g.append('text').attr('x', 0).attr('y', yy - 5).attr('font-size', 12).attr('fill', 'var(--color-fg-muted)').text(o.label(r));
      else g.append('text').attr('x', left - 8).attr('y', yy + 13).attr('dy', '0.32em').attr('text-anchor', 'end').attr('font-size', 12).attr('fill', 'var(--color-fg-muted)').text(o.label(r));
      g.append('rect').attr('x', left).attr('y', yy + 2).attr('width', x(1) - left).attr('height', stacked ? 18 : 22).attr('rx', 4).attr('fill', 'var(--cds-chart-grid)');
      g.append('rect').attr('x', left).attr('y', yy + 2).attr('width', Math.max(1, x(Math.max(0, Math.min(1, o.value(r) || 0))) - left)).attr('height', stacked ? 18 : 22).attr('rx', 4).attr('fill', o.color(r));
      g.append('text').attr('x', x(1) + 6).attr('y', yy + (stacked ? 11 : 13)).attr('dy', '0.32em').attr('font-size', 12).attr('fill', 'var(--color-fg)').text(o.text(r));
    });
  };
  // Monthly bars: a grey total with a coloured part (delivered, completed) in front.
  const monthBars = (el, rows, o) => {
    const W = Math.max(280, el.clientWidth), H = 240;
    const svg = d3.select(el).append('svg').attr('width', '100%').attr('viewBox', `0 0 ${W} ${H}`).attr('role', 'img').attr('aria-label', o.aria);
    const y = d3.scaleLinear([0, (d3.max(rows, (r) => r.total) || 1) * 1.15], [H - 28, 14]).nice();
    const left = textWidth(svg, fmtInt(y.domain()[1])) + 12;
    const x = d3.scaleBand(rows.map((r) => r.month), [left, W - 8]).padding(0.35);
    const muted = 'var(--color-fg-muted)';
    svg.append('g').selectAll('line').data(y.ticks(4)).join('line').attr('x1', left).attr('x2', W - 8).attr('y1', (d) => y(d)).attr('y2', (d) => y(d)).attr('stroke', 'var(--cds-chart-grid)');
    svg.append('g').selectAll('text').data(y.ticks(4)).join('text').attr('x', left - 6).attr('y', (d) => y(d)).attr('dy', '0.32em').attr('text-anchor', 'end').attr('font-size', 12).attr('fill', muted).text((d) => fmtInt(d));
    const every = Math.max(1, Math.ceil(rows.length / Math.max(2, Math.floor((W - left) / 64))));
    svg.append('g').selectAll('text').data(rows.filter((r, i) => i % every === 0)).join('text').attr('x', (r) => x(r.month) + x.bandwidth() / 2).attr('y', H - 8).attr('text-anchor', 'middle').attr('font-size', 12).attr('fill', muted).text((r) => R.monthLabel(r.month).replace(/ 20(\d\d)$/, " '$1"));
    const bw = Math.min(36, x.bandwidth());
    const bx = (r) => x(r.month) + (x.bandwidth() - bw) / 2;
    const fade = (r) => (o.picked && o.picked !== 'all' && o.picked !== r.month ? 0.35 : 1);
    svg.append('g').selectAll('rect').data(rows).join('rect').attr('x', bx).attr('width', bw).attr('y', (r) => y(r.total)).attr('height', (r) => y(0) - y(r.total)).attr('rx', 4)
      .attr('fill', o.part ? 'var(--cds-chart-muted)' : C[0]).attr('opacity', (r) => (o.part ? 0.6 : 1) * fade(r));
    if (o.part) svg.append('g').selectAll('rect').data(rows).join('rect').attr('x', bx).attr('width', bw).attr('y', (r) => y(r.part)).attr('height', (r) => y(0) - y(r.part)).attr('rx', 4).attr('fill', C[0]).attr('opacity', fade);
    if (bw >= 22 || rows.length <= 8) svg.append('g').selectAll('text').data(rows).join('text').attr('x', (r) => x(r.month) + x.bandwidth() / 2).attr('y', (r) => y(r.total) - 6).attr('text-anchor', 'middle').attr('font-size', 12).attr('fill', 'var(--color-fg)').text((r) => fmtInt(r.total));
    const tt = tip(el);
    svg.append('g').selectAll('rect').data(rows).join('rect').attr('x', (r) => x(r.month)).attr('width', x.bandwidth()).attr('y', 0).attr('height', H - 24).attr('fill', 'transparent').style('cursor', o.onClick ? 'pointer' : 'default')
      .on('pointermove pointerdown', (event, r) => tt.show(event, R.monthLabel(r.month), o.lines(r)))
      .on('pointerleave', () => tt.hide())
      .on('click', (event, r) => o.onClick && o.onClick(r));
  };
  const savedItem = (ul, d, open, acts) => {
    const li = document.createElement('li');
    li.className = 'ito-saved-item' + (open ? ' is-open' : '');
    const info = document.createElement('div');
    const t = document.createElement('div'); t.className = 'ito-saved-name'; t.textContent = d.name || 'Saved report';
    const m = document.createElement('div'); m.className = 'ito-sub';
    const who = d.by ? CAP.names[d.by] || 'Someone' : 'Someone';
    m.textContent = fmtInt(d.rowsCount || 0) + ' rows · saved by ' + who + (d.at ? ', ' + fmtWhen(d.at) : '');
    info.append(t, m);
    const box = document.createElement('div'); box.className = 'ito-saved-acts';
    acts.forEach(([label, act, cls]) => { const b = document.createElement('button'); b.type = 'button'; b.className = 'ito-btn' + (cls ? ' ' + cls : ''); b.dataset.act = act; b.dataset.id = d.id; b.textContent = label; box.append(b); });
    li.append(info, box);
    ul.append(li);
  };
  const compact = (results) => Object.fromEntries(Object.entries(results).map(([id, rows]) => {
    const c = rows.length ? Object.keys(rows[0]) : [];
    return [id, { c, r: rows.map((row) => c.map((k) => (typeof row[k] === 'number' && !Number.isInteger(row[k]) ? Math.round(row[k] * 1e4) / 1e4 : row[k] ?? null))) }];
  }));
  const expand = (packed) => Object.fromEntries(Object.entries(packed || {}).map(([id, { c, r }]) => [id, (r || []).map((v) => Object.fromEntries(c.map((k, i) => [k, v[i]])))]));

  // ---------- departments and tabs ----------

  const lastTab = {};
  const deptOf = (tab) => tab.dataset.dept || 'orders';
  all('[data-dept-pick]').forEach((b) => b.addEventListener('click', () => {
    const d = b.dataset.deptPick;
    const tabs = [...all('[data-tab]')].filter((t) => deptOf(t) === d);
    const t = tabs.find((x) => x.dataset.tab === lastTab[d]) || tabs[0];
    if (t) t.click();
  }));
  const DSP_FILTER_TABS = ['dsp-overview', 'dsp-couriers', 'dsp-regions', 'dsp-followup'];
  const syncDspFilters = (name) => { all('.dsp-filters').forEach((f) => { f.hidden = !DSP_FILTER_TABS.includes(name); }); };
  all('[data-tab]').forEach((t) => t.addEventListener('click', () => { lastTab[deptOf(t)] = t.dataset.tab; syncDspFilters(t.dataset.tab); }));
  const goTab = (name) => { const t = [...all('[data-tab]')].find((x) => x.dataset.tab === name); if (t) t.click(); };
  all('[data-goto]').forEach((b) => b.addEventListener('click', () => goTab(b.dataset.goto)));
  syncDspFilters(location.hash.slice(1));

  // ---------- date ranges (Orders and Dispatch) ----------

  const DAY = 864e5;
  const isoAdd = (iso, n) => new Date(Date.parse(iso + 'T00:00:00Z') + n * DAY).toISOString().slice(0, 10);
  const spanDays = (a, b) => Math.round((Date.parse(b) - Date.parse(a)) / DAY) + 1;
  const inRange = (d, from, to) => !!d && (!from || d >= from) && (!to || d <= to);
  const fmtShort = (iso) => d3.utcFormat('%-d %b')(new Date(iso + 'T00:00:00Z'));
  const RANGE = {
    orders: { from: 'from', to: 'to', reset: { month: 'all' } },
    dispatch: { from: 'dfrom', to: 'dto', reset: { dmonth: 'all' } },
  };
  // The first and last date in the data, from the per-day results (small, and already loaded where the dates show).
  RANGE.orders.bounds = () => {
    const d = dash.data('daily');
    if (d.status !== 'ok' || !d.data.length) return null;
    const ds = d.data.filter((r) => r.subject === 'all' && r.tier === 'all').map((r) => r.date);
    return ds.length ? { min: d3.min(ds), max: d3.max(ds) } : null;
  };
  RANGE.dispatch.bounds = () => {
    const d = baseD('dsp_daily');
    if (d.status !== 'ok' || !d.data.length) return null;
    const ds = d.data.filter((r) => r.exam === 'all' && r.courier === 'all').map((r) => r.date);
    return ds.length ? { min: d3.min(ds), max: d3.max(ds) } : null;
  };
  const rangeOf = (scope) => { const q = dash.params(); const c = RANGE[scope]; return { from: q[c.from] || '', to: q[c.to] || '' }; };
  const rangeText = (scope) => { const r = rangeOf(scope); return r.from || r.to ? (r.from ? fmtShort(r.from) : 'start') + ' – ' + (r.to ? fmtDate(r.to) : 'latest') : ''; };
  const presetRange = (preset, b) => {
    if (!b || preset === 'all') return { from: null, to: null };
    if (preset === 'month') return { from: b.max.slice(0, 8) + '01', to: b.max };
    return { from: isoAdd(b.max, -(Number(preset) - 1)), to: b.max };
  };
  const PRESET = { orders: 'all', dispatch: 'all' };
  const setRange = (scope, from, to, preset) => {
    const c = RANGE[scope];
    if (from && to && from > to) [from, to] = [to, from];
    PRESET[scope] = preset || (from || to ? '' : 'all');
    dash.setParams({ [c.from]: from || null, [c.to]: to || null, ...(from || to ? c.reset : {}) });
  };

  dash.draw('date-input', (el) => {
    const scope = el.dataset.range, c = RANGE[scope];
    if (!c) return;
    const b = c.bounds();
    if (b) { el.min = b.min; el.max = b.max; }
    const key = c[el.dataset.end];
    if (el !== document.activeElement) el.value = dash.params()[key] || '';
    el.classList.toggle('is-set', !!el.value);
    el.onchange = () => {
      const r = rangeOf(scope);
      if (!el.value) setRange(scope, el.dataset.end === 'from' ? '' : r.from, el.dataset.end === 'to' ? '' : r.to);
      else if (el.validity.valid && /^\d{4}-\d{2}-\d{2}$/.test(el.value)) setRange(scope, el.dataset.end === 'from' ? el.value : r.from, el.dataset.end === 'to' ? el.value : r.to);
    };
    el.onblur = () => { el.value = dash.params()[key] || ''; };
  });
  all('.ito-field-presets[data-range] [data-preset]').forEach((b) => b.addEventListener('click', () => {
    const scope = b.closest('[data-range]').dataset.range;
    const r = presetRange(b.dataset.preset, RANGE[scope].bounds());
    setRange(scope, r.from, r.to, b.dataset.preset);
  }));
  all('.dsp-basis').forEach((sel) => sel.addEventListener('change', () => dash.setParams({ dbasis: sel.value })));
  dash.onData(() => {
    const q = dash.params();
    all('.ito-field-presets[data-range]').forEach((box) => {
      const scope = box.dataset.range;
      const r = rangeOf(scope);
      const on = r.from || r.to ? PRESET[scope] : 'all';
      box.querySelectorAll('[data-preset]').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.preset === on)));
    });
    all('.dsp-basis').forEach((sel) => { sel.value = q.dbasis || 'order'; });
    const ranged = (q.from || q.to) && (q.month || 'all') === 'all';
    all('[data-period-label]').forEach((el) => { el.hidden = el.dataset.periodLabel !== (ranged ? 'range' : 'month'); });
  });

  // Every calc of the page (window.calcDefs), run in page order over the inputs given; calcs whose inputs are missing are skipped.
  const runDefs = (inputs) => {
    const res = { ...inputs };
    Object.entries(window.calcDefs || {}).forEach(([id, def]) => {
      if (!def.inputs.every((i) => res[i] !== undefined)) return;
      try { res[id] = def.fn(...def.inputs.map((i) => res[i])); } catch (e) { console.error('calc ' + id, e); res[id] = []; }
    });
    return res;
  };
  const asResult = (data) => ({ status: 'ok', data: data || [], columns: data && data.length ? Object.keys(data[0]) : [], refreshing: false, meta: [] });
  const DV = { okey: null, ores: null, dkey: null, dres: null };
  window.dateView = {
    rangeText,
    orders(id) {
      const q = dash.params();
      if (!q.from && !q.to) return dash.data(id);
      const src = dash.data('orders');
      if (src.status !== 'ok') return { ...src, data: [] };
      const key = q.from + '|' + q.to + '|' + src.data.length;
      if (DV.okey !== key) {
        const dateOf = (o) => String(o.date || '').slice(0, 10);
        const res = runDefs({ orders: src.data.filter((o) => inRange(dateOf(o), q.from, q.to)) });
        // The period before, of the same length, for the changes shown on the headline figures.
        const ds = src.data.map(dateOf).filter(Boolean);
        const from = q.from || d3.min(ds), to = q.to || d3.max(ds);
        const pf = isoAdd(from, -spanDays(from, to)), pt = isoAdd(from, -1);
        const prevRows = src.data.filter((o) => inRange(dateOf(o), pf, pt));
        const prev = prevRows.length && window.calcDefs.cube ? new Map(window.calcDefs.cube.fn(prevRows).filter((r) => r.month === 'all').map((r) => [r.k, r])) : new Map();
        (res.cube || []).forEach((r) => {
          if (r.month !== 'all') return;
          const p = prev.get(r.k);
          r.paid_change = p && p.paid ? r.paid / p.paid - 1 : null;
          r.success_change = p && p.success !== null && r.success !== null ? r.success - p.success : null;
          r.prev_paid = p ? p.paid : null; r.prev_attempts = p ? p.attempts : null;
        });
        res.__prev = { from: pf, to: pt, rows: prevRows.length };
        DV.okey = key; DV.ores = res;
      }
      return id === '__prev' ? DV.ores.__prev : asResult(DV.ores[id]);
    },
    dispatch(id) {
      const q = dash.params();
      if (!q.dfrom && !q.dto) return baseD(id);
      let rows, material;
      if (DSP.override) {
        if (!DSP.rows) return baseD(id);
        rows = DSP.rows; material = DSP.override.dsp_material || [];
      } else {
        const s = dash.data('dispatch');
        if (s.status !== 'ok') return { ...s, data: [] };
        const m = dash.data('dispatch_material');
        rows = s.data; material = m.status === 'ok' ? m.data : [];
      }
      const basis = q.dbasis || 'order';
      const key = [q.dfrom, q.dto, basis, rows.length, DSP.name].join('|');
      if (DV.dkey !== key) {
        const asOf = rows.reduce((mx, r) => [r.order_date, r.dispatch_date, r.delivery_date].reduce((x, d) => (d && d > x ? d : x), mx), '');
        const field = basis === 'dispatch' ? 'dispatch_date' : basis === 'delivery' ? 'delivery_date' : 'order_date';
        const sel = rows.filter((r) => inRange(r[field] || (basis === 'order' ? r.dispatch_date : ''), q.dfrom, q.dto)).map((r) => ({ ...r, sheet_as_of: asOf }));
        DV.dres = window.dspCalcs.run(sel, material);
        DV.dkey = key;
      }
      return asResult(DV.dres[id]);
    },
  };

  // ---------- Dispatch ----------

  const DSP = { override: null, rows: null, name: '', files: [], savedId: null, busy: false, error: '', errorKind: '', saveState: '', pick: 'open', query: '', sort: { key: 'days_open', dir: -1 }, rsort: { key: 'orders', dir: -1 }, rquery: '', saved: [], savedState: 'loading', confirmDel: null };
  const baseD = (id) => (DSP.override ? { status: 'ok', data: DSP.override[id] || [] } : dash.data(id));
  // With dates picked, every Dispatch figure is worked out again from the rows in that range (window.dateView.dispatch).
  const D = (id) => window.dateView.dispatch(id);
  const dq = () => { const q = dash.params(); return { month: q.dmonth || 'all', courier: q.courier || 'all', exam: q.exam || 'all' }; };
  const dk = (q) => q.exam + '|' + q.courier + '|' + q.month;
  const cubeRow = (q) => { const c = D('dsp_cube'); return c.status === 'ok' ? c.data.find((r) => r.k === dk(q)) || null : null; };
  const cite = (el, k) => { if (DSP.override) el.removeAttribute('data-where'); else el.setAttribute('data-where', 'k=' + enc(k)); };
  const REGIONS = { AP: 'Andhra Pradesh', AS: 'Assam', BR: 'Bihar', CG: 'Chhattisgarh', CH: 'Chandigarh', DL: 'Delhi', DN: 'Dadra & Nagar Haveli', GA: 'Goa', GJ: 'Gujarat', HP: 'Himachal Pradesh', HR: 'Haryana', JH: 'Jharkhand', JK: 'Jammu & Kashmir', KA: 'Karnataka', KL: 'Kerala', MH: 'Maharashtra', ML: 'Meghalaya', MN: 'Manipur', MP: 'Madhya Pradesh', OD: 'Odisha', PB: 'Punjab', PY: 'Puducherry', RJ: 'Rajasthan', SK: 'Sikkim', TN: 'Tamil Nadu', TR: 'Tripura', TS: 'Telangana', UK: 'Uttarakhand', UP: 'Uttar Pradesh', WB: 'West Bengal' };
  const regionLabel = (code) => (REGIONS[code] ? code + ' · ' + REGIONS[code] : code || '—');
  const examLabel = (e) => (e === 'LC' ? 'Little Champ' : e === 'ITO' ? 'ITO' : 'ITO and Little Champ');
  const STAGE_CLS = { Delivered: 'is-ok', 'In transit': 'is-warn', 'On hold': 'is-bad', 'Not shipped yet': 'is-bad' };

  const DSP_DRAWS = ['dsp-month-select', 'dsp-courier-select', 'dsp-scope', 'dsp-asof', 'dsp-kpi', 'dsp-pipeline', 'dsp-trend', 'dsp-courier-bars', 'dsp-courier-rows', 'dsp-daily', 'dsp-days',
    'dsp-region-rows', 'dsp-units', 'dsp-material-rows', 'dsp-open-n', 'dsp-open-count', 'dsp-open-rows', 'dsp-source', 'dsp-source-name', 'dsp-error', 'dsp-saved', 'ins-dispatch', 'date-input'];
  const dspSync = () => {
    all('[data-dsp-needs]').forEach((el) => {
      const k = el.dataset.dspNeeds;
      el.hidden = !(k === 'override' ? !!DSP.override : k === 'rows' ? !!DSP.rows : k === 'save' ? !!(DSP.override && !DSP.savedId && CAP.db && CAP.canEdit)
        : k === 'noedit' ? !!(DSP.override && !DSP.savedId && CAP.db && !CAP.canEdit) : k === 'nomaterial' ? !!DSP.override && !(DSP.override.dsp_material || []).length
        : k === 'norange' ? !!(DSP.override && !DSP.rows && (dash.params().dfrom || dash.params().dto)) : false);
    });
    all('[data-dsp-msg]').forEach((el) => { const k = el.dataset.dspMsg; el.hidden = !(k === 'reading' ? DSP.busy : !!DSP.error && DSP.errorKind === k); });
    all('[data-dsp-save]').forEach((el) => { el.hidden = DSP.saveState !== el.dataset.dspSave; });
    all('.dsp-save').forEach((b) => { b.disabled = DSP.saveState === 'saving' || DSP.saveState === 'saved'; });
    all('[data-dsp-saved-state]').forEach((el) => { el.hidden = el.dataset.dspSavedState !== DSP.savedState; });
    const q = dq();
    all('[data-exam-pick]').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.examPick === q.exam)));
    all('[data-open-pick]').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.openPick === DSP.pick)));
  };
  const dspRender = () => { dspSync(); DSP_DRAWS.forEach((n) => dash.draw(n)); };
  dash.onData(dspSync);

  all('[data-exam-pick]').forEach((b) => b.addEventListener('click', () => dash.setParams({ exam: b.dataset.examPick })));
  all('.dsp-reset').forEach((b) => b.addEventListener('click', () => { PRESET.dispatch = 'all'; dash.setParams({ dmonth: 'all', courier: 'all', exam: 'all', dfrom: null, dto: null, dbasis: 'order' }); }));
  all('[data-open-pick]').forEach((b) => b.addEventListener('click', () => { DSP.pick = b.dataset.openPick; dspSync(); ['dsp-open-rows', 'dsp-open-count'].forEach((n) => dash.draw(n)); }));
  all('.dsp-open-search').forEach((inp) => inp.addEventListener('input', () => { DSP.query = inp.value.trim().toLowerCase(); ['dsp-open-rows', 'dsp-open-count'].forEach((n) => dash.draw(n)); }));
  all('.dsp-region-search').forEach((inp) => inp.addEventListener('input', () => { DSP.rquery = inp.value.trim().toLowerCase(); dash.draw('dsp-region-rows'); }));
  const sortable = (sel, state, draw, textKeys) => all(sel + ' th[data-sort]').forEach((th) => th.addEventListener('click', () => {
    const k = th.dataset.sort;
    state.dir = state.key === k ? -state.dir : (textKeys.includes(k) ? 1 : -1);
    state.key = k;
    all(sel + ' th[data-sort]').forEach((h) => h.setAttribute('aria-sort', h.dataset.sort === k ? (state.dir > 0 ? 'ascending' : 'descending') : 'none'));
    dash.draw(draw);
  }));
  sortable('.dsp-region-table', DSP.rsort, 'dsp-region-rows', ['region']);
  sortable('.dsp-open-table', DSP.sort, 'dsp-open-rows', ['sno', 'school', 'region', 'courier', 'stage', 'order_date', 'dispatch_date']);
  const sortRows = (rows, s) => rows.sort((a, b) => {
    const x = a[s.key], y = b[s.key];
    if (typeof x === 'number' || typeof y === 'number') return s.dir * ((x ?? -1) - (y ?? -1));
    return s.dir * String(x ?? '').localeCompare(String(y ?? ''), undefined, { numeric: true });
  });

  dash.draw('dsp-month-select', (sel) => {
    const c = D('dsp_cube');
    const months = c.status === 'ok' ? [...new Set(c.data.map((r) => r.month))].filter((m) => m !== 'all').sort() : [];
    sel.replaceChildren(new Option('All months', 'all'), ...months.map((m) => new Option(R.monthLabel(m), m)));
    sel.value = months.includes(dq().month) ? dq().month : 'all';
    sel.onchange = () => dash.setParams({ dmonth: sel.value });
  });
  dash.draw('dsp-courier-select', (sel) => {
    const c = D('dsp_cube');
    const list = c.status === 'ok' ? [...new Set(c.data.map((r) => r.courier))].filter((m) => m !== 'all').sort() : [];
    sel.replaceChildren(new Option('All couriers', 'all'), ...list.map((m) => new Option(m, m)));
    sel.value = list.includes(dq().courier) ? dq().courier : 'all';
    sel.onchange = () => dash.setParams({ courier: sel.value });
  });
  dash.draw('dsp-scope', (el) => {
    const q = dq();
    const dates = rangeText('dispatch');
    el.textContent = 'Showing ' + (q.month === 'all' ? 'all months' : R.monthLabel(q.month)) + ' · ' + (q.courier === 'all' ? 'all couriers' : q.courier) + ' · ' + examLabel(q.exam) + (dates ? ' · ' + dates : '');
  });
  dash.draw('dsp-asof', (el) => {
    const r = cubeRow({ month: 'all', courier: 'all', exam: 'all' });
    cite(el, 'all|all|all');
    el.textContent = r && r.as_of ? fmtDate(r.as_of) : '—';
  });
  dash.draw('dsp-kpi', (el) => {
    const q = dq();
    cite(el, dk(q));
    const r = cubeRow(q);
    el.textContent = r ? fmtv(r[el.dataset.field], el.dataset.fmt) : '—';
  });

  dash.draw('dsp-pipeline', (el) => {
    el.replaceChildren();
    const r = cubeRow(dq());
    if (!r || !r.orders) return;
    const reach = el.clientWidth < 420 ? 0.5 : el.clientWidth < 560 ? 0.62 : 0.74;
    const head = (t) => { const h = document.createElement('div'); h.className = 'ito-fsep'; h.textContent = t; el.append(h); };
    const step = (label, v, color, bad) => {
      const row = document.createElement('div'); row.className = 'ito-frow' + (bad && v ? ' bad' : '');
      const lab = document.createElement('div'); lab.className = 'lab'; lab.textContent = label;
      const bar = document.createElement('div'); bar.className = 'bar';
      const fill = document.createElement('div'); fill.className = 'fill';
      const w = Math.max(0.5, (v / r.orders) * 100);
      fill.style.width = w * reach + '%';
      fill.style.setProperty('--ito-c', bad ? 'var(--color-bad)' : color);
      const val = document.createElement('div'); val.className = 'val';
      val.textContent = fmtInt(v) + ' · ' + (reach < 0.6 ? d3.format('.0%') : fmtPct)(v / r.orders);
      val.style.left = 'calc(' + w * reach + '% + 4px)';
      bar.append(fill, val); row.append(lab, bar); el.append(row);
    };
    head('Orders');
    step('School orders', r.orders, 'var(--cds-chart-muted)');
    step('Shipped', r.shipped, C[0]);
    step('Delivered', r.delivered, C[2]);
    head('Not delivered yet');
    step('In transit, up to 7 days', r.transit_7, C[1]);
    step('In transit, over 7 days', r.transit_late, null, true);
    step('On hold', r.on_hold, 'var(--cds-chart-muted)');
    step('Packed, no courier yet', r.not_shipped, 'var(--cds-chart-muted)');
  });

  dash.draw('dsp-trend', (el) => {
    el.replaceChildren();
    const c = D('dsp_cube');
    if (c.status !== 'ok') return;
    const q = dq();
    const rows = c.data.filter((r) => r.exam === q.exam && r.courier === q.courier && r.month !== 'all' && r.month !== 'unknown' && r.orders > 0)
      .sort((a, b) => (a.month < b.month ? -1 : 1)).map((r) => ({ ...r, total: r.orders, part: r.delivered }));
    if (!rows.length) return;
    monthBars(el, rows, {
      aria: 'School orders and deliveries by month', picked: q.month, part: true,
      lines: (r) => [[null, 'School orders', fmtInt(r.orders)], [null, 'Shipped', fmtInt(r.shipped)], [C[0], 'Delivered', fmtInt(r.delivered)], [null, 'Delivery rate', fmtv(r.delivery_rate, 'pct')], [null, 'Still in transit', fmtInt(r.in_transit)]],
      onClick: (r) => dash.setParams({ dmonth: dq().month === r.month ? 'all' : r.month }),
    });
  });

  dash.draw('dsp-courier-bars', (el) => {
    el.replaceChildren();
    const c = D('dsp_cube');
    if (c.status !== 'ok') return;
    const q = dq();
    const rows = c.data.filter((r) => r.exam === q.exam && r.month === q.month && r.courier !== 'all' && r.shipped > 0).sort((a, b) => b.shipped - a.shipped);
    if (!rows.length) return;
    const worst = d3.min(rows.filter((r) => r.shipped >= 20), (r) => r.delivery_rate);
    hbars(el, rows, {
      aria: 'Delivery rate by courier', label: (r) => r.courier, value: (r) => r.delivery_rate,
      text: (r) => fmtv(r.delivery_rate, 'pct') + ' of ' + fmtInt(r.shipped) + ' · ' + fmtInt(r.in_transit) + ' in transit',
      color: (r) => (r.delivery_rate === worst && rows.length > 1 ? 'var(--cds-chart-status-critical)' : C[0]),
      picked: (r) => r.courier === q.courier,
      onClick: (r) => dash.setParams({ courier: dq().courier === r.courier ? 'all' : r.courier }),
    });
  });

  const COURIER_COLS = [
    { key: 'courier', show: (r) => (r.courier === 'all' ? 'All couriers' : r.courier) }, { key: 'shipped', type: 'int' }, { key: 'delivered', type: 'int' },
    { key: 'delivery_rate', type: 'pct' }, { key: 'in_transit', type: 'int' }, { key: 'transit_late', type: 'int' }, { key: 'avg_days', type: 'days' },
    { key: 'median_days', type: 'days' }, { key: 'within_7', type: 'pct' }, { key: 'issues', type: 'int' },
  ];
  dash.draw('dsp-courier-rows', (tbody) => {
    const c = D('dsp_cube');
    if (c.status !== 'ok') { tbody.replaceChildren(); return; }
    const q = dq();
    const rows = c.data.filter((r) => r.exam === q.exam && r.month === q.month && r.courier !== 'all' && r.shipped > 0).sort((a, b) => b.shipped - a.shipped)
      .map((r) => ({ ...r, __cls: r.courier === q.courier ? 'is-picked' : '' }));
    const total = c.data.find((r) => r.k === q.exam + '|all|' + q.month);
    if (total && rows.length > 1) rows.push({ ...total, __cls: 'is-total' });
    fillBody(tbody, rows, COURIER_COLS, (r) => r.k);
  });

  dash.draw('dsp-daily', (el) => {
    el.replaceChildren();
    const d = D('dsp_daily');
    if (d.status !== 'ok') return;
    const q = dq();
    const rows = d.data.filter((r) => r.exam === q.exam && r.courier === q.courier && (q.month === 'all' || r.month === q.month));
    if (!rows.length || !rows.some((r) => r.shipped)) return;
    const W = Math.max(280, el.clientWidth), H = 300;
    const svg = d3.select(el).append('svg').attr('width', '100%').attr('viewBox', `0 0 ${W} ${H}`).attr('role', 'img').attr('aria-label', 'Shipments per day');
    const parse = (s) => new Date(s + 'T00:00:00Z');
    const y = d3.scaleLinear([0, (d3.max(rows, (r) => Math.max(r.shipped, r.delivered)) || 1) * 1.1], [H - 26, 10]).nice();
    const left = textWidth(svg, fmtInt(y.domain()[1])) + 12;
    const x = d3.scaleBand(rows.map((r) => r.date), [left, W - 10]).padding(0.15);
    svg.append('g').selectAll('line').data(y.ticks(5)).join('line').attr('x1', left).attr('x2', W - 10).attr('y1', (v) => y(v)).attr('y2', (v) => y(v)).attr('stroke', 'var(--cds-chart-grid)');
    svg.append('g').selectAll('text').data(y.ticks(5)).join('text').attr('x', left - 6).attr('y', (v) => y(v)).attr('dy', '0.32em').attr('text-anchor', 'end').attr('font-size', 12).attr('fill', 'var(--color-fg-muted)').text((v) => fmtInt(v));
    const every = Math.max(1, Math.ceil(rows.length / Math.max(2, Math.floor((W - left) / 70))));
    svg.append('g').selectAll('text').data(rows.filter((r, i) => i % every === 0)).join('text').attr('x', (r) => x(r.date) + x.bandwidth() / 2).attr('y', H - 6).attr('text-anchor', 'middle').attr('font-size', 12).attr('fill', 'var(--color-fg-muted)').text((r) => d3.utcFormat('%b %-d')(parse(r.date)));
    svg.append('g').selectAll('rect').data(rows).join('rect').attr('x', (r) => x(r.date)).attr('width', x.bandwidth()).attr('y', (r) => y(r.shipped)).attr('height', (r) => y(0) - y(r.shipped)).attr('rx', Math.min(3, x.bandwidth() / 2)).attr('fill', C[0]);
    svg.append('path').datum(rows).attr('d', d3.line((r) => x(r.date) + x.bandwidth() / 2, (r) => y(r.delivered_7d))).attr('fill', 'none').attr('stroke', C[2]).attr('stroke-width', 2).attr('stroke-linecap', 'round').attr('stroke-linejoin', 'round');
    const tt = tip(el);
    const hover = svg.append('rect').attr('fill', 'none').attr('stroke', 'var(--color-fg)').attr('stroke-width', 1).attr('visibility', 'hidden');
    svg.append('g').selectAll('rect').data(rows).join('rect').attr('x', (r) => x(r.date) - (x.step() - x.bandwidth()) / 2).attr('width', x.step()).attr('y', 0).attr('height', H - 24).attr('fill', 'transparent')
      .on('pointermove pointerdown', (event, r) => {
        hover.attr('x', x(r.date)).attr('width', x.bandwidth()).attr('y', y(r.shipped)).attr('height', y(0) - y(r.shipped)).attr('visibility', r.shipped ? 'visible' : 'hidden');
        tt.show(event, d3.utcFormat('%a %-d %b %Y')(parse(r.date)), [[C[0], 'Shipped', fmtInt(r.shipped)], [null, 'Delivered that day', fmtInt(r.delivered)], [C[2], '7-day average delivered', f1(r.delivered_7d)]]);
      })
      .on('pointerleave', () => { tt.hide(); hover.attr('visibility', 'hidden'); });
  });

  dash.draw('dsp-days', (el) => {
    el.replaceChildren();
    const d = D('dsp_days');
    if (d.status !== 'ok') return;
    const q = dq();
    const rows = d.data.filter((r) => r.k.startsWith(dk(q) + '|')).sort((a, b) => a.order - b.order);
    if (!rows.length || !rows.some((r) => r.delivered)) return;
    const W = Math.max(260, el.clientWidth), H = 200;
    const svg = d3.select(el).append('svg').attr('width', '100%').attr('viewBox', `0 0 ${W} ${H}`).attr('role', 'img').attr('aria-label', 'Delivered orders by days in transit');
    const y = d3.scaleLinear([0, d3.max(rows, (r) => r.delivered) || 1], [H - 24, 40]);
    const x = d3.scaleBand(rows.map((r) => r.bucket), [4, W - 4]).padding(0.25);
    const bw = Math.min(64, x.bandwidth());
    rows.forEach((r) => {
      const cx = x(r.bucket) + x.bandwidth() / 2;
      svg.append('rect').attr('x', cx - bw / 2).attr('width', bw).attr('y', y(r.delivered)).attr('height', y(0) - y(r.delivered)).attr('rx', 4).attr('fill', r.order >= 3 ? 'var(--cds-chart-status-critical)' : r.order === 2 ? C[1] : C[2]);
      svg.append('text').attr('x', cx).attr('y', y(r.delivered) - 20).attr('text-anchor', 'middle').attr('font-size', 12).attr('fill', 'var(--color-fg)').text(fmtInt(r.delivered));
      svg.append('text').attr('x', cx).attr('y', y(r.delivered) - 6).attr('text-anchor', 'middle').attr('font-size', 11).attr('fill', 'var(--color-fg-muted)').text(d3.format('.0%')(r.share || 0));
      svg.append('text').attr('x', cx).attr('y', H - 6).attr('text-anchor', 'middle').attr('font-size', 12).attr('fill', 'var(--color-fg-muted)').text(r.bucket);
    });
  });

  dash.draw('dsp-region-rows', (tbody) => {
    const d = D('dsp_regions');
    if (d.status !== 'ok') { tbody.replaceChildren(); return; }
    const q = dq();
    const rows = d.data.filter((r) => r.k.startsWith(dk(q) + '|') && (!DSP.rquery || regionLabel(r.region).toLowerCase().includes(DSP.rquery)));
    sortRows(rows, DSP.rsort);
    fillBody(tbody, rows, [
      { key: 'region', show: (r) => regionLabel(r.region) }, { key: 'orders', type: 'int' }, { key: 'shipped', type: 'int' }, { key: 'delivered', type: 'int' },
      { key: 'delivery_rate', type: 'pct' }, { key: 'in_transit', type: 'int' }, { key: 'pending_ship', type: 'int' }, { key: 'units', type: 'int' }, { key: 'avg_days', type: 'days' },
    ], (r) => r.k);
  });

  const UNITS = [['ITO workbooks', 'ito_workbooks'], ['Question papers (PYQPs)', 'pyqps'], ['Guides', 'guides'], ['LC workbooks', 'lc_workbooks'], ['LC textbooks', 'textbooks']];
  dash.draw('dsp-units', (el) => {
    el.replaceChildren();
    const r = cubeRow(dq());
    if (!r) return;
    const rows = UNITS.map(([label, f]) => ({ label, v: r[f] || 0 })).filter((x) => x.v > 0);
    if (!rows.length) return;
    const max = d3.max(rows, (x) => x.v);
    hbars(el, rows, { aria: 'Copies sent by type', label: (x) => x.label, value: (x) => x.v / max, text: (x) => fmtInt(x.v), color: () => C[0] });
  });

  dash.draw('dsp-material-rows', (tbody) => {
    tbody.replaceChildren();
    const d = D('dsp_material');
    if (d.status !== 'ok') return;
    const q = dq();
    d.data.filter((r) => q.exam === 'all' || r.exam === q.exam).forEach((r) => {
      const tr = document.createElement('tr');
      tr.dataset.key = r.k;
      [examLabel(r.exam), r.item, fmtInt(r.count), fmtInt(r.packed), fmtInt(r.dispatch_today), fmtInt(r.pending), fmtv(r.packed_share, 'pct')].forEach((v, i) => {
        const td = document.createElement('td');
        td.textContent = v;
        if (i === 6 && r.packed_share !== null) {
          const b = document.createElement('span'); b.className = 'ito-cellbar';
          b.style.width = Math.round(r.packed_share * 48) + 'px';
          b.style.setProperty('--ito-c', r.packed_share < 0.95 ? C[1] : C[2]);
          td.append(b);
        }
        tr.append(td);
      });
      tbody.append(tr);
    });
  });

  const PICKS = {
    open: (r) => r.stage !== 'Delivered', late: (r) => r.stage === 'In transit' && r.days_open > 7, transit: (r) => r.stage === 'In transit',
    hold: (r) => r.stage === 'On hold', notshipped: (r) => r.stage === 'Not shipped yet', issues: (r) => !!r.issue,
  };
  const openRows = (pick, withSearch) => {
    const d = D('dsp_open');
    if (d.status !== 'ok') return null;
    const q = dq();
    const s = withSearch ? DSP.query : '';
    return d.data.filter((r) => (q.exam === 'all' || r.exam === q.exam) && (q.courier === 'all' || r.courier === q.courier) && (q.month === 'all' || r.month === q.month) && PICKS[pick](r)
      && (!s || [r.school, r.pod, r.remark, r.courier, r.sno, regionLabel(r.region)].some((v) => String(v || '').toLowerCase().includes(s))));
  };
  dash.draw('dsp-open-n', (el) => { const rows = openRows(el.dataset.pick, false); el.textContent = rows ? fmtInt(rows.length) : '—'; });
  dash.draw('dsp-open-count', (el) => { const rows = openRows(DSP.pick, true); el.textContent = rows ? fmtInt(rows.length) + (rows.length === 1 ? ' shipment' : ' shipments') : '—'; });
  dash.draw('dsp-open-rows', (tbody) => {
    tbody.replaceChildren();
    const rows = openRows(DSP.pick, true);
    if (!rows) return;
    sortRows(rows, DSP.sort).forEach((r) => {
      const tr = document.createElement('tr');
      tr.dataset.key = r.k;
      const cells = [r.sno, r.school, regionLabel(r.region), examLabel(r.exam), r.courier || '—', r.pod || '—', fmtDate(r.order_date) || '—', fmtDate(r.dispatch_date) || '—', null, fmtv(r.days_open, 'int'), r.issue ? r.issue + (r.remark ? ': ' + r.remark : '') : r.remark];
      cells.forEach((v, i) => {
        const td = document.createElement('td');
        if (i === 8) {
          const p = document.createElement('span');
          p.className = 'dsp-pill ' + (r.stage === 'In transit' && r.days_open > 7 ? 'is-bad' : STAGE_CLS[r.stage] || '');
          p.textContent = r.stage;
          td.append(p);
        } else td.textContent = v || '';
        if (i === 9) td.className = 'num';
        tr.append(td);
      });
      tbody.append(tr);
    });
  });

  // where the Dispatch tabs' data comes from
  dash.draw('dsp-source', (el) => {
    if (DSP.override) { el.textContent = (DSP.savedId ? 'Saved report: ' : 'Your file: ') + DSP.name + ' · ' + fmtInt(DSP.rowsCount || 0) + ' dispatch rows.'; return; }
    const r = cubeRow({ month: 'all', courier: 'all', exam: 'all' });
    el.textContent = r ? 'The dashboard\'s dispatch sheet: ' + fmtInt(r.orders) + ' school orders, data up to ' + fmtDate(r.as_of) + '.' : '—';
  });
  dash.draw('dsp-source-name', (el) => { el.textContent = DSP.override ? DSP.name : '—'; });
  dash.draw('dsp-error', (el) => { el.textContent = DSP.errorKind === 'error' ? DSP.error : '—'; });

  const useDispatch = (rows, material, name, files) => {
    DSP.rows = rows; DSP.rowsCount = rows.length; DSP.name = name; DSP.files = files; DSP.savedId = null; DSP.saveState = '';
    DSP.override = window.dspCalcs.run(rows, material);
    dash.setParams({ dmonth: 'all', courier: 'all' });
    dspRender();
  };
  const dspHandle = async (list) => {
    if (!list.length) return;
    DSP.busy = true; DSP.error = ''; DSP.errorKind = ''; dspSync();
    try {
      const files = await readAll(list);
      const books = files.map((f) => ({ f, b: dispatchBook(f) }));
      const bad = books.find((x) => !x.b);
      if (bad) throw new Error(bad.f.name + " doesn't look like a dispatch sheet. It needs columns for the courier, the school code, and the POD number or dispatch date. For other files, use the MIS or Support Team pages.");
      const rows = books.flatMap((x, i) => (books.length > 1 ? x.b.rows.map((r) => ({ ...r, sno: (i + 1) + '-' + r.sno })) : x.b.rows));
      const material = (books.find((x) => x.b.material.length) || { b: { material: [] } }).b.material;
      useDispatch(rows, material, files.length === 1 ? files[0].name : files.length + ' files', files.map((f) => f.name));
    } catch (e) {
      [DSP.errorKind, DSP.error] = errorOf(e);
    }
    DSP.busy = false;
    dspRender();
  };
  all('.dsp-file').forEach((input) => input.addEventListener('change', () => { dspHandle([...input.files]); input.value = ''; }));
  const dropzone = (sel, fn) => all(sel).forEach((zone) => {
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('is-over'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('is-over'));
    zone.addEventListener('drop', (e) => { e.preventDefault(); zone.classList.remove('is-over'); if (e.dataTransfer) fn([...e.dataTransfer.files], zone); });
  });
  dropzone('.dsp-drop', (files) => dspHandle(files));
  all('.dsp-back').forEach((b) => b.addEventListener('click', () => {
    DSP.override = null; DSP.rows = null; DSP.savedId = null; DSP.saveState = ''; DSP.error = '';
    dash.setParams({ dmonth: 'all', courier: 'all' });
    dspRender();
  }));

  // dispatch downloads
  const DCOLS = {
    view: [
      { key: 'label', label: 'View', type: 'text' }, { key: 'orders', label: 'School orders', type: 'int' }, { key: 'schools', label: 'Schools', type: 'int' },
      { key: 'shipped', label: 'Shipped', type: 'int' }, { key: 'delivered', label: 'Delivered', type: 'int' }, { key: 'delivery_rate', label: 'Delivery rate', type: 'pct' },
      { key: 'in_transit', label: 'In transit', type: 'int' }, { key: 'transit_late', label: 'In transit over 7 days', type: 'int' }, { key: 'on_hold', label: 'On hold', type: 'int' },
      { key: 'not_shipped', label: 'Packed, no courier yet', type: 'int' }, { key: 'avg_days', label: 'Avg delivery days', type: 'num' }, { key: 'median_days', label: 'Median delivery days', type: 'num' },
      { key: 'within_7', label: 'Delivered within 7 days', type: 'pct' }, { key: 'avg_lag', label: 'Avg days order to dispatch', type: 'num' }, { key: 'issues', label: 'Returns & re-routes', type: 'int' },
      { key: 'ito_workbooks', label: 'ITO workbooks', type: 'int' }, { key: 'pyqps', label: 'Question papers', type: 'int' }, { key: 'guides', label: 'Guides', type: 'int' },
      { key: 'lc_workbooks', label: 'LC workbooks', type: 'int' }, { key: 'textbooks', label: 'LC textbooks', type: 'int' },
    ],
    region: [{ key: 'label', label: 'Region', type: 'text' }, { key: 'orders', label: 'School orders', type: 'int' }, { key: 'shipped', label: 'Shipped', type: 'int' }, { key: 'delivered', label: 'Delivered', type: 'int' }, { key: 'delivery_rate', label: 'Delivery rate', type: 'pct' }, { key: 'in_transit', label: 'In transit', type: 'int' }, { key: 'transit_late', label: 'In transit over 7 days', type: 'int' }, { key: 'pending_ship', label: 'Not shipped yet', type: 'int' }, { key: 'units', label: 'Books & papers', type: 'int' }, { key: 'avg_days', label: 'Avg delivery days', type: 'num' }],
    days: [{ key: 'bucket', label: 'Days from dispatch to delivery', type: 'text' }, { key: 'delivered', label: 'Delivered orders', type: 'int' }, { key: 'share', label: 'Share', type: 'pct' }],
    open: [{ key: 'sno', label: 'S.No.', type: 'text' }, { key: 'school', label: 'School code', type: 'text' }, { key: 'region_label', label: 'Region', type: 'text' }, { key: 'exam_label', label: 'Programme', type: 'text' }, { key: 'courier', label: 'Courier', type: 'text' }, { key: 'pod', label: 'POD / AWB', type: 'text' }, { key: 'order_date', label: 'Order date', type: 'text' }, { key: 'dispatch_date', label: 'Dispatch date', type: 'text' }, { key: 'stage', label: 'Status', type: 'text' }, { key: 'days_open', label: 'Days open', type: 'int' }, { key: 'issue', label: 'Issue', type: 'text' }, { key: 'remark', label: 'Note', type: 'text' }, { key: 'units', label: 'Books & papers', type: 'int' }],
    material: [{ key: 'exam_label', label: 'Programme', type: 'text' }, { key: 'item', label: 'Item', type: 'text' }, { key: 'count', label: 'Ordered', type: 'int' }, { key: 'packed', label: 'Packed', type: 'int' }, { key: 'dispatch_today', label: 'Dispatched today', type: 'int' }, { key: 'pending', label: 'Pending', type: 'int' }, { key: 'packed_share', label: 'Packed share', type: 'pct' }],
    rows: [{ key: 'sno', label: 'S.No.' }, { key: 'school', label: 'School code' }, { key: 'exam', label: 'Programme' }, { key: 'region', label: 'Region' }, { key: 'order_date', label: 'Order date' }, { key: 'courier', label: 'Courier' }, { key: 'pod', label: 'POD / AWB' }, { key: 'ito_workbooks', label: 'ITO workbooks' }, { key: 'pyqps', label: 'Question papers' }, { key: 'guides', label: 'Guides' }, { key: 'lc_workbooks', label: 'LC workbooks' }, { key: 'textbooks', label: 'LC textbooks' }, { key: 'order_status', label: 'Order status' }, { key: 'dispatch_date', label: 'Dispatch date' }, { key: 'delivery_date', label: 'Delivery date' }, { key: 'stage', label: 'Stage' }, { key: 'issue', label: 'Issue' }, { key: 'remark', label: 'Note' }],
  };
  const openOut = (rows) => rows.map((r) => ({ ...r, region_label: regionLabel(r.region), exam_label: examLabel(r.exam) }));
  const dspTag = () => { const q = dq(); return [q.month === 'all' ? 'all_months' : q.month, q.courier !== 'all' ? q.courier : '', q.exam !== 'all' ? q.exam : ''].filter(Boolean).join('_'); };
  const dspXlsx = () => {
    const q = dq();
    const cube = D('dsp_cube'), regions = D('dsp_regions'), days = D('dsp_days'), mat = D('dsp_material');
    if (cube.status !== 'ok') return;
    const now = cubeRow(q);
    if (!now) return;
    const summary = [
      ['Source', DSP.override ? DSP.name : "The dashboard's dispatch sheet", 'text'], ['Data up to', fmtDate(now.as_of), 'text'],
      ['Order month', q.month === 'all' ? 'All months' : R.monthLabel(q.month), 'text'], ['Courier', q.courier === 'all' ? 'All couriers' : q.courier, 'text'], ['Programme', examLabel(q.exam), 'text'],
      ...DCOLS.view.slice(1).map((c) => [c.label, now[c.key], c.type === 'pct' ? 'pct' : c.type === 'num' ? 'num' : 'int']),
    ];
    const book = R.toXlsx([
      { name: 'Summary', columns: [{ key: 'm', label: 'Measure', type: 'text' }, { key: 'v', label: 'Value', type: 'num' }, { key: 'p', label: 'Rate', type: 'pct' }, { key: 't', label: 'Detail', type: 'text' }],
        rows: summary.map(([m, v, t]) => ({ m, v: t === 'int' || t === 'num' ? (typeof v === 'number' ? Math.round(v * 10) / 10 : null) : null, p: t === 'pct' ? v : null, t: t === 'text' ? v : null })) },
      { name: 'By courier', columns: DCOLS.view, rows: cube.data.filter((r) => r.exam === q.exam && r.month === q.month && r.shipped > 0).sort((a, b) => (a.courier === 'all') - (b.courier === 'all') || b.shipped - a.shipped).map((r) => ({ ...r, label: r.courier === 'all' ? 'All couriers' : r.courier })) },
      { name: 'By month', columns: DCOLS.view, rows: cube.data.filter((r) => r.exam === q.exam && r.courier === q.courier && r.orders > 0).sort((a, b) => (a.month === 'all') - (b.month === 'all') || (a.month < b.month ? -1 : 1)).map((r) => ({ ...r, label: R.monthLabel(r.month) })) },
      { name: 'By region', columns: DCOLS.region, rows: regions.status === 'ok' ? regions.data.filter((r) => r.k.startsWith(dk(q) + '|')).sort((a, b) => b.orders - a.orders).map((r) => ({ ...r, label: regionLabel(r.region) })) : [] },
      { name: 'Delivery time', columns: DCOLS.days, rows: days.status === 'ok' ? days.data.filter((r) => r.k.startsWith(dk(q) + '|')).sort((a, b) => a.order - b.order) : [] },
      { name: 'Follow-up list', columns: DCOLS.open, rows: openOut(sortRows(openRows('open', false) || [], { key: 'days_open', dir: -1 })) },
      { name: 'Returns & re-routes', columns: DCOLS.open, rows: openOut(openRows('issues', false) || []) },
      { name: 'Study material', columns: DCOLS.material, rows: mat.status === 'ok' ? mat.data.filter((r) => q.exam === 'all' || r.exam === q.exam).map((r) => ({ ...r, exam_label: examLabel(r.exam) })) : [] },
    ]);
    save('Dispatch_report_' + dspTag() + '.xlsx', new Blob([book]));
  };
  all('.dsp-dl-xlsx').forEach((b) => b.addEventListener('click', dspXlsx));
  all('.dsp-dl-open').forEach((b) => b.addEventListener('click', () => {
    const rows = openRows(b.dataset.pick || DSP.pick, !b.dataset.pick);
    if (rows) save('Dispatch_follow_up_' + dspTag() + '.csv', R.toCSV(DCOLS.open, openOut(sortRows(rows, DSP.sort))));
  }));
  all('.dsp-dl-rows').forEach((b) => b.addEventListener('click', () => { if (DSP.rows) save(fileStem(DSP.name) + '_dispatch_rows.csv', R.toCSV(DCOLS.rows, DSP.rows)); }));

  // ---------- MIS and Support Team for Schools ----------

  const newG = () => ({ busy: false, error: '', errorKind: '', files: [], sheetIdx: 0, table: null, a: null, name: '', fileNames: [], savedId: null, saveState: '', by: '', sum: '', group: null, hint: '', book: null, closed: false, saved: [], confirmDel: null, view: null, full: null, from: '', to: '', preset: 'all', prev: null });
  const G = { mis: newG(), support: newG() };
  const scopeOf = (el) => { const s = el.closest('[data-dscope]'); return s && G[s.dataset.dscope] ? s.dataset.dscope : null; };
  const inScope = (d, sel) => document.querySelectorAll('[data-dscope="' + d + '"] ' + sel);
  const GX_DRAWS = ['gx-name', 'gx-meta', 'gx-kpi', 'gx-col', 'gx-month', 'gx-status', 'gx-cats', 'gx-nums', 'gx-cols', 'gx-preview', 'gx-by', 'gx-sum', 'gx-group', 'gx-saved', 'gx-sheet', 'gx-error', 'gx-date-input', 'gx-range-scope', 'ins-gx'];
  const gxSync = (d) => {
    const g = G[d], a = g.a;
    const has = { report: !!a, empty: !a, saved: !!g.savedId, local: !!g.table, status: !!(a && a.statusCol), duration: !!(a && a.closedCount), date: !!(a && a.dateCol), cats: !!(a && a.categories.length), nums: !!(a && a.numbers.length),
      save: !!(a && !g.savedId && CAP.db && CAP.canEdit), noedit: !!(a && !g.savedId && CAP.db && !CAP.canEdit), group: !!(g.by && g.group), rows: !!g.table };
    inScope(d, '[data-gx-when]').forEach((el) => { el.hidden = !has[el.dataset.gxWhen]; });
    inScope(d, '[data-gx-needs]').forEach((el) => { el.hidden = !el.dataset.gxNeeds.split(' ').every((k) => has[k]); });
    inScope(d, '[data-gx-msg]').forEach((el) => { const k = el.dataset.gxMsg; el.hidden = !(k === 'reading' ? g.busy : !!g.error && g.errorKind === k); });
    inScope(d, '[data-gx-hint]').forEach((el) => { el.hidden = g.hint !== el.dataset.gxHint; });
    inScope(d, '[data-gx-save]').forEach((el) => { el.hidden = g.saveState !== el.dataset.gxSave; });
    inScope(d, '.gx-save').forEach((b) => { b.disabled = g.saveState === 'saving' || g.saveState === 'saved'; });
    inScope(d, '.gx-presets [data-preset]').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.preset === (g.from || g.to ? g.preset : 'all'))));
    inScope(d, '[data-gx-sheets]').forEach((el) => { el.hidden = !(g.files.length === 1 && g.files[0].sheets.length > 1 && !g.savedId); });
    inScope(d, '[data-gx-saved-state]').forEach((el) => { el.hidden = el.dataset.gxSavedState !== (CAP.db ? (g.saved.length ? 'list' : DSP.savedState === 'loading' ? 'loading' : 'empty') : DSP.savedState === 'loading' ? 'loading' : 'off'); });
  };
  const gxRender = (d) => { gxSync(d); GX_DRAWS.forEach((n) => dash.draw(n)); };

  const gxBuild = (d) => {
    const g = G[d];
    g.a = null; g.table = null; g.savedId = null; g.saveState = ''; g.group = null; g.hint = ''; g.book = null;
    if (!g.files.length) return;
    const biggest = (f) => f.sheets.reduce((a, s) => (s.rows.length > a.rows.length ? s : a), f.sheets[0]);
    const tables = g.files.map((f) => (g.files.length === 1 && f.sheets.length > 1 ? f.sheets[Math.min(g.sheetIdx, f.sheets.length - 1)] : biggest(f)));
    const names = g.files.map((f) => f.name);
    const many = tables.length > 1;
    const cols = [...new Set(tables.flatMap((t) => t.columns))];
    g.table = { columns: many ? ['Source file', ...cols] : cols, rows: tables.flatMap((t, i) => t.rows.map((r) => (many ? { 'Source file': names[i], ...r } : r))) };
    g.name = names.length === 1 ? names[0] : names.length + ' files';
    g.fileNames = names;
    g.a = analyze(g.table);
    g.full = g.a; g.view = g.table; g.from = ''; g.to = ''; g.preset = 'all'; g.prev = null;
    const book = g.files.length === 1 ? dispatchBook(g.files[0]) : null;
    if (book) { g.hint = 'dispatch'; g.book = book; }
    else if (R.isIto(g.table.columns)) g.hint = 'orders';
    const firstCat = g.a.categories[0];
    g.by = g.a.statusCol || (firstCat ? firstCat.column : '');
    g.sum = '';
    g.group = g.by ? R.groupBy(g.table, g.by, g.sum) : null;
  };
  // Dates on an uploaded sheet: the report is worked out again for the rows whose date is in the range, and the period before is counted for comparison.
  const gxRowDate = (g) => { const col = g.full.dateCol; return (r) => { const v = r[col]; return v === '' || v === null || v === undefined ? '' : R.toISODate(v); }; };
  const gxApplyRange = (d) => {
    const g = G[d];
    if (!g.table || !g.full || !g.full.dateCol) return;
    if (!g.from && !g.to) { g.view = g.table; g.a = g.full; g.prev = null; }
    else {
      const dateOf = gxRowDate(g);
      g.view = { columns: g.table.columns, rows: g.table.rows.filter((r) => inRange(dateOf(r), g.from, g.to)) };
      const a = analyze(g.view);
      g.a = { ...a, dateCol: g.full.dateCol };
      const from = g.from || g.full.from, to = g.to || g.full.to;
      const pf = isoAdd(from, -spanDays(from, to)), pt = isoAdd(from, -1);
      const prevRows = g.table.rows.filter((r) => inRange(dateOf(r), pf, pt));
      const sc = g.full.statusCol;
      g.prev = { from: pf, to: pt, rows: prevRows.length, done: sc ? prevRows.filter((r) => isDoneValue(r[sc])).length : null };
    }
    g.group = g.by && g.view ? R.groupBy(g.view, g.by, g.sum) : g.group;
  };
  const gxSetRange = (d, from, to, preset) => {
    const g = G[d];
    if (from && to && from > to) [from, to] = [to, from];
    g.from = from || ''; g.to = to || ''; g.preset = preset || (from || to ? '' : 'all');
    gxApplyRange(d);
    gxRender(d);
  };
  dash.draw('gx-date-input', (el) => {
    const d = scopeOf(el); const g = d && G[d];
    if (!g || !g.full) { el.value = ''; return; }
    if (g.full.from) { el.min = g.full.from; el.max = g.full.to; }
    if (el !== document.activeElement) el.value = g[el.dataset.end] || '';
    el.classList.toggle('is-set', !!el.value);
    el.onchange = () => {
      if (el.value && !(el.validity.valid && /^d{4}-d{2}-d{2}$/.test(el.value))) return;
      gxSetRange(d, el.dataset.end === 'from' ? el.value : g.from, el.dataset.end === 'to' ? el.value : g.to);
    };
    el.onblur = () => { el.value = g[el.dataset.end] || ''; };
  });
  dash.draw('gx-range-scope', (el) => {
    const d = scopeOf(el); const g = d && G[d];
    if (!g || !g.a || !g.full) { el.textContent = '—'; return; }
    el.textContent = g.from || g.to ? 'Showing ' + fmtInt(g.a.rows) + ' of ' + fmtInt(g.full.rows) + ' entries · ' + (g.from ? fmtShort(g.from) : 'start') + ' – ' + (g.to ? fmtDate(g.to) : 'latest') : 'Showing all ' + fmtInt(g.full.rows) + ' entries';
  });
  all('.gx-presets [data-preset]').forEach((b) => b.addEventListener('click', () => {
    const d = scopeOf(b); const g = d && G[d];
    if (!g || !g.full || !g.full.to) return;
    const r = presetRange(b.dataset.preset, { min: g.full.from, max: g.full.to });
    gxSetRange(d, r.from, r.to, b.dataset.preset);
  }));
  const gxHandle = async (d, list) => {
    if (!list.length) return;
    const g = G[d];
    g.busy = true; g.error = ''; g.errorKind = ''; gxSync(d);
    try {
      g.files = await readAll(list);
      g.sheetIdx = 0;
      if (g.files.length === 1 && g.files[0].sheets.length > 1) {
        const s = g.files[0].sheets;
        g.sheetIdx = s.reduce((bi, x, i) => (x.rows.length > s[bi].rows.length ? i : bi), 0);
      }
      gxBuild(d);
      goTab(d + '-report');
    } catch (e) {
      [g.errorKind, g.error] = errorOf(e);
    }
    g.busy = false;
    gxRender(d);
  };
  all('.gx-file').forEach((input) => input.addEventListener('change', () => { const d = scopeOf(input); if (d) gxHandle(d, [...input.files]); input.value = ''; }));
  dropzone('.gx-drop', (files, zone) => { const d = scopeOf(zone); if (d) gxHandle(d, files); });
  all('.gx-clear').forEach((b) => b.addEventListener('click', () => { const d = scopeOf(b); if (!d) return; Object.assign(G[d], { files: [], table: null, a: null, savedId: null, error: '', hint: '', closed: true }); gxRender(d); }));
  all('.gx-to-dispatch').forEach((b) => b.addEventListener('click', () => {
    const d = scopeOf(b);
    if (!d || !G[d].book) return;
    useDispatch(G[d].book.rows, G[d].book.material, G[d].name, G[d].fileNames);
    goTab('dsp-overview');
  }));

  const gxMeta = (g) => {
    const a = g.a;
    if (!a) return '';
    const parts = [fmtInt(a.rows) + ' rows', fmtInt(a.columns) + ' columns'];
    if (a.from) parts.push(fmtDate(a.from) + ' – ' + fmtDate(a.to) + ' (by ' + a.dateCol + ')');
    if (g.fileNames.length > 1) parts.push(g.fileNames.length + ' files combined');
    if (g.savedId) parts.push('saved report');
    return parts.join(' · ');
  };
  dash.draw('gx-name', (el) => { const d = scopeOf(el); el.textContent = d && G[d].a ? G[d].name : '—'; });
  dash.draw('gx-meta', (el) => { const d = scopeOf(el); el.textContent = d && G[d].a ? gxMeta(G[d]) : '—'; });
  dash.draw('gx-error', (el) => { const d = scopeOf(el); el.textContent = d && G[d].errorKind === 'error' ? G[d].error : '—'; });
  dash.draw('gx-col', (el) => { const d = scopeOf(el); const a = d && G[d].a; const v = a ? a[el.dataset.field] ?? G[d][el.dataset.field] : null; el.textContent = v || '—'; });
  dash.draw('gx-kpi', (el) => {
    const d = scopeOf(el); const a = d && G[d].a;
    const f = el.dataset.field;
    if (!a) { el.textContent = '—'; return; }
    if (f === 'period') { el.textContent = a.from ? fmtDate(a.from) + ' – ' + fmtDate(a.to) : '—'; return; }
    el.textContent = fmtv(a[f], el.dataset.fmt);
  });
  dash.draw('gx-sheet', (sel) => {
    const d = scopeOf(sel); const f = d && G[d].files[0];
    sel.replaceChildren(...(f ? f.sheets : []).map((s, i) => new Option(s.name + ' (' + fmtInt(s.rows.length) + ' rows)', String(i))));
    sel.value = String(d ? G[d].sheetIdx : 0);
    sel.onchange = () => { G[d].sheetIdx = Number(sel.value); gxBuild(d); gxRender(d); };
  });
  dash.draw('gx-month', (el) => {
    el.replaceChildren();
    const d = scopeOf(el); const a = d && G[d].a;
    if (!a || !a.byMonth.length) return;
    const withDone = !!a.statusCol;
    monthBars(el, a.byMonth.map((r) => ({ ...r, total: r.rows, part: withDone ? r.done : null })), {
      aria: 'Entries by month', part: withDone,
      lines: (r) => [[null, 'Entries', fmtInt(r.rows)], ...(withDone ? [[C[0], 'Completed', fmtInt(r.done)], [null, 'Completed share', fmtv(r.rows ? r.done / r.rows : null, 'pct')]] : [])],
    });
  });
  dash.draw('gx-status', (el) => {
    el.replaceChildren();
    const d = scopeOf(el); const a = d && G[d].a;
    if (!a || !a.byStatus.length) return;
    hbars(el, a.byStatus, { aria: 'Entries by status', label: (r) => r.value, value: (r) => r.share, text: (r) => fmtPct(r.share) + ' · ' + fmtInt(r.rows), color: (r) => (r.is_done ? C[2] : r.value === '(blank)' ? 'var(--cds-chart-muted)' : C[1]) });
  });
  dash.draw('gx-cats', (el) => {
    el.replaceChildren();
    const d = scopeOf(el); const a = d && G[d].a;
    if (!a) return;
    // Lay out every block first, so each chart is drawn at its final column width.
    const charts = a.categories.map((c) => {
      const box = document.createElement('div'); box.className = 'gx-cat';
      const h = document.createElement('h3'); h.textContent = c.column;
      const s = document.createElement('div'); s.className = 'ito-sub'; s.textContent = fmtInt(c.distinct) + ' different values';
      const chart = document.createElement('div'); chart.className = 'ito-carts';
      box.append(h, s, chart);
      el.append(box);
      return [c, chart];
    });
    charts.forEach(([c, chart]) => {
      hbars(chart, c.values, { aria: 'Entries by ' + c.column, label: (r) => (r.value.length > 28 ? r.value.slice(0, 27) + '…' : r.value), value: (r) => r.share, text: (r) => fmtInt(r.rows) + ' · ' + d3.format('.0%')(r.share), color: (r) => (/^Other \(/.test(r.value) || r.value === '(blank)' ? 'var(--cds-chart-muted)' : C[0]) });
    });
  });
  dash.draw('gx-nums', (tbody) => {
    const d = scopeOf(tbody); const a = d && G[d].a;
    fillBody(tbody, a ? a.numbers : [], [{ key: 'column', type: 'text' }, { key: 'filled', type: 'int' }, { key: 'total', type: 'num' }, { key: 'avg', show: (r) => (r.avg === null ? '—' : d3.format(',.2~f')(r.avg)) }, { key: 'min', type: 'num' }, { key: 'max', type: 'num' }]);
  });
  dash.draw('gx-cols', (tbody) => {
    const d = scopeOf(tbody); const a = d && G[d].a;
    fillBody(tbody, a ? a.profile : [], [{ key: 'column', show: (r) => r.column + (r.personal ? ' (personal)' : '') }, { key: 'type', type: 'text' }, { key: 'filled', type: 'int' }, { key: 'empty', type: 'int' }, { key: 'distinct', type: 'int' }, { key: 'total', type: 'num' }, { key: 'min', type: 'num' }, { key: 'max', type: 'num' }]);
  });
  dash.draw('gx-preview', (table) => {
    const d = scopeOf(table); const g = d && G[d];
    const v = g && (g.view || g.table);
    if (!v) { table.replaceChildren(); return; }
    fillTable(table, v.columns.map((c) => ({ key: c, label: c, type: 'text' })), v.rows.slice(0, 50));
  });
  const groupCols = (g) => [{ key: 'group', label: g.by, type: 'text' }, { key: 'rows', label: 'Rows', type: 'int' }, { key: 'share', label: 'Share of rows', type: 'pct' }, ...(g.sum ? [{ key: 'total', label: 'Total ' + g.sum, type: 'num' }] : [])];
  const gxPick = (sel, first, filter, prop) => {
    const d = scopeOf(sel); const g = d && G[d];
    const cols = g && g.a ? g.a.profile.filter(filter).map((c) => c.column) : [];
    sel.replaceChildren(new Option(first, ''), ...cols.map((c) => new Option(c, c)));
    sel.value = g && cols.includes(g[prop]) ? g[prop] : '';
    sel.disabled = !(g && g.table);
    sel.onchange = () => { g[prop] = sel.value; g.group = g.by && g.table ? R.groupBy(g.view || g.table, g.by, g.sum) : null; gxRender(d); };
  };
  dash.draw('gx-by', (sel) => gxPick(sel, 'Pick a column', (c) => c.type !== 'Number' || c.distinct <= 50, 'by'));
  dash.draw('gx-sum', (sel) => gxPick(sel, 'Count rows only', (c) => c.type === 'Number', 'sum'));
  dash.draw('gx-group', (table) => {
    const d = scopeOf(table); const g = d && G[d];
    if (!g || !g.group || !g.by) { table.replaceChildren(); return; }
    fillTable(table, groupCols(g), g.group.slice(0, 500));
  });

  const gxSheets = (d) => {
    const g = G[d], a = g.a;
    const summary = [
      ['Department', d === 'mis' ? 'MIS' : 'Support Team for Schools'], ['Source', g.fileNames.join(', ')], ...(g.from || g.to ? [['Dates picked', (g.from || 'start') + ' to ' + (g.to || 'latest')]] : []), ['Rows', a.rows], ['Columns', a.columns],
      ...(a.from ? [['Period', fmtDate(a.from) + ' – ' + fmtDate(a.to)], ['Dated by', a.dateCol]] : []),
      ...(a.statusCol ? [['Status column', a.statusCol], ['Completed / closed', a.done], ['Completed share', a.doneRate, 'pct'], ['Not completed', a.open]] : []),
      ...(a.closedCount ? [['Average days to close', Math.round(a.avgClose * 10) / 10], ['Median days to close', a.medianClose], ['Closed with both dates', a.closedCount]] : []),
    ];
    const sheets = [{ name: 'Summary', columns: [{ key: 'm', label: 'Measure', type: 'text' }, { key: 'v', label: 'Value', type: 'num' }, { key: 'p', label: 'Rate', type: 'pct' }, { key: 't', label: 'Detail', type: 'text' }],
      rows: summary.map(([m, v, t]) => ({ m, v: typeof v === 'number' && t !== 'pct' ? v : null, p: t === 'pct' ? v : null, t: typeof v === 'number' ? null : v })) }];
    if (a.byMonth.length) sheets.push({ name: 'By month', columns: [{ key: 'label', label: 'Month', type: 'text' }, { key: 'rows', label: 'Entries', type: 'int' }, ...(a.statusCol ? [{ key: 'done', label: 'Completed', type: 'int' }] : [])], rows: a.byMonth.map((r) => ({ ...r, label: R.monthLabel(r.month) })) });
    if (a.byStatus.length) sheets.push({ name: 'By status', columns: [{ key: 'value', label: a.statusCol, type: 'text' }, { key: 'rows', label: 'Entries', type: 'int' }, { key: 'share', label: 'Share', type: 'pct' }], rows: a.byStatus });
    if (a.categories.length) sheets.push({ name: 'Top values', columns: [{ key: 'column', label: 'Column', type: 'text' }, { key: 'value', label: 'Value', type: 'text' }, { key: 'rows', label: 'Entries', type: 'int' }, { key: 'share', label: 'Share', type: 'pct' }, ...(a.statusCol ? [{ key: 'done_share', label: 'Completed share', type: 'pct' }] : [])], rows: a.categories.flatMap((c) => c.values.map((v) => ({ column: c.column, ...v }))) });
    if (a.numbers.length) sheets.push({ name: 'Numbers', columns: [{ key: 'column', label: 'Column', type: 'text' }, { key: 'filled', label: 'Filled', type: 'int' }, { key: 'total', label: 'Total', type: 'num' }, { key: 'avg', label: 'Average', type: 'num' }, { key: 'min', label: 'Lowest', type: 'num' }, { key: 'max', label: 'Highest', type: 'num' }], rows: a.numbers });
    sheets.push({ name: 'Columns', columns: R.COLS.profile, rows: a.profile });
    if (g.group && g.by) sheets.push({ name: 'By ' + g.by, columns: groupCols(g), rows: g.group });
    const v = g.view || g.table;
    if (v) sheets.push({ name: 'Data', columns: v.columns.map((c) => ({ key: c, label: c, type: 'auto' })), rows: v.rows });
    return sheets;
  };
  all('.gx-dl-xlsx').forEach((b) => b.addEventListener('click', () => { const d = scopeOf(b); if (d && G[d].a) save(fileStem(G[d].name) + '_report.xlsx', new Blob([R.toXlsx(gxSheets(d))])); }));
  all('.gx-dl-group').forEach((b) => b.addEventListener('click', () => { const d = scopeOf(b); const g = d && G[d]; if (g && g.group) save(fileStem(g.name) + '_by_' + fileStem(g.by) + '.csv', R.toCSV(groupCols(g), g.group)); }));
  all('.gx-dl-cols').forEach((b) => b.addEventListener('click', () => { const d = scopeOf(b); const g = d && G[d]; if (g && g.a) save(fileStem(g.name) + '_columns.csv', R.toCSV(R.COLS.profile, g.a.profile)); }));

  // ---------- saving for the team (summaries only; no row data or personal columns) ----------

  const newId = () => 'd' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const saveDoc = async (doc, setState) => {
    if (!CAP.db) return;
    setState('saving');
    try { await CAP.db.collection('dept_reports').doc(newId()).set({ ...doc, at: new Date().toISOString(), by: CAP.uid || null }); setState('saved'); }
    catch (e) { setState('failed'); }
  };
  all('.dsp-save').forEach((b) => b.addEventListener('click', () => {
    if (!DSP.override) return;
    const results = { ...DSP.override };
    saveDoc({ dept: 'dispatch', kind: 'dispatch', name: DSP.name, files: DSP.files, rowsCount: DSP.rowsCount, results: compact(results) }, (s) => { DSP.saveState = s; dspSync(); });
  }));
  all('.gx-save').forEach((b) => b.addEventListener('click', () => {
    const d = scopeOf(b); const g = d && G[d];
    if (!g || !g.a) return;
    const a = { ...g.a, categories: g.a.categories.filter((c) => !PERSONAL.test(c.column)) };
    saveDoc({ dept: d, kind: 'generic', name: g.name, files: g.fileNames, rowsCount: a.rows, analysis: a, group: g.group && g.by ? { by: g.by, sum: g.sum || null, rows: g.group.slice(0, 500) } : null }, (s) => { g.saveState = s; gxSync(d); });
  }));
  const openSavedDispatch = (doc) => {
    DSP.override = expand(doc.results); DSP.rows = null; DSP.rowsCount = doc.rowsCount; DSP.name = doc.name; DSP.files = doc.files || []; DSP.savedId = doc.id; DSP.saveState = '';
    dash.setParams({ dmonth: 'all', courier: 'all' });
    dspRender();
  };
  const openSavedG = (d, doc) => {
    const g = G[d];
    Object.assign(g, { files: [], table: null, view: null, full: doc.analysis, from: '', to: '', prev: null, error: '', hint: '', book: null, savedId: doc.id, saveState: '', name: doc.name || 'Saved report', fileNames: doc.files || [doc.name], a: doc.analysis });
    g.by = doc.group ? doc.group.by : ''; g.sum = doc.group ? doc.group.sum || '' : ''; g.group = doc.group ? doc.group.rows : null;
    gxRender(d);
  };
  const deleteSaved = async (doc, state, redraw) => {
    if (state.confirmDel !== doc.id) { state.confirmDel = doc.id; redraw(); return; }
    state.confirmDel = null;
    try { await CAP.db.collection('dept_reports').doc(doc.id).delete(); } catch (e) { /* the list stays as the store has it */ }
  };
  dash.draw('dsp-saved', (ul) => {
    ul.replaceChildren();
    DSP.saved.forEach((d) => savedItem(ul, d, DSP.savedId === d.id, [['Open', 'open'], ...(CAP.canEdit ? [[DSP.confirmDel === d.id ? 'Confirm delete' : 'Delete', 'delete', 'is-quiet']] : [])]));
  });
  dash.draw('gx-saved', (ul) => {
    ul.replaceChildren();
    const d = scopeOf(ul); if (!d) return;
    const g = G[d];
    g.saved.forEach((x) => savedItem(ul, x, g.savedId === x.id, [['Open', 'open'], ['Download Excel', 'download'], ...(CAP.canEdit ? [[g.confirmDel === x.id ? 'Confirm delete' : 'Delete', 'delete', 'is-quiet']] : [])]));
  });
  all('[data-draw="dsp-saved"]').forEach((ul) => ul.addEventListener('click', (e) => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const doc = DSP.saved.find((x) => x.id === b.dataset.id); if (!doc) return;
    if (b.dataset.act === 'open') { openSavedDispatch(doc); goTab('dsp-overview'); }
    else if (b.dataset.act === 'delete') deleteSaved(doc, DSP, () => dash.draw('dsp-saved'));
  }));
  all('[data-draw="gx-saved"]').forEach((ul) => ul.addEventListener('click', (e) => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const d = scopeOf(ul); const g = d && G[d];
    const doc = g && g.saved.find((x) => x.id === b.dataset.id); if (!doc) return;
    if (b.dataset.act === 'open') { openSavedG(d, doc); goTab(d + '-report'); }
    else if (b.dataset.act === 'download') { openSavedG(d, doc); save(fileStem(g.name) + '_report.xlsx', new Blob([R.toXlsx(gxSheets(d))])); }
    else if (b.dataset.act === 'delete') deleteSaved(doc, g, () => dash.draw('gx-saved'));
  }));

  // ---------- What stands out: findings written from the numbers in view ----------

  // parts: plain strings, or [text] for a bold figure. tone: good, bad, warn or ''.
  const insights = (ul, items) => {
    ul.replaceChildren();
    const list = items.filter(Boolean).slice(0, 8);
    if (!list.length) list.push({ tone: '', parts: ['Nothing to report for this view.'] });
    list.forEach(({ tone, parts }) => {
      const li = document.createElement('li');
      if (tone) li.className = 'is-' + tone;
      const span = document.createElement('span');
      parts.forEach((x) => { if (Array.isArray(x)) { const b = document.createElement('b'); b.textContent = x[0]; span.append(b); } else span.append(x); });
      li.append(span);
      ul.append(li);
    });
  };
  const signed = (v, f) => (v > 0 ? '+' : '') + f(v);
  const pts = (v) => signed(v * 100, d3.format('.1f')) + ' pts';
  const WEEKDAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

  dash.draw('ins-orders', (ul) => {
    const od = window.dateView.orders;
    const q = dash.params();
    const month = q.month || 'all', subject = q.subject || 'all', tier = q.tier || 'all';
    const cube = od('cube');
    if (cube.status !== 'ok') { ul.replaceChildren(); return; }
    const r = cube.data.find((x) => x.k === month + '|' + subject + '|' + tier);
    if (!r || !r.attempts) { insights(ul, [{ tone: '', parts: ['No orders in this view. Widen the dates or filters.'] }]); return; }
    const out = [];
    if (r.paid_change !== null && r.paid_change !== undefined) {
      const ranged = (q.from || q.to) && month === 'all';
      out.push({ tone: r.paid_change >= 0 ? 'good' : 'bad', parts: ['Paid orders ', [signed(r.paid_change, fmtPct)], ' ' + (ranged ? 'on the period before' : 'on the month before') + (r.prev_paid !== undefined && r.prev_paid !== null ? ' (' + fmtInt(r.paid) + ' vs ' + fmtInt(r.prev_paid) + ')' : '') + (r.success_change !== null && r.success_change !== undefined ? '; payment success ' : '.'), ...(r.success_change !== null && r.success_change !== undefined ? [[pts(r.success_change)], '.'] : [])] });
    }
    if (r.never_paid) out.push({ tone: 'bad', parts: [[fmtInt(r.never_paid)], ' students (' + fmtPct(r.never_paid / r.students) + ' of those at checkout) never paid. That is the win-back list in the order upload tab.'] });
    if (r.attempts - r.paid > 0) out.push({ tone: r.lost > r.recovered ? 'warn' : 'good', parts: ['Of ' + fmtInt(r.attempts - r.paid) + ' unpaid orders, ', [fmtInt(r.recovered)], ' were paid later by the same student and ', [fmtInt(r.lost)], ' never were.'] });
    const plans = od('plans');
    if (plans.status === 'ok') {
      const ps = plans.data.filter((x) => x.month === month && (subject === 'all' || x.subject === subject) && (tier === 'all' || x.tier === tier));
      const top = ps.reduce((a, x) => (!a || x.paid > a.paid ? x : a), null);
      const weak = ps.filter((x) => x.attempts >= 20).reduce((a, x) => (!a || x.success < a.success ? x : a), null);
      if (top && top.paid) out.push({ tone: 'good', parts: ['Top plan: ', [top.plan], ', ' + fmtInt(top.paid) + ' paid orders (' + fmtPct(top.share) + ' of paid plans).'] });
      if (weak && weak !== top && weak.success < (r.success || 1)) out.push({ tone: 'warn', parts: ['Weakest plan to pay: ', [weak.plan], ', only ' + fmtPct(weak.success) + ' of ' + fmtInt(weak.attempts) + ' orders paid.'] });
    }
    const daily = od('daily');
    if (daily.status === 'ok') {
      const ds = daily.data.filter((x) => x.subject === subject && x.tier === tier && (month === 'all' || x.month === month) && x.attempts);
      const peak = ds.reduce((a, x) => (!a || x.attempts > a.attempts ? x : a), null);
      if (peak) out.push({ tone: '', parts: ['Busiest day: ', [d3.utcFormat('%a %-d %b')(new Date(peak.date + 'T00:00:00Z'))], ' with ' + fmtInt(peak.attempts) + ' orders, ' + fmtInt(peak.paid) + ' paid.'] });
    }
    const wd = od('weekday');
    if (wd.status === 'ok') {
      const ws = wd.data.filter((x) => x.month === month && x.subject === subject && x.tier === tier && x.attempts >= 30);
      if (ws.length >= 3) {
        const best = ws.reduce((a, x) => (x.success > a.success ? x : a)), worst = ws.reduce((a, x) => (x.success < a.success ? x : a));
        if (best.success - worst.success >= 0.05) out.push({ tone: '', parts: ['Orders pay best on ', [WEEKDAYS[best.weekday]], ' (' + fmtPct(best.success) + ') and worst on ', [WEEKDAYS[worst.weekday]], ' (' + fmtPct(worst.success) + '): a good day for reminders.'] });
      }
    }
    insights(ul, out);
  });

  dash.draw('ins-dispatch', (ul) => {
    const q = dq();
    const cube = D('dsp_cube');
    if (cube.status !== 'ok') { ul.replaceChildren(); return; }
    const r = cubeRow(q);
    if (!r || !r.orders) { insights(ul, [{ tone: '', parts: ['No school orders in this view. Widen the dates or filters.'] }]); return; }
    const out = [];
    const couriers = cube.data.filter((x) => x.exam === q.exam && x.month === q.month && x.courier !== 'all' && x.shipped > 0 && (q.courier === 'all' || x.courier === q.courier));
    if (r.transit_late) {
      const worst = couriers.reduce((a, x) => (!a || x.transit_late > a.transit_late ? x : a), null);
      out.push({ tone: 'bad', parts: [[fmtInt(r.transit_late)], ' shipments have been in transit for over 7 days' + (worst && q.courier === 'all' && couriers.length > 1 ? '; ' : '.'), ...(worst && q.courier === 'all' && couriers.length > 1 ? [[worst.courier], ' holds ' + fmtInt(worst.transit_late) + ' of them.'] : [])] });
    }
    const timed = couriers.filter((x) => x.delivered >= 20 && x.avg_days !== null);
    if (timed.length > 1) {
      const fast = timed.reduce((a, x) => (x.avg_days < a.avg_days ? x : a)), slow = timed.reduce((a, x) => (x.avg_days > a.avg_days ? x : a));
      out.push({ tone: '', parts: ['Fastest courier: ', [fast.courier], ' at ' + fmtv(fast.avg_days, 'days') + ' on average; slowest: ', [slow.courier], ' at ' + fmtv(slow.avg_days, 'days') + '.'] });
    }
    const regions = D('dsp_regions');
    if (regions.status === 'ok') {
      const rs = regions.data.filter((x) => x.k.startsWith(dk(q) + '|') && x.shipped >= 20);
      const low = rs.reduce((a, x) => (!a || x.delivery_rate < a.delivery_rate ? x : a), null);
      if (low && low.delivery_rate < (r.delivery_rate || 1)) out.push({ tone: 'warn', parts: ['Lowest delivery rate by region: ', [regionLabel(low.region)], ', ' + fmtPct(low.delivery_rate) + ' of ' + fmtInt(low.shipped) + ' shipped (' + fmtInt(low.in_transit) + ' in transit).'] });
    }
    const open = openRows('open', false);
    if (open && open.length) {
      const old = open.reduce((a, x) => ((x.days_open ?? -1) > (a.days_open ?? -1) ? x : a));
      if (old.days_open !== null) out.push({ tone: 'bad', parts: ['Oldest open order: ', [old.school], ' (' + old.stage.toLowerCase() + (old.courier ? ', ' + old.courier : '') + '), open for ' + fmtInt(old.days_open) + ' days.'] });
    }
    if (r.pending_ship) out.push({ tone: 'warn', parts: [[fmtInt(r.pending_ship)], ' orders aren\'t shipped yet: ' + fmtInt(r.on_hold) + ' on hold and ' + fmtInt(r.not_shipped) + ' packed, waiting for a courier.'] });
    if (r.issues) {
      const most = couriers.reduce((a, x) => (!a || x.issues > a.issues ? x : a), null);
      out.push({ tone: 'warn', parts: [[fmtInt(r.issues)], ' shipments were returned (RTO), re-routed or lost' + (most && most.issues && couriers.length > 1 ? '; most were with ' : '.'), ...(most && most.issues && couriers.length > 1 ? [[most.courier], ' (' + fmtInt(most.issues) + ').'] : [])] });
    }
    if (r.within_7 !== null) out.push({ tone: r.within_7 >= 0.8 ? 'good' : 'warn', parts: [[fmtPct(r.within_7)], ' of deliveries arrived within 7 days of dispatch (median ' + fmtv(r.median_days, 'days') + ').'] });
    const mat = D('dsp_material');
    if (mat.status === 'ok') {
      const ms = mat.data.filter((x) => (q.exam === 'all' || x.exam === q.exam) && x.item !== 'School orders' && x.pending > 0);
      const big = ms.reduce((a, x) => (!a || x.pending > a.pending ? x : a), null);
      if (big) out.push({ tone: '', parts: ['Most material still to pack: ', [examLabel(big.exam) + ' ' + big.item.toLowerCase()], ', ' + fmtInt(big.pending) + ' copies (' + fmtv(big.packed_share, 'pct') + ' packed).'] });
    }
    insights(ul, out);
  });

  dash.draw('ins-gx', (ul) => {
    const d = scopeOf(ul); const g = d && G[d]; const a = g && g.a;
    if (!a) { ul.replaceChildren(); return; }
    if (!a.rows) { insights(ul, [{ tone: '', parts: ['No entries in these dates.'] }]); return; }
    const out = [];
    const noun = d === 'support' ? 'requests' : 'entries';
    if (g.prev && g.prev.rows) {
      const ch = a.rows / g.prev.rows - 1;
      out.push({ tone: '', parts: [[fmtInt(a.rows)], ' ' + noun + ' in these dates, ', [signed(ch, fmtPct)], ' on the period before (' + fmtInt(g.prev.rows) + ').'] });
    } else if (a.byMonth.length >= 2) {
      const [p, l] = a.byMonth.slice(-2);
      out.push({ tone: '', parts: [[R.monthLabel(l.month)], ': ' + fmtInt(l.rows) + ' ' + noun + ', ', [signed(l.rows / p.rows - 1, fmtPct)], ' on ' + R.monthLabel(p.month) + '.'] });
    }
    if (a.byMonth.length >= 2) { const peak = a.byMonth.reduce((x, y) => (y.rows > x.rows ? y : x)); out.push({ tone: '', parts: ['Busiest month: ', [R.monthLabel(peak.month)], ' with ' + fmtInt(peak.rows) + ' ' + noun + '.'] }); }
    if (a.statusCol && a.statusFilled) out.push({ tone: a.doneRate >= 0.8 ? 'good' : a.doneRate >= 0.5 ? 'warn' : 'bad', parts: [[fmtPct(a.doneRate)], ' completed and ', [fmtInt(a.open)], ' still open, going by ' + a.statusCol + '.'] });
    // Open entries more than a week old, counted from the rows (not kept in a saved report).
    const v = g.view || g.table;
    if (v && a.statusCol && a.dateCol && a.to) {
      const dateOf = gxRowDate(g);
      const cutoff = isoAdd(g.full.to, -7);
      const stale = v.rows.filter((r) => !isDoneValue(r[a.statusCol]) && txt(r[a.statusCol]) && dateOf(r) && dateOf(r) < cutoff).length;
      if (stale) out.push({ tone: 'bad', parts: [[fmtInt(stale)], ' open ' + noun + ' are more than 7 days old.'] });
    }
    if (a.closedCount) out.push({ tone: '', parts: ['Average time to close: ', [fmtv(a.avgClose, 'days')], ' (median ' + fmtv(a.medianClose, 'days') + ', from ' + fmtInt(a.closedCount) + ' closed).'] });
    a.categories.slice(0, 3).forEach((c) => {
      const top = c.values.find((x) => x.value !== '(blank)' && !/^Other \(/.test(x.value));
      if (top) out.push({ tone: '', parts: ['Most common ' + c.column + ': ', [top.value], ' (' + fmtPct(top.share) + ' of ' + noun + ').'] });
    });
    if (a.statusCol) {
      const weak = a.categories.flatMap((c) => c.values.filter((x) => x.rows >= 10 && x.value !== '(blank)' && !/^Other \(/.test(x.value)).map((x) => ({ ...x, column: c.column })))
        .reduce((m, x) => (!m || x.done_share < m.done_share ? x : m), null);
      if (weak && weak.done_share < (a.doneRate || 0) - 0.1) out.push({ tone: 'warn', parts: ['Lowest completion: ', [weak.value], ' (' + weak.column + '), ' + fmtPct(weak.done_share) + ' of ' + fmtInt(weak.rows) + ' completed.'] });
    }
    insights(ul, out);
  });

  dspSync();
  ['mis', 'support'].forEach(gxSync);

  // capabilities light up when the viewer's page grants them
  (async () => {
    const use = window.claude && typeof window.claude.use === 'function' ? (n) => window.claude.use(n).catch(() => null) : async () => null;
    const [dl, db, user] = await Promise.all([use('downloads'), use('db'), use('user')]);
    CAP.dl = dl; CAP.db = db;
    if (user) { CAP.canEdit = await user.canEdit(); CAP.uid = await user.id(); }
    if (!db) { DSP.savedState = 'off'; dspRender(); ['mis', 'support'].forEach(gxRender); return; }
    db.collection('dept_reports').orderBy('at', 'desc').limit(200).onSnapshot(async (snap) => {
      const docs = snap.docs.map((x) => ({ id: x.id, ...x.data() }));
      DSP.saved = docs.filter((x) => x.dept === 'dispatch' && x.kind === 'dispatch');
      DSP.savedState = DSP.saved.length ? 'list' : 'empty';
      ['mis', 'support'].forEach((d) => {
        const g = G[d];
        g.saved = docs.filter((x) => x.dept === d && x.kind === 'generic' && x.analysis);
        // With no file of their own open, a department page shows the team's latest saved report.
        if (!g.a && !g.closed && g.saved.length) openSavedG(d, g.saved[0]);
      });
      dspRender(); ['mis', 'support'].forEach(gxRender);
      const ids = [...new Set(docs.map((x) => x.by).filter(Boolean))].filter((id) => !CAP.names[id]);
      if (user && ids.length) {
        const ps = await user.profiles(ids);
        ids.forEach((id) => { CAP.names[id] = (ps[id] && ps[id].name) || 'Someone'; });
        dash.draw('dsp-saved'); dash.draw('gx-saved');
      }
    }, () => { DSP.savedState = 'off'; dspRender(); ['mis', 'support'].forEach(gxRender); });
    dspRender(); ['mis', 'support'].forEach(gxRender);
  })();
})();
