// Minimal stand-in for the claude.ai dashboard runtime, so dashboard/page.html runs as a static site.
// Supports what the page uses: calc, data, draw, onData, params, setParams, resetParams, setLink, colors.
// Data comes precomputed (scripts/precompute.js runs the page's calcs), so calcs are only registered here.
(function () {
  const DEFAULT_PARAMS = { month: 'all', subject: 'all', tier: 'all', metric: 'paid', dmonth: 'all', courier: 'all', exam: 'all' };
  const results = {};
  const draws = new Map();
  const listeners = [];
  let status = 'loading';
  let message = '';
  let params = { ...DEFAULT_PARAMS };
  let queued = false;

  // Like the real runtime, drawings under a hidden ancestor are not run until it is shown.
  function render() {
    queued = false;
    listeners.forEach((fn) => { try { fn(); } catch (e) { console.error(e); } });
    draws.forEach((fn, name) => drawOne(name, fn));
  }
  function drawOne(name, fn) {
    document.querySelectorAll(`[data-draw="${name}"]`).forEach((el) => {
      if (el.closest('[hidden]')) return;
      try { fn(el); } catch (e) { console.error('draw ' + name, e); }
    });
  }
  function schedule() {
    if (queued) return;
    queued = true;
    requestAnimationFrame(render);
  }

  window.dash = {
    colors: ['#3366cc', '#e0861a', '#2a9d8f', '#9b5de5', '#d1495b', '#6c8a2e', '#b5838d', '#577590'],
    calc() {},
    loader() {},
    data(id) {
      if (status !== 'ok') return { status, data: [], message, refreshing: false, meta: [] };
      const data = results[id];
      return data ? { status: 'ok', data, columns: data.length ? Object.keys(data[0]) : [], refreshing: false, meta: [] } : { status: 'missing', data: [], meta: [] };
    },
    onData(fn) { listeners.push(fn); try { fn(); } catch (e) { console.error(e); } },
    draw(name, fn) {
      if (typeof fn === 'function') draws.set(name, fn);
      else if (draws.has(name)) drawOne(name, draws.get(name));
    },
    params() { return { ...params }; },
    setParams(next) { params = { ...params, ...next }; schedule(); },
    resetParams() { params = { ...DEFAULT_PARAMS }; schedule(); },
    setLink(name) { history.replaceState(null, '', '#' + name); },
    refresh() {},
    async start(url) {
      render();
      // Redraw when a tab panel or the filter bar is shown or hidden.
      const mo = new MutationObserver(schedule);
      document.querySelectorAll('[data-panel], .ito-filters, [data-dept-only], [data-gx-when], [data-gx-needs]').forEach((el) => mo.observe(el, { attributes: true, attributeFilter: ['hidden'] }));
      try {
        const res = await fetch(url);
        if (!res.ok) throw new Error('HTTP ' + res.status);
        Object.entries(await res.json()).forEach(([id, { c, r }]) => {
          results[id] = r.map((v) => Object.fromEntries(c.map((col, i) => [col, v[i]])));
        });
        status = 'ok';
      } catch (e) {
        status = 'error';
        message = String(e);
        console.error('Could not load data', e);
      }
      render();
      let t;
      window.addEventListener('resize', () => { clearTimeout(t); t = setTimeout(render, 150); });
    },
  };
})();
