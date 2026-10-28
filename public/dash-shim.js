// Minimal stand-in for the claude.ai dashboard runtime, so dashboard/page.html runs as a static site.
// Supports what the page uses: calc, data, draw, onData, params, setParams, resetParams, colors.
(function () {
  const DEFAULT_PARAMS = { month: 'all', subject: 'all', tier: 'all' };
  const calcs = [];
  const results = {};
  const draws = new Map();
  const listeners = [];
  let status = 'loading';
  let message = '';
  let params = { ...DEFAULT_PARAMS };

  function render() {
    listeners.forEach((fn) => { try { fn(); } catch (e) { console.error(e); } });
    draws.forEach((fn, name) => {
      document.querySelectorAll(`[data-draw="${name}"]`).forEach((el) => {
        try { fn(el); } catch (e) { console.error('draw ' + name, e); }
      });
    });
  }

  window.dash = {
    colors: ['#3366cc', '#e0861a', '#2a9d8f', '#9b5de5', '#d1495b', '#6c8a2e', '#b5838d', '#577590'],
    calc(id, def) { calcs.push({ id, def }); },
    loader() {},
    data(id) {
      if (status !== 'ok') return { status, data: [], message, refreshing: false, meta: [] };
      const data = results[id];
      return data ? { status: 'ok', data, columns: data.length ? Object.keys(data[0]) : [], refreshing: false, meta: [] } : { status: 'missing', data: [], meta: [] };
    },
    onData(fn) { listeners.push(fn); try { fn(); } catch (e) { console.error(e); } },
    draw(name, fn) { draws.set(name, fn); },
    params() { return { ...params }; },
    setParams(next) { params = { ...params, ...next }; render(); },
    resetParams() { params = { ...DEFAULT_PARAMS }; render(); },
    setLink(name) { history.replaceState(null, '', '#' + name); },
    refresh() {},
    async start(url) {
      render();
      try {
        const res = await fetch(url);
        if (!res.ok) throw new Error('HTTP ' + res.status);
        results.orders = await res.json();
        calcs.forEach(({ id, def }) => { results[id] = def.fn(...def.inputs.map((i) => results[i])); });
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
