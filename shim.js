// Capa mínima que reemplaza el motor del dashboard de Claude: inicio de sesión Microsoft + lectura de archivos con Graph.
(function () {
  const CFG = window.APP_CONFIG;
  const SCOPES = ['Files.Read.All'];
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const loaders = {}, calcs = {}, cbs = [];
  const params = {};
  const state = {
    matriz: { status: 'loading' }, recibidas: { status: 'loading' }, inspecciones: { status: 'loading' },
    ocs: { status: 'loading' }, resumen: { status: 'loading' },
  };
  const emit = () => cbs.forEach((f) => f());
  window.dash = {
    colors: ['#c96442'],
    loader: (id, o) => { loaders[id] = o.fn; },
    calc: (id, o) => { calcs[id] = o; },
    setParams: (p) => { Object.assign(params, p); emit(); },
    params: () => params,
    data: (id) => state[id],
    onData: (f) => { cbs.push(f); f(); },
  };
  window.addEventListener('resize', () => emit());

  const $ = (id) => document.getElementById(id);
  let app = null;

  async function token() {
    const accounts = app.getAllAccounts();
    if (!accounts.length) return null;
    try {
      return (await app.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] })).accessToken;
    } catch (e) {
      await app.acquireTokenRedirect({ scopes: SCOPES, account: accounts[0] });
      return null;
    }
  }

  async function graph(path, tk, as) {
    const r = await fetch(GRAPH + path, { headers: { Authorization: 'Bearer ' + tk } });
    if (r.status === 401 || r.status === 403) { const e = new Error('Sin permiso (' + r.status + ')'); e.declined = true; throw e; }
    if (!r.ok) throw new Error('Graph ' + r.status + ' en ' + path.slice(0, 60));
    return as === 'buf' ? r.arrayBuffer() : as === 'json' ? r.json() : r.text();
  }

  function leerMatriz(buf) {
    const wb = XLSX.read(buf, { type: 'array', sheets: 'Matriz' });
    const ws = wb.Sheets['Matriz'];
    if (!ws) throw new Error('No se encontró la hoja Matriz en el libro.');
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    const low = (v) => String(v).trim().toLowerCase();
    let h = aoa.findIndex((r, i) => i < 20 && r.some((c) => low(c) === 'oc') && r.some((c) => low(c) === 'status'));
    if (h < 0) h = 0;
    const columns = aoa[h].map((c) => String(c).trim());
    const ocI = columns.findIndex((c) => c.toUpperCase() === 'OC');
    if (ocI >= 0 && !columns.some((c) => c.toUpperCase() === 'PROVEEDOR')) columns[ocI + 1] = 'PROVEEDOR';
    const rows = aoa.slice(h + 1)
      .filter((r) => r.some((c) => String(c).trim() !== ''))
      .map((r) => { const o = {}; columns.forEach((c, j) => { if (c) o[c] = String(r[j] == null ? '' : r[j]).trim(); }); return o; });
    return { columns, rows };
  }

  function recompute() {
    const ids = ['matriz', 'recibidas', 'inspecciones'];
    const sts = ids.map((i) => state[i].status);
    const worst = sts.every((s) => s === 'ok') ? 'ok'
      : sts.includes('connect') ? 'connect' : sts.includes('declined') ? 'declined'
      : sts.includes('error') ? 'error' : 'loading';
    if (worst === 'ok') {
      try {
        const ocs = calcs.ocs.fn(state.matriz.data, state.recibidas.data, state.inspecciones.data);
        state.ocs = { status: 'ok', data: ocs };
        state.resumen = { status: 'ok', data: calcs.resumen.fn(ocs) };
      } catch (e) { window.__lastError = e.message; state.ocs = state.resumen = { status: 'error' }; }
    } else state.ocs = state.resumen = { status: worst };
    $('ic-login').hidden = worst !== 'connect';
    emit();
  }


  async function boot() {
    if (!CFG || /PEGAR/.test(CFG.clientId)) {
      window.__lastError = 'falta configurar clientId en config.js';
      ['matriz', 'recibidas', 'inspecciones'].forEach((i) => { state[i] = { status: 'error' }; });
      return recompute();
    }
    app = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: 'https://login.microsoftonline.com/' + (CFG.tenant || 'organizations'), redirectUri: CFG.redirectUri || location.origin + location.pathname },
      cache: { cacheLocation: 'localStorage' },
    });
    if (typeof app.initialize === 'function') await app.initialize();
    try { await app.handleRedirectPromise(); } catch (e) { window.__lastError = e.message; }
    $('ic-login').addEventListener('click', () => { if (window.top !== window) window.top.postMessage({ type: 'ic-login' }, location.origin); else app.loginRedirect({ scopes: SCOPES }); });
    const tk = await token();
    if (!tk) { ['matriz', 'recibidas', 'inspecciones'].forEach((i) => { state[i] = { status: 'connect' }; }); return recompute(); }
    const d = CFG.driveId;
    await Promise.all(['matriz', 'recibidas', 'inspecciones'].map(async (id) => {
      try {
        const as = id === 'matriz' ? 'buf' : 'text';
        const raw = await graph('/drives/' + d + '/items/' + CFG.files[id] + '/content', tk, as);
        const answer = id === 'matriz' ? leerMatriz(raw) : raw.replace(/^﻿/, '');
        let rows = loaders[id]({ answer }).rows;
        if (id === 'inspecciones' && CFG.files.historial) {
          // La lista se depura cada mes: se une el historial guardado con la lista actual (por Id; manda la lista).
          const hist = await graph('/drives/' + d + '/items/' + CFG.files.historial + '/content', tk, 'text');
          const m = new Map();
          loaders[id]({ answer: hist.replace(/^\uFEFF/, '') }).rows.forEach((r) => m.set(r.Id, r));
          rows.forEach((r) => m.set(r.Id, r));
          rows = [...m.values()];
        }
        state[id] = { status: 'ok', data: rows };
      } catch (e) {
        window.__lastError = e.message;
        state[id] = { status: e.declined ? 'declined' : 'error' };
      }
      recompute();
    }));
    try {
      const fechas = await Promise.all(['recibidas', 'inspecciones'].map((id) =>
        graph('/drives/' + d + '/items/' + CFG.files[id] + '?select=lastModifiedDateTime', tk, 'json').then((j) => new Date(j.lastModifiedDateTime))));
      const f = new Date(Math.min(...fechas));
      $('ic-upd').textContent = 'Datos actualizados por el flujo diario: ' + f.toLocaleString('es-CR', { dateStyle: 'short', timeStyle: 'short' }) + ' · Se refresca cada noche a las 10 pm.';
    } catch (e) { /* opcional */ }
  }
  document.addEventListener('DOMContentLoaded', () => window.icGate(boot));
})();
