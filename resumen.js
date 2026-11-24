// Resumen mensual de producto off spec y comparación con la inspección de contenedores.
(function () {
  const CFG = window.APP_CONFIG, G = 'https://graph.microsoft.com/v1.0', SC = ['Sites.ReadWrite.All', 'Files.Read.All'];
  const $ = (id) => document.getElementById(id);
  const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'setiembre', 'octubre', 'noviembre', 'diciembre'];
  const st = { tk: null, off: [], rec: [], insp: new Set() };
  const mk = (d) => d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0');
  const mlabel = (k) => MESES[Number(k.slice(5)) - 1] + ' ' + k.slice(0, 4);
  const fmt = (n) => Number(n).toLocaleString('es-CR', { maximumFractionDigits: 1 });
  const pct = (n) => n.toLocaleString('es-CR', { maximumFractionDigits: 1 }) + '%';
  const serial = (n) => (Number(n) > 20000 ? new Date(Math.round((Number(n) - 25569) * 86400000)) : null);

  async function g(path, raw) {
    const r = await fetch(path.startsWith('http') ? path : G + path, { headers: { Authorization: 'Bearer ' + st.tk } });
    if (r.status === 401 || r.status === 403) { const e = new Error('Sin permiso'); e.denied = true; throw e; }
    if (!r.ok) throw new Error('Error ' + r.status);
    return raw ? r.text() : r.json();
  }
  function parseCsv(t) {
    const rows = []; let r = [], c = '', q = false;
    for (let i = 0; i < t.length; i++) {
      const ch = t[i];
      if (q) { if (ch === '"') { if (t[i + 1] === '"') { c += '"'; i++; } else q = false; } else c += ch; }
      else if (ch === '"') q = true; else if (ch === ',') { r.push(c); c = ''; }
      else if (ch === '\n') { r.push(c); rows.push(r); r = []; c = ''; } else if (ch !== '\r') c += ch;
    }
    if (c || r.length) { r.push(c); rows.push(r); }
    const h = rows[0].map((x) => x.trim().replace(/^﻿/, ''));
    return rows.slice(1).filter((x) => x.length > 1).map((x) => { const o = {}; h.forEach((k, i) => { o[k] = (x[i] || '').trim(); }); return o; });
  }
  async function pages(url) { let out = []; while (url) { const j = await g(url); out = out.concat(j.value); url = j['@odata.nextLink'] || null; } return out; }

  async function load() {
    const site = (await g('/sites/' + CFG.site.host + ':' + CFG.site.path + '?$select=id')).id;
    const dr = (id) => g('/drives/' + CFG.driveId + '/items/' + CFG.files[id] + '/content', true);
    const [items, rec, insp, hist] = await Promise.all([
      pages('/sites/' + site + '/lists/' + CFG.lists.offspec + '/items?$expand=fields&$top=999'),
      dr('recibidas'), dr('inspecciones'), CFG.files.historial ? dr('historial') : Promise.resolve(''),
    ]);
    st.off = items.map((x) => x.fields).filter((f) => f.FechaDefecto).map((f) => {
      let nf = 0; try { nf = JSON.parse(f.Fotos || '[]').length; } catch (e) { nf = 0; }
      return { mes: String(f.FechaDefecto).slice(0, 7), prod: f.Title || '', cod: (f.CodigoProducto || '').trim(), def: f.TipoDefecto || '(sin defecto)', trat: f.Tratamiento || '(sin tratamiento)', q: Number(f.Cantidad) || 0, und: (f.Unidad || '').trim(), oc: String(f.OC || '').trim(), fotos: nf };
    });
    st.rec = parseCsv(rec).filter((r) => r.OC);
    [insp, hist].forEach((t) => { if (t) parseCsv(t).forEach((r) => { if (r.OrdenCompra) st.insp.add(String(r.OrdenCompra).trim()); }); });
  }

  // OC -> { oc, pr, ata(Date|null), lines: Map(cod -> {q}) }
  function ocs() {
    const m = new Map();
    st.rec.forEach((r) => {
      let o = m.get(r.OC); if (!o) { o = { oc: r.OC, pr: r.PROVEEDOR, ata: null, lines: new Map() }; m.set(r.OC, o); }
      const d = serial(r['Dia en planta']) || serial(r['Mes en planta']) || serial(r.ATA) || serial(r.ETA); if (d && (!o.ata || d > o.ata)) o.ata = d;
      o.lines.set(r.Codigo, (o.lines.get(r.Codigo) || 0) + (Number(r.Cantidad) || 0));
    });
    return m;
  }
  const sumBy = (arr, key) => { const m = new Map(); arr.forEach((x) => { const k = x[key]; const o = m.get(k) || { k, n: 0, u: new Map() }; o.n++; o.u.set(x.und || 'unid.', (o.u.get(x.und || 'unid.') || 0) + x.q); m.set(k, o); }); return [...m.values()].sort((a, b) => b.n - a.n); };
  const qtxt = (u) => [...u].map(([k, v]) => fmt(v) + ' ' + k).join(' · ');
  function rows(tb, list, cols) {
    tb.replaceChildren();
    if (!list.length) { const tr = document.createElement('tr'), td = document.createElement('td'); td.colSpan = 8; td.textContent = 'Sin datos'; td.className = 'no'; tr.appendChild(td); tb.appendChild(tr); return; }
    list.forEach((it) => { const tr = document.createElement('tr'); cols(it).forEach((c) => { const td = document.createElement('td'); if (c && c.n) td.className = 'n'; if (c && c.cls) td.className = c.cls; if (c && c.meter != null) { const m = document.createElement('div'); m.className = 'meter'; const i = document.createElement('i'); i.style.width = c.meter + '%'; m.appendChild(i); td.appendChild(m); } else td.textContent = c && c.t != null ? c.t : c; tr.appendChild(td); }); tb.appendChild(tr); });
  }

  function chart(months, cur) {
    const by = new Map(); st.off.forEach((x) => by.set(x.mes, (by.get(x.mes) || 0) + 1));
    const ks = months.slice().sort().slice(-12), mx = Math.max(1, ...ks.map((k) => by.get(k) || 0)), w = 56, H = 120;
    let s = '<svg viewBox="0 0 ' + ks.length * w + ' ' + (H + 28) + '" width="100%" role="img" aria-label="Registros off spec por mes">';
    ks.forEach((k, i) => { const v = by.get(k) || 0, h = Math.round((v / mx) * H); s += '<rect x="' + (i * w + 10) + '" y="' + (H - h + 4) + '" width="' + (w - 20) + '" height="' + h + '" rx="3" fill="' + (k === cur ? '#c96442' : '#b8b4a6') + '"/><text x="' + (i * w + w / 2) + '" y="' + (H - h) + '" text-anchor="middle">' + v + '</text><text x="' + (i * w + w / 2) + '" y="' + (H + 20) + '" text-anchor="middle">' + MESES[Number(k.slice(5)) - 1].slice(0, 3) + ' ' + k.slice(2, 4) + '</text>'; });
    $('chart').innerHTML = s + '</svg>';
  }

  function show(mes) {
    const O = ocs(), off = st.off.filter((x) => x.mes === mes);
    // OC del mes: llegadas en el mes, o con registros off spec del mes
    const inMonth = [...O.values()].filter((o) => o.ata && mk(o.ata) === mes);
    const withOff = new Set(off.filter((x) => x.oc).map((x) => x.oc));
    const set = new Map(); inMonth.forEach((o) => set.set(o.oc, o)); withOff.forEach((id) => { if (O.has(id)) set.set(id, O.get(id)); });
    const insMes = inMonth.filter((o) => st.insp.has(o.oc)).length;
    const nPhoto = off.filter((x) => x.fotos > 0).length;
    // % defectos por OC: off spec con mismo código y unidad compatible / recibido de esos códigos
    const perOc = [...set.values()].map((o) => {
      const mine = st.off.filter((x) => x.oc === o.oc);
      const matched = mine.filter((x) => o.lines.has(x.cod));
      const recv = [...new Set(matched.map((x) => x.cod))].reduce((a, c) => a + o.lines.get(c), 0);
      const q = matched.reduce((a, x) => a + x.q, 0);
      return { o, n: mine.length, q, recv, p: recv ? (q / recv) * 100 : null, unm: mine.length - matched.length, insp: st.insp.has(o.oc) };
    }).sort((a, b) => (b.p || 0) - (a.p || 0) || b.n - a.n);
    const avg = (arr) => { const v = arr.filter((x) => x.p != null && x.n); return v.length ? v.reduce((a, x) => a + x.p, 0) / v.length : null; };
    const aI = avg(perOc.filter((x) => x.insp)), aN = avg(perOc.filter((x) => !x.insp));

    const kp = [
      [off.length, 'registros off spec en ' + mlabel(mes)],
      [nPhoto + ' de ' + off.length, 'registros con foto'],
      [inMonth.length, 'OC recibidas en el mes'],
      [inMonth.length ? pct((insMes / inMonth.length) * 100) : '—', 'OC inspeccionadas (' + insMes + ' de ' + inMonth.length + ')'],
      [perOc.filter((x) => x.n).length, 'OC con producto off spec'],
      [aI == null ? '—' : pct(aI), '% defectos promedio, OC inspeccionadas' + (aN == null ? '' : ' (no inspeccionadas: ' + pct(aN) + ')')],
    ];
    $('kpis').replaceChildren(...kp.map(([v, l]) => { const d = document.createElement('div'); d.className = 'card kpi'; const a = document.createElement('div'); a.className = 'v'; a.textContent = v; const b = document.createElement('div'); b.className = 'l'; b.textContent = l; d.append(a, b); return d; }));
    const mx = (arr) => Math.max(1, ...arr.map((x) => x.n));
    const d = sumBy(off, 'def'), t = sumBy(off, 'trat');
    rows($('t-def'), d, (x) => [x.k, { t: x.n, n: 1 }, qtxt(x.u), { meter: (x.n / mx(d)) * 100 }]);
    rows($('t-trat'), t, (x) => [x.k, { t: x.n, n: 1 }, qtxt(x.u), { meter: (x.n / mx(t)) * 100 }]);
    rows($('t-oc'), perOc, (x) => [x.o.oc, x.o.pr, x.o.ata ? x.o.ata.toISOString().slice(0, 10) : '', { t: x.insp ? 'Inspeccionada' : 'Sin inspección', cls: x.insp ? 'yes' : 'no' }, { t: fmt(x.recv || 0), n: 1 }, { t: fmt(x.q), n: 1 }, { t: x.p == null ? '—' : pct(x.p), n: 1 }, x.n + (x.unm ? ' (' + x.unm + ' sin producto en la OC)' : '')]);
    const sin = off.filter((x) => !x.oc).length;
    $('oc-note').textContent = '% defectos = cantidad off spec registrada contra esa OC ÷ cantidad recibida de los mismos productos en la OC. ' + (sin ? sin + ' registro' + (sin === 1 ? '' : 's') + ' del mes no tiene OC y no entra en esta tabla.' : 'Todos los registros del mes tienen OC.');
    const byP = new Map(); off.forEach((x) => { const o = byP.get(x.cod || x.prod) || { k: (x.cod ? x.cod + ' · ' : '') + x.prod, n: 0, u: new Map() }; o.n++; o.u.set(x.und || 'unid.', (o.u.get(x.und || 'unid.') || 0) + x.q); byP.set(x.cod || x.prod, o); });
    rows($('t-prod'), [...byP.values()].sort((a, b) => b.n - a.n).slice(0, 10), (x) => [x.k, { t: x.n, n: 1 }, qtxt(x.u)]);
    const all = new Set([...st.off.map((x) => x.mes), ...[...O.values()].filter((o) => o.ata).map((o) => mk(o.ata))]);
    chart([...all], mes);
  }

  async function start() {
    try { await window.icAuth.init(); st.tk = await window.icAuth.token(SC); } catch (e) { return need(e.message); }
    if (!st.tk) return need();
    try { await load(); } catch (e) { return need(e.denied ? 'Falta el permiso para leer SharePoint.' : 'No se pudo cargar. ' + e.message, e.denied); }
    $('gate').style.display = 'none'; $('main').hidden = false;
    const O = ocs(), ms = new Set([...st.off.map((x) => x.mes), ...[...O.values()].filter((o) => o.ata).map((o) => mk(o.ata)), mk(new Date())]);
    const sel = $('mes'); [...ms].sort().reverse().forEach((k) => sel.appendChild(new Option(mlabel(k), k)));
    sel.value = mk(new Date()); sel.addEventListener('change', () => show(sel.value)); show(sel.value);
  }
  function need(msg, nb) { $('gate').style.display = 'block'; $('main').hidden = true; if (msg) $('gate-msg').textContent = msg; $('btn-login').hidden = !!nb; }
  $('btn-login').addEventListener('click', () => window.icAuth.login(SC));
  document.addEventListener('DOMContentLoaded', () => window.icGate(start));
})();
