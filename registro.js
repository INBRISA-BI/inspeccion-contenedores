// Registro de producto off spec: productos desde listas de SharePoint; registros en CALI_ProductoOffSpec; fotos en CALI_OffSpecFotos.
(function () {
  const CFG = window.APP_CONFIG;
  const G = 'https://graph.microsoft.com/v1.0';
  const SC = ['Sites.ReadWrite.All', 'Files.Read.All'];
  const $ = (id) => document.getElementById(id);
  const DEFECTOS = ['Fuera de especificación dimensional', 'Color o tono diferente', 'Contaminación o suciedad', 'Mal sellado o empaque defectuoso', 'Mal corte o rebaba', 'Peso fuera de rango', 'Etiquetado o codificación incorrecta', 'Producto deformado o dañado', 'Humedad u oxidación', 'Otro'];
  const TRATAMIENTOS = ['Reproceso', 'Reempaque', 'Reclasificar como segunda calidad', 'Reincorporar al proceso', 'Descarte', 'En espera de decisión'];
  const MAXOPT = 40, MAXPX = 1600;
  const st = { tk: null, siteId: null, tipo: null, cat: { terminado: null, semi: null }, sel: null, active: -1, shown: [],
    edit: null, items: new Map(), keep: [], gone: [], news: [] };

  const norm = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const hoy = () => new Date().toLocaleDateString('en-CA');
  const say = (t, cls) => { const m = $('msg'); m.textContent = t || ''; m.className = 'msg' + (cls ? ' ' + cls : ''); };

  async function g(path, opt, raw) {
    const url = path.startsWith('http') ? path : G + path;
    const o = opt || {};
    const r = await fetch(url, { ...o, headers: { Authorization: 'Bearer ' + st.tk, ...(o.body && typeof o.body === 'string' ? { 'Content-Type': 'application/json' } : {}), ...(o.headers || {}) } });
    if (r.status === 401 || r.status === 403) { const e = new Error('Sin permiso (' + r.status + ')'); e.denied = true; throw e; }
    if (!r.ok) { let d = ''; try { d = (await r.json()).error.message; } catch (x) { /* sin detalle */ } throw new Error('Error ' + r.status + (d ? ': ' + d : '')); }
    if (raw) return r;
    return r.status === 204 ? null : r.json();
  }
  async function pages(path) {
    let out = [], url = path;
    while (url) { const j = await g(url); out = out.concat(j.value); url = j['@odata.nextLink'] || null; }
    return out;
  }
  const L = (id) => '/sites/' + st.siteId + '/lists/' + id;
  const fieldsOf = (list, sel) => pages(L(list) + '/items?$expand=fields($select=' + sel + ')&$top=999');
  const photoPath = (item, name) => L(CFG.lists.fotos) + '/drive/root:/' + item + '/' + encodeURIComponent(name);

  // ----- catálogo de productos
  async function loadCatalog(tipo) {
    if (st.cat[tipo]) return st.cat[tipo];
    const m = new Map();
    if (tipo === 'terminado') {
      for (const it of await fieldsOf(CFG.lists.productos, 'CodigoSotfland,field_3,field_4,field_6')) {
        const f = it.fields, cod = (f.CodigoSotfland || '').trim();
        if (cod && !m.has(cod)) m.set(cod, { cod, desc: (f.field_3 || '').trim(), und: (f.field_4 || '').trim(), cat: (f.field_6 || '').trim() });
      }
    } else {
      for (const it of await fieldsOf(CFG.lists.ps, 'CodProducto,Descripci_x00f3_n,CodEntrada,DescripcionEntrada,UndEntrada,FamiliaMaterial')) {
        const f = it.fields;
        if ((f.FamiliaMaterial || '').trim() === 'Semielaborado') {
          const cod = (f.CodEntrada || '').trim();
          if (cod && !m.has(cod)) m.set(cod, { cod, desc: (f.DescripcionEntrada || '').trim(), und: (f.UndEntrada || '').trim(), cat: '' });
        }
        const pc = (f.CodProducto || '').trim();
        if (/^(PS|MP)-/.test(pc) && !m.has(pc)) m.set(pc, { cod: pc, desc: (f.Descripci_x00f3_n || '').trim(), und: '', cat: '' });
      }
    }
    const arr = [...m.values()].map((p) => ({ ...p, key: norm(p.cod + ' ' + p.desc + ' ' + p.cat) })).sort((a, b) => a.desc.localeCompare(b.desc, 'es'));
    st.cat[tipo] = arr;
    return arr;
  }

  // ----- OC o contenedor (Recibidas resumen.csv)
  function parseCsv(t) {
    const rows = []; let r = [], c = '', q = false;
    for (let i = 0; i < t.length; i++) {
      const ch = t[i];
      if (q) { if (ch === '"') { if (t[i + 1] === '"') { c += '"'; i++; } else q = false; } else c += ch; }
      else if (ch === '"') q = true;
      else if (ch === ',') { r.push(c); c = ''; }
      else if (ch === '\n') { r.push(c); rows.push(r); r = []; c = ''; }
      else if (ch !== '\r') c += ch;
    }
    if (c || r.length) { r.push(c); rows.push(r); }
    return rows;
  }
  async function loadOcs() {
    const sel = $('oc'); sel.replaceChildren(new Option('Sin OC', ''));
    try {
      const raw = await (await g('/drives/' + CFG.driveId + '/items/' + CFG.files.recibidas + '/content', null, true)).text();
      const rows = parseCsv(raw.replace(/^﻿/, '')), h = rows[0].map((x) => x.trim());
      const iOc = h.indexOf('OC'), iPr = h.indexOf('PROVEEDOR'), dI = ['Dia en planta', 'Mes en planta', 'ATA', 'ETA'].map((k) => h.indexOf(k));
      const m = new Map();
      rows.slice(1).forEach((r) => { const oc = (r[iOc] || '').trim(); if (!oc) return; const a = dI.map((i) => Number(r[i]) || 0).find(Boolean) || 0; if (!m.has(oc) || a > m.get(oc).a) m.set(oc, { oc, pr: (r[iPr] || '').trim(), a: a || 0 }); });
      [...m.values()].sort((a, b) => b.a - a.a || Number(b.oc) - Number(a.oc)).forEach((o) => sel.appendChild(new Option('OC ' + o.oc + (o.pr ? ' · ' + o.pr : ''), o.oc)));
    } catch (e) { /* el campo OC es opcional */ }
  }

  // ----- selector de producto con búsqueda
  const inp = $('prod'), list = $('opts');
  function render() {
    const arr = st.cat[st.tipo] || [];
    const words = norm(inp.value).split(/\s+/).filter(Boolean);
    const hits = arr.filter((p) => words.every((w) => p.key.includes(w)));
    st.shown = hits.slice(0, MAXOPT);
    list.replaceChildren();
    st.shown.forEach((p, i) => {
      const li = document.createElement('li');
      li.id = 'o' + i; li.setAttribute('role', 'option'); li.setAttribute('aria-selected', String(i === st.active));
      const a = document.createElement('span'); a.textContent = p.desc || p.cod;
      const b = document.createElement('span'); b.className = 'cod'; b.textContent = p.cod + (p.und ? ' · ' + p.und : '') + (p.cat ? ' · ' + p.cat : '');
      li.append(a, b);
      li.addEventListener('mousedown', (e) => { e.preventDefault(); pick(p); });
      list.appendChild(li);
    });
    if (!hits.length) { const d = document.createElement('li'); d.className = 'none'; d.textContent = arr.length ? 'Sin resultados' : 'Cargando productos…'; list.appendChild(d); }
    else if (hits.length > MAXOPT) { const d = document.createElement('li'); d.className = 'more'; d.textContent = 'Hay ' + (hits.length - MAXOPT) + ' más. Sigue escribiendo para filtrar.'; list.appendChild(d); }
    open(true);
  }
  function open(v) { list.hidden = !v; inp.setAttribute('aria-expanded', String(v)); if (!v) inp.removeAttribute('aria-activedescendant'); }
  function mark() {
    [...list.querySelectorAll('[role=option]')].forEach((li, i) => li.setAttribute('aria-selected', String(i === st.active)));
    const cur = list.querySelector('[aria-selected=true]');
    if (cur) { inp.setAttribute('aria-activedescendant', cur.id); cur.scrollIntoView({ block: 'nearest' }); }
  }
  function pick(p) {
    st.sel = p; inp.value = p.cod + ' · ' + p.desc; open(false);
    $('picked').textContent = 'Seleccionado: ' + p.desc + (p.und ? ' (unidad: ' + p.und + ')' : '');
    $('und').textContent = p.und ? '(' + p.und + ')' : '';
    check();
  }
  inp.addEventListener('input', () => { st.sel = null; st.active = -1; $('picked').textContent = ''; $('und').textContent = ''; render(); check(); });
  inp.addEventListener('focus', () => { if (!inp.disabled && !st.sel) render(); });
  inp.addEventListener('blur', () => open(false));
  inp.addEventListener('keydown', (e) => {
    if (list.hidden && (e.key === 'ArrowDown')) { render(); return; }
    if (e.key === 'ArrowDown') { e.preventDefault(); st.active = Math.min(st.active + 1, st.shown.length - 1); mark(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); st.active = Math.max(st.active - 1, 0); mark(); }
    else if (e.key === 'Enter' && !list.hidden) { e.preventDefault(); const p = st.shown[st.active >= 0 ? st.active : (st.shown.length === 1 ? 0 : -1)]; if (p) pick(p); }
    else if (e.key === 'Escape') open(false);
  });

  // ----- tipo de producto
  const PH = { terminado: 'Busca por código o nombre del producto terminado', semi: 'Busca por código o nombre del semielaborado' };
  async function setTipo(t, keepSel) {
    st.tipo = t; if (!keepSel) { st.sel = null; inp.value = ''; $('picked').textContent = ''; $('und').textContent = ''; }
    st.active = -1;
    document.querySelectorAll('#tipo button').forEach((b) => b.setAttribute('aria-checked', String(b.dataset.tipo === t)));
    inp.disabled = false; inp.placeholder = PH[t];
    if (!keepSel) say('');
    if (!st.cat[t]) {
      inp.placeholder = 'Cargando productos…'; inp.disabled = true;
      try { await loadCatalog(t); } catch (e) { say(e.denied ? 'Falta el permiso para leer las listas de productos.' : 'No se pudieron cargar los productos. ' + e.message, 'bad'); return; }
      inp.disabled = false; inp.placeholder = PH[t];
    }
    if (!keepSel) inp.focus();
    check();
  }
  document.querySelectorAll('#tipo button').forEach((b) => b.addEventListener('click', () => setTipo(b.dataset.tipo)));

  // ----- fotos
  function compress(file) {
    return new Promise((res, rej) => {
      const url = URL.createObjectURL(file), im = new Image();
      im.onload = () => {
        const k = Math.min(1, MAXPX / Math.max(im.width, im.height)), c = document.createElement('canvas');
        c.width = Math.round(im.width * k); c.height = Math.round(im.height * k);
        c.getContext('2d').drawImage(im, 0, 0, c.width, c.height);
        c.toBlob((b) => { URL.revokeObjectURL(url); b ? res(b) : rej(new Error('No se pudo procesar la foto')); }, 'image/jpeg', 0.82);
      };
      im.onerror = () => { URL.revokeObjectURL(url); rej(new Error('Archivo de imagen no válido')); };
      im.src = url;
    });
  }
  function thumb(src, onDel, label) {
    const d = document.createElement('div'); d.className = 'th';
    const i = document.createElement('img'); i.src = src; i.alt = label || 'Foto'; i.addEventListener('click', () => window.open(src, '_blank'));
    const b = document.createElement('button'); b.type = 'button'; b.textContent = '×'; b.setAttribute('aria-label', 'Quitar foto'); b.addEventListener('click', onDel);
    d.append(i, b); return d;
  }
  function drawThumbs() {
    const box = $('thumbs'); box.replaceChildren();
    st.keep.forEach((p) => {
      box.appendChild(thumb(p.url || '', () => { st.keep = st.keep.filter((x) => x !== p); st.gone.push(p.n); drawThumbs(); }, p.n));
    });
    st.news.forEach((p) => {
      box.appendChild(thumb(p.url, () => { URL.revokeObjectURL(p.url); st.news = st.news.filter((x) => x !== p); drawThumbs(); }, 'Foto nueva'));
    });
  }
  async function loadPhotoUrls(itemId) {
    for (const p of st.keep) {
      try { const r = await g(photoPath(itemId, p.n) + ':/content', null, true); p.url = URL.createObjectURL(await r.blob()); } catch (e) { /* foto no disponible */ }
      if (st.edit && st.edit.id === itemId) drawThumbs();
    }
  }
  $('addph').addEventListener('click', () => $('fotos').click());
  $('fotos').addEventListener('change', async () => {
    const files = [...$('fotos').files]; $('fotos').value = '';
    for (const f of files) {
      try { const b = await compress(f); st.news.push({ blob: b, url: URL.createObjectURL(b) }); } catch (e) { say(e.message, 'bad'); }
    }
    drawThumbs();
  });

  // ----- formulario
  function opts(sel, arr, ph) { sel.replaceChildren(new Option(ph, ''), ...arr.map((x) => new Option(x, x))); }
  function missing() {
    const f = [];
    if (!st.tipo) f.push('tipo de producto');
    if (!st.sel) f.push('producto');
    if (!$('defecto').value) f.push('tipo de defecto');
    if (!$('trat').value) f.push('tratamiento');
    if (!$('fecha').value) f.push('fecha');
    if (!(Number($('cant').value) > 0)) f.push('cantidad');
    if ($('defecto').value === 'Otro' && !$('obs').value.trim()) f.push('observaciones (obligatorias con "Otro")');
    return f;
  }
  function check() { $('send').disabled = missing().length > 0; $('obs-req').textContent = $('defecto').value === 'Otro' ? '(obligatorio con "Otro")' : '(opcional)'; }
  ['defecto', 'trat', 'fecha', 'cant', 'obs', 'oc'].forEach((id) => { $(id).addEventListener('input', check); $(id).addEventListener('change', check); });

  function clearForm(keepDate) {
    st.sel = null; inp.value = ''; $('picked').textContent = ''; $('und').textContent = '';
    ['defecto', 'trat', 'oc'].forEach((id) => { $(id).value = ''; });
    $('cant').value = ''; $('obs').value = '';
    if (!keepDate) $('fecha').value = hoy();
    st.news.forEach((p) => URL.revokeObjectURL(p.url)); st.keep.forEach((p) => p.url && URL.revokeObjectURL(p.url));
    st.keep = []; st.gone = []; st.news = []; drawThumbs();
  }
  function setMode(item) {
    st.edit = item || null;
    $('send').textContent = item ? 'Guardar cambios' : 'Registrar';
    $('cancel').hidden = !item;
    $('form').classList.toggle('editing', !!item);
    document.querySelectorAll('#rows tr').forEach((tr) => tr.classList.toggle('sel', !!item && tr.dataset.id === item.id));
  }
  $('cancel').addEventListener('click', () => { clearForm(); setMode(null); say(''); check(); });

  async function startEdit(id) {
    const it = st.items.get(id); if (!it) return;
    const f = it.fields;
    clearForm(); setMode({ id });
    say('Editando el registro de ' + (f.Title || '') + '.');
    const tipo = f.TipoProducto === 'Semielaborado' ? 'semi' : 'terminado';
    await setTipo(tipo, true);
    const p = (st.cat[tipo] || []).find((x) => x.cod === f.CodigoProducto) || { cod: f.CodigoProducto || '', desc: f.Title || '', und: f.Unidad || '', cat: '', key: '' };
    pick(p);
    const has = (sel, v) => { if (v && ![...sel.options].some((o) => o.value === v)) sel.appendChild(new Option(v, v)); sel.value = v || ''; };
    has($('defecto'), f.TipoDefecto); has($('trat'), f.Tratamiento); has($('oc'), f.OC);
    $('fecha').value = String(f.FechaDefecto || '').slice(0, 10); $('cant').value = f.Cantidad == null ? '' : f.Cantidad; $('obs').value = f.Observaciones || '';
    let names = []; try { names = JSON.parse(f.Fotos || '[]'); } catch (e) { names = []; }
    st.keep = names.map((n) => ({ n })); drawThumbs(); loadPhotoUrls(id);
    check(); $('form').scrollIntoView({ block: 'start', behavior: 'smooth' });
  }

  async function uploadNew(itemId) {
    const names = [];
    for (const p of st.news) {
      const n = Date.now() + '-' + names.length + '.jpg';
      await g(photoPath(itemId, n) + ':/content', { method: 'PUT', headers: { 'Content-Type': 'image/jpeg' }, body: p.blob });
      names.push(n);
    }
    return names;
  }

  $('form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const m = missing();
    if (m.length) { say('Falta: ' + m.join(', ') + '.', 'bad'); return; }
    const p = st.sel, editing = st.edit;
    const fields = {
      Title: p.desc || p.cod, TipoProducto: st.tipo === 'terminado' ? 'Producto terminado' : 'Semielaborado', CodigoProducto: p.cod,
      TipoDefecto: $('defecto').value, FechaDefecto: $('fecha').value + 'T12:00:00Z', Cantidad: Number($('cant').value),
      Unidad: p.und || '', Tratamiento: $('trat').value, Observaciones: $('obs').value.trim(), OC: $('oc').value,
    };
    $('send').disabled = true; say('Guardando…');
    let id = editing ? editing.id : null, stage = 'registro';
    try {
      if (!id) id = (await g(L(CFG.lists.offspec) + '/items', { method: 'POST', body: JSON.stringify({ fields }) })).id;
      else await g(L(CFG.lists.offspec) + '/items/' + id + '/fields', { method: 'PATCH', body: JSON.stringify(fields) });
      stage = 'fotos';
      const added = await uploadNew(id);
      for (const n of st.gone) { try { await g(photoPath(id, n), { method: 'DELETE' }); } catch (x) { /* ya no existía */ } }
      const names = st.keep.map((k) => k.n).concat(added);
      if (added.length || st.gone.length || !editing) await g(L(CFG.lists.offspec) + '/items/' + id + '/fields', { method: 'PATCH', body: JSON.stringify({ Fotos: JSON.stringify(names) }) });
      say((editing ? 'Cambios guardados: ' : 'Registrado: ') + fields.Title + ', ' + fields.Cantidad + (fields.Unidad ? ' ' + fields.Unidad : '') + (names.length ? ', ' + names.length + ' foto' + (names.length === 1 ? '' : 's') : '') + '.', 'ok');
      clearForm(true); setMode(null); inp.focus(); await loadRecent();
    } catch (er) {
      const base = er.denied ? 'No tienes permiso para guardar en SharePoint.' : 'No se pudo guardar. ' + er.message;
      say(stage === 'fotos' && id ? 'El registro se guardó, pero fallaron las fotos. Ábrelo desde "Últimos registros" y agrégalas otra vez. ' + base : base, 'bad');
      if (stage === 'fotos' && id) { clearForm(true); setMode(null); loadRecent(); }
    }
    check();
  });

  async function loadRecent() {
    try {
      const its = await pages(L(CFG.lists.offspec) + '/items?$expand=fields&$top=999');
      its.sort((a, b) => String(b.createdDateTime).localeCompare(String(a.createdDateTime)));
      st.items = new Map(its.map((x) => [String(x.id), x]));
      const tb = $('rows'); tb.replaceChildren();
      its.slice(0, 20).forEach((it) => {
        const f = it.fields, tr = document.createElement('tr'); tr.dataset.id = String(it.id); tr.tabIndex = 0;
        let nf = 0; try { nf = JSON.parse(f.Fotos || '[]').length; } catch (e) { nf = 0; }
        [String(f.FechaDefecto || '').slice(0, 10), (f.CodigoProducto ? f.CodigoProducto + ' · ' : '') + (f.Title || ''), f.TipoDefecto || '', (f.Cantidad == null ? '' : f.Cantidad) + (f.Unidad ? ' ' + f.Unidad : ''), f.Tratamiento || '', String(nf)].forEach((v, i) => {
          const td = document.createElement('td'); if (i === 3 || i === 5) td.className = 'n'; td.textContent = v; tr.appendChild(td);
        });
        tr.addEventListener('click', () => startEdit(String(it.id)));
        tr.addEventListener('keydown', (e) => { if (e.key === 'Enter') startEdit(String(it.id)); });
        tb.appendChild(tr);
      });
      $('rnote').textContent = its.length + ' registro' + (its.length === 1 ? '' : 's') + ' en total. Se muestran los 20 más recientes.';
      $('recent').hidden = false;
      if (st.edit) setMode(st.edit);
    } catch (e) { /* los últimos registros son opcionales */ }
  }

  async function loadChoices() {
    let d = DEFECTOS, t = TRATAMIENTOS;
    try {
      const cols = (await g(L(CFG.lists.offspec) + '/columns?$select=name,choice')).value;
      const c = (n) => { const x = cols.find((k) => k.name === n); return x && x.choice && x.choice.choices && x.choice.choices.length ? x.choice.choices : null; };
      d = c('TipoDefecto') || d; t = c('Tratamiento') || t;
    } catch (e) { /* se usan las opciones por defecto */ }
    opts($('defecto'), d, 'Selecciona el defecto'); opts($('trat'), t, 'Selecciona el tratamiento');
  }

  async function start() {
    $('fecha').value = hoy(); $('fecha').max = hoy();
    try {
      await window.icAuth.init();
      st.tk = await window.icAuth.token(SC);
    } catch (e) { return needLogin(e.message); }
    if (!st.tk) return needLogin();
    try {
      st.siteId = (await g('/sites/' + CFG.site.host + ':' + CFG.site.path + '?$select=id')).id;
    } catch (e) { return needLogin(e.denied ? 'Falta el permiso para leer y escribir en SharePoint. Pide al administrador que apruebe la aplicación.' : e.message, e.denied); }
    $('gate').style.display = 'none'; $('form').hidden = false;
    await loadChoices(); check(); loadRecent(); loadOcs();
  }
  function needLogin(msg, noButton) {
    $('gate').style.display = 'block'; $('form').hidden = true;
    if (msg) $('gate-msg').textContent = msg;
    $('btn-login').hidden = !!noButton;
  }
  $('btn-login').addEventListener('click', () => window.icAuth.login(SC));
  document.addEventListener('DOMContentLoaded', () => window.icGate(start));
})();
