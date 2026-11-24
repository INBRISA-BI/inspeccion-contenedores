// Registro de producto off spec: lee los productos de las listas de SharePoint y guarda cada registro en CALI_ProductoOffSpec.
(function () {
  const CFG = window.APP_CONFIG;
  const G = 'https://graph.microsoft.com/v1.0';
  const SC = ['Sites.ReadWrite.All'];
  const $ = (id) => document.getElementById(id);
  const DEFECTOS = ['Fuera de especificación dimensional', 'Color o tono diferente', 'Contaminación o suciedad', 'Mal sellado o empaque defectuoso', 'Mal corte o rebaba', 'Peso fuera de rango', 'Etiquetado o codificación incorrecta', 'Producto deformado o dañado', 'Humedad u oxidación', 'Otro'];
  const TRATAMIENTOS = ['Reproceso', 'Reempaque', 'Reclasificar como segunda calidad', 'Reincorporar al proceso', 'Descarte', 'En espera de decisión'];
  const MAXOPT = 40;
  const st = { tk: null, siteId: null, tipo: null, cat: { terminado: null, semi: null }, sel: null, active: -1, shown: [] };

  const norm = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const hoy = () => new Date().toLocaleDateString('en-CA');
  const say = (t, cls) => { const m = $('msg'); m.textContent = t || ''; m.className = 'msg' + (cls ? ' ' + cls : ''); };

  async function g(path, opt) {
    const url = path.startsWith('http') ? path : G + path;
    const r = await fetch(url, { ...(opt || {}), headers: { Authorization: 'Bearer ' + st.tk, 'Content-Type': 'application/json', ...((opt && opt.headers) || {}) } });
    if (r.status === 401 || r.status === 403) { const e = new Error('Sin permiso (' + r.status + ')'); e.denied = true; throw e; }
    if (!r.ok) { let d = ''; try { d = (await r.json()).error.message; } catch (x) { /* sin detalle */ } throw new Error('Error ' + r.status + (d ? ': ' + d : '')); }
    return r.status === 204 ? null : r.json();
  }
  async function pages(path) {
    let out = [], url = path;
    while (url) { const j = await g(url); out = out.concat(j.value); url = j['@odata.nextLink'] || null; }
    return out;
  }
  const fieldsOf = (list, sel) => pages('/sites/' + st.siteId + '/lists/' + list + '/items?$expand=fields($select=' + sel + ')&$top=999');

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

  // ----- selector con búsqueda
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
  async function setTipo(t) {
    st.tipo = t; st.sel = null; st.active = -1;
    document.querySelectorAll('#tipo button').forEach((b) => b.setAttribute('aria-checked', String(b.dataset.tipo === t)));
    inp.value = ''; $('picked').textContent = ''; $('und').textContent = '';
    inp.disabled = false;
    inp.placeholder = t === 'terminado' ? 'Busca por código o nombre del producto terminado' : 'Busca por código o nombre del semielaborado';
    say('');
    if (!st.cat[t]) { inp.placeholder = 'Cargando productos…'; inp.disabled = true; try { await loadCatalog(t); } catch (e) { say(e.denied ? 'Falta el permiso para leer las listas de productos.' : 'No se pudieron cargar los productos. ' + e.message, 'bad'); return; } inp.disabled = false; inp.placeholder = t === 'terminado' ? 'Busca por código o nombre del producto terminado' : 'Busca por código o nombre del semielaborado'; }
    inp.focus(); check();
  }
  document.querySelectorAll('#tipo button').forEach((b) => b.addEventListener('click', () => setTipo(b.dataset.tipo)));

  // ----- formulario
  function opts(sel, arr, ph) { sel.replaceChildren(new Option(ph, ''), ...arr.map((x) => new Option(x, x))); }
  function missing() {
    const f = [];
    if (!st.tipo) f.push('tipo de producto');
    if (!st.sel) f.push('producto');
    if (!$('defecto').value) f.push('tipo de defecto');
    if (!$('trat').value) f.push('tratamiento');
    if (!$('fecha').value) f.push('fecha');
    const c = Number($('cant').value);
    if (!(c > 0)) f.push('cantidad');
    if ($('defecto').value === 'Otro' && !$('obs').value.trim()) f.push('observaciones (obligatorias con "Otro")');
    return f;
  }
  function check() { $('send').disabled = missing().length > 0; $('obs-req').textContent = $('defecto').value === 'Otro' ? '(obligatorio con "Otro")' : '(opcional)'; }
  ['defecto', 'trat', 'fecha', 'cant', 'obs'].forEach((id) => { $(id).addEventListener('input', check); $(id).addEventListener('change', check); });

  $('form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const m = missing();
    if (m.length) { say('Falta: ' + m.join(', ') + '.', 'bad'); return; }
    const p = st.sel;
    const fields = {
      Title: p.desc || p.cod, TipoProducto: st.tipo === 'terminado' ? 'Producto terminado' : 'Semielaborado', CodigoProducto: p.cod,
      TipoDefecto: $('defecto').value, FechaDefecto: $('fecha').value + 'T12:00:00Z', Cantidad: Number($('cant').value),
      Unidad: p.und || '', Tratamiento: $('trat').value, Observaciones: $('obs').value.trim(),
    };
    $('send').disabled = true; say('Guardando…');
    try {
      await g('/sites/' + st.siteId + '/lists/' + CFG.lists.offspec + '/items', { method: 'POST', body: JSON.stringify({ fields }) });
      say('Registrado: ' + fields.Title + ', ' + fields.Cantidad + (fields.Unidad ? ' ' + fields.Unidad : '') + '.', 'ok');
      st.sel = null; inp.value = ''; $('picked').textContent = ''; $('und').textContent = '';
      ['defecto', 'trat'].forEach((id) => { $(id).value = ''; });
      $('cant').value = ''; $('obs').value = '';
      inp.focus(); loadRecent();
    } catch (er) {
      say(er.denied ? 'No tienes permiso para guardar en la lista CALI_ProductoOffSpec.' : 'No se pudo guardar. ' + er.message, 'bad');
    }
    check();
  });

  async function loadRecent() {
    try {
      const its = await pages('/sites/' + st.siteId + '/lists/' + CFG.lists.offspec + '/items?$expand=fields($select=Title,CodigoProducto,TipoDefecto,FechaDefecto,Cantidad,Unidad,Tratamiento)&$top=999');
      its.sort((a, b) => String(b.createdDateTime).localeCompare(String(a.createdDateTime)));
      const tb = $('rows'); tb.replaceChildren();
      its.slice(0, 10).forEach((it) => {
        const f = it.fields, tr = document.createElement('tr');
        [String(f.FechaDefecto || '').slice(0, 10), (f.CodigoProducto ? f.CodigoProducto + ' · ' : '') + (f.Title || ''), f.TipoDefecto || '', (f.Cantidad == null ? '' : f.Cantidad) + (f.Unidad ? ' ' + f.Unidad : ''), f.Tratamiento || ''].forEach((v, i) => {
          const td = document.createElement('td'); if (i === 3) td.className = 'n'; td.textContent = v; tr.appendChild(td);
        });
        tb.appendChild(tr);
      });
      $('rnote').textContent = its.length + ' registro' + (its.length === 1 ? '' : 's') + ' en total.';
      $('recent').hidden = false;
    } catch (e) { /* los últimos registros son opcionales */ }
  }

  async function loadChoices() {
    let d = DEFECTOS, t = TRATAMIENTOS;
    try {
      const cols = (await g('/sites/' + st.siteId + '/lists/' + CFG.lists.offspec + '/columns?$select=name,choice')).value;
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
    await loadChoices(); check(); loadRecent();
  }
  function needLogin(msg, noButton) {
    $('gate').style.display = 'block'; $('form').hidden = true;
    if (msg) $('gate-msg').textContent = msg;
    $('btn-login').hidden = !!noButton;
  }
  $('btn-login').addEventListener('click', () => window.icAuth.login(['Files.Read.All', 'Sites.ReadWrite.All']));
  document.addEventListener('DOMContentLoaded', () => window.icGate(start));
})();
