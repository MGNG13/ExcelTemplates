/* Excel Template Builder v12 — dashboard.js
 * Vista "Actividad": uso en tiempo real por plantilla, usuarios conectados y feed de eventos.
 * Se monta dentro de #dashboard-view (contrato §3.4) y además llena los slots .tpl-usage de Inicio.
 * API pública: ETB.dashboard = { refresh() }.
 * Sin dependencias; todo texto externo pasa por esc() o textContent (CSP/XSS). Sin on*="" inline.
 */
(function () {
  'use strict';

  const REFRESH_MS = 60000;      // recarga del resumen mientras la vista es visible
  const REL_TICK_MS = 30000;     // refresco de "hace X min"
  const AGO_TICK_MS = 5000;      // refresco de "Actualizado hace Xs"
  const HOME_CACHE_MS = 30000;   // caché del resumen para tarjetas de Inicio
  const FEED_MAX = 100;
  const MERGE_MS = 15000;        // ventana para fusionar eventos duplicados (usage 'save' ↔ live 'saved')

  const S = {
    mounted: false, visible: false,
    summary: null, fetchedAt: 0, error: null, inflight: null,
    sort: { key: 'today', dir: -1 }, query: '', showDeleted: false,
    feed: [], feedIds: new Set(),
    presence: null,               // último detail.clients de 'live:presence' (respaldo si ETB.live no responde)
    timers: { refresh: 0, rel: 0, ago: 0 },
    drawerId: null, drawerReq: 0, drawerReturn: null,
  };

  let root = null;
  const reduceMotion = () => !!(window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches);

  // ── Utilidades ──────────────────────────────────────────────────────────────
  const ETB = () => window.ETB || {};
  function esc(s) {
    if (typeof ETB().esc === 'function') return ETB().esc(s);
    return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  /** Ícono del sprite (§3.7). El nombre se valida para no inyectar nada en href. */
  function ico(name, cls) {
    const n = /^i-[a-z0-9-]+$/.test(name) ? name : 'i-dot';
    return `<svg class="ico${cls ? ' ' + cls : ''}" aria-hidden="true" focusable="false"><use href="#${n}"/></svg>`;
  }
  const nf = new Intl.NumberFormat('es-MX');
  const fmt = n => nf.format(Math.round(Number(n) || 0));
  const pct = (a, b) => (b > 0 ? (a / b * 100) : 0);
  const fmtPct = v => (v > 0 && v < 0.1 ? '<0.1' : v.toLocaleString('es-MX', { maximumFractionDigits: v < 10 ? 1 : 0 })) + ' %';
  function relTime(iso) {
    if (!iso) return '—';
    const t = new Date(iso).getTime();
    if (!isFinite(t)) return '—';
    const s = Math.max(0, Math.round((Date.now() - t) / 1000));
    if (s < 45) return 'hace un momento';
    const m = Math.round(s / 60);
    if (m < 60) return `hace ${m} min`;
    const h = Math.round(m / 60);
    if (h < 24) return `hace ${h} h`;
    const d = Math.round(h / 24);
    return d === 1 ? 'ayer' : `hace ${d} días`;
  }
  const absTime = iso => { const d = new Date(iso); return isFinite(d) ? d.toLocaleString('es-MX') : ''; };
  function initials(name) {
    const p = String(name || '?').trim().split(/\s+/).filter(Boolean);
    return ((p[0] || '?')[0] + (p[1] ? p[1][0] : '')).toUpperCase();
  }
  const safeColor = c => (/^#[0-9a-fA-F]{6}$/.test(c || '') ? c : '#217346');
  const cssEsc = s => (window.CSS && CSS.escape ? CSS.escape(s) : String(s).replace(/["\\]/g, '\\$&'));
  const cap = s => (s ? s.charAt(0).toUpperCase() + s.slice(1) : s);

  async function apiGet(path) {
    // Usa ETB.api si existe (añade identidad); si no, fetch directo.
    if (typeof ETB().api === 'function') {
      try { return await ETB().api('GET', path); }
      catch (e) { throw e instanceof Error ? e : Object.assign(new Error(String((e && (e.error || e.message)) || 'Error de red')), { status: e && e.status }); }
    }
    const r = await fetch(path, { headers: { Accept: 'application/json' } });
    if (!r.ok) { const err = new Error(r.status === 404 ? 'No disponible (404)' : 'HTTP ' + r.status); err.status = r.status; throw err; }
    return r.json();
  }
  /** Detecta un 404 aunque ETB.api lo reporte solo en el mensaje. */
  const is404 = err => !!err && (err.status === 404 || /\b404\b|not found/i.test(String(err.message || '')));

  function countUp(el, to) {
    if (!el) return;
    const prev = el.dataset.val;
    el.dataset.val = String(to);
    if (prev === String(to) && el.textContent) return;
    const fx = ETB().fx;
    if (fx && typeof fx.countUp === 'function' && !reduceMotion()) { try { fx.countUp(el, to, 700); return; } catch (_) { /* respaldo abajo */ } }
    el.textContent = fmt(to);
  }
  function flash(el) {
    if (!el || reduceMotion()) return;
    const fx = ETB().fx;
    if (fx && typeof fx.flash === 'function') { try { fx.flash(el); return; } catch (_) { /* respaldo abajo */ } }
    el.classList.remove('db-flash'); void el.offsetWidth; el.classList.add('db-flash');
  }

  // ── Gráficas SVG (sin librerías) ──────────────────────────────────────────────
  /** Sparkline: línea 2px + lavado al 10 %, punto final. vals = number[] viejo→nuevo. */
  function sparkline(vals, w, h, label, hits = true) {
    const v = (Array.isArray(vals) ? vals : []).map(n => Math.max(0, Number(n) || 0));
    if (!v.length) return '';
    const max = Math.max(...v), total = v.reduce((a, b) => a + b, 0), top = Math.max(1, max);
    const step = v.length > 1 ? (w - 6) / (v.length - 1) : 0;
    const pts = v.map((n, i) => [3 + i * step, h - 3 - (n / top) * (h - 7)]);
    const line = pts.map((p, i) => (i ? 'L' : 'M') + p[0].toFixed(1) + ' ' + p[1].toFixed(1)).join('');
    const area = `${line}L${pts[pts.length - 1][0].toFixed(1)} ${h - 1}L3 ${h - 1}Z`;
    const last = pts[pts.length - 1];
    const aria = `${label}: ${fmt(total)} en total, máximo ${fmt(max)} en una hora, ${fmt(v[v.length - 1])} en la hora actual`;
    // <title> por punto (tooltip nativo) sobre círculos invisibles más grandes que la marca.
    const hitEls = hits && v.length <= 48 ? v.map((n, i) => `<circle class="h" cx="${pts[i][0].toFixed(1)}" cy="${pts[i][1].toFixed(1)}" r="${Math.max(3, step / 2).toFixed(1)}"><title>${esc(hourLabel(v.length - 1 - i))}: ${fmt(n)}</title></circle>`).join('') : '';
    return `<svg class="db-spark${total ? '' : ' is-empty'}" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" role="img" aria-label="${esc(aria)}">` +
      `<title>${esc(aria)}</title><path class="a" d="${area}"/><path class="l" d="${line}"/>` +
      (total ? `<circle class="d" cx="${last[0].toFixed(1)}" cy="${last[1].toFixed(1)}" r="2.5"/>` : '') + hitEls + `</svg>`;
  }
  function hourLabel(ago) {
    if (ago === 0) return 'Hora actual';
    const d = new Date(Date.now() - ago * 3600000);
    return `${String(d.getHours()).padStart(2, '0')}:00 (hace ${ago} h)`;
  }

  /** Rectángulo con esquinas superiores redondeadas (extremo de dato) y base cuadrada. */
  function barPath(x, y, w, h, r) {
    r = Math.min(r, w / 2, h);
    return `M${x} ${y + h}V${y + r}Q${x} ${y} ${x + r} ${y}H${x + w - r}Q${x + w} ${y} ${x + w} ${y + r}V${y + h}Z`;
  }
  /**
   * Columnas apiladas (series = [{key,label,cls}], abajo→arriba). rows: [{tick, tip, [key]:n}].
   * Se dibuja al ancho real del contenedor (opts.w) para que el texto no se deforme y los bordes
   * caigan en píxeles enteros (nítido a 1x/2x).
   */
  function barChart(rows, series, opts = {}) {
    const W = Math.max(240, Math.round(opts.w || 540)), H = opts.h || 156, padL = 34, padR = 4, padB = 20, padT = 8;
    const n = rows.length || 1;
    const totals = rows.map(r => series.reduce((a, s) => a + (Number(r[s.key]) || 0), 0));
    const peak = Math.max(0, ...totals), max = niceMax(Math.max(1, peak));
    const plotW = W - padL - padR, plotH = H - padB - padT;
    const slot = plotW / n, bw = Math.max(2, Math.min(24, Math.floor(slot * 0.7)));
    const y = val => Math.round(padT + plotH - (val / max) * plotH);
    let g = '';
    // Rejilla recesiva: 0, mitad, máximo (sin eje secundario).
    [0, max / 2, max].forEach(t => {
      const yy = y(t) + 0.5;
      g += `<line class="${t ? 'grid' : 'base'}" x1="${padL}" x2="${W - padR}" y1="${yy}" y2="${yy}"/>` +
        `<text class="ax" x="${padL - 8}" y="${yy + 3}" text-anchor="end">${fmt(t)}</text>`;
    });
    rows.forEach((r, i) => {
      const x = Math.round(padL + i * slot + (slot - bw) / 2);
      const tip = `${r.tip || r.tick || ''}: ` + series.map(s => `${s.label} ${fmt(r[s.key])}`).join(', ');
      g += `<g class="bar"><title>${esc(tip)}</title><rect class="hit" x="${Math.round(padL + i * slot)}" y="${padT}" width="${Math.max(1, Math.round(slot))}" height="${plotH}"/>`;
      let base = 0;
      const visible = series.filter(s => Number(r[s.key]) > 0);
      visible.forEach((s, si) => {
        const val = Number(r[s.key]);
        const y0 = y(base) - (si > 0 ? 2 : 0);            // 2px de separación entre segmentos
        const y1 = Math.min(y(base + val), y0 - 1);
        const isTop = si === visible.length - 1;
        g += isTop ? `<path class="${s.cls}" d="${barPath(x, y1, bw, y0 - y1, Math.min(4, bw / 2))}"/>`
          : `<rect class="${s.cls}" x="${x}" y="${y1}" width="${bw}" height="${y0 - y1}"/>`;
        base += val;
      });
      g += `</g>`;
      if (opts.tickEvery && r.tick && (i % opts.tickEvery === 0 || i === n - 1) && (i === n - 1 || n - 1 - i >= opts.tickEvery / 2))
        g += `<text class="ax" x="${x + bw / 2}" y="${H - 5}" text-anchor="middle">${esc(r.tick)}</text>`;
    });
    const sum = series.map(s => `${s.label} ${fmt(rows.reduce((a, r) => a + (Number(r[s.key]) || 0), 0))}`).join(', ');
    const aria = `${opts.label || 'Gráfica'}: ${sum}; máximo ${fmt(peak)} en un periodo`;
    return `<svg class="db-bars" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="${esc(aria)}">${g}</svg>`;
  }
  function niceMax(v) {
    if (v <= 4) return 4;
    const p = Math.pow(10, Math.floor(Math.log10(v))), m = v / p;
    return (m <= 2 ? 2 : m <= 5 ? 5 : 10) * p;
  }
  const legend = items => `<div class="db-legend">${items.map(i => `<span><i class="sw ${i.cls}"></i>${esc(i.label)}</span>`).join('')}</div>`;

  // ── Estructura de la vista ────────────────────────────────────────────────────
  function mount() {
    if (S.mounted) return true;
    root = document.getElementById('dashboard-view');
    if (!root) return false;
    S.mounted = true;
    root.classList.add('db-root');
    root.innerHTML = `
      <div class="db-wrap">
        <header class="db-head db-enter">
          <div class="db-title">
            <span class="db-title-ico">${ico('i-activity')}</span>
            <div><h1>Actividad en tiempo real</h1>
            <p class="db-sub"><span class="db-live" data-state="connecting"><i class="db-dot"></i><b class="db-live-txt">Conectando…</b></span>
            <span class="db-sep" aria-hidden="true">·</span><span class="db-ago">Cargando…</span></p></div>
          </div>
          <div class="db-head-actions">
            <button type="button" class="db-btn" data-act="refresh" title="Actualizar ahora">${ico('i-refresh')}<span>Actualizar</span></button>
          </div>
        </header>
        <section class="db-kpis" aria-label="Indicadores"></section>
        <div class="db-grid">
          <section class="db-card db-usage db-enter" style="--i:2">
            <div class="db-card-h"><h2>${ico('i-file-spreadsheet')} Uso por plantilla</h2>
              <div class="db-tools">
                <label class="db-search">${ico('i-search')}<input type="search" placeholder="Buscar plantilla…" aria-label="Buscar plantilla" autocomplete="off"></label>
                <label class="db-check"><input type="checkbox" data-act="deleted"> Mostrar eliminadas</label>
              </div></div>
            <div class="db-table-wrap"></div>
          </section>
          <aside class="db-side">
            <section class="db-card db-online db-enter" style="--i:3"><div class="db-card-h"><h2>${ico('i-users')} Conectados ahora <span class="db-badge db-online-n">0</span></h2></div><ul class="db-people"></ul></section>
            <section class="db-card db-feed-card db-enter" style="--i:4"><div class="db-card-h"><h2>${ico('i-history')} Actividad reciente</h2></div><ol class="db-feed" aria-live="polite" aria-relevant="additions"></ol></section>
          </aside>
        </div>
      </div>
      <div class="db-drawer-ov" hidden></div>
      <aside class="db-drawer" role="dialog" aria-modal="true" aria-label="Detalle de uso" tabindex="-1" hidden></aside>`;

    root.addEventListener('click', onClick);
    root.addEventListener('keydown', onKey);
    const inp = root.querySelector('.db-search input');
    inp.addEventListener('input', () => { S.query = inp.value.trim().toLowerCase(); renderTable(); });
    root.querySelector('[data-act="deleted"]').addEventListener('change', e => { S.showDeleted = e.target.checked; renderTable(); });
    renderKpis(); renderTable(); renderPeople(); renderFeed();
    return true;
  }

  function onKey(e) {
    if (e.key === 'Escape' && S.drawerId) { e.stopPropagation(); closeDrawer(); return; }
    if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('tr[data-id], li[data-open], th[data-sort]')) { e.preventDefault(); e.target.click(); return; }
    // Trampa de foco simple dentro del drawer.
    if (e.key === 'Tab' && S.drawerId) {
      const dr = root.querySelector('.db-drawer');
      const f = [...dr.querySelectorAll('button, [href], input, [tabindex]:not([tabindex="-1"])')].filter(x => !x.disabled && x.offsetParent);
      if (!f.length) return;
      if (e.shiftKey && (document.activeElement === f[0] || document.activeElement === dr)) { e.preventDefault(); f[f.length - 1].focus(); }
      else if (!e.shiftKey && document.activeElement === f[f.length - 1]) { e.preventDefault(); f[0].focus(); }
    }
  }

  function onClick(e) {
    const t = e.target.closest('[data-act], th[data-sort], tr[data-id], li[data-open], .db-drawer-ov');
    if (!t || !root.contains(t)) return;
    if (t.classList.contains('db-drawer-ov')) return closeDrawer();
    if (t.matches('th[data-sort]')) {
      const k = t.dataset.sort;
      S.sort = S.sort.key === k ? { key: k, dir: -S.sort.dir } : { key: k, dir: k === 'name' ? 1 : -1 };
      renderTable();
      root.querySelector(`th[data-sort="${cssEsc(k)}"]`)?.focus();
      return;
    }
    if (t.matches('tr[data-id]')) return openDrawer(t.dataset.id, t);
    if (t.matches('li[data-open]')) return openTpl(t.dataset.open);
    const act = t.dataset.act, id = t.dataset.id;
    if (act === 'refresh' || act === 'retry') return load(true);
    if (act === 'retry-drawer') return S.drawerId && openDrawer(S.drawerId);
    if (act === 'close') return closeDrawer();
    if (act === 'open') { closeDrawer(true); return openTpl(id); }
    if (act === 'automate') {
      if (typeof ETB().openAutomation === 'function') { closeDrawer(true); try { ETB().openAutomation(id); } catch (err) { toast('No se pudo abrir la automatización.', 'error'); } }
      else toast('La automatización no está disponible en este momento.', 'warn');
    }
  }
  function openTpl(id) {
    if (!id) return;
    if (typeof ETB().openTemplate === 'function') { try { ETB().openTemplate(id); } catch (_) { toast('No se pudo abrir la plantilla.', 'error'); } }
  }
  function toast(m, t) { if (typeof ETB().toast === 'function') ETB().toast(m, t); }

  // ── Carga de datos ─────────────────────────────────────────────────────────────
  function load(force) {
    if (S.inflight) return S.inflight;
    if (!force && S.summary && Date.now() - S.fetchedAt < 5000) return Promise.resolve(S.summary);
    root?.querySelector('[data-act="refresh"]')?.classList.add('is-spin');
    S.inflight = apiGet('/api/usage').then(sum => {
      if (!sum || typeof sum !== 'object' || !sum.totals) throw new Error('Respuesta inválida del servidor');
      S.summary = normalizeSummary(sum); S.fetchedAt = Date.now(); S.error = null;
      sum.recent.slice().reverse().forEach(ev => pushFeed(usageItem(ev), true));
      return S.summary;
    }).catch(err => { S.error = err || new Error('Error de red'); return null; }).finally(() => {
      S.inflight = null;
      root?.querySelector('[data-act="refresh"]')?.classList.remove('is-spin');
      if (S.mounted) { renderKpis(); renderTable(); renderFeed(); updateAgo(); }
      paintHome();
    });
    return S.inflight;
  }
  function normalizeSummary(sum) {
    sum.templates = (Array.isArray(sum.templates) ? sum.templates : []).filter(t => t && t.id);
    sum.recent = Array.isArray(sum.recent) ? sum.recent : [];
    sum.totals = sum.totals || {};
    return sum;
  }

  // ── KPIs ───────────────────────────────────────────────────────────────────────
  function kpiDefs() {
    const T = S.summary && S.summary.totals;
    const clients = liveClients();
    const users = groupUsers(clients);
    const openTpls = new Set(clients.filter(c => c.templateId && (c.mode === 'edit' || c.mode === 'view')).map(c => c.templateId)).size;
    const editing = users.filter(u => u.mode === 'edit').length, viewing = users.filter(u => u.mode === 'view').length;
    const tToday = (T && T.today) || {};
    const errRate = pct(tToday.errors || 0, (tToday.renders || 0) + (tToday.errors || 0));
    const hoursSoFar = new Date().getHours() + 1;
    const activeTpls = S.summary ? S.summary.templates.filter(t => !t.deleted).length : 0;
    return [
      { k: 'today', icon: 'i-download', label: 'Generados hoy', v: T ? tToday.renders : null, tone: 'g',
        spark: T ? hourlyTotals() : null, foot: T ? 'Últimas 24 h' : '' },
      { k: 'err', icon: !T || tToday.errors ? 'i-alert-triangle' : 'i-check-circle', label: 'Errores hoy', v: T ? tToday.errors : null, tone: T && tToday.errors ? 'r' : 'n',
        foot: !T ? '' : tToday.errors ? `Tasa de error ${fmtPct(errRate)}` : 'Sin errores hoy' },
      { k: 'online', icon: 'i-users', label: 'Conectados ahora', v: users.length, tone: 'b', live: true,
        foot: users.length ? `${fmt(editing)} editando · ${fmt(viewing)} viendo` : (liveState() === 'offline' ? 'Sin conexión en vivo' : 'Nadie conectado') },
      { k: 'open', icon: 'i-file-spreadsheet', label: 'Plantillas abiertas ahora', v: openTpls, tone: 'b', live: true,
        foot: S.summary && activeTpls ? `de ${fmt(activeTpls)} con actividad registrada` : 'En el editor ahora' },
      { k: 'hour', icon: 'i-trending-up', label: 'Generados última hora', v: T ? T.rendersLastHour : null, tone: 'g',
        foot: T ? `Promedio hoy: ${fmt((tToday.renders || 0) / hoursSoFar)} por hora` : '' },
    ];
  }
  function hourlyTotals() {
    const out = new Array(24).fill(0);
    (S.summary?.templates || []).forEach(t => (t.spark24h || []).forEach((n, i) => { if (i < 24) out[i] += Number(n) || 0; }));
    return out;
  }
  function renderKpis() {
    if (!root) return;
    const box = root.querySelector('.db-kpis');
    const defs = kpiDefs();
    if (!box.children.length) {
      box.innerHTML = defs.map((d, i) => `<div class="db-kpi db-enter tone-${d.tone}" data-k="${d.k}" style="--i:${i + 1}">
        <div class="db-kpi-h"><span class="db-kpi-ico">${ico(d.icon)}</span><span class="db-kpi-l">${esc(d.label)}</span>${d.live ? '<span class="db-kpi-live" title="En vivo"></span>' : ''}</div>
        <strong class="db-kpi-v fx-count"><span class="db-skel-txt"></span></strong>
        <div class="db-kpi-f"><span class="db-kpi-ft"></span><span class="db-kpi-sp"></span></div></div>`).join('');
    }
    defs.forEach(d => {
      const el = box.querySelector(`[data-k="${d.k}"]`);
      ['g', 'r', 'b', 'n'].forEach(t => el.classList.toggle('tone-' + t, t === d.tone));
      const icoUse = el.querySelector('.db-kpi-ico use');
      if (icoUse && icoUse.getAttribute('href') !== '#' + d.icon) icoUse.setAttribute('href', '#' + d.icon);
      el.querySelector('.db-kpi-ft').textContent = d.foot || '';
      const v = el.querySelector('.db-kpi-v');
      if (d.v == null) { if (S.error && !d.live) { v.textContent = '—'; delete v.dataset.val; } return; }
      countUp(v, Number(d.v) || 0);
      if (d.spark) el.querySelector('.db-kpi-sp').innerHTML = sparkline(d.spark, 96, 26, 'Generados por hora, todas las plantillas, últimas 24 h');
    });
  }

  // ── Tabla de uso ───────────────────────────────────────────────────────────────
  const COLS = [
    { k: 'name', l: 'Plantilla' }, { k: 'live', l: 'Ahora', c: 'col-live' }, { k: 'today', l: 'Hoy', n: 1 },
    { k: 'total', l: 'Total', n: 1 }, { k: 'errors', l: 'Errores', n: 1 }, { k: 'avg', l: 'Prom. ms', n: 1, c: 'col-avg' },
    { k: 'last', l: 'Último uso', c: 'col-last' }, { k: 'spark', l: 'Últimas 24 h', ns: 1, c: 'col-spark' },
  ];
  function sortVal(t, k) {
    switch (k) {
      case 'name': return String(t.name || t.id).toLowerCase();
      case 'live': { const b = byTemplate(t.id); return b.viewers.length + b.editors.length * 2; }
      case 'today': return t.today?.renders || 0;
      case 'total': return t.totals?.renders || 0;
      case 'errors': return t.totals?.errors || 0;
      case 'avg': return t.totals?.avgMs || 0;
      case 'last': return Date.parse(lastUse(t)) || 0;
      default: return 0;
    }
  }
  function lastUse(t) {
    const l = t.last || {};
    return [l.renderAt, l.viewAt, l.saveAt].filter(Boolean).sort().pop() || null;
  }
  function renderTable() {
    if (!root) return;
    const wrap = root.querySelector('.db-table-wrap');
    if (!S.summary) {
      wrap.innerHTML = S.error ? errorState() : `<div class="db-skel-table" aria-busy="true" aria-label="Cargando">${'<div class="db-skel-row"><i></i><i></i><i></i><i></i><i></i></div>'.repeat(6)}</div>`;
      return;
    }
    let list = S.summary.templates.slice();
    const hiddenDeleted = list.filter(t => t.deleted).length;
    if (!S.showDeleted) list = list.filter(t => !t.deleted);
    if (S.query) list = list.filter(t => String(t.name || '').toLowerCase().includes(S.query) || String(t.id).toLowerCase().includes(S.query));
    const { key, dir } = S.sort;
    list.sort((a, b) => { const x = sortVal(a, key), y = sortVal(b, key); return ((x < y ? -1 : x > y ? 1 : 0) * dir) || String(a.name || '').localeCompare(String(b.name || '')); });

    if (!list.length) {
      const q = !!S.query;
      wrap.innerHTML = `<div class="db-empty">${ico(q ? 'i-search' : 'i-chart-bar', 'ico-xl')}<b>${q ? 'Sin coincidencias' : 'Aún no hay uso registrado'}</b>
        <span>${q ? 'Prueba con otro nombre.' : 'Genera un Excel desde una plantilla o vía API y aparecerá aquí al instante.'}</span></div>` +
        (!S.showDeleted && hiddenDeleted ? deletedNote(hiddenDeleted) : '');
      return;
    }
    const head = COLS.map(c => {
      const act = c.k === key, cls = [c.n ? 'num' : '', c.c || '', act ? 'is-sorted' : ''].filter(Boolean).join(' ');
      return c.ns ? `<th class="${cls}" scope="col">${c.l}</th>` :
        `<th scope="col" data-sort="${c.k}" class="${cls}" aria-sort="${act ? (dir > 0 ? 'ascending' : 'descending') : 'none'}" tabindex="0">${c.l}<span class="db-sort" aria-hidden="true">${act ? (dir > 0 ? '▲' : '▼') : ''}</span></th>`;
    }).join('');
    wrap.innerHTML = `<table class="db-table"><thead><tr>${head}</tr></thead><tbody>${list.map(rowHtml).join('')}</tbody></table>` +
      (!S.showDeleted && hiddenDeleted ? deletedNote(hiddenDeleted) : '');
  }
  const deletedNote = n => `<p class="db-foot">${ico('i-eye-off')}${fmt(n)} ${n === 1 ? 'plantilla eliminada oculta' : 'plantillas eliminadas ocultas'}.</p>`;
  function liveChips(id) {
    const b = byTemplate(id);
    if (!b.viewers.length && !b.editors.length) return '<span class="db-muted">—</span>';
    const people = [...b.editors.map(p => [p, 1]), ...b.viewers.map(p => [p, 0])];
    const av = people.slice(0, 3).map(([p, ed]) => `<span class="db-av sm${ed ? ' ed' : ''}" style="--c:${safeColor(p.color)}" title="${esc(p.name || 'Invitado')} · ${ed ? 'editando' : 'viendo'}">${esc(initials(p.name))}</span>`).join('');
    const more = people.length > 3 ? `<span class="db-av sm more">+${people.length - 3}</span>` : '';
    const txt = [b.editors.length ? `<span title="Editando">${ico('i-pencil')}${b.editors.length}</span>` : '', b.viewers.length ? `<span title="Viendo">${ico('i-eye')}${b.viewers.length}</span>` : ''].filter(Boolean).join('');
    return `<span class="db-chips"><span class="db-avs">${av}${more}</span><span class="db-chip-n">${txt}</span></span>`;
  }
  function rowHtml(t) {
    const lu = lastUse(t);
    const err = t.totals?.errors || 0;
    const name = t.name || t.id;
    return `<tr data-id="${esc(t.id)}" tabindex="0" class="${t.deleted ? 'is-deleted' : ''}" aria-label="${esc(`Ver detalle de ${name}`)}">
      <td class="nm"><span class="db-nm">${ico('i-file-spreadsheet')}<span title="${esc(name)}">${esc(name)}</span>${t.deleted ? '<span class="db-tag">Eliminada</span>' : ''}</span></td>
      <td class="lv col-live">${liveChips(t.id)}</td>
      <td class="num strong">${fmt(t.today?.renders)}</td>
      <td class="num">${fmt(t.totals?.renders)}</td>
      <td class="num">${err ? `<span class="db-err" title="${esc(`${fmt(err)} errores en total · ${fmt(t.today?.errors)} hoy`)}">${ico('i-alert-circle')}${fmt(err)}</span>` : '<span class="db-muted">0</span>'}</td>
      <td class="num col-avg">${t.totals?.avgMs ? fmt(t.totals.avgMs) : '<span class="db-muted">—</span>'}</td>
      <td class="rel col-last" data-ts="${esc(lu || '')}" title="${esc(lu ? absTime(lu) : 'Sin uso')}">${relTime(lu)}</td>
      <td class="sp col-spark">${sparkline(t.spark24h, 92, 24, `${name}: generados por hora, últimas 24 h`)}</td></tr>`;
  }
  function errorState() {
    const nf404 = is404(S.error);
    return `<div class="db-empty is-error" role="alert">${ico(nf404 ? 'i-server' : 'i-wifi-off', 'ico-xl')}<b>${nf404 ? 'Las estadísticas de uso no están disponibles' : 'No pudimos cargar la actividad'}</b>
      <span>${nf404 ? 'El servidor todavía no ofrece el registro de uso. Inténtalo más tarde.' : 'Revisa tu conexión e inténtalo de nuevo.'}</span>
      <button type="button" class="db-btn primary" data-act="retry">${ico('i-refresh')}<span>Reintentar</span></button></div>`;
  }
  function refreshRelTimes() {
    if (!root) return;
    root.querySelectorAll('[data-ts]').forEach(el => { if (el.dataset.ts) el.textContent = relTime(el.dataset.ts); });
  }

  // ── Presencia ─────────────────────────────────────────────────────────────────
  function liveClients() {
    try { const c = ETB().live && typeof ETB().live.clients === 'function' && ETB().live.clients(); if (Array.isArray(c)) return c; } catch (_) { /* respaldo */ }
    return Array.isArray(S.presence) ? S.presence : [];
  }
  function byTemplate(id) {
    try {
      const b = ETB().live && typeof ETB().live.byTemplate === 'function' && ETB().live.byTemplate(id);
      if (b) return { viewers: b.viewers || [], editors: b.editors || [] };
    } catch (_) { /* respaldo */ }
    const on = liveClients().filter(c => c.templateId === id);
    return { viewers: on.filter(c => c.mode === 'view'), editors: on.filter(c => c.mode === 'edit') };
  }
  const RANK = { edit: 3, view: 2, dashboard: 1, home: 0 };
  /** Agrupa pestañas por usuario; la pestaña más activa representa al usuario. */
  function groupUsers(clients) {
    const users = new Map();
    const score = c => (c.idle ? 0 : 10) + (RANK[c.mode] || 0);
    clients.forEach(c => {
      if (!c) return;
      const k = c.userId || c.clientId; const cur = users.get(k);
      if (!cur) users.set(k, { ...c, tabs: 1 });
      else users.set(k, score(c) > score(cur) ? { ...c, tabs: cur.tabs + 1 } : { ...cur, tabs: cur.tabs + 1 });
    });
    return [...users.values()];
  }
  function activity(c) {
    const tn = c.templateName || 'una plantilla';
    const where = c.sheet ? ` · ${c.sheet}${c.cell ? '!' + c.cell : ''}` : '';
    switch (c.mode) {
      case 'edit': return { icon: 'i-pencil', txt: `Editando ${tn}${where}`, cls: 'ed' };
      case 'view': return { icon: 'i-eye', txt: `Viendo ${tn}${where}`, cls: 'vw' };
      case 'dashboard': return { icon: 'i-activity', txt: 'En Actividad', cls: 'db' };
      default: return { icon: 'i-home', txt: 'En inicio', cls: 'hm' };
    }
  }
  function selfUserId() {
    try { return (ETB().live?.self?.() || {}).userId || ETB().identity?.userId || null; } catch (_) { return null; }
  }
  function renderPeople() {
    if (!root) return;
    const ul = root.querySelector('.db-people');
    const selfId = selfUserId();
    const list = groupUsers(liveClients()).sort((a, b) => (!!a.idle - !!b.idle) || (RANK[b.mode] || 0) - (RANK[a.mode] || 0) || String(a.name || '').localeCompare(String(b.name || '')));
    root.querySelector('.db-online-n').textContent = fmt(list.length);
    const st = liveState();
    if (!list.length) {
      const off = st === 'offline';
      ul.innerHTML = `<li class="db-empty sm">${ico(off ? 'i-wifi-off' : 'i-user', 'ico-xl')}<b>${off ? 'Sin conexión en vivo' : 'Nadie conectado'}</b><span>${off ? 'La presencia se mostrará al reconectar.' : 'Cuando alguien abra la app aparecerá aquí.'}</span></li>`;
      return;
    }
    ul.innerHTML = list.map(c => {
      const a = activity(c);
      const canOpen = !!c.templateId && (c.mode === 'edit' || c.mode === 'view');
      const me = !!selfId && c.userId === selfId;
      const status = c.idle ? 'Inactivo' : a.cls === 'ed' ? 'Editando' : a.cls === 'vw' ? 'Viendo' : 'Conectado';
      return `<li class="db-person${c.idle ? ' is-idle' : ''}${canOpen ? ' is-link' : ''}"${canOpen ? ` data-open="${esc(c.templateId)}" tabindex="0" title="${esc(`Abrir ${c.templateName || 'plantilla'}`)}"` : ''}>
        <span class="db-av" style="--c:${safeColor(c.color)}" aria-hidden="true">${esc(initials(c.name))}<i class="st ${c.idle ? 'idle' : a.cls}"></i></span>
        <span class="db-pb"><b><span class="db-pn">${esc(c.name || 'Invitado')}</span>${me ? ' <em>(tú)</em>' : ''}${c.tabs > 1 ? ` <em>· ${c.tabs} pestañas</em>` : ''}</b>
        <span class="db-act ${a.cls}">${ico(a.icon)}<span title="${esc(a.txt)}">${esc(a.txt)}</span></span></span>
        <span class="db-pr"><span class="sr-only">${status}</span>${c.idle ? '<span class="db-idle" aria-hidden="true">Inactivo</span>' : ''}${canOpen ? ico('i-arrow-right') : ''}</span></li>`;
    }).join('');
  }
  function liveState() {
    try { const s = ETB().live && typeof ETB().live.status === 'function' && ETB().live.status(); if (s) return s; } catch (_) { /* respaldo */ }
    return ETB().live ? 'connecting' : 'offline';
  }
  function renderStatus(state) {
    if (!root) return;
    const el = root.querySelector('.db-live');
    const st = ['online', 'connecting', 'offline'].includes(state) ? state : liveState();
    el.dataset.state = st;
    root.dataset.live = st; // atenúa los indicadores "en vivo" de los KPIs si no hay conexión
    el.querySelector('.db-live-txt').textContent = st === 'online' ? 'En vivo' : st === 'connecting' ? 'Conectando…' : 'Sin conexión en vivo';
  }
  function updateAgo() {
    if (!root) return;
    const el = root.querySelector('.db-ago');
    if (!S.fetchedAt) { el.textContent = S.error ? 'Sin datos' : 'Cargando…'; return; }
    const s = Math.round((Date.now() - S.fetchedAt) / 1000);
    el.textContent = s < 5 ? 'Actualizado ahora' : s < 60 ? `Actualizado hace ${s} s` : `Actualizado ${relTime(new Date(S.fetchedAt).toISOString())}`;
    el.title = 'Última carga completa: ' + absTime(new Date(S.fetchedAt).toISOString());
  }

  // ── Feed ───────────────────────────────────────────────────────────────────────
  // tipo → [ícono, verbo (minúsculas), tono, familia para fusionar duplicados]
  const EV = {
    render: ['i-download', 'generó', 'g', 'render'], batch: ['i-package', 'generó un lote de', 'g', 'batch'], view: ['i-eye', 'abrió', 'b', 'view'],
    save: ['i-save', 'guardó', 'n', 'save'], create: ['i-plus', 'creó', 'g', 'create'], import: ['i-upload', 'importó', 'n', 'create'],
    duplicate: ['i-copy', 'duplicó', 'n', 'create'], delete: ['i-trash', 'eliminó', 'r', 'delete'],
    saved: ['i-save', 'guardó', 'n', 'save'], created: ['i-plus', 'creó', 'g', 'create'], deleted: ['i-trash', 'eliminó', 'r', 'delete'],
  };
  const SRC = { ui: 'desde la app', api: 'vía API', link: 'por enlace' };
  function usageItem(ev) {
    if (!ev || typeof ev !== 'object') return null;
    const def = EV[ev.type] || ['i-dot', String(ev.type || 'evento'), 'n', ev.type];
    const fail = ev.ok === false;
    return {
      id: 'u:' + (ev.id || `${ev.type}-${ev.templateId}-${ev.ts}`), ts: ev.ts || new Date().toISOString(),
      icon: fail ? 'i-alert-triangle' : def[0], tone: fail ? 'r' : def[2], fam: def[3],
      who: ev.client || (ev.source === 'api' ? 'API' : ev.source === 'link' ? 'Enlace' : ''),
      verb: fail ? (ev.type === 'batch' ? 'falló un lote de' : 'falló al generar') : def[1], what: ev.templateName || ev.templateId || '',
      extra: [ev.type === 'batch' && ev.count ? `${fmt(ev.count)} archivos` : '', SRC[ev.source] || '', ev.ms ? `${fmt(ev.ms)} ms` : '', fail && ev.error ? String(ev.error).slice(0, 160) : ''].filter(Boolean).join(' · '),
      tid: ev.templateId,
    };
  }
  /** Inserta en el feed (orden por fecha desc, máx. FEED_MAX). Fusiona el mismo hecho reportado por dos canales. */
  function pushFeed(item, initial) {
    if (!item || S.feedIds.has(item.id)) return false;
    const t = Date.parse(item.ts) || Date.now();
    if (item.fam === 'save' || item.fam === 'create' || item.fam === 'delete') {
      const twin = S.feed.find(f => f.fam === item.fam && f.tid === item.tid && Math.abs((Date.parse(f.ts) || 0) - t) < MERGE_MS);
      if (twin) {
        S.feedIds.add(item.id);
        if (!twin.who && item.who) twin.who = item.who;
        if (!twin.extra && item.extra) twin.extra = item.extra;
        return false;
      }
    }
    S.feedIds.add(item.id);
    item.fresh = !initial;
    S.feed.unshift(item); // al frente: con sort estable, el más reciente gana en empates de ts
    S.feed.sort((a, b) => (Date.parse(b.ts) || 0) - (Date.parse(a.ts) || 0));
    while (S.feed.length > FEED_MAX) S.feed.pop();
    if (S.feedIds.size > FEED_MAX * 4) S.feedIds = new Set(S.feed.map(f => f.id)); // acota memoria
    return true;
  }
  function feedHtml(f) {
    const sentence = f.who ? `<b>${esc(f.who)}</b> ${esc(f.verb)}` : esc(cap(f.verb));
    return `<li class="db-ev tone-${f.tone}${f.fresh ? ' is-new' : ''}">
      <span class="db-ev-ico" aria-hidden="true">${ico(f.icon)}</span>
      <span class="db-ev-b"><span class="db-ev-t">${sentence} <b class="db-ev-w">${esc(f.what)}</b></span>
      ${f.extra ? `<span class="db-ev-x">${esc(f.extra)}</span>` : ''}</span>
      <time datetime="${esc(f.ts)}" data-ts="${esc(f.ts)}" title="${esc(absTime(f.ts))}">${relTime(f.ts)}</time></li>`;
  }
  function renderFeed() {
    if (!root) return;
    const ol = root.querySelector('.db-feed');
    if (!S.feed.length) {
      ol.innerHTML = !S.summary && !S.error ? '<li class="db-skel-row" aria-hidden="true"><i></i><i></i></li>'.repeat(5)
        : `<li class="db-empty sm">${ico('i-history', 'ico-xl')}<b>Sin actividad reciente</b><span>Los eventos aparecerán aquí en vivo.</span></li>`;
      return;
    }
    ol.innerHTML = S.feed.map(feedHtml).join('');
    S.feed.forEach(f => { f.fresh = false; });
  }

  // ── Eventos en vivo ─────────────────────────────────────────────────────────────
  function onUsage(d) {
    if (!d || typeof d !== 'object') return;
    const tpl = d.template && d.template.id ? d.template : null;
    if (S.summary) {
      if (d.totals && typeof d.totals === 'object') S.summary.totals = d.totals;
      if (tpl) {
        const arr = S.summary.templates;
        const i = arr.findIndex(t => t.id === tpl.id);
        if (i >= 0) arr[i] = tpl; else arr.push(tpl);
      }
    }
    const added = pushFeed(usageItem(d.event));
    if (S.mounted) {
      renderKpis();
      if (tpl && S.summary) {
        const tr = root.querySelector(`tr[data-id="${cssEsc(tpl.id)}"]`);
        if (tr) {
          const tmp = document.createElement('tbody'); tmp.innerHTML = rowHtml(tpl);
          const nr = tmp.firstElementChild; tr.replaceWith(nr); flash(nr);
        } else renderTable();
      }
      if (added) renderFeed();
      if (S.visible && S.drawerId && tpl && tpl.id === S.drawerId) loadDrawer(S.drawerId, true);
    }
    if (tpl) paintHome([tpl.id], true);
  }
  function onTemplate(d) {
    if (!d || !d.type || !EV[d.type]) return;
    const known = S.summary?.templates.find(t => t.id === d.templateId);
    const def = EV[d.type];
    pushFeed({
      id: `t:${d.type}-${d.templateId}-${d.revision || ''}-${d.updatedAt || ''}`, ts: d.updatedAt || new Date().toISOString(),
      icon: def[0], tone: def[2], fam: def[3], who: (d.by && d.by.name) || '', verb: def[1],
      what: d.name || (known && known.name) || d.templateId || '', extra: d.type === 'saved' && d.revision ? `revisión ${d.revision}` : '', tid: d.templateId,
    });
    let tableDirty = false;
    if (known) {
      if (d.type === 'deleted' && !known.deleted) { known.deleted = true; tableDirty = true; }
      if (d.type === 'saved' && d.name && d.name !== known.name) { known.name = d.name; tableDirty = true; }
    }
    // Siempre re-pintar el feed: aunque se haya fusionado con un evento previo, pudo ganar autor.
    if (S.mounted) { renderFeed(); if (tableDirty) renderTable(); }
  }

  // ── Drawer de detalle ───────────────────────────────────────────────────────────
  function openDrawer(id, from) {
    if (!root || !id) return;
    if (from) S.drawerReturn = from;
    S.drawerId = id;
    const dr = root.querySelector('.db-drawer'), ov = root.querySelector('.db-drawer-ov');
    const wasOpen = !dr.hidden;
    dr.hidden = false; ov.hidden = false;
    if (!wasOpen) requestAnimationFrame(() => { dr.classList.add('open'); ov.classList.add('open'); });
    const t = S.summary?.templates.find(x => x.id === id);
    dr.innerHTML = drawerHead(t || { id, name: id }) + `<div class="db-dr-body" aria-busy="true"><div class="db-skel-block sm"></div><div class="db-skel-block"></div><div class="db-skel-block"></div></div>`;
    dr.focus({ preventScroll: true });
    loadDrawer(id);
  }
  function drawerHead(t) {
    return `<header class="db-dr-h"><div class="db-dr-ht"><span class="db-dr-k">${ico('i-chart-bar')} Uso de plantilla</span><h2>${esc(t.name || t.id)}</h2>
      ${t.deleted ? '<span class="db-tag">Eliminada</span>' : ''}</div>
      <button type="button" class="db-icon-btn" data-act="close" aria-label="Cerrar detalle" title="Cerrar (Esc)">${ico('i-x')}</button></header>`;
  }
  function loadDrawer(id, silent) {
    const req = ++S.drawerReq;
    return apiGet('/api/usage/' + encodeURIComponent(id)).then(d => {
      if (req !== S.drawerReq || S.drawerId !== id) return;
      if (!d || typeof d !== 'object') throw new Error('Respuesta inválida del servidor');
      renderDrawer(d, id);
    }).catch(err => {
      if (req !== S.drawerReq || S.drawerId !== id || silent) return;
      const body = root.querySelector('.db-dr-body');
      if (!body) return;
      body.removeAttribute('aria-busy');
      body.innerHTML = `<div class="db-empty is-error" role="alert">${ico('i-alert-circle', 'ico-xl')}<b>No se pudo cargar el detalle</b>
        <span>${is404(err) ? 'El detalle de uso no está disponible en este servidor.' : esc((err && err.message) || 'Error de red')}</span>
        <button type="button" class="db-btn primary" data-act="retry-drawer">${ico('i-refresh')}<span>Reintentar</span></button></div>`;
    });
  }
  function renderDrawer(d, id) {
    const dr = root.querySelector('.db-drawer');
    const scroll = dr.querySelector('.db-dr-body')?.scrollTop || 0;
    d.id = d.id || id;
    const h48 = (Array.isArray(d.hourly48) ? d.hourly48 : []).map(r => {
      const hh = String(r.hour || '').slice(11, 13), dd = String(r.hour || '').slice(0, 10);
      const dt = new Date(dd + 'T00:00:00');
      return { ...r, tick: hh ? `${hh}:00` : '', tip: `${isFinite(dt) ? dt.toLocaleDateString('es-MX', { weekday: 'short', day: 'numeric', month: 'short' }) : dd} ${hh}:00` };
    });
    const d30 = (Array.isArray(d.daily30) ? d.daily30 : []).map(r => {
      const dt = new Date(String(r.day) + 'T00:00:00');
      return { ...r, tick: isFinite(dt) ? dt.toLocaleDateString('es-MX', { day: 'numeric', month: 'short' }) : '', tip: isFinite(dt) ? dt.toLocaleDateString('es-MX', { weekday: 'long', day: 'numeric', month: 'long' }) : String(r.day || '') };
    });
    const src = d.sources || {}, srcTot = (src.ui || 0) + (src.api || 0) + (src.link || 0);
    const srcRows = [['ui', 'App (UI)', 'i-grid'], ['api', 'API', 'i-code'], ['link', 'Enlace', 'i-link']];
    const tc = Array.isArray(d.topClients) ? d.topClients : [], tcMax = Math.max(1, ...tc.map(c => c.renders || 0));
    const T = d.totals || {};
    const errRate = pct(T.errors || 0, (T.renders || 0) + (T.errors || 0));
    const mini = (l, v, ic, cls) => `<div class="db-mini${cls ? ' ' + cls : ''}"><span>${ico(ic)}${l}</span><b>${v}</b></div>`;
    const series = [{ key: 'renders', label: 'Generados', cls: 'ok' }, { key: 'errors', label: 'Errores', cls: 'bad' }];
    const lg = legend([{ cls: 'ok', label: 'Generados' }, { cls: 'bad', label: 'Errores' }]);

    // Ancho real disponible para las gráficas (sección = cuerpo − padding).
    const bodyW = (dr.querySelector('.db-dr-body')?.clientWidth || dr.clientWidth || 580) - 40 - 30;
    dr.innerHTML = drawerHead(d) + `<div class="db-dr-body">
      ${d.deleted ? '' : `<div class="db-dr-actions">
        <button type="button" class="db-btn primary" data-act="open" data-id="${esc(d.id)}">${ico('i-external')}<span>Abrir plantilla</span></button>
        <button type="button" class="db-btn" data-act="automate" data-id="${esc(d.id)}">${ico('i-zap')}<span>Automatizar</span></button></div>`}
      <div class="db-minis">
        ${mini('Hoy', fmt(d.today?.renders), 'i-download')}${mini('Total generados', fmt(T.renders), 'i-hash')}
        ${mini('Errores', `${fmt(T.errors)}<small>${T.errors ? fmtPct(errRate) : ''}</small>`, 'i-alert-triangle', T.errors ? 'is-bad' : '')}
        ${mini('Tiempo promedio', T.avgMs ? fmt(T.avgMs) + '<small>ms</small>' : '—', 'i-gauge')}
        ${mini('Vistas', fmt(T.views), 'i-eye')}${mini('Guardados', fmt(T.saves), 'i-save')}
      </div>
      <section class="db-sec"><div class="db-sec-h"><h3>Últimas 48 horas</h3>${lg}</div>
        ${h48.some(r => r.renders || r.errors) ? barChart(h48, series, { w: bodyW, label: 'Generados y errores por hora, últimas 48 h', tickEvery: 12 }) : emptyChart('Sin generados en las últimas 48 horas.')}</section>
      <section class="db-sec"><div class="db-sec-h"><h3>Últimos 30 días</h3>${lg}</div>
        ${d30.some(r => r.renders || r.errors) ? barChart(d30, series, { w: bodyW, label: 'Generados y errores por día, últimos 30 días', tickEvery: 7 }) : emptyChart('Sin generados en los últimos 30 días.')}</section>
      <div class="db-two">
        <section class="db-sec"><div class="db-sec-h"><h3>Origen</h3></div>
          ${srcTot ? `<div class="db-stack" role="img" aria-label="${esc('Origen de los generados: ' + srcRows.map(([k, l]) => `${l} ${fmt(src[k])}`).join(', '))}">${srcRows.filter(([k]) => src[k] > 0).map(([k, l]) => `<i class="s-${k}" style="flex:${Number(src[k]) || 0}" title="${esc(`${l}: ${fmt(src[k])} (${Math.round(src[k] / srcTot * 100)} %)`)}"></i>`).join('')}</div>
          <ul class="db-src">${srcRows.map(([k, l, ic]) => `<li><i class="sw s-${k}"></i>${ico(ic)}<span>${l}</span><b>${fmt(src[k])}</b><em>${srcTot ? Math.round((src[k] || 0) / srcTot * 100) : 0} %</em></li>`).join('')}</ul>`
          : '<p class="db-muted db-note">Sin generados aún.</p>'}</section>
        <section class="db-sec"><div class="db-sec-h"><h3>Principales clientes</h3></div>
          ${tc.length ? `<ul class="db-top" role="list">${tc.slice(0, 6).map(c => `<li title="${esc(`${c.client || 'Anónimo'}: ${fmt(c.renders)} generados`)}"><span class="nm">${esc(c.client || 'Anónimo')}</span><span class="bar" aria-hidden="true"><i style="width:${Math.max(3, (c.renders || 0) / tcMax * 100).toFixed(1)}%"></i></span><b>${fmt(c.renders)}</b></li>`).join('')}</ul>` : '<p class="db-muted db-note">Sin clientes registrados.</p>'}</section>
      </div>
      <section class="db-sec"><div class="db-sec-h"><h3>Eventos recientes</h3></div>
        ${(Array.isArray(d.recent) && d.recent.length) ? `<ol class="db-feed compact">${d.recent.slice(0, 25).map(e => usageItem(e)).filter(Boolean).map(feedHtml).join('')}</ol>` : '<p class="db-muted db-note">Sin eventos.</p>'}</section>
    </div>`;
    const body = dr.querySelector('.db-dr-body');
    if (body && scroll) body.scrollTop = scroll;
  }
  const emptyChart = msg => `<div class="db-chart-empty">${ico('i-chart-bar')}<span>${esc(msg)}</span></div>`;
  function closeDrawer(noFocus) {
    if (!root || !S.drawerId) return;
    S.drawerId = null; S.drawerReq++;
    const dr = root.querySelector('.db-drawer'), ov = root.querySelector('.db-drawer-ov');
    dr.classList.remove('open'); ov.classList.remove('open');
    setTimeout(() => { if (!S.drawerId) { dr.hidden = true; ov.hidden = true; } }, reduceMotion() ? 0 : 280);
    const back = S.drawerReturn; S.drawerReturn = null;
    if (!noFocus && back && back.isConnected) back.focus({ preventScroll: true });
  }

  // ── Visibilidad / ciclo de vida ──────────────────────────────────────────────────
  function stopTimers() { Object.keys(S.timers).forEach(k => { clearInterval(S.timers[k]); S.timers[k] = 0; }); }
  function show() {
    if (!mount()) return;
    const first = !S.visible;
    S.visible = true;
    renderStatus(); renderPeople(); renderKpis(); renderTable(); renderFeed(); updateAgo();
    load(!S.summary || Date.now() - S.fetchedAt > 15000);
    if (first) {
      stopTimers();
      S.timers.refresh = setInterval(() => { if (!document.hidden) load(true); }, REFRESH_MS);
      S.timers.rel = setInterval(() => { if (!document.hidden) refreshRelTimes(); }, REL_TICK_MS);
      S.timers.ago = setInterval(updateAgo, AGO_TICK_MS);
    }
  }
  function hide() {
    S.visible = false;
    stopTimers();
    closeDrawer(true);
  }

  // ── Tarjetas de Inicio (.tpl-usage) ─────────────────────────────────────────────────
  function paintHome(onlyIds, animate) {
    if (!S.summary) return;
    const slots = document.querySelectorAll('.tpl-usage[data-usage-for]');
    if (!slots.length) return;
    const map = new Map(S.summary.templates.map(t => [t.id, t]));
    slots.forEach(el => {
      const id = el.getAttribute('data-usage-for');
      if (onlyIds && !onlyIds.includes(id)) return;
      const t = map.get(id);
      const total = t?.totals?.renders || 0, today = t?.today?.renders || 0;
      const html = !t || (!total && !today)
        ? `<span class="db-hu is-empty">${ico('i-chart-bar')}<span>Sin uso aún</span></span>`
        : `<span class="db-hu" title="${esc(`${fmt(today)} generados hoy · ${fmt(total)} en total`)}">${ico('i-trending-up')}<span class="db-hu-t"><b>${fmt(today)}</b> hoy · ${fmt(total)} total</span>${sparkline(t.spark24h, 52, 16, `${t.name || id}: generados por hora, últimas 24 h`, false)}</span>`;
      if (el.dataset.dbHtml !== html) {
        el.innerHTML = html; el.dataset.dbHtml = html;
        if (animate) flash(el);
      }
    });
  }
  function onHomePainted() {
    // Las tarjetas se re-crean al repintar: forzar repintado de los slots nuevos.
    document.querySelectorAll('.tpl-usage[data-usage-for]').forEach(el => { if (!el.firstChild) delete el.dataset.dbHtml; });
    if (S.summary && Date.now() - S.fetchedAt < HOME_CACHE_MS) return paintHome();
    if (S.summary) paintHome();      // pinta lo que hay mientras llega lo nuevo
    load(true);                      // paintHome se ejecuta al terminar
  }

  // ── Init ──────────────────────────────────────────────────────────────────────────
  let inited = false;
  function init() {
    const E = ETB();
    if (inited || !E || typeof E.on !== 'function') return;
    inited = true;
    E.dashboard = { refresh: () => load(true) };
    E.on('view', d => (d && d.view === 'dashboard' ? show() : (S.visible && hide())));
    E.on('live:status', d => { if (S.mounted) { renderStatus(d && d.state); renderPeople(); renderKpis(); } });
    E.on('live:presence', d => {
      if (d && Array.isArray(d.clients)) S.presence = d.clients;
      if (S.mounted) { renderPeople(); renderKpis(); refreshLiveCells(); }
    });
    E.on('live:usage', onUsage);
    E.on('live:template', onTemplate);
    E.on('home:painted', onHomePainted);
    document.addEventListener('visibilitychange', () => { if (!document.hidden && S.visible && Date.now() - S.fetchedAt > REFRESH_MS) load(true); });
    // Si la vista ya está activa al cargar el módulo.
    try { if (typeof E.currentView === 'function' && E.currentView() === 'dashboard') show(); } catch (_) { /* noop */ }
    if (document.querySelector('.tpl-usage[data-usage-for]')) onHomePainted();
  }
  function refreshLiveCells() {
    if (!root || !S.summary) return;
    root.querySelectorAll('tr[data-id]').forEach(tr => { const td = tr.querySelector('td.lv'); if (td) td.innerHTML = liveChips(tr.dataset.id); });
  }

  function boot() {
    if (window.ETB && typeof window.ETB.whenReady === 'function') window.ETB.whenReady(init); else init();
  }
  if (window.ETB && typeof window.ETB.whenReady === 'function') boot();
  else if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
