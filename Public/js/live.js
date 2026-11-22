/* ══════════════════════════════════════════════════════════════════════════
 *  live.js — Presencia en tiempo real (Excel Template Builder v12)
 *  Contrato §3.1, §3.3, §3.4, §3.5.
 *  - Identidad del cliente (síncrona, antes de cualquier api()).
 *  - Conexión SSE /api/live/stream + POST /presence + beacon /leave.
 *  - Avatares, cursores remotos, insignias en tarjetas, banner de cambios remotos.
 *  Todo es tolerante a fallos: si falta un punto de montaje o un método de ETB, no hace nada.
 * ══════════════════════════════════════════════════════════════════════════ */
(function () {
    'use strict';

    const ETB = window.ETB = window.ETB || {};
    // 8 colores armónicos (paleta Office) con buen contraste para texto blanco.
    const PALETTE = ['#217346', '#0078d4', '#8764b8', '#c239b3', '#d13438', '#ca5010', '#986f0b', '#038387'];
    const EDIT_WINDOW_MS = 15000;      // "editando" si hubo actividad en los últimos 15 s
    const OFFLINE_AFTER_MS = 10000;    // sin reconectar en 10 s → offline
    const STALE_AFTER_MS = 70000;      // sin ningún evento (ping cada 25 s) → conexión muerta
    const KEEPALIVE_MS = 30000;
    const THROTTLE_MS = 250;
    const NEAR_LIVE_IDLE_MS = 3000;
    const HOME_REFRESH_MS = 800;
    const JOIN_TOAST_GAP_MS = 4000;
    const ID_RE = /^[A-Za-z0-9_-]{8,64}$/;
    const CELL_RE = /^[A-Z]{1,3}[0-9]{1,7}$/;

    // ── Utilidades ───────────────────────────────────────────────────────────
    const ico = name => `<svg class="ico" aria-hidden="true"><use href="#${name}"/></svg>`;
    const $ = id => document.getElementById(id);
    const has = name => typeof ETB[name] === 'function';
    const call = (name, ...args) => { try { return has(name) ? ETB[name](...args) : undefined; } catch (e) { console.warn('[live]', name, e); } };
    const toast = (msg, type) => call('toast', msg, type || 'info');
    const reduced = () => !!(window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches);
    const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`;

    function uuid() {
        try { if (window.crypto && crypto.randomUUID) return crypto.randomUUID(); } catch (_) { }
        const b = new Uint8Array(16);
        try { crypto.getRandomValues(b); } catch (_) { for (let i = 0; i < 16; i++) b[i] = Math.random() * 256 | 0; }
        b[6] = (b[6] & 15) | 64; b[8] = (b[8] & 63) | 128;
        const h = Array.from(b, x => x.toString(16).padStart(2, '0')).join('');
        return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
    }
    function hashColor(s) { let h = 0; for (const ch of String(s)) h = (h * 31 + ch.charCodeAt(0)) >>> 0; return PALETTE[h % PALETTE.length]; }
    const cleanName = n => String(n || '').replace(/[\u0000-\u001f\u007f]/g, '').replace(/\s+/g, ' ').trim().slice(0, 40);
    const validColor = c => /^#[0-9a-fA-F]{6}$/.test(c || '');
    const colorOf = c => (validColor(c && c.color) ? c.color : hashColor((c && (c.userId || c.name)) || ''));
    function initials(name) {
        const parts = cleanName(name).split(/[\s._-]+/).filter(Boolean);
        if (!parts.length) return '?';
        const a = Array.from(parts[0]), z = parts.length > 1 ? Array.from(parts[parts.length - 1]) : a.slice(1);
        return ((a[0] || '') + (z[0] || '')).toUpperCase();
    }
    function storeGet(store, key) { try { return window[store].getItem(key); } catch (_) { return null; } }
    function storeSet(store, key, val) { try { window[store].setItem(key, val); } catch (_) { } }
    function textEl(tag, cls, txt) { const d = document.createElement(tag); if (cls) d.className = cls; d.textContent = txt; return d; }

    // ── Identidad (SÍNCRONA: index.html la usa en los headers de api()) ──────
    let user = null;
    try { user = JSON.parse(storeGet('localStorage', 'etb.user') || 'null'); } catch (_) { user = null; }
    if (!user || typeof user !== 'object') user = {};
    if (typeof user.userId !== 'string' || !ID_RE.test(user.userId)) user.userId = uuid();
    user.name = cleanName(user.name);
    try { const q = cleanName(new URLSearchParams(location.search).get('user')); if (q) user.name = q; } catch (_) { }
    if (!validColor(user.color)) user.color = hashColor(user.userId);
    let clientId = storeGet('sessionStorage', 'etb.clientId');
    if (!clientId || !ID_RE.test(clientId)) { clientId = uuid(); storeSet('sessionStorage', 'etb.clientId', clientId); }
    const needsWelcome = !user.name;
    // userId/color se guardan ya (estables entre pestañas); el nombre solo cuando el usuario lo confirma.
    storeSet('localStorage', 'etb.user', JSON.stringify({ userId: user.userId, name: user.name, color: user.color }));
    if (!user.name) user.name = 'Invitado-' + user.userId.replace(/-/g, '').slice(0, 4).toUpperCase();

    function publishIdentity(persist) {
        ETB.identity = { clientId, userId: user.userId, name: user.name, color: user.color };
        if (persist) storeSet('localStorage', 'etb.user', JSON.stringify({ userId: user.userId, name: user.name, color: user.color }));
    }
    publishIdentity(!needsWelcome);

    // Pestaña duplicada: el navegador copia sessionStorage → mismo clientId en dos pestañas,
    // y el servidor se las robaría mutuamente el stream. La pestaña nueva pregunta y, si otra
    // ya usa su clientId, genera uno nuevo.
    let channel = null;
    try {
        if (typeof BroadcastChannel === 'function') {
            channel = new BroadcastChannel('etb-live');
            channel.onmessage = ev => {
                const m = ev.data || {};
                if (m.clientId !== clientId) return;
                if (m.type === 'who') channel.postMessage({ type: 'mine', clientId });
                else if (m.type === 'mine') renewClientId();
            };
            channel.postMessage({ type: 'who', clientId });
        }
    } catch (_) { channel = null; }
    // Breve ventana para recibir la respuesta antes de abrir el stream (evita robarle el stream a la otra pestaña).
    const claimSettled = new Promise(res => setTimeout(res, channel ? 150 : 0));
    function renewClientId() {
        clientId = uuid();
        storeSet('sessionStorage', 'etb.clientId', clientId);
        publishIdentity(false);
        if (es) connect();
    }

    // ── Estado ────────────────────────────────────────────────────────────────
    let clients = [];               // PresenceEntry[] (incluye la propia)
    let status = 'connecting';
    let unsupported = false;
    let es = null, offlineTimer = null, retryTimer = null, retryDelay = 3000, watchdog = null;
    let keepaliveTimer = null, sendTimer = null, sendActive = false, lastSend = 0, lastEventAt = 0;
    let pendingResend = false;
    let lastActivity = 0, lastActivityTpl = null, editEndTimer = null;
    let nearLiveTimer = null, nearLiveBlockedUntil = 0;
    const sel = { cell: null, range: null };

    // ── Presencia local ───────────────────────────────────────────────────────
    function activeTpl() { return call('currentView') === 'editor' ? (call('activeTemplate') || null) : null; }
    function isCellEditing() { const st = call('state'); return !!(st && st.isEditing); }
    function currentPresence() {
        const view = call('currentView') || 'home';
        const t = activeTpl();
        let mode = view === 'dashboard' ? 'dashboard' : view === 'editor' ? 'view' : 'home';
        if (t && (isCellEditing() || (Date.now() - lastActivity < EDIT_WINDOW_MS && (!lastActivityTpl || lastActivityTpl === t.id)))) mode = 'edit';
        const p = {
            clientId, userId: user.userId, name: user.name, color: user.color, mode,
            templateId: t && t.id ? String(t.id) : null,
            templateName: t ? String(t.name || '').slice(0, 120) : null,
            sheet: null, cell: null, range: null
        };
        if (t) {
            p.sheet = String(call('activeSheetName') || '').slice(0, 64) || null;
            p.cell = sel.cell && CELL_RE.test(sel.cell) ? sel.cell : null;
            p.range = sel.range && /^[A-Z]{1,3}[0-9]{1,7}:[A-Z]{1,3}[0-9]{1,7}$/.test(sel.range) && sel.range.split(':')[0] !== sel.range.split(':')[1] ? sel.range : null;
        }
        return p;
    }
    function headers() {
        return { 'Content-Type': 'application/json', 'X-Client-Id': clientId, 'X-User-Id': user.userId, 'X-Client-Name': encodeURIComponent(user.name) };
    }
    function sendPresence(active) {
        if (unsupported || typeof fetch !== 'function') return;
        lastSend = Date.now();
        const body = currentPresence(); if (active) body.active = true;
        fetch('/api/live/presence', { method: 'POST', headers: headers(), body: JSON.stringify(body), keepalive: true, credentials: 'same-origin' })
            .then(r => { if (r.status === 404) pendingResend = true; })   // NO_STREAM → reenviar tras el próximo open
            .catch(() => { pendingResend = true; });
        updateSelfEntry(body);
    }
    // Throttle 250 ms (borde inicial + final); "active" se acumula entre envíos.
    function schedulePresence(active) {
        if (active) sendActive = true;
        if (sendTimer) return;
        const wait = Math.max(0, THROTTLE_MS - (Date.now() - lastSend));
        sendTimer = setTimeout(() => { sendTimer = null; const a = sendActive; sendActive = false; sendPresence(a); }, wait);
    }
    // Refleja nuestra presencia en la lista local para que la UI reaccione sin esperar al servidor.
    function updateSelfEntry(p) {
        const i = clients.findIndex(c => c.clientId === clientId);
        if (i < 0) return;
        clients[i] = Object.assign({}, clients[i], p); delete clients[i].active;
        renderAll();
    }

    // ── Conexión SSE ─────────────────────────────────────────────────────────
    function setStatus(s) {
        if (status === s) return;
        status = s;
        if (s === 'online') clearTimeout(offlineTimer);
        call('emit', 'live:status', { state: s });
        renderAll();
    }
    function armOffline() {
        if (offlineTimer) return;
        offlineTimer = setTimeout(() => { offlineTimer = null; if (status !== 'online') setStatus('offline'); }, OFFLINE_AFTER_MS);
    }
    function disarmOffline() { clearTimeout(offlineTimer); offlineTimer = null; }
    function markAlive() {
        lastEventAt = Date.now(); retryDelay = 3000;
        disarmOffline();
        if (status !== 'online') setStatus('online');
    }
    function closeStream() { try { if (es) es.close(); } catch (_) { } es = null; }
    function connect() {
        if (unsupported) return;
        if (typeof EventSource !== 'function') { degrade(); return; }
        closeStream(); clearTimeout(retryTimer);
        const qs = new URLSearchParams({ clientId, userId: user.userId, name: user.name, color: user.color });
        if (status === 'online') setStatus('connecting');
        armOffline();
        const src = es = new EventSource('/api/live/stream?' + qs.toString());
        src.addEventListener('open', () => {
            if (src !== es) return;
            markAlive(); pendingResend = false;
            sendPresence(true);   // (re)apertura → presencia completa
        });
        src.addEventListener('error', () => {
            if (src !== es) return;
            if (status === 'online') setStatus('connecting');
            armOffline();
            // EventSource reintenta solo (CONNECTING); si queda CLOSED (404, MIME, 503) lo decidimos nosotros.
            if (src.readyState === 2) probeStream();
        });
        const on = (name, fn) => src.addEventListener(name, ev => {
            if (src !== es) return;
            markAlive();
            let d = null; try { d = JSON.parse(ev.data); } catch (_) { return; }
            if (!d || typeof d !== 'object') return;
            try { fn(d); } catch (e) { console.warn('[live]', name, e); }
        });
        on('hello', d => { if (Array.isArray(d.clients)) setClients(d.clients); if (pendingResend) { pendingResend = false; sendPresence(true); } });
        on('presence', d => { if (Array.isArray(d.clients)) setClients(d.clients); if (pendingResend) { pendingResend = false; sendPresence(); } });
        on('ping', () => { });
        on('usage', d => call('emit', 'live:usage', d));
        ['saved', 'created', 'deleted'].forEach(t => on('template:' + t, d => onTemplateEvent(t, d)));
    }
    // Stream cerrado definitivamente: ¿existe el endpoint? 404/no-JSON → modo degradado; si no, reintento con backoff.
    function probeStream() {
        closeStream();
        fetch('/api/live/presence', { method: 'GET', credentials: 'same-origin', headers: { Accept: 'application/json' } })
            .then(r => {
                const json = /json/i.test(r.headers.get('content-type') || '');
                if (r.status === 404 || (r.ok && !json)) degrade(); else scheduleReconnect();
            })
            .catch(() => scheduleReconnect());
    }
    function scheduleReconnect() {
        clearTimeout(retryTimer);
        retryTimer = setTimeout(connect, retryDelay);
        retryDelay = Math.min(30000, retryDelay * 2);
    }
    function degrade() {
        unsupported = true;
        closeStream();
        disarmOffline(); clearTimeout(retryTimer); clearInterval(keepaliveTimer); clearInterval(watchdog);
        clients = [];
        status = 'offline';
        call('emit', 'live:status', { state: 'offline' });
        renderAll();
    }

    function setClients(list) {
        clients = list.filter(c => c && typeof c.clientId === 'string');
        call('emit', 'live:presence', { clients: clients.slice(), self: selfEntry() });
        trackJoins();
        renderAll();
        scheduleNearLive();
    }
    function selfEntry() { return clients.find(c => c.clientId === clientId) || Object.assign({ connectedAt: null, lastActive: null, idle: false }, currentPresence()); }

    // ── Agrupación por usuario (varias pestañas → un avatar; gana el "mejor" modo) ──
    const MODE_RANK = { edit: 4, view: 3, dashboard: 2, home: 1 };
    function groupUsers(list) {
        const m = new Map();
        for (const c of list) {
            const key = c.userId || c.clientId;
            const g = m.get(key);
            if (!g) { m.set(key, Object.assign({}, c, { tabs: 1 })); continue; }
            const tabs = g.tabs + 1, idle = g.idle && c.idle;
            const rc = (MODE_RANK[c.mode] || 0) + (c.idle ? 0 : 0.5), rg = (MODE_RANK[g.mode] || 0) + (g.idle ? 0 : 0.5);
            if (rc > rg) Object.assign(g, c);
            g.tabs = tabs; g.idle = idle;
        }
        return [...m.values()];
    }
    const byRank = (a, b) => ((MODE_RANK[b.mode] || 0) - (MODE_RANK[a.mode] || 0)) || ((a.idle ? 1 : 0) - (b.idle ? 1 : 0)) || String(a.name).localeCompare(String(b.name), 'es');
    function othersOn(templateId) {
        if (!templateId) return [];
        return clients.filter(c => c.templateId === templateId && c.userId !== user.userId && c.clientId !== clientId);
    }

    // ── Avatares ──────────────────────────────────────────────────────────────
    const MODE_LABEL = { edit: 'editando', view: 'viendo', dashboard: 'en Actividad', home: 'en Inicio' };
    function describe(c) {
        let s = c.name || 'Invitado';
        if (c.templateName && (c.mode === 'edit' || c.mode === 'view')) s += ` · ${MODE_LABEL[c.mode]} «${c.templateName}»`;
        else s += ` · ${MODE_LABEL[c.mode] || 'conectado'}`;
        if (c.sheet && (c.cell || c.range)) s += ` (${c.sheet}!${c.range || c.cell})`;
        if (c.idle) s += ' · inactivo';
        if (c.tabs > 1) s += ` · ${c.tabs} pestañas`;
        return s;
    }
    function avatarEl(c, opts = {}) {
        const el = document.createElement(opts.button ? 'button' : 'span');
        if (opts.button) el.type = 'button';
        const ini = document.createElement('span'); ini.className = 'live-av-ini';
        el.appendChild(ini);
        paintAvatar(el, c, opts);
        return el;
    }
    // Actualiza un avatar existente (sin recrearlo → sin reiniciar animaciones ni perder el foco).
    function paintAvatar(el, c, opts = {}) {
        el.className = 'live-av' + (opts.size ? ' live-av-' + opts.size : '') + (c.idle ? ' is-idle' : '') +
            (c.mode === 'edit' && !opts.noBadge ? ' is-editing' : '') + (opts.extraClass ? ' ' + opts.extraClass : '');
        el.style.setProperty('--av', colorOf(c));
        const ini = el.querySelector('.live-av-ini'); const txt = initials(c.name);
        if (ini.textContent !== txt) ini.textContent = txt;
        let badge = el.querySelector('.live-av-badge');
        const wantBadge = c.mode === 'edit' && !opts.noBadge;
        if (wantBadge && !badge) { badge = document.createElement('span'); badge.className = 'live-av-badge'; badge.innerHTML = ico('i-pencil'); el.appendChild(badge); }
        else if (!wantBadge && badge) badge.remove();
        const title = opts.title || describe(c);
        el.title = title;                       // propiedad DOM: nunca se interpreta como HTML
        el.setAttribute('aria-label', title);
    }
    // Reconciliación por clave: reutiliza nodos existentes y solo anima los nuevos.
    function syncKeyed(host, items, keyOf, make, update) {
        const existing = new Map();
        Array.from(host.children).forEach(ch => { if (ch.dataset && ch.dataset.key) existing.set(ch.dataset.key, ch); else ch.remove(); });
        let prev = null;
        items.forEach((it, i) => {
            const k = keyOf(it);
            let el = existing.get(k);
            if (el) existing.delete(k); else { el = make(it, i); el.dataset.key = k; }
            update(el, it, i);
            const ref = prev ? prev.nextSibling : host.firstChild;
            if (el !== ref) host.insertBefore(el, ref);
            prev = el;
        });
        existing.forEach(el => el.remove());
    }
    function statusLabel(n) {
        if (unsupported) return 'Sin tiempo real';
        if (status === 'offline') return 'Sin conexión';
        if (status === 'connecting') return 'Conectando…';
        return n === 1 ? '1 conectado' : `${n} conectados`;
    }

    // #live-indicator: punto de conexión + nº conectados + avatar propio (perfil)
    let ind = null;
    function renderIndicator() {
        const host = $('live-indicator'); if (!host) return;
        if (!ind || ind.host !== host || !host.contains(ind.btn)) {
            host.classList.add('live-indicator');
            host.innerHTML = '';
            const btn = document.createElement('button'); btn.type = 'button'; btn.className = 'live-ind-btn';
            btn.innerHTML = `<span class="live-dot" aria-hidden="true"></span><span class="live-ind-ico"></span><span class="live-ind-txt"></span>`;
            btn.addEventListener('click', e => { e.stopPropagation(); toggleUsersPopover(btn); });
            const me = avatarEl({ name: user.name, color: user.color }, { button: true, size: 'sm', noBadge: true });
            me.addEventListener('click', e => { e.stopPropagation(); openProfile(me); });
            host.append(btn, me);
            ind = { host, btn, me, icon: '' };
        }
        const n = groupUsers(clients).length;
        host.dataset.state = unsupported ? 'unsupported' : status;
        const icon = status === 'online' && !unsupported ? 'i-users' : 'i-wifi-off';
        if (ind.icon !== icon) { ind.btn.querySelector('.live-ind-ico').innerHTML = ico(icon); ind.icon = icon; }
        const label = statusLabel(n);
        const txt = ind.btn.querySelector('.live-ind-txt'); if (txt.textContent !== label) txt.textContent = label;
        ind.btn.title = unsupported ? 'El servidor no ofrece presencia en tiempo real' : status === 'online' ? 'Ver quién está conectado' : 'Reconectando con el servidor…';
        paintAvatar(ind.me, { name: user.name, color: user.color }, { size: 'sm', noBadge: true, extraClass: 'live-av-self', title: `${user.name} (tú) · Cambiar nombre o color` });
    }

    // #presence-stack: OTROS usuarios en la plantilla activa (máx. 5 + "+N"), clic → seguir
    function renderPresenceStack() {
        const host = $('presence-stack'); if (!host) return;
        host.classList.add('live-stack');
        const t = activeTpl();
        const users = t ? groupUsers(othersOn(t.id)).sort(byRank) : [];
        host.hidden = !users.length;
        host.classList.toggle('is-stale', status !== 'online');   // presencia desactualizada mientras reconecta
        host.setAttribute('aria-label', users.length ? `${plural(users.length, 'persona', 'personas')} en esta plantilla` : '');
        const items = users.slice(0, 5);
        if (users.length > 5) items.push({ __more: users.length - 5, rest: users.slice(5) });
        syncKeyed(host, items, u => (u.__more ? '__more' : 'u:' + (u.userId || u.clientId)),
            (u, i) => {
                if (u.__more) {
                    const m = document.createElement('button'); m.type = 'button'; m.className = 'live-av live-av-more';
                    m.addEventListener('click', e => { e.stopPropagation(); toggleUsersPopover(m); });
                    return m;
                }
                const el = avatarEl(u, { button: true });
                el.style.setProperty('--i', i);
                el.addEventListener('click', () => { if (el._user) follow(el._user); });
                return el;
            },
            (el, u) => {
                if (u.__more) { el.textContent = '+' + u.__more; el.title = u.rest.map(x => x.name).join(', '); return; }
                el._user = u;
                paintAvatar(el, u, { button: true, title: describe(u) + ' · Clic para ir a su celda' });
            });
    }
    function follow(u) {
        const cur = call('activeSheetName');
        const switching = u.sheet && u.sheet !== cur;
        if (switching) call('switchSheetByName', u.sheet);
        const target = u.cell || (u.range ? u.range.split(':')[0] : null);
        if (target) setTimeout(() => call('selectCell', target), switching ? 80 : 0);
        else if (!switching) toast(`${u.name} no tiene una celda seleccionada`, 'info');
    }

    // Popover de usuarios conectados (se actualiza en vivo mientras está abierto)
    let popEl = null, popAnchor = null;
    function closePopover() { if (popEl) { popEl.remove(); popEl = null; popAnchor = null; } }
    function placePopover(el, anchor) {
        if (!el.parentNode) document.body.appendChild(el);
        const r = anchor.getBoundingClientRect(), w = el.offsetWidth, h = el.offsetHeight;
        const left = Math.min(Math.max(8, r.right - w), innerWidth - w - 8);
        let top = r.bottom + 8; if (top + h > innerHeight - 8) top = Math.max(8, r.top - h - 8);
        el.style.left = Math.max(8, left) + 'px'; el.style.top = top + 'px';
    }
    function toggleUsersPopover(anchor) {
        if (popEl && popEl.dataset.kind === 'users') { closePopover(); return; }
        closePopover();
        const el = document.createElement('div'); el.className = 'live-pop'; el.dataset.kind = 'users';
        el.setAttribute('role', 'dialog'); el.setAttribute('aria-label', 'Usuarios conectados');
        el.addEventListener('click', e => e.stopPropagation());
        popEl = el; popAnchor = anchor;
        fillUsersPopover();
        placePopover(el, anchor);
    }
    function fillUsersPopover() {
        const el = popEl; if (!el || el.dataset.kind !== 'users') return;
        const users = groupUsers(clients).sort((a, b) => (a.userId === user.userId ? -1 : b.userId === user.userId ? 1 : byRank(a, b)));
        // Reconstruir solo si cambió algo visible (evita perder clics/hover con cada evento de presencia).
        const sig = [unsupported, status].concat(users.map(u => [u.userId, u.name, u.color, u.mode, u.idle, u.templateId, u.templateName, u.sheet, u.cell, u.range, u.tabs].join('·'))).join('|');
        if (el.dataset.sig === sig) return;
        el.dataset.sig = sig;
        el.innerHTML = '';
        const head = document.createElement('div'); head.className = 'live-pop-head';
        head.innerHTML = `${ico('i-users')}<span>Conectados ahora</span><span class="live-pop-count"></span>`;
        head.querySelector('.live-pop-count').textContent = unsupported || status !== 'online' ? '—' : String(users.length);
        el.appendChild(head);
        const list = document.createElement('div'); list.className = 'live-pop-list';
        if (unsupported) list.appendChild(textEl('div', 'live-pop-empty', 'El servidor no ofrece presencia en tiempo real.'));
        else if (status !== 'online') list.appendChild(textEl('div', 'live-pop-empty', 'Sin conexión con el servidor. Reintentando…'));
        else if (!users.length) list.appendChild(textEl('div', 'live-pop-empty', 'Nadie más conectado.'));
        if (!unsupported && status === 'online') users.forEach(u => {
            const me = u.userId === user.userId;
            const row = document.createElement('div'); row.className = 'live-pop-row' + (me ? ' is-self' : '');
            row.appendChild(avatarEl(u, { size: 'sm' }));
            const tx = document.createElement('div'); tx.className = 'live-pop-txt';
            const nm = document.createElement('div'); nm.className = 'live-pop-name';
            nm.appendChild(textEl('span', 'live-pop-nm', u.name));
            if (me) nm.appendChild(textEl('span', 'live-pop-you', 'Tú'));
            tx.appendChild(nm);
            const where = u.templateName && (u.mode === 'edit' || u.mode === 'view')
                ? `${u.mode === 'edit' ? 'Editando' : 'Viendo'} «${u.templateName}»${u.sheet && (u.cell || u.range) ? ` · ${u.sheet}!${u.range || u.cell}` : ''}`
                : (u.mode === 'dashboard' ? 'En Actividad' : 'En Inicio');
            tx.appendChild(textEl('div', 'live-pop-sub', where + (u.idle ? ' · inactivo' : '') + (u.tabs > 1 ? ` · ${u.tabs} pestañas` : '')));
            row.appendChild(tx);
            if (u.templateId && !me && has('openTemplate')) {
                const go = document.createElement('button'); go.type = 'button'; go.className = 'live-pop-go';
                go.title = 'Ir con esta persona'; go.setAttribute('aria-label', `Ir con ${u.name}`); go.innerHTML = ico('i-arrow-right');
                go.addEventListener('click', () => {
                    closePopover();
                    const t = activeTpl();
                    if (t && t.id === u.templateId) { follow(u); return; }
                    Promise.resolve(call('openTemplate', u.templateId)).then(() => setTimeout(() => follow(u), 350)).catch(() => { });
                });
                row.appendChild(go);
            }
            list.appendChild(row);
        });
        el.appendChild(list);
    }
    document.addEventListener('click', e => { if (popEl && !popEl.contains(e.target)) closePopover(); });
    document.addEventListener('keydown', e => { if (e.key === 'Escape' && popEl) closePopover(); });
    window.addEventListener('resize', () => { if (popEl && popAnchor && popAnchor.isConnected) placePopover(popEl, popAnchor); else closePopover(); });

    // ── Perfil / bienvenida ───────────────────────────────────────────────────
    const COLOR_NAMES = { '#217346': 'Verde', '#0078d4': 'Azul', '#8764b8': 'Lavanda', '#c239b3': 'Magenta', '#d13438': 'Rojo', '#ca5010': 'Naranja', '#986f0b': 'Ámbar', '#038387': 'Turquesa' };
    function colorSwatches(current, onPick) {
        const wrap = document.createElement('div'); wrap.className = 'live-swatches'; wrap.setAttribute('role', 'radiogroup'); wrap.setAttribute('aria-label', 'Color');
        PALETTE.forEach(c => {
            const b = document.createElement('button'); b.type = 'button'; b.className = 'live-swatch' + (c === current ? ' is-on' : '');
            b.style.setProperty('--sw', c); b.setAttribute('role', 'radio'); b.setAttribute('aria-checked', String(c === current));
            b.title = COLOR_NAMES[c] || c; b.setAttribute('aria-label', b.title);
            b.addEventListener('click', () => {
                wrap.querySelectorAll('.live-swatch').forEach(x => { x.classList.remove('is-on'); x.setAttribute('aria-checked', 'false'); });
                b.classList.add('is-on'); b.setAttribute('aria-checked', 'true'); onPick(c);
            });
            wrap.appendChild(b);
        });
        return wrap;
    }
    function applyIdentity(name, color) {
        user.name = cleanName(name) || user.name;
        if (validColor(color)) user.color = color;
        publishIdentity(true);
        sendPresence(true);   // el servidor fusiona name/color en la entrada del stream
        renderAll();
    }
    function openProfile(anchor) {
        if (popEl && popEl.dataset.kind === 'profile') { closePopover(); return; }
        closePopover();
        const el = document.createElement('div'); el.className = 'live-pop live-profile'; el.dataset.kind = 'profile';
        el.setAttribute('role', 'dialog'); el.setAttribute('aria-label', 'Tu perfil');
        el.addEventListener('click', e => e.stopPropagation());
        let color = user.color;
        el.innerHTML = `<div class="live-pop-head">${ico('i-user')}<span>Tu perfil</span></div>
            <div class="live-profile-row"><span class="live-profile-av"></span>
            <label class="live-field"><span>Nombre visible</span><input type="text" maxlength="40" autocomplete="nickname" spellcheck="false"></label></div>
            <div class="live-field-lbl">Color</div>`;
        const inp = el.querySelector('input'); inp.value = user.name;
        const avHost = el.querySelector('.live-profile-av');
        const preview = avatarEl({ name: user.name, color }, { size: 'md', title: 'Vista previa' }); avHost.appendChild(preview);
        const repaint = () => paintAvatar(preview, { name: inp.value || user.name, color }, { size: 'md', title: 'Vista previa' });
        el.appendChild(colorSwatches(color, c => { color = c; repaint(); }));
        inp.addEventListener('input', repaint);
        const act = document.createElement('div'); act.className = 'live-actions';
        const cancel = textEl('button', 'live-btn live-btn-ghost', 'Cancelar'); cancel.type = 'button';
        const ok = textEl('button', 'live-btn live-btn-primary', 'Guardar'); ok.type = 'button';
        act.append(cancel, ok); el.appendChild(act);
        const save = () => {
            const n = cleanName(inp.value);
            if (!n) { inp.classList.add('is-invalid'); inp.focus(); return; }
            const changed = n !== user.name || color !== user.color;
            applyIdentity(n, color); closePopover();
            if (changed) toast('Perfil actualizado', 'success');
        };
        ok.addEventListener('click', save);
        cancel.addEventListener('click', closePopover);
        inp.addEventListener('keydown', e => { if (e.key === 'Enter') save(); });
        popEl = el; popAnchor = anchor;
        placePopover(el, anchor);
        setTimeout(() => { inp.focus(); inp.select(); }, 0);
    }
    function openWelcome() {
        if (document.querySelector('.live-welcome')) return;
        const ov = document.createElement('div'); ov.className = 'live-welcome';
        ov.innerHTML = `<div class="live-welcome-card fx-glass" role="dialog" aria-modal="true" aria-labelledby="live-wel-title" aria-describedby="live-wel-desc">
            <div class="live-welcome-hero">${ico('i-users')}</div>
            <h2 id="live-wel-title">¡Te damos la bienvenida!</h2>
            <p id="live-wel-desc">¿Cómo te llamas? Tu equipo verá en tiempo real quién está viendo y editando cada plantilla.</p>
            <div class="live-wel-preview"></div>
            <label class="live-field"><span>Tu nombre</span><input type="text" maxlength="40" placeholder="Ej. Ana López" autocomplete="name" spellcheck="false"></label>
            <div class="live-field-lbl">Elige tu color</div>
            <div class="live-actions"><button type="button" class="live-btn live-btn-ghost" data-act="skip">Omitir</button><button type="button" class="live-btn live-btn-primary" data-act="ok">Continuar${ico('i-arrow-right')}</button></div>
        </div>`;
        const card = ov.querySelector('.live-welcome-card');
        const inp = ov.querySelector('input');
        let color = user.color;
        const preview = avatarEl({ name: user.name, color }, { size: 'lg', title: 'Vista previa' });
        ov.querySelector('.live-wel-preview').appendChild(preview);
        const repaint = () => { paintAvatar(preview, { name: inp.value || user.name, color }, { size: 'lg', title: 'Vista previa' }); inp.classList.remove('is-invalid'); };
        card.insertBefore(colorSwatches(color, c => { color = c; repaint(); }), card.querySelector('.live-actions'));
        inp.addEventListener('input', repaint);
        let closed = false;
        const close = () => {
            if (closed) return; closed = true;
            ov.classList.add('is-leaving');
            setTimeout(() => ov.remove(), reduced() ? 0 : 260);
        };
        const accept = () => {
            const n = cleanName(inp.value);
            if (!n) { inp.classList.add('is-invalid'); inp.focus(); return; }
            applyIdentity(n, color); close(); toast(`¡Hola, ${n}!`, 'success');
        };
        const skip = () => { applyIdentity(user.name, color); close(); };
        ov.querySelector('[data-act="ok"]').addEventListener('click', accept);
        ov.querySelector('[data-act="skip"]').addEventListener('click', skip);
        inp.addEventListener('keydown', e => { if (e.key === 'Enter') accept(); });
        ov.addEventListener('keydown', e => {
            if (e.key === 'Escape') { e.stopPropagation(); skip(); }
            if (e.key === 'Tab') {   // atrapa el foco dentro del diálogo
                const f = Array.from(card.querySelectorAll('input,button'));
                const i = f.indexOf(document.activeElement);
                if (e.shiftKey && i <= 0) { e.preventDefault(); f[f.length - 1].focus(); }
                else if (!e.shiftKey && i === f.length - 1) { e.preventDefault(); f[0].focus(); }
            }
        });
        document.body.appendChild(ov);
        setTimeout(() => inp.focus(), 60);
    }

    // ── Home ──────────────────────────────────────────────────────────────────
    let homeCount = -1;
    function renderHomeOnline() {
        const host = $('home-online'); if (!host) return;
        if (unsupported) { host.hidden = true; return; }
        host.classList.add('live-home-online');
        host.hidden = false;
        host.dataset.state = status;
        let head = host.querySelector(':scope > .live-ho-head'), row = host.querySelector(':scope > .live-ho-avatars');
        if (!head || !row) {
            host.innerHTML = '';
            head = document.createElement('div'); head.className = 'live-ho-head';
            head.innerHTML = `<span class="live-dot" aria-hidden="true"></span><span class="live-ho-title">Conectados ahora</span><span class="live-ho-count"></span>`;
            row = document.createElement('div'); row.className = 'live-ho-avatars';
            host.append(head, row); homeCount = -1;
        }
        const users = groupUsers(clients).sort((a, b) => (a.userId === user.userId ? -1 : b.userId === user.userId ? 1 : byRank(a, b)));
        const cnt = head.querySelector('.live-ho-count');
        if (status !== 'online') { cnt.textContent = '—'; homeCount = -1; }
        else if (users.length !== homeCount) {
            if (ETB.fx && typeof ETB.fx.countUp === 'function' && homeCount >= 0) ETB.fx.countUp(cnt, users.length, 400);
            else cnt.textContent = String(users.length);
            homeCount = users.length;
        }
        const shown = status === 'online' ? users.slice(0, 12) : [];
        const items = shown.slice();
        if (status === 'online' && users.length > 12) items.push({ __more: users.length - 12 });
        if (!items.length) items.push({ __empty: true });
        syncKeyed(row, items, u => (u.__more ? '__more' : u.__empty ? '__empty' : 'u:' + (u.userId || u.clientId)),
            u => (u.__empty ? textEl('div', 'live-ho-empty', '') : u.__more ? textEl('span', 'live-av live-av-md live-av-more', '') : avatarEl(u, { size: 'md' })),
            (el, u) => {
                if (u.__empty) { el.textContent = status === 'connecting' ? 'Conectando…' : status === 'online' ? 'Solo tú por ahora' : 'Sin conexión'; return; }
                if (u.__more) { el.textContent = '+' + u.__more; return; }
                const me = u.userId === user.userId;
                paintAvatar(el, u, { size: 'md', extraClass: me ? 'live-av-self' : '', title: describe(u) + (me ? ' (tú)' : '') });
            });
    }
    // Insignias en tarjetas: mini avatares + "N viendo" / "N editando"
    function renderCardBadges() {
        document.querySelectorAll('.tpl-live[data-live-for]').forEach(host => {
            const id = host.getAttribute('data-live-for');
            const users = unsupported ? [] : groupUsers(clients.filter(c => c.templateId === id && c.clientId !== clientId && c.userId !== user.userId)).sort(byRank);
            const sig = users.map(u => [u.userId, u.mode, u.idle ? 1 : 0, u.color, u.name].join('·')).join('|');
            host.classList.toggle('is-on', users.length > 0);
            if (host.dataset.sig === sig && (users.length ? host.firstChild : !host.firstChild)) {
                const pill = host.firstChild; if (pill) pill.title = users.map(describe).join('\n');
                return;
            }
            host.dataset.sig = sig;
            host.innerHTML = '';
            if (!users.length) return;
            const editors = users.filter(u => u.mode === 'edit').length;
            const pill = document.createElement('div'); pill.className = 'live-badge' + (editors ? ' is-editing' : '');
            const avs = document.createElement('span'); avs.className = 'live-badge-avs';
            users.slice(0, 3).forEach(u => avs.appendChild(avatarEl(u, { size: 'xs', noBadge: true })));
            pill.appendChild(avs);
            const tx = document.createElement('span'); tx.className = 'live-badge-txt';
            tx.innerHTML = ico(editors ? 'i-pencil' : 'i-eye');
            const viewers = users.length - editors;
            const label = editors ? `${editors} editando` + (viewers ? ` · ${viewers} viendo` : '') : `${viewers} viendo`;
            tx.appendChild(document.createTextNode(label));
            pill.appendChild(tx);
            pill.title = users.map(describe).join('\n');
            pill.setAttribute('aria-label', label + ' ahora');
            host.appendChild(pill);
        });
    }
    function renderStatusBar() {
        const host = $('status-live'); if (!host) return;
        if (!host.querySelector(':scope > .live-dot')) {
            host.classList.add('live-status');
            host.innerHTML = '<span class="live-dot" aria-hidden="true"></span><span class="live-status-txt"></span>';
        }
        host.dataset.state = unsupported ? 'unsupported' : status;
        const label = statusLabel(groupUsers(clients).length);
        const t = host.querySelector('.live-status-txt'); if (t && t.textContent !== label) t.textContent = label;
    }

    // ── Cursores remotos ─────────────────────────────────────────────────────
    // Capa overlay absoluta dentro de #grid-scroll (se desplaza con el contenido). Las posiciones se
    // miden con getBoundingClientRect, que ya incluye el zoom CSS de #grid-zoom.
    let cursorRaf = 0;
    function scheduleCursors() {
        if (cursorRaf) return;
        cursorRaf = (window.requestAnimationFrame || (f => setTimeout(f, 16)))(() => { cursorRaf = 0; try { renderCursors(); } catch (e) { console.warn('[live] cursores', e); } });
    }
    function colNum(letters) { let n = 0; for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
    function findCellEl(table, addr) {
        if (!CELL_RE.test(addr)) return null;
        let el = null;
        try { el = has('gridCellEl') ? ETB.gridCellEl(addr) : null; } catch (_) { el = null; }
        if (!el) el = table.querySelector(`[data-addr="${addr}"]`);
        if (el) return el;
        // Celda cubierta por una combinación: buscar la maestra que la contiene.
        const m = /^([A-Z]+)(\d+)$/.exec(addr), c = colNum(m[1]), r = +m[2];
        for (const td of table.querySelectorAll('td[colspan], td[rowspan]')) {
            const mc = +td.dataset.col, mr = +td.dataset.row;
            if (mc && mr && c >= mc && c < mc + (td.colSpan || 1) && r >= mr && r < mr + (td.rowSpan || 1)) return td;
        }
        return null;
    }
    function cellRect(table, addr, base) {
        const el = findCellEl(table, addr); if (!el) return null;
        const r = el.getBoundingClientRect();
        if (r.width < 1 || r.height < 1) return null;   // columna/fila oculta
        return { l: r.left - base.left, t: r.top - base.top, r: r.right - base.left, b: r.bottom - base.top };
    }
    function renderCursors() {
        const scroll = $('grid-scroll'), table = $('grid-table');
        let layer = $('live-cursor-layer');
        const t = activeTpl();
        if (!scroll || !table || !t || unsupported) { if (layer) layer.textContent = ''; return; }
        if (!layer || layer.parentNode !== scroll) {
            if (layer) layer.remove();
            layer = document.createElement('div'); layer.id = 'live-cursor-layer'; layer.setAttribute('aria-hidden', 'true');
            scroll.appendChild(layer);
        }
        layer.classList.toggle('is-stale', status !== 'online');
        const sheet = call('activeSheetName');
        const others = groupUsers(othersOn(t.id).filter(c => c.sheet === sheet && (c.cell || c.range)));
        // Origen del contenido desplazable = rect del scroll + borde − desplazamiento.
        const sr = scroll.getBoundingClientRect();
        const base = { left: sr.left + scroll.clientLeft - scroll.scrollLeft, top: sr.top + scroll.clientTop - scroll.scrollTop };
        const head = table.tHead ? table.tHead.getBoundingClientRect().height : 26;
        const seen = new Set();
        others.forEach(u => {
            const [a, b] = String(u.range || u.cell).split(':');
            const r1 = cellRect(table, a, base); if (!r1) return;
            const r2 = b ? (cellRect(table, b, base) || r1) : r1;
            const box = { l: Math.min(r1.l, r2.l), t: Math.min(r1.t, r2.t), r: Math.max(r1.r, r2.r), b: Math.max(r1.b, r2.b) };
            const key = 'c-' + (u.userId || u.clientId); seen.add(key);
            let el = null;
            for (const ch of layer.children) if (ch.dataset.k === key) { el = ch; break; }
            const fresh = !el;
            if (fresh) {
                el = document.createElement('div'); el.className = 'live-cursor'; el.dataset.k = key;
                const flag = document.createElement('span'); flag.className = 'live-cursor-flag';
                flag.innerHTML = `<span class="live-cursor-ico">${ico('i-pencil')}</span><span class="live-cursor-name"></span>`;
                el.appendChild(flag);
                layer.appendChild(el);
            }
            el.style.setProperty('--c', colorOf(u));
            el.style.transform = `translate(${Math.round(box.l)}px, ${Math.round(box.t)}px)`;
            el.style.width = Math.max(4, Math.round(box.r - box.l)) + 'px';
            el.style.height = Math.max(4, Math.round(box.b - box.t)) + 'px';
            if (fresh) { el.classList.add('is-new'); setTimeout(() => el.classList.remove('is-new'), 400); }
            el.classList.toggle('is-idle', !!u.idle);
            el.classList.toggle('is-editing', u.mode === 'edit');
            // Etiqueta debajo si quedaría oculta tras los encabezados fijos.
            el.classList.toggle('flag-below', box.t - scroll.scrollTop < head + 20);
            const nm = el.querySelector('.live-cursor-name'); if (nm.textContent !== u.name) nm.textContent = u.name;
            el.title = describe(u);
        });
        Array.from(layer.children).forEach(el => { if (!seen.has(el.dataset.k)) el.remove(); });
    }
    let mo = null, moTarget = null, zoomMo = null, zoomTarget = null, scrollTarget = null;
    function observeGrid() {
        if (typeof MutationObserver !== 'function') return;
        const table = $('grid-table');
        if (table && moTarget !== table) {
            if (mo) mo.disconnect();
            mo = new MutationObserver(scheduleCursors);
            mo.observe(table, { childList: true, subtree: true, attributes: true, attributeFilter: ['style', 'colspan', 'rowspan', 'class'] });
            moTarget = table;
        }
        const zoom = $('grid-zoom');
        if (zoom && zoomTarget !== zoom) {
            if (zoomMo) zoomMo.disconnect();
            zoomMo = new MutationObserver(scheduleCursors);
            zoomMo.observe(zoom, { attributes: true, attributeFilter: ['style'] });
            zoomTarget = zoom;
        }
        const scroll = $('grid-scroll');
        if (scroll && scrollTarget !== scroll) {
            // Solo para reubicar etiquetas cerca del borde superior; la capa ya se desplaza sola.
            scroll.addEventListener('scroll', scheduleCursors, { passive: true });
            scrollTarget = scroll;
        }
    }
    window.addEventListener('resize', scheduleCursors);

    // ── Entradas / salidas (toasts, solo plantilla activa, con throttle) ──────
    const knownOnTemplate = new Map();
    let trackedTpl = undefined, joinToastAt = 0;
    function trackJoins() {
        const t = activeTpl(); const tid = t ? t.id : null;
        const now = new Map(groupUsers(othersOn(tid)).map(u => [u.userId || u.clientId, u.name]));
        if (trackedTpl !== tid) { trackedTpl = tid; knownOnTemplate.clear(); now.forEach((n, id) => knownOnTemplate.set(id, n)); return; }
        if (!tid) return;
        const joined = [...now].filter(([id]) => !knownOnTemplate.has(id)).map(([, n]) => n);
        const left = [...knownOnTemplate].filter(([id]) => !now.has(id)).map(([, n]) => n);
        knownOnTemplate.clear(); now.forEach((n, id) => knownOnTemplate.set(id, n));
        if (!(joined.length || left.length) || Date.now() - joinToastAt < JOIN_TOAST_GAP_MS) return;
        joinToastAt = Date.now();
        if (joined.length) toast(joined.length === 1 ? `${joined[0]} se unió a esta plantilla` : `${joined.length} personas se unieron a esta plantilla`, 'info');
        else toast(left.length === 1 ? `${left[0]} salió de esta plantilla` : `${left.length} personas salieron de esta plantilla`, 'info');
    }

    // ── Eventos de plantilla remotos (§3.5) ───────────────────────────────────
    let homeRefreshTimer = null;
    function refreshHomeSoon() { clearTimeout(homeRefreshTimer); homeRefreshTimer = setTimeout(() => { if (call('currentView') === 'home') call('refreshHome'); }, HOME_REFRESH_MS); }
    function openDocs() {
        try {
            const st = call('state') || {};
            return (Array.isArray(st.docs) ? st.docs : []).map(d => d && d.template).filter(t => t && t.id);
        } catch (_) { return []; }
    }
    async function fetchTemplate(id) {
        if (has('api')) {
            let r = await ETB.api('GET', '/api/templates/' + encodeURIComponent(id));
            if (r && typeof r.json === 'function') r = await r.json();
            return r;
        }
        const res = await fetch('/api/templates/' + encodeURIComponent(id), { headers: headers(), credentials: 'same-origin' });
        return res.ok ? res.json() : null;
    }
    function onTemplateEvent(type, d) {
        call('emit', 'live:template', Object.assign({ type }, d));
        const mine = !!(d.by && d.by.clientId === clientId);
        if (!mine && call('currentView') === 'home') refreshHomeSoon();
        if (mine || typeof d.templateId !== 'string') return;
        const who = cleanName(d.by && d.by.name) || 'Otro usuario';
        const t = activeTpl();
        const isActive = !!(t && t.id === d.templateId);
        if (type === 'saved') {
            if (isActive) {
                if (typeof t.revision === 'number' && typeof d.revision === 'number' && d.revision <= t.revision) return;   // ya la tenemos
                if (call('isDirty') || isCellEditing()) showRemoteBanner(who, d);
                else if (has('reloadActiveFromServer')) {
                    Promise.resolve(call('reloadActiveFromServer'))
                        .then(() => { hideRemoteBanner(); toast(`Actualizado por ${who}`, 'info'); })
                        .catch(() => showRemoteBanner(who, d));
                } else showRemoteBanner(who, d);
            } else if (has('replaceTemplateInDocs')) {
                const doc = openDocs().find(x => x.id === d.templateId);
                if (doc && !(typeof doc.revision === 'number' && typeof d.revision === 'number' && d.revision <= doc.revision)) {
                    fetchTemplate(d.templateId).then(tpl => { if (tpl && tpl.id === d.templateId) call('replaceTemplateInDocs', tpl); }).catch(() => { });
                }
            }
        } else if (type === 'deleted') {
            const doc = isActive ? t : openDocs().find(x => x.id === d.templateId);
            if (doc) toast(`${who} eliminó «${doc.name || 'la plantilla'}». Tu pestaña sigue abierta: guárdala si quieres conservarla.`, 'warn');
            if (isActive) hideRemoteBanner();
        }
    }
    function showRemoteBanner(who, d) {
        const host = $('remote-banner'); if (!host) { toast(`${who} guardó cambios en esta plantilla`, 'warn'); return; }
        host.classList.add('live-remote-banner');
        host.innerHTML = `<span class="live-rb-ico">${ico('i-refresh')}</span><span class="live-rb-txt"><strong></strong><span></span></span>` +
            `<button type="button" class="live-btn live-btn-primary live-btn-sm" data-act="reload">${ico('i-refresh')}Recargar</button>` +
            `<button type="button" class="live-btn live-btn-ghost live-btn-sm" data-act="ignore">Ignorar</button>`;
        host.querySelector('.live-rb-txt strong').textContent = `${who} guardó cambios`;
        host.querySelector('.live-rb-txt span').textContent = ` · Tienes cambios sin guardar${typeof d.revision === 'number' ? ` (versión ${d.revision})` : ''}`;
        host.querySelector('[data-act="reload"]').addEventListener('click', e => {
            e.currentTarget.disabled = true;
            Promise.resolve(call('reloadActiveFromServer')).catch(() => toast('No se pudo recargar la plantilla', 'error')).finally(hideRemoteBanner);
        });
        host.querySelector('[data-act="ignore"]').addEventListener('click', hideRemoteBanner);
        host.dataset.templateId = d.templateId;
        host.hidden = false; host.classList.remove('hidden'); host.classList.add('is-shown');
        host.setAttribute('role', 'status');
    }
    function hideRemoteBanner() {
        const host = $('remote-banner'); if (!host || !host.classList.contains('is-shown')) return;
        host.classList.remove('is-shown'); host.hidden = true; host.innerHTML = '';
        delete host.dataset.templateId;
    }

    // ── Sincronización casi en vivo (§3.5) ────────────────────────────────────
    function scheduleNearLive() {
        clearTimeout(nearLiveTimer); nearLiveTimer = null;
        const t = activeTpl();
        if (!t || !t.id || !has('quickSave') || !othersOn(t.id).length || !call('isDirty')) return;
        if (Date.now() < nearLiveBlockedUntil) return;
        const wait = Math.max(200, NEAR_LIVE_IDLE_MS - (Date.now() - lastActivity));
        nearLiveTimer = setTimeout(() => {
            nearLiveTimer = null;
            const cur = activeTpl();
            if (!cur || cur.id !== t.id || !call('isDirty') || !othersOn(t.id).length || Date.now() < nearLiveBlockedUntil) return;
            if (isCellEditing() || Date.now() - lastActivity < NEAR_LIVE_IDLE_MS) { scheduleNearLive(); return; }
            Promise.resolve(call('quickSave'))
                .then(ok => { if (ok === false) nearLiveBlockedUntil = Date.now() + 30000; })   // conflicto/error: no insistir en bucle
                .catch(() => { nearLiveBlockedUntil = Date.now() + 30000; });
        }, wait);
    }

    // ── Render global ─────────────────────────────────────────────────────────
    function renderAll() {
        const parts = [renderIndicator, renderPresenceStack, renderHomeOnline, renderCardBadges, renderStatusBar, fillUsersPopover];
        for (const fn of parts) { try { fn(); } catch (e) { console.warn('[live] render', e); } }
        scheduleCursors();
    }

    // ── API pública ───────────────────────────────────────────────────────────
    ETB.live = {
        clients: () => clients.slice(),
        self: () => selfEntry(),
        byTemplate: id => { const l = clients.filter(c => c.templateId === id); return { viewers: l.filter(c => c.mode === 'view'), editors: l.filter(c => c.mode === 'edit') }; },
        status: () => (unsupported ? 'offline' : status),
        supported: () => !unsupported,
        editProfile: () => { const a = ind && ind.me; if (a && a.isConnected) openProfile(a); }
    };

    // ── Inicialización ────────────────────────────────────────────────────────
    let started = false;
    function init() {
        if (started) return; started = true;
        const on = (type, fn) => { if (has('on')) ETB.on(type, d => { try { fn(d || {}); } catch (e) { console.warn('[live]', type, e); } }); };
        on('view', () => {
            if (call('currentView') !== 'editor') { sel.cell = sel.range = null; hideRemoteBanner(); }
            schedulePresence(true); trackJoins(); renderAll();
        });
        on('doc:activated', () => { hideRemoteBanner(); sel.cell = sel.range = null; nearLiveBlockedUntil = 0; schedulePresence(true); trackJoins(); renderAll(); });
        on('selection', d => { sel.cell = d.cell || null; sel.range = d.range || null; schedulePresence(true); scheduleCursors(); });
        on('sheet:switched', () => { schedulePresence(true); scheduleCursors(); });
        on('activity', d => {
            const wasEdit = Date.now() - lastActivity < EDIT_WINDOW_MS;
            lastActivity = Date.now(); lastActivityTpl = d.templateId || null;
            if (!wasEdit) schedulePresence(true);
            clearTimeout(editEndTimer);
            editEndTimer = setTimeout(() => schedulePresence(), EDIT_WINDOW_MS + 100);   // vuelve a "viendo"
            scheduleNearLive();
        });
        on('saved', () => { clearTimeout(nearLiveTimer); nearLiveBlockedUntil = 0; });
        on('conflict', () => { nearLiveBlockedUntil = Date.now() + 60000; });
        on('grid:rendered', () => { observeGrid(); scheduleCursors(); });
        on('home:painted', () => { renderCardBadges(); renderHomeOnline(); });
        on('docs:changed', () => renderAll());
        on('live:presence', () => scheduleCursors());
        observeGrid();
        renderAll();
        if (needsWelcome) openWelcome();
        if (typeof EventSource !== 'function' || typeof fetch !== 'function') { degrade(); return; }
        claimSettled.then(() => { if (!es && !unsupported) connect(); });
        keepaliveTimer = setInterval(() => { if (!unsupported && status === 'online') sendPresence(); }, KEEPALIVE_MS);
        // Vigilante: el servidor envía ping cada 25 s; si no llega nada en ~70 s la conexión está muerta.
        watchdog = setInterval(() => {
            if (unsupported || status !== 'online' || !lastEventAt) return;
            if (Date.now() - lastEventAt > STALE_AFTER_MS) { setStatus('connecting'); connect(); }
        }, 5000);
        window.addEventListener('pagehide', () => {
            try {
                const blob = new Blob([JSON.stringify({ clientId })], { type: 'application/json' });
                if (navigator.sendBeacon) navigator.sendBeacon('/api/live/leave', blob);
            } catch (_) { }
            closeStream();
        });
        window.addEventListener('pageshow', e => { if (e.persisted && !unsupported) connect(); });
        document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && !unsupported) schedulePresence(true); });
        window.addEventListener('online', () => { if (!unsupported && status !== 'online') connect(); });
    }
    if (has('whenReady')) ETB.whenReady(init);
    else if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();
