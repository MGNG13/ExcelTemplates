'use strict';
/**
 * Library/live.js — Hub de tiempo real por Server-Sent Events (SSE).
 *
 * Mantiene la presencia de cada pestaña conectada (quién está viendo/editando qué plantilla),
 * difunde eventos (`presence`, `template:*`, `usage`, `ping`) y emite 'enter'/'leave' en Node
 * cuando un cliente empieza/deja de ver una plantilla. Ver CONTRACT.md §2.2.
 */
const { EventEmitter } = require('events');
const express = require('express');

const CLIENT_ID_RE = /^[A-Za-z0-9_-]{8,64}$/;
const USER_ID_RE = /^[A-Za-z0-9_-]{1,64}$/;
const SAFE_ID = /^[A-Za-z0-9_-]{1,64}$/;
const COLOR_RE = /^#[0-9a-fA-F]{6}$/;
const CELL_RE = /^[A-Z]{1,3}[0-9]{1,7}$/;
const RANGE_RE = /^[A-Z]{1,3}[0-9]{1,7}:[A-Z]{1,3}[0-9]{1,7}$/;
const EVENT_RE = /^[a-z:]+$/;
const MODES = new Set(['home', 'view', 'edit', 'dashboard']);
const PALETTE = ['#217346', '#2563eb', '#d97706', '#db2777', '#7c3aed', '#0891b2', '#dc2626', '#65a30d', '#ea580c', '#4f46e5', '#0d9488', '#9333ea'];
const BROADCAST_MS = 150;
const DEFAULT_NAME = 'Invitado';

// Caracteres de control, separadores de línea y overrides bidi; además < > como defensa en profundidad.
const STRIP_RE = /[\u0000-\u001F\u007F-\u009F\u200B-\u200F\u2028-\u202E\u2066-\u2069\uFEFF<>]/g;

/** Texto limpio: sin control chars, espacios colapsados, recortado a `max` code points. */
function cleanText(value, max) {
  if (typeof value !== 'string' && typeof value !== 'number') return '';
  const s = String(value).slice(0, max * 4).replace(STRIP_RE, '').replace(/\s+/g, ' ').trim();
  return Array.from(s).slice(0, max).join('').trim();
}
const sanitizeName = (v) => cleanText(v, 40) || DEFAULT_NAME;
const optText = (v, max) => (v == null ? null : cleanText(v, max) || null);
const matchOr = (v, re, fallback = null) => (typeof v === 'string' && re.test(v) ? v : fallback);

/** Color determinista de la paleta a partir de una semilla. */
function paletteColor(seed) {
  let h = 0;
  for (const ch of String(seed || '')) h = (Math.imul(h, 31) + ch.codePointAt(0)) | 0;
  return PALETTE[Math.abs(h) % PALETTE.length];
}

function createLiveHub(opts = {}) {
  const heartbeatMs = Number(opts.heartbeatMs) > 0 ? Number(opts.heartbeatMs) : 25000;
  const idleMs = Number(opts.idleMs) > 0 ? Number(opts.idleMs) : 120000;
  const maxClients = Number(opts.maxClients) > 0 ? Number(opts.maxClients) : 500;
  const maxPerIp = Number(opts.maxPerIp) > 0 ? Number(opts.maxPerIp) : 50;

  const hub = new EventEmitter();
  hub.setMaxListeners(50);
  /** clientId → { clientId, userId, res, ip, entry } */
  const conns = new Map();
  let closed = false;
  let broadcastTimer = null;
  let lastBroadcast = 0;
  let cachedRouter = null;

  const heartbeat = setInterval(() => hub.publish('ping', { serverTime: Date.now() }), heartbeatMs);
  heartbeat.unref?.();

  // ---- Utilidades internas ------------------------------------------------
  const clientInfo = (c) => ({ clientId: c.clientId, userId: c.entry.userId, name: c.entry.name });

  /** Vista pública de una entrada (idle calculado al leer). */
  function publicEntry(e, now = Date.now()) {
    return {
      clientId: e.clientId, userId: e.userId, name: e.name, color: e.color, mode: e.mode,
      templateId: e.templateId, templateName: e.templateName, sheet: e.sheet, cell: e.cell, range: e.range,
      connectedAt: e.connectedAt, lastActive: e.lastActive, idle: now - e.lastActive > idleMs,
    };
  }

  /** Escribe un evento SSE; si falla, descarta la conexión. */
  function send(conn, event, data) {
    const { res } = conn;
    if (res.writableEnded || res.destroyed) { drop(conn); return false; }
    try {
      res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
      res.flush?.();
      return true;
    } catch (_) {
      drop(conn);
      return false;
    }
  }

  /** Quita la conexión (si sigue siendo la vigente) y emite 'leave' si veía una plantilla. */
  function drop(conn) {
    if (conns.get(conn.clientId) !== conn) return;
    conns.delete(conn.clientId);
    try { if (!conn.res.writableEnded) conn.res.end(); } catch (_) { /* socket ya cerrado */ }
    if (conn.entry.templateId) safeEmit('leave', { templateId: conn.entry.templateId, client: clientInfo(conn) });
    scheduleBroadcast();
  }

  // Un listener que lanza no debe tumbar el hub.
  function safeEmit(name, payload) {
    try { hub.emit(name, payload); } catch (err) { console.error(`[live] listener '${name}' falló:`, err && err.message); }
  }

  function scheduleBroadcast() {
    if (closed || broadcastTimer) return;
    const wait = Math.max(0, lastBroadcast + BROADCAST_MS - Date.now());
    broadcastTimer = setTimeout(() => {
      broadcastTimer = null;
      lastBroadcast = Date.now();
      hub.publish('presence', { clients: hub.getPresence(), serverTime: lastBroadcast });
    }, wait);
    broadcastTimer.unref?.();
  }

  /** Aplica un PresenceUpdate parcial y emite enter/leave en transiciones de plantilla. */
  function applyUpdate(conn, u) {
    const e = conn.entry;
    const has = (k) => Object.prototype.hasOwnProperty.call(u, k);
    const prevTpl = e.templateId;
    const prevInfo = clientInfo(conn);

    if (has('userId')) { const uid = matchOr(u.userId, USER_ID_RE); if (uid) e.userId = conn.userId = uid; }
    if (has('name')) e.name = sanitizeName(u.name);
    if (has('color')) e.color = matchOr(u.color, COLOR_RE) || paletteColor(e.userId);
    if (has('mode') && MODES.has(u.mode)) e.mode = u.mode;
    if (has('templateId')) {
      e.templateId = matchOr(u.templateId, SAFE_ID);
      if (e.templateId !== prevTpl && !has('templateName')) e.templateName = null;
      if (e.templateId !== prevTpl) { e.sheet = e.cell = e.range = null; }
    }
    if (has('templateName')) e.templateName = optText(u.templateName, 120);
    if (has('sheet')) e.sheet = optText(u.sheet, 64);
    if (has('cell')) e.cell = matchOr(u.cell, CELL_RE);
    if (has('range')) e.range = matchOr(u.range, RANGE_RE);
    if (u.active === true) e.lastActive = Date.now();

    if (e.templateId !== prevTpl) {
      if (prevTpl) safeEmit('leave', { templateId: prevTpl, client: prevInfo });
      if (e.templateId) safeEmit('enter', { templateId: e.templateId, templateName: e.templateName, client: clientInfo(conn) });
    }
    scheduleBroadcast();
  }

  const jsonError = (res, status, error, code) => res.status(status).json({ error, code });
  const ipOf = (req) => req.ip || req.socket?.remoteAddress || 'unknown';

  // ---- Rutas -------------------------------------------------------------
  function handleStream(req, res) {
    if (closed) return jsonError(res, 503, 'Servicio en tiempo real no disponible', 'CLOSED');
    const clientId = matchOr(req.query.clientId, CLIENT_ID_RE);
    if (!clientId) return jsonError(res, 400, 'clientId inválido', 'BAD_CLIENT_ID');

    const old = conns.get(clientId);
    const ip = ipOf(req);
    if (!old && conns.size >= maxClients) return jsonError(res, 503, 'Demasiadas conexiones en tiempo real', 'TOO_MANY_CLIENTS');
    let perIp = 0;
    for (const c of conns.values()) if (c.ip === ip && c !== old) perIp++;
    if (perIp >= maxPerIp) return jsonError(res, 429, 'Demasiadas conexiones desde esta dirección', 'TOO_MANY_STREAMS');

    const now = Date.now();
    const userId = matchOr(req.query.userId, USER_ID_RE) || clientId;
    const entry = {
      clientId, userId,
      name: sanitizeName(req.query.name),
      color: matchOr(req.query.color, COLOR_RE) || paletteColor(userId),
      mode: 'home', templateId: null, templateName: null, sheet: null, cell: null, range: null,
      connectedAt: now, lastActive: now,
    };
    // Reconexión del mismo clientId: conservar dónde estaba para no emitir enter/leave espurios.
    if (old) {
      for (const k of ['mode', 'templateId', 'templateName', 'sheet', 'cell', 'range']) entry[k] = old.entry[k];
      conns.delete(clientId);
      try { old.res.end(); } catch (_) { /* ignorar */ }
    }
    const conn = { clientId, userId, res, ip, entry };
    conns.set(clientId, conn);

    req.socket?.setTimeout?.(0);
    req.socket?.setNoDelay?.(true);
    res.status(200).set({
      'Content-Type': 'text/event-stream; charset=utf-8',
      'Cache-Control': 'no-cache, no-transform',
      Connection: 'keep-alive',
      'X-Accel-Buffering': 'no',
    });
    res.flushHeaders?.();
    try { res.write('retry: 3000\n\n'); res.flush?.(); } catch (_) { drop(conn); return undefined; }

    const onClose = () => drop(conn);
    req.on('close', onClose);
    res.on('close', onClose);
    res.on('error', onClose);

    send(conn, 'hello', { clientId, serverTime: now, clients: hub.getPresence() });
    scheduleBroadcast();
    return undefined;
  }

  // ---- API pública -------------------------------------------------------
  hub.router = function router() {
    if (cachedRouter) return cachedRouter;
    const r = express.Router();
    r.get('/stream', handleStream);
    r.get('/presence', (req, res) => {
      res.set('Cache-Control', 'no-store').json({ clients: hub.getPresence(), serverTime: Date.now() });
    });
    const json = express.json({ limit: '16kb', type: ['application/json', 'text/plain'] });
    r.post('/presence', json, (req, res) => {
      const body = req.body && typeof req.body === 'object' && !Array.isArray(req.body) ? req.body : {};
      const conn = typeof body.clientId === 'string' ? conns.get(body.clientId) : null;
      if (!conn) return jsonError(res, 404, 'No hay conexión en tiempo real para este cliente', 'NO_STREAM');
      applyUpdate(conn, body);
      return res.status(204).end();
    });
    r.post('/leave', json, (req, res) => {
      const conn = req.body && typeof req.body.clientId === 'string' ? conns.get(req.body.clientId) : null;
      if (conn) applyUpdate(conn, { mode: 'home', templateId: null });
      res.status(204).end();
    });
    // Errores de parseo/tamaño del body → JSON sin reflejar la entrada.
    // eslint-disable-next-line no-unused-vars
    r.use((err, req, res, next) => {
      const status = err && err.status >= 400 && err.status < 500 ? err.status : 400;
      jsonError(res, status, status === 413 ? 'Cuerpo demasiado grande' : 'Cuerpo inválido', 'BAD_BODY');
    });
    cachedRouter = r;
    return r;
  };

  /** Difunde un evento a todos los streams. Devuelve el número de entregas. */
  hub.publish = function publish(event, data) {
    if (typeof event !== 'string' || !EVENT_RE.test(event)) throw new TypeError(`Nombre de evento inválido: ${String(event).slice(0, 40)}`);
    let n = 0;
    for (const conn of [...conns.values()]) if (send(conn, event, data === undefined ? null : data)) n++;
    return n;
  };

  hub.getPresence = function getPresence() {
    const now = Date.now();
    return [...conns.values()].map((c) => publicEntry(c.entry, now));
  };

  hub.countByTemplate = function countByTemplate() {
    const out = {};
    for (const { entry: e } of conns.values()) {
      if (!e.templateId) continue;
      const t = out[e.templateId] || (out[e.templateId] = { viewers: 0, editors: 0, users: [] });
      if (e.mode === 'edit') t.editors++; else t.viewers++;
      t.users.push({ userId: e.userId, name: e.name, color: e.color, mode: e.mode });
    }
    return out;
  };

  hub.clientFromRequest = function clientFromRequest(req) {
    const h = (k) => { const v = req && req.headers ? req.headers[k] : undefined; return Array.isArray(v) ? v[0] : v; };
    const clientId = matchOr(h('x-client-id'), CLIENT_ID_RE);
    const conn = clientId ? conns.get(clientId) : null;
    if (conn) return clientInfo(conn);
    const userId = matchOr(h('x-user-id'), USER_ID_RE);
    let rawName = h('x-client-name');
    if (typeof rawName === 'string') { try { rawName = decodeURIComponent(rawName); } catch (_) { /* usar tal cual */ } }
    const name = cleanText(rawName, 40);
    if (!clientId && !userId && !name) return null;
    return { clientId, userId, name: name || DEFAULT_NAME };
  };

  hub.close = function close() {
    closed = true;
    clearInterval(heartbeat);
    if (broadcastTimer) { clearTimeout(broadcastTimer); broadcastTimer = null; }
    const all = [...conns.values()];
    conns.clear();
    for (const c of all) { try { c.res.end(); } catch (_) { /* ignorar */ } }
  };

  return hub;
}

module.exports = { createLiveHub };
