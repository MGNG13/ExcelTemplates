'use strict';
/**
 * Library/usage.js — analítica de uso por plantilla (contrato §2.3).
 *
 * Modelo en memoria por plantilla + anillo global de eventos recientes.
 * Persistencia debounced (2 s) y atómica en <dataDir>/usage/usage.json.
 * Buckets en hora LOCAL del servidor: dayKey 'YYYY-MM-DD', hourKey 'YYYY-MM-DDTHH'.
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const SAFE_ID = /^[A-Za-z0-9_-]{1,64}$/;
const TYPES = new Set(['render', 'batch', 'view', 'save', 'create', 'import', 'duplicate', 'delete']);
const SOURCES = new Set(['ui', 'api', 'link']);
const RING_MAX = 1000;
const CLIENTS_MAX = 50;
const HOUR_RETENTION_MS = 8 * 24 * 3600e3;
const DAY_RETENTION_MS = 120 * 24 * 3600e3;
const PERSIST_DELAY_MS = 2000;
const FILE_VERSION = 1;

// ---------- utilidades de tiempo (hora local) ----------
const pad = (n) => String(n).padStart(2, '0');
const dayKey = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const hourKey = (d) => `${dayKey(d)}T${pad(d.getHours())}`;
function parseHourKey(k) {
  const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2})$/.exec(k);
  return m ? new Date(+m[1], +m[2] - 1, +m[3], +m[4]).getTime() : NaN;
}
function parseDayKey(k) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(k);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]).getTime() : NaN;
}
/** Últimas n horas (viejo→nuevo) terminando en la hora de `now`. */
function lastHours(now, n) {
  const base = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours());
  const out = [];
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(base);
    d.setHours(d.getHours() - i); // setHours maneja DST y cambios de día
    out.push(hourKey(d));
  }
  return out;
}
/** Últimos n días (viejo→nuevo) terminando hoy. */
function lastDays(now, n) {
  const out = [];
  for (let i = n - 1; i >= 0; i--) {
    out.push(dayKey(new Date(now.getFullYear(), now.getMonth(), now.getDate() - i)));
  }
  return out;
}

// ---------- normalización ----------
function cleanStr(v, max) {
  if (v === undefined || v === null) return undefined;
  // eslint-disable-next-line no-control-regex
  const s = String(v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim();
  return s ? s.slice(0, max) : undefined;
}
function nonNegNum(v, max = Number.MAX_SAFE_INTEGER) {
  const n = Number(v);
  return Number.isFinite(n) && n >= 0 ? Math.min(n, max) : undefined;
}
const num = (v) => (Number.isFinite(Number(v)) ? Number(v) : 0);

function blankCounts() { return { renders: 0, errors: 0, views: 0 }; }
function blankTemplate(id) {
  return {
    id,
    name: '',
    deleted: false,
    totals: { renders: 0, errors: 0, views: 0, saves: 0, bytes: 0, totalMs: 0, msCount: 0 },
    last: { renderAt: null, viewAt: null, saveAt: null },
    hourly: {},
    daily: {},
    sources: { ui: 0, api: 0, link: 0 },
    clients: {},
  };
}
/** Repara una entrada leída del disco (tolerante a campos faltantes). */
function reviveTemplate(id, raw) {
  const t = blankTemplate(id);
  if (!raw || typeof raw !== 'object') return t;
  t.name = typeof raw.name === 'string' ? raw.name.slice(0, 120) : '';
  t.deleted = !!raw.deleted;
  for (const k of Object.keys(t.totals)) t.totals[k] = num(raw.totals && raw.totals[k]);
  for (const k of Object.keys(t.last)) t.last[k] = raw.last && typeof raw.last[k] === 'string' ? raw.last[k] : null;
  for (const [field, parse] of [['hourly', parseHourKey], ['daily', parseDayKey]]) {
    const src = raw[field] && typeof raw[field] === 'object' ? raw[field] : {};
    for (const [k, v] of Object.entries(src)) {
      if (!Number.isFinite(parse(k)) || !v) continue;
      t[field][k] = { renders: num(v.renders), errors: num(v.errors), views: num(v.views) };
    }
  }
  for (const k of SOURCES) t.sources[k] = num(raw.sources && raw.sources[k]);
  if (raw.clients && typeof raw.clients === 'object') {
    for (const [k, v] of Object.entries(raw.clients)) t.clients[String(k).slice(0, 80)] = num(v);
  }
  return t;
}

// ---------- escritura atómica (con reintentos para Windows: EPERM/EBUSY) ----------
async function atomicWrite(file, data) {
  const tmp = `${file}.${process.pid}.${Date.now()}.tmp`;
  await fs.promises.writeFile(tmp, data, 'utf8');
  for (let attempt = 0; ; attempt++) {
    try {
      await fs.promises.rename(tmp, file);
      return;
    } catch (err) {
      if (attempt >= 5 || !['EPERM', 'EBUSY', 'EACCES'].includes(err.code)) {
        await fs.promises.unlink(tmp).catch(() => {});
        throw err;
      }
      await new Promise((r) => setTimeout(r, 50 * (attempt + 1)));
    }
  }
}

/**
 * @param {object} opts
 * @param {string} opts.dataDir  raíz de datos
 * @param {object} [opts.live]   hub SSE opcional (live.publish)
 * @param {() => number} [opts.now] reloj inyectable (pruebas)
 * @param {number} [opts.persistDelayMs]
 */
function createUsageTracker({ dataDir, live = null, now: clock = Date.now, persistDelayMs = PERSIST_DELAY_MS } = {}) {
  if (!dataDir) throw new Error('createUsageTracker: dataDir requerido');
  const dir = path.join(dataDir, 'usage');
  const file = path.join(dir, 'usage.json');

  /** @type {Map<string, ReturnType<typeof blankTemplate>>} */
  const templates = new Map();
  let recent = []; // anillo global (viejo→nuevo)
  const dedupe = new Map(); // dedupeKey → expira (ms)

  // ---- carga síncrona ----
  try {
    fs.mkdirSync(dir, { recursive: true });
    if (fs.existsSync(file)) {
      const text = fs.readFileSync(file, 'utf8');
      try {
        const data = JSON.parse(text);
        if (!data || typeof data !== 'object') throw new Error('formato inválido');
        for (const [id, raw] of Object.entries(data.templates || {})) {
          if (SAFE_ID.test(id)) templates.set(id, reviveTemplate(id, raw));
        }
        if (Array.isArray(data.recent)) {
          recent = data.recent.filter((e) => e && typeof e === 'object' && typeof e.ts === 'string').slice(-RING_MAX);
        }
      } catch (err) {
        const bad = `${file}.corrupt-${Date.now()}`;
        try { fs.renameSync(file, bad); } catch (_) { /* ignorar */ }
        console.warn(`[usage] usage.json corrupto; renombrado a ${path.basename(bad)}. Se inicia vacío.`);
      }
    }
  } catch (err) {
    console.warn('[usage] no se pudo cargar el historial de uso:', err.message);
  }

  // ---- persistencia ----
  let timer = null;
  let dirty = false;
  let writing = Promise.resolve();

  function serialize() {
    const out = { version: FILE_VERSION, savedAt: new Date(clock()).toISOString(), templates: {}, recent };
    for (const [id, t] of templates) out.templates[id] = t;
    return JSON.stringify(out);
  }
  function writeNow() {
    dirty = false;
    const payload = serialize();
    writing = writing
      .then(() => fs.promises.mkdir(dir, { recursive: true }))
      .then(() => atomicWrite(file, payload))
      .catch((err) => console.warn('[usage] error al guardar usage.json:', err.message));
    return writing;
  }
  function schedule() {
    dirty = true;
    if (timer) return;
    timer = setTimeout(() => { timer = null; if (dirty) writeNow(); }, persistDelayMs);
    if (timer.unref) timer.unref();
  }
  async function flush() {
    if (timer) { clearTimeout(timer); timer = null; }
    if (dirty) await writeNow();
    else await writing;
  }

  // ---- helpers ----
  function getT(id) {
    let t = templates.get(id);
    if (!t) { t = blankTemplate(id); templates.set(id, t); }
    return t;
  }
  function bump(map, key, field, n) {
    const b = map[key] || (map[key] = blankCounts());
    b[field] += n;
  }
  function prune(nowMs) {
    for (const t of templates.values()) {
      for (const k of Object.keys(t.hourly)) if (!(nowMs - parseHourKey(k) <= HOUR_RETENTION_MS)) delete t.hourly[k];
      for (const k of Object.keys(t.daily)) if (!(nowMs - parseDayKey(k) <= DAY_RETENTION_MS)) delete t.daily[k];
    }
    for (const [k, exp] of dedupe) if (exp <= nowMs) dedupe.delete(k);
  }
  function pruneClients(t) {
    const entries = Object.entries(t.clients);
    if (entries.length <= CLIENTS_MAX) return;
    entries.sort((a, b) => b[1] - a[1]);
    t.clients = Object.fromEntries(entries.slice(0, CLIENTS_MAX));
  }

  function normalize(evt) {
    if (!evt || typeof evt !== 'object') return null;
    const type = String(evt.type || '');
    if (!TYPES.has(type)) return null;
    const templateId = String(evt.templateId || '');
    if (!SAFE_ID.test(templateId)) return null;
    const e = { type, templateId, ok: evt.ok !== false };
    const name = cleanStr(evt.templateName, 120); if (name) e.templateName = name;
    const ms = nonNegNum(evt.ms, 36e5); if (ms !== undefined) e.ms = Math.round(ms);
    const bytes = nonNegNum(evt.bytes, 1e12); if (bytes !== undefined) e.bytes = Math.round(bytes);
    if (type === 'batch') e.count = Math.max(1, Math.round(nonNegNum(evt.count, 100000) || 1));
    if (SOURCES.has(evt.source)) e.source = evt.source;
    const client = cleanStr(evt.client, 80); if (client) e.client = client;
    const error = cleanStr(evt.error, 300); if (error) e.error = error;
    return e;
  }

  // ---- vistas ----
  function summaryOf(t, nowMs) {
    const now = new Date(nowMs);
    const today = t.daily[dayKey(now)] || blankCounts();
    return {
      id: t.id,
      name: t.name || t.id,
      deleted: t.deleted,
      totals: {
        renders: t.totals.renders,
        errors: t.totals.errors,
        views: t.totals.views,
        saves: t.totals.saves,
        bytes: t.totals.bytes,
        avgMs: t.totals.msCount ? Math.round(t.totals.totalMs / t.totals.msCount) : 0,
      },
      today: { renders: today.renders, errors: today.errors, views: today.views },
      last: { ...t.last },
      spark24h: lastHours(now, 24).map((k) => (t.hourly[k] ? t.hourly[k].renders : 0)),
      spark7d: lastDays(now, 7).map((k) => (t.daily[k] ? t.daily[k].renders : 0)),
    };
  }
  function globalTotals(nowMs) {
    const tk = dayKey(new Date(nowMs));
    const g = { renders: 0, errors: 0, views: 0, saves: 0, bytes: 0, today: blankCounts(), rendersLastHour: 0, templatesUsed: 0 };
    for (const t of templates.values()) {
      g.renders += t.totals.renders; g.errors += t.totals.errors; g.views += t.totals.views;
      g.saves += t.totals.saves; g.bytes += t.totals.bytes;
      if (t.totals.renders > 0) g.templatesUsed++;
      const d = t.daily[tk];
      if (d) { g.today.renders += d.renders; g.today.errors += d.errors; g.today.views += d.views; }
    }
    const cutoff = nowMs - 3600e3;
    for (let i = recent.length - 1; i >= 0; i--) {
      const e = recent[i];
      const ts = Date.parse(e.ts);
      if (!(ts > cutoff)) continue; // el anillo puede contener eventos fuera de orden si se inyecta "now"
      if (e.ok === false) continue;
      if (e.type === 'render') g.rendersLastHour += 1;
      else if (e.type === 'batch') g.rendersLastHour += e.count || 1;
    }
    return g;
  }

  // ---- API pública ----
  function record(evt, { dedupeKey, dedupeMs, now: nowOverride } = {}) {
    const e = normalize(evt);
    if (!e) return null;
    const nowMs = Number.isFinite(nowOverride) ? nowOverride : clock();
    if (dedupeKey && dedupeMs > 0) {
      const k = String(dedupeKey);
      const exp = dedupe.get(k);
      if (exp && exp > nowMs) return null;
      dedupe.set(k, nowMs + Number(dedupeMs));
    }
    const d = new Date(nowMs);
    const ts = d.toISOString();
    const event = { id: crypto.randomUUID(), ts, ...e };
    const t = getT(e.templateId);
    if (e.templateName) t.name = e.templateName;
    const hk = hourKey(d), dk = dayKey(d);

    switch (e.type) {
      case 'render':
      case 'batch': {
        const n = e.type === 'batch' ? e.count : 1;
        if (e.ok) {
          t.totals.renders += n;
          bump(t.hourly, hk, 'renders', n); bump(t.daily, dk, 'renders', n);
          t.last.renderAt = ts;
          if (e.source) t.sources[e.source] += n;
          if (e.client) { t.clients[e.client] = (t.clients[e.client] || 0) + n; pruneClients(t); }
        } else {
          t.totals.errors += 1;
          bump(t.hourly, hk, 'errors', 1); bump(t.daily, dk, 'errors', 1);
        }
        if (e.bytes) t.totals.bytes += e.bytes;
        if (e.ms !== undefined) { t.totals.totalMs += e.ms; t.totals.msCount += 1; }
        break;
      }
      case 'view':
        t.totals.views += 1;
        bump(t.hourly, hk, 'views', 1); bump(t.daily, dk, 'views', 1);
        t.last.viewAt = ts;
        break;
      case 'save':
      case 'create':
        t.totals.saves += 1;
        t.last.saveAt = ts;
        if (e.type === 'create') t.deleted = false;
        break;
      case 'delete':
        t.deleted = true;
        break;
      default: // import, duplicate: solo se registran en el anillo
        if (e.type === 'import' || e.type === 'duplicate') t.deleted = false;
    }

    recent.push(event);
    if (recent.length > RING_MAX) recent.splice(0, recent.length - RING_MAX);
    prune(nowMs);
    schedule();

    if (live && typeof live.publish === 'function') {
      try {
        live.publish('usage', { event, template: summaryOf(t, nowMs), totals: globalTotals(nowMs) });
      } catch (err) {
        console.warn('[usage] live.publish falló:', err.message);
      }
    }
    return event;
  }

  function summary() {
    const nowMs = clock();
    const list = [...templates.values()].map((t) => summaryOf(t, nowMs));
    list.sort((a, b) => b.today.renders - a.today.renders || b.totals.renders - a.totals.renders
      || String(a.name).localeCompare(String(b.name), 'es'));
    return {
      generatedAt: new Date(nowMs).toISOString(),
      totals: globalTotals(nowMs),
      templates: list,
      recent: recent.slice(-50).reverse(),
    };
  }

  function detail(id) {
    const nowMs = clock();
    const now = new Date(nowMs);
    const known = SAFE_ID.test(String(id)) ? templates.get(String(id)) : null;
    const t = known || blankTemplate(String(id));
    const base = summaryOf(t, nowMs);
    return {
      ...base,
      hourly48: lastHours(now, 48).map((hour) => ({ hour, ...(t.hourly[hour] ? { ...t.hourly[hour] } : blankCounts()) })),
      daily30: lastDays(now, 30).map((day) => ({ day, ...(t.daily[day] ? { ...t.daily[day] } : blankCounts()) })),
      sources: { ...t.sources },
      topClients: Object.entries(t.clients).sort((a, b) => b[1] - a[1]).slice(0, 10)
        .map(([client, renders]) => ({ client, renders })),
      recent: known ? recent.filter((e) => e.templateId === t.id).slice(-100).reverse() : [],
    };
  }

  function rename(id, name) {
    if (!SAFE_ID.test(String(id))) return false;
    const n = cleanStr(name, 120);
    if (!n) return false;
    getT(String(id)).name = n;
    schedule();
    return true;
  }

  function markDeleted(id) {
    if (!SAFE_ID.test(String(id))) return false;
    getT(String(id)).deleted = true;
    schedule();
    return true;
  }

  function router() {
    // express se requiere de forma perezosa para que el módulo funcione sin él en pruebas unitarias
    const express = require('express');
    const r = express.Router();
    r.use((req, res, next) => { res.set('Cache-Control', 'no-store'); next(); });
    r.get('/', (req, res) => res.json(summary()));
    r.get('/events', (req, res) => {
      let limit = parseInt(req.query.limit, 10);
      if (!Number.isFinite(limit)) limit = 100;
      limit = Math.min(1000, Math.max(1, limit));
      const tid = typeof req.query.templateId === 'string' && req.query.templateId ? req.query.templateId : null;
      if (tid && !SAFE_ID.test(tid)) return res.status(400).json({ error: 'templateId inválido', code: 'BAD_ID' });
      const src = tid ? recent.filter((e) => e.templateId === tid) : recent;
      res.json({ events: src.slice(-limit).reverse() });
    });
    r.get('/:id', (req, res) => {
      if (!SAFE_ID.test(req.params.id)) return res.status(400).json({ error: 'ID de plantilla inválido', code: 'BAD_ID' });
      res.json(detail(req.params.id));
    });
    return r;
  }

  return { record, summary, detail, rename, markDeleted, router, flush };
}

module.exports = { createUsageTracker };
