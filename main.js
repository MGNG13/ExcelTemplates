// Excel Template Builder v12 — servidor HTTP (wiring delgado).
// Persistencia: Library/store.js · Render: Library/render.js · Import: Library/importer.js
// Tiempo real: Library/live.js · Uso: Library/usage.js · Automatización: Library/automation.js
'use strict';
const express = require('express');
const multer = require('multer');
const path = require('path');
const cors = require('cors');
const helmet = require('helmet');
const compression = require('compression');
const rateLimit = require('express-rate-limit');
const AdmZip = require('adm-zip');

const { createStore, SAFE_ID } = require('./Library/store');
const { renderWorkbook } = require('./Library/render');
const { normalizeZones, zoneOfCell } = require('./Library/repeat-zones');
const { importWorkbook } = require('./Library/importer');
const pkg = require('./package.json');

// Peer modules are loaded defensively: a missing/broken module degrades its feature, not the server.
function optionalRequire(rel) {
  try { return require(rel); }
  catch (e) { console.warn(`⚠️  Módulo ${rel} no disponible: ${e.message}`); return null; }
}
const liveMod = optionalRequire('./Library/live');
const usageMod = optionalRequire('./Library/usage');
const automation = optionalRequire('./Library/automation') || {};

const PORT = process.argv[2] || process.env.PORT || 3001;
const HOST = process.env.HOST || '0.0.0.0';
const DATA_ROOT = process.env.ETB_DATA_DIR || path.join(__dirname, 'data');
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const BATCH_MAX = 200;

const store = createStore({ dataDir: DATA_ROOT });
const live = liveMod?.createLiveHub
  ? liveMod.createLiveHub({ heartbeatMs: 25000, idleMs: 120000, maxClients: 500 })
  : null;
const usage = usageMod?.createUsageTracker ? usageMod.createUsageTracker({ dataDir: DATA_ROOT, live }) : null;

// ── Small helpers around peers ────────────────────────────────────────────────
const clientOf = req => { try { return live?.clientFromRequest(req) || null; } catch { return null; } };
const publish = (evt, data) => { try { live?.publish(evt, data); } catch (e) { console.warn('live.publish:', e.message); } };
const record = (evt, opts) => { try { usage?.record(evt, opts); } catch (e) { console.warn('usage.record:', e.message); } };
const truncate = (s, n = 200) => String(s ?? '').slice(0, n);

function safeFileName(name, fallback = 'resultado') {
  if (typeof automation.safeFileName === 'function') return automation.safeFileName(name, fallback);
  // Fallback compatible con §2.4
  let n = String(name || '').replace(/[\/\\:*?"<>|\x00-\x1f\x7f]/g, '').trim().replace(/\.xlsx$/i, '');
  if (!n) n = fallback;
  return n.slice(0, 115) + '.xlsx';
}
const contentDisposition = fn => {
  const ascii = String(fn).normalize('NFKD').replace(/[̀-ͯ]/g, '').replace(/[^ -~]|["\%;]/g, '_');
  return `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(fn)}`;
};
const isPlainObject = v => v !== null && typeof v === 'object' && !Array.isArray(v);
const httpError = (status, error, extra) => Object.assign(new Error(error), { status, expose: true, extra });

if (live && usage && typeof live.on === 'function') {
  // A client starts viewing a template → one "view" per user/template every 10 min
  live.on('enter', ({ templateId, templateName, client } = {}) => {
    if (!templateId) return;
    const userKey = client?.userId || client?.clientId || 'anon';
    record({ type: 'view', templateId, templateName, source: 'ui', client: client?.name || undefined },
      { dedupeKey: userKey + templateId, dedupeMs: 600000 });
  });
}

// ── App & middleware ──────────────────────────────────────────────────────────
const app = express();
app.disable('x-powered-by');
app.set('trust proxy', process.env.TRUST_PROXY || 'loopback');
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'", "'unsafe-inline'", 'https://cdn.jsdelivr.net'],
      scriptSrcAttr: ["'none'"],
      styleSrc: ["'self'", "'unsafe-inline'", 'https://fonts.googleapis.com'],
      fontSrc: ["'self'", 'https://fonts.gstatic.com', 'data:'],
      imgSrc: ["'self'", 'data:', 'blob:'],
      connectSrc: ["'self'"],
      objectSrc: ["'none'"],
      frameAncestors: ["'self'"],
    },
  },
  crossOriginEmbedderPolicy: false,
}));
// Never compress SSE: it would buffer events
app.use(compression({
  filter: (req, res) => {
    if (req.path.startsWith('/api/live/stream')) return false;
    if (String(res.getHeader('Content-Type') || '').includes('text/event-stream')) return false;
    return compression.filter(req, res);
  },
}));
const allowedOrigins = (process.env.CORS_ORIGINS || '').split(',').map(o => o.trim()).filter(Boolean);
app.use(cors({ origin: allowedOrigins.length ? allowedOrigins : false, exposedHeaders: ['Content-Disposition', 'X-Render-Ms', 'X-Render-Warnings'] }));

const limitMsg = { error: 'Demasiadas solicitudes, intenta de nuevo en un momento', code: 'RATE_LIMIT' };
app.use('/api', rateLimit({ windowMs: 60000, limit: 1200, standardHeaders: 'draft-7', legacyHeaders: false,
  message: limitMsg, skip: req => req.path.startsWith('/live') }));
app.use('/api/render', rateLimit({ windowMs: 60000, limit: 120, standardHeaders: 'draft-7', legacyHeaders: false, message: limitMsg }));

// Live router has its own small body parser → mount before the 50 MB JSON parser
if (live) app.use('/api/live', live.router());
else app.use('/api/live', (req, res) => res.status(503).json({ error: 'Tiempo real no disponible', code: 'LIVE_UNAVAILABLE' }));

app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'Public'), { dotfiles: 'ignore', index: 'index.html' }));

// Path traversal guard: template ids are alphanumeric / dash / underscore
app.param('id', (req, res, next, id) => {
  if (!SAFE_ID.test(id)) return res.status(400).json({ error: 'ID inválido' });
  next();
});

// defParamCharset utf8: los navegadores envían el nombre del archivo en UTF-8 sin declararlo (si no, «Cotización» llega como «CotizaciÃ³n»).
const upload = multer({ storage: multer.memoryStorage(), defParamCharset: 'utf8', limits: { fileSize: 50 * 1024 * 1024, files: 1 } });
const wrap = fn => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

async function loadTemplate(id) {
  const t = await store.get(id);
  if (!t) throw httpError(404, 'Plantilla no encontrada');
  return t;
}

// ── Templates ─────────────────────────────────────────────────────────────────
app.get('/api/health', (_, res) => res.json({ ok: true, version: pkg.version, uptime: Math.round(process.uptime()) }));

app.get('/api/templates', wrap(async (_, res) => res.json(await store.list())));

app.get('/api/templates/:id', wrap(async (req, res) => res.json(await loadTemplate(req.params.id))));

app.post('/api/templates', wrap(async (req, res) => {
  const t = req.body;
  if (!isPlainObject(t) || !Array.isArray(t.sheets)) return res.status(400).json({ error: 'Datos de plantilla inválidos' });
  if (t.id != null && t.id !== '' && !SAFE_ID.test(String(t.id))) return res.status(400).json({ error: 'ID inválido' });
  const by = clientOf(req);
  const force = req.query.force === '1' || req.query.force === 'true';
  const baseRevision = typeof t.revision === 'number' ? t.revision : null;
  const r = await store.save(t, { baseRevision, force, by });
  if (!r.ok) {
    return res.status(409).json({ error: 'La plantilla fue modificada por otra persona', code: 'CONFLICT', current: r.current });
  }
  const saved = r.template;
  // New templates announce "created"; updates announce "saved" (one event per write)
  if (r.created) publish('template:created', { templateId: saved.id, name: saved.name, by });
  else publish('template:saved', { templateId: saved.id, revision: saved.revision, name: saved.name, updatedAt: saved.updatedAt, by });
  record({ type: r.created ? 'create' : 'save', templateId: saved.id, templateName: saved.name,
    source: req.get('X-Client-Id') ? 'ui' : 'api', client: by?.name || req.ip });
  try { usage?.rename(saved.id, saved.name); } catch { /* optional */ }
  res.json({ success: true, id: saved.id, revision: saved.revision, updatedAt: saved.updatedAt });
}));

app.put('/api/templates/:id/thumbnail', wrap(async (req, res) => {
  const thumb = req.body?.thumbnail || null;   // empty → clear (v11 behaviour)
  const ok = await store.setThumbnail(req.params.id, thumb);
  if (ok === null) return res.status(400).json({ error: 'Miniatura inválida (PNG/JPEG/WebP en base64, máx. 2 MB)' });
  if (!ok) return res.status(404).json({ error: 'No encontrado' });
  res.json({ success: true });
}));

app.post('/api/templates/:id/duplicate', wrap(async (req, res) => {
  const by = clientOf(req);
  const copy = await store.duplicate(req.params.id, { by });
  if (!copy) return res.status(404).json({ error: 'No encontrado' });
  publish('template:created', { templateId: copy.id, name: copy.name, by });
  record({ type: 'duplicate', templateId: copy.id, templateName: copy.name, client: by?.name || req.ip });
  res.json({ success: true, id: copy.id, name: copy.name });
}));

app.delete('/api/templates/:id', wrap(async (req, res) => {
  const by = clientOf(req);
  const existing = await store.get(req.params.id);
  if (!(await store.remove(req.params.id))) return res.status(404).json({ error: 'No encontrado' });
  publish('template:deleted', { templateId: req.params.id, by });
  record({ type: 'delete', templateId: req.params.id, templateName: existing?.name, client: by?.name || req.ip });
  try { usage?.markDeleted(req.params.id); } catch { /* optional */ }
  res.json({ success: true });
}));

app.get('/api/templates/:id/schema', wrap(async (req, res) => {
  const t = await loadTemplate(req.params.id);
  const schema = (t.sheets || []).map(s => ({
    name: s.name, isRepeatable: !!s.isRepeatable,
    repeatVariable: s.repeatVariable || null, rowTemplate: s.rowTemplate || null,
    variables: Object.entries(s.cells || {}).filter(([, c]) => c && c.variable).map(([addr, c]) => ({ addr, ...c.variable })),
  }));
  let extra = {};
  if (typeof automation.buildSchema === 'function') {
    try {
      const b = automation.buildSchema(t);
      // Additive: implicit variables (e.g. {{var}} in text), repeat zones and per-variable zone from automation
      extra.sheets = b.sheets;
      for (const sh of schema) {
        const bs = (b.sheets || []).find(x => x.name === sh.name);
        if (!bs) continue;
        sh.rowTemplate = bs.rowTemplate;
        sh.repeatZones = bs.repeatZones || [];
        const known = new Set(sh.variables.map(v => v.name));
        const tSheet = (t.sheets || []).find(x => x && x.name === sh.name);
        const zones = tSheet ? normalizeZones(tSheet).zones : [];
        for (const v of sh.variables) {
          const z = zoneOfCell(zones, v.addr);
          v.zone = z ? { arrayVariable: z.arrayVariable, direction: z.direction, id: z.id } : null;
          v.inRowTemplate = !!z;
        }
        for (const v of bs.variables) if (!known.has(v.name)) sh.variables.push({ ...v, implicit: true });
      }
    } catch (e) { console.warn('buildSchema:', e.message); }
  }
  res.json({ id: t.id, name: t.name, revision: t.revision || 0, schema, ...extra });
}));

function needAutomation(fn) {
  if (typeof automation[fn] !== 'function') throw httpError(503, 'Módulo de automatización no disponible');
  return automation[fn];
}

app.get('/api/templates/:id/example', wrap(async (req, res) => {
  res.json(needAutomation('buildExample')(await loadTemplate(req.params.id)));
}));

app.post('/api/templates/:id/validate', wrap(async (req, res) => {
  const t = await loadTemplate(req.params.id);
  res.json(needAutomation('validateBody')(t, req.body ?? {}));
}));

app.get('/api/openapi.json', (req, res, next) => {
  try { res.json(needAutomation('buildOpenApi')({ baseUrl: `${req.protocol}://${req.get('host')}` })); }
  catch (e) { next(e); }
});

// ── Import ────────────────────────────────────────────────────────────────────
app.post('/api/import', upload.single('file'), wrap(async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No se recibió archivo' });
  const t0 = Date.now();
  const by = clientOf(req);
  let template;
  try { template = await importWorkbook(req.file.buffer, req.file.originalname); }
  catch (e) {
    console.error('import error:', e);
    return res.status(400).json({ error: 'No se pudo leer el archivo Excel: ' + truncate(e.message) });
  }
  // Import events are not recorded: the template has no id until saved (the save records 'create').
  res.json(template);
}));

// ── Render ────────────────────────────────────────────────────────────────────
function renderSource(req) {
  if (req.method === 'GET') return 'link';
  return req.get('X-Client-Id') ? 'ui' : 'api';
}
function parseFormat(req) {
  const f = String(req.query.format || 'xlsx').toLowerCase();
  if (f !== 'xlsx' && f !== 'base64') throw httpError(400, 'Formato inválido (usa xlsx o base64)');
  return f;
}
function strictCheck(req, t, body) {
  const strict = req.query.strict === '1' || req.query.strict === 'true';
  if (!strict) return;
  const v = needAutomation('validateBody')(t, body);
  if (!v.valid) throw httpError(422, 'Los datos no cumplen el esquema de la plantilla', { code: 'VALIDATION', errors: v.errors, warnings: v.warnings });
}

async function handleRender(req, res, body, preloaded = null) {
  const t0 = Date.now();
  const source = renderSource(req);
  const client = clientOf(req)?.name || req.ip;
  let t = preloaded;
  try {
    const format = parseFormat(req);
    t ||= await loadTemplate(req.params.id);
    if (!isPlainObject(body)) throw httpError(400, 'El cuerpo debe ser un objeto JSON con datos por hoja');
    strictCheck(req, t, body);
    const warnings = [];
    const buf = await renderWorkbook(t, body, { warnings });
    const ms = Date.now() - t0;
    if (warnings.length) res.setHeader('X-Render-Warnings', encodeURIComponent(JSON.stringify(warnings.slice(0, 20).map(w => w.message))));
    const pattern = typeof req.query.filename === 'string' && req.query.filename ? req.query.filename : '';
    const filled = pattern && typeof automation.fillPattern === 'function' ? automation.fillPattern(pattern, body) : pattern;
    const fileName = safeFileName(filled || t.name, 'plantilla');
    record({ type: 'render', templateId: t.id, templateName: t.name, ok: true, ms, bytes: buf.length, source, client });
    res.setHeader('X-Render-Ms', String(ms));
    if (format === 'base64') return res.json({ fileName, mimeType: XLSX_MIME, size: buf.length, base64: buf.toString('base64'), ...(warnings.length ? { warnings } : {}) });
    res.setHeader('Content-Type', XLSX_MIME);
    res.setHeader('Content-Disposition', contentDisposition(fileName));
    res.send(buf);
  } catch (e) {
    if (t) record({ type: 'render', templateId: t.id, templateName: t.name, ok: false, ms: Date.now() - t0,
      source, client, error: truncate(e.message) });
    throw e;
  }
}

app.post('/api/render/:id', wrap((req, res) => handleRender(req, res, req.body ?? {})));

app.get('/api/render/:id', wrap(async (req, res) => {
  const t = await loadTemplate(req.params.id);
  const body = typeof automation.bodyFromQuery === 'function' ? automation.bodyFromQuery(t, req.query) : {};
  return handleRender(req, res, body, t);
}));

app.post('/api/render/:id/batch', wrap(async (req, res) => {
  const t0 = Date.now();
  const source = renderSource(req);
  const client = clientOf(req)?.name || req.ip;
  const format = parseFormat(req);
  const t = await loadTemplate(req.params.id);
  const items = Array.isArray(req.body) ? req.body : req.body?.items;
  const pattern = Array.isArray(req.body) ? null : req.body?.fileName;
  if (!Array.isArray(items) || items.length < 1 || items.length > BATCH_MAX)
    throw httpError(400, `Se requiere "items" con entre 1 y ${BATCH_MAX} elementos`);
  const bad = items.findIndex(it => !isPlainObject(it));
  if (bad >= 0) throw httpError(400, `El elemento ${bad + 1} no es un objeto JSON`);
  try {
    if (req.query.strict === '1' || req.query.strict === 'true') {
      const validate = needAutomation('validateBody');
      const failures = items.map((it, i) => ({ i, v: validate(t, it) })).filter(x => !x.v.valid);
      if (failures.length) throw httpError(422, 'Algunos elementos no cumplen el esquema', {
        code: 'VALIDATION', errors: failures.flatMap(f => f.v.errors.map(er => ({ ...er, item: f.i }))), warnings: [] });
    }
    const zip = new AdmZip();
    const used = new Set();
    for (let i = 0; i < items.length; i++) {
      const buf = await renderWorkbook(t, items[i]);
      let base = typeof pattern === 'string' && pattern && typeof automation.fillPattern === 'function'
        ? automation.fillPattern(pattern, items[i]) : `${t.name || 'plantilla'} ${i + 1}`;
      let name = safeFileName(base, `resultado ${i + 1}`);
      for (let k = 2; used.has(name.toLowerCase()); k++) name = safeFileName(`${base} (${k})`, `resultado ${i + 1}`);
      used.add(name.toLowerCase());
      zip.addFile(name, buf);
    }
    const out = zip.toBuffer();
    const ms = Date.now() - t0;
    const fileName = safeFileName(req.query.filename || t.name, 'plantilla').replace(/\.xlsx$/i, '.zip');
    record({ type: 'batch', templateId: t.id, templateName: t.name, ok: true, ms, bytes: out.length, count: items.length, source, client });
    res.setHeader('X-Render-Ms', String(ms));
    if (format === 'base64') return res.json({ fileName, mimeType: 'application/zip', size: out.length, base64: out.toString('base64'), count: items.length });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', contentDisposition(fileName));
    res.send(out);
  } catch (e) {
    record({ type: 'batch', templateId: t.id, templateName: t.name, ok: false, ms: Date.now() - t0, count: items.length,
      source, client, error: truncate(e.message) });
    throw e;
  }
}));

// ── Usage ─────────────────────────────────────────────────────────────────────
if (usage) app.use('/api/usage', usage.router());
else app.use('/api/usage', (req, res) => res.status(503).json({ error: 'Estadísticas de uso no disponibles' }));

// 404 JSON
app.use((req, res) => res.status(404).json({ error: 'Ruta no encontrada' }));

// Error handler: never leaks stack traces
// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) return res.status(400).json({ error: 'Archivo inválido: ' + err.message });
  if (err.type === 'entity.too.large') return res.status(413).json({ error: 'Payload demasiado grande' });
  if (err.type === 'entity.parse.failed') return res.status(400).json({ error: 'JSON inválido' });
  if (err.expose && err.status) return res.status(err.status).json({ error: err.message, ...(err.extra || {}) });
  console.error(err);
  res.status(500).json({ error: 'Error interno del servidor' });
});

// ── Start / shutdown ──────────────────────────────────────────────────────────
const server = app.listen(PORT, HOST, () => {
  const base = `http://localhost:${PORT}`;
  console.log(`\n🟢  Excel Template Builder v${pkg.version}  →  ${base}`);
  console.log(`📁  Datos:          ${DATA_ROOT}`);
  console.log(`🩺  Salud:          ${base}/api/health`);
  console.log(`📘  OpenAPI:        ${base}/api/openapi.json`);
  console.log(`📊  Uso:            ${base}/api/usage${usage ? '' : '  (no disponible)'}`);
  console.log(`📡  Tiempo real:    ${base}/api/live/presence${live ? '' : '  (no disponible)'}\n`);
});
server.on('error', err => { console.error(err); process.exit(1); });
server.requestTimeout = 120000;
server.headersTimeout = 65000;

let shuttingDown = false;
async function shutdown() {
  if (shuttingDown) return;
  shuttingDown = true;
  console.log('\n⏹️  Cerrando servidor…');
  try { live?.close(); } catch (e) { console.warn('live.close:', e.message); }
  try { await usage?.flush(); } catch (e) { console.warn('usage.flush:', e.message); }
  const force = setTimeout(() => process.exit(0), 5000);
  force.unref();
  server.close(() => process.exit(0));
}
process.on('SIGTERM', shutdown);
process.on('SIGINT', shutdown);

module.exports = { app, server };
