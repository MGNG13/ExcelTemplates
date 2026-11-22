// Template persistence: one JSON file per template in <dataDir>/templates.
// - atomic writes (tmp file + rename, retried on transient Windows locks)
// - per-id promise-chain mutex so writes to the same template serialize
// - optimistic concurrency via a monotonically increasing `revision`
// - in-memory summary cache for list(), updated on every write
'use strict';
const fsp = require('fs').promises;
const path = require('path');
const crypto = require('crypto');

const SAFE_ID = /^[A-Za-z0-9_-]{1,64}$/;
const THUMB_RE = /^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+/=\s]+$/;
const THUMB_MAX = 2 * 1024 * 1024;
const RETRY_CODES = new Set(['EPERM', 'EBUSY', 'EACCES']);

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function atomicWrite(file, text) {
  const tmp = `${file}.${process.pid}.${crypto.randomBytes(4).toString('hex')}.tmp`;
  await fsp.writeFile(tmp, text, 'utf8');
  for (let attempt = 0; ; attempt++) {
    try { await fsp.rename(tmp, file); return; }
    catch (e) {
      if (!RETRY_CODES.has(e.code) || attempt >= 6) {
        await fsp.unlink(tmp).catch(() => {});
        throw e;
      }
      await sleep(15 * (attempt + 1));   // antivirus / indexer briefly holding the file
    }
  }
}

function summarize(t) {
  const sheets = Array.isArray(t.sheets) ? t.sheets : [];
  return {
    id: t.id,
    name: t.name,
    sheetCount: sheets.length,
    updatedAt: t.updatedAt || null,
    createdAt: t.createdAt || t.updatedAt || null,
    revision: typeof t.revision === 'number' ? t.revision : 0,
    updatedBy: t.updatedBy || null,
    thumbnail: t.thumbnail || null,
    sheetNames: sheets.filter(s => s.state === undefined || s.state === 'visible').map(s => s.name),
    variables: sheets.reduce((n, s) => n + Object.values(s.cells || {}).filter(c => c && c.variable).length, 0),
  };
}

function sanitizeBy(by) {
  if (!by || typeof by !== 'object') return null;
  const s = v => (v == null ? null : String(v).slice(0, 120));
  return { clientId: s(by.clientId), userId: s(by.userId), name: s(by.name) };
}

function createStore({ dataDir } = {}) {
  if (!dataDir) throw new TypeError('createStore: dataDir es obligatorio');
  const dir = path.join(dataDir, 'templates');
  const locks = new Map();      // id → tail promise
  let cache = null;             // Map id → summary (null = not loaded)
  let ready = null;

  const ensureDir = () => (ready ||= fsp.mkdir(dir, { recursive: true }));
  const isValidId = id => typeof id === 'string' && SAFE_ID.test(id);
  const fileOf = id => path.join(dir, `${id}.json`);

  // Serialize async work per id; errors don't break the chain
  function withLock(id, fn) {
    const prev = locks.get(id) || Promise.resolve();
    const run = prev.then(fn, fn);
    const tail = run.catch(() => {});
    locks.set(id, tail);
    tail.then(() => { if (locks.get(id) === tail) locks.delete(id); });
    return run;
  }

  async function readRaw(id) {
    try { return JSON.parse(await fsp.readFile(fileOf(id), 'utf8')); }
    catch (e) {
      if (e.code === 'ENOENT') return null;
      if (e instanceof SyntaxError) { console.warn(`plantilla corrupta ${id}:`, e.message); return null; }
      throw e;
    }
  }

  async function writeTpl(t) {
    await ensureDir();
    await atomicWrite(fileOf(t.id), JSON.stringify(t, null, 2));
    if (cache) cache.set(t.id, summarize(t));
  }

  async function loadCache() {
    await ensureDir();
    const files = (await fsp.readdir(dir)).filter(f => f.endsWith('.json'));
    const map = new Map();
    await Promise.all(files.map(async f => {
      const id = f.slice(0, -5);
      if (!isValidId(id)) return;
      const t = await readRaw(id).catch(() => null);
      if (t && typeof t === 'object') map.set(id, summarize({ ...t, id: t.id || id }));
    }));
    cache = map;
  }

  async function list() {
    if (!cache) await loadCache();
    const ts = s => Date.parse(s.updatedAt) || 0;   // legacy / missing dates sort last
    return [...cache.values()].sort((a, b) => ts(b) - ts(a));
  }

  async function get(id) {
    if (!isValidId(id)) return null;
    await ensureDir();
    return readRaw(id);
  }

  async function save(tpl, { baseRevision = null, force = false, by = null } = {}) {
    if (!tpl || typeof tpl !== 'object' || !Array.isArray(tpl.sheets)) throw new TypeError('Datos de plantilla inválidos');
    const t = { ...tpl };
    if (t.id == null || t.id === '') t.id = crypto.randomUUID();
    t.id = String(t.id);
    if (!isValidId(t.id)) { const e = new Error('ID inválido'); e.status = 400; throw e; }
    return withLock(t.id, async () => {
      await ensureDir();
      const disk = await readRaw(t.id);
      const now = new Date().toISOString();
      const cur = disk ? (typeof disk.revision === 'number' ? disk.revision : 0) : 0;
      if (disk && !force && baseRevision !== null && baseRevision !== cur) {
        return { ok: false, conflict: true, current: { revision: cur, updatedAt: disk.updatedAt || null, updatedBy: disk.updatedBy || null } };
      }
      t.revision = cur + 1;
      t.updatedAt = now;
      t.updatedBy = sanitizeBy(by);
      t.createdAt = (disk && (disk.createdAt || disk.updatedAt)) || t.createdAt || now;
      // Thumbnails are managed by setThumbnail; keep the stored one if the client omitted it
      if (disk && t.thumbnail === undefined && disk.thumbnail) t.thumbnail = disk.thumbnail;
      await writeTpl(t);
      return { ok: true, template: t, created: !disk };
    });
  }

  async function setThumbnail(id, dataUrl) {
    if (!isValidId(id)) return null;
    if (dataUrl !== null && (typeof dataUrl !== 'string' || dataUrl.length > THUMB_MAX || !THUMB_RE.test(dataUrl))) return null;
    return withLock(id, async () => {
      const t = await readRaw(id);
      if (!t) return false;
      t.thumbnail = dataUrl;     // revision / updatedAt intentionally untouched
      await writeTpl(t);
      return true;
    });
  }

  async function duplicate(id, { by = null } = {}) {
    const src = await get(id);
    if (!src) return null;
    const now = new Date().toISOString();
    const copy = { ...src, id: crypto.randomUUID(), name: (src.name || 'Plantilla') + ' (copia)',
      revision: 1, createdAt: now, updatedAt: now, updatedBy: sanitizeBy(by) };
    await withLock(copy.id, () => writeTpl(copy));
    return copy;
  }

  async function remove(id) {
    if (!isValidId(id)) return false;
    return withLock(id, async () => {
      try { await fsp.unlink(fileOf(id)); }
      catch (e) { if (e.code === 'ENOENT') { cache?.delete(id); return false; } throw e; }
      cache?.delete(id);
      return true;
    });
  }

  return { list, get, save, setThumbnail, duplicate, remove, isValidId, dir };
}

module.exports = { createStore, SAFE_ID };
