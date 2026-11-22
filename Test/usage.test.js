'use strict';
// Pruebas de Library/usage.js — ejecutar: node --test test/usage.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');
const { createUsageTracker } = require('../Library/usage');

const ROOT = process.env.ETB_DATA_DIR || fs.mkdtempSync(path.join(os.tmpdir(), 'etb-usage-'));
let n = 0;
const freshDir = () => { const d = path.join(ROOT, `case-${process.pid}-${n++}`); fs.mkdirSync(d, { recursive: true }); return d; };
const H = 3600e3;

test('agregados, errores, batch, bytes y avgMs', () => {
  const u = createUsageTracker({ dataDir: freshDir() });
  u.record({ type: 'render', templateId: 'a', templateName: 'Factura', ms: 100, bytes: 1000, source: 'ui', client: 'c1' });
  u.record({ type: 'render', templateId: 'a', ms: 300, bytes: 500, source: 'api' });
  u.record({ type: 'render', templateId: 'a', ok: false, error: 'boom' });
  u.record({ type: 'batch', templateId: 'a', count: 5, source: 'api', client: 'c1' });
  u.record({ type: 'view', templateId: 'a' });
  u.record({ type: 'save', templateId: 'a' });
  const d = u.detail('a');
  assert.equal(d.name, 'Factura');
  assert.equal(d.totals.renders, 7);
  assert.equal(d.totals.errors, 1);
  assert.equal(d.totals.views, 1);
  assert.equal(d.totals.saves, 1);
  assert.equal(d.totals.bytes, 1500);
  assert.equal(d.totals.avgMs, 200);
  assert.deepEqual(d.sources, { ui: 1, api: 6, link: 0 });
  assert.deepEqual(d.topClients[0], { client: 'c1', renders: 6 });
  assert.equal(d.today.renders, 7);
  assert.equal(d.today.errors, 1);
  const s = u.summary();
  assert.equal(s.totals.rendersLastHour, 7);
  assert.equal(s.totals.templatesUsed, 1);
  assert.equal(s.totals.today.views, 1);
  assert.equal(s.recent.length, 6);
});

test('validación: tipo y templateId inválidos se ignoran', () => {
  const u = createUsageTracker({ dataDir: freshDir() });
  assert.equal(u.record({ type: 'hack', templateId: 'a' }), null);
  assert.equal(u.record({ type: 'render', templateId: '../x' }), null);
  assert.equal(u.record(null), null);
  assert.equal(u.summary().templates.length, 0);
});

test('dedupe', () => {
  let t = Date.now();
  const u = createUsageTracker({ dataDir: freshDir(), now: () => t });
  assert.ok(u.record({ type: 'view', templateId: 'a' }, { dedupeKey: 'u1a', dedupeMs: 600000 }));
  assert.equal(u.record({ type: 'view', templateId: 'a' }, { dedupeKey: 'u1a', dedupeMs: 600000 }), null);
  t += 600001;
  assert.ok(u.record({ type: 'view', templateId: 'a' }, { dedupeKey: 'u1a', dedupeMs: 600000 }));
  assert.equal(u.detail('a').totals.views, 2);
});

test('spark24h / spark7d longitudes y posiciones', () => {
  const now = new Date(2026, 10, 19, 15, 30).getTime();
  const u = createUsageTracker({ dataDir: freshDir(), now: () => now });
  u.record({ type: 'render', templateId: 'a' });
  u.record({ type: 'render', templateId: 'a' }, { now: now - 2 * H });
  u.record({ type: 'render', templateId: 'a' }, { now: now - 23 * H });
  u.record({ type: 'render', templateId: 'a' }, { now: now - 24 * H }); // fuera de 24h
  u.record({ type: 'render', templateId: 'a' }, { now: now - 3 * 24 * H });
  const d = u.detail('a');
  assert.equal(d.spark24h.length, 24);
  assert.equal(d.spark24h[23], 1);
  assert.equal(d.spark24h[21], 1);
  assert.equal(d.spark24h[0], 1);
  assert.equal(d.spark24h.reduce((a, b) => a + b), 3);
  assert.equal(d.spark7d.length, 7);
  assert.equal(d.spark7d[6], 2); // hoy: 15:30 y 13:30
  assert.equal(d.spark7d[5], 2); // ayer 16:30 y 15:30
  assert.equal(d.spark7d[3], 1);
  assert.equal(d.hourly48.length, 48);
  assert.equal(d.daily30.length, 30);
  assert.equal(d.daily30[29].day, '2026-11-19');
  assert.equal(d.hourly48[47].hour, '2026-11-19T15');
  assert.equal(u.summary().totals.rendersLastHour, 1);
  assert.equal(u.summary().totals.today.renders, 2);
});

test('orden del summary: renders hoy desc, luego total', () => {
  const now = Date.now();
  const u = createUsageTracker({ dataDir: freshDir(), now: () => now });
  for (let i = 0; i < 5; i++) u.record({ type: 'render', templateId: 'old' }, { now: now - 10 * 24 * H });
  u.record({ type: 'render', templateId: 'hoy' });
  u.record({ type: 'view', templateId: 'nada' });
  assert.deepEqual(u.summary().templates.map((t) => t.id), ['hoy', 'old', 'nada']);
  assert.equal(u.summary().totals.templatesUsed, 2);
});

test('persistencia y recarga desde disco', async () => {
  const dir = freshDir();
  const u = createUsageTracker({ dataDir: dir });
  u.record({ type: 'render', templateId: 'p', templateName: 'Persistida', bytes: 10 });
  await u.flush();
  assert.ok(fs.existsSync(path.join(dir, 'usage', 'usage.json')));
  const u2 = createUsageTracker({ dataDir: dir });
  const d = u2.detail('p');
  assert.equal(d.totals.renders, 1);
  assert.equal(d.name, 'Persistida');
  assert.equal(d.recent.length, 1);
});

test('archivo corrupto: se renombra y se inicia vacío', () => {
  const dir = freshDir();
  fs.mkdirSync(path.join(dir, 'usage'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'usage', 'usage.json'), '{not json');
  const u = createUsageTracker({ dataDir: dir });
  assert.equal(u.summary().templates.length, 0);
  const files = fs.readdirSync(path.join(dir, 'usage'));
  assert.ok(files.some((f) => f.startsWith('usage.json.corrupt-')));
  assert.ok(!files.includes('usage.json'));
});

test('retención: horas > 8 días y días > 120 días se podan', () => {
  const now = Date.now();
  const u = createUsageTracker({ dataDir: freshDir(), now: () => now });
  u.record({ type: 'render', templateId: 'r' }, { now: now - 200 * 24 * H });
  u.record({ type: 'render', templateId: 'r' }, { now: now - 30 * 24 * H });
  u.record({ type: 'render', templateId: 'r' }); // dispara poda con now actual
  const d = u.detail('r');
  assert.equal(d.totals.renders, 3); // totales históricos se conservan
  assert.equal(d.daily30.reduce((a, x) => a + x.renders, 0), 1); // solo hoy (hace 30 días cae fuera de daily30)
});

test('retención inspeccionando el archivo', async () => {
  const dir = freshDir();
  const now = Date.now();
  const u = createUsageTracker({ dataDir: dir, now: () => now });
  u.record({ type: 'render', templateId: 'r' }, { now: now - 200 * 24 * H });
  u.record({ type: 'render', templateId: 'r' }, { now: now - 9 * 24 * H });
  u.record({ type: 'render', templateId: 'r' });
  await u.flush();
  const t = JSON.parse(fs.readFileSync(path.join(dir, 'usage', 'usage.json'), 'utf8')).templates.r;
  assert.equal(Object.keys(t.hourly).length, 1);
  assert.equal(Object.keys(t.daily).length, 2);
});

test('markDeleted y rename', () => {
  const u = createUsageTracker({ dataDir: freshDir() });
  u.record({ type: 'render', templateId: 'm', templateName: 'Uno' });
  u.rename('m', 'Dos');
  assert.equal(u.detail('m').name, 'Dos');
  u.markDeleted('m');
  const s = u.summary().templates.find((t) => t.id === 'm');
  assert.equal(s.deleted, true);
  u.record({ type: 'delete', templateId: 'z' });
  assert.equal(u.detail('z').deleted, true);
  const unknown = u.detail('nope');
  assert.equal(unknown.totals.renders, 0);
  assert.equal(unknown.spark24h.length, 24);
  assert.deepEqual(unknown.recent, []);
});

test('live.publish recibe evento usage', () => {
  const calls = [];
  const u = createUsageTracker({ dataDir: freshDir(), live: { publish: (name, data) => calls.push([name, data]) } });
  u.record({ type: 'render', templateId: 'l' });
  assert.equal(calls.length, 1);
  assert.equal(calls[0][0], 'usage');
  assert.equal(calls[0][1].event.templateId, 'l');
  assert.equal(calls[0][1].template.id, 'l');
  assert.equal(calls[0][1].totals.renders, 1);
});

test('router HTTP en puerto 4103', async () => {
  const express = require('express');
  const u = createUsageTracker({ dataDir: freshDir() });
  for (let i = 0; i < 3; i++) u.record({ type: 'render', templateId: 'h' });
  u.record({ type: 'view', templateId: 'k' });
  const app = express();
  app.use('/api/usage', u.router());
  const server = await new Promise((r) => { const s = app.listen(4103, () => r(s)); });
  const get = (p) => new Promise((resolve, reject) => {
    http.get(`http://127.0.0.1:4103${p}`, (res) => {
      let body = ''; res.on('data', (c) => (body += c));
      res.on('end', () => resolve({ status: res.statusCode, cc: res.headers['cache-control'], json: JSON.parse(body) }));
    }).on('error', reject);
  });
  try {
    const s = await get('/api/usage');
    assert.equal(s.status, 200);
    assert.equal(s.cc, 'no-store');
    assert.equal(s.json.totals.renders, 3);
    const ev = await get('/api/usage/events?limit=2');
    assert.equal(ev.json.events.length, 2);
    const evk = await get('/api/usage/events?templateId=k');
    assert.equal(evk.json.events.length, 1);
    const d = await get('/api/usage/h');
    assert.equal(d.json.totals.renders, 3);
    assert.equal(d.json.hourly48.length, 48);
    const bad = await get('/api/usage/bad%20id!');
    assert.equal(bad.status, 400);
  } finally {
    await new Promise((r) => server.close(r));
    await u.flush();
  }
});
