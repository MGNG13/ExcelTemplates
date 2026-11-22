'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { createStore, SAFE_ID } = require('../Library/store');

// Temp dirs go under ETB_TEST_DIR when given (never the real data/ folder)
const base = process.env.ETB_TEST_DIR || os.tmpdir();
fs.mkdirSync(base, { recursive: true });
const freshDir = () => fs.mkdtempSync(path.join(base, 'etb-store-'));
const tpl = (extra = {}) => ({ name: 'Prueba', sheets: [{ name: 'Hoja1', cells: { A1: { value: '{{x}}', variable: { name: 'x', type: 'text' } } } }], ...extra });
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

test('create assigns id, revision 1 and timestamps', async () => {
  const store = createStore({ dataDir: freshDir() });
  const by = { clientId: 'c1', userId: 'u1', name: 'Ana' };
  const r = await store.save(tpl(), { by });
  assert.equal(r.ok, true); assert.equal(r.created, true);
  assert.ok(SAFE_ID.test(r.template.id));
  assert.equal(r.template.revision, 1);
  assert.deepEqual(r.template.updatedBy, by);
  assert.ok(r.template.createdAt && r.template.updatedAt);
  const list = await store.list();
  assert.equal(list.length, 1);
  assert.equal(list[0].variables, 1);
  assert.deepEqual(list[0].sheetNames, ['Hoja1']);
  assert.equal(list[0].revision, 1);
});

test('save increments revision and preserves createdAt; conflict and force', async () => {
  const store = createStore({ dataDir: freshDir() });
  const { template: t1 } = await store.save(tpl({ id: 'abc' }));
  const r2 = await store.save({ ...t1, name: 'v2' }, { baseRevision: 1 });
  assert.equal(r2.ok, true); assert.equal(r2.created, false); assert.equal(r2.template.revision, 2);
  assert.equal(r2.template.createdAt, t1.createdAt);
  const stale = await store.save({ ...t1, name: 'stale' }, { baseRevision: 1 });
  assert.equal(stale.ok, false); assert.equal(stale.conflict, true);
  assert.equal(stale.current.revision, 2);
  assert.equal((await store.get('abc')).name, 'v2');
  const forced = await store.save({ ...t1, name: 'forzado' }, { baseRevision: 1, force: true });
  assert.equal(forced.ok, true); assert.equal(forced.template.revision, 3);
  const noBase = await store.save({ ...t1, name: 'sin base' });
  assert.equal(noBase.template.revision, 4);
  assert.equal((await store.list())[0].name, 'sin base');
});

test('legacy template without revision is treated as revision 0', async () => {
  const dir = freshDir();
  fs.mkdirSync(path.join(dir, 'templates'));
  fs.writeFileSync(path.join(dir, 'templates', 'legacy.json'),
    JSON.stringify({ id: 'legacy', name: 'Viejo', updatedAt: '2024-01-01T00:00:00.000Z', sheets: [] }));
  const store = createStore({ dataDir: dir });
  const [s] = await store.list();
  assert.equal(s.revision, 0);
  assert.equal(s.createdAt, '2024-01-01T00:00:00.000Z');
  assert.equal((await store.save({ id: 'legacy', name: 'x', sheets: [] }, { baseRevision: 5 })).conflict, true);
  const r = await store.save({ id: 'legacy', name: 'Nuevo', sheets: [] }, { baseRevision: 0 });
  assert.equal(r.ok, true); assert.equal(r.template.revision, 1);
  assert.equal(r.template.createdAt, '2024-01-01T00:00:00.000Z');
});

test('thumbnail validation; thumbnails do not bump revision', async () => {
  const store = createStore({ dataDir: freshDir() });
  const { template } = await store.save(tpl({ id: 'th' }));
  assert.equal(await store.setThumbnail('th', PNG), true);
  const after = await store.get('th');
  assert.equal(after.thumbnail, PNG); assert.equal(after.revision, 1); assert.equal(after.updatedAt, template.updatedAt);
  assert.equal(await store.setThumbnail('th', 'data:image/svg+xml;base64,PHN2Zz4='), null);
  assert.equal(await store.setThumbnail('th', 'javascript:alert(1)'), null);
  assert.equal(await store.setThumbnail('th', 'data:image/png;base64,' + 'A'.repeat(2 * 1024 * 1024)), null);
  assert.equal(await store.setThumbnail('nope', PNG), false);
  assert.equal((await store.list())[0].thumbnail, PNG);
  await store.save({ id: 'th', name: 'y', sheets: [] });   // save without thumbnail keeps it
  assert.equal((await store.get('th')).thumbnail, PNG);
});

test('duplicate and remove', async () => {
  const store = createStore({ dataDir: freshDir() });
  await store.save(tpl({ id: 'orig', name: 'Factura' }));
  await store.save(tpl({ id: 'orig', name: 'Factura' }));
  const copy = await store.duplicate('orig', { by: { name: 'Luis' } });
  assert.equal(copy.name, 'Factura (copia)'); assert.equal(copy.revision, 1); assert.notEqual(copy.id, 'orig');
  assert.equal(copy.updatedBy.name, 'Luis');
  assert.equal(await store.duplicate('missing'), null);
  assert.equal((await store.list()).length, 2);
  assert.equal(await store.remove('orig'), true);
  assert.equal(await store.remove('orig'), false);
  assert.equal(await store.get('orig'), null);
  assert.deepEqual((await store.list()).map(s => s.id), [copy.id]);
});

test('invalid ids are rejected', async () => {
  const store = createStore({ dataDir: freshDir() });
  for (const id of ['../x', 'a/b', '', 'x'.repeat(65), 'a.b', null, 5]) assert.equal(store.isValidId(id), false);
  assert.equal(await store.get('../etc'), null);
  assert.equal(await store.remove('..\\x'), false);
  assert.equal(await store.setThumbnail('../x', PNG), null);
  await assert.rejects(store.save({ id: '../evil', sheets: [] }), /ID inválido/);
  await assert.rejects(store.save({ id: 'ok' }), /inválidos/);
});

test('concurrent saves serialize; revisions increase monotonically', async () => {
  const store = createStore({ dataDir: freshDir() });
  await store.save(tpl({ id: 'conc' }));
  const results = await Promise.all(Array.from({ length: 25 }, (_, i) => store.save(tpl({ id: 'conc', name: 'n' + i }))));
  const revs = results.map(r => r.template.revision).sort((a, b) => a - b);
  assert.deepEqual(revs, Array.from({ length: 25 }, (_, i) => i + 2));
  assert.equal((await store.get('conc')).revision, 26);
  // Many writers with the same base revision → exactly one wins
  const race = await Promise.all(Array.from({ length: 10 }, () => store.save(tpl({ id: 'conc' }), { baseRevision: 26 })));
  assert.equal(race.filter(r => r.ok).length, 1);
  assert.equal(race.filter(r => r.conflict).length, 9);
  assert.deepEqual(fs.readdirSync(store.dir), ['conc.json']);   // no tmp leftovers
});
