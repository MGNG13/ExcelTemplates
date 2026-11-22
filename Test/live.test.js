'use strict';
// Pruebas del hub de tiempo real (Library/live.js). Ejecutar: node --test test/live.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('http');
const express = require('express');
const { createLiveHub } = require('../Library/live');

const PORT = Number(process.env.LIVE_TEST_PORT) || 4102;
const BASE = `http://127.0.0.1:${PORT}/api/live`;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let live;
let server;

/** Abre un stream SSE y acumula eventos parseados. */
function openStream(params) {
  return new Promise((resolve, reject) => {
    const qs = new URLSearchParams(params).toString();
    const req = http.get(`${BASE}/stream?${qs}`, (res) => {
      const s = { res, req, status: res.statusCode, headers: res.headers, events: [], ended: false, body: '' };
      s.close = () => { req.destroy(); };
      s.waitFor = async (pred, ms = 2000) => {
        const t0 = Date.now();
        while (Date.now() - t0 < ms) {
          const hit = s.events.find(pred);
          if (hit) return hit;
          await sleep(15);
        }
        throw new Error('Evento no recibido');
      };
      let buf = '';
      res.setEncoding('utf8');
      res.on('data', (chunk) => {
        s.body += chunk;
        buf += chunk;
        let i;
        while ((i = buf.indexOf('\n\n')) >= 0) {
          const block = buf.slice(0, i);
          buf = buf.slice(i + 2);
          let event = 'message'; const data = [];
          for (const line of block.split('\n')) {
            if (line.startsWith('event: ')) event = line.slice(7);
            else if (line.startsWith('data: ')) data.push(line.slice(6));
          }
          if (data.length) s.events.push({ event, data: JSON.parse(data.join('\n')) });
        }
      });
      res.on('end', () => { s.ended = true; });
      res.on('error', () => { s.ended = true; });
      resolve(s);
    });
    req.on('error', reject);
  });
}

function request(method, path, { body, headers = {} } = {}) {
  return new Promise((resolve, reject) => {
    const req = http.request(`${BASE}${path}`, { method, headers }, (res) => {
      let data = '';
      res.on('data', (c) => { data += c; });
      res.on('end', () => {
        let json = null; try { json = data ? JSON.parse(data) : null; } catch (_) { /* no JSON */ }
        resolve({ status: res.statusCode, json, text: data });
      });
    });
    req.on('error', reject);
    if (body !== undefined) req.end(typeof body === 'string' ? body : JSON.stringify(body)); else req.end();
  });
}
const postJson = (path, body, type = 'application/json') => request('POST', path, { body, headers: { 'Content-Type': type } });

test.before(async () => {
  live = createLiveHub({ heartbeatMs: 200, idleMs: 60000, maxClients: 5, maxPerIp: 50 });
  const app = express();
  app.use('/api/live', live.router());
  await new Promise((r) => { server = app.listen(PORT, '127.0.0.1', r); });
});

test.after(async () => {
  live.close();
  server.closeAllConnections?.();
  await new Promise((r) => server.close(r));
});

test('hello, cabeceras SSE, retry y ping', async () => {
  const s = await openStream({ clientId: 'clientAAAA1', userId: 'userA', name: 'Ana', color: '#112233' });
  assert.equal(s.status, 200);
  assert.match(s.headers['content-type'], /^text\/event-stream/);
  assert.equal(s.headers['cache-control'], 'no-cache, no-transform');
  assert.equal(s.headers['x-accel-buffering'], 'no');
  const hello = await s.waitFor((e) => e.event === 'hello');
  assert.ok(s.body.startsWith('retry: 3000'));
  assert.equal(hello.data.clientId, 'clientAAAA1');
  const me = hello.data.clients.find((c) => c.clientId === 'clientAAAA1');
  assert.equal(me.name, 'Ana');
  assert.equal(me.color, '#112233');
  assert.equal(me.mode, 'home');
  assert.equal(me.idle, false);
  await s.waitFor((e) => e.event === 'ping', 1000);
  s.close();
  await sleep(50);
  assert.equal(live.getPresence().length, 0);
});

test('clientId inválido → 400', async () => {
  const r = await request('GET', '/stream?clientId=x');
  assert.equal(r.status, 400);
});

test('presence: merge, broadcast coalescido, enter/leave y countByTemplate', async () => {
  const a = await openStream({ clientId: 'clientAAAA1', userId: 'userA', name: 'Ana' });
  const b = await openStream({ clientId: 'clientBBBB1', userId: 'userB', name: 'Beto' });
  await a.waitFor((e) => e.event === 'hello');
  const enters = []; const leaves = [];
  const onEnter = (p) => enters.push(p); const onLeave = (p) => leaves.push(p);
  live.on('enter', onEnter); live.on('leave', onLeave);

  let r = await postJson('/presence', { clientId: 'clientAAAA1', mode: 'view', templateId: 'tpl-1', templateName: 'Factura', active: true });
  assert.equal(r.status, 204);
  r = await postJson('/presence', { clientId: 'clientAAAA1', mode: 'edit', sheet: 'Hoja1', cell: 'B2' });
  assert.equal(r.status, 204);
  assert.equal(enters.length, 1, 'solo una transición');
  assert.deepEqual(enters[0], { templateId: 'tpl-1', templateName: 'Factura', client: { clientId: 'clientAAAA1', userId: 'userA', name: 'Ana' } });

  const pres = await b.waitFor((e) => e.event === 'presence' && e.data.clients.some((c) => c.cell === 'B2'));
  const ana = pres.data.clients.find((c) => c.clientId === 'clientAAAA1');
  assert.equal(ana.templateName, 'Factura'); // merge conservó el nombre
  assert.equal(ana.mode, 'edit');
  assert.equal(ana.sheet, 'Hoja1');

  await postJson('/presence', { clientId: 'clientBBBB1', mode: 'view', templateId: 'tpl-1' });
  assert.deepEqual(live.countByTemplate()['tpl-1'].editors, 1);
  assert.deepEqual(live.countByTemplate()['tpl-1'].viewers, 1);
  assert.equal(live.countByTemplate()['tpl-1'].users.length, 2);

  // Coalescencia: ráfaga de 10 updates → pocas difusiones
  const before = a.events.filter((e) => e.event === 'presence').length;
  for (let i = 1; i <= 10; i++) await postJson('/presence', { clientId: 'clientAAAA1', cell: `C${i}` });
  await sleep(400);
  const after = a.events.filter((e) => e.event === 'presence').length;
  assert.ok(after - before <= 4, `difusiones: ${after - before}`);

  // Cambio de plantilla → leave + enter
  await postJson('/presence', { clientId: 'clientAAAA1', templateId: 'tpl-2' });
  assert.equal(leaves.at(-1).templateId, 'tpl-1');
  assert.equal(enters.at(-1).templateId, 'tpl-2');
  assert.equal(enters.at(-1).templateName, null);

  // Desconexión → leave
  b.close();
  await sleep(50);
  assert.equal(leaves.at(-1).templateId, 'tpl-1');
  assert.equal(leaves.at(-1).client.clientId, 'clientBBBB1');
  live.off('enter', onEnter); live.off('leave', onLeave);
  a.close();
  await sleep(50);
});

test('sanitización de nombres, colores, celdas y templateId', async () => {
  const long = '<b>' + 'x'.repeat(100) + '</b>\u0007';
  const s = await openStream({ clientId: 'clientSAN01', name: long, color: 'red' });
  const hello = await s.waitFor((e) => e.event === 'hello');
  const me = hello.data.clients[0];
  assert.ok(me.name.length <= 40);
  assert.ok(!/[<>\u0007]/.test(me.name));
  assert.match(me.color, /^#[0-9a-fA-F]{6}$/);
  assert.equal(me.userId, 'clientSAN01');

  await postJson('/presence', {
    clientId: 'clientSAN01', name: '   ', color: 'javascript:1', cell: 'a1<script>', range: 'A1:B', templateId: '../etc/passwd',
    templateName: 'T'.repeat(500), sheet: 'S'.repeat(100), mode: 'admin',
  });
  const [e] = live.getPresence();
  assert.equal(e.name, 'Invitado');
  assert.match(e.color, /^#[0-9a-fA-F]{6}$/);
  assert.equal(e.cell, null);
  assert.equal(e.range, null);
  assert.equal(e.templateId, null);
  assert.equal(e.templateName.length, 120);
  assert.equal(e.sheet.length, 64);
  assert.equal(e.mode, 'home');
  await postJson('/presence', { clientId: 'clientSAN01', range: 'A1:C10', cell: 'AB12' });
  assert.equal(live.getPresence()[0].range, 'A1:C10');
  assert.equal(live.getPresence()[0].cell, 'AB12');
  s.close();
  await sleep(50);
});

test('NO_STREAM 404, body inválido 400, GET /presence', async () => {
  let r = await postJson('/presence', { clientId: 'desconocido1' });
  assert.equal(r.status, 404);
  assert.equal(r.json.code, 'NO_STREAM');
  r = await postJson('/presence', '{malo');
  assert.equal(r.status, 400);
  assert.ok(!r.text.includes('malo'));
  r = await request('GET', '/presence');
  assert.equal(r.status, 200);
  assert.ok(Array.isArray(r.json.clients));
  assert.equal(typeof r.json.serverTime, 'number');
});

test('leave por JSON y text/plain (sendBeacon)', async () => {
  const s = await openStream({ clientId: 'clientLEAVE1' });
  await s.waitFor((e) => e.event === 'hello');
  const leaves = []; const on = (p) => leaves.push(p); live.on('leave', on);
  await postJson('/presence', { clientId: 'clientLEAVE1', templateId: 'tplL', mode: 'view' });
  let r = await postJson('/leave', { clientId: 'clientLEAVE1' });
  assert.equal(r.status, 204);
  assert.equal(leaves.length, 1);
  assert.equal(live.getPresence()[0].templateId, null);
  await postJson('/presence', { clientId: 'clientLEAVE1', templateId: 'tplL' });
  r = await postJson('/leave', JSON.stringify({ clientId: 'clientLEAVE1' }), 'text/plain;charset=UTF-8');
  assert.equal(r.status, 204);
  assert.equal(leaves.length, 2);
  r = await postJson('/leave', { clientId: 'nadie' });
  assert.equal(r.status, 204);
  live.off('leave', on);
  s.close();
  await sleep(50);
});

test('reemplazo del mismo clientId sin leave espurio', async () => {
  const leaves = []; const on = (p) => leaves.push(p); live.on('leave', on);
  const s1 = await openStream({ clientId: 'clientREP01', name: 'Uno' });
  await s1.waitFor((e) => e.event === 'hello');
  await postJson('/presence', { clientId: 'clientREP01', templateId: 'tplR', mode: 'view' });
  const s2 = await openStream({ clientId: 'clientREP01', name: 'Dos' });
  await s2.waitFor((e) => e.event === 'hello');
  await sleep(50);
  assert.equal(s1.ended, true, 'el stream viejo se cerró');
  const p = live.getPresence();
  assert.equal(p.length, 1);
  assert.equal(p[0].name, 'Dos');
  assert.equal(p[0].templateId, 'tplR');
  assert.equal(leaves.length, 0);
  s2.close();
  await sleep(50);
  assert.equal(leaves.length, 1);
  live.off('leave', on);
});

test('clientFromRequest', async () => {
  assert.equal(live.clientFromRequest({ headers: {} }), null);
  assert.deepEqual(live.clientFromRequest({ headers: { 'x-user-id': 'u1', 'x-client-name': encodeURIComponent('José <i>') } }),
    { clientId: null, userId: 'u1', name: 'José i' });
  assert.equal(live.clientFromRequest({ headers: { 'x-client-name': '%E0%A4%A' } }).name, '%E0%A4%A');
  const s = await openStream({ clientId: 'clientCFR01', userId: 'uReg', name: 'Registrado' });
  await s.waitFor((e) => e.event === 'hello');
  assert.deepEqual(live.clientFromRequest({ headers: { 'x-client-id': 'clientCFR01', 'x-client-name': 'Otro' } }),
    { clientId: 'clientCFR01', userId: 'uReg', name: 'Registrado' });
  s.close();
  await sleep(50);
});

test('maxClients → 503', async () => {
  const streams = [];
  for (let i = 0; i < 5; i++) streams.push(await openStream({ clientId: `clientMAX0${i}` }));
  const r = await request('GET', '/stream?clientId=clientMAX99');
  assert.equal(r.status, 503);
  for (const s of streams) s.close();
  await sleep(50);
});

test('publish difunde y valida el nombre', async () => {
  const s = await openStream({ clientId: 'clientPUB01' });
  await s.waitFor((e) => e.event === 'hello');
  const n = live.publish('template:saved', { templateId: 't1', revision: 3, name: 'A\nB' });
  assert.equal(n, 1);
  const ev = await s.waitFor((e) => e.event === 'template:saved');
  assert.equal(ev.data.name, 'A\nB');
  assert.throws(() => live.publish('bad event\n', {}), TypeError);
  assert.throws(() => live.publish('Upper', {}), TypeError);
  s.close();
  await sleep(50);
});
