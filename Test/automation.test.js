'use strict';
// Pruebas de Library/automation.js — ejecutar con: node --test test/automation.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const A = require('../Library/automation');

// Fixture compartido (scratchpad) con respaldo embebido si no existe
const FIXTURE = process.env.ETB_FIXTURE || 'C:/Users/MAGNUS~1/AppData/Local/Temp/claude/97a6d3a3288351ce/f7d38315-7959-4674-a505-060f7df6a440/scratchpad/fixtures/template-basic.json';
function loadFixture() {
  try { return JSON.parse(fs.readFileSync(FIXTURE, 'utf8')); } catch (_) { /* respaldo */ }
  const v = (name, type, required) => ({ value: `{{${name}}}`, variable: { name, type, hint: '', required, defaultValue: '' } });
  return {
    id: 'fixture-factura', name: 'Factura de prueba',
    sheets: [
      { name: 'Factura', isRepeatable: false, rowTemplate: { startRow: 5, endRow: 5, arrayVariable: 'items' },
        cells: { A1: { value: 'Cliente:' }, B1: v('cliente', 'text', true), B2: v('fecha', 'date', false),
          A5: v('descripcion', 'text', false), B5: v('cantidad', 'number', false), C5: v('precio', 'number', false), C7: v('total', 'number', true) } },
      { name: 'Pedido', isRepeatable: true, repeatVariable: 'pedidos', rowTemplate: null,
        cells: { B1: v('folio', 'text', true), B2: v('pagado', 'boolean', false) } },
    ],
  };
}
const clone = o => JSON.parse(JSON.stringify(o));
const base = loadFixture();

// Plantilla con variables implícitas (texto y fórmulas) y una variable obligatoria en la zona de filas
function withImplicit() {
  const t = clone(base);
  const f = t.sheets[0].cells;
  f.A9 = { value: 'Atendió: {{vendedor}} el {{fecha}}' };            // fecha ya explícita → no se duplica
  f.A10 = { value: 'Otra vez {{vendedor}}' };                         // dedupe
  f.D5 = { value: '', formula: '=B5*C5*{{tasa}}' };                    // implícita dentro de la zona
  f.B5.variable.required = true;                                       // cantidad obligatoria por item
  return t;
}
const findVar = (sheet, name, inRow) => sheet.variables.find(v => v.name === name && (inRow === undefined || v.inRowTemplate === inRow));

test('buildSchema: variables explícitas, tipos y zona de filas', () => {
  const s = A.buildSchema(base);
  assert.equal(s.id, 'fixture-factura');
  assert.equal(s.revision, 0, 'plantillas sin revision → 0');
  assert.equal(s.sheets.length, 2);
  const [fac, ped] = s.sheets;
  assert.deepEqual(fac.rowTemplate, { startRow: 5, endRow: 5, arrayVariable: 'items' });
  assert.equal(findVar(fac, 'cliente').required, true);
  assert.equal(findVar(fac, 'cliente').inRowTemplate, false);
  assert.equal(findVar(fac, 'cantidad').type, 'number');
  assert.equal(findVar(fac, 'cantidad').inRowTemplate, true);
  assert.equal(findVar(fac, 'total').inRowTemplate, false);
  assert.equal(findVar(fac, 'total').addr, 'C7');
  assert.equal(ped.isRepeatable, true);
  assert.equal(ped.rowTemplate, null);
  assert.equal(findVar(ped, 'pagado').type, 'boolean');
  for (const v of fac.variables) for (const k of ['name', 'type', 'required', 'hint', 'defaultValue', 'addr', 'inRowTemplate']) assert.ok(k in v, k);
});

test('buildSchema: variables implícitas en texto y fórmulas, deduplicadas', () => {
  const fac = A.buildSchema(withImplicit()).sheets[0];
  const vend = fac.variables.filter(v => v.name === 'vendedor');
  assert.equal(vend.length, 1);
  assert.equal(vend[0].implicit, true);
  assert.equal(vend[0].type, 'text');
  assert.equal(vend[0].required, false);
  assert.equal(fac.variables.filter(v => v.name === 'fecha').length, 1, 'la explícita gana');
  assert.equal(findVar(fac, 'fecha').implicit, false);
  assert.equal(findVar(fac, 'tasa').inRowTemplate, true);
  assert.equal(findVar(fac, 'tasa').implicit, true);
});

test('buildSchema: entradas raras no lanzan', () => {
  for (const x of [null, undefined, 5, 'x', [], { sheets: 'no' }, { sheets: [null, 3, { cells: null }] }]) {
    const s = A.buildSchema(x);
    assert.ok(Array.isArray(s.sheets));
  }
  const s = A.buildSchema({ sheets: [{ name: 'H', rowTemplate: { startRow: 'a' }, cells: { ZZ: {}, A1: { variable: { name: '__proto__' } } } }] });
  assert.equal(s.sheets[0].rowTemplate, null);
  assert.equal(s.sheets[0].variables.length, 0);
});

test('buildExample: forma del cuerpo', () => {
  const ex = A.buildExample(withImplicit());
  assert.ok(ex.Factura && !Array.isArray(ex.Factura));
  assert.equal(typeof ex.Factura.total, 'number');
  assert.ok('vendedor' in ex.Factura);
  assert.ok(!('cantidad' in ex.Factura), 'variables de la zona van en items');
  assert.equal(ex.Factura.items.length, 2);
  assert.deepEqual(Object.keys(ex.Factura.items[0]).sort(), ['cantidad', 'descripcion', 'precio', 'tasa']);
  assert.ok(Array.isArray(ex.Pedido) && ex.Pedido.length === 2);
  assert.equal(typeof ex.Pedido[0].pagado, 'boolean');
  // El ejemplo siempre es válido
  assert.equal(A.validateBody(withImplicit(), ex).valid, true);
  assert.equal(A.validateBody(base, A.buildExample(base)).valid, true);
});

test('validateBody: cuerpo válido', () => {
  const r = A.validateBody(base, {
    Factura: { cliente: 'Ana', fecha: '2026-01-15', total: '1160.5', items: [{ descripcion: 'X', cantidad: 2, precio: '10' }] },
    Pedido: [{ folio: 'P-1', pagado: 'true' }, { folio: 'P-2', pagado: 0 }],
  });
  assert.deepEqual(r, { valid: true, errors: [], warnings: [] });
});

test('validateBody: obligatorias faltantes, incluidos items de la zona', () => {
  const t = withImplicit();
  const r = A.validateBody(t, { Factura: { cliente: '  ', items: [{ cantidad: 1 }, { descripcion: 'sin cantidad' }] }, Pedido: [{ pagado: true }] });
  assert.equal(r.valid, false);
  const paths = r.errors.filter(e => e.code === 'missing').map(e => e.path).sort();
  assert.deepEqual(paths, ['Factura.cliente', 'Factura.items[1].cantidad', 'Factura.total', 'Pedido[0].folio'].sort());
  assert.ok(r.errors.every(e => /obligatoria/.test(e.message)));
  // Sin arreglo items pero con variables obligatorias en la zona
  const r2 = A.validateBody(t, { Factura: { cliente: 'A', total: 1 }, Pedido: [{ folio: 'x' }] });
  assert.equal(r2.errors.length, 1);
  assert.equal(r2.errors[0].path, 'Factura.items');
});

test('validateBody: errores de tipo', () => {
  const t = clone(base);
  t.sheets[0].cells.D1 = { value: '', variable: { name: 'logo', type: 'image', required: false } };
  const r = A.validateBody(t, {
    Factura: { cliente: { a: 1 }, fecha: 'no es fecha', total: 'mil', logo: 'http://x/y.png', items: [{ cantidad: 'dos' }] },
    Pedido: [{ folio: 'a', pagado: 'quizá' }],
  });
  const types = r.errors.filter(e => e.code === 'type').map(e => e.path).sort();
  assert.deepEqual(types, ['Factura.cliente', 'Factura.fecha', 'Factura.items[0].cantidad', 'Factura.logo', 'Factura.total', 'Pedido[0].pagado'].sort());
  const ok = A.validateBody(t, { Factura: { cliente: 5, total: 1, logo: 'data:image/jpeg;base64,/9j/4AAQ' }, Pedido: [{ folio: 'a' }] });
  assert.equal(ok.valid, true, JSON.stringify(ok.errors));
});

test('validateBody: errores de forma y advertencias', () => {
  let r = A.validateBody(base, { Factura: [], Pedido: 'x' });
  assert.deepEqual(r.errors.filter(e => e.code === 'shape').map(e => e.path).sort(), ['Factura', 'Pedido']);
  r = A.validateBody(base, { Factura: { cliente: 'a', total: 1, items: { a: 1 } }, Pedido: [5] });
  assert.deepEqual(r.errors.map(e => [e.code, e.path]).sort(), [['shape', 'Factura.items'], ['shape', 'Pedido[0]']]);
  r = A.validateBody(base, { Factura: { cliente: 'a', total: 1, items: [null] }, Pedido: [{ folio: '1' }] });
  assert.equal(r.errors[0].path, 'Factura.items[0]');
  // Objeto en hoja repetible → advertencia, se trata como un elemento
  r = A.validateBody(base, { Factura: { cliente: 'a', total: 1, extra: 1 }, Pedido: { folio: 'x' }, Fantasma: {} });
  assert.equal(r.valid, true);
  const wp = r.warnings.map(w => w.path).sort();
  assert.deepEqual(wp, ['Factura.extra', 'Fantasma', 'Pedido']);
  assert.ok(r.warnings.every(w => w.code === 'unknown'));
  // Variable de la zona puesta a nivel de hoja → advertencia con pista
  r = A.validateBody(base, { Factura: { cliente: 'a', total: 1, cantidad: 3 }, Pedido: [{ folio: '1' }] });
  assert.equal(r.valid, true);
  assert.equal(r.warnings.length, 1);
  assert.match(r.warnings[0].message, /zona de filas/);
  assert.match(r.warnings[0].message, /Factura\.items/);
});

test('validateBody: entradas basura nunca lanzan', () => {
  for (const b of [null, undefined, 0, 'texto', [], [1, 2], true]) {
    const r = A.validateBody(base, b);
    assert.equal(r.valid, false);
    assert.equal(r.errors[0].code, 'shape');
  }
  for (const t of [null, 'x', { sheets: null }]) assert.doesNotThrow(() => A.validateBody(t, { a: 1 }));
  const evil = JSON.parse('{"__proto__": {"x": 1}, "Factura": {"__proto__": {"cliente": "a"}, "constructor": 1}}');
  const r = A.validateBody(base, evil);
  assert.ok(r.errors.some(e => e.path === 'Factura.cliente'), 'no lee claves heredadas');
  assert.equal({}.x, undefined);
});

test('bodyFromQuery: coerción, reservados y claves peligrosas', () => {
  const t = withImplicit();
  const q = JSON.parse('{"cliente":"Ana","total":"1160.5","vendedor":"Luis","format":"base64","filename":"x","strict":"1","download":"1","folio":"P-1","cantidad":"3","nada":"z","__proto__":"x","constructor":"y","fecha":["2026-01-01","2027"]}');
  const body = A.bodyFromQuery(t, q);
  assert.deepEqual(body, { Factura: { cliente: 'Ana', total: 1160.5, vendedor: 'Luis', fecha: '2026-01-01' } });
  assert.equal(Object.getPrototypeOf(body), Object.prototype);
  // booleanos (hoja no repetible)
  const t2 = clone(base); t2.sheets[1].isRepeatable = false;
  assert.deepEqual(A.bodyFromQuery(t2, { pagado: 'sí' }), { Pedido: { pagado: true } });
  assert.deepEqual(A.bodyFromQuery(t2, { pagado: '0' }), { Pedido: { pagado: false } });
  assert.deepEqual(A.bodyFromQuery(t2, { pagado: 'tal vez', total: 'abc' }), { Pedido: { pagado: 'tal vez' }, Factura: { total: 'abc' } });
  assert.deepEqual(A.bodyFromQuery(base, null), {});
  assert.deepEqual(A.bodyFromQuery(null, { a: 1 }), {});
});

test('safeFileName: casos límite', () => {
  assert.equal(A.safeFileName('Factura Ana'), 'Factura Ana.xlsx');
  assert.equal(A.safeFileName('Factura.xlsx'), 'Factura.xlsx');
  assert.equal(A.safeFileName('a/b\\c:d*e?f"g<h>i|j'), 'a_b_c_d_e_f_g_h_i_j.xlsx');
  assert.equal(A.safeFileName('con\u0000tro\u001fl\n'), 'control.xlsx');
  assert.equal(A.safeFileName(''), 'resultado.xlsx');
  assert.equal(A.safeFileName(null, 'Mi plantilla'), 'Mi plantilla.xlsx');
  assert.equal(A.safeFileName('   ', '   '), 'resultado.xlsx');
  assert.equal(A.safeFileName('...'), 'resultado.xlsx');
  assert.equal(A.safeFileName('CON'), '_CON.xlsx');
  assert.equal(A.safeFileName('lpt1.txt'), '_lpt1.txt.xlsx');
  assert.equal(A.safeFileName('Año señal Ñandú café'), 'Año señal Ñandú café.xlsx');
  const long = A.safeFileName('x'.repeat(500));
  assert.ok(long.length <= 120 && long.endsWith('.xlsx'));
  const longAcc = A.safeFileName('é'.repeat(300));
  assert.ok(longAcc.length <= 120 && longAcc.endsWith('.xlsx'));
  // Emoji en el límite: nunca deja sustitutos sueltos (encodeURIComponent lanzaría URIError)
  for (let pad = 0; pad < 4; pad++) {
    const emo = A.safeFileName('a'.repeat(pad) + '😀'.repeat(100));
    assert.ok(emo.length <= 120 && emo.endsWith('.xlsx'));
    assert.doesNotThrow(() => encodeURIComponent(emo));
  }
  assert.doesNotThrow(() => encodeURIComponent(A.safeFileName('mal\ud800formado\udc00')));
  // Controles bidireccionales / invisibles eliminados
  assert.equal(A.safeFileName('factura‮slx.exe'), 'facturaslx.exe.xlsx');
  assert.equal(A.safeFileName(12345), '12345.xlsx');
});

test('fillPattern: valores aplanados', () => {
  const body = { Factura: { cliente: 'Ana', total: 10, items: [{ descripcion: 'X' }] }, Pedido: [{ folio: 'P-1' }] };
  assert.equal(A.fillPattern('Factura {{cliente}} {{folio}} {{descripcion}}', body), 'Factura Ana P-1 X');
  assert.equal(A.fillPattern('F {{nada}} {{total}}', body), 'F 10');
  assert.equal(A.fillPattern('Fijo', null), 'Fijo');
  assert.equal(A.fillPattern(null, body), '');
  const cyc = { a: { n: 'z' } }; cyc.a.self = cyc;
  assert.equal(A.fillPattern('{{n}}', cyc), 'z');
  assert.equal(A.fillPattern('{{x}}', JSON.parse('{"__proto__":{"x":"mal"}}')), '');
});

test('buildOpenApi: documenta todas las rutas', () => {
  const doc = A.buildOpenApi({ baseUrl: 'http://localhost:4104' });
  assert.equal(doc.openapi, '3.1.0');
  assert.equal(doc.servers[0].url, 'http://localhost:4104');
  const expected = {
    '/api/health': ['get'], '/api/templates': ['get', 'post'], '/api/templates/{id}': ['get', 'delete'],
    '/api/templates/{id}/thumbnail': ['put'], '/api/templates/{id}/duplicate': ['post'], '/api/templates/{id}/schema': ['get'],
    '/api/templates/{id}/example': ['get'], '/api/templates/{id}/validate': ['post'], '/api/import': ['post'],
    '/api/render/{id}': ['post', 'get'], '/api/render/{id}/batch': ['post'], '/api/openapi.json': ['get'],
    '/api/live/stream': ['get'], '/api/live/presence': ['get', 'post'], '/api/live/leave': ['post'],
    '/api/usage': ['get'], '/api/usage/events': ['get'], '/api/usage/{id}': ['get'],
  };
  for (const [p, methods] of Object.entries(expected)) for (const m of methods) assert.ok(doc.paths[p]?.[m], `${m} ${p}`);
  assert.ok(doc.paths['/api/templates'].post.responses[409]);
  assert.ok(doc.paths['/api/render/{id}'].post.responses[422]);
  for (const n of ['TemplateSummary', 'ValidationResult', 'UsageSummary', 'PresenceEntry', 'Error']) assert.ok(doc.components.schemas[n], n);
  assert.deepEqual(doc.tags.map(t => t.name), ['Plantillas', 'Generación', 'Automatización', 'Tiempo real', 'Uso']);
  // Todas las $ref resuelven
  const refs = JSON.stringify(doc).match(/#\/components\/schemas\/\w+/g);
  for (const r of refs) assert.ok(doc.components.schemas[r.split('/').pop()], r);
  for (const r of JSON.stringify(doc).match(/#\/components\/parameters\/\w+/g) || []) assert.ok(doc.components.parameters[r.split('/').pop()], r);
  // operationId únicos y cada operación con al menos una respuesta 2xx
  const ops = Object.values(doc.paths).flatMap(p => Object.values(p));
  const opIds = ops.map(o => o.operationId);
  assert.equal(new Set(opIds).size, opIds.length);
  for (const o of ops) assert.ok(Object.keys(o.responses).some(c => /^2/.test(c)), o.operationId);
  assert.ok(doc.paths['/api/render/{id}/batch'].post.parameters.some(p => p.name === 'strict'));
  assert.doesNotThrow(() => A.buildOpenApi());
});

// ── Zonas repetibles (filas y columnas) ──────────────────────────────────────
function withZones() {
  const v = (name, type, required) => ({ value: `{{${name}}}`, variable: { name, type, hint: '', required, defaultValue: '' } });
  return {
    id: 'zonas', name: 'Ventas',
    sheets: [{
      name: 'Hoja', isRepeatable: false,
      cells: {
        A1: { value: 'Ventas {{anio}}' }, B1: v('mes', 'text', true), C1: { value: 'Total' },
        B2: v('ventas', 'number', true), C2: { value: '', formula: '=SUM(B2:B2)' },
        A4: v('producto', 'text', true), C4: v('cantidad', 'number', false), A6: v('responsable', 'text', false),
      },
      repeatZones: [
        { id: 'meses', direction: 'cols', start: 2, end: 2, arrayVariable: 'meses' },
        { id: 'prods', direction: 'rows', start: 4, end: 4, arrayVariable: 'productos' },
      ],
      rowTemplate: { startRow: 9, endRow: 9, arrayVariable: 'ignorado' },
    }],
  };
}

test('buildSchema: repeatZones y zona de cada variable', () => {
  const s = A.buildSchema(withZones()).sheets[0];
  assert.deepEqual(s.repeatZones.map(z => [z.direction, z.start, z.end, z.arrayVariable]), [['cols', 2, 2, 'meses'], ['rows', 4, 4, 'productos']]);
  assert.equal(s.repeatZones[0].label, 'columna B');
  assert.deepEqual(s.rowTemplate, { startRow: 4, endRow: 4, arrayVariable: 'productos' }, 'compatibilidad: primera zona de filas');
  assert.deepEqual(findVar(s, 'mes').zone, { arrayVariable: 'meses', direction: 'cols', id: 'meses' });
  assert.equal(findVar(s, 'ventas').inRowTemplate, true);
  assert.equal(findVar(s, 'producto').zone.direction, 'rows');
  assert.equal(findVar(s, 'responsable').zone, null);
  assert.equal(findVar(s, 'anio').zone, null);
  assert.deepEqual(s.zoneIssues, []);
  // El rowTemplate heredado se sigue entendiendo como zona
  const legacy = A.buildSchema(base).sheets[0];
  assert.deepEqual(legacy.repeatZones.map(z => [z.direction, z.start, z.arrayVariable]), [['rows', 5, 'items']]);
  assert.deepEqual(findVar(legacy, 'cantidad').zone, { arrayVariable: 'items', direction: 'rows', id: 'rowTemplate' });
});

test('buildExample: un arreglo de 2 elementos por zona y siempre válido', () => {
  const ex = A.buildExample(withZones());
  assert.equal(ex.Hoja.meses.length, 2);
  assert.deepEqual(Object.keys(ex.Hoja.meses[0]).sort(), ['mes', 'ventas']);
  assert.deepEqual(Object.keys(ex.Hoja.productos[0]).sort(), ['cantidad', 'producto']);
  assert.ok('anio' in ex.Hoja && 'responsable' in ex.Hoja);
  assert.ok(!('mes' in ex.Hoja));
  assert.equal(typeof ex.Hoja.meses[1].ventas, 'number');
  assert.equal(A.validateBody(withZones(), ex).valid, true);
});

test('validateBody: errores por elemento con rutas Hoja.meses[1].ventas', () => {
  const r = A.validateBody(withZones(), { Hoja: {
    meses: [{ mes: 'Ene', ventas: 10 }, { mes: 'Feb', ventas: 'mucho' }, 'x'],
    productos: { producto: 'no es arreglo' },
    ventas: 5,
  } });
  assert.equal(r.valid, false);
  const byPath = Object.fromEntries(r.errors.map(e => [e.path, e.code]));
  assert.equal(byPath['Hoja.meses[1].ventas'], 'type');
  assert.equal(byPath['Hoja.meses[2]'], 'shape');
  assert.equal(byPath['Hoja.productos'], 'shape');
  assert.match(r.errors.find(e => e.path === 'Hoja.productos').message, /zona de filas/);
  const w = r.warnings.find(x => x.path === 'Hoja.ventas');
  assert.match(w.message, /zona de columnas/);
  assert.match(w.message, /Hoja\.meses/);
  const miss = A.validateBody(withZones(), { Hoja: { meses: [{ ventas: 1 }] } });
  assert.ok(miss.errors.some(e => e.path === 'Hoja.meses[0].mes' && e.code === 'missing'));
  assert.ok(miss.errors.some(e => e.path === 'Hoja.productos' && e.code === 'missing'));
});

test('validateBody: avisa zonas superpuestas o cruzadas', () => {
  const t = withZones();
  t.sheets[0].repeatZones.push({ direction: 'cols', start: 2, end: 3, arrayVariable: 'otra' });
  t.sheets[0].cells.B4 = { value: 'cruce' };
  t.sheets[0].cells.A4.variable.required = false;
  t.sheets[0].cells.B1.variable.required = false;
  t.sheets[0].cells.B2.variable.required = false;
  const r = A.validateBody(t, { Hoja: {} });
  const codes = r.warnings.filter(x => x.code === 'zone').map(x => x.message);
  assert.ok(codes.some(m => /se superpone/.test(m)), codes.join(' | '));
  assert.ok(codes.some(m => /se cruzan en B4/.test(m)), codes.join(' | '));
});

test('buildOpenApi: documenta RepeatZone', () => {
  const doc = A.buildOpenApi({});
  assert.deepEqual(doc.components.schemas.RepeatZone.properties.direction.enum, ['rows', 'cols']);
  assert.ok(doc.components.schemas.ValidationIssue.properties.code.enum.includes('zone'));
});
