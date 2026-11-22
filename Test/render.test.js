'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const AdmZip = require('adm-zip');
const { renderWorkbook } = require('../Library/render');

// Same template as the shared fixture (fixtures/template-basic.json), inlined so the test is self-contained
const v = (name, type, extra = {}) => ({ value: `{{${name}}}`, variable: { name, type, hint: '', required: false, defaultValue: '', ...extra } });
const bold = value => ({ value, style: { bold: true } });
const FIXTURE = {
  id: 'fixture-factura', name: 'Factura de prueba',
  sheets: [
    { id: 's1', name: 'Factura',
      cells: { A1: bold('Cliente:'), B1: v('cliente', 'text', { required: true }), A2: bold('Fecha:'), B2: v('fecha', 'date'),
        A4: bold('Descripción'), B4: bold('Cantidad'), C4: bold('Precio'),
        A5: v('descripcion', 'text'), B5: v('cantidad', 'number'), C5: v('precio', 'number'),
        A7: bold('Total:'), C7: v('total', 'number', { required: true }) },
      columnWidths: { A: 160, B: 100, C: 100 }, rowHeights: {}, merges: [], isRepeatable: false, repeatVariable: '',
      rowTemplate: { startRow: 5, endRow: 5, arrayVariable: 'items' }, floatingImages: [], shapes: [], hiddenRows: [], hiddenCols: [], state: 'visible' },
    { id: 's2', name: 'Pedido',
      cells: { A1: bold('Pedido:'), B1: v('folio', 'text', { required: true }), A2: bold('Pagado:'), B2: v('pagado', 'boolean') },
      columnWidths: {}, rowHeights: {}, merges: [], isRepeatable: true, repeatVariable: 'pedidos', rowTemplate: null,
      floatingImages: [], shapes: [], hiddenRows: [], hiddenCols: [], state: 'visible' },
  ],
};
const BODY = {
  Factura: { cliente: 'Ana', fecha: '2026-11-19', total: 123.5,
    items: [{ descripcion: 'Uno', cantidad: 2, precio: 10 }, { descripcion: 'Dos', cantidad: 1, precio: 5 }] },
  Pedido: [{ folio: 'P-1', pagado: true }, { folio: 'P-2', pagado: false }],
};

test('renders the fixture into a valid xlsx', async () => {
  const template = structuredClone(FIXTURE);
  const buf = await renderWorkbook(template, BODY);
  assert.ok(Buffer.isBuffer(buf) && buf.length > 1000);
  assert.deepEqual(template, FIXTURE, 'the template must not be mutated');
  const zip = new AdmZip(buf);
  const wbXml = zip.readAsText('xl/workbook.xml');
  for (const n of ['Factura', 'Pedido 1', 'Pedido 2']) assert.match(wbXml, new RegExp(`name="${n}"`));
  const texts = zip.getEntries()
    .filter(e => /^xl\/(sharedStrings|worksheets\/sheet\d+)\.xml$/.test(e.entryName))
    .map(e => e.getData().toString('utf8')).join('');
  assert.match(texts, /Ana/);
  assert.match(texts, /Dos/);   // second row-template item
});

test('single repeatable item keeps the sheet name; empty body renders', async () => {
  const one = new AdmZip(await renderWorkbook(FIXTURE, { Pedido: [{ folio: 'X' }] })).readAsText('xl/workbook.xml');
  assert.match(one, /name="Pedido"/);
  assert.doesNotMatch(one, /Pedido 1/);
  const empty = await renderWorkbook(FIXTURE, {});
  assert.ok(empty.length > 1000);
});

test('rejects invalid templates', async () => {
  await assert.rejects(renderWorkbook(null, {}), /inválida/);
  await assert.rejects(renderWorkbook({ sheets: 'x' }, {}), /inválida/);
});

// ── Zonas repetibles (filas y columnas) ──────────────────────────────────────
const ExcelJS = require('exceljs');
const { mapFormulaRefs, adjustFormulaCols, normalizeZones } = require('../Library/repeat-zones');
const { adjustFormulaRows } = require('../Library/excel-utils');

async function load(buf) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}
const val = c => (c.value && typeof c.value === 'object' && 'formula' in c.value) ? '=' + c.value.formula : c.value;
const merges = ws => Object.keys(ws._merges || {}).map(k => ws._merges[k].range || String(ws._merges[k])).sort();
const sheetOf = (cells, extra = {}) => ({ id: 'z', name: 'Hoja', cells, columnWidths: {}, rowHeights: {}, merges: [], hiddenRows: [], hiddenCols: [], ...extra });

test('zona de filas con 3 elementos: copia filas, desplaza lo siguiente y el total crece', async () => {
  const tpl = { name: 'T', sheets: [sheetOf({
    A1: bold('Producto'), B1: bold('Monto'),
    A2: v('producto', 'text'), B2: v('monto', 'number'), C2: { value: '', formula: '=B2*2' }, D2: v('moneda', 'text'),
    A4: bold('Total'), B4: { value: '', formula: '=SUMA(B2:B2)' }, C4: { value: '', formula: '=B4+$B$2+B5' },
    A5: { value: 'Fin {{cliente}}' }, A6: { value: 'oculta' },
  }, { rowHeights: { 2: 30, 5: 40 }, merges: ['A5:C5'], hiddenRows: [6],
       repeatZones: [{ id: 'z1', direction: 'rows', start: 2, end: 2, arrayVariable: 'lineas' }] })] };
  const body = { Hoja: { cliente: 'Ana', moneda: 'MXN', lineas: [{ producto: 'A', monto: 1 }, { producto: 'B', monto: 2 }, { producto: 'C', monto: 3, moneda: 'USD' }] } };
  const ws = (await load(await renderWorkbook(tpl, body))).getWorksheet('Hoja');
  assert.deepEqual([2, 3, 4].map(r => ws.getCell('A' + r).value), ['A', 'B', 'C']);
  assert.deepEqual([2, 3, 4].map(r => ws.getCell('B' + r).value), [1, 2, 3]);
  assert.deepEqual([2, 3, 4].map(r => val(ws.getCell('C' + r))), ['=B2*2', '=B3*2', '=B4*2'], 'refs relativas por copia');
  assert.deepEqual([2, 3, 4].map(r => ws.getCell('D' + r).value), ['MXN', 'MXN', 'USD'], 'respaldo en datos de la hoja');
  assert.equal(ws.getCell('A6').value, 'Total');
  assert.equal(val(ws.getCell('B6')), '=SUM(B2:B4)', 'el rango que cubre la zona abarca todas las copias');
  assert.equal(val(ws.getCell('C6')), '=B6+$B$2+B7', 'refs posteriores se desplazan; absolutas a la zona se quedan');
  assert.equal(ws.getCell('A7').value, 'Fin Ana');
  assert.deepEqual([2, 3, 4].map(r => ws.getRow(r).height), [22.5, 22.5, 22.5]);
  assert.equal(ws.getRow(7).height, 30);
  assert.ok(merges(ws).includes('A7:C7'), 'combinación desplazada: ' + merges(ws));
  assert.equal(ws.getRow(8).hidden, true, 'fila oculta desplazada');
});

test('zona de columnas con 4 meses: copia columnas, anchos y combinaciones; desplaza lo siguiente', async () => {
  const tpl = { name: 'T', sheets: [sheetOf({
    A1: bold('Concepto'), B1: v('mes', 'text'), C1: bold('Total'), D1: { value: '{{nota}}' },
    A2: bold('Ventas'), B2: v('ventas', 'number'), C2: { value: '', formula: '=SUM(B2:B2)' }, D2: { value: '', formula: '=C2*2' },
    B3: { value: 'x' },
  }, { columnWidths: { A: 120, B: 75, C: 110 }, merges: ['B3:B4'], hiddenCols: [4],
       repeatZones: [{ id: 'm', direction: 'cols', start: 2, end: 2, arrayVariable: 'meses' }] })] };
  const meses = ['Ene', 'Feb', 'Mar', 'Abr'].map((mes, i) => ({ mes, ventas: (i + 1) * 100 }));
  const ws = (await load(await renderWorkbook(tpl, { Hoja: { nota: 'ok', meses } }))).getWorksheet('Hoja');
  assert.deepEqual(['B', 'C', 'D', 'E'].map(c => ws.getCell(c + '1').value), ['Ene', 'Feb', 'Mar', 'Abr']);
  assert.deepEqual(['B', 'C', 'D', 'E'].map(c => ws.getCell(c + '2').value), [100, 200, 300, 400]);
  assert.equal(ws.getCell('F1').value, 'Total');
  assert.equal(val(ws.getCell('F2')), '=SUM(B2:E2)');
  assert.equal(val(ws.getCell('G2')), '=F2*2');
  assert.equal(ws.getCell('G1').value, 'ok');
  const w = c => Math.round(ws.getColumn(c).width * 100) / 100;
  assert.deepEqual([2, 3, 4, 5].map(w), [w(2), w(2), w(2), w(2)]);
  assert.ok(Math.abs(w(2) - 10) < 0.01, 'ancho de la zona copiado: ' + w(2));
  assert.ok(Math.abs(w(6) - 15) < 0.01, 'ancho de Total desplazado: ' + w(6));
  assert.equal(ws.getColumn(7).hidden, true, 'columna oculta desplazada');
  for (const m of ['B3:B4', 'C3:C4', 'D3:D4', 'E3:E4']) assert.ok(merges(ws).includes(m), m + ' en ' + merges(ws));
});

test('zonas de filas y columnas en la misma hoja', async () => {
  const tpl = { name: 'T', sheets: [sheetOf({
    A1: bold('Año'), B1: v('mes', 'text'), C1: bold('Fin'),
    A3: v('producto', 'text'), A4: { value: '', formula: '=COUNTA(A3:A3)' },
  }, { repeatZones: [
    { id: 'c', direction: 'cols', start: 2, end: 2, arrayVariable: 'meses' },
    { id: 'r', direction: 'rows', start: 3, end: 3, arrayVariable: 'productos' },
  ] })] };
  const warnings = [];
  const buf = await renderWorkbook(tpl, { Hoja: { meses: [{ mes: 'Ene' }, { mes: 'Feb' }, { mes: 'Mar' }], productos: [{ producto: 'P1' }, { producto: 'P2' }] } }, { warnings });
  assert.deepEqual(warnings, []);
  const ws = (await load(buf)).getWorksheet('Hoja');
  assert.deepEqual(['B1', 'C1', 'D1', 'E1'].map(a => ws.getCell(a).value), ['Ene', 'Feb', 'Mar', 'Fin']);
  assert.deepEqual(['A3', 'A4'].map(a => ws.getCell(a).value), ['P1', 'P2']);
  assert.equal(val(ws.getCell('A5')), '=COUNTA(A3:A4)');
});

test('zona sin elementos: queda una vez, sin desplazar; datos que no son arreglo → aviso', async () => {
  const tpl = { name: 'T', sheets: [sheetOf({ A1: v('x', 'text'), A2: { value: 'Fin' } },
    { repeatZones: [{ direction: 'rows', start: 1, end: 1, arrayVariable: 'items' }] })] };
  let ws = (await load(await renderWorkbook(tpl, { Hoja: { items: [] } }))).getWorksheet('Hoja');
  assert.equal(ws.getCell('A2').value, 'Fin');
  const warnings = [];
  ws = (await load(await renderWorkbook(tpl, { Hoja: { items: 'no' } }, { warnings }))).getWorksheet('Hoja');
  assert.equal(ws.getCell('A2').value, 'Fin');
  assert.equal(warnings[0].code, 'zone-data');
});

test('rowTemplate heredado genera exactamente lo mismo que la zona equivalente', async () => {
  const modern = structuredClone(FIXTURE);
  modern.sheets[0].repeatZones = [{ id: 'items', direction: 'rows', start: 5, end: 5, arrayVariable: 'items' }];
  modern.sheets[0].rowTemplate = { startRow: 99, endRow: 99, arrayVariable: 'ignorado' };   // repeatZones manda
  const parts = buf => new Map(new AdmZip(buf).getEntries().filter(e => e.entryName !== 'docProps/core.xml').map(e => [e.entryName, e.getData().toString('utf8')]));
  const a = parts(await renderWorkbook(FIXTURE, BODY)), b = parts(await renderWorkbook(modern, BODY));
  assert.deepEqual([...a.keys()].sort(), [...b.keys()].sort());
  for (const [k, x] of a) assert.equal(b.get(k), x, k);
  const ws = (await load(await renderWorkbook(FIXTURE, BODY))).getWorksheet('Factura');
  assert.deepEqual([ws.getCell('A5').value, ws.getCell('A6').value, ws.getCell('C8').value], ['Uno', 'Dos', 123.5]);
});

test('zonas superpuestas se ignoran con aviso; las cruzadas con contenido se avisan', async () => {
  const sheet = sheetOf({ B2: { value: 'x' } }, { repeatZones: [
    { direction: 'rows', start: 2, end: 4, arrayVariable: 'a' }, { direction: 'rows', start: 4, end: 5, arrayVariable: 'b' },
    { direction: 'cols', start: 2, end: 2, arrayVariable: 'c' }, { direction: 'cols', start: 3, end: 1, arrayVariable: 'd' } ] });
  const { zones, issues } = normalizeZones(sheet);
  assert.deepEqual(zones.map(z => z.arrayVariable), ['a', 'c']);
  assert.deepEqual(issues.map(i => i.code).sort(), ['zone-cross', 'zone-invalid', 'zone-overlap']);
  const warnings = [];
  await renderWorkbook({ sheets: [sheet] }, { Hoja: { a: [{}, {}] } }, { warnings });
  assert.ok(warnings.some(w => w.code === 'zone-overlap' && w.sheet === 'Hoja'));
  // columnas por letra
  assert.equal(normalizeZones({ repeatZones: [{ direction: 'cols', start: 'C', end: 'E', arrayVariable: 'm' }] }).zones[0].start, 3);
});

test('mapFormulaRefs / adjustFormulaCols: respetan cadenas, otras hojas, funciones y $', () => {
  const plus = r => ({ col: r.col === undefined ? undefined : r.col + 1, row: r.row === undefined ? undefined : r.row + 10 });
  assert.equal(mapFormulaRefs('SUM(B4:B5)*$C$2', plus), 'SUM(C14:C15)*$D$12');
  assert.equal(mapFormulaRefs('IF(A1>0,"B5",Hoja2!B5)', plus), 'IF(B11>0,"B5",Hoja2!B5)');
  assert.equal(mapFormulaRefs("'Mi hoja'!C3+LOG10(A1)+{{A1}}", plus), "'Mi hoja'!C3+LOG10(B11)+{{A1}}");
  assert.equal(mapFormulaRefs('SUM(5:7)+SUM(C:D)', plus), 'SUM(15:17)+SUM(D:E)');
  const same = 'SUMA(B4:B5;C$1)&"x"';
  assert.equal(mapFormulaRefs(same, r => ({ col: r.col, row: r.row })), same, 'sin cambios → texto idéntico');
  assert.equal(adjustFormulaCols('B5*$C5+SUM(D1:E1)', 2), 'D5*$C5+SUM(F1:G1)');
  assert.equal(adjustFormulaRows('B5*C$5', 2), 'B7*C$5', 'adjustFormulaRows sin cambios');
});
