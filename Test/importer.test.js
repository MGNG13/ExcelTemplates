'use strict';
// Pruebas del importador (Library/importer.js). Ejecutar: node --test test/importer.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const { importWorkbook } = require('../Library/importer');

test('las celdas ocultas de una combinación no repiten el valor de la celda principal', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Ventas');
  ws.mergeCells('B15:C15');
  ws.getCell('B15').value = '{{producto}}';
  ws.getCell('C15').border = { top: { style: 'thin' } };
  ws.mergeCells('A1:D1');
  ws.getCell('A1').value = 'Título';
  const buf = Buffer.from(await wb.xlsx.writeBuffer());
  const t = await importWorkbook(buf, 'Cotización.xlsx');
  const c = t.sheets[0].cells;
  assert.equal(t.name, 'Cotización');
  assert.equal(c.B15.value, '{{producto}}');
  assert.equal(c.C15?.value ?? '', '', 'C15 forma parte de la combinación: solo conserva su estilo');
  assert.equal(c.A1.value, 'Título');
  for (const a of ['B1', 'C1', 'D1']) assert.equal(c[a]?.value ?? '', '', a);
  assert.ok(t.sheets[0].merges.includes('B15:C15'));
});
