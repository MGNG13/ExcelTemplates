'use strict';
// Comprueba que los scripts de la interfaz (sin paso de compilación) no tengan errores de sintaxis.
// Ejecutar: node --test test/ui-syntax.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const PUB = path.join(__dirname, '..', 'Public');

test('los <script> en línea de index.html compilan', () => {
  const html = fs.readFileSync(path.join(PUB, 'index.html'), 'utf8');
  const blocks = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
  assert.ok(blocks.length > 0, 'se esperaba al menos un script en línea');
  blocks.forEach((code, i) => assert.doesNotThrow(() => new vm.Script(code, { filename: `index.html#script${i}` })));
});

test('los módulos de Public/js compilan', () => {
  for (const f of fs.readdirSync(path.join(PUB, 'js')).filter(f => f.endsWith('.js'))) {
    const code = fs.readFileSync(path.join(PUB, 'js', f), 'utf8');
    assert.doesNotThrow(() => new vm.Script(code, { filename: f }), f);
  }
});
