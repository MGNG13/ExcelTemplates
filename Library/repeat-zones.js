// Zonas repetibles (filas y columnas) — normalización, validación y expansión.
//
// Modelo (en cada hoja de la plantilla):
//   sheet.repeatZones = [{ id, direction: 'rows'|'cols', start, end, arrayVariable }]
//     · rows → start/end son números de fila (1-based)
//     · cols → start/end son números de columna (1-based); también se aceptan letras ('C')
//   Compatibilidad: si la hoja NO tiene `repeatZones` (array) se usa el antiguo
//   `rowTemplate = { startRow, endRow, arrayVariable }` como una zona de filas.
//   Si `repeatZones` existe es la fuente de verdad (rowTemplate se ignora).
//
// Datos: sheetData[arrayVariable] es un array de objetos; la zona se copia una vez por
// elemento. Las variables dentro de la zona leen del elemento y, si no está, de los
// datos de la hoja. Con 0 elementos (o sin la clave) la zona queda UNA vez, con las
// variables que no estén en la hoja vacías (mismo comportamiento que el rowTemplate v11).
//
// Sin dependencias de ExcelJS: funciones puras usadas por render.js y automation.js.
'use strict';

const MAX_ROWS = 1048576;
const MAX_COLS = 16384;
const FORBIDDEN = new Set(['__proto__', 'constructor', 'prototype']);

const isObj = v => v !== null && typeof v === 'object' && !Array.isArray(v);

function colToIdx(letter) {
  return String(letter).toUpperCase().split('').reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0);
}
function idxToCol(idx) {
  let r = '';
  while (idx > 0) { const m = (idx - 1) % 26; r = String.fromCharCode(65 + m) + r; idx = Math.floor((idx - 1) / 26); }
  return r;
}
function parseCell(a) {
  const m = /^([A-Z]{1,3})(\d{1,7})$/i.exec(String(a));
  return m ? { col: colToIdx(m[1]), row: parseInt(m[2], 10) } : null;
}
function parseRange(r) {
  const m = /^([A-Z]{1,3})(\d{1,7}):([A-Z]{1,3})(\d{1,7})$/i.exec(String(r));
  if (!m) return null;
  const c1 = colToIdx(m[1]), r1 = +m[2], c2 = colToIdx(m[3]), r2 = +m[4];
  return { sc: Math.min(c1, c2), sr: Math.min(r1, r2), ec: Math.max(c1, c2), er: Math.max(r1, r2) };
}

function zoneLabel(z) {
  return z.direction === 'cols'
    ? (z.start === z.end ? `columna ${idxToCol(z.start)}` : `columnas ${idxToCol(z.start)}–${idxToCol(z.end)}`)
    : (z.start === z.end ? `fila ${z.start}` : `filas ${z.start}–${z.end}`);
}

function toIndex(v, direction) {
  if (typeof v === 'string' && direction === 'cols' && /^[A-Z]{1,3}$/i.test(v.trim())) return colToIdx(v.trim());
  const n = Number(v);
  return Number.isInteger(n) ? n : NaN;
}

/** Zonas tal como vienen en la plantilla (repeatZones o rowTemplate heredado). */
function rawZones(sheet) {
  if (!isObj(sheet)) return [];
  if (Array.isArray(sheet.repeatZones)) return sheet.repeatZones;
  const rt = sheet.rowTemplate;
  if (isObj(rt)) return [{ id: 'rowTemplate', direction: 'rows', start: rt.startRow, end: rt.endRow, arrayVariable: rt.arrayVariable, legacy: true }];
  return [];
}

/** ¿La celda tiene contenido (valor, fórmula, variable o imagen)? */
function hasContent(cd) {
  if (!isObj(cd)) return false;
  return !!(cd.formula || cd.variable || cd.image || (cd.value !== undefined && cd.value !== null && String(cd.value) !== ''));
}

/**
 * Normaliza y valida las zonas de una hoja.
 * @returns {{ zones: Array<{id,direction,start,end,arrayVariable,legacy?}>, issues: Array<{code,zone,message}> }}
 *   issues.code: 'zone-invalid' (descartada), 'zone-overlap' (descartada), 'zone-cross' (se conserva).
 */
function normalizeZones(sheet) {
  const zones = [], issues = [];
  rawZones(sheet).forEach((z, i) => {
    if (!isObj(z)) { issues.push({ code: 'zone-invalid', zone: null, message: `La zona repetible #${i + 1} no es un objeto; se ignora.` }); return; }
    const direction = z.direction === 'cols' ? 'cols' : 'rows';
    const start = toIndex(z.start ?? z.startRow ?? z.startCol, direction);
    const end = toIndex(z.end ?? z.endRow ?? z.endCol ?? z.start ?? z.startRow ?? z.startCol, direction);
    const arrayVariable = typeof z.arrayVariable === 'string' ? z.arrayVariable.trim() : '';
    const max = direction === 'cols' ? MAX_COLS : MAX_ROWS;
    const id = typeof z.id === 'string' && z.id ? z.id.slice(0, 64) : `z${i + 1}`;
    if (!arrayVariable || FORBIDDEN.has(arrayVariable) || arrayVariable.length > 64) {
      issues.push({ code: 'zone-invalid', zone: arrayVariable || null, message: `La zona repetible #${i + 1} no tiene un nombre de arreglo válido; se ignora.` });
      return;
    }
    if (!Number.isInteger(start) || !Number.isInteger(end) || start < 1 || end < start || end > max) {
      issues.push({ code: 'zone-invalid', zone: arrayVariable, message: `La zona repetible "${arrayVariable}" tiene un rango inválido; se ignora.` });
      return;
    }
    const nz = { id, direction, start, end, arrayVariable };
    if (z.legacy) nz.legacy = true;
    const clash = zones.find(o => o.direction === direction && o.start <= end && start <= o.end);
    if (clash) {
      issues.push({ code: 'zone-overlap', zone: arrayVariable,
        message: `La zona "${arrayVariable}" (${zoneLabel(nz)}) se superpone con "${clash.arrayVariable}" (${zoneLabel(clash)}); se ignora.` });
      return;
    }
    zones.push(nz);
  });
  // Una zona de filas y una de columnas se cruzan si hay celdas con contenido en la intersección
  const rowsZ = zones.filter(z => z.direction === 'rows'), colsZ = zones.filter(z => z.direction === 'cols');
  if (rowsZ.length && colsZ.length && isObj(sheet.cells)) {
    for (const r of rowsZ) for (const c of colsZ) {
      const hit = Object.keys(sheet.cells).find(a => {
        const p = parseCell(a);
        return p && p.row >= r.start && p.row <= r.end && p.col >= c.start && p.col <= c.end && hasContent(sheet.cells[a]);
      });
      if (hit) issues.push({ code: 'zone-cross', zone: c.arrayVariable,
        message: `Las zonas "${r.arrayVariable}" (${zoneLabel(r)}) y "${c.arrayVariable}" (${zoneLabel(c)}) se cruzan en ${hit}: esas celdas se repetirán en ambas direcciones.` });
    }
  }
  return { zones, issues };
}

/** Zona (de la lista) que contiene la celda; prioridad a filas si cae en ambas. */
function zoneOfCell(zones, addr) {
  const p = parseCell(addr); if (!p) return null;
  return zones.find(z => z.direction === 'rows' && p.row >= z.start && p.row <= z.end)
      || zones.find(z => z.direction === 'cols' && p.col >= z.start && p.col <= z.end) || null;
}

// ── Referencias en fórmulas ──────────────────────────────────────────────────

const CELL_RE = /(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7}))?/y;
const ROWS_RE = /(\$?)([0-9]{1,7}):(\$?)([0-9]{1,7})/y;
const COLS_RE = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/y;
const WORDCH = /[A-Za-z0-9_.]/;
const AFTER_BAD = /[A-Za-z0-9_(!.]/;

function matchAt(re, f, i) { re.lastIndex = i; const m = re.exec(f); return m && m.index === i ? m : null; }

/**
 * Recorre las referencias A1 de una fórmula (sin '=') y las reescribe con fn.
 * fn({ kind: 'cell'|'rows'|'cols', col, row, colAbs, rowAbs, role: 'single'|'start'|'end' }) → { col, row }
 * Se respetan cadenas "…", marcadores {{var}} y referencias a otras hojas (Hoja!A1, 'Mi hoja'!A1),
 * que no se tocan salvo con { ignoreSheet: true }. Si nada cambia devuelve el texto idéntico.
 */
function mapFormulaRefs(f, fn, opts = {}) {
  if (typeof f !== 'string' || !f) return f;
  const ignoreSheet = !!opts.ignoreSheet;
  let out = '', i = 0;
  const n = f.length;
  const fmtCell = (m, o, role) => {
    // m: [$c, letters, $r, digits]
    const col = colToIdx(m[1]), row = +m[3];
    const r = fn({ kind: 'cell', col, row, colAbs: m[0] === '$', rowAbs: m[2] === '$', role });
    if (!r || (r.col === col && r.row === row)) return o;
    return m[0] + (r.col === col ? m[1] : idxToCol(r.col)) + m[2] + r.row;
  };
  const tryRef = (at, sheetQualified) => {
    // Devuelve [texto, longitud] o null
    const passthrough = sheetQualified && !ignoreSheet;
    let m = matchAt(CELL_RE, f, at);
    if (m && !AFTER_BAD.test(f[at + m[0].length] || '')) {
      if (passthrough) return [m[0], m[0].length];
      const first = m[0].slice(0, m[1].length + m[2].length + m[3].length + m[4].length);
      if (m[5] === undefined && m[6] === undefined) return [fmtCell([m[1], m[2], m[3], m[4]], first, 'single'), m[0].length];
      const second = m[0].slice(first.length + 1);
      return [fmtCell([m[1], m[2], m[3], m[4]], first, 'start') + ':' + fmtCell([m[5], m[6], m[7], m[8]], second, 'end'), m[0].length];
    }
    m = matchAt(ROWS_RE, f, at);
    if (m && !AFTER_BAD.test(f[at + m[0].length] || '')) {
      if (passthrough) return [m[0], m[0].length];
      const one = (abs, digits, role) => {
        const row = +digits, r = fn({ kind: 'rows', row, rowAbs: abs === '$', role });
        return abs + (r && r.row !== row ? r.row : digits);
      };
      return [one(m[1], m[2], 'start') + ':' + one(m[3], m[4], 'end'), m[0].length];
    }
    m = matchAt(COLS_RE, f, at);
    if (m && !AFTER_BAD.test(f[at + m[0].length] || '')) {
      if (passthrough) return [m[0], m[0].length];
      const one = (abs, letters, role) => {
        const col = colToIdx(letters), r = fn({ kind: 'cols', col, colAbs: abs === '$', role });
        return abs + (r && r.col !== col ? idxToCol(r.col) : letters);
      };
      return [one(m[1], m[2], 'start') + ':' + one(m[3], m[4], 'end'), m[0].length];
    }
    return null;
  };
  while (i < n) {
    const ch = f[i];
    if (ch === '"') {
      let j = i + 1;
      while (j < n) { if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; } j++; }
      out += f.slice(i, j + 1); i = j + 1; continue;
    }
    if (ch === '{' && f[i + 1] === '{') {
      const j = f.indexOf('}}', i + 2);
      if (j >= 0) { out += f.slice(i, j + 2); i = j + 2; continue; }
    }
    if (ch === "'") {
      let j = i + 1;
      while (j < n) { if (f[j] === "'") { if (f[j + 1] === "'") { j += 2; continue; } break; } j++; }
      out += f.slice(i, j + 1); i = j + 1;
      if (f[i] === '!') {
        out += '!'; i++;
        const r = tryRef(i, true);
        if (r) { out += r[0]; i += r[1]; }
      }
      continue;
    }
    const prev = i > 0 ? f[i - 1] : '';
    if ((prev && WORDCH.test(prev)) || !/[A-Za-z0-9$]/.test(ch)) { out += ch; i++; continue; }
    if (prev === ':' || (prev === '$' && i > 1 && WORDCH.test(f[i - 2]))) { out += ch; i++; continue; }
    const r = tryRef(i, prev === '!');
    if (r) { out += r[0]; i += r[1]; continue; }
    out += ch; i++;
  }
  return out;
}

/** Desplaza columnas relativas (sin $) `offset` posiciones; contraparte de adjustFormulaRows. */
function adjustFormulaCols(formula, offset) {
  if (!offset) return formula;
  return mapFormulaRefs(formula, r => ({ col: (r.col !== undefined && !r.colAbs) ? Math.max(1, r.col + offset) : r.col, row: r.row }));
}

// ── Expansión ────────────────────────────────────────────────────────────────

/**
 * Mapeo de un índice (fila o columna) al expandir una zona [S,E] en N copias.
 * k ≥ 0: la referencia pertenece a la copia k (copia de celdas dentro de la zona).
 * role 'end' en ≥ E crece con la zona → los rangos que cubren la zona (SUMA(C5:C7)) abarcan todas las copias.
 */
function idxMapper(S, E, H, extra) {
  return (v, abs, role, k) => {
    if (v === undefined) return v;
    if (k >= 0 && v >= S && v <= E) return abs ? v : v + k * H;
    if (role === 'end' ? v >= E : v > E) return v + extra;
    return v;
  };
}

function numMap(o) {
  const m = new Map();
  if (isObj(o)) for (const [k, v] of Object.entries(o)) { const n = +k; if (Number.isInteger(n)) m.set(n, v); }
  return m;
}
function dupMap(map, S, E, H, N, extra) {
  const out = new Map();
  for (const [k, v] of map) {
    if (k < S) out.set(k, v);
    else if (k > E) out.set(k + extra, v);
    else for (let c = 0; c < N; c++) out.set(k + c * H, v);
  }
  return out;
}
function dupList(list, S, E, H, N, extra) {
  const out = [];
  for (const k of list) {
    if (k < S) out.push(k);
    else if (k > E) out.push(k + extra);
    else for (let c = 0; c < N; c++) out.push(k + c * H);
  }
  return out;
}

/**
 * Expande las zonas con elementos de una hoja.
 * @param {object} sheet plantilla de la hoja
 * @param {object} data  datos de la hoja
 * @param {object} [o]   { warn(issue), rowFrac(px), colFrac(px) } — rowFrac/colFrac convierten las
 *                       imágenes posicionadas por píxeles (sin ancla) a fracciones de fila/columna.
 * @returns {null | { sheet, dataAt(addr) }} null → no hay nada que expandir (render normal).
 */
function expandSheet(sheet, data, o = {}) {
  const warn = typeof o.warn === 'function' ? o.warn : () => {};
  const { zones, issues } = normalizeZones(sheet);
  issues.forEach(is => warn({ sheet: sheet.name, ...is }));
  if (!zones.length) return null;
  const d = isObj(data) ? data : {};
  const active = [];
  for (const z of zones) {
    const arr = Object.prototype.hasOwnProperty.call(d, z.arrayVariable) ? d[z.arrayVariable] : undefined;
    if (arr !== undefined && arr !== null && !Array.isArray(arr))
      warn({ sheet: sheet.name, code: 'zone-data', zone: z.arrayVariable, message: `"${z.arrayVariable}" debe ser un arreglo; la zona ${zoneLabel(z)} se deja sin repetir.` });
    if (Array.isArray(arr) && arr.length) active.push({ ...z, items: arr });
  }
  if (!active.length) return null;

  // Estado de trabajo
  let cells = [];
  for (const [a, cd] of Object.entries(sheet.cells || {})) {
    const p = parseCell(a);
    if (p && cd) cells.push({ col: p.col, row: p.row, cd, rowItem: undefined, colItem: undefined });
  }
  let rowHeights = numMap(sheet.rowHeights), autoRowHeights = numMap(sheet.autoRowHeights);
  let colWidths = new Map(Object.entries(sheet.columnWidths || {}).map(([l, v]) => [colToIdx(l), v]));
  let colStyles = new Map(Object.entries(sheet.colStyles || {}).map(([l, v]) => [colToIdx(l), v]));
  const rawMerges = [];
  let merges = [];
  (sheet.merges || []).forEach(m => { const r = parseRange(m); if (r) merges.push(r); else rawMerges.push(m); });
  let hiddenRows = (sheet.hiddenRows || []).map(Number).filter(Number.isInteger);
  let hiddenCols = (sheet.hiddenCols || []).map(c => (typeof c === 'string' && /^[A-Z]+$/i.test(c)) ? colToIdx(c) : Number(c)).filter(Number.isInteger);
  let rowOutline = numMap(sheet.rowOutline), colOutline = numMap(sheet.colOutline);
  const images = (sheet.floatingImages || []).map(fi => {
    const c = JSON.parse(JSON.stringify(fi));
    if (!(c.anchor && c.anchor.from) && o.rowFrac && o.colFrac) c.__tl = { col: o.colFrac(c.x || 0), row: o.rowFrac(c.y || 0) };
    return c;
  });
  const shapes = Array.isArray(sheet.shapes) ? JSON.parse(JSON.stringify(sheet.shapes)) : [];
  const view = isObj(sheet.view) ? { ...sheet.view } : sheet.view;
  const pageSetup = isObj(sheet.pageSetup) ? { ...sheet.pageSetup } : sheet.pageSetup;

  const ordered = active.filter(z => z.direction === 'rows').sort((a, b) => b.start - a.start)
    .concat(active.filter(z => z.direction === 'cols').sort((a, b) => b.start - a.start));

  for (const z of ordered) {
    const rowsAxis = z.direction === 'rows';
    const S = z.start, E = z.end, H = E - S + 1, N = z.items.length, extra = (N - 1) * H;
    const mapIdx = idxMapper(S, E, H, extra);
    // Límites de Excel
    let maxUsed = 0;
    const see = v => { if (v > maxUsed) maxUsed = v; };
    if (rowsAxis) { cells.forEach(c => see(c.row)); rowHeights.forEach((_, k) => see(k)); hiddenRows.forEach(see); }
    else { cells.forEach(c => see(c.col)); colWidths.forEach((_, k) => see(k)); hiddenCols.forEach(see); }
    if (maxUsed + extra > (rowsAxis ? MAX_ROWS : MAX_COLS))
      throw new RangeError(`La zona repetible "${z.arrayVariable}" generaría más ${rowsAxis ? 'filas' : 'columnas'} de las que admite Excel (${N} elementos).`);
    const refFn = k => r => rowsAxis
      ? { col: r.col, row: r.kind === 'cols' ? r.row : mapIdx(r.row, r.rowAbs, r.role, k) }
      : { row: r.row, col: r.kind === 'rows' ? r.col : mapIdx(r.col, r.colAbs, r.role, k) };
    const xform = (cd, k) => {
      if (!cd.formula || typeof cd.formula !== 'string') return cd;
      const body = cd.formula.charAt(0) === '=' ? cd.formula.slice(1) : cd.formula;
      const nf = mapFormulaRefs(body, refFn(k));
      return nf === body ? cd : { ...cd, formula: '=' + nf };
    };
    const outside = refFn(-1);
    // Celdas
    const next = [];
    for (const c of cells) {
      const v = rowsAxis ? c.row : c.col;
      if (v < S || v > E) {
        const nc = { ...c, cd: xform(c.cd, -1) };
        if (v > E) { if (rowsAxis) nc.row += extra; else nc.col += extra; }
        next.push(nc);
      } else {
        for (let k = 0; k < N; k++) {
          const nc = { ...c, cd: xform(c.cd, k) };
          if (rowsAxis) { nc.row += k * H; nc.rowItem = z.items[k]; } else { nc.col += k * H; nc.colItem = z.items[k]; }
          next.push(nc);
        }
      }
    }
    cells = next;
    // Combinaciones: dentro de la zona se duplican; las que la cruzan crecen
    const nm = [];
    for (const m of merges) {
      const a = rowsAxis ? m.sr : m.sc, b = rowsAxis ? m.er : m.ec;
      const put = (na, nb) => nm.push(rowsAxis ? { ...m, sr: na, er: nb } : { ...m, sc: na, ec: nb });
      if (b < S) put(a, b);
      else if (a > E) put(a + extra, b + extra);
      else if (a >= S && b <= E) for (let k = 0; k < N; k++) put(a + k * H, b + k * H);
      else put(a > E ? a + extra : a, b >= E ? b + extra : b);
    }
    merges = nm;
    if (rowsAxis) {
      rowHeights = dupMap(rowHeights, S, E, H, N, extra);
      autoRowHeights = dupMap(autoRowHeights, S, E, H, N, extra);
      hiddenRows = dupList(hiddenRows, S, E, H, N, extra);
      rowOutline = dupMap(rowOutline, S, E, H, N, extra);
    } else {
      colWidths = dupMap(colWidths, S, E, H, N, extra);
      colStyles = dupMap(colStyles, S, E, H, N, extra);
      hiddenCols = dupList(hiddenCols, S, E, H, N, extra);
      colOutline = dupMap(colOutline, S, E, H, N, extra);
    }
    // Imágenes flotantes y formas: lo que está después de la zona se desplaza
    const key = rowsAxis ? 'row' : 'col';
    const mv = pt => { if (pt && typeof pt[key] === 'number' && pt[key] + 1 > E) pt[key] += extra; };
    for (const fi of images) {
      if (fi.anchor && fi.anchor.from) { mv(fi.anchor.from); if (fi.anchor.to) mv(fi.anchor.to); }
      else if (fi.__tl && Math.floor(fi.__tl[key]) + 1 > E) fi.__tl[key] += extra;
    }
    for (const sh of shapes) {
      const a = sh && sh.anchor;
      if (a && a.from) { mv(a.from); if (a.to) mv(a.to); }
    }
    // Paneles inmovilizados y área de impresión
    if (isObj(view)) {
      const sk = rowsAxis ? 'ySplit' : 'xSplit';
      if (typeof view[sk] === 'number' && view[sk] > E) view[sk] += extra;
    }
    if (isObj(pageSetup) && typeof pageSetup.printArea === 'string' && pageSetup.printArea)
      pageSetup.printArea = mapFormulaRefs(pageSetup.printArea, outside, { ignoreSheet: true });
  }

  // Hoja expandida
  const out = { ...sheet };
  delete out.repeatZones; delete out.rowTemplate;
  out.cells = {};
  const meta = new Map();
  for (const c of cells) {
    const a = idxToCol(c.col) + c.row;
    out.cells[a] = c.cd;
    if (c.rowItem !== undefined || c.colItem !== undefined) meta.set(a, c);
  }
  const toObj = m => { const r = {}; for (const [k, v] of [...m].sort((x, y) => x[0] - y[0])) r[k] = v; return r; };
  const toLetters = m => { const r = {}; for (const [k, v] of [...m].sort((x, y) => x[0] - y[0])) r[idxToCol(k)] = v; return r; };
  out.rowHeights = toObj(rowHeights);
  if (sheet.autoRowHeights) out.autoRowHeights = toObj(autoRowHeights);
  out.columnWidths = toLetters(colWidths);
  if (sheet.colStyles) out.colStyles = toLetters(colStyles);
  out.merges = merges.map(m => `${idxToCol(m.sc)}${m.sr}:${idxToCol(m.ec)}${m.er}`).concat(rawMerges);
  out.hiddenRows = [...new Set(hiddenRows)];
  out.hiddenCols = [...new Set(hiddenCols)];
  if (sheet.rowOutline) out.rowOutline = toObj(rowOutline);
  if (sheet.colOutline) out.colOutline = toObj(colOutline);
  out.floatingImages = images;
  out.shapes = shapes;
  if (view !== undefined) out.view = view;
  if (pageSetup !== undefined) out.pageSetup = pageSetup;

  // Datos por celda: elemento de fila > elemento de columna > datos de la hoja
  const cache = new Map();
  const asObj = v => (isObj(v) ? v : {});
  const dataAt = a => {
    const c = meta.get(a);
    if (!c) return d;
    let byRow = cache.get(c.rowItem);
    if (!byRow) { byRow = new Map(); cache.set(c.rowItem, byRow); }
    let merged = byRow.get(c.colItem);
    if (!merged) {
      merged = Object.assign({}, d, c.colItem !== undefined ? asObj(c.colItem) : {}, c.rowItem !== undefined ? asObj(c.rowItem) : {});
      byRow.set(c.colItem, merged);
    }
    return merged;
  };
  return { sheet: out, dataAt };
}

module.exports = {
  normalizeZones, rawZones, zoneOfCell, zoneLabel, mapFormulaRefs, adjustFormulaCols, expandSheet,
  MAX_ROWS, MAX_COLS,
};
