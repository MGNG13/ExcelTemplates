'use strict';
/**
 * Library/automation.js — helpers puros para automatizar la generación de Excel.
 *
 * Sin dependencias ni E/S: todo se calcula a partir del JSON de la plantilla.
 * La semántica de datos replica exactamente renderSheet/setCellValue de main.js:
 *   - El body se indexa por NOMBRE DE HOJA.
 *   - Hoja repetible → array de objetos (una hoja por item); un objeto se renderiza como hoja única.
 *   - Zonas repetibles (sheet.repeatZones, filas o columnas; el antiguo rowTemplate cuenta como una
 *     zona de filas) → sheetData[arrayVariable] es un array de items; las celdas de la zona toman sus
 *     valores de cada item (con respaldo en los datos de la hoja). Ver Library/repeat-zones.js.
 *   - Celdas de texto y fórmulas sin `variable` pueden contener {{nombre}} (substituteVars):
 *     se exponen como variables implícitas (tipo text, no obligatorias).
 */

const { normalizeZones, zoneOfCell, zoneLabel } = require('./repeat-zones');

const PLACEHOLDER_RE = /\{\{(\w+)\}\}/g;           // idéntico a substituteVars en main.js
const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype']);
const RESERVED_QUERY = new Set(['format', 'filename', 'strict', 'download']);
const VAR_TYPES = new Set(['text', 'number', 'boolean', 'date', 'image']);
// ExcelJS (embedImage) admite png, jpeg y gif; guessExt lee el subtipo del data URL.
const IMAGE_RE = /^data:image\/(png|jpe?g|gif);base64,[A-Za-z0-9+/=\s]+$/i;
const WIN_RESERVED = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])$/i;

const isPlainObject = v => v !== null && typeof v === 'object' && !Array.isArray(v);
const isSafeKey = k => typeof k === 'string' && !FORBIDDEN_KEYS.has(k);
const own = (o, k) => isPlainObject(o) && isSafeKey(k) && Object.prototype.hasOwnProperty.call(o, k);
const isEmpty = v => v === undefined || v === null || (typeof v === 'string' && v.trim() === '');

/** Fila (1-based) de una dirección tipo 'B5'; null si no es válida. */
function rowOf(addr) {
  const m = /^([A-Z]+)(\d+)$/i.exec(String(addr));
  return m ? parseInt(m[2], 10) : null;
}

/** Zonas repetibles válidas de la hoja (+ avisos de zonas inválidas/superpuestas/cruzadas). */
function sheetZones(sheet) {
  try { return normalizeZones(sheet); } catch (_) { return { zones: [], issues: [] }; }
}

/** rowTemplate de compatibilidad: la primera zona de filas, en el formato antiguo. */
function legacyRowTemplate(zones) {
  const z = zones.find(x => x.direction === 'rows');
  return z ? { startRow: z.start, endRow: z.end, arrayVariable: z.arrayVariable } : null;
}

/** Zonas agrupadas por nombre de arreglo (dos zonas pueden compartir el mismo arreglo). */
function zoneGroups(zones) {
  const m = new Map();
  for (const z of zones) {
    if (!m.has(z.arrayVariable)) m.set(z.arrayVariable, { arrayVariable: z.arrayVariable, direction: z.direction, zones: [] });
    m.get(z.arrayVariable).zones.push(z);
  }
  return [...m.values()];
}

const zoneKind = dir => (dir === 'cols' ? 'zona de columnas' : 'zona de filas');

function sheetsOf(template) {
  return isPlainObject(template) && Array.isArray(template.sheets) ? template.sheets.filter(isPlainObject) : [];
}

// ── Schema ───────────────────────────────────────────────────────────────────

/** Variables de una hoja (explícitas + implícitas), deduplicadas por nombre y zona. */
function sheetVariables(sheet, zones) {
  const vars = [];
  const seen = new Map();  // key name|zone → índice en vars
  const cells = isPlainObject(sheet.cells) ? sheet.cells : {};
  // Orden estable por fila/columna para que el esquema sea predecible
  const addrs = Object.keys(cells).filter(a => rowOf(a) !== null).sort((a, b) => {
    const ra = rowOf(a), rb = rowOf(b);
    if (ra !== rb) return ra - rb;
    const ca = a.replace(/\d+$/, ''), cb = b.replace(/\d+$/, '');
    return ca.length - cb.length || (ca < cb ? -1 : ca > cb ? 1 : 0);
  });
  const add = (v) => {
    if (!isSafeKey(v.name) || !v.name) return;
    const key = v.name + '|' + (v.zone ? v.zone.arrayVariable : '');
    if (seen.has(key)) {
      const prev = vars[seen.get(key)];
      if (prev.implicit && !v.implicit) vars[seen.get(key)] = v;  // la definición explícita gana
      return;
    }
    seen.set(key, vars.length);
    vars.push(v);
  };
  for (const addr of addrs) {
    const cd = cells[addr];
    if (!isPlainObject(cd) || cd.image) continue;          // las imágenes fijas no usan datos
    const z = zoneOfCell(zones, addr);
    const zone = z ? { arrayVariable: z.arrayVariable, direction: z.direction, id: z.id } : null;
    const inRowTemplate = !!zone;   // compatibilidad: true = va dentro de cada item de una zona
    // setCellValue: la fórmula tiene prioridad sobre la variable
    if (typeof cd.formula === 'string' && cd.formula) {
      for (const m of cd.formula.matchAll(PLACEHOLDER_RE)) add(implicitVar(m[1], addr, zone));
      continue;
    }
    if (isPlainObject(cd.variable) && typeof cd.variable.name === 'string' && cd.variable.name.trim()) {
      const v = cd.variable;
      add({
        name: v.name.trim(),
        type: VAR_TYPES.has(v.type) ? v.type : 'text',
        required: !!v.required,
        hint: typeof v.hint === 'string' ? v.hint : '',
        defaultValue: v.defaultValue === undefined ? '' : v.defaultValue,
        addr, inRowTemplate, zone, implicit: false,
      });
      continue;
    }
    if (typeof cd.value === 'string') {
      for (const m of cd.value.matchAll(PLACEHOLDER_RE)) add(implicitVar(m[1], addr, zone));
    }
  }
  return vars;
}

function implicitVar(name, addr, zone) {
  return { name, type: 'text', required: false, hint: '', defaultValue: '', addr, inRowTemplate: !!zone, zone, implicit: true };
}

/**
 * Esquema de automatización de la plantilla.
 * @returns {{id,name,revision,sheets:Array<{name,isRepeatable,repeatVariable,rowTemplate,repeatZones,zoneIssues,variables}>}}
 *   repeatZones: [{ id, direction, start, end, arrayVariable, label }] · rowTemplate: primera zona de filas (compatibilidad)
 *   variables[].zone: { arrayVariable, direction, id } | null
 */
function buildSchema(template) {
  const t = isPlainObject(template) ? template : {};
  return {
    id: typeof t.id === 'string' ? t.id : null,
    name: typeof t.name === 'string' ? t.name : '',
    revision: typeof t.revision === 'number' ? t.revision : 0,   // plantillas antiguas sin revision
    sheets: sheetsOf(t).map(s => {
      const { zones, issues } = sheetZones(s);
      return {
        name: typeof s.name === 'string' ? s.name : '',
        isRepeatable: !!s.isRepeatable,
        repeatVariable: s.repeatVariable || null,
        rowTemplate: legacyRowTemplate(zones),
        repeatZones: zones.map(z => ({ id: z.id, direction: z.direction, start: z.start, end: z.end, arrayVariable: z.arrayVariable, label: zoneLabel(z) })),
        zoneIssues: issues.map(i => ({ code: i.code, zone: i.zone, message: i.message })),
        variables: sheetVariables(s, zones),
      };
    }),
  };
}

// ── Ejemplo ──────────────────────────────────────────────────────────────────

function exampleValue(v, i) {
  if (!isEmpty(v.defaultValue) && i === 0) return v.defaultValue;
  switch (v.type) {
    case 'number': return (i + 1) * 10;
    case 'boolean': return i % 2 === 0;
    case 'date': return i === 0 ? '2026-01-15' : '2026-02-15';
    case 'image': return 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
    default: return i === 0 ? `Ejemplo ${v.name}` : `Ejemplo ${v.name} ${i + 1}`;
  }
}

function exampleSheetData(sheet, idx) {
  const out = {};
  for (const v of sheet.variables) if (!v.zone && !own(out, v.name)) out[v.name] = exampleValue(v, idx);
  // Cada zona → arreglo de 2 elementos con sus variables
  for (const g of zoneGroups(sheet.repeatZones || [])) {
    const zVars = sheet.variables.filter(v => v.zone && v.zone.arrayVariable === g.arrayVariable);
    out[g.arrayVariable] = [0, 1].map(i => {
      const item = {};
      for (const v of zVars) if (!own(item, v.name)) item[v.name] = exampleValue(v, i);
      return item;
    });
  }
  return out;
}

/** Body de ejemplo listo para POST /api/render/:id. */
function buildExample(template) {
  const body = {};
  for (const s of buildSchema(template).sheets) {
    if (!isSafeKey(s.name) || !s.name) continue;
    body[s.name] = s.isRepeatable ? [exampleSheetData(s, 0), exampleSheetData(s, 1)] : exampleSheetData(s, 0);
  }
  return body;
}

// ── Validación ───────────────────────────────────────────────────────────────

const TYPE_LABEL = { text: 'texto', number: 'número', boolean: 'booleano', date: 'fecha', image: 'imagen' };

/** Devuelve mensaje de error si el valor no es del tipo, o null si es válido. Asume valor no vacío. */
function typeError(type, value) {
  switch (type) {
    case 'number':
      if (typeof value === 'number' && Number.isFinite(value)) return null;
      if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value.trim()))) return null;
      return 'debe ser un número (p. ej. 12.5)';
    case 'boolean':
      if (value === true || value === false || value === 0 || value === 1 || value === 'true' || value === 'false') return null;
      return 'debe ser booleano: true, false, "true", "false", 0 o 1';
    case 'date':
      if ((typeof value === 'string' || typeof value === 'number') && !Number.isNaN(new Date(value).getTime())) return null;
      return 'debe ser una fecha válida (p. ej. "2026-01-15")';
    case 'image':
      if (typeof value === 'string' && IMAGE_RE.test(value)) return null;
      return 'debe ser una imagen como data URL: data:image/png;base64,… (png, jpeg o gif)';
    default:
      if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return null;
      return 'debe ser texto, número o booleano';
  }
}

function describe(v) {
  if (v === null) return 'null';
  if (Array.isArray(v)) return 'un arreglo';
  return typeof v === 'object' ? 'un objeto' : typeof v;
}

/**
 * Valida un body contra la plantilla. Nunca lanza.
 * @returns {{valid:boolean, errors:Array, warnings:Array}}
 */
function validateBody(template, body) {
  const errors = [], warnings = [];
  const err = (sheet, path, variable, code, message) => errors.push({ sheet, path, variable, code, message });
  const warn = (sheet, path, message) => warnings.push({ sheet, path, code: 'unknown', message });

  let schema;
  try { schema = buildSchema(template); } catch (_) { schema = { sheets: [] }; }

  if (!isPlainObject(body)) {
    err(null, '', null, 'shape', `El cuerpo debe ser un objeto JSON con una clave por hoja; se recibió ${describe(body)}.`);
    return { valid: false, errors, warnings };
  }

  // Valida un objeto de datos (hoja o item de zona) contra un conjunto de variables
  const checkVars = (sheetName, basePath, data, vars, extraKnown, zoneHint) => {
    for (const v of vars) {
      const path = `${basePath}.${v.name}`;
      const value = own(data, v.name) ? data[v.name] : undefined;
      if (isEmpty(value)) {
        if (v.required) err(sheetName, path, v.name, 'missing', `Falta la variable obligatoria "${v.name}" (${path}).`);
        continue;
      }
      const te = typeError(v.type, value);
      if (te) err(sheetName, path, v.name, 'type', `"${path}" ${te}; se recibió ${describe(value)}.`);
    }
    const known = new Set(vars.map(v => v.name).concat(extraKnown || []));
    for (const k of Object.keys(data)) {
      if (known.has(k)) continue;
      // Pista útil: la variable existe, pero pertenece a una zona repetible (va dentro de cada item)
      const zh = zoneHint && zoneHint.get(k);
      if (zh) {
        warn(sheetName, `${basePath}.${k}`, `"${k}" es una variable de la ${zoneKind(zh.direction)}: va dentro de cada elemento de "${basePath}.${zh.arrayVariable}"; aquí se ignorará.`);
      } else {
        warn(sheetName, `${basePath}.${k}`, `"${k}" no es una variable de ${basePath}; se ignorará.`);
      }
    }
  };

  const checkSheetObject = (s, basePath, data) => {
    const groups = zoneGroups(s.repeatZones || []);
    const topVars = s.variables.filter(v => !v.zone);
    const topNames = new Set(topVars.map(v => v.name));
    // Pista: variable de zona enviada a nivel de hoja (va dentro de cada elemento del arreglo)
    const zoneHint = new Map();
    for (const v of s.variables) if (v.zone && !topNames.has(v.name) && !zoneHint.has(v.name)) zoneHint.set(v.name, v.zone);
    checkVars(s.name, basePath, data, topVars, groups.map(g => g.arrayVariable), zoneHint);
    for (const g of groups) {
      const zVars = s.variables.filter(v => v.zone && v.zone.arrayVariable === g.arrayVariable);
      const kind = zoneKind(g.direction);
      const arrPath = `${basePath}.${g.arrayVariable}`;
      const items = own(data, g.arrayVariable) ? data[g.arrayVariable] : undefined;
      if (items === undefined || items === null) {
        if (zVars.some(v => v.required))
          err(s.name, arrPath, g.arrayVariable, 'missing', `Falta "${arrPath}": la ${kind} tiene variables obligatorias.`);
        continue;
      }
      if (!Array.isArray(items)) {
        err(s.name, arrPath, g.arrayVariable, 'shape', `"${arrPath}" debe ser un arreglo de objetos (${kind}); se recibió ${describe(items)}.`);
        continue;
      }
      if (!items.length && zVars.some(v => v.required))
        err(s.name, arrPath, g.arrayVariable, 'missing', `"${arrPath}" está vacío: la ${kind} tiene variables obligatorias.`);
      items.forEach((item, i) => {
        const ip = `${arrPath}[${i}]`;
        if (!isPlainObject(item)) {
          err(s.name, ip, g.arrayVariable, 'shape', `"${ip}" debe ser un objeto; se recibió ${describe(item)}.`);
          return;
        }
        checkVars(s.name, ip, item, zVars);
      });
    }
  };

  const knownSheets = new Set();
  for (const s of schema.sheets) {
    if (!s.name || !isSafeKey(s.name)) continue;
    knownSheets.add(s.name);
    // Zonas inválidas, superpuestas o cruzadas: aviso (la generación las ignora o las repite en ambas direcciones)
    for (const zi of s.zoneIssues || [])
      warnings.push({ sheet: s.name, path: zi.zone ? `${s.name}.${zi.zone}` : s.name, code: 'zone', message: zi.message });
    const data = own(body, s.name) ? body[s.name] : undefined;
    if (data === undefined || data === null) {
      // Sin datos: la hoja se genera vacía; solo es error si hay obligatorias
      checkSheetObject(s, s.name, {});
      continue;
    }
    if (s.isRepeatable) {
      if (Array.isArray(data)) {
        if (!data.length) warn(s.name, s.name, `"${s.name}" es un arreglo vacío: la hoja repetible no se incluirá en el archivo.`);
        data.forEach((item, i) => {
          const p = `${s.name}[${i}]`;
          if (!isPlainObject(item)) err(s.name, p, null, 'shape', `"${p}" debe ser un objeto; se recibió ${describe(item)}.`);
          else checkSheetObject(s, p, item);
        });
      } else if (isPlainObject(data)) {
        warnings.push({ sheet: s.name, path: s.name, code: 'unknown', message: `"${s.name}" es una hoja repetible: se esperaba un arreglo; el objeto se usará como un solo elemento.` });
        checkSheetObject(s, s.name, data);
      } else {
        err(s.name, s.name, null, 'shape', `"${s.name}" debe ser un arreglo de objetos (hoja repetible); se recibió ${describe(data)}.`);
      }
    } else if (isPlainObject(data)) {
      checkSheetObject(s, s.name, data);
    } else {
      err(s.name, s.name, null, 'shape', `"${s.name}" debe ser un objeto; se recibió ${describe(data)}.`);
    }
  }
  for (const k of Object.keys(body)) {
    if (!knownSheets.has(k)) warn(null, k, `La hoja "${k}" no existe en la plantilla; se ignorará.`);
  }
  return { valid: errors.length === 0, errors, warnings };
}

// ── Query string → body ──────────────────────────────────────────────────────

function coerce(type, raw) {
  const s = String(raw).trim();
  if (type === 'number') { const n = Number(s); return s !== '' && Number.isFinite(n) ? n : s; }
  if (type === 'boolean') {
    const l = s.toLowerCase();
    if (['true', '1', 'si', 'sí', 'yes', 'on'].includes(l)) return true;
    if (['false', '0', 'no', 'off', ''].includes(l)) return false;
    return s;
  }
  return String(raw);
}

/**
 * ?cliente=Ana&total=5 → { Factura: { cliente: 'Ana', total: 5 } }.
 * Solo hojas NO repetibles y variables fuera de la zona de filas.
 */
function bodyFromQuery(template, query) {
  const body = {};
  if (!isPlainObject(query)) return body;
  let schema;
  try { schema = buildSchema(template); } catch (_) { return body; }
  for (const key of Object.keys(query)) {
    if (!isSafeKey(key) || RESERVED_QUERY.has(key.toLowerCase())) continue;
    let raw = query[key];
    if (Array.isArray(raw)) raw = raw[0];
    if (raw === undefined || raw === null || typeof raw === 'object') continue;
    for (const s of schema.sheets) {
      if (s.isRepeatable || !s.name || !isSafeKey(s.name)) continue;
      const v = s.variables.find(x => x.name === key && !x.inRowTemplate);
      if (!v) continue;
      if (!own(body, s.name)) body[s.name] = {};
      body[s.name][key] = coerce(v.type, raw);
    }
  }
  return body;
}

// ── Nombres de archivo ───────────────────────────────────────────────────────

/** Recorta a `max` unidades UTF-16 sin partir pares sustitutos (emoji): evita URIError en encodeURIComponent. */
function truncateUtf16(s, max) {
  if (s.length <= max) return s;
  let out = '';
  for (const ch of s) {             // itera por punto de código
    if (out.length + ch.length > max) break;
    out += ch;
  }
  return out;
}

/** Nombre seguro para Windows/macOS/Linux, siempre terminado en .xlsx y ≤ 120 caracteres. */
function safeFileName(name, fallback = 'resultado') {
  const clean = (s) => {
    let out = String(s == null ? '' : s)
      .normalize('NFC')
      .replace(/[\u0000-\u001f\u007f-\u009f]/g, '')
      // Caracteres invisibles / de control bidireccional (evitan nombres engañosos tipo "fdp.exe")
      .replace(/[​-‏‪-‮⁦-⁩﻿]/g, '')
      // Sustitutos sueltos (texto UTF-16 mal formado)
      .replace(/[\ud800-\udbff](?![\udc00-\udfff])|(^|[^\ud800-\udbff])[\udc00-\udfff]/g, '$1')
      .replace(/[/\\:*?"<>|]/g, '_')
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/\.xlsx$/i, '')
      .replace(/^[.\s]+|[.\s]+$/g, '');
    if (WIN_RESERVED.test(out.split('.')[0])) out = '_' + out;
    const max = 120 - '.xlsx'.length;
    if (out.length > max) out = truncateUtf16(out, max).replace(/[.\s]+$/, '');
    return out;
  };
  const base = clean(name) || clean(fallback) || 'resultado';
  return base + '.xlsx';
}

/** Primer valor primitivo por nombre de clave (búsqueda en anchura, sin seguir prototipos). */
function flatten(body) {
  const flat = new Map();
  const queue = [body];
  const seen = new Set();
  let steps = 0;
  while (queue.length && steps++ < 10000) {
    const cur = queue.shift();
    if (cur === null || typeof cur !== 'object' || seen.has(cur)) continue;
    seen.add(cur);
    if (Array.isArray(cur)) { cur.forEach(x => queue.push(x)); continue; }
    for (const k of Object.keys(cur)) {
      if (!isSafeKey(k)) continue;
      const v = cur[k];
      if (v !== null && typeof v === 'object') queue.push(v);
      else if (!flat.has(k) && v !== undefined && v !== null && v !== '') flat.set(k, v);
    }
  }
  return flat;
}

/** 'Factura {{cliente}}' + body → 'Factura Ana' (variables ausentes → ''). */
function fillPattern(pattern, body) {
  if (typeof pattern !== 'string') return '';
  const flat = flatten(body);
  return pattern.replace(PLACEHOLDER_RE, (_, k) => (flat.has(k) ? String(flat.get(k)) : '')).replace(/\s+/g, ' ').trim();
}

// ── OpenAPI ──────────────────────────────────────────────────────────────────

function buildOpenApi({ baseUrl } = {}) {
  const ref = n => ({ $ref: `#/components/schemas/${n}` });
  const json = (schema, example) => ({ 'application/json': example === undefined ? { schema } : { schema, example } });
  const err = (desc, example) => ({ description: desc, content: json(ref('Error'), example) });
  const idParam = { name: 'id', in: 'path', required: true, description: 'ID de la plantilla', schema: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,64}$' } };
  const r404 = err('Plantilla no encontrada', { error: 'Plantilla no encontrada' });
  const r400 = err('Solicitud inválida (ID, formato o cuerpo)', { error: 'ID inválido' });
  // Encabezados opcionales para atribuir cambios y uso a una persona (los envía la interfaz web)
  const who = ['XClientId', 'XUserId', 'XClientName'].map(n => ({ $ref: `#/components/parameters/${n}` }));
  const r429 = err('Demasiadas solicitudes (límite de tasa)', { error: 'Demasiadas solicitudes, intenta de nuevo en un momento', code: 'RATE_LIMIT' });
  const exampleBody = { Factura: { cliente: 'Ana López', total: 1160, items: [{ descripcion: 'Servicio', cantidad: 2, precio: 500 }] } };
  const xlsx = { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': { schema: { type: 'string', format: 'binary' } } };
  const renderParams = [
    idParam,
    { name: 'format', in: 'query', description: '`xlsx` (binario, predeterminado) o `base64` (JSON)', schema: { type: 'string', enum: ['xlsx', 'base64'], default: 'xlsx' } },
    { name: 'filename', in: 'query', description: 'Nombre del archivo descargado (se sanea y se le agrega .xlsx); si se omite, el nombre de la plantilla', schema: { type: 'string' } },
    { name: 'strict', in: 'query', description: 'Con `1` valida el cuerpo y responde 422 si hay errores', schema: { type: 'string', enum: ['0', '1'] } },
    ...who,
  ];
  const renderOk = {
    description: 'Archivo Excel generado (o JSON en modo base64)',
    headers: {
      'Content-Disposition': { description: "attachment; filename*=UTF-8''…", schema: { type: 'string' } },
      'X-Render-Ms': { description: 'Milisegundos que tardó la generación', schema: { type: 'integer' } },
    },
    content: { ...xlsx, ...json(ref('Base64File'), { fileName: 'Factura Ana.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', size: 6789, base64: 'UEsDBBQABgAIAAAAIQ…' }) },
  };
  const r422 = { description: 'Cuerpo inválido en modo estricto', content: json(ref('ValidationError'), {
    error: 'Los datos no cumplen el esquema de la plantilla', code: 'VALIDATION',
    errors: [{ sheet: 'Factura', path: 'Factura.total', variable: 'total', code: 'missing', message: 'Falta la variable obligatoria "total" (Factura.total).' }],
    warnings: [],
  }) };
  const tplSaved = { success: true, id: 'mi-plantilla', revision: 4, updatedAt: '2026-11-19T18:00:00.000Z' };

  return {
    openapi: '3.1.0',
    info: {
      title: 'Excel Template Builder API',
      version: '12.0.0',
      description: 'API para gestionar plantillas de Excel y generar archivos rellenando sus variables {{…}}. ' +
        'El cuerpo de generación se indexa por **nombre de hoja**; las hojas repetibles reciben un arreglo y cada zona repetible (filas o columnas, ' +
        '`repeatZones` de la hoja) un arreglo de items bajo su `arrayVariable`: la zona se copia una vez por elemento y las fórmulas que la abarcan (p. ej. SUMA) crecen para cubrir todas las copias. ' +
        'Envía los encabezados opcionales `X-Client-Id`, `X-User-Id` y `X-Client-Name` para identificar quién hace los cambios.',
    },
    servers: [{ url: baseUrl || '/', description: 'Servidor actual' }],
    tags: [
      { name: 'Plantillas', description: 'Crear, leer, guardar, duplicar y eliminar plantillas' },
      { name: 'Generación', description: 'Generar archivos Excel (uno o en lote)' },
      { name: 'Automatización', description: 'Esquema, ejemplos, validación y esta especificación' },
      { name: 'Tiempo real', description: 'Presencia y eventos en vivo por SSE' },
      { name: 'Uso', description: 'Métricas de uso de cada plantilla' },
    ],
    paths: {
      '/api/health': { get: { tags: ['Automatización'], summary: 'Estado del servidor', operationId: 'health',
        responses: { 200: { description: 'OK', content: json({ type: 'object', properties: { ok: { type: 'boolean' }, version: { type: 'string' }, uptime: { type: 'number' } } }, { ok: true, version: '12.0.0', uptime: 1234.5 }) } } } },
      '/api/templates': {
        get: { tags: ['Plantillas'], summary: 'Listar plantillas', operationId: 'listTemplates',
          responses: { 200: { description: 'Plantillas ordenadas por última modificación', content: json({ type: 'array', items: ref('TemplateSummary') }, [{
            id: 'factura', name: 'Factura', sheetCount: 2, updatedAt: '2026-11-19T18:00:00.000Z', createdAt: '2026-10-01T15:00:00.000Z', revision: 4,
            updatedBy: { clientId: 'k3j2h1g0f9', userId: 'u-ana', name: 'Ana' }, thumbnail: null, sheetNames: ['Factura', 'Pedido'], variables: [] }]) } } },
        post: { tags: ['Plantillas'], summary: 'Crear o guardar una plantilla', operationId: 'saveTemplate',
          description: 'Si el cuerpo incluye `revision` (número) se usa como revisión base: si otra persona guardó antes, responde 409. Usa `?force=1` para sobrescribir.',
          parameters: [{ name: 'force', in: 'query', description: '`1` sobrescribe aunque haya conflicto', schema: { type: 'string', enum: ['0', '1'] } }, ...who],
          requestBody: { required: true, content: json(ref('Template')) },
          responses: {
            200: { description: 'Guardada', content: json({ type: 'object', properties: { success: { type: 'boolean' }, id: { type: 'string' }, revision: { type: 'integer' }, updatedAt: { type: 'string', format: 'date-time' } } }, tplSaved) },
            409: { description: 'Conflicto: la plantilla cambió en el servidor', content: json(ref('ConflictError'), { error: 'La plantilla fue modificada por otra persona', code: 'CONFLICT', current: { revision: 5, updatedAt: '2026-11-19T18:01:00.000Z', updatedBy: { clientId: 'abc12345', userId: 'u1', name: 'Luis' } } }) },
            400: err('Datos de plantilla inválidos', { error: 'Datos de plantilla inválidos' }),
            429: r429,
          } },
      },
      '/api/templates/{id}': {
        get: { tags: ['Plantillas'], summary: 'Obtener una plantilla', operationId: 'getTemplate', parameters: [idParam],
          responses: { 200: { description: 'Plantilla completa (incluye `revision`)', content: json(ref('Template')) }, 404: r404 } },
        delete: { tags: ['Plantillas'], summary: 'Eliminar una plantilla', operationId: 'deleteTemplate', parameters: [idParam, ...who],
          responses: { 200: { description: 'Eliminada', content: json(ref('Success'), { success: true }) }, 404: r404 } },
      },
      '/api/templates/{id}/thumbnail': { put: { tags: ['Plantillas'], summary: 'Actualizar la miniatura', operationId: 'putThumbnail', parameters: [idParam],
        requestBody: { required: true, content: json({ type: 'object', properties: { thumbnail: { type: 'string', description: 'data:image/(png|jpeg|webp);base64 ≤ 2 MB' } } }) },
        responses: { 200: { description: 'OK', content: json(ref('Success'), { success: true }) }, 400: err('Imagen inválida'), 404: r404 } } },
      '/api/templates/{id}/duplicate': { post: { tags: ['Plantillas'], summary: 'Duplicar una plantilla', operationId: 'duplicateTemplate', parameters: [idParam, ...who],
        responses: { 200: { description: 'Copia creada', content: json({ type: 'object', properties: { success: { type: 'boolean' }, id: { type: 'string' }, name: { type: 'string' } } }, { success: true, id: '3f2a…', name: 'Factura (copia)' }) }, 404: r404 } } },
      '/api/templates/{id}/schema': { get: { tags: ['Automatización'], summary: 'Esquema de variables', operationId: 'getSchema', parameters: [idParam],
        responses: { 200: { description: 'Hojas y variables (formato compatible con versiones anteriores + `revision`)', content: json({ type: 'object', properties: { id: { type: 'string' }, name: { type: 'string' }, revision: { type: 'integer' }, schema: { type: 'array', items: { type: 'object' } } } }) }, 404: r404 } } },
      '/api/templates/{id}/example': { get: { tags: ['Automatización'], summary: 'Cuerpo de ejemplo', operationId: 'getExample', parameters: [idParam],
        responses: { 200: { description: 'Body listo para POST /api/render/{id}', content: json({ type: 'object' }, exampleBody) }, 404: r404 } } },
      '/api/templates/{id}/validate': { post: { tags: ['Automatización'], summary: 'Validar datos sin generar', operationId: 'validateBody', parameters: [idParam],
        requestBody: { required: true, content: json({ type: 'object' }, exampleBody) },
        responses: { 200: { description: 'Resultado (siempre 200; revisa `valid`)', content: json(ref('ValidationResult'), { valid: true, errors: [], warnings: [] }) }, 404: r404 } } },
      '/api/import': { post: { tags: ['Plantillas'], summary: 'Importar un .xlsx como plantilla', operationId: 'importXlsx',
        requestBody: { required: true, content: { 'multipart/form-data': { schema: { type: 'object', properties: { file: { type: 'string', format: 'binary' } }, required: ['file'] } } } },
        responses: { 200: { description: 'Plantilla convertida a JSON (no se guarda hasta enviarla a POST /api/templates)', content: json({ type: 'object' }) }, 400: err('No se recibió archivo', { error: 'No se recibió archivo' }) } } },
      '/api/render/{id}': {
        post: { tags: ['Generación'], summary: 'Generar un Excel', operationId: 'render', parameters: renderParams,
          requestBody: { required: true, content: json({ type: 'object', description: 'Claves = nombres de hoja' }, exampleBody) },
          responses: { 200: renderOk, 400: r400, 404: r404, 422: r422, 429: r429 } },
        get: { tags: ['Generación'], summary: 'Generar un Excel desde un enlace', operationId: 'renderLink',
          description: 'Cada parámetro de consulta que coincida con una variable de una hoja NO repetible se usa como valor (`?cliente=Ana&total=5`). Ideal para enlaces directos.',
          parameters: renderParams.concat([{ name: 'variables', in: 'query', style: 'form', explode: true, schema: { type: 'object', additionalProperties: { type: 'string' } }, description: 'Variables como parámetros sueltos' }]),
          responses: { 200: renderOk, 400: r400, 404: r404, 422: r422, 429: r429 } },
      },
      '/api/render/{id}/batch': { post: { tags: ['Generación'], summary: 'Generar varios Excel en un ZIP', operationId: 'renderBatch',
        description: 'Genera un .xlsx por elemento y los entrega en un ZIP. `fileName` es un patrón con {{variables}} (se toma el primer valor encontrado en cada elemento); los nombres repetidos reciben " (2)", " (3)"…',
        parameters: [idParam,
          { name: 'format', in: 'query', schema: { type: 'string', enum: ['xlsx', 'base64'] }, description: '`base64` devuelve JSON; si se omite, el ZIP binario' },
          { name: 'filename', in: 'query', schema: { type: 'string' }, description: 'Nombre del ZIP (sin extensión)' },
          { name: 'strict', in: 'query', schema: { type: 'string', enum: ['0', '1'] }, description: 'Con `1` valida cada elemento; los errores incluyen `item` (índice base 0)' },
          ...who],
        requestBody: { required: true, content: json({ oneOf: [
          { type: 'object', properties: { items: { type: 'array', minItems: 1, maxItems: 200, items: { type: 'object' } }, fileName: { type: 'string', description: 'Patrón, p. ej. "Factura {{cliente}}"' } }, required: ['items'] },
          { type: 'array', minItems: 1, maxItems: 200, items: { type: 'object' } },
        ] }, { items: [exampleBody, { Factura: { cliente: 'Luis', total: 580 } }], fileName: 'Factura {{cliente}}' }) },
        responses: {
          200: { description: 'ZIP con un .xlsx por item (nombres únicos)', content: { 'application/zip': { schema: { type: 'string', format: 'binary' } },
            ...json({ allOf: [ref('Base64File'), { type: 'object', properties: { count: { type: 'integer' } } }] }, { fileName: 'Factura.zip', mimeType: 'application/zip', size: 12345, base64: 'UEsDB…', count: 2 }) } },
          400: err('Lote vacío, con más de 200 elementos o con elementos que no son objetos', { error: 'Se requiere "items" con entre 1 y 200 elementos' }), 404: r404, 422: r422, 429: r429,
        } } },
      '/api/openapi.json': { get: { tags: ['Automatización'], summary: 'Esta especificación OpenAPI', operationId: 'openapi', responses: { 200: { description: 'Documento OpenAPI 3.1', content: json({ type: 'object' }) } } } },
      '/api/live/stream': { get: { tags: ['Tiempo real'], summary: 'Flujo de eventos en vivo (SSE)', operationId: 'liveStream',
        description: 'Eventos: `hello`, `presence`, `template:saved`, `template:created`, `template:deleted`, `usage`, `ping`.',
        parameters: ['clientId', 'userId', 'name', 'color'].map(n => ({ name: n, in: 'query', required: n === 'clientId', schema: { type: 'string' } })),
        responses: { 200: { description: 'text/event-stream', content: { 'text/event-stream': { schema: { type: 'string' }, example: 'event: presence\ndata: {"clients":[],"serverTime":1763575261000}\n\n' } } },
          400: err('clientId inválido', { error: 'clientId inválido', code: 'BAD_CLIENT_ID' }), 429: err('Demasiadas conexiones desde la misma dirección'), 503: err('Demasiados clientes conectados o tiempo real no disponible') } } },
      '/api/live/presence': {
        get: { tags: ['Tiempo real'], summary: 'Usuarios conectados', operationId: 'getPresence',
          responses: { 200: { description: 'Presencia actual', content: json({ type: 'object', properties: { clients: { type: 'array', items: ref('PresenceEntry') }, serverTime: { type: 'integer', description: 'Epoch en ms' } } }, { clients: [{ clientId: 'k3j2h1g0f9', userId: 'u-ana', name: 'Ana', color: '#217346', mode: 'edit', templateId: 'factura', templateName: 'Factura', sheet: 'Factura', cell: 'B5', range: null, connectedAt: 1763575200000, lastActive: 1763575260000, idle: false }], serverTime: 1763575261000 }) } } },
        post: { tags: ['Tiempo real'], summary: 'Actualizar mi presencia', operationId: 'postPresence',
          requestBody: { required: true, content: json({ type: 'object', properties: { clientId: { type: 'string' } }, required: ['clientId'], additionalProperties: true }) },
          responses: { 204: { description: 'Actualizada' }, 404: err('No hay un flujo abierto para ese clientId', { error: 'Sin conexión en vivo', code: 'NO_STREAM' }) } },
      },
      '/api/live/leave': { post: { tags: ['Tiempo real'], summary: 'Desconectarse', operationId: 'leave',
        requestBody: { content: json({ type: 'object', properties: { clientId: { type: 'string' } } }) }, responses: { 204: { description: 'OK' } } } },
      '/api/usage': { get: { tags: ['Uso'], summary: 'Resumen de uso', operationId: 'usageSummary', responses: { 200: { description: 'Totales y por plantilla', content: json(ref('UsageSummary')) } } } },
      '/api/usage/events': { get: { tags: ['Uso'], summary: 'Eventos recientes', operationId: 'usageEvents',
        parameters: [{ name: 'limit', in: 'query', schema: { type: 'integer', default: 100, minimum: 1, maximum: 1000 } }, { name: 'templateId', in: 'query', schema: { type: 'string' } }],
        responses: { 200: { description: 'Eventos (más reciente primero)', content: json({ type: 'object', properties: { events: { type: 'array', items: ref('UsageEvent') } } }) }, 400: err('templateId inválido') } } },
      '/api/usage/{id}': { get: { tags: ['Uso'], summary: 'Detalle de uso de una plantilla', operationId: 'usageDetail', parameters: [idParam],
        responses: { 200: { description: 'Series por hora/día, fuentes y clientes (con ceros si no hay datos)', content: json(ref('TemplateUsageDetail')) }, 400: err('ID inválido') } } },
    },
    components: {
      parameters: {
        XClientId: { name: 'X-Client-Id', in: 'header', required: false, description: 'Identificador de la pestaña/cliente (la interfaz web lo envía; con él la generación cuenta como origen "ui")', schema: { type: 'string' } },
        XUserId: { name: 'X-User-Id', in: 'header', required: false, description: 'Identificador estable de la persona', schema: { type: 'string' } },
        XClientName: { name: 'X-Client-Name', in: 'header', required: false, description: 'Nombre visible, codificado con encodeURIComponent', schema: { type: 'string' } },
      },
      schemas: {
        Error: { type: 'object', description: 'Error con mensaje en español y código opcional (CONFLICT, VALIDATION, RATE_LIMIT, NO_STREAM…)', properties: { error: { type: 'string' }, code: { type: 'string' } }, required: ['error'] },
        Success: { type: 'object', properties: { success: { type: 'boolean' } } },
        Template: { type: 'object', description: 'Plantilla completa (hojas, celdas, estilos…). Cada hoja puede tener `repeatZones` (ver RepeatZone).', properties: { id: { type: 'string' }, name: { type: 'string' }, revision: { type: 'integer' }, sheets: { type: 'array', items: { type: 'object', properties: { name: { type: 'string' }, repeatZones: { type: 'array', items: ref('RepeatZone') } }, additionalProperties: true } } }, additionalProperties: true },
        RepeatZone: { type: 'object', description: 'Filas o columnas que se repiten por cada elemento de sheetData[arrayVariable]. Las zonas de la misma dirección no pueden superponerse. Con 0 elementos la zona queda una vez, sin datos de item. El antiguo `rowTemplate` {startRow,endRow,arrayVariable} equivale a una zona de filas.',
          properties: { id: { type: 'string' }, direction: { type: 'string', enum: ['rows', 'cols'] }, start: { type: 'integer', minimum: 1, description: 'Primera fila / columna (1 = A)' }, end: { type: 'integer', minimum: 1 }, arrayVariable: { type: 'string' } },
          required: ['direction', 'start', 'end', 'arrayVariable'], example: { id: 'meses', direction: 'cols', start: 2, end: 2, arrayVariable: 'meses' } },
        Client: { type: ['object', 'null'], properties: { clientId: { type: 'string' }, userId: { type: 'string' }, name: { type: 'string' } } },
        TemplateSummary: { type: 'object', properties: {
          id: { type: 'string' }, name: { type: 'string' }, sheetCount: { type: 'integer' }, updatedAt: { type: 'string', format: 'date-time' },
          createdAt: { type: 'string', format: 'date-time' }, revision: { type: 'integer' }, updatedBy: ref('Client'),
          thumbnail: { type: ['string', 'null'] }, sheetNames: { type: 'array', items: { type: 'string' } }, variables: { type: 'array', items: { type: 'object' } } } },
        ValidationIssue: { type: 'object', properties: { sheet: { type: ['string', 'null'] }, path: { type: 'string' }, variable: { type: ['string', 'null'] }, code: { type: 'string', enum: ['missing', 'type', 'shape', 'unknown', 'zone'] }, message: { type: 'string' }, item: { type: 'integer', description: 'Solo en lotes: índice del elemento' } } },
        ValidationResult: { type: 'object', properties: { valid: { type: 'boolean' }, errors: { type: 'array', items: ref('ValidationIssue') }, warnings: { type: 'array', items: ref('ValidationIssue') } }, required: ['valid', 'errors', 'warnings'] },
        ValidationError: { allOf: [ref('Error'), { type: 'object', properties: { errors: { type: 'array', items: ref('ValidationIssue') }, warnings: { type: 'array', items: ref('ValidationIssue') } } }] },
        ConflictError: { allOf: [ref('Error'), { type: 'object', properties: { current: { type: 'object', properties: { revision: { type: 'integer' }, updatedAt: { type: 'string' }, updatedBy: ref('Client') } } } }] },
        Base64File: { type: 'object', properties: { fileName: { type: 'string' }, mimeType: { type: 'string' }, size: { type: 'integer' }, base64: { type: 'string' } } },
        PresenceEntry: { type: 'object', properties: {
          clientId: { type: 'string' }, userId: { type: 'string' }, name: { type: 'string' }, color: { type: 'string' },
          mode: { type: 'string', enum: ['home', 'view', 'edit', 'dashboard'] }, templateId: { type: ['string', 'null'] }, templateName: { type: ['string', 'null'] },
          sheet: { type: ['string', 'null'] }, cell: { type: ['string', 'null'] }, range: { type: ['string', 'null'] },
          connectedAt: { type: 'integer', description: 'Epoch en ms' }, lastActive: { type: 'integer', description: 'Epoch en ms' }, idle: { type: 'boolean' } } },
        UsageEvent: { type: 'object', properties: { id: { type: 'string' }, ts: { type: 'string', format: 'date-time' }, type: { type: 'string', enum: ['render', 'batch', 'view', 'save', 'create', 'import', 'duplicate', 'delete'] }, templateId: { type: 'string' }, templateName: { type: 'string' }, ok: { type: 'boolean' }, ms: { type: 'number' }, bytes: { type: 'number' }, count: { type: 'integer' }, source: { type: 'string', enum: ['ui', 'api', 'link'] } } },
        GlobalTotals: { type: 'object', properties: {
          renders: { type: 'integer' }, errors: { type: 'integer' }, views: { type: 'integer' }, saves: { type: 'integer' }, bytes: { type: 'integer' },
          today: { type: 'object', properties: { renders: { type: 'integer' }, errors: { type: 'integer' }, views: { type: 'integer' } } },
          rendersLastHour: { type: 'integer' }, templatesUsed: { type: 'integer' } } },
        TemplateUsageSummary: { type: 'object', properties: {
          id: { type: 'string' }, name: { type: 'string' }, deleted: { type: 'boolean' },
          totals: { type: 'object', properties: { renders: { type: 'integer' }, errors: { type: 'integer' }, views: { type: 'integer' }, saves: { type: 'integer' }, bytes: { type: 'integer' }, avgMs: { type: 'number' } } },
          today: { type: 'object', properties: { renders: { type: 'integer' }, errors: { type: 'integer' }, views: { type: 'integer' } } },
          last: { type: 'object', properties: { renderAt: { type: ['string', 'null'] }, viewAt: { type: ['string', 'null'] }, saveAt: { type: ['string', 'null'] } } },
          spark24h: { type: 'array', items: { type: 'integer' }, minItems: 24, maxItems: 24, description: 'Generaciones por hora (antigua → actual)' },
          spark7d: { type: 'array', items: { type: 'integer' }, minItems: 7, maxItems: 7, description: 'Generaciones por día' } } },
        TemplateUsageDetail: { allOf: [ref('TemplateUsageSummary'), { type: 'object', properties: {
          hourly48: { type: 'array', items: { type: 'object', properties: { hour: { type: 'string' }, renders: { type: 'integer' }, errors: { type: 'integer' }, views: { type: 'integer' } } } },
          daily30: { type: 'array', items: { type: 'object', properties: { day: { type: 'string' }, renders: { type: 'integer' }, errors: { type: 'integer' }, views: { type: 'integer' } } } },
          sources: { type: 'object', properties: { ui: { type: 'integer' }, api: { type: 'integer' }, link: { type: 'integer' } } },
          topClients: { type: 'array', items: { type: 'object', properties: { client: { type: 'string' }, renders: { type: 'integer' } } } },
          recent: { type: 'array', items: ref('UsageEvent') } } }] },
        UsageSummary: { type: 'object', properties: {
          generatedAt: { type: 'string', format: 'date-time' },
          totals: ref('GlobalTotals'),
          templates: { type: 'array', items: ref('TemplateUsageSummary') },
          recent: { type: 'array', items: ref('UsageEvent') } } },
      },
    },
  };
}

module.exports = {
  buildSchema, buildExample, validateBody, bodyFromQuery, safeFileName, fillPattern, buildOpenApi,
  // utilidades expuestas para pruebas/integración
  RESERVED_QUERY, IMAGE_RE,
};
