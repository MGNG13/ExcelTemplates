// Shared Excel helpers: address math, colour resolution, unit conversion,
// Spanish→English formula normalisation and cell styling/value writing.
// Extracted verbatim from the v11 main.js so render/import behave identically.
'use strict';
const { applyTint, DEFAULT_THEME } = require('./shapes.js');
// Column counterpart of adjustFormulaRows + a reference-aware formula rewriter (repeat zones)
const { adjustFormulaCols, mapFormulaRefs } = require('./repeat-zones.js');

function colIdxToLetter(idx) {
  let r = '';
  while (idx > 0) { const m = (idx - 1) % 26; r = String.fromCharCode(65 + m) + r; idx = Math.floor((idx - 1) / 26); }
  return r;
}
function colLetterToIdx(letter) {
  return String(letter).toUpperCase().split('').reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0);
}
function parseAddr(addr) {
  const m = String(addr).match(/^([A-Z]+)(\d+)$/i);
  if (!m) return null;
  return { col: colLetterToIdx(m[1]), row: parseInt(m[2]) };
}
function guessExt(src) {
  const m = String(src).match(/data:image\/(\w+);/);
  if (m) return m[1] === 'jpeg' ? 'jpg' : m[1];
  return 'png';
}
// Excel legacy indexed colour palette (indices 0-63)
const INDEXED_COLORS = [
  '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
  '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF',
  '800000','008000','000080','808000','800080','008080','C0C0C0','808080',
  '9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF',
  '000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF',
  '00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99',
  '3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696',
  '003366','339966','003300','333300','993300','993366','333399','333333',
];

// Theme slot order used by Excel's cell styles (differs from DrawingML naming)
const THEME_SLOTS = ['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'];

// Resolve an ExcelJS colour object → '#RRGGBB' (or null)
function resolveExcelColor(c, theme) {
  if (!c) return null;
  if (c.argb) {
    const a = String(c.argb);
    return '#' + (a.length === 8 ? a.slice(2) : a).toUpperCase();
  }
  if (c.theme !== undefined) {
    const slot = THEME_SLOTS[c.theme];
    let hex = (theme && theme[slot]) || DEFAULT_THEME[slot] || '000000';
    if (c.tint) hex = applyTint(hex, c.tint);
    return '#' + String(hex).replace('#', '').toUpperCase();
  }
  if (c.indexed !== undefined) {
    const hex = INDEXED_COLORS[c.indexed];
    if (hex) return '#' + hex;
  }
  return null;
}

function hexToArgb(hex) {
  if (!hex || hex === 'transparent') return null;
  return 'FF' + hex.replace('#', '').toUpperCase().padStart(6, '0');
}
const ES_ESCAPE = '\x00ESCAPED\x00';
function _hideStrings(formula) {
  const stash = [];
  let i = 0, out = '', inStr = false, quote = '';
  while (i < formula.length) {
    const ch = formula[i];
    if (!inStr && (ch === '"' || ch === "'")) { inStr = true; quote = ch; out += ch; }
    else if (inStr && ch === quote) {
      if (formula[i + 1] === quote) { out += ch + ch; i += 2; continue; }
      inStr = false; quote = ''; out += ch;
    } else if (inStr) {
      out += ES_ESCAPE + stash.length + ES_ESCAPE;
      stash.push(ch);
    } else {
      out += ch;
    }
    i++;
  }
  return { out, stash };
}
function _restoreStrings(str, stash) {
  return str.replace(new RegExp(ES_ESCAPE + '(\\d+)' + ES_ESCAPE, 'g'), (_, i) => stash[+i] || '');
}
const ES_TO_EN_FUNCTIONS = {
  SI: 'IF', SIERROR: 'IFERROR', Y: 'AND', O: 'OR', NO: 'NOT',
  SUMA: 'SUM', SUMAR: 'SUM', SUMARSI: 'SUMIF', 'SUMAR.SI': 'SUMIF', 'SUMAR.SI.CONJUNTO': 'SUMIFS', SUMARSISET: 'SUMIFS',
  CONTAR: 'COUNT', CONTARA: 'COUNTA', 'CONTAR.SI': 'COUNTIF', CONTARSI: 'COUNTIF', 'CONTAR.SI.CONJUNTO': 'COUNTIFS', CONTARSISET: 'COUNTIFS', 'CONTAR.BLANCO': 'COUNTBLANK', CONTARBLANCO: 'COUNTBLANK',
  PROMEDIO: 'AVERAGE', MEDIA: 'AVERAGE', 'PROMEDIO.SI': 'AVERAGEIF', PROMEDIOSI: 'AVERAGEIF', 'PROMEDIO.SI.CONJUNTO': 'AVERAGEIFS', PROMEDIOSISET: 'AVERAGEIFS',
  MINIMO: 'MIN', 'MÍNIMO': 'MIN', MAXIMO: 'MAX', 'MÁXIMO': 'MAX',
  BUSCARV: 'VLOOKUP', BUSCARH: 'HLOOKUP', BUSCARX: 'XLOOKUP', COINCIDIR: 'MATCH', INDICE: 'INDEX',
  'INDICE.COINCIDIR': 'XMATCH',
  CONCATENAR: 'CONCAT', CONCAT: 'CONCAT', TEXTO: 'TEXT', VALOR: 'VALUE', NUMERO: 'NUMBERVALUE',
  IZQUIERDA: 'LEFT', DERECHA: 'RIGHT', MEDIO: 'MID', LARGO: 'LEN', LONGITUD: 'LEN',
  ESPACIOS: 'TRIM', MINUSCULA: 'LOWER', MAYUSCULA: 'UPPER', PROPER: 'PROPER',
  REEMPLAZAR: 'SUBSTITUTE', SUSTITUIR: 'SUBSTITUTE', BUSCAR: 'SEARCH', HALLAR: 'FIND',
  REDONDEAR: 'ROUND', 'REDONDEAR.MAS': 'ROUNDUP', 'REDONDEAR.MENOS': 'ROUNDDOWN', TRUNCAR: 'TRUNC',
  ENTERO: 'INT', RESIDUO: 'MOD', 'DIVISION.ENTERA': 'QUOTIENT',
  ALEATORIO: 'RAND', 'ALEATORIO.ENTRE': 'RANDBETWEEN',
  HOY: 'TODAY', AHORA: 'NOW', FECHA: 'DATE', MES: 'MONTH', ANO: 'YEAR', 'AÑO': 'YEAR', DIA: 'DAY', HORA: 'HOUR', MINUTO: 'MINUTE', SEGUNDO: 'SECOND',
  DIAS: 'DAYS', 'DIAS.LABORABLES': 'WORKDAY', 'DIAS.LABORABLES.INTL': 'WORKDAY.INTL',
  SEMANA: 'WEEKDAY', 'NUM.DE.SEMANA': 'WEEKNUM',
  'ES.ERROR': 'ISERROR', ESERR: 'ISERR', 'ES.ND': 'ISNA', ESBLANCO: 'ISBLANK',
  LOG: 'LOG', LOG10: 'LOG10', LN: 'LN', EXP: 'EXP', POTENCIA: 'POWER', RAIZ: 'SQRT',
  ABS: 'ABS', PI: 'PI',
  UNIR: 'TEXTJOIN', UNIRCADENAS: 'TEXTJOIN', 'TEXTO.UNIR': 'TEXTJOIN',
  DESREF: 'OFFSET', INDIRECTO: 'INDIRECT',
  DECABSCISO: 'CEILING', MATOABSCISO: 'FLOOR',
};
function normalizeFormulaEStoEN(formula) {
  if (typeof formula !== 'string' || !formula) return formula;
  let f = formula.startsWith('=') ? formula.slice(1) : formula;
  const { out, stash } = _hideStrings(f);
  let norm = out.replace(/;/g, ',');
  const longestFirst = Object.keys(ES_TO_EN_FUNCTIONS).sort((a, b) => b.length - a.length);
  const escPattern = new RegExp('\\b(' + longestFirst.map(k => k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|') + ')\\s*\\(', 'gi');
  norm = norm.replace(escPattern, m => {
    const hasTrailing = /\(\s*$/.test(m);
    const raw = m.replace(/\(\s*$/, '').trim();
    const lookup = raw.toUpperCase();
    const replaceFn = ES_TO_EN_FUNCTIONS[lookup];
    return (replaceFn || m.replace(/\s*\(\s*$/, '')) + (hasTrailing ? m.match(/\(\s*$/)[0] : '');
  });
  norm = norm.replace(/\bVERDADERO\b/gi, 'TRUE').replace(/\bFALSO\b/gi, 'FALSE');
  return _restoreStrings(norm, stash);
}

function substituteVars(str, data) {
  if (typeof str !== 'string') return str;
  return str.replace(/\{\{(\w+)\}\}/g, (_, k) => (data && Object.prototype.hasOwnProperty.call(data, k) && data[k] !== undefined && data[k] !== null ? data[k] : ''));
}
function adjustFormulaRows(formula, offset) {
  if (!offset) return formula;
  return formula.replace(/(\$?[A-Z]+)(\$?)(\d+)/g, (_, col, abs, row) =>
    abs === '$' ? _ : col + (parseInt(row) + offset));
}

function applyStyle(excelCell, style, numFmt) {
  if (numFmt) excelCell.numFmt = numFmt;
  if (!style) return;
  const font = { name: style.fontFamily || 'Calibri', size: style.fontSize || 11 };
  if (style.bold)      font.bold      = true;
  if (style.italic)    font.italic    = true;
  if (style.underline) font.underline = true;
  const ca = hexToArgb(style.color);
  if (ca && style.color !== '#000000') font.color = { argb: ca };
  excelCell.font = font;
  const ba = hexToArgb(style.bgColor);
  if (ba) excelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ba } };
  excelCell.alignment = {
    horizontal: style.align  || 'left',
    vertical:   style.valign || 'middle',
    wrapText:   !!style.wrap,
  };
  if (style.border) {
    const bs = style.borderStyles || {};
    const b  = {};
    ['top','bottom','left','right'].forEach(side => {
      if (!style.border[side]) return;
      const spec = bs[side];
      b[side] = {
        style: spec?.style || 'thin',
        color: { argb: hexToArgb(spec?.color || '#000000') },
      };
    });
    if (Object.keys(b).length) excelCell.border = b;
  }
}

const DEFAULT_COL_WIDTH_PX = 64;   // Excel default 8.43 chars
const DEFAULT_ROW_HEIGHT_PX = 20;  // Excel default 15 pt

// Excel column width is in "characters" (MDW=7px for Calibri 11)
const colCharsToPx = w => Math.round((Number(w) || 0) * 7 + 5);
const colPxToChars = px => Math.max(0, ((Number(px) || 0) - 5) / 7);
// Row height is in points; 96 dpi / 72 pt = 4/3
const rowPtToPx = pt => Math.round((Number(pt) || 0) * 4 / 3);
const rowPxToPt = px => (Number(px) || 0) * 3 / 4;

// Pixel → fractional col index (0-based) for ExcelJS image placement
function colPxToFrac(px, colWidths) {
  let rem = px, col = 0;
  while (rem > 0 && col < 26) {
    const w = colWidths[colIdxToLetter(col + 1)] || DEFAULT_COL_WIDTH_PX;
    if (rem <= w) return col + rem / w;
    rem -= w; col++;
  }
  return col;
}
// Pixel → fractional row index (0-based)
function rowPxToFrac(px, rowHeights) {
  let rem = px, row = 0;
  while (rem > 0 && row < 100) {
    const h = rowHeights[row + 1] || DEFAULT_ROW_HEIGHT_PX;
    if (rem <= h) return row + rem / h;
    rem -= h; row++;
  }
  return row;
}
// Fractional col index → pixel position (for import)
function fracColToPx(frac, colWidths) {
  const whole = Math.floor(frac), part = frac - whole;
  let px = 0;
  for (let i = 0; i < whole; i++) px += colWidths[colIdxToLetter(i + 1)] || DEFAULT_COL_WIDTH_PX;
  px += part * (colWidths[colIdxToLetter(whole + 1)] || DEFAULT_COL_WIDTH_PX);
  return px;
}
// Fractional row index → pixel position (for import)
function fracRowToPx(frac, rowHeights) {
  const whole = Math.floor(frac), part = frac - whole;
  let px = 0;
  for (let i = 0; i < whole; i++) px += rowHeights[i + 1] || DEFAULT_ROW_HEIGHT_PX;
  px += part * (rowHeights[whole + 1] || DEFAULT_ROW_HEIGHT_PX);
  return px;
}

function embedImage(wb, ws, col0, row0, src, w, h) {
  try {
    const base64 = src.includes(',') ? src.split(',')[1] : src;
    const ext    = guessExt(src);
    const imgId  = wb.addImage({ base64, extension: ext });
    ws.addImage(imgId, { tl: { col: col0, row: row0 }, ext: { width: w || 150, height: h || 100 } });
  } catch (e) { console.warn('img embed:', e.message); }
}

function setCellValue(cell, cd, data) {
  if (cd.image) return; // handled separately
  let val;
  if (cd.formula) {
    const normalized = normalizeFormulaEStoEN(substituteVars(cd.formula.slice(1), data));
    val = { formula: normalized };
  } else if (cd.variable) {
    const raw = data && Object.prototype.hasOwnProperty.call(data, cd.variable.name) && data[cd.variable.name] !== undefined ? data[cd.variable.name] : '';
    switch (cd.variable.type) {
      case 'number':  val = isNaN(parseFloat(raw)) ? 0 : parseFloat(raw); break;
      case 'boolean': {
        const sv = typeof raw === 'string' ? raw.trim().toLowerCase() : raw;
        val = (sv === 'false' || sv === '0' || sv === '' || sv === 'no' || sv == null) ? false : (sv === 'true' || sv === '1' || sv === 'si' || sv === 'sí') ? true : Boolean(sv);
        break;
      }
      case 'date':    val = raw ? new Date(raw) : null; break;
      case 'image':   return; // handled in image pass
      default:        val = raw;
    }
  } else if (cd.emptyStyled && !cd.value) {
    val = null;   // empty cell that only carries formatting
  } else {
    val = substituteVars(cd.value || '', data);
    // Restore the original Excel data type so number formats render correctly
    if (cd.dataType === 'number' && val !== '' && !isNaN(Number(val))) val = Number(val);
    else if (cd.dataType === 'boolean') val = val === 'true' || val === true;
    else if (cd.dataType === 'date' && val) { const d = new Date(val); if (!isNaN(d)) val = d; }
  }
  cell.value = val;
}

module.exports = {
  colIdxToLetter, colLetterToIdx, parseAddr, guessExt, resolveExcelColor, hexToArgb,
  normalizeFormulaEStoEN, substituteVars, adjustFormulaRows, adjustFormulaCols, mapFormulaRefs, applyStyle, setCellValue, embedImage,
  DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PX, colCharsToPx, colPxToChars, rowPtToPx, rowPxToPt,
  colPxToFrac, rowPxToFrac, fracColToPx, fracRowToPx, INDEXED_COLORS, THEME_SLOTS,
};
