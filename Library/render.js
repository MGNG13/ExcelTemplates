// XLSX render pipeline: template JSON + data body → workbook Buffer.
// Behaviour is identical to the v11 inline implementation in main.js for sheets without
// repeat zones; zones (rows/cols, see repeat-zones.js) are expanded first into a plain sheet.
'use strict';
const ExcelJS = require('exceljs');
const { injectShapes, applyDefaultFont, stripPageSetup, restoreSheetFormatPr } = require('./shapes.js');
const {
  colIdxToLetter, colLetterToIdx, parseAddr, guessExt, hexToArgb,
  applyStyle, setCellValue, embedImage, DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PX,
  colPxToChars, rowPxToPt, colPxToFrac, rowPxToFrac,
} = require('./excel-utils.js');
const { expandSheet } = require('./repeat-zones.js');

/**
 * Render one sheet into the workbook.
 * @returns {{ shapes: Array|null }} shapes moved by repeat zones (null → use the template's)
 */
function renderSheet(wb, sheetTpl, data = {}, opts = {}) {
  // Repeat zones (rows/cols): expand into a plain sheet + per-cell data lookup
  const orig = sheetTpl;
  const ex = expandSheet(sheetTpl, data, {
    warn: opts.warn,
    rowFrac: px => rowPxToFrac(px, orig.rowHeights || {}),
    colFrac: px => colPxToFrac(px, orig.columnWidths || {}),
  });
  if (ex) sheetTpl = ex.sheet;
  const dataAt = ex ? ex.dataAt : () => data;
  const v  = sheetTpl.view || {};
  const wsOpts = {
    views: [{
      zoomScale:     v.zoomScale || 100,
      showGridLines: v.showGridLines !== false,
      state:         v.state || 'normal',
      xSplit:        v.xSplit || 0,
      ySplit:        v.ySplit || 0,
    }],
    properties: {
      defaultColWidth:  colPxToChars(sheetTpl.defaultColWidth  || DEFAULT_COL_WIDTH_PX),
      defaultRowHeight: (sheetTpl.rawCustomHeight || !sheetTpl.rawDefaultRowHeightPt)
                          ? rowPxToPt(sheetTpl.defaultRowHeight || DEFAULT_ROW_HEIGHT_PX)
                          : sheetTpl.rawDefaultRowHeightPt,
    },
  };
  if (sheetTpl.state && sheetTpl.state !== 'visible') wsOpts.state = sheetTpl.state;
  if (sheetTpl.pageSetup) {
    const ps = { ...sheetTpl.pageSetup };
    if (ps.margins) wsOpts.pageSetup = { margins: ps.margins };
    wsOpts.pageSetup = { ...(wsOpts.pageSetup || {}), ...ps };
  }
  if (sheetTpl.tabColor) wsOpts.properties.tabColor = { argb: hexToArgb(sheetTpl.tabColor) };
  const ws = wb.addWorksheet(sheetTpl.name, wsOpts);

  // Column widths + default styles
  Object.entries(sheetTpl.columnWidths || {}).forEach(([letter, px]) => {
    ws.getColumn(colLetterToIdx(letter)).width = Math.max(colPxToChars(px), 0.5);
  });
  const colStyleToExcel = st => {
    if (!st) return null;
    const o = {};
    const font = {};
    if (st.fontFamily) font.name = st.fontFamily;
    if (st.fontSize) font.size = st.fontSize;
    if (st.bold) font.bold = true; if (st.italic) font.italic = true;
    if (st.color && st.color !== '#000000') font.color = { argb: hexToArgb(st.color) };
    if (Object.keys(font).length) o.font = font;
    if (st.bgColor) o.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(st.bgColor) } };
    if (st.align || st.valign || st.wrap) o.alignment = { horizontal: st.align, vertical: st.valign, wrapText: !!st.wrap };
    if (st.numFmt) o.numFmt = st.numFmt;
    return o;
  };
  const colStyles = sheetTpl.colStyles || {};
  Object.entries(colStyles).forEach(([letter, st]) => {
    const ex = colStyleToExcel(st); if (ex) ws.getColumn(colLetterToIdx(letter)).style = ex;
  });
  if (sheetTpl.defaultColStyle) {
    // Apply the catch-all default to every column up to a wide bound
    const ex = colStyleToExcel(sheetTpl.defaultColStyle);
    const maxUsed = Math.max(26, ...Object.keys(colStyles).map(colLetterToIdx), ...Object.keys(sheetTpl.columnWidths || {}).map(colLetterToIdx));
    for (let c = 1; c <= maxUsed + 40; c++) if (!colStyles[colIdxToLetter(c)] && ex) ws.getColumn(c).style = ex;
  }

  // Normal rendering (with repeat zones the sheet was expanded above; dataAt gives each cell's data)
  Object.entries(sheetTpl.rowHeights || {}).forEach(([r, px]) => ws.getRow(+r).height = rowPxToPt(px));
  Object.entries(sheetTpl.cells || {}).forEach(([a, cd]) => {
    const cell = ws.getCell(a);
    setCellValue(cell, cd, dataAt(a));
    applyStyle(cell, cd.style, cd.numFmt);
  });

  // Images pass (after cells so images overlay)
  Object.entries(sheetTpl.cells || {}).forEach(([a, cd]) => {
    const p = parseAddr(a); if (!p) return;
    const c0 = p.col - 1, r0 = p.row - 1;
    if (cd.image) {
      embedImage(wb, ws, c0, r0, cd.image.src, cd.image.width, cd.image.height);
      // Auto-size row to image height
      const pts = rowPxToPt(cd.image.height || 100);
      if (!ws.getRow(p.row).height || ws.getRow(p.row).height < pts)
        ws.getRow(p.row).height = pts;
    }
    const cdata = dataAt(a);
    if (cd.variable?.type === 'image' && cdata && cdata[cd.variable.name]) {
      const src = cdata[cd.variable.name];
      embedImage(wb, ws, c0, r0, src, 150, 100);
    }
  });

  // Hidden rows / columns and outline levels (preserve collapsed groups)
  (sheetTpl.hiddenCols || []).forEach(c => { ws.getColumn(c).hidden = true; });
  Object.entries(sheetTpl.colOutline || {}).forEach(([c, lvl]) => { ws.getColumn(+c).outlineLevel = lvl; });
  // (rows below repeat zones were already shifted by expandSheet)
  (sheetTpl.hiddenRows || []).forEach(r => { ws.getRow(r).hidden = true; });
  Object.entries(sheetTpl.rowOutline || {}).forEach(([r, lvl]) => { ws.getRow(+r).outlineLevel = lvl; });

  // Merges
  (sheetTpl.merges || []).forEach(m => { try { ws.mergeCells(m); } catch (_) {} });

  // Floating images — anchored to cells like Excel does
  const cwPx = c => (sheetTpl.columnWidths || {})[colIdxToLetter(c)] || sheetTpl.defaultColWidth || DEFAULT_COL_WIDTH_PX;
  const rhPx = r => (sheetTpl.rowHeights || {})[r] || (sheetTpl.autoRowHeights || {})[r] || sheetTpl.defaultRowHeight || DEFAULT_ROW_HEIGHT_PX;
  (sheetTpl.floatingImages || []).forEach(fi => {
    try {
      const base64 = fi.src.includes(',') ? fi.src.split(',')[1] : fi.src;
      const ext    = fi.ext || guessExt(fi.src) || 'png';
      const imgId  = wb.addImage({ base64, extension: ext });
      let tl;
      const a = fi.anchor;
      if (fi.__tl) {
        tl = { col: fi.__tl.col, row: fi.__tl.row };   // pixel-positioned image moved by a repeat zone
      } else if (a && a.from) {
        tl = { col: a.from.col + (a.from.colOff / cwPx(a.from.col + 1)),
               row: a.from.row + (a.from.rowOff / rhPx(a.from.row + 1)) };
      } else {
        tl = { col: colPxToFrac(fi.x, sheetTpl.columnWidths || {}),
               row: rowPxToFrac(fi.y, sheetTpl.rowHeights   || {}) };
      }
      ws.addImage(imgId, { tl, ext: { width: fi.width || 200, height: fi.height || 150 } });
    } catch (e) { console.warn('float img render:', e.message); }
  });
  return { shapes: ex ? sheetTpl.shapes : null };
}

// Sheet names produced for a template sheet given the body (repeatable sheets
// with >1 item become "Name 1", "Name 2", …).
function expandedNames(s, body) {
  if (s.isRepeatable && Array.isArray(body[s.name]) && body[s.name].length > 1)
    return body[s.name].map((_, i) => `${s.name} ${i + 1}`);
  return [s.name];
}

/**
 * Render a template with data.
 * @param {object} template  template JSON (sheets[])
 * @param {object} body      data keyed by sheet name (array for repeatable sheets)
 * @param {object} [opts]    { warnings: [] } receives { sheet, code, zone, message } for
 *                           invalid/overlapping repeat zones or zone data that is not an array
 * @returns {Promise<Buffer>} xlsx file
 */
async function renderWorkbook(template, body, opts = {}) {
  if (!template || !Array.isArray(template.sheets)) throw new TypeError('Plantilla inválida');
  body = (body && typeof body === 'object') ? body : {};
  // Work on a copy: the caller's object must never be touched
  template = JSON.parse(JSON.stringify(template));
  const seenWarn = new Set();
  const warn = w => {
    const k = (w.sheet || '') + '|' + w.message;
    if (seenWarn.has(k)) return;
    seenWarn.add(k);
    if (Array.isArray(opts.warnings)) opts.warnings.push(w);
  };
  const results = new Map();   // rendered sheet name → renderSheet result
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Excel Template Builder'; wb.created = new Date();

  for (const sheetTpl of template.sheets) {
    const sd = body[sheetTpl.name];
    if (sheetTpl.isRepeatable && Array.isArray(sd)) {
      sd.forEach((item, i) => {
        const clone = JSON.parse(JSON.stringify(sheetTpl));
        clone.name  = sd.length === 1 ? sheetTpl.name : `${sheetTpl.name} ${i + 1}`;
        results.set(clone.name, renderSheet(wb, clone, item, { warn }));
      });
    } else {
      results.set(sheetTpl.name, renderSheet(wb, sheetTpl, sd || {}, { warn }));
    }
  }

  let buf = Buffer.from(await wb.xlsx.writeBuffer());

  // Restore the original workbook default font (affects column width rendering)
  const defFont = template.defaultFont || template.sheets.find(s => s.defaultFont)?.defaultFont;
  if (defFont) { try { buf = applyDefaultFont(Buffer.from(buf), defFont); } catch (e) { console.warn('default font:', e.message); } }

  // Sheets whose original had no <pageSetup> keep printer defaults
  const noPS = [];
  template.sheets.filter(s => s.hasPageSetup === false).forEach(s => noPS.push(...expandedNames(s, body)));
  if (noPS.length) { try { buf = stripPageSetup(Buffer.from(buf), noPS); } catch (e) { console.warn('strip pageSetup:', e.message); } }

  // Restore each sheet's original <sheetFormatPr> (row auto-fit behaviour)
  const sfpMap = {};
  template.sheets.forEach(s => {
    if (s.rawSheetFormatPr) expandedNames(s, body).forEach(n => { sfpMap[n] = s.rawSheetFormatPr; });
  });
  if (Object.keys(sfpMap).length) { try { buf = restoreSheetFormatPr(Buffer.from(buf), sfpMap); } catch (e) { console.warn('sheetFormatPr:', e.message); } }

  // Re-inject drawing shapes (ExcelJS cannot write these)
  const shapesBySheet = {};
  template.sheets.forEach(s => {
    if (!s.shapes?.length) return;
    expandedNames(s, body).forEach(n => { shapesBySheet[n] = results.get(n)?.shapes || s.shapes; });
  });
  if (Object.keys(shapesBySheet).length) {
    try { buf = injectShapes(Buffer.from(buf), shapesBySheet); }
    catch (e) { console.warn('shape inject:', e.message); }
  }
  return Buffer.from(buf);
}

module.exports = { renderWorkbook, renderSheet };
