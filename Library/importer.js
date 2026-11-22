// XLSX import pipeline: workbook Buffer → template JSON (id: null).
// Extracted verbatim from the v11 POST /api/import handler.
'use strict';
const ExcelJS = require('exceljs');
const { extractShapes, readTheme, readSheetFormats } = require('./shapes.js');
const {
  colIdxToLetter, colLetterToIdx, resolveExcelColor, DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PX,
  colCharsToPx, rowPtToPx,
} = require('./excel-utils.js');

/**
 * @param {Buffer} buffer        xlsx contents
 * @param {string} originalName  uploaded file name (template name = name without extension)
 * @returns {Promise<object>} template JSON
 */
async function importWorkbook(buffer, originalName) {
  const file = { buffer, originalname: String(originalName || 'Plantilla.xlsx') };
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(file.buffer);
  const theme   = readTheme(file.buffer);        // resolve theme colours accurately
  const formats = readSheetFormats(file.buffer); // real default row/col sizes

  const template = {
    id: null,
    name: file.originalname.replace(/\.xlsx?$/i, ''),
    sheets: [],
  };

  wb.eachSheet(ws => {
    const sheet = {
      id: Date.now().toString(36) + Math.random().toString(36).slice(2),
      name: ws.name, cells: {}, columnWidths: {}, rowHeights: [],
      merges: [], isRepeatable: false, repeatVariable: '', rowTemplate: null,
      state: ws.state || 'visible',   // visible | hidden | veryHidden
      tabColor: ws.properties?.tabColor?.argb ? '#' + String(ws.properties.tabColor.argb).slice(-6) : null,
    };
    sheet.rowHeights = {};

    // Column widths
    // Sheet view state (zoom, gridlines, frozen panes)
    const view = (ws.views && ws.views[0]) || {};
    sheet.view = {
      zoomScale:     view.zoomScale || 100,
      showGridLines: view.showGridLines !== false,
      state:         view.state || 'normal',
      xSplit:        view.xSplit || 0,
      ySplit:        view.ySplit || 0,
    };
    // Page setup / print area / margins (so PDF & print match the original)
    const ps = ws.pageSetup || {};
    sheet.pageSetup = {
      orientation:  ps.orientation || 'portrait',
      paperSize:    ps.paperSize,
      scale:        ps.scale,
      fitToPage:    !!ps.fitToPage,
      fitToWidth:   ps.fitToWidth,
      fitToHeight:  ps.fitToHeight,
      printArea:    ps.printArea,
      printTitlesRow: ps.printTitlesRow,
      printTitlesColumn: ps.printTitlesColumn,
      horizontalCentered: !!ps.horizontalCentered,
      verticalCentered:   !!ps.verticalCentered,
      showGridLines: !!ps.showGridLines,
      blackAndWhite: !!ps.blackAndWhite,
      margins: ps.margins ? { ...ps.margins } : undefined,
    };
    Object.keys(sheet.pageSetup).forEach(k => sheet.pageSetup[k] === undefined && delete sheet.pageSetup[k]);
    sheet.hasPageSetup = !!(formats.sheets[ws.name]?.hasPageSetup);
    sheet.rawSheetFormatPr = formats.sheets[ws.name]?.rawSheetFormatPr || null;
    sheet.hiddenRows = [];
    sheet.hiddenCols = [];
    sheet.rowOutline = {};
    sheet.colOutline = {};

    // ── Excel's real defaults ────────────────────────────────────────────
    const fmt   = formats.sheets[ws.name] || {};
    const fontPt = formats.defaultFont.size || 11;
    // Row: honour defaultRowHeight only when customHeight=1; otherwise Excel
    // derives it from the default font (Calibri 11 → 15pt → 20px)
    let defRowPx;
    if (fmt.customHeight && fmt.defaultRowHeightPt) defRowPx = rowPtToPx(fmt.defaultRowHeightPt);
    else defRowPx = Math.round(fontPt * 1.818);
    // (refined below once column default styles are known)
    // Column: defaultColWidth if given, else baseColWidth + padding (~0.71 chars)
    let defColChars = fmt.defaultColWidth || (fmt.baseColWidth ? fmt.baseColWidth + 0.71 : null) || ws.properties?.defaultColWidth || 8.43;
    sheet.defaultColWidth  = colCharsToPx(defColChars);
    sheet.defaultRowHeight = defRowPx;
    sheet.defaultFont      = { ...formats.defaultFont };
    sheet.rawDefaultRowHeightPt = fmt.defaultRowHeightPt || null;
    sheet.rawCustomHeight       = !!fmt.customHeight;
    template.defaultFont = { ...formats.defaultFont };
    if (fmt.state && fmt.state !== 'visible') sheet.state = fmt.state;
    // Column widths — use the raw <col min max> ranges, NOT ws.columns (which
    // expands a <col max="16384"> catch-all into 16k entries and freezes the UI).
    // A range spanning many columns is the sheet's effective default width.
    let lastUsedCol = 1;
    ws.eachRow({ includeEmpty: false }, row => {
      if (row.cellCount) lastUsedCol = Math.max(lastUsedCol, row.actualCellCount ? row.cellCount : row.cellCount);
    });
    (ws.model?.merges || []).forEach(m => {
      const rng = typeof m === 'string' ? m : (m?.range || '');
      const mm = rng.match(/^[A-Z]+\d+:([A-Z]+)\d+$/);
      if (mm) lastUsedCol = Math.max(lastUsedCol, colLetterToIdx(mm[1]));
    });
    const colCap = Math.min(16384, lastUsedCol + 40);
    const rawCols = ws.model?.cols || [];
    // Column default styles (apply to empty cells → affect auto row height)
    sheet.colStyles = {};
    let sheetDefaultColStyle = null;
    const styleFromCol = st => {
      if (!st) return null;
      const o = {};
      if (st.font) {
        o.fontFamily = st.font.name; o.fontSize = st.font.size;
        o.bold = !!st.font.bold; o.italic = !!st.font.italic;
        const fc = resolveExcelColor(st.font.color, theme); if (fc) o.color = fc;
      }
      if (st.fill?.type === 'pattern' && st.fill.pattern !== 'none') {
        const bg = resolveExcelColor(st.fill.fgColor, theme); if (bg) o.bgColor = bg;
      }
      if (st.alignment) { o.align = st.alignment.horizontal; o.valign = st.alignment.vertical; o.wrap = !!st.alignment.wrapText; }
      if (st.numFmt) o.numFmt = st.numFmt;
      return Object.keys(o).length ? o : null;
    };
    rawCols.forEach(c => {
      const min = c.min || 1, max = c.max || min;
      const span = max - min + 1;
      if (c.width > 0 && span >= 50) {
        // Catch-all range → sheet default
        sheet.defaultColWidth = colCharsToPx(c.width);
      }
      const cst = styleFromCol(c.style);
      if (cst && span >= 50) sheetDefaultColStyle = cst;
      for (let ci = min; ci <= Math.min(max, colCap); ci++) {
        if (c.width > 0 && span < 50) sheet.columnWidths[colIdxToLetter(ci)] = colCharsToPx(c.width);
        if (c.hidden) sheet.hiddenCols.push(ci);
        if (c.outlineLevel) sheet.colOutline[ci] = c.outlineLevel;
        if (cst && span < 50) sheet.colStyles[colIdxToLetter(ci)] = cst;
      }
    });

    if (sheetDefaultColStyle) sheet.defaultColStyle = sheetDefaultColStyle;
    // Excel auto-fits default rows to the largest font present; with a column
    // default style smaller than the workbook font, rows shrink accordingly.
    if (!fmt.customHeight) {
      const colFontPt = sheetDefaultColStyle?.fontSize
        || Math.max(0, ...Object.values(sheet.colStyles).map(c => c.fontSize || 0)) || null;
      if (colFontPt) sheet.defaultRowHeight = Math.round(Math.max(colFontPt, fmt.defaultRowHeightPt ? fmt.defaultRowHeightPt / 1.364 : 0) * 1.818);
    }

    // Rows & cells
    ws.eachRow({ includeEmpty: false }, (row, rn) => {
      if (row.height) sheet.rowHeights[rn] = rowPtToPx(row.height);
      if (row.hidden) sheet.hiddenRows.push(rn);
      if (row.outlineLevel) sheet.rowOutline[rn] = row.outlineLevel;
      row.eachCell({ includeEmpty: true }, (cell, cn) => {
        if (cn > colCap) return;
        const a = colIdxToLetter(cn) + rn;
        const cd = {
          value: '', formula: null, variable: null,
          style: {
            bold: false, italic: false, underline: false,
            color: '#000000', bgColor: '#ffffff', fontSize: 11,
            align: 'left', valign: 'bottom', wrap: false,
            border: { top: false, bottom: false, left: false, right: false },
          },
        };
        // Value — the hidden cells of a merge only keep their style: ExcelJS reports the master's value for
        // them, which would duplicate texts/{{variables}} and make empty cells look filled (e.g. zone crossings).
        const mergedSlave = cell.type === ExcelJS.ValueType.Merge || (cell.isMerged && cell.master && cell.master.address !== cell.address);
        if (mergedSlave) {
          /* sin valor propio */
        } else if (cell.type === ExcelJS.ValueType.Formula) {
          cd.formula = '=' + (cell.formula || '');
          // Use computed result (cell.result) so cells show values, not formula text
          const res = cell.result;
          if (res !== null && res !== undefined) {
            if (typeof res === 'object' && res !== null && res.error) {
              cd.value = ''; // Excel error result (#VALUE!, etc.)
            } else if (res instanceof Date) {
              cd.value = res.toISOString().split('T')[0]; cd.dataType = 'date';
            } else if (typeof res === 'number') {
              cd.value = String(res); cd.dataType = 'number';   // so numFmt applies
            } else if (typeof res === 'boolean') {
              cd.value = String(res); cd.dataType = 'boolean';
            } else {
              cd.value = String(res);
            }
          } else {
            cd.value = '';
          }
        } else if (cell.value !== null && cell.value !== undefined) {
          if (cell.value instanceof Date) {
            cd.value = cell.value.toISOString().split('T')[0];
            cd.dataType = 'date';
          } else if (typeof cell.value === 'object' && cell.value.richText) {
            cd.value = cell.value.richText.map(r => r.text).join('');
          } else if (typeof cell.value === 'object' && cell.value.result !== undefined) {
            cd.value = String(cell.value.result);
          } else if (typeof cell.value === 'number') {
            cd.value = String(cell.value);
            cd.dataType = 'number';        // keep numeric so numFmt applies on export
          } else if (typeof cell.value === 'boolean') {
            cd.value = String(cell.value);
            cd.dataType = 'boolean';
          } else {
            cd.value = String(cell.value);
          }
        }
        // Font
        if (cell.font) {
          cd.style.bold      = !!cell.font.bold;
          cd.style.italic    = !!cell.font.italic;
          cd.style.underline = !!cell.font.underline;
          if (cell.font.size) cd.style.fontSize = cell.font.size;
          const fc = resolveExcelColor(cell.font.color, theme);
          if (fc) cd.style.color = fc;
          if (cell.font.name)   cd.style.fontFamily = cell.font.name;
        }
        // Fill — resolve argb, theme+tint and indexed colours alike
        if (cell.fill?.type === 'pattern' && cell.fill.pattern !== 'none') {
          const bg = resolveExcelColor(cell.fill.fgColor, theme);
          if (bg && bg.toUpperCase() !== '#FFFFFF') cd.style.bgColor = bg;
        }
        // Number format (so 10,000 / $ / % render like Excel)
        if (cell.numFmt) cd.numFmt = cell.numFmt;
        // Alignment
        if (cell.alignment) {
          cd.style.align  = cell.alignment.horizontal || 'left';
          cd.style.valign = cell.alignment.vertical   || 'bottom';
          cd.style.wrap   = !!cell.alignment.wrapText;
        }
        // Borders — keep the real line style and colour
        if (cell.border) {
          cd.style.border = {
            top:    !!cell.border.top?.style,
            bottom: !!cell.border.bottom?.style,
            left:   !!cell.border.left?.style,
            right:  !!cell.border.right?.style,
          };
          const bs = {};
          ['top','bottom','left','right'].forEach(side => {
            const b = cell.border[side];
            if (b?.style) bs[side] = { style: b.style, color: resolveExcelColor(b.color, theme) || '#000000' };
          });
          if (Object.keys(bs).length) cd.style.borderStyles = bs;
        }
        // Keep empty cells that carry formatting (Excel uses them for auto-fit, fills, borders)
        const df = formats.defaultFont || {};
        const colSt = sheet.colStyles[colIdxToLetter(cn)] || sheet.defaultColStyle || {};
        const baseFont = colSt.fontFamily || df.name, baseSize = colSt.fontSize || df.size || 11;
        const styled = !!(cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern !== 'none' && cd.style.bgColor !== '#ffffff')
                    || Object.values(cd.style.border).some(Boolean)
                    || (cell.font && ((cell.font.name && cell.font.name !== baseFont) || (cell.font.size && cell.font.size !== baseSize) || cell.font.bold || cell.font.italic));
        if (cd.value || cd.formula || styled) { if (!cd.value && !cd.formula) cd.emptyStyled = true; sheet.cells[a] = cd; }
      });
    });

    // Rows that are hidden but hold no cells are skipped by eachRow — catch them
    const lastRow = ws.actualRowCount ? ws.rowCount : ws.rowCount;
    for (let rn = 1; rn <= lastRow; rn++) {
      const r = ws.getRow(rn);
      if (r?.hidden && !sheet.hiddenRows.includes(rn)) sheet.hiddenRows.push(rn);
      if (r?.height && !sheet.rowHeights[rn]) sheet.rowHeights[rn] = rowPtToPx(r.height);
    }
    sheet.hiddenRows.sort((a, b) => a - b);

    // Merges — ExcelJS exposes these in several shapes depending on version
    const rawMerges = ws.model?.merges || ws._merges || [];
    if (Array.isArray(rawMerges)) {
      rawMerges.forEach(m => {
        if (typeof m === 'string') sheet.merges.push(m);
        else if (m?.range) sheet.merges.push(m.range);
        else if (m?.model?.range) sheet.merges.push(m.model.range);
      });
    } else if (typeof rawMerges === 'object') {
      Object.values(rawMerges).forEach(m => {
        const r = typeof m === 'string' ? m : (m?.range || m?.model?.range);
        if (r) sheet.merges.push(r);
      });
    }
    sheet.merges = [...new Set(sheet.merges)];

    // Images now extracted by extractShapes() via ZIP/drawing XML (more reliable)

    template.sheets.push(sheet);
  });

  // ── Extract shapes + images from drawing XML (ZIP-based, reliable) ──
  try {
    const geometry = {};
    template.sheets.forEach(s => {
      const cw = {};
      Object.entries(s.columnWidths).forEach(([L, px]) => { cw[colLetterToIdx(L)] = px; });
      geometry[s.name] = {
        colWidths: cw, rowHeights: s.rowHeights,
        defColPx: s.defaultColWidth  || DEFAULT_COL_WIDTH_PX,
        defRowPx: s.defaultRowHeight || DEFAULT_ROW_HEIGHT_PX,
      };
    });
    const extracted = extractShapes(file.buffer, geometry);
    let totalShapes = 0, totalImages = 0;
    template.sheets.forEach(s => {
      const data = extracted.get(s.name) || { shapes: [], images: [] };
      s.shapes = data.shapes || [];
      // Images from drawing XML override ExcelJS-extracted ones
      if (data.images && data.images.length > 0) {
        s.floatingImages = [...data.images, ...(s.floatingImages || [])];
        totalImages += data.images.length;
      }
      totalShapes += s.shapes.length;
    });
    if (totalShapes) console.log(`  → ${totalShapes} formas de dibujo extraídas`);
    if (totalImages) console.log(`  → ${totalImages} imágenes extraídas del dibujo`);
  } catch (e) { console.warn('shape/image extract:', e.message); }

  return template;
}

module.exports = { importWorkbook };
