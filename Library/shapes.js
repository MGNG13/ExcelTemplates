'use strict';
// ═════════════════════════════════════════════════════════════════════════════
//  shapes.js — Extract & inject Excel drawing shapes + images (xl/drawings/)
//  ExcelJS cannot reliably handle shapes/pictures; we parse the OOXML directly.
// ═════════════════════════════════════════════════════════════════════════════

const AdmZip     = require('adm-zip');
const { XMLParser } = require('fast-xml-parser');

const EMU_PER_PX = 9525;
const emuToPx    = e => Math.round((Number(e) || 0) / EMU_PER_PX);
const pxToEmu    = p => Math.round((Number(p) || 0) * EMU_PER_PX);

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@',
  removeNSPrefix: true,    // xdr:sp → sp, a:t → t, r:embed → embed
  parseTagValue: false,
  trimValues: false,
});

const arr = x => (x === undefined || x === null) ? [] : (Array.isArray(x) ? x : [x]);

// ── Office default theme palette ──────────────────────────────────────────────
const DEFAULT_THEME = {
  dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
  accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5',
  accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
  hlink: '0563C1', folHlink: '954F72',
};

function hexToRgb(h) {
  const n = parseInt(String(h).replace('#', ''), 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}
function rgbToHex({ r, g, b }) {
  const c = v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0');
  return (c(r) + c(g) + c(b)).toUpperCase();
}
function applyTint(hex, tint) {
  if (!tint) return hex;
  const { r, g, b } = hexToRgb(hex);
  const f = v => tint > 0 ? v + (255 - v) * tint : v * (1 + tint);
  return rgbToHex({ r: f(r), g: f(g), b: f(b) });
}
function applyLum(hex, lumMod, lumOff) {
  if (lumMod === undefined && lumOff === undefined) return hex;
  let { r, g, b } = hexToRgb(hex);
  const m = lumMod !== undefined ? lumMod / 100000 : 1;
  const o = lumOff !== undefined ? lumOff / 100000 : 0;
  const f = v => (v / 255 * m + o) * 255;
  return rgbToHex({ r: f(r), g: f(g), b: f(b) });
}

function resolveClr(node, theme) {
  if (!node) return null;
  const t = theme || DEFAULT_THEME;
  let hex = null, mods = null;
  if (node.srgbClr)  { hex = node.srgbClr['@val'];  mods = node.srgbClr; }
  else if (node.schemeClr) {
    let name = node.schemeClr['@val'];
    if (name === 'tx1') name = 'dk1'; if (name === 'tx2') name = 'dk2';
    if (name === 'bg1') name = 'lt1'; if (name === 'bg2') name = 'lt2';
    hex = t[name] || DEFAULT_THEME[name] || '000000'; mods = node.schemeClr;
  } else if (node.sysClr) { hex = node.sysClr['@lastClr'] || '000000'; mods = node.sysClr; }
  if (!hex) return null;
  if (mods) {
    const lm = mods.lumMod?.['@val'] !== undefined ? parseInt(mods.lumMod['@val']) : undefined;
    const lo = mods.lumOff?.['@val'] !== undefined ? parseInt(mods.lumOff['@val']) : undefined;
    hex = applyLum(hex, lm, lo);
    if (mods.tint?.['@val'])  hex = applyTint(hex,  parseInt(mods.tint['@val'])  / 100000);
    if (mods.shade?.['@val']) hex = applyTint(hex, -1 + parseInt(mods.shade['@val']) / 100000);
  }
  return '#' + String(hex).toUpperCase();
}

function parseTheme(zip, parserRef) {
  const e = zip.getEntry('xl/theme/theme1.xml');
  if (!e) return { ...DEFAULT_THEME };
  try {
    const x = parserRef.parse(zip.readAsText(e));
    const scheme = x?.theme?.themeElements?.clrScheme;
    if (!scheme) return { ...DEFAULT_THEME };
    const out = { ...DEFAULT_THEME };
    ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink']
      .forEach(k => {
        const n = scheme[k];
        if (!n) return;
        const v = n.srgbClr?.['@val'] || n.sysClr?.['@lastClr'];
        if (v) out[k] = String(v).toUpperCase();
      });
    return out;
  } catch { return { ...DEFAULT_THEME }; }
}

// ── Anchor position → pixel coords ───────────────────────────────────────────
function anchorToPx(anchor, colWidths, rowHeights, defColPx, defRowPx) {
  const col    = parseInt(anchor?.col    ?? 0) || 0;
  const colOff = parseInt(anchor?.colOff ?? 0) || 0;
  const row    = parseInt(anchor?.row    ?? 0) || 0;
  const rowOff = parseInt(anchor?.rowOff ?? 0) || 0;
  let x = 0; for (let i = 0; i < col; i++) x += colWidths[i + 1] || defColPx;
  x += emuToPx(colOff);
  let y = 0; for (let i = 0; i < row; i++) y += rowHeights[i + 1] || defRowPx;
  y += emuToPx(rowOff);
  return { x, y };
}

// ── Extract text from txBody ──────────────────────────────────────────────────
function extractText(txBody, theme, styleNode) {
  const empty = { text: '', bold: false, size: 11, color: null, align: 'ctr', vAnchor: 'ctr', lines: [] };
  if (!txBody) return empty;
  const paras = arr(txBody.p);
  const lines = [];   // [{ text, bold, size, color, align }]
  let bold = false, size = null, color = null, align = null;
  const vAnchor = txBody.bodyPr?.['@anchor'] || 'ctr';

  paras.forEach(p => {
    const pAlign = p?.pPr?.['@algn'] || null;
    if (pAlign && !align) align = pAlign;
    const runs = arr(p?.r);
    let lBold = false, lSize = null, lColor = null;
    const text = runs.map(r => {
      const rPr = r?.rPr;
      if (rPr?.['@b'] === '1' || rPr?.['@b'] === 'true') { bold = true; lBold = true; }
      if (rPr?.['@sz']) { const sz = Math.round(parseInt(rPr['@sz']) / 100); size = size || sz; lSize = sz; }
      const c = rPr?.solidFill ? resolveClr(rPr.solidFill, theme) : null;
      if (c) { color = color || c; lColor = c; }
      const t = r?.t;
      return typeof t === 'object' ? (t['#text'] ?? '') : (t ?? '');
    }).join('');
    // endParaRPr can carry size for empty paragraphs
    if (!lSize && p?.endParaRPr?.['@sz']) lSize = Math.round(parseInt(p.endParaRPr['@sz']) / 100);
    lines.push({ text, bold: lBold, size: lSize, color: lColor, align: pAlign });
  });

  // Fallback text colour from <xdr:style><a:fontRef>
  if (!color && styleNode?.fontRef) color = resolveClr(styleNode.fontRef, theme);

  const a = align || 'ctr';
  return {
    text: lines.map(l => l.text).join('\n'),
    lines, bold, size: size || 11, color: color || '#000000',
    align: a === 'l' ? 'left' : a === 'r' ? 'right' : 'center',
    vAnchor: vAnchor === 't' ? 'top' : vAnchor === 'b' ? 'bottom' : 'middle',
  };
}

// ── Extract style from spPr ───────────────────────────────────────────────────
function extractSpPr(spPr, theme, styleNode, kind) {
  const geom = spPr?.prstGeom?.['@prst'] || 'rect';
  let fill = null, stroke = null, strokeW = 1, dash = null;
  let headEnd = null, tailEnd = null;

  // Explicit fill in spPr wins
  if (spPr?.noFill !== undefined)      fill = 'none';
  else if (spPr?.solidFill)            fill = resolveClr(spPr.solidFill, theme);
  else if (spPr?.gradFill) {
    const first = arr(spPr.gradFill?.gsLst?.gs)[0];
    fill = resolveClr(first, theme);
  }

  const ln = spPr?.ln;
  if (ln) {
    if (ln.noFill !== undefined)  stroke = 'none';
    else if (ln.solidFill)        stroke = resolveClr(ln.solidFill, theme);
    if (ln['@w']) strokeW = Math.max(1, Math.round(Number(ln['@w']) / 12700));
    const d = ln.prstDash?.['@val'];
    if (d && d !== 'solid') dash = d;
    // Arrowheads (headEnd = start of line, tailEnd = end of line)
    if (ln.headEnd?.['@type'] && ln.headEnd['@type'] !== 'none') headEnd = ln.headEnd['@type'];
    if (ln.tailEnd?.['@type'] && ln.tailEnd['@type'] !== 'none') tailEnd = ln.tailEnd['@type'];
  }

  // ── Fall back to the shape's <xdr:style> theme references ──────────────
  //    fillRef idx="0" → no fill; idx>=1 → use referenced colour
  //    lnRef   idx="0" → no line; idx>=1 → use referenced colour
  if (styleNode) {
    if (fill === null) {
      const fr = styleNode.fillRef;
      const idx = parseInt(fr?.['@idx'] ?? '0');
      if (idx === 0) fill = 'none';
      else fill = resolveClr(fr, theme) || null;
    }
    if (stroke === null) {
      const lr = styleNode.lnRef;
      const idx = parseInt(lr?.['@idx'] ?? '0');
      if (idx === 0) stroke = 'none';
      else stroke = resolveClr(lr, theme) || '#000000';
    }
  }
  // Connectors without any explicit line info still need a visible stroke
  if (kind === 'connector' && (stroke === null || stroke === 'none')) stroke = '#000000';
  // Plain shapes: default black outline if nothing said otherwise
  if (kind !== 'connector' && stroke === null) stroke = '#000000';
  if (fill === null) fill = 'none';
  const adj = {};
  arr(spPr?.prstGeom?.avLst?.gd).forEach(g => {
    if (g['@name']) adj[g['@name']] = parseInt(g['@fmla']?.replace('val ', '') || '0');
  });
  const xfrm = spPr?.xfrm;
  return { geom, fill, stroke, strokeW, dash, adj, headEnd, tailEnd,
           rot:   xfrm?.['@rot'] ? Math.round(parseInt(xfrm['@rot']) / 60000) : 0,
           flipH: xfrm?.['@flipH'] === '1', flipV: xfrm?.['@flipV'] === '1' };
}

// ── Normalize a relative path (resolve .. segments) ───────────────────────────
function normalizePath(base, rel) {
  const parts = [...base.split('/'), ...rel.split('/')];
  const out = [];
  parts.forEach(s => {
    if (s === '..') out.pop();
    else if (s && s !== '.') out.push(s);
  });
  return out.join('/');
}

// ── Parse one drawing anchor → { shapes[], pictures[] } ──────────────────────
function parseAnchor(node, kind, colWidths, rowHeights, defColPx, defRowPx, theme) {
  const shapes   = [];
  const pictures = [];   // { embedRid, x, y, width, height }

  // ── Compute position & size from the anchor ────────────────────────────
  const mk = a => ({
    col: parseInt(a?.col ?? 0) || 0, row: parseInt(a?.row ?? 0) || 0,
    colOff: emuToPx(a?.colOff ?? 0), rowOff: emuToPx(a?.rowOff ?? 0),
  });
  let pos, size, anchor;
  if (kind === 'two') {
    pos  = anchorToPx(node.from, colWidths, rowHeights, defColPx, defRowPx);
    const to = anchorToPx(node.to, colWidths, rowHeights, defColPx, defRowPx);
    size = { width: Math.max(4, to.x - pos.x), height: Math.max(4, to.y - pos.y) };
    anchor = { type: 'two', from: mk(node.from), to: mk(node.to) };
  } else if (kind === 'one') {
    pos  = anchorToPx(node.from, colWidths, rowHeights, defColPx, defRowPx);
    size = { width: emuToPx(node.ext?.['@cx'] ?? node.ext?.cx ?? 0),
             height: emuToPx(node.ext?.['@cy'] ?? node.ext?.cy ?? 0) };
    anchor = { type: 'one', from: mk(node.from), ext: { w: size.width, h: size.height } };
  } else { // absolute
    pos  = { x: emuToPx(node.pos?.['@x'] ?? 0), y: emuToPx(node.pos?.['@y'] ?? 0) };
    size = { width: emuToPx(node.ext?.['@cx'] ?? 0), height: emuToPx(node.ext?.['@cy'] ?? 0) };
    anchor = { type: 'abs', x: pos.x, y: pos.y, w: size.width, h: size.height };
  }
  if (!pos || isNaN(pos.x)) pos = { x: 0, y: 0 };
  if (!size || isNaN(size.width) || size.width < 4) size = { width: 100, height: 60 };

  // ── Collect shapes (sp + cxnSp) ───────────────────────────────────────
  const shapeNodes = [
    ...arr(node.sp).map(s    => ({ n: s, type: 'shape' })),
    ...arr(node.cxnSp).map(s => ({ n: s, type: 'connector' })),
  ];
  arr(node.grpSp).forEach(g => {
    arr(g.sp).forEach(s    => shapeNodes.push({ n: s, type: 'shape' }));
    arr(g.cxnSp).forEach(s => shapeNodes.push({ n: s, type: 'connector' }));
  });

  shapeNodes.forEach(({ n, type }) => {
    const spPr  = n.spPr;
    const style = extractSpPr(spPr, theme, n.style, type);
    const txt   = extractText(n.txBody, theme, n.style);
    let w = size.width, h = size.height, x = pos.x, y = pos.y;
    if (spPr?.xfrm?.ext && shapeNodes.length > 1) {
      const ew = emuToPx(spPr.xfrm.ext['@cx']); const eh = emuToPx(spPr.xfrm.ext['@cy']);
      if (ew > 0) w = ew; if (eh > 0) h = eh;
    }
    shapes.push({
      id: Math.random().toString(36).slice(2) + Date.now().toString(36),
      kind: type, geom: style.geom, x, y, width: Math.max(4, w), height: Math.max(4, h),
      fill: style.fill, stroke: style.stroke, strokeWidth: style.strokeW,
      dash: style.dash, adj: style.adj,
      headEnd: style.headEnd, tailEnd: style.tailEnd,
      anchor,
      rot: style.rot, flipH: style.flipH, flipV: style.flipV,
      text: txt.text, textBold: txt.bold, textSize: txt.size,
      textColor: txt.color, textAlign: txt.align, textVAnchor: txt.vAnchor,
      textLines: txt.lines,
    });
  });

  // ── Collect pictures (xdr:pic) — images embedded in the drawing ────────
  const picNodes = [...arr(node.pic)];
  arr(node.grpSp).forEach(g => arr(g.pic).forEach(p => picNodes.push(p)));

  picNodes.forEach(pic => {
    // r:embed is stripped to @embed by the namespace-removing parser
    const embedRid = pic?.blipFill?.blip?.['@embed']
                  || pic?.blipFill?.blip?.['@r:embed'];
    if (!embedRid) return;
    pictures.push({ embedRid, x: pos.x, y: pos.y,
                    width: Math.max(4, size.width), height: Math.max(4, size.height), anchor });
  });

  return { shapes, pictures };
}

// ═════════════════════════════════════════════════════════════════════════════
//  PUBLIC: extract shapes AND images for every sheet in an xlsx buffer
//  Returns: Map<sheetName, { shapes: shape[], images: floatingImage[] }>
// ═════════════════════════════════════════════════════════════════════════════
function extractShapes(buffer, sheetGeometry) {
  const out = new Map();
  let zip;
  try { zip = new AdmZip(buffer); } catch { return out; }

  const theme  = parseTheme(zip, parser);
  const readXml = p => {
    const e = zip.getEntry(p); if (!e) return null;
    try { return parser.parse(zip.readAsText(e)); } catch { return null; }
  };

  // workbook → sheet list
  const wbXml  = readXml('xl/workbook.xml'); if (!wbXml) return out;
  const sheets = arr(wbXml.workbook?.sheets?.sheet);
  const wbRels = readXml('xl/_rels/workbook.xml.rels');
  const relMap = {};
  arr(wbRels?.Relationships?.Relationship).forEach(r => { relMap[r['@Id']] = r['@Target']; });

  sheets.forEach(sh => {
    const name   = sh['@name'];
    const rId    = sh['@id'] || sh['@r:id'];
    let   target = relMap[rId];
    if (!target) return;
    target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
    const sheetPath = 'xl/' + target;

    // Find this sheet's drawing relationship
    const shBase    = sheetPath.split('/').pop();
    const shDir     = sheetPath.slice(0, -shBase.length);
    const shRels    = readXml(shDir + '_rels/' + shBase + '.rels');
    let   drawRid   = null;
    const shXml     = readXml(sheetPath);
    drawRid = shXml?.worksheet?.drawing?.['@id']
           || shXml?.worksheet?.drawing?.['@r:id'];
    if (!drawRid) return;

    let drawTarget = null;
    arr(shRels?.Relationships?.Relationship).forEach(r => {
      if (r['@Id'] === drawRid) drawTarget = r['@Target'];
    });
    if (!drawTarget) return;

    const drawPath  = normalizePath(shDir.replace(/\/$/, ''), drawTarget);
    const dXml      = readXml(drawPath); if (!dXml) return;
    const ws        = dXml?.wsDr; if (!ws) return;

    // Drawing rels (for resolving image embedRid → file path)
    const drawBase  = drawPath.split('/').pop();
    const drawDir   = drawPath.slice(0, -drawBase.length);
    const drawRels  = readXml(drawDir + '_rels/' + drawBase + '.rels');
    const drawRelMap = {};
    arr(drawRels?.Relationships?.Relationship).forEach(r => { drawRelMap[r['@Id']] = r['@Target']; });

    const geom = sheetGeometry[name] || { colWidths: {}, rowHeights: {}, defColPx: 64, defRowPx: 20 };

    // Collect shapes and pictures from every anchor type
    const allShapes   = [];
    const rawPictures = [];

    ['twoCellAnchor','oneCellAnchor','absoluteAnchor'].forEach((kind, ki) => {
      const kindKey = ['two','one','abs'][ki];
      arr(ws[kind]).forEach(a => {
        const r = parseAnchor(a, kindKey, geom.colWidths, geom.rowHeights, geom.defColPx, geom.defRowPx, theme);
        allShapes.push(...r.shapes);
        rawPictures.push(...r.pictures);
      });
    });

    // Resolve pictures → base64 floating images
    const images = [];
    rawPictures.forEach(pic => {
      const target = drawRelMap[pic.embedRid];
      if (!target) { console.warn(`No rel for ${pic.embedRid} in ${drawPath}`); return; }

      const imgPath = normalizePath(drawDir.replace(/\/$/, ''), target);
      const imgEntry = zip.getEntry(imgPath);
      if (!imgEntry) { console.warn('Image not found in zip:', imgPath); return; }

      const buf = imgEntry.getData();
      if (!buf || buf.length === 0) { console.warn('Empty image buffer:', imgPath); return; }

      const rawExt = imgPath.split('.').pop().toLowerCase();
      const ext    = rawExt === 'jpg' ? 'jpeg' : rawExt;
      const src    = `data:image/${ext};base64,${buf.toString('base64')}`;

      images.push({
        id: Math.random().toString(36).slice(2) + Date.now().toString(36),
        src, ext, anchor: pic.anchor,
        x: Math.max(0, Math.round(pic.x)),
        y: Math.max(0, Math.round(pic.y)),
        width:  Math.max(20, Math.round(pic.width)),
        height: Math.max(20, Math.round(pic.height)),
      });
    });

    if (allShapes.length || images.length) {
      out.set(name, { shapes: allShapes, images });
    }
  });

  return out;
}

// ═════════════════════════════════════════════════════════════════════════════
//  Inject shapes into generated xlsx buffer (ExcelJS can't write shapes)
// ═════════════════════════════════════════════════════════════════════════════
function xmlEsc(s) {
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}

function shapeToXml(sh, idx) {
  const cx = pxToEmu(sh.width), cy = pxToEmu(sh.height);
  const x  = pxToEmu(sh.x),    y  = pxToEmu(sh.y);
  const geom = sh.geom || 'rect';
  const fillXml = (!sh.fill || sh.fill === 'none')
    ? '<a:noFill/>'
    : `<a:solidFill><a:srgbClr val="${sh.fill.replace('#','').toUpperCase()}"/></a:solidFill>`;
  const lnXml = (!sh.stroke || sh.stroke === 'none')
    ? '<a:ln><a:noFill/></a:ln>'
    : `<a:ln w="${Math.round((sh.strokeWidth||1)*12700)}"><a:solidFill><a:srgbClr val="${sh.stroke.replace('#','').toUpperCase()}"/></a:solidFill>${sh.dash?`<a:prstDash val="${sh.dash}"/>`:''}${sh.headEnd?`<a:headEnd type="${sh.headEnd}"/>`:''}${sh.tailEnd?`<a:tailEnd type="${sh.tailEnd}"/>`:''}</a:ln>`;
  const rotAttr  = sh.rot   ? ` rot="${Math.round(sh.rot * 60000)}"` : '';
  const flipAttr = (sh.flipH?' flipH="1"':'') + (sh.flipV?' flipV="1"':'');
  let txBody = '';
  if (sh.text) {
    const defAlgn = sh.textAlign === 'left' ? 'l' : sh.textAlign === 'right' ? 'r' : 'ctr';
    const anchor  = sh.textVAnchor === 'top' ? 't' : sh.textVAnchor === 'bottom' ? 'b' : 'ctr';
    const lineDefs = (sh.textLines && sh.textLines.length)
      ? sh.textLines
      : String(sh.text).split('\n').map(t => ({ text: t }));
    const paras = lineDefs.map(l => {
      const algn = l.align ? (l.align === 'l' ? 'l' : l.align === 'r' ? 'r' : 'ctr') : defAlgn;
      const sz   = (l.size || sh.textSize || 11) * 100;
      const b    = (l.bold !== undefined ? l.bold : sh.textBold) ? ' b="1"' : '';
      const col  = (l.color || sh.textColor || '#000000').replace('#','').toUpperCase();
      if (!l.text) return `<a:p><a:pPr algn="${algn}"/><a:endParaRPr lang="es-MX" sz="${sz}"/></a:p>`;
      return `<a:p><a:pPr algn="${algn}"/><a:r><a:rPr lang="es-MX" sz="${sz}"${b}><a:solidFill><a:srgbClr val="${col}"/></a:solidFill></a:rPr><a:t>${xmlEsc(l.text)}</a:t></a:r></a:p>`;
    }).join('');
    txBody = `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" anchor="${anchor}"/><a:lstStyle/>${paras}</xdr:txBody>`;
  } else {
    txBody = '<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="es-MX"/></a:p></xdr:txBody>';
  }
  const adjEntries = Object.entries(sh.adj || {});
  const avLst = adjEntries.length
    ? `<a:avLst>${adjEntries.map(([k, v]) => `<a:gd name="${k}" fmla="val ${v}"/>`).join('')}</a:avLst>`
    : '<a:avLst/>';
  const tag  = sh.kind === 'connector' ? 'cxnSp' : 'sp';
  const nvPr = sh.kind === 'connector'
    ? `<xdr:nvCxnSpPr><xdr:cNvPr id="${idx+2}" name="Connector ${idx+1}"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>`
    : `<xdr:nvSpPr><xdr:cNvPr id="${idx+2}" name="Shape ${idx+1}"/><xdr:cNvSpPr/></xdr:nvSpPr>`;
  const body = sh.kind === 'connector'
    ? `${nvPr}<xdr:spPr><a:xfrm${rotAttr}${flipAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${geom}">${avLst}</a:prstGeom>${lnXml}</xdr:spPr>`
    : `${nvPr}<xdr:spPr><a:xfrm${rotAttr}${flipAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${geom}">${avLst}</a:prstGeom>${fillXml}${lnXml}</xdr:spPr>${txBody}`;
  const a = sh.anchor;
  const mkA = (t, o) => `<xdr:${t}><xdr:col>${o.col}</xdr:col><xdr:colOff>${pxToEmu(o.colOff)}</xdr:colOff><xdr:row>${o.row}</xdr:row><xdr:rowOff>${pxToEmu(o.rowOff)}</xdr:rowOff></xdr:${t}>`;
  if (a && a.type === 'two' && a.from && a.to)
    return `<xdr:twoCellAnchor>${mkA('from', a.from)}${mkA('to', a.to)}<xdr:${tag}>${body}</xdr:${tag}><xdr:clientData/></xdr:twoCellAnchor>`;
  if (a && a.type === 'one' && a.from)
    return `<xdr:oneCellAnchor>${mkA('from', a.from)}<xdr:ext cx="${cx}" cy="${cy}"/><xdr:${tag}>${body}</xdr:${tag}><xdr:clientData/></xdr:oneCellAnchor>`;
  return `<xdr:absoluteAnchor><xdr:pos x="${x}" y="${y}"/><xdr:ext cx="${cx}" cy="${cy}"/><xdr:${tag}>${body}</xdr:${tag}><xdr:clientData/></xdr:absoluteAnchor>`;
}

function buildDrawingXml(shapes) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${shapes.map((s,i)=>shapeToXml(s,i)).join('')}</xdr:wsDr>`;
}

function injectShapes(buffer, shapesBySheet) {
  const entries = Object.entries(shapesBySheet).filter(([,s])=>s&&s.length);
  if (!entries.length) return buffer;
  let zip; try { zip = new AdmZip(buffer); } catch { return buffer; }
  const readText = p => { const e=zip.getEntry(p); return e?zip.readAsText(e):null; };
  const wbXml  = readText('xl/workbook.xml');
  const wbRels = readText('xl/_rels/workbook.xml.rels');
  if (!wbXml || !wbRels) return buffer;
  const wbParsed = parser.parse(wbXml);
  const wbRelsParsed = parser.parse(wbRels);
  const relMap = {};
  arr(wbRelsParsed?.Relationships?.Relationship).forEach(r => { relMap[r['@Id']] = r['@Target']; });
  const sheets = arr(wbParsed.workbook?.sheets?.sheet);

  // Find highest existing drawingN.xml so new ones don't collide
  let dc = 0;
  zip.getEntries().forEach(e => {
    const m = e.entryName.match(/^xl\/drawings\/drawing(\d+)\.xml$/);
    if (m) dc = Math.max(dc, parseInt(m[1]));
  });
  const ctAdd = [];

  entries.forEach(([sheetName, shapes]) => {
    const sh = sheets.find(s => s['@name'] === sheetName); if (!sh) return;
    const rId = sh['@id'] || sh['@r:id'];
    let target = relMap[rId]; if (!target) return;
    target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
    const sheetPath = 'xl/' + target;
    let shXml = readText(sheetPath); if (!shXml) return;
    const base = sheetPath.split('/').pop(), dir = sheetPath.slice(0,-base.length);
    const relPath = dir + '_rels/' + base + '.rels';
    const existRel = readText(relPath);

    // ── Case A: sheet already has a drawing (e.g. ExcelJS wrote images) → MERGE ──
    const dm = shXml.match(/<drawing[^>]*r:id="([^"]+)"/);
    if (dm && existRel) {
      const relParsed = parser.parse(existRel);
      let drawTarget = null;
      arr(relParsed?.Relationships?.Relationship).forEach(r => { if (r['@Id'] === dm[1]) drawTarget = r['@Target']; });
      if (drawTarget) {
        const drawPath = normalizePath(dir.replace(/\/$/, ''), drawTarget);
        let dXml = readText(drawPath);
        if (dXml) {
          // Make sure the DrawingML "a" namespace is declared on the root
          if (!/xmlns:a=/.test(dXml))
            dXml = dXml.replace('<xdr:wsDr ', '<xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ');
          // Use high, non-colliding cNvPr ids
          const extra = shapes.map((s, i) => shapeToXml(s, 5000 + i)).join('');
          dXml = dXml.replace('</xdr:wsDr>', extra + '</xdr:wsDr>');
          zip.updateFile(drawPath, Buffer.from(dXml, 'utf8'));
          return;
        }
      }
    }

    // ── Case B: no drawing yet → create one ─────────────────────────────
    dc++;
    const drawFile = `drawing${dc}.xml`;
    const drawPath = `xl/drawings/${drawFile}`;
    const newRid   = 'rIdShapes' + dc;
    zip.addFile(drawPath, Buffer.from(buildDrawingXml(shapes), 'utf8'));
    ctAdd.push(`<Override PartName="/${drawPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
    const newRel = `<Relationship Id="${newRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawFile}"/>`;
    if (existRel) zip.updateFile(relPath, Buffer.from(existRel.replace('</Relationships>', newRel+'</Relationships>'), 'utf8'));
    else zip.addFile(relPath, Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${newRel}</Relationships>`, 'utf8'));
    // <drawing> must come before <legacyDrawing>/<tableParts>/<extLst>; safest: before </worksheet> unless those exist
    const drawingTag = `<drawing r:id="${newRid}"/>`;
    const beforeTags = ['<legacyDrawing', '<legacyDrawingHF', '<picture', '<oleObjects', '<controls', '<webPublishItems', '<tableParts', '<extLst'];
    let inserted = false;
    for (const t of beforeTags) {
      const i = shXml.indexOf(t);
      if (i !== -1) { shXml = shXml.slice(0, i) + drawingTag + shXml.slice(i); inserted = true; break; }
    }
    if (!inserted) shXml = shXml.replace('</worksheet>', drawingTag + '</worksheet>');
    if (!/xmlns:r=/.test(shXml)) shXml = shXml.replace('<worksheet ','<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
    zip.updateFile(sheetPath, Buffer.from(shXml, 'utf8'));
  });
  if (ctAdd.length) {
    let ct = readText('[Content_Types].xml');
    if (ct) zip.updateFile('[Content_Types].xml', Buffer.from(ct.replace('</Types>', ctAdd.join('')+'</Types>'), 'utf8'));
  }
  return zip.toBuffer();
}

// Per-sheet <sheetFormatPr> + workbook default font — needed to reproduce
// Excel's real default row height (which is font-based unless customHeight=1)
function readSheetFormats(buffer) {
  const out = { sheets: {}, defaultFont: { name: 'Calibri', size: 11 } };
  let zip; try { zip = new AdmZip(buffer); } catch { return out; }
  const readXml = p => { const e = zip.getEntry(p); if (!e) return null; try { return parser.parse(zip.readAsText(e)); } catch { return null; } };

  // Default font = first <font> in styles.xml
  const styles = readXml('xl/styles.xml');
  const f0 = arr(styles?.styleSheet?.fonts?.font)[0];
  if (f0) {
    if (f0.sz?.['@val'])     out.defaultFont.size   = parseFloat(f0.sz['@val']);
    if (f0.name?.['@val'])   out.defaultFont.name   = f0.name['@val'];
    if (f0.family?.['@val']) out.defaultFont.family = parseInt(f0.family['@val']);
    if (f0.scheme?.['@val']) out.defaultFont.scheme = f0.scheme['@val'];
  }

  const wbXml  = readXml('xl/workbook.xml'); if (!wbXml) return out;
  const wbRels = readXml('xl/_rels/workbook.xml.rels');
  const relMap = {};
  arr(wbRels?.Relationships?.Relationship).forEach(r => { relMap[r['@Id']] = r['@Target']; });
  arr(wbXml.workbook?.sheets?.sheet).forEach(sh => {
    const rId = sh['@id'] || sh['@r:id'];
    let target = relMap[rId]; if (!target) return;
    target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
    const x = readXml('xl/' + target);
    const fp = x?.worksheet?.sheetFormatPr || {};
    const rawSheet = zip.getEntry('xl/' + target) ? zip.readAsText(zip.getEntry('xl/' + target)) : '';
    const sfp = rawSheet.match(/<sheetFormatPr[^>]*\/>/);
    out.sheets[sh['@name']] = {
      hasPageSetup: /<pageSetup[\s\/>]/.test(rawSheet),
      rawSheetFormatPr: sfp ? sfp[0] : null,
      defaultRowHeightPt: fp['@defaultRowHeight'] !== undefined ? parseFloat(fp['@defaultRowHeight']) : null,
      customHeight:       fp['@customHeight'] === '1' || fp['@customHeight'] === 'true',
      baseColWidth:       fp['@baseColWidth'] !== undefined ? parseFloat(fp['@baseColWidth']) : null,
      defaultColWidth:    fp['@defaultColWidth'] !== undefined ? parseFloat(fp['@defaultColWidth']) : null,
      state: sh['@state'] || 'visible',
    };
  });
  return out;
}

// Replace the workbook default font (fonts[0] in styles.xml) so column
// widths — expressed in "characters" of that font — render identically.
function applyDefaultFont(buffer, font) {
  if (!font || !font.name) return buffer;
  let zip; try { zip = new AdmZip(buffer); } catch { return buffer; }
  const e = zip.getEntry('xl/styles.xml'); if (!e) return buffer;
  let xml = zip.readAsText(e);
  const m = xml.match(/<fonts[^>]*>\s*<font>[\s\S]*?<\/font>/);
  if (!m) return buffer;
  const first = m[0];
  const sz   = font.size || 11;
  const fam  = font.family || 2;
  const repl = first.replace(/<font>[\s\S]*?<\/font>/,
    `<font><sz val="${sz}"/><color theme="1"/><name val="${font.name.replace(/&/g,'&amp;').replace(/"/g,'&quot;')}"/><family val="${fam}"/>${font.scheme?`<scheme val="${font.scheme}"/>`:''}</font>`);
  xml = xml.replace(first, repl);
  zip.updateFile('xl/styles.xml', Buffer.from(xml, 'utf8'));
  return zip.toBuffer();
}

// Apply per-sheet XML patch functions: { sheetName: xml => xml }
function patchSheets(buffer, patches) {
  const names = Object.keys(patches || {});
  if (!names.length) return buffer;
  let zip; try { zip = new AdmZip(buffer); } catch { return buffer; }
  const readText = p => { const e = zip.getEntry(p); return e ? zip.readAsText(e) : null; };
  const wbXml = readText('xl/workbook.xml'), wbRels = readText('xl/_rels/workbook.xml.rels');
  if (!wbXml || !wbRels) return buffer;
  const relMap = {};
  arr(parser.parse(wbRels)?.Relationships?.Relationship).forEach(r => { relMap[r['@Id']] = r['@Target']; });
  arr(parser.parse(wbXml).workbook?.sheets?.sheet).forEach(sh => {
    const fn = patches[sh['@name']]; if (!fn) return;
    const rId = sh['@id'] || sh['@r:id']; let target = relMap[rId]; if (!target) return;
    target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
    const path = 'xl/' + target; let xml = readText(path); if (!xml) return;
    const out = fn(xml);
    if (out && out !== xml) zip.updateFile(path, Buffer.from(out, 'utf8'));
  });
  return zip.toBuffer();
}

// Remove <pageSetup> from sheets whose original had none (keeps printer defaults)
function stripPageSetup(buffer, sheetNames) {
  const patches = {};
  (sheetNames || []).forEach(n => { patches[n] = xml => xml.replace(/<pageSetup[^>]*\/>/, ''); });
  return patchSheets(buffer, patches);
}

// Restore the original <sheetFormatPr> verbatim (ExcelJS forces customHeight=1)
function restoreSheetFormatPr(buffer, map) {
  const patches = {};
  Object.entries(map || {}).forEach(([n, raw]) => {
    if (!raw) return;
    patches[n] = xml => xml.replace(/<sheetFormatPr[^>]*\/>/, raw.replace(/\$/g, '$$$$'));
  });
  return patchSheets(buffer, patches);
}

// Read just the theme palette from an xlsx buffer
function readTheme(buffer) {
  try { return parseTheme(new AdmZip(buffer), parser); }
  catch { return { ...DEFAULT_THEME }; }
}

module.exports = { extractShapes, injectShapes, readTheme, readSheetFormats, applyDefaultFont, stripPageSetup, restoreSheetFormatPr, patchSheets, applyTint, DEFAULT_THEME };
