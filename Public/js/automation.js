/* ═══════════════════════════════════════════════════════════════════════════
 *  Excel Template Builder v12 — Panel "Automatizar" (automation.js)
 *  Expone ETB.openAutomation(templateId | null).
 *  - Sin dependencias, sin handlers inline (CSP), todo texto externo escapado.
 *  - Degrada con gracia si faltan endpoints (/example, /validate, /usage).
 * ═══════════════════════════════════════════════════════════════════════════ */
(function () {
    'use strict';

    const ETB = window.ETB = window.ETB || {};

    // ── Utilidades con fallback (si index.html aún no definió el puente) ─────
    const esc = s => (typeof ETB.esc === 'function' ? ETB.esc(s) : String(s ?? '')
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&#39;'));
    const ico = name => (typeof ETB.ico === 'function' ? ETB.ico(name)
        : `<svg class="ico" aria-hidden="true"><use href="#${esc(name)}"/></svg>`);
    const toast = (msg, type) => { try { if (typeof ETB.toast === 'function') ETB.toast(msg, type || 'info'); } catch { /* sin toast */ } };

    /** Headers de identidad (§3.1) para que el servidor atribuya el uso a la UI. */
    function identityHeaders() {
        const h = {}, id = ETB.identity;
        if (id && id.clientId) {
            h['X-Client-Id'] = id.clientId;
            if (id.userId) h['X-User-Id'] = id.userId;
            if (id.name) h['X-Client-Name'] = encodeURIComponent(id.name);
        }
        return h;
    }

    /** fetch JSON propio: devuelve { ok, status, data } sin lanzar por HTTP (necesitamos distinguir 404/422). */
    async function req(method, path, body) {
        const opts = { method, headers: { Accept: 'application/json', ...identityHeaders() } };
        if (body !== undefined) { opts.headers['Content-Type'] = 'application/json'; opts.body = typeof body === 'string' ? body : JSON.stringify(body); }
        let res;
        try { res = await fetch(path, opts); }
        catch { return { ok: false, status: 0, data: { error: 'No se pudo conectar con el servidor.' } }; }
        let data = null;
        try { data = await res.json(); } catch { data = null; }
        return { ok: res.ok, status: res.status, data };
    }

    const origin = () => location.origin;
    const fmtBytes = n => (n == null || isNaN(n)) ? '—' : n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(1)} KB` : `${(n / 1048576).toFixed(2)} MB`;
    const fmtDate = iso => { if (!iso) return 'Nunca'; const d = new Date(iso); return isNaN(d) ? '—' : d.toLocaleString('es-MX', { dateStyle: 'medium', timeStyle: 'short' }); };
    /** 'hace 5 min' / 'hace 2 h' / fecha corta; con el texto completo para el title. */
    function fmtRel(iso) {
        if (!iso) return 'Nunca';
        const d = new Date(iso), s = (Date.now() - d.getTime()) / 1000;
        if (isNaN(s)) return '—';
        if (s < 45) return 'Hace un momento';
        if (s < 3600) return `Hace ${Math.round(s / 60)} min`;
        if (s < 86400) return `Hace ${Math.round(s / 3600)} h`;
        if (s < 86400 * 7) return `Hace ${Math.round(s / 86400)} d`;
        return d.toLocaleDateString('es-MX', { day: 'numeric', month: 'short' });
    }
    const reduceMotion = () => window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;

    // ── Estado del panel ──────────────────────────────────────────────────────
    const S = {
        root: null, lastFocus: null, open: false,
        templates: [], tplId: null, name: '', revision: null,
        sheets: [],            // normalizado: [{ name, isRepeatable, repeatVariable, rowTemplate, zones:[{direction,start,end,arrayVariable}], variables:[... zone] }]
        example: null,         // body de ejemplo
        opts: { format: 'xlsx', strict: false, fileName: '' },
        tab: 'curl', batchTab: 'curl', loadSeq: 0, unsubUsage: null,
        loadedId: null, editorTouched: false,  // conserva el cuerpo editado al reabrir la misma plantilla
    };

    const TABS = [
        ['curl', 'cURL', 'i-terminal'], ['powershell', 'PowerShell', 'i-terminal'], ['python', 'Python', 'i-code'],
        ['node', 'Node.js', 'i-code'], ['csharp', 'C#', 'i-code'], ['pa', 'Power Automate', 'i-zap'],
        ['n8n', 'n8n', 'i-webhook'], ['link', 'Enlace directo', 'i-link'],
    ];
    const TYPE_LABEL = { text: 'texto', number: 'número', boolean: 'booleano', date: 'fecha', image: 'imagen' };

    // ── Construcción del modal (una sola vez) ────────────────────────────────
    function build() {
        if (S.root) return S.root;
        const ov = document.createElement('div');
        ov.id = 'au-overlay';
        ov.className = 'au-overlay';
        ov.hidden = true;
        ov.innerHTML = `
<div class="au-modal" role="dialog" aria-modal="true" aria-labelledby="au-title">
  <header class="au-head">
    <div class="au-head-icon">${ico('i-zap')}</div>
    <div class="au-head-text">
      <h2 id="au-title">Automatizar <span class="au-tname" id="au-tname"></span> <span class="au-rev" id="au-rev" hidden></span></h2>
      <p class="au-sub">Genera archivos Excel desde cualquier sistema enviando datos JSON.</p>
    </div>
    <label class="au-picker"><span class="au-sr">Plantilla</span>
      <select id="au-tpl-select" aria-label="Plantilla a automatizar"><option value="">Cargando plantillas…</option></select>
    </label>
    <button type="button" class="au-icon-btn au-close" data-au="close" aria-label="Cerrar">${ico('i-x')}</button>
  </header>

  <ol class="au-steps" aria-label="Cómo funciona">
    <li><span class="au-step-ico">${ico('i-send')}</span><span><b>1</b> Envía tus datos JSON</span></li>
    <li class="au-step-arrow" aria-hidden="true">${ico('i-arrow-right')}</li>
    <li><span class="au-step-ico">${ico('i-server')}</span><span><b>2</b> El servidor llena la plantilla</span></li>
    <li class="au-step-arrow" aria-hidden="true">${ico('i-arrow-right')}</li>
    <li><span class="au-step-ico">${ico('i-file-spreadsheet')}</span><span><b>3</b> Recibes el Excel</span></li>
  </ol>

  <div class="au-body" id="au-body">
    <div class="au-empty" id="au-empty" hidden></div>
    <div class="au-grid" id="au-main">
      <div class="au-col">
        <section class="au-card au-c-ep" aria-labelledby="au-ep-h">
          <h3 id="au-ep-h">${ico('i-globe')} Endpoint</h3>
          <div class="au-endpoint"><span class="au-method">POST</span><code id="au-url"></code>
            <button type="button" class="au-icon-btn" data-au="copy-url" aria-label="Copiar URL">${ico('i-copy')}</button></div>
          <p class="au-help">Envía un JSON con <code>Content-Type: application/json</code>. Cada clave del JSON es el nombre de una hoja.</p>
        </section>

        <section class="au-card au-c-vars" aria-labelledby="au-vars-h">
          <h3 id="au-vars-h">${ico('i-braces')} Variables <span class="au-count" id="au-var-count"></span></h3>
          <div id="au-vars" class="au-vars"></div>
        </section>

        <section class="au-card au-c-usage" id="au-usage-card" aria-labelledby="au-usage-h" hidden>
          <h3 id="au-usage-h">${ico('i-activity')} Uso de esta plantilla</h3>
          <div class="au-usage" id="au-usage"></div>
        </section>

        <section class="au-card au-c-batch" aria-labelledby="au-batch-h">
          <h3 id="au-batch-h">${ico('i-package')} Lotes (varios archivos en un ZIP)</h3>
          <p class="au-help">¿Necesitas muchos archivos a la vez? Envía un arreglo <code>items</code> (1 a 200 cuerpos como el de la derecha)
            a <code>/batch</code> y recibe un <b>.zip</b> con un Excel por elemento. Usa <code>fileName</code> con <code>{{variables}}</code> para nombrar cada archivo.</p>
          <div class="au-mini-seg" role="group" aria-label="Lenguaje del ejemplo de lote">
            <button type="button" data-au="batch-tab" data-lang="curl" aria-pressed="true">cURL</button>
            <button type="button" data-au="batch-tab" data-lang="powershell" aria-pressed="false">PowerShell</button>
          </div>
          <div class="au-code-wrap"><pre class="au-code" id="au-batch-code"></pre>
            <button type="button" class="au-copy-btn" data-au="copy-batch">${ico('i-copy')}<span>Copiar</span></button></div>
        </section>
      </div>

      <div class="au-col">
        <section class="au-card au-c-opt" aria-labelledby="au-opt-h">
          <h3 id="au-opt-h">${ico('i-settings')} Opciones</h3>
          <div class="au-opts">
            <fieldset class="au-seg"><legend>Formato de respuesta</legend>
              <label><input type="radio" name="au-format" value="xlsx" checked> Archivo .xlsx</label>
              <label><input type="radio" name="au-format" value="base64"> JSON base64</label>
            </fieldset>
            <label class="au-switch"><input type="checkbox" id="au-strict"><span class="au-switch-ui" aria-hidden="true"></span>
              <span>Validación estricta <small>(rechaza datos inválidos con 422)</small></span></label>
            <label class="au-field"><span>Nombre de archivo</span>
              <input type="text" id="au-filename" maxlength="120" placeholder="p. ej. Factura {{cliente}}" spellcheck="false" autocomplete="off" aria-describedby="au-fn-preview"></label>
            <div class="au-chips" id="au-fn-chips" aria-label="Insertar variable en el nombre de archivo"></div>
            <p class="au-fn-preview" id="au-fn-preview" aria-live="polite"></p>
          </div>
        </section>

        <section class="au-card au-card-editor au-c-editor" aria-labelledby="au-body-h">
          <h3 id="au-body-h">${ico('i-code')} Cuerpo JSON de prueba</h3>
          <label class="au-sr" for="au-editor">Cuerpo JSON</label>
          <textarea id="au-editor" class="au-editor" spellcheck="false" autocomplete="off" autocapitalize="off"></textarea>
          <div class="au-json-state" id="au-json-state" aria-live="polite"></div>
          <div class="au-actions">
            <button type="button" class="au-btn" data-au="validate">${ico('i-check-circle')}<span>Validar</span></button>
            <button type="button" class="au-btn" data-au="format">${ico('i-sparkles')}<span>Formatear</span></button>
            <button type="button" class="au-btn" data-au="reset">${ico('i-refresh')}<span>Restablecer ejemplo</span></button>
            <button type="button" class="au-btn au-btn-primary" data-au="render">${ico('i-download')}<span id="au-render-label">Generar y descargar</span></button>
          </div>
          <div class="au-result" id="au-result" aria-live="polite"></div>
        </section>

        <section class="au-card au-c-snip" aria-labelledby="au-snip-h">
          <h3 id="au-snip-h">${ico('i-terminal')} Código listo para copiar</h3>
          <div class="au-tabs" role="tablist" aria-label="Lenguaje del ejemplo" id="au-tabs">
            ${TABS.map(([k, label, i]) => `<button type="button" role="tab" id="au-tab-${k}" data-tab="${k}" aria-controls="au-snip-panel" aria-selected="false" tabindex="-1">${ico(i)}<span>${label}</span></button>`).join('')}
          </div>
          <div role="tabpanel" id="au-snip-panel" class="au-snip-panel" tabindex="0">
            <p class="au-help" id="au-snip-note"></p>
            <div class="au-code-wrap"><pre class="au-code" id="au-snip-code"></pre>
              <button type="button" class="au-copy-btn" data-au="copy-snip">${ico('i-copy')}<span>Copiar</span></button></div>
            <div class="au-snip-extra" id="au-snip-extra"></div>
          </div>
        </section>
      </div>
    </div>
  </div>

  <footer class="au-foot">
    <a href="/api/openapi.json" target="_blank" rel="noopener" class="au-link">${ico('i-book')} OpenAPI (importar en Postman/Insomnia)</a>
    <span class="au-foot-note">${ico('i-info')} Consulta el README del proyecto para más ejemplos.</span>
    <button type="button" class="au-btn" data-au="close">Cerrar</button>
  </footer>
</div>`;
        document.body.appendChild(ov);
        S.root = ov;
        wire(ov);
        return ov;
    }

    const $ = id => document.getElementById(id);

    // ── Eventos (delegación; sin atributos inline) ───────────────────────────
    function wire(ov) {
        ov.addEventListener('mousedown', e => { if (e.target === ov) close(); });
        ov.addEventListener('click', e => {
            const b = e.target.closest('[data-au]');
            if (b) return onAction(b.dataset.au, b);
            const t = e.target.closest('[data-tab]');
            if (t) selectTab(t.dataset.tab, true);
        });
        $('au-tpl-select').addEventListener('change', e => { if (e.target.value) loadTemplate(e.target.value); });
        ov.querySelectorAll('input[name="au-format"]').forEach(r => r.addEventListener('change', () => { S.opts.format = r.value; refreshSnippets(); }));
        $('au-strict').addEventListener('change', e => { S.opts.strict = e.target.checked; refreshSnippets(); });
        $('au-filename').addEventListener('input', e => { S.opts.fileName = e.target.value; refreshSnippetsSoon(); });
        $('au-editor').addEventListener('input', () => { S.editorTouched = true; checkJson(); refreshSnippetsSoon(); });
    }

    const focusables = () => [...S.root.querySelectorAll('button:not([disabled]), [href], input:not([disabled]), select, textarea, [tabindex="0"]')]
        .filter(el => el.offsetParent !== null && !el.closest('[hidden]'));

    /**
     * Enrutador único de teclado, en fase de CAPTURA sobre window mientras el panel está abierto.
     * index.html escucha atajos del editor en captura sobre document (flechas, Supr, Tab, letras
     * mueven/borran celdas), así que solo deteniendo la propagación aquí —antes que nadie— el
     * editor de atrás queda aislado. stopPropagation no cancela la acción por defecto: escribir,
     * activar botones con Enter/Espacio o cambiar un <select> siguen funcionando.
     */
    function onWinKey(e) {
        if (!S.open) return;
        const t = e.target, inside = S.root.contains(t);
        // Otro diálogo abierto encima (p. ej. bienvenida de live.js): no interferir
        if (!inside && t && t.closest && t.closest('[role="dialog"], [aria-modal="true"], .modal-overlay:not(.hidden)')) return;
        e.stopPropagation();
        if (e.key === 'Escape') { e.preventDefault(); close(); return; }
        if (!inside) {
            // El foco quedó fuera (p. ej. en <body> tras un clic en zona no enfocable): Tab lo regresa
            if (e.key === 'Tab') { e.preventDefault(); const f = focusables(); (e.shiftKey ? f[f.length - 1] : f[0])?.focus(); }
            return;
        }
        // Tab inserta sangría dentro del editor JSON (Shift+Tab o Ctrl+Tab salen)
        if (t.id === 'au-editor' && e.key === 'Tab' && !e.shiftKey && !e.ctrlKey) {
            e.preventDefault();
            t.setRangeText('  ', t.selectionStart, t.selectionEnd, 'end');
            t.dispatchEvent(new Event('input'));
            return;
        }
        // Pestañas de snippets: flechas / Inicio / Fin (patrón WAI-ARIA tabs)
        if (t.getAttribute && t.getAttribute('role') === 'tab' && t.closest('#au-tabs')) {
            const keys = TABS.map(x => x[0]); let i = keys.indexOf(S.tab);
            if (e.key === 'ArrowRight') i = (i + 1) % keys.length;
            else if (e.key === 'ArrowLeft') i = (i - 1 + keys.length) % keys.length;
            else if (e.key === 'Home') i = 0;
            else if (e.key === 'End') i = keys.length - 1;
            else i = -1;
            if (i >= 0) { e.preventDefault(); selectTab(keys[i], true); return; }
        }
        // Trampa de foco básica
        if (e.key === 'Tab') {
            const f = focusables();
            if (!f.length) return;
            const first = f[0], last = f[f.length - 1];
            if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
            else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
        }
    }

    function onAction(act, btn) {
        switch (act) {
            case 'close': return close();
            case 'copy-url': return copyText(renderUrl(), btn);
            case 'copy-snip': return copyText($('au-snip-code').textContent, btn);
            case 'copy-batch': return copyText($('au-batch-code').textContent, btn);
            case 'copy-link': return copyText(directLink(), btn);
            case 'open-link': return window.open(directLink(), '_blank', 'noopener');
            case 'validate': return validate();
            case 'format': return formatJson();
            case 'reset': return resetExample();
            case 'render': return renderTest(btn);
            case 'dl-b64': return downloadB64(btn);
            case 'retry': return S.tplId ? loadTemplate(S.tplId) : loadList(null);
            case 'ins-var': return insertFileVar(btn.dataset.var || '');
            case 'batch-tab': S.batchTab = btn.dataset.lang === 'powershell' ? 'powershell' : 'curl'; return refreshSnippets();
        }
    }

    // ── Abrir / cerrar ────────────────────────────────────────────────────────
    function openAutomation(templateId) {
        build();
        if (!S.open) {
            S.lastFocus = document.activeElement;
            S.open = true;
            S.root.hidden = false;
            S.root.classList.remove('au-closing');
            void S.root.offsetWidth;            // reinicia la animación
            S.root.classList.add('au-open');
            document.documentElement.classList.add('au-lock');
            window.addEventListener('keydown', onWinKey, true);
        }
        const id = templateId && /^[A-Za-z0-9_-]{1,64}$/.test(String(templateId)) ? String(templateId) : null;
        loadList(id);
        if (id) loadTemplate(id);
        else if (S.tplId && S.sheets.length) showEmpty(null);   // reabrir sin id: conserva la última plantilla vista
        else showEmpty('pick');
        setTimeout(() => { (id ? $('au-editor') : $('au-tpl-select'))?.focus({ preventScroll: true }); }, 30);
        if (!S.unsubUsage && typeof ETB.on === 'function') {
            // Actualiza el mini-panel de uso en vivo cuando llega un evento SSE de esta plantilla
            S.unsubUsage = ETB.on('live:usage', d => {
                const ev = d && (d.detail || d);
                if (S.open && ev && ev.event && ev.event.templateId === S.tplId) loadUsage(S.tplId);
            });
        }
    }

    function close() {
        if (!S.open) return;
        S.open = false;
        document.documentElement.classList.remove('au-lock');
        window.removeEventListener('keydown', onWinKey, true);
        const done = () => { if (!S.open) { S.root.hidden = true; S.root.classList.remove('au-closing'); } };
        S.root.classList.remove('au-open');
        if (reduceMotion()) done();
        else { S.root.classList.add('au-closing'); setTimeout(done, 220); }
        try { if (S.lastFocus && document.contains(S.lastFocus)) S.lastFocus.focus({ preventScroll: true }); } catch { /* ignorar */ }
    }

    // ── Carga de datos ────────────────────────────────────────────────────────
    async function loadList(selectId) {
        const sel = $('au-tpl-select');
        const r = await req('GET', '/api/templates');
        if (!r.ok || !Array.isArray(r.data)) {
            sel.innerHTML = '<option value="">No se pudieron cargar las plantillas</option>';
            if (!S.tplId) showEmpty('error', (r.data && r.data.error) || 'No se pudo obtener la lista de plantillas.');
            return;
        }
        S.templates = r.data;
        const cur = selectId || S.tplId || '';
        sel.innerHTML = `<option value=""${cur ? '' : ' selected'} disabled>Elige una plantilla…</option>` +
            r.data.map(t => `<option value="${esc(t.id)}"${t.id === cur ? ' selected' : ''}>${esc(t.name || 'Sin nombre')}</option>`).join('');
        if (!r.data.length && !S.tplId) showEmpty('none');
    }

    function showEmpty(kind, msg) {
        const e = $('au-empty'), m = $('au-main');
        if (!kind) { e.hidden = true; m.hidden = false; return; }
        m.hidden = true; e.hidden = false;
        S.tplId = null; $('au-tname').textContent = ''; $('au-rev').hidden = true;
        if (kind === 'pick') e.innerHTML = `${ico('i-file-spreadsheet')}<h3>Elige una plantilla</h3><p>Selecciona arriba la plantilla que quieres generar automáticamente.</p>`;
        else if (kind === 'none') e.innerHTML = `${ico('i-folder')}<h3>Aún no hay plantillas</h3><p>Crea y guarda una plantilla con variables <code>{{nombre}}</code> para automatizarla.</p>`;
        else e.innerHTML = `${ico('i-alert-triangle')}<h3>Algo salió mal</h3><p>${esc(msg || 'Error desconocido')}</p><button type="button" class="au-btn" data-au="retry">${ico('i-refresh')}<span>Reintentar</span></button>`;
    }

    const colName = n => { let r = ''; while (n > 0) { const m = (n - 1) % 26; r = String.fromCharCode(65 + m) + r; n = Math.floor((n - 1) / 26); } return r; };
    const colNum = l => String(l).toUpperCase().split('').reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0);
    /** "filas 5–7" / "columnas B–D" */
    function zoneRange(z) {
        if (z.direction === 'cols') return z.start === z.end ? `columna ${colName(z.start)}` : `columnas ${colName(z.start)}–${colName(z.end)}`;
        return z.start === z.end ? `fila ${z.start}` : `filas ${z.start}–${z.end}`;
    }
    /** Zona (de la lista) que contiene la celda; filas primero. */
    function zoneOfAddr(zones, addr) {
        const m = /^([A-Z]+)(\d+)$/i.exec(addr || ''); if (!m) return null;
        const c = colNum(m[1]), r = +m[2];
        return zones.find(z => z.direction === 'rows' && r >= z.start && r <= z.end) || zones.find(z => z.direction === 'cols' && c >= z.start && c <= z.end) || null;
    }

    /** Normaliza /schema (formato actual `schema` o buildSchema `sheets`). Entiende repeatZones y el rowTemplate heredado. */
    function normalizeSchema(d) {
        const list = Array.isArray(d && d.schema) ? d.schema : Array.isArray(d && d.sheets) ? d.sheets : [];
        return list.map(s => {
            const rt = s.rowTemplate && s.rowTemplate.arrayVariable ? s.rowTemplate : null;
            let zones = Array.isArray(s.repeatZones)
                ? s.repeatZones.filter(z => z && z.arrayVariable).map(z => ({ direction: z.direction === 'cols' ? 'cols' : 'rows', start: +z.start, end: +z.end, arrayVariable: String(z.arrayVariable) }))
                : (rt ? [{ direction: 'rows', start: +rt.startRow, end: +rt.endRow, arrayVariable: String(rt.arrayVariable) }] : []);
            return {
                name: String(s.name ?? ''), isRepeatable: !!s.isRepeatable, repeatVariable: s.repeatVariable || '',
                rowTemplate: rt, zones,
                variables: (s.variables || []).map(v => {
                    let zone = v.zone && v.zone.arrayVariable ? zones.find(z => z.arrayVariable === v.zone.arrayVariable) || null : null;
                    if (!zone && v.zone === undefined && v.addr) zone = zoneOfAddr(zones, v.addr);
                    if (!zone && v.zone === undefined && v.inRowTemplate && zones.length) zone = zones[0];
                    return { name: String(v.name ?? ''), type: v.type || 'text', required: !!v.required, hint: v.hint || '', addr: v.addr || '', defaultValue: v.defaultValue, inRowTemplate: !!zone, zone, implicit: !!v.implicit };
                }),
            };
        });
    }

    function sampleForType(t, name) {
        switch (t) {
            case 'number': return 100;
            case 'boolean': return true;
            case 'date': return new Date().toISOString().slice(0, 10);
            case 'image': return 'data:image/png;base64,iVBORw0KGgo=';
            default: return name ? `Ejemplo ${name}` : 'texto';
        }
    }

    /** Ejemplo construido en el cliente (fallback si /example no existe). */
    function buildExampleClient(sheets) {
        const ex = {};
        sheets.forEach(s => {
            const obj = {};
            s.variables.forEach(v => { if (!v.zone) obj[v.name] = sampleForType(v.type, v.name); });
            s.zones.forEach(z => {
                const item = {};
                s.variables.forEach(v => { if (v.zone && v.zone.arrayVariable === z.arrayVariable) item[v.name] = sampleForType(v.type, v.name); });
                obj[z.arrayVariable] = [item, { ...item }];
            });
            ex[s.name] = s.isRepeatable ? [obj, JSON.parse(JSON.stringify(obj))] : obj;
        });
        return ex;
    }

    async function loadTemplate(id) {
        const seq = ++S.loadSeq;
        if (id !== S.loadedId) {
            // Otra plantilla: vuelve arriba y limpia el patrón de nombre (sus {{variables}} no aplicarían)
            $('au-body').scrollTop = 0;
            if (S.loadedId) { S.opts.fileName = ''; $('au-filename').value = ''; }
        }
        S.tplId = id; showEmpty(null);
        const sel = $('au-tpl-select'); if (sel.value !== id) sel.value = id;
        $('au-vars').innerHTML = '<div class="au-skel"></div><div class="au-skel"></div><div class="au-skel"></div>';
        $('au-result').innerHTML = ''; $('au-usage-card').hidden = true;
        $('au-url').textContent = renderUrl();
        const known = S.templates.find(t => t.id === id);
        $('au-tname').textContent = known ? `· ${known.name || ''}` : '';

        const [sc, exr] = await Promise.all([req('GET', `/api/templates/${encodeURIComponent(id)}/schema`), req('GET', `/api/templates/${encodeURIComponent(id)}/example`)]);
        if (seq !== S.loadSeq) return;           // el usuario cambió de plantilla mientras cargaba
        if (!sc.ok || !sc.data) {
            showEmpty('error', sc.status === 404 ? 'La plantilla no existe o aún no se ha guardado.' : ((sc.data && sc.data.error) || `Error HTTP ${sc.status}`));
            S.tplId = id; return;
        }
        S.name = sc.data.name || (known && known.name) || 'Sin nombre';
        S.revision = typeof sc.data.revision === 'number' ? sc.data.revision : (known && typeof known.revision === 'number' ? known.revision : null);
        S.sheets = normalizeSchema(sc.data);
        const exampleOk = exr.ok && exr.data && typeof exr.data === 'object' && !Array.isArray(exr.data);
        S.example = exampleOk ? exr.data : buildExampleClient(S.sheets);
        if (exampleOk) mergeImplicitVars(S.sheets, S.example);

        $('au-tname').textContent = `· ${S.name}`;
        const rev = $('au-rev'); rev.hidden = !S.revision; rev.textContent = `rev. ${S.revision}`;   // 0/null = plantilla antigua sin revisión
        renderVars();
        // Al reabrir la misma plantilla se conserva el cuerpo que el usuario editó
        if (S.loadedId === id && S.editorTouched) { checkJson(); refreshSnippets(); }
        else resetExample(true);
        S.loadedId = id;
        loadUsage(id);
    }

    /** Infiere el tipo de un valor de ejemplo (para variables implícitas). */
    function typeOfValue(v) {
        if (typeof v === 'number') return 'number';
        if (typeof v === 'boolean') return 'boolean';
        if (typeof v === 'string' && /^data:image\//.test(v)) return 'image';
        if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v)) return 'date';
        return 'text';
    }

    /**
     * /schema (formato heredado) solo lista variables explícitas; /example (buildExample) también
     * incluye marcadores {{x}} escritos como texto o en fórmulas. Las agrega como "implícitas".
     */
    function mergeImplicitVars(sheets, example) {
        sheets.forEach(s => {
            let data = example[s.name];
            if (Array.isArray(data)) data = data[0];
            if (!data || typeof data !== 'object' || Array.isArray(data)) return;
            const have = new Set(s.variables.map(v => v.name + '|' + (v.zone ? v.zone.arrayVariable : '')));
            const add = (name, val, zone) => {
                const key = name + '|' + (zone ? zone.arrayVariable : '');
                if (have.has(key)) return;
                have.add(key);
                s.variables.push({ name, type: typeOfValue(val), required: false, hint: '', addr: '', defaultValue: '', inRowTemplate: !!zone, zone, implicit: true });
            };
            Object.keys(data).forEach(k => {
                const zone = s.zones.find(z => z.arrayVariable === k);
                if (zone && Array.isArray(data[k])) {
                    const item = data[k][0];
                    if (item && typeof item === 'object') Object.keys(item).forEach(rk => add(rk, item[rk], zone));
                } else if (data[k] === null || typeof data[k] !== 'object') add(k, data[k], null);
            });
        });
    }

    /** Variables simples (no de filas, no imagen) útiles para el nombre de archivo y el enlace directo. */
    function flatVarNames() {
        const out = [];
        S.sheets.forEach(s => s.variables.forEach(v => {
            if (!v.inRowTemplate && v.type !== 'image' && !out.includes(v.name)) out.push(v.name);
        }));
        return out;
    }

    function renderFileChips() {
        const names = flatVarNames().slice(0, 12);
        $('au-fn-chips').innerHTML = names.length
            ? `<span class="au-chips-l">Insertar:</span>` + names.map(n => `<button type="button" class="au-chip" data-au="ins-var" data-var="${esc(n)}">{{${esc(n)}}}</button>`).join('')
            : '';
    }

    function insertFileVar(name) {
        if (!/^\w+$/.test(name)) return;
        const inp = $('au-filename'), tok = `{{${name}}}`;
        const s = inp.selectionStart ?? inp.value.length, e = inp.selectionEnd ?? s;
        const pre = inp.value.slice(0, s), sep = pre && !/\s$/.test(pre) ? ' ' : '';
        inp.setRangeText(sep + tok, s, e, 'end');
        inp.focus();
        inp.dispatchEvent(new Event('input'));
    }

    function renderVars() {
        renderFileChips();
        const total = S.sheets.reduce((n, s) => n + s.variables.length, 0);
        $('au-var-count').textContent = String(total);
        if (!S.sheets.length) { $('au-vars').innerHTML = '<p class="au-help">Esta plantilla no tiene hojas.</p>'; return; }
        $('au-vars').innerHTML = S.sheets.map(s => {
            const badges = (s.isRepeatable ? `<span class="au-badge au-badge-rep" title="La clave de esta hoja es un arreglo: se crea una hoja por elemento">${ico('i-repeat')}repetible</span>` : '') +
                s.zones.map(z => {
                    const cols = z.direction === 'cols';
                    return `<span class="au-badge au-badge-rows${cols ? ' au-badge-cols' : ''}" title="${esc(zoneRange(z))}: se repiten por cada elemento de &quot;${esc(z.arrayVariable)}&quot;">${ico(cols ? 'i-columns' : 'i-rows')}${cols ? 'columnas repetibles' : 'filas repetibles'}: ${esc(z.arrayVariable)}</span>`;
                }).join('');
            // Orden: variables simples primero, luego las de cada zona
            const vars = s.variables.filter(v => !v.zone).concat(...s.zones.map(z => s.variables.filter(v => v.zone && v.zone.arrayVariable === z.arrayVariable)));
            const rows = vars.map(v => `<tr>
                <td><code class="au-var">${esc(v.zone ? `${v.zone.arrayVariable}[].${v.name}` : v.name)}</code>${v.implicit ? '<span class="au-badge au-badge-mini au-badge-imp" title="Detectada como {{marcador}} en el texto de una celda; no tiene tipo configurado">implícita</span>' : ''}</td>
                <td><span class="au-type au-type-${esc(TYPE_LABEL[v.type] ? v.type : 'text')}">${esc(TYPE_LABEL[v.type] || v.type)}</span></td>
                <td>${v.required ? `<span class="au-req">${ico('i-dot')}Sí</span>` : '<span class="au-opt">No</span>'}</td>
                <td class="au-hint">${v.hint ? esc(v.hint) : '<span class="au-opt">—</span>'}</td>
                <td><span class="au-addr">${esc(v.addr || '—')}</span></td></tr>`).join('');
            // Cómo se ve esta hoja dentro del JSON (ayuda para no expertos)
            const shape = s.isRepeatable
                ? `<code>"${esc(s.name)}": [ { … }, { … } ]</code> <span>una hoja por elemento</span>`
                : `<code>"${esc(s.name)}": { … }</code>`;
            const rowsInfo = s.zones.map(z => ` · <code>"${esc(z.arrayVariable)}": [ … ]</code> <span>repite ${esc(zoneRange(z))} por elemento</span>`).join('');
            return `<div class="au-sheet">
              <div class="au-sheet-h">${ico('i-file-spreadsheet')}<b>${esc(s.name)}</b>${badges}</div>
              <div class="au-shape">${ico('i-braces')}<span>${shape}${rowsInfo}</span></div>
              ${rows ? `<div class="au-table-wrap"><table class="au-table"><thead><tr><th scope="col">Variable</th><th scope="col">Tipo</th><th scope="col">Requerida</th><th scope="col">Descripción</th><th scope="col">Celda</th></tr></thead><tbody>${rows}</tbody></table></div>`
                : '<p class="au-help">Sin variables en esta hoja.</p>'}
            </div>`;
        }).join('');
    }

    async function loadUsage(id) {
        const card = $('au-usage-card');
        const r = await req('GET', `/api/usage/${encodeURIComponent(id)}`);
        if (id !== S.tplId) return;
        if (!r.ok || !r.data || !r.data.totals) { card.hidden = true; return; }   // endpoint ausente → ocultar
        const d = r.data, today = d.today || {}, tot = d.totals || {}, last = d.last || {};
        card.hidden = false;
        const stat = (icon, label, val, cls, title) => `<div class="au-stat ${cls || ''}"${title ? ` title="${esc(title)}"` : ''}>${ico(icon)}<span class="au-stat-v">${esc(val)}</span><span class="au-stat-l">${esc(label)}</span></div>`;
        $('au-usage').innerHTML =
            stat('i-calendar', 'Generados hoy', today.renders ?? 0) +
            stat('i-chart-bar', 'Total generados', tot.renders ?? 0) +
            stat('i-clock', 'Último uso', fmtRel(last.renderAt), 'au-stat-text', last.renderAt ? fmtDate(last.renderAt) : '') +
            stat('i-alert-circle', 'Errores', tot.errors ?? 0, (tot.errors ? 'au-stat-bad' : '')) +
            sparkHtml(d.spark24h);
    }

    /** Mini gráfica de barras (24 h) en SVG; vacía si no hay datos. */
    function sparkHtml(arr) {
        if (!Array.isArray(arr) || !arr.length) return '';
        const vals = arr.slice(-24).map(n => Math.max(0, Number(n) || 0)), max = Math.max(1, ...vals);
        const w = 100 / vals.length;
        const bars = vals.map((v, i) => {
            const h = v ? Math.max(4, (v / max) * 100) : 2;
            return `<rect x="${(i * w + w * 0.15).toFixed(2)}" y="${(100 - h).toFixed(2)}" width="${(w * 0.7).toFixed(2)}" height="${h.toFixed(2)}" rx="1"${i === vals.length - 1 ? ' class="now"' : ''}><title>${vals.length - 1 - i === 0 ? 'Esta hora' : `Hace ${vals.length - 1 - i} h`}: ${v}</title></rect>`;
        }).join('');
        const total = vals.reduce((a, b) => a + b, 0);
        return `<div class="au-spark"><div class="au-spark-l">${ico('i-trending-up')}Últimas 24 h <b>${esc(total)}</b></div>
            <svg viewBox="0 0 100 100" preserveAspectRatio="none" role="img" aria-label="Archivos generados por hora en las últimas 24 horas">${bars}</svg></div>`;
    }

    // ── Editor JSON ───────────────────────────────────────────────────────────
    function parseBody() {
        try { return { ok: true, value: JSON.parse($('au-editor').value) }; }
        catch (e) { return { ok: false, error: e.message }; }
    }
    function checkJson() {
        const p = parseBody(), st = $('au-json-state');
        st.className = 'au-json-state ' + (p.ok ? 'ok' : 'bad');
        st.innerHTML = p.ok ? `${ico('i-check')}JSON válido` : `${ico('i-alert-circle')}${esc('JSON inválido: ' + p.error)}`;
        return p;
    }
    function formatJson() {
        const p = checkJson();
        if (!p.ok) { toast('Corrige el JSON antes de formatear', 'warn'); return; }
        $('au-editor').value = JSON.stringify(p.value, null, 2);
        refreshSnippets();
    }
    function resetExample(silent) {
        $('au-editor').value = JSON.stringify(S.example || {}, null, 2);
        S.editorTouched = false;
        checkJson(); refreshSnippets();
        if (silent !== true) { $('au-result').innerHTML = ''; toast('Ejemplo restablecido', 'info'); }
    }

    // ── Validar ───────────────────────────────────────────────────────────────
    function issuesHtml(errors, warnings) {
        const li = (cls, icon, it) => {
            const where = it.path || it.sheet || '';
            // El servidor suele repetir la ruta al final del mensaje: "(Factura.cliente)"
            const msg = String(it.message || it.code || '').replace(where ? ` (${where})` : /^$/, '');
            const item = typeof it.item === 'number' ? `<span class="au-issue-item">elemento ${it.item + 1}</span> ` : '';
            return `<li class="${cls}">${ico(icon)}<span>${item}${where ? `<code>${esc(where)}</code> ` : ''}${esc(msg)}</span></li>`;
        };
        return `<ul class="au-issues">${(errors || []).map(e => li('err', 'i-alert-circle', e)).join('')}${(warnings || []).map(w => li('warn', 'i-alert-triangle', w)).join('')}</ul>`;
    }
    function setResult(kind, title, extraHtml) {
        const icon = kind === 'ok' ? 'i-check-circle' : kind === 'warn' ? 'i-alert-triangle' : kind === 'info' ? 'i-info' : 'i-alert-circle';
        $('au-result').innerHTML = `<div class="au-res au-res-${kind}"><div class="au-res-h">${ico(icon)}<b>${esc(title)}</b></div>${extraHtml || ''}</div>`;
    }

    async function validate() {
        if (!S.tplId) return;
        const p = checkJson();
        if (!p.ok) { setResult('err', 'El JSON no es válido', `<p>${esc(p.error)}</p>`); return; }
        setResult('info', 'Validando…');
        const r = await req('POST', `/api/templates/${encodeURIComponent(S.tplId)}/validate`, p.value);
        if (r.status === 404 && !(r.data && 'valid' in r.data)) {
            setResult('info', 'El servidor no ofrece validación detallada', '<p>La sintaxis JSON es correcta. Prueba con «Generar y descargar» para confirmar.</p>'); return;
        }
        if (!r.ok || !r.data) { setResult('err', 'No se pudo validar', `<p>${esc((r.data && r.data.error) || `Error HTTP ${r.status}`)}</p>`); return; }
        const { valid, errors = [], warnings = [] } = r.data;
        if (valid && !warnings.length) setResult('ok', 'Todo correcto: los datos coinciden con la plantilla');
        else if (valid) setResult('warn', `Válido con ${warnings.length} advertencia(s)`, issuesHtml([], warnings));
        else setResult('err', `${errors.length} error(es) encontrado(s)`, issuesHtml(errors, warnings));
    }

    // ── Generar y descargar (prueba real) ────────────────────────────────────
    /** Query de opciones. fillBody: rellena el patrón {{…}} del nombre con ese cuerpo (prueba real),
     *  así el nombre es correcto aunque el servidor no aplique el patrón en renders individuales. */
    function queryString(extra, fillBody) {
        const q = new URLSearchParams();
        if (S.opts.format === 'base64') q.set('format', 'base64');
        if (S.opts.strict) q.set('strict', '1');
        const fn = S.opts.fileName.trim();
        if (fn) q.set('filename', fillBody ? (fillPattern(fn, fillBody) || fn) : fn);
        if (extra) Object.entries(extra).forEach(([k, v]) => q.append(k, v));
        const s = q.toString();
        return s ? '?' + s : '';
    }
    const renderPath = () => `/api/render/${encodeURIComponent(S.tplId || '')}`;
    const renderUrl = () => origin() + renderPath();
    const fullUrl = () => renderUrl() + queryString();

    function fileNameFromDisposition(cd) {
        if (!cd) return null;
        const star = /filename\*\s*=\s*(?:UTF-8|utf-8)''([^;]+)/.exec(cd);
        if (star) { try { return decodeURIComponent(star[1].trim().replace(/^"|"$/g, '')); } catch { /* sigue */ } }
        const plain = /filename\s*=\s*"?([^";]+)"?/i.exec(cd);
        return plain ? plain[1].trim() : null;
    }
    function saveBlob(blob, name) {
        const url = URL.createObjectURL(blob), a = document.createElement('a');
        a.href = url; a.download = name || 'resultado.xlsx'; a.rel = 'noopener';
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 4000);
    }

    let lastB64 = null;
    async function renderTest(btn) {
        if (!S.tplId) return;
        const p = checkJson();
        if (!p.ok) { setResult('err', 'El JSON no es válido', `<p>${esc(p.error)}</p>`); return; }
        btn.disabled = true; btn.classList.add('au-busy');
        const use = btn.querySelector('use'), prevIco = use && use.getAttribute('href');
        if (use) use.setAttribute('href', '#i-refresh');            // ícono giratorio mientras genera
        const t0 = performance.now();
        try {
            let res;
            try {
                res = await fetch(renderPath() + queryString(null, p.value), { method: 'POST', headers: { 'Content-Type': 'application/json', ...identityHeaders() }, body: JSON.stringify(p.value) });
            } catch { throw new Error('No se pudo conectar con el servidor.'); }
            const ms = res.headers.get('X-Render-Ms') || Math.round(performance.now() - t0);
            if (!res.ok) {
                const d = await res.json().catch(() => null);
                if (res.status === 422 && d) { setResult('err', 'Rechazado por validación estricta', issuesHtml(d.errors, d.warnings)); return; }
                throw new Error((d && d.error) || `Error HTTP ${res.status}`);
            }
            const ctype = res.headers.get('Content-Type') || '';
            if (S.opts.format === 'base64' || ctype.includes('application/json')) {
                const d = await res.json();
                lastB64 = d;
                const preview = JSON.stringify({ ...d, base64: String(d.base64 || '').slice(0, 60) + '…' }, null, 2);
                setResult('ok', `Respuesta JSON recibida en ${ms} ms`,
                    `<p class="au-meta">${ico('i-file-spreadsheet')}<b>${esc(d.fileName || '')}</b> · ${esc(fmtBytes(d.size))}</p>
                     <pre class="au-code au-code-sm">${esc(preview)}</pre>
                     <button type="button" class="au-btn" data-au="dl-b64">${ico('i-download')}<span>Descargar archivo decodificado</span></button>`);
            } else {
                const blob = await res.blob();
                const name = fileNameFromDisposition(res.headers.get('Content-Disposition')) || outName();
                saveBlob(blob, name);
                setResult('ok', 'Archivo generado y descargado', `<p class="au-meta">${ico('i-file-spreadsheet')}<b>${esc(name)}</b> · ${esc(fmtBytes(blob.size))} · ${ico('i-gauge')}${esc(ms)} ms</p>`);
            }
            toast('Excel generado', 'success');
            setTimeout(() => S.tplId && loadUsage(S.tplId), 600);
        } catch (e) {
            setResult('err', 'No se pudo generar el archivo', `<p>${esc(e.message)}</p>`);
        } finally {
            btn.disabled = false; btn.classList.remove('au-busy');
            if (use && prevIco) use.setAttribute('href', prevIco);
        }
    }
    function downloadB64() {
        if (!lastB64 || !lastB64.base64) return;
        try {
            const bin = atob(lastB64.base64), u8 = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
            saveBlob(new Blob([u8], { type: lastB64.mimeType || 'application/octet-stream' }), lastB64.fileName || 'resultado.xlsx');
        } catch { toast('El base64 recibido no es válido', 'error'); }
    }

    // ── Snippets ──────────────────────────────────────────────────────────────
    const shQ = s => `'${String(s).replace(/'/g, `'\\''`)}'`;          // bash: comillas simples
    const psQ = s => `'${String(s).replace(/'/g, "''")}'`;              // PowerShell: comillas simples
    const csVerb = s => `@"${String(s).replace(/"/g, '""')}"`;          // C#: cadena literal @"..."
    const jsStr = s => JSON.stringify(String(s));                       // JS/Python: literal JSON (válido en ambos)

    /** Convierte un valor JSON a literal Python legible. */
    function toPy(v, ind) {
        ind = ind || '';
        const nx = ind + '    ';
        if (v === null) return 'None';
        if (v === true) return 'True';
        if (v === false) return 'False';
        if (typeof v === 'number') return isFinite(v) ? String(v) : 'None';
        if (typeof v === 'string') return jsStr(v).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
        if (Array.isArray(v)) return v.length ? `[\n${v.map(x => nx + toPy(x, nx)).join(',\n')}\n${ind}]` : '[]';
        const ks = Object.keys(v);
        return ks.length ? `{\n${ks.map(k => `${nx}${jsStr(k)}: ${toPy(v[k], nx)}`).join(',\n')}\n${ind}}` : '{}';
    }

    /** Primer valor primitivo por clave (búsqueda en anchura), igual que flatten() del servidor. */
    function flattenBody(body) {
        const flat = new Map(), queue = [body]; let steps = 0;
        while (queue.length && steps++ < 5000) {
            const cur = queue.shift();
            if (!cur || typeof cur !== 'object') continue;
            if (Array.isArray(cur)) { cur.forEach(x => queue.push(x)); continue; }
            Object.keys(cur).forEach(k => {
                const v = cur[k];
                if (v !== null && typeof v === 'object') queue.push(v);
                else if (!flat.has(k) && v !== undefined && v !== null && v !== '') flat.set(k, v);
            });
        }
        return flat;
    }
    /** 'Factura {{cliente}}' → 'Factura Ana' (como fillPattern del servidor). */
    function fillPattern(pattern, body) {
        const flat = flattenBody(body);
        return String(pattern || '').replace(/\{\{(\w+)\}\}/g, (_, k) => (flat.has(k) ? String(flat.get(k)) : '')).replace(/\s+/g, ' ').trim();
    }
    /** Nombre seguro (.xlsx) como safeFileName del servidor (versión simplificada). */
    function safeName(name, fallback) {
        const clean = s => String(s || '').replace(/[\u0000-\u001f\u007f]/g, '').replace(/[/\\:*?"<>|]/g, '_')
            .replace(/\s+/g, ' ').trim().replace(/\.xlsx$/i, '').replace(/^[.\s]+|[.\s]+$/g, '').slice(0, 115);
        return (clean(name) || clean(fallback) || 'resultado') + '.xlsx';
    }
    /** Nombre del archivo que se obtendrá con el cuerpo actual. */
    function outName() { return safeName(fillPattern(S.opts.fileName.trim(), currentBody()) || S.name, 'resultado'); }

    function currentBody() { const p = parseBody(); return p.ok ? p.value : (S.example || {}); }

    function snippet(kind) {
        const body = currentBody(), pretty = JSON.stringify(body, null, 2);
        const url = fullUrl(), b64 = S.opts.format === 'base64', file = outName();
        switch (kind) {
            case 'curl': return {
                // El JSON va por stdin con un heredoc literal ('JSON'): sin escapes de comillas y sin que
                // Windows altere acentos al pasarlo como argumento.
                note: 'Bash / macOS / Linux / Git Bash. Para Windows CMD usa la pestaña PowerShell.',
                code: b64
                    ? `curl -sS -f -X POST ${shQ(url)} \\\n  -H 'Content-Type: application/json' \\\n  --data-binary @- <<'JSON' \\\n  | python3 -c 'import sys,json,base64; d=json.load(sys.stdin.buffer); open(d["fileName"],"wb").write(base64.b64decode(d["base64"])); print("Guardado:", d["fileName"])'\n${pretty}\nJSON`
                    : `curl -sS -f -X POST ${shQ(url)} \\\n  -H 'Content-Type: application/json' \\\n  --data-binary @- \\\n  -o ${shQ(file)} <<'JSON'\n${pretty}\nJSON`,
            };
            case 'powershell': return {
                note: 'Windows PowerShell 5.1 o PowerShell 7+.',
                code: `$body = @'\n${pretty}\n'@\n` + (b64
                    ? `$r = Invoke-RestMethod -Method Post -Uri ${psQ(url)} \`\n  -ContentType 'application/json; charset=utf-8' \`\n  -Body ([System.Text.Encoding]::UTF8.GetBytes($body))\n[System.IO.File]::WriteAllBytes((Join-Path (Get-Location) $r.fileName), [Convert]::FromBase64String($r.base64))`
                    : `Invoke-WebRequest -Method Post -Uri ${psQ(url)} \`\n  -ContentType 'application/json; charset=utf-8' \`\n  -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) \`\n  -OutFile ${psQ(file)}`),
            };
            case 'python': return {
                note: 'Requiere la librería requests (pip install requests).',
                code: `import ${b64 ? 'base64\nimport ' : ''}requests\n\nurl = ${jsStr(url)}\npayload = ${toPy(body)}\n\nr = requests.post(url, json=payload, timeout=60)\nr.raise_for_status()\n` + (b64
                    ? `data = r.json()\nwith open(data["fileName"], "wb") as f:\n    f.write(base64.b64decode(data["base64"]))\nprint("Guardado:", data["fileName"])\n`
                    : `with open(${jsStr(file)}, "wb") as f:\n    f.write(r.content)\nprint("Guardado:", ${jsStr(file)})\n`),
            };
            case 'node': return {
                note: 'Node.js 18 o superior (fetch incluido). Guarda como generar.mjs y ejecuta: node generar.mjs',
                code: `import { writeFile } from 'node:fs/promises';\n\nconst url = ${jsStr(url)};\nconst payload = ${pretty};\n\nconst res = await fetch(url, {\n  method: 'POST',\n  headers: { 'Content-Type': 'application/json' },\n  body: JSON.stringify(payload),\n});\nif (!res.ok) throw new Error(\`HTTP \${res.status}: \${await res.text()}\`);\n` + (b64
                    ? `const data = await res.json();\nawait writeFile(data.fileName, Buffer.from(data.base64, 'base64'));\nconsole.log('Guardado:', data.fileName);\n`
                    : `await writeFile(${jsStr(file)}, Buffer.from(await res.arrayBuffer()));\nconsole.log('Guardado:', ${jsStr(file)});\n`),
            };
            case 'csharp': return {
                note: '.NET 6+ (top-level statements). Usa HttpClient y System.Text.Json.',
                code: `using System.Net.Http;\nusing System.Text;\n${b64 ? 'using System.Text.Json;\n' : ''}\nvar url = ${jsStr(url)};\nvar json = ${csVerb(pretty)};\n\nusing var http = new HttpClient();\nusing var content = new StringContent(json, Encoding.UTF8, "application/json");\nvar res = await http.PostAsync(url, content);\nres.EnsureSuccessStatusCode();\n` + (b64
                    ? `using var doc = JsonDocument.Parse(await res.Content.ReadAsStringAsync());\nvar fileName = doc.RootElement.GetProperty("fileName").GetString()!;\nawait File.WriteAllBytesAsync(fileName, Convert.FromBase64String(doc.RootElement.GetProperty("base64").GetString()!));\nConsole.WriteLine($"Guardado: {fileName}");\n`
                    : `await File.WriteAllBytesAsync(${jsStr(file)}, await res.Content.ReadAsByteArrayAsync());\nConsole.WriteLine("Guardado: ${file.replace(/["\\{}]/g, '')}");\n`),
            };
            case 'pa': {
                const pUrl = (() => { const q = new URLSearchParams(queryString().slice(1)); q.set('format', 'base64'); return renderUrl() + '?' + q.toString(); })();
                return {
                    note: 'Power Automate: acción «HTTP» + «Crear archivo» (OneDrive/SharePoint). Siempre usa formato base64; escribe las expresiones en la pestaña «Expresión» (fx).',
                    code: `Acción: HTTP\n  Método:   POST\n  URI:      ${pUrl}\n  Encabezados:\n    Content-Type: application/json\n  Cuerpo:\n${pretty.replace(/^/gm, '    ')}\n\nAcción: Crear archivo (OneDrive para la Empresa / SharePoint)\n  Nombre de archivo:    body('HTTP')?['fileName']\n  Contenido de archivo: base64ToBinary(body('HTTP')?['base64'])\n\n# Nota: no envuelvas base64ToBinary en @{…} (lo convertiría a texto y dañaría el archivo).`,
                };
            }
            case 'n8n': {
                const node = {
                    nodes: [{
                        parameters: {
                            method: 'POST', url, sendBody: true, specifyBody: 'json', jsonBody: pretty,
                            options: b64 ? {} : { response: { response: { responseFormat: 'file', outputPropertyName: 'data' } } },
                        },
                        name: 'Generar Excel', type: 'n8n-nodes-base.httpRequest', typeVersion: 4.2, position: [0, 0],
                    }],
                    connections: {},
                };
                return { note: 'Copia y pega (Ctrl+V) directamente en el lienzo de n8n. ' + (b64 ? 'La respuesta trae fileName y base64.' : 'El archivo queda en la propiedad binaria «data».'), code: JSON.stringify(node, null, 2) };
            }
            case 'link': {
                const link = directLink();
                return {
                    note: 'Abre o comparte esta URL (GET) para descargar el Excel sin programar. Solo incluye variables simples de hojas no repetibles.',
                    code: link || 'Esta plantilla no tiene variables simples para usar en un enlace.',
                    link: !!link,
                };
            }
        }
        return { note: '', code: '' };
    }

    /** GET /api/render/:id?var=… con las variables planas de hojas no repetibles. */
    function directLink() {
        if (!S.tplId) return '';
        const body = currentBody(), q = new URLSearchParams();
        S.sheets.forEach(s => {
            if (s.isRepeatable) return;
            const data = body && typeof body[s.name] === 'object' && !Array.isArray(body[s.name]) ? body[s.name] : {};
            s.variables.forEach(v => {
                if (v.inRowTemplate || v.type === 'image') return;
                const val = data[v.name];
                if (val == null || typeof val === 'object') return;
                if (!q.has(v.name)) q.set(v.name, String(val));
            });
        });
        if (S.opts.fileName.trim()) q.set('filename', S.opts.fileName.trim());
        if (S.opts.format === 'base64') q.set('format', 'base64');
        if (S.opts.strict) q.set('strict', '1');
        const s = q.toString();
        return renderUrl() + (s ? '?' + s : '');
    }

    /** Snippet de lote: POST /api/render/:id/batch → ZIP con un Excel por elemento de items. */
    function batchSnippet() {
        const body = currentBody(), items = [body, body];
        const pattern = S.opts.fileName.trim() || `${firstVarPattern() || 'Archivo'}`;
        const payload = JSON.stringify({ fileName: pattern, items }, null, 2);
        const url = renderUrl() + '/batch' + (S.opts.strict ? '?strict=1' : '');
        if (S.batchTab === 'powershell') {
            return `$body = @'\n${payload}\n'@\nInvoke-WebRequest -Method Post -Uri ${psQ(url)} \`\n  -ContentType 'application/json; charset=utf-8' \`\n  -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) \`\n  -OutFile 'lote.zip'`;
        }
        return `curl -sS -f -X POST ${shQ(url)} \\\n  -H 'Content-Type: application/json' \\\n  --data-binary @- \\\n  -o lote.zip <<'JSON'\n${payload}\nJSON`;
    }
    function firstVarPattern() {
        for (const s of S.sheets) { if (s.isRepeatable) continue; const v = s.variables.find(x => !x.inRowTemplate && x.type === 'text'); if (v) return `${S.name} {{${v.name}}}`; }
        return S.name;
    }

    function selectTab(k, focus) {
        S.tab = k;
        S.root.querySelectorAll('#au-tabs [role="tab"]').forEach(t => {
            const on = t.dataset.tab === k;
            t.setAttribute('aria-selected', on ? 'true' : 'false'); t.tabIndex = on ? 0 : -1;
            if (on && focus) t.focus();
        });
        $('au-snip-panel').setAttribute('aria-labelledby', 'au-tab-' + k);
        refreshSnippets();
    }

    let snipTimer = 0;
    function refreshSnippetsSoon() { clearTimeout(snipTimer); snipTimer = setTimeout(refreshSnippets, 180); }
    function refreshSnippets() {
        if (!S.root || !S.tplId) return;
        $('au-url').textContent = renderUrl();
        const sn = snippet(S.tab);
        $('au-snip-note').textContent = sn.note || '';
        $('au-snip-code').textContent = sn.code || '';
        $('au-snip-extra').innerHTML = sn.link
            ? `<button type="button" class="au-btn" data-au="copy-link">${ico('i-copy')}<span>Copiar enlace</span></button><button type="button" class="au-btn" data-au="open-link">${ico('i-external')}<span>Abrir</span></button>`
            : '';
        $('au-batch-code').textContent = batchSnippet();
        $('au-render-label').textContent = S.opts.format === 'base64' ? 'Generar respuesta JSON (base64)' : 'Generar y descargar';
        $('au-snip-code').classList.toggle('au-code-wrapline', S.tab === 'link');
        S.root.querySelectorAll('[data-au="batch-tab"]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.lang === S.batchTab)));
        // Vista previa del nombre de archivo resultante
        const pv = $('au-fn-preview');
        pv.innerHTML = `${ico('i-file-spreadsheet')}<span>Se descargará como <b>${esc(outName())}</b></span>`;
        if (!S.root.querySelector('#au-tabs [aria-selected="true"]')) selectTab(S.tab);
    }

    // ── Portapapeles con fallback ────────────────────────────────────────────
    async function copyText(text, btn) {
        let ok = false;
        try { if (navigator.clipboard && window.isSecureContext) { await navigator.clipboard.writeText(text); ok = true; } } catch { /* fallback */ }
        if (!ok) {
            const ta = document.createElement('textarea');
            ta.value = text; ta.setAttribute('readonly', ''); ta.style.cssText = 'position:fixed;top:-1000px;opacity:0';
            S.root.appendChild(ta); ta.select();
            try { ok = document.execCommand('copy'); } catch { ok = false; }
            ta.remove();
        }
        if (!ok) { toast('No se pudo copiar; selecciona el texto manualmente', 'error'); return; }
        if (btn) {
            const span = btn.querySelector('span');
            btn.classList.add('au-copied');
            if (span) { span.dataset.orig = span.dataset.orig || span.textContent; span.textContent = 'Copiado'; }
            clearTimeout(btn._auT);
            btn._auT = setTimeout(() => { btn.classList.remove('au-copied'); if (span) span.textContent = span.dataset.orig; }, 1600);
        }
        toast('Copiado al portapapeles', 'success');
    }

    // ── Registro ──────────────────────────────────────────────────────────────
    ETB.openAutomation = openAutomation;
    // Si index.html redefine el puente después de cargar este módulo, vuelve a registrarse.
    if (typeof ETB.whenReady === 'function') {
        try { ETB.whenReady(() => { if (window.ETB && window.ETB.openAutomation !== openAutomation) window.ETB.openAutomation = openAutomation; }); } catch { /* opcional */ }
    }
    ETB.automation = { open: openAutomation, close, isOpen: () => S.open };
    // Para pruebas internas (no forma parte del contrato)
    ETB.automation._snippet = k => snippet(k);
})();
