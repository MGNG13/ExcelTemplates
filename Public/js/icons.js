/*
 * icons.js — Excel Template Builder v12
 * Inyecta de forma síncrona los íconos nuevos (contrato §3.7) en el sprite SVG
 * (#icon-sprite). Estilo trazo 24×24: el color/grosor se hereda de la clase .ico
 * (stroke: currentColor; fill: none). Idempotente: omite ids ya existentes.
 * Sin dependencias.
 */
(function () {
    'use strict';
    var SVG_NS = 'http://www.w3.org/2000/svg';

    // Geometría estilo Lucide (solo trazos). Cada valor es el contenido interno del <symbol>.
    var ICONS = {
        'i-users': '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>',
        'i-user': '<path d="M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>',
        'i-user-plus': '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M19 8v6M22 11h-6"/>',
        'i-activity': '<path d="M22 12h-4l-3 9L9 3l-3 9H2"/>',
        'i-chart-bar': '<path d="M3 3v18h18"/><path d="M8 17v-5M13 17V8M18 17v-9"/>',
        'i-trending-up': '<path d="M22 7l-8.5 8.5-5-5L2 17"/><path d="M16 7h6v6"/>',
        'i-bell': '<path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/>',
        'i-download': '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5 5 5-5M12 15V3"/>',
        'i-check-circle': '<circle cx="12" cy="12" r="10"/><path d="M9 12l2 2 4-4"/>',
        'i-alert-triangle': '<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><path d="M12 9v4M12 17h.01"/>',
        'i-alert-circle': '<circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/>',
        'i-info': '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/>',
        'i-wifi': '<path d="M5 12.55a11 11 0 0 1 14.08 0M1.42 9a16 16 0 0 1 21.16 0M8.53 16.11a6 6 0 0 1 6.95 0M12 20h.01"/>',
        'i-wifi-off': '<path d="M2 2l20 20M8.5 16.5a5 5 0 0 1 7 0M2 8.82a15 15 0 0 1 4.17-2.65M10.66 5c4.01-.36 8.14.9 11.34 3.76M16.85 11.25a10 10 0 0 1 2.22 1.68M5 13a10 10 0 0 1 5.24-2.76M12 20h.01"/>',
        'i-code': '<path d="M16 18l6-6-6-6M8 6l-6 6 6 6"/>',
        'i-terminal': '<path d="M4 17l6-6-6-6M12 19h8"/>',
        'i-link': '<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>',
        'i-external': '<path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><path d="M15 3h6v6M10 14L21 3"/>',
        'i-refresh': '<path d="M21 12a9 9 0 0 1-15.36 6.36L3 16M3 12a9 9 0 0 1 15.36-6.36L21 8"/><path d="M21 3v5h-5M3 21v-5h5"/>',
        'i-play': '<path d="M6 3l14 9-14 9V3z"/>',
        'i-server': '<rect x="2" y="2" width="20" height="8" rx="2"/><rect x="2" y="14" width="20" height="8" rx="2"/><path d="M6 6h.01M6 18h.01"/>',
        'i-key': '<circle cx="7.5" cy="15.5" r="5.5"/><path d="M11.4 11.6L21 2M15.5 7.5l3 3M18 5l2 2"/>',
        'i-book': '<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/>',
        'i-sparkles': '<path d="M12 3l1.9 5.1L19 10l-5.1 1.9L12 17l-1.9-5.1L5 10l5.1-1.9z"/><path d="M19 15v4M17 17h4M5 3v3M3.5 4.5h3"/>',
        'i-package': '<path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><path d="M3.3 7L12 12l8.7-5M12 22V12M7.5 4.27l9 5.15"/>',
        'i-list': '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>',
        'i-filter': '<path d="M22 3H2l8 9.46V19l4 2v-8.54z"/>',
        'i-pencil': '<path d="M17 3a2.85 2.85 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5z"/><path d="M15 5l4 4"/>',
        'i-cursor': '<path d="M3 3l7.07 16.97 2.51-7.39 7.39-2.51z"/><path d="M13 13l6 6"/>',
        'i-globe': '<circle cx="12" cy="12" r="10"/><path d="M2 12h20M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>',
        'i-calendar': '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
        'i-hash': '<path d="M4 9h16M4 15h16M10 3L8 21M16 3l-2 18"/>',
        'i-database': '<ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/>',
        'i-send': '<path d="M22 2L11 13"/><path d="M22 2l-7 20-4-9-9-4z"/>',
        'i-clipboard': '<rect x="8" y="2" width="8" height="4" rx="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/>',
        'i-gauge': '<path d="M12 14l4-4"/><path d="M3.34 19a10 10 0 1 1 17.32 0"/>',
        'i-history': '<path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5M12 7v5l4 2"/>',
        'i-settings': '<path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>',
        'i-grid': '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/>',
        'i-webhook': '<path d="M18 16.98h-5.99c-1.1 0-1.95.94-2.48 1.9A4 4 0 0 1 2 17c.01-.7.2-1.4.57-2"/><path d="M6 17l3.13-5.78c.53-.97.1-2.18-.5-3.1a4 4 0 1 1 6.89-4.06"/><path d="M12 6l3.13 5.73C15.66 12.7 16.9 13 18 13a4 4 0 0 1 0 8"/>',
        'i-file-spreadsheet': '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><rect x="8" y="12" width="8" height="6" rx="1"/><path d="M8 15h8M12 12v6"/>',
        'i-arrow-right': '<path d="M5 12h14M12 5l7 7-7 7"/>',
        'i-more': '<circle cx="12" cy="12" r="1" fill="currentColor"/><circle cx="19" cy="12" r="1" fill="currentColor"/><circle cx="5" cy="12" r="1" fill="currentColor"/>',
        'i-dot': '<circle cx="12" cy="12" r="4" fill="currentColor"/>'
    };

    /** Localiza el sprite: #icon-sprite, el <svg> que ya contiene <symbol>, o crea uno oculto. */
    function getSprite() {
        var sprite = document.getElementById('icon-sprite');
        if (sprite) return sprite;
        var sym = document.querySelector('svg > symbol');
        if (sym && sym.parentNode) return sym.parentNode;
        sprite = document.createElementNS(SVG_NS, 'svg');
        sprite.setAttribute('id', 'icon-sprite');
        sprite.setAttribute('width', '0');
        sprite.setAttribute('height', '0');
        sprite.setAttribute('aria-hidden', 'true');
        sprite.style.position = 'absolute';
        var host = document.body || document.documentElement;
        host.insertBefore(sprite, host.firstChild);
        return sprite;
    }

    function inject() {
        var sprite = getSprite();
        var frag = document.createDocumentFragment();
        Object.keys(ICONS).forEach(function (id) {
            if (document.getElementById(id)) return; // idempotente
            // Contenido estático propio (no externo): se parsea dentro de un <svg> temporal.
            var tmp = document.createElementNS(SVG_NS, 'svg');
            tmp.innerHTML = ICONS[id];
            var symbol = document.createElementNS(SVG_NS, 'symbol');
            symbol.setAttribute('id', id);
            symbol.setAttribute('viewBox', '0 0 24 24');
            while (tmp.firstChild) symbol.appendChild(tmp.firstChild);
            frag.appendChild(symbol);
        });
        sprite.appendChild(frag);
    }

    try { inject(); } catch (err) { if (window.console) console.warn('[icons] no se pudieron inyectar los íconos:', err); }
    // Lista pública (útil para pruebas / galerías).
    window.ETB_ICON_IDS = Object.keys(ICONS);
})();
