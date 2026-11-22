/*
 * fx.js — Excel Template Builder v12
 * Micro-animaciones reutilizables (contrato §3.6):
 *   ETB.fx.blurIn(el)                  entrada con desenfoque
 *   ETB.fx.stagger(container, sel?)    entrada escalonada de hijos
 *   ETB.fx.countUp(el, to, ms?)        contador animado (es-MX)
 *   ETB.fx.flash(el)                   destello de "valor actualizado"
 * Además, de forma automática y sin tocar el HTML:
 *   - Tarjetas de Inicio: entrada escalonada solo en la primera pintura tras entrar a
 *     Inicio (no en cada búsqueda ni al regenerar miniaturas).
 *   - Pestañas de documento: pop-in solo para la pestaña recién abierta.
 * Las clases/keyframes viven en css/effects.css. Respeta prefers-reduced-motion.
 * Ninguna animación deja filter/transform aplicado al terminar. Sin dependencias.
 */
(function () {
    'use strict';
    var ETB = window.ETB = window.ETB || {};
    if (ETB.fx && ETB.fx.__v12) return; // idempotente si el script se carga dos veces

    var mq = window.matchMedia ? window.matchMedia('(prefers-reduced-motion: reduce)') : null;
    function reduced() { return !!(mq && mq.matches); }

    /** Reinicia una clase de animación y la retira al terminar (sin dejar estilos residuales). */
    function play(el, cls, maxMs) {
        if (!el || !el.classList) return el;
        if (el._fxDone) el._fxDone(); // cancela una reproducción previa de este elemento
        el.classList.remove(cls);
        void el.offsetWidth; // fuerza reflow para reiniciar la animación
        el.classList.add(cls);
        var t;
        var onEnd = function (e) { if (e.target === el) done(); };
        var done = function () {
            el.classList.remove(cls);
            el.removeEventListener('animationend', onEnd);
            clearTimeout(t);
            if (el._fxDone === done) el._fxDone = null;
        };
        el._fxDone = done;
        el.addEventListener('animationend', onEnd);
        t = setTimeout(done, maxMs || 1500); // respaldo si no hay animación (CSS ausente, display:none…)
        return el;
    }

    function blurIn(el) {
        if (!el || reduced()) return el;
        return play(el, 'fx-blur-in', 1500);
    }

    function stagger(container, selector) {
        if (!container) return;
        var items = selector ? container.querySelectorAll(selector) : container.children;
        Array.prototype.forEach.call(items, function (it, i) {
            if (!it.style) return;
            it.style.setProperty('--i', String(Math.min(i, 12))); // retraso acotado
            if (!reduced()) play(it, 'fx-blur-in', 2500);
        });
    }

    var fmtCache = {};
    function format(n, decimals) {
        try {
            var f = fmtCache[decimals] || (fmtCache[decimals] = new Intl.NumberFormat('es-MX', {
                minimumFractionDigits: decimals, maximumFractionDigits: decimals
            }));
            return f.format(n);
        } catch (_) { return decimals ? n.toFixed(decimals) : String(Math.round(n)); }
    }

    /** Anima el texto de `el` desde su último valor hasta `to` (ms por defecto 700). */
    function countUp(el, to, ms) {
        if (!el) return;
        var target = Number(to);
        if (to === null || to === undefined || to === '' || !isFinite(target)) {
            el.textContent = to == null ? '' : String(to);
            return;
        }
        var decimals = Math.min(2, (String(to).split('.')[1] || '').length);
        if (el._fxCount) { cancelAnimationFrame(el._fxCount); el._fxCount = 0; }
        var from = el.dataset ? Number(el.dataset.fxValue) : NaN;
        if (!isFinite(from)) from = 0;
        if (el.dataset) el.dataset.fxValue = String(target);
        if (reduced() || !window.requestAnimationFrame || from === target) {
            el.textContent = format(target, decimals);
            return;
        }
        var dur = Math.max(0, ms == null ? 700 : Number(ms) || 0), start = null;
        var step = function (now) {
            if (start === null) start = now;
            var p = dur ? Math.min(1, (now - start) / dur) : 1;
            var eased = 1 - Math.pow(1 - p, 3); // easeOutCubic
            el.textContent = format(p < 1 ? from + (target - from) * eased : target, decimals);
            el._fxCount = p < 1 ? requestAnimationFrame(step) : 0;
        };
        el._fxCount = requestAnimationFrame(step);
    }

    function flash(el) {
        if (!el || reduced()) return el;
        return play(el, 'fx-flash', 1500);
    }

    // ── Automático: tarjetas de Inicio ─────────────────────────────────────
    // index.html repinta #home-grid completo en cada búsqueda y tras cada miniatura;
    // la entrada escalonada solo se "arma" al entrar a Inicio y se consume en la
    // primera pintura con tarjetas.
    function watchHomeCards() {
        var view = document.getElementById('home-view');
        var grid = document.getElementById('home-grid');
        if (!view || !grid || !window.MutationObserver) return;
        var armed = true, timer = 0, t0 = 0;
        var CARD_WINDOW_MS = 1000; // retraso máx. (440 ms) + duración (420 ms) + margen

        function now() { return window.performance && performance.now ? performance.now() : Date.now(); }
        function finish() {
            clearTimeout(timer); timer = 0;
            grid.classList.remove('fx-cards-in');
            grid.style.removeProperty('--fx-elapsed');
        }
        function hasCards() { return !!grid.querySelector('.tpl-card'); }
        function maybeAnimate() {
            if (timer) {
                // Repintado durante la entrada (p.ej. llegó una miniatura): las tarjetas nuevas
                // continúan la animación en el punto actual en lugar de reiniciarla.
                grid.style.setProperty('--fx-elapsed', Math.round(now() - t0) + 'ms');
                return;
            }
            if (!armed || reduced() || view.classList.contains('hidden') || !hasCards()) return;
            armed = false;
            t0 = now();
            grid.style.setProperty('--fx-elapsed', '0ms');
            grid.classList.add('fx-cards-in');
            timer = setTimeout(finish, CARD_WINDOW_MS);
        }
        new MutationObserver(maybeAnimate).observe(grid, { childList: true });
        new MutationObserver(function () {
            if (view.classList.contains('hidden')) { armed = true; finish(); }
        }).observe(view, { attributes: true, attributeFilter: ['class'] });
        if (hasCards()) maybeAnimate(); // por si ya se pintó antes de cargar fx.js
    }

    // ── Automático: pestaña de documento recién abierta ────────────────────
    function watchDocTabs() {
        var list = document.getElementById('doc-tabs-list');
        if (!list || !window.MutationObserver) return;
        var count = list.querySelectorAll('.doc-tab').length;
        new MutationObserver(function () {
            var tabs = list.querySelectorAll('.doc-tab');
            if (tabs.length > count && !reduced()) {
                // openDoc() agrega al final; renderDocTabs() reconstruye todo, así que se
                // anima solo la última pestaña y solo cuando aumenta el número.
                play(tabs[tabs.length - 1], 'fx-tab-new', 800);
            }
            count = tabs.length;
        }).observe(list, { childList: true });
    }

    function initAuto() {
        try { watchHomeCards(); } catch (e) { if (window.console) console.warn('[fx] tarjetas:', e); }
        try { watchDocTabs(); } catch (e) { if (window.console) console.warn('[fx] pestañas:', e); }
    }
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initAuto, { once: true });
    else initAuto();

    ETB.fx = { blurIn: blurIn, stagger: stagger, countUp: countUp, flash: flash, reducedMotion: reduced, __v12: true };
})();
