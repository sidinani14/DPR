/* ============================================================================
 * searchselect.js — lightweight, dependency-free searchable dropdown.
 *
 * Replaces the native <select> wheel (painful on iPad/iPhone for long lists)
 * with a tap-to-search control: tap → keyboard opens → type to filter → pick.
 *
 * Contract preserved: the real <select> stays in the DOM. We only write back
 * to select.value and dispatch a bubbling 'change' event, so every existing
 * handler (inline onchange="" and addEventListener) and all collection code
 * that reads select.value keeps working unchanged.
 *
 * Usage (per page, after this script is included):
 *     SearchSelect.auto('.p-name, .t-cat, .t-type');
 * Selectors accumulate across calls. Works for selects added later
 * (dynamic blocks) — a MutationObserver upgrades them as they appear, and a
 * per-select observer keeps the control in sync when its <option>s change.
 *
 * This is a browser file: it is auto-excluded from clasp by the all-JS
 * ignore rule in .claspignore, so it never gets pushed as server code.
 * ==========================================================================*/
(function () {
  'use strict';
  if (window.SearchSelect) return;

  var STYLE_ID = 'ss-styles';
  var registry = [];          // selectors to enhance
  var observing = false;
  var rafPending = false;

  // ── styles (dark-form friendly; override with --ss-* CSS vars) ──────────
  function injectStyles() {
    if (document.getElementById(STYLE_ID)) return;
    var css =
      '.ss-wrap{position:relative;display:block;width:100%;font-family:inherit}' +
      '.ss-trigger{display:flex;align-items:center;justify-content:space-between;gap:8px;' +
        'width:100%;box-sizing:border-box;cursor:pointer;text-align:left;' +
        'font:inherit;font-size:13px;line-height:1.3;padding:8px 11px;border-radius:8px;' +
        'background:var(--ss-bg,#1A1A1A);color:var(--ss-fg,#F0EDE8);' +
        'border:1px solid var(--ss-border,rgba(255,255,255,0.16));' +
        '-webkit-appearance:none;appearance:none}' +
      '.ss-trigger:focus{outline:none;border-color:var(--ss-accent,#C8A96E)}' +
      '.ss-trigger.ss-open{border-color:var(--ss-accent,#C8A96E)}' +
      '.ss-trigger .ss-cur{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;min-width:0}' +
      '.ss-trigger .ss-cur.ss-ph{color:var(--ss-muted,rgba(255,255,255,0.4))}' +
      '.ss-caret{flex-shrink:0;width:0;height:0;border-left:4px solid transparent;' +
        'border-right:4px solid transparent;border-top:5px solid var(--ss-muted,rgba(255,255,255,0.45))}' +
      '.ss-pop{position:fixed;z-index:99999;' +
        'background:var(--ss-pop-bg,#161616);border:1px solid var(--ss-border,rgba(255,255,255,0.16));' +
        'border-radius:10px;box-shadow:0 12px 32px rgba(0,0,0,0.55);overflow:hidden;display:none}' +
      '.ss-pop.ss-show{display:block}' +
      '.ss-search-wrap{padding:8px;border-bottom:1px solid var(--ss-border,rgba(255,255,255,0.1))}' +
      '.ss-search{width:100%;box-sizing:border-box;font:inherit;font-size:13px;padding:8px 10px;' +
        'border-radius:7px;background:var(--ss-bg,#1A1A1A);color:var(--ss-fg,#F0EDE8);' +
        'border:1px solid var(--ss-border,rgba(255,255,255,0.16))}' +
      '.ss-search:focus{outline:none;border-color:var(--ss-accent,#C8A96E)}' +
      '.ss-list{max-height:240px;overflow-y:auto;-webkit-overflow-scrolling:touch;padding:4px}' +
      '.ss-opt{padding:9px 11px;font-size:13px;border-radius:6px;cursor:pointer;color:var(--ss-fg,#F0EDE8);' +
        'white-space:normal;word-break:break-word}' +
      '.ss-opt:hover,.ss-opt.ss-active{background:var(--ss-hover,rgba(200,169,110,0.18))}' +
      '.ss-opt.ss-sel{color:var(--ss-accent,#C8A96E);font-weight:600}' +
      '.ss-opt.ss-ph{color:var(--ss-muted,rgba(255,255,255,0.4))}' +
      '.ss-empty{padding:12px;text-align:center;font-size:12px;color:var(--ss-muted,rgba(255,255,255,0.4))}';
    var el = document.createElement('style');
    el.id = STYLE_ID;
    el.textContent = css;
    (document.head || document.documentElement).appendChild(el);
  }

  function curText(select) {
    var o = select.options[select.selectedIndex];
    return o ? o.text : '';
  }
  function isPlaceholder(select) {
    var o = select.options[select.selectedIndex];
    return !o || o.value === '';
  }

  // ── enhance a single <select> ───────────────────────────────────────────
  function enhance(select) {
    if (!select || select.__ss) return;
    select.__ss = true;

    var wrap = document.createElement('div');
    wrap.className = 'ss-wrap';
    var trigger = document.createElement('button');
    trigger.type = 'button';
    trigger.className = 'ss-trigger';
    trigger.innerHTML = '<span class="ss-cur"></span><span class="ss-caret"></span>';
    var pop = document.createElement('div');
    pop.className = 'ss-pop';
    pop.innerHTML =
      '<div class="ss-search-wrap"><input type="text" class="ss-search" ' +
      'placeholder="Type to search…" autocomplete="off" autocorrect="off" ' +
      'autocapitalize="off" spellcheck="false"></div><div class="ss-list"></div>';

    // place wrapper where the select is, then move the (now hidden) select inside
    select.parentNode.insertBefore(wrap, select);
    wrap.appendChild(trigger);
    wrap.appendChild(pop);
    wrap.appendChild(select);
    select.style.display = 'none';

    var curEl = trigger.querySelector('.ss-cur');
    var search = pop.querySelector('.ss-search');
    var list = pop.querySelector('.ss-list');
    var activeIdx = -1;
    var visibleOpts = [];   // [{value,text,el}]

    function syncTrigger() {
      curEl.textContent = curText(select) || 'Select…';
      curEl.classList.toggle('ss-ph', isPlaceholder(select));
    }

    function buildList() {
      var q = (search.value || '').trim().toLowerCase();
      list.innerHTML = '';
      visibleOpts = [];
      activeIdx = -1;
      for (var i = 0; i < select.options.length; i++) {
        var o = select.options[i];
        var txt = o.text;
        if (q && txt.toLowerCase().indexOf(q) === -1) continue;
        var row = document.createElement('div');
        row.className = 'ss-opt';
        if (o.value === '') row.className += ' ss-ph';
        if (i === select.selectedIndex) row.className += ' ss-sel';
        row.textContent = txt;
        row.setAttribute('data-i', i);
        list.appendChild(row);
        visibleOpts.push({ idx: i, el: row });
      }
      if (!visibleOpts.length) {
        var e = document.createElement('div');
        e.className = 'ss-empty';
        e.textContent = 'No matches';
        list.appendChild(e);
      }
    }

    function setActive(n) {
      if (!visibleOpts.length) return;
      if (n < 0) n = 0;
      if (n > visibleOpts.length - 1) n = visibleOpts.length - 1;
      visibleOpts.forEach(function (v) { v.el.classList.remove('ss-active'); });
      activeIdx = n;
      var el = visibleOpts[n].el;
      el.classList.add('ss-active');
      var elTop = el.offsetTop, elBot = elTop + el.offsetHeight;
      if (elTop < list.scrollTop) list.scrollTop = elTop;
      else if (elBot > list.scrollTop + list.clientHeight) list.scrollTop = elBot - list.clientHeight;
    }

    function choose(optIdx) {
      if (select.selectedIndex !== optIdx) {
        select.selectedIndex = optIdx;
        select.dispatchEvent(new Event('input', { bubbles: true }));
        select.dispatchEvent(new Event('change', { bubbles: true }));
      }
      syncTrigger();
      close();
    }

    // Position the (position:fixed) popup against the trigger. Fixed positioning
    // escapes ancestor overflow:hidden (cards/sections clip an absolute popup).
    function position() {
      var r = trigger.getBoundingClientRect();
      var vh = window.innerHeight, gap = 4, pad = 10;
      pop.style.left = r.left + 'px';
      pop.style.width = r.width + 'px';
      var below = vh - r.bottom - gap - pad;
      var above = r.top - gap - pad;
      var up = below < 220 && above > below;
      if (up) { pop.style.top = 'auto'; pop.style.bottom = (vh - r.top + gap) + 'px'; }
      else    { pop.style.bottom = 'auto'; pop.style.top = (r.bottom + gap) + 'px'; }
      // let the option list scroll within the space available in the viewport
      list.style.maxHeight = Math.max(96, (up ? above : below) - 52) + 'px';
    }
    var _reposition = function () { if (pop.classList.contains('ss-show')) position(); };
    function open() {
      if (pop.classList.contains('ss-show')) return;
      closeAll();
      injectStyles();
      pop.classList.add('ss-show');
      trigger.classList.add('ss-open');
      search.value = '';
      buildList();
      position();
      window.addEventListener('scroll', _reposition, true);
      window.addEventListener('resize', _reposition);
      // focus the search box → opens the mobile keyboard
      setTimeout(function () { search.focus(); }, 0);
      openInstance = api;
    }
    function close() {
      pop.classList.remove('ss-show');
      trigger.classList.remove('ss-open');
      window.removeEventListener('scroll', _reposition, true);
      window.removeEventListener('resize', _reposition);
      if (openInstance === api) openInstance = null;
    }

    // events
    trigger.addEventListener('click', function (e) {
      e.preventDefault();
      pop.classList.contains('ss-show') ? close() : open();
    });
    search.addEventListener('input', buildList);
    search.addEventListener('keydown', function (e) {
      if (e.key === 'ArrowDown') { e.preventDefault(); setActive(activeIdx + 1); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(activeIdx - 1); }
      else if (e.key === 'Enter') {
        e.preventDefault();
        if (activeIdx >= 0 && visibleOpts[activeIdx]) choose(visibleOpts[activeIdx].idx);
        else if (visibleOpts.length === 1) choose(visibleOpts[0].idx);
      } else if (e.key === 'Escape') { e.preventDefault(); close(); trigger.focus(); }
    });
    list.addEventListener('click', function (e) {
      var row = e.target.closest('.ss-opt');
      if (!row) return;
      choose(parseInt(row.getAttribute('data-i'), 10));
    });

    // keep the control in sync when the <select>'s options change (dynamic fill)
    var mo = new MutationObserver(function () {
      syncTrigger();
      if (pop.classList.contains('ss-show')) buildList();
    });
    mo.observe(select, { childList: true });
    // and when other code changes the value
    select.addEventListener('change', syncTrigger);

    var api = { close: close, sync: syncTrigger };
    syncTrigger();
  }

  // ── single open instance + outside-click handling ───────────────────────
  var openInstance = null;
  function closeAll() { if (openInstance) openInstance.close(); }
  document.addEventListener('click', function (e) {
    if (openInstance && !e.target.closest('.ss-wrap')) closeAll();
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') closeAll();
  });

  // ── sweep current matches for all registered selectors ──────────────────
  function sweep() {
    rafPending = false;
    if (!registry.length) return;
    var sel = registry.join(',');
    var nodes;
    try { nodes = document.querySelectorAll(sel); } catch (e) { return; }
    for (var i = 0; i < nodes.length; i++) {
      if (nodes[i].tagName === 'SELECT') enhance(nodes[i]);
    }
  }
  function scheduleSweep() {
    if (rafPending) return;
    rafPending = true;
    (window.requestAnimationFrame || window.setTimeout)(sweep, 16);
  }
  function startObserver() {
    if (observing) return;
    observing = true;
    var run = function () {
      injectStyles();
      sweep();
      new MutationObserver(scheduleSweep)
        .observe(document.documentElement, { childList: true, subtree: true });
    };
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', run);
    else run();
  }

  window.SearchSelect = {
    // register selector(s); enhances current + future matching <select>s
    auto: function (selector) {
      if (selector) registry.push(selector);
      startObserver();
      scheduleSweep();
    },
    // manually enhance a specific element
    enhance: enhance
  };
})();
