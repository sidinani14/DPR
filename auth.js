/* ============================================================
   Ideaform Design Studio — shared access gate
   Include in <head> on EVERY page:  <script src="auth.js"></script>
   - Hides the whole page until an allowlisted team Google account signs in
   - Attaches the Google ID token to every backend API call (fetch patch)
   - Pages should start loading data from window.idsOnAuth(user) (called once
     after a successful, authorized sign-in) OR by listening for the
     'ids-auth' event on document.
   The real security boundary is the backend token check; this is the UI wall.
   ============================================================ */
(function () {
  var CLIENT_ID = '48052407111-mantkqn708ejp5otfc34nch2ngl8o9ot.apps.googleusercontent.com';
  var API_HOST = 'script.google.com/macros';
  var ACCESS_CHECK_URL = 'https://script.google.com/macros/s/AKfycbxoYY488eYAomVcsP9h3TwlYZIWDg0gmn4qrCyUiJbriAUIRJr_19VH0RM3NRZPBUoKYA/exec';
  var STORE_KEY = 'ids_token';

  // ── Storage helpers — localStorage so sign-in persists across all tabs ──
  // Falls back to sessionStorage if localStorage is blocked (private browsing).
  function storeGet(k)    { try { return localStorage.getItem(k); }    catch(e){ try { return sessionStorage.getItem(k); }    catch(_){} } return null; }
  function storeSet(k, v) { try { localStorage.setItem(k, v); }         catch(e){ try { sessionStorage.setItem(k, v); }         catch(_){} } }
  function storeDel(k)    { try { localStorage.removeItem(k); }          catch(e){ try { sessionStorage.removeItem(k); }          catch(_){} } }

  // ── 1. Cover the page with an opaque overlay until auth ──
  try {
    var st = document.createElement('style');
    st.id = 'ids-lock-style';
    st.textContent = 'html{background:#0F0F0F}';
    (document.head || document.documentElement).appendChild(st);
  } catch (e) {}

  // ── 2. Patch fetch so every backend call carries the ID token, and so an
  //       "unauthorized" reply surfaces as a clear access-denied screen ──
  var _fetch = window.fetch;
  window.fetch = function (url, opts) {
    opts = opts || {};
    var tok = window.IDS_TOKEN;
    var isBackend = typeof url === 'string' && url.indexOf(API_HOST) > -1;
    if (isBackend && tok) {
      var m = (opts.method || 'GET').toUpperCase();
      if (m === 'GET') {
        url += (url.indexOf('?') > -1 ? '&' : '?') + 'idToken=' + encodeURIComponent(tok);
      } else if (typeof opts.body === 'string') {
        try { var b = JSON.parse(opts.body); b.idToken = tok; opts.body = JSON.stringify(b); } catch (e) {}
      }
    }
    return _fetch.call(this, url, opts).then(function (resp) {
      if (isBackend && !window.__idsDenied) {
        try {
          resp.clone().json().then(function (j) {
            if (j && j.code === 'unauthorized') {
              window.__idsDenied = true;
              window.IDS_TOKEN = null;
              try { storeDel(STORE_KEY); google.accounts.id.disableAutoSelect(); } catch (e) {}
              buildGate('denied',
                'Your account isn’t authorized for this system. Ask your admin to add your exact Google email to the TEAM tab (and mark it active).',
                (window.IDS_USER && window.IDS_USER.email) || '');
            }
          }).catch(function () {});
        } catch (e) {}
      }
      return resp;
    });
  };

  function parseJwt(t) {
    try { return JSON.parse(atob(t.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'))); }
    catch (e) { return null; }
  }

  // ── 3. Gate UI ──
  function buildGate(state, msg, who) {
    var g = document.getElementById('ids-gate');
    if (!g) {
      g = document.createElement('div');
      g.id = 'ids-gate';
      g.style.cssText = 'position:fixed;inset:0;z-index:2147483647;background:#0F0F0F;color:#F0EDE8;' +
        'display:flex;align-items:center;justify-content:center;font-family:system-ui,-apple-system,Segoe UI,sans-serif';
      document.body.appendChild(g);
    }
    var inner = '<div style="text-align:center;max-width:340px;padding:28px">' +
      '<div style="font-family:Syne,system-ui,sans-serif;font-weight:700;font-size:18px;color:#C8A96E;letter-spacing:.5px">IDEAFORM DESIGN STUDIO</div>' +
      '<div style="font-size:12px;color:#5A5652;margin-top:4px;font-family:DM Mono,monospace">Internal · authorized team access only</div>';
    if (state === 'signin') {
      inner += '<div style="margin-top:26px;font-size:13px;color:#9A9690">Sign in with your Ideaform Google account</div>' +
        '<div id="ids-btn" style="margin-top:16px;display:flex;justify-content:center"></div>' +
        '<div style="margin-top:14px;font-size:11px;color:#5A5652">Only emails on the team list can enter.</div>';
    } else if (state === 'denied') {
      inner += '<div style="margin-top:26px;font-size:34px">🔒</div>' +
        '<div style="margin-top:10px;font-size:14px;color:#C05050;font-weight:600">Access denied</div>' +
        '<div style="margin-top:8px;font-size:12px;color:#9A9690">' + (msg || '') +
        (who ? '<br><span style="color:#5A5652;font-family:DM Mono,monospace">' + who + '</span>' : '') + '</div>' +
        '<button id="ids-retry" style="margin-top:18px;padding:8px 16px;font-size:12px;color:#C8A96E;background:rgba(200,169,110,.1);border:1px solid rgba(200,169,110,.3);border-radius:6px;cursor:pointer">Try another account</button>';
    } else {
      inner += '<div style="margin-top:26px;font-size:13px;color:#9A9690">Checking access…</div>';
    }
    inner += '</div>';
    g.innerHTML = inner;
    if (state === 'signin') renderButton();
    if (state === 'denied') {
      var rb = document.getElementById('ids-retry');
      if (rb) rb.onclick = function () {
        try { storeDel(STORE_KEY); google.accounts.id.disableAutoSelect(); } catch (e) {}
        buildGate('signin');
      };
    }
  }

  function renderButton() {
    if (!(window.google && google.accounts && google.accounts.id)) { setTimeout(renderButton, 200); return; }
    var el = document.getElementById('ids-btn');
    if (el) google.accounts.id.renderButton(el, { theme: 'filled_black', size: 'large', shape: 'pill', text: 'signin_with' });
  }

  function onCredential(resp) {
    var token = resp && resp.credential;
    var p = token ? parseJwt(token) : null;
    if (!p) { buildGate('denied', 'Sign-in failed. Please try again.'); return; }
    grant(token, p);
  }

  function grant(token, p) {
    window.IDS_TOKEN = token;
    window.IDS_USER = { email: String(p.email || '').toLowerCase(), name: p.name || p.email || '', picture: p.picture || '' };
    storeSet(STORE_KEY, token);
    // reveal page
    var s = document.getElementById('ids-lock-style'); if (s) s.remove();
    var g = document.getElementById('ids-gate'); if (g) g.remove();
    document.dispatchEvent(new CustomEvent('ids-auth', { detail: window.IDS_USER }));
    if (typeof window.idsOnAuth === 'function') { try { window.idsOnAuth(window.IDS_USER); } catch (e) {} }
  }

  // ── 4. Boot ──
  function boot() {
    // Reuse a stored token — localStorage means it's shared across all open tabs.
    // Only use it if not expired (keep 60s margin) AND issued today (end-of-day reset).
    try {
      var saved = storeGet(STORE_KEY);
      if (saved) {
        var sp = parseJwt(saved);
        var nowMs = Date.now();
        // Midnight tonight (local time) — token is cleared after working day ends
        var midnight = new Date(); midnight.setHours(24,0,0,0);
        if (sp && sp.exp && sp.exp * 1000 > nowMs + 60000 && nowMs < midnight.getTime()) {
          grant(saved, sp);
          return;
        }
        storeDel(STORE_KEY);
      }
    } catch (e) {}

    // No valid stored token — try silent FedCM re-auth first, then show sign-in UI
    buildGate('checking');
    loadGis(function () {
      try {
        google.accounts.id.initialize({
          client_id: CLIENT_ID,
          callback: onCredential,
          auto_select: true,
          use_fedcm_for_prompt: true,
        });
      } catch (e) {}
      // prompt() will silently call onCredential if the browser can auto-sign in.
      // If it can't (token fully expired, multiple accounts, etc.), it does nothing
      // and we fall through to show the sign-in button.
      try { google.accounts.id.prompt(function(n){ if(n.isNotDisplayed()||n.isSkippedMoment()) buildGate('signin'); }); }
      catch(e) { buildGate('signin'); }
    });
  }

  function loadGis(cb) {
    if (window.google && google.accounts && google.accounts.id) { cb(); return; }
    var existing = document.querySelector('script[src*="accounts.google.com/gsi/client"]');
    if (existing) { existing.addEventListener('load', cb); if (window.google && google.accounts) cb(); return; }
    var s = document.createElement('script');
    s.src = 'https://accounts.google.com/gsi/client';
    s.async = true; s.defer = true; s.onload = cb;
    document.head.appendChild(s);
  }

  // Expose sign-out so pages can wire it to a button
  window.IDS_SIGNOUT = function () {
    storeDel(STORE_KEY);
    window.IDS_TOKEN = null;
    try { google.accounts.id.disableAutoSelect(); } catch (e) {}
    location.href = 'index.html';
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
