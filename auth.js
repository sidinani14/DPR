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
  var EXP_KEY   = 'ids_session_exp';
  var HINT_KEY  = 'ids_email_hint'; // persistent — never deleted, survives expiry
  var SESSION_MS = 12 * 60 * 60 * 1000;  // 12 hours

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
              var who = (window.IDS_USER && window.IDS_USER.email) || '';
              try { storeDel(STORE_KEY); storeDel(EXP_KEY); google.accounts.id.disableAutoSelect(); } catch (e) {}
              buildGate('denied',
                'The signed-in account (' + (who || 'unknown') + ') is not on the authorised team list.' +
                ' If you have multiple Google accounts in this browser, tap "Try another account" to switch.',
                who);
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
    var inner = '<div style="text-align:center;max-width:360px;padding:28px">' +
      '<div style="font-family:Syne,system-ui,sans-serif;font-weight:700;font-size:18px;color:#C8A96E;letter-spacing:.5px">IDEAFORM DESIGN STUDIO</div>' +
      '<div style="font-size:12px;color:#5A5652;margin-top:4px;font-family:DM Mono,monospace">Internal · authorized team access only</div>';
    if (state === 'signin') {
      // Show hint email if we know their account
      var hint = storeGet(HINT_KEY);
      var hintHtml = hint
        ? '<div style="margin-top:10px;font-size:11px;color:#5A5652">Signing in as <span style="color:#C8A96E;font-family:DM Mono,monospace">' + hint + '</span></div>'
        : '<div style="margin-top:10px;font-size:11px;color:#5A5652">Only emails on the team list can enter.</div>';
      inner += '<div style="margin-top:26px;font-size:13px;color:#9A9690">Sign in with your Ideaform Google account</div>' +
        '<div id="ids-btn" style="margin-top:16px;display:flex;justify-content:center"></div>' +
        hintHtml;
    } else if (state === 'denied') {
      inner += '<div style="margin-top:26px;font-size:34px">🔒</div>' +
        '<div style="margin-top:10px;font-size:14px;color:#C05050;font-weight:600">Access denied</div>' +
        '<div style="margin-top:8px;font-size:12px;color:#9A9690;line-height:1.5">' + (msg || '') + '</div>' +
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
        window.__idsDenied = false;
        try { storeDel(STORE_KEY); storeDel(EXP_KEY); google.accounts.id.disableAutoSelect(); } catch (e) {}
        buildGate('signin');
        // Re-initialize GIS without auto_select so user sees the account picker
        try {
          google.accounts.id.initialize({ client_id: CLIENT_ID, callback: onCredential, auto_select: false, use_fedcm_for_prompt: true });
          google.accounts.id.prompt(function(n){ if(n.isNotDisplayed()||n.isSkippedMoment()) renderButton(); });
        } catch(e) { renderButton(); }
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

    // Guard: if a valid session exists within the 12-hour window for a DIFFERENT email,
    // discard this token — Chrome may have auto-selected a wrong secondary account.
    // We protect for the full 12h session window (not just while Google JWT is fresh),
    // so a token refresh mid-session can't sneak in the wrong account.
    try {
      var existing = storeGet(STORE_KEY);
      if (existing) {
        var ep  = parseJwt(existing);
        var exp = parseInt(storeGet(EXP_KEY) || '0');
        var withinSession = ep && Date.now() < exp; // within 12h window
        if (withinSession && String(ep.email||'').toLowerCase() !== String(p.email||'').toLowerCase()) {
          return; // wrong account — keep existing session, discard this token
        }
      }
    } catch(e) {}

    grant(token, p);
  }

  function grant(token, p) {
    window.IDS_TOKEN = token;
    window.IDS_USER = { email: String(p.email || '').toLowerCase(), name: p.name || p.email || '', picture: p.picture || '' };
    storeSet(STORE_KEY, token);
    storeSet(EXP_KEY, String(Date.now() + SESSION_MS));
    storeSet(HINT_KEY, String(p.email || '').toLowerCase()); // persist email — never deleted
    window.__idsDenied = false;
    // reveal page
    var s = document.getElementById('ids-lock-style'); if (s) s.remove();
    var g = document.getElementById('ids-gate'); if (g) g.remove();
    document.dispatchEvent(new CustomEvent('ids-auth', { detail: window.IDS_USER }));
    if (typeof window.idsOnAuth === 'function') { try { window.idsOnAuth(window.IDS_USER); } catch (e) {} }
  }

  // ── 4. Boot ──
  function boot() {
    // Reuse a stored token — localStorage means it's shared across all open tabs.
    // Only use it if not expired (keep 60s margin) AND within 12-hour session window.
    try {
      var saved = storeGet(STORE_KEY);
      if (saved) {
        var sp = parseJwt(saved);
        var nowMs = Date.now();
        var sessionExp = parseInt(storeGet(EXP_KEY) || '0');
        // Valid if: Google token not expired AND within 12-hour session window
        if (sp && sp.exp && sp.exp * 1000 > nowMs + 60000 && nowMs < sessionExp) {
          grant(saved, sp);
          return;
        }
        // Token expired — save email hint before deleting (fixes loginHint bug)
        if (sp && sp.email) storeSet(HINT_KEY, String(sp.email).toLowerCase());
        storeDel(STORE_KEY); storeDel(EXP_KEY);
      }
    } catch (e) {}

    // No valid stored token — try silent FedCM re-auth first, then show sign-in UI.
    // Read the persistent HINT_KEY (not the deleted STORE_KEY) so Chrome picks the
    // right account even after session expiry.
    var loginHint = storeGet(HINT_KEY) || '';

    buildGate('checking');
    loadGis(function () {
      try {
        var initOpts = { client_id: CLIENT_ID, callback: onCredential, auto_select: true, use_fedcm_for_prompt: true };
        if (loginHint) initOpts.login_hint = loginHint;
        google.accounts.id.initialize(initOpts);
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
    storeDel(STORE_KEY); storeDel(EXP_KEY);
    // NOTE: HINT_KEY intentionally kept — helps pick right account on next sign-in
    window.IDS_TOKEN = null;
    try { google.accounts.id.disableAutoSelect(); } catch (e) {}
    location.href = 'index.html';
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
