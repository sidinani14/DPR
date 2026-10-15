// reportday.js — which day a daily report (DPR / DPER / CRM) is for.
// Rule (2026-10-15, Siddharth): today's report, or the PREVIOUS working day's
// until 12 noon (so a report forgotten in the evening isn't lost). The server
// enforces it (allowedReportDay in Code.js) whatever the page sends; this
// shows the choice, the cut-off notices, and tells the server the form was
// opened (dailyFormOpened — emails Siddharth after noon if a day was missed).
// Usage: <script src="reportday.js" data-form="DPR"></script>; needs #f-date.
// Exposes: window.getReportDay(), window.REPORT_DAY (state),
//          window.reloadReportDay(), window.lockReportDateToToday() (re-render),
//          and fires a 'reportday-change' event when the day or status changes.
(function(){
  const FORM = (document.currentScript && document.currentScript.dataset.form) || 'DPR';
  const S = window.REPORT_DAY = { form: FORM, today: '', prev: '', windowOpen: false, graceHour: 12,
    filed: {}, edit: {}, loaded: false, choice: '' };

  function localToday(){ const n = new Date(); return n.getFullYear()+'-'+String(n.getMonth()+1).padStart(2,'0')+'-'+String(n.getDate()).padStart(2,'0'); }
  function pretty(d){ try { return new Date(d+'T12:00:00').toLocaleDateString('en-IN', {weekday:'short', day:'numeric', month:'short'}); } catch(e){ return d; } }
  function windowOpenNow(){ return S.loaded && S.windowOpen && !!S.prev && new Date().getHours() < S.graceHour; }
  function current(){ const t = S.today || localToday(); return (windowOpenNow() && S.choice === S.prev) ? S.prev : t; }
  window.getReportDay = current;
  function fire(){ try { document.dispatchEvent(new CustomEvent('reportday-change', { detail: { day: current() } })); } catch(e){} }

  function render(){
    const el = document.getElementById('f-date'); if (!el) return;
    if (!S.loaded) S.today = localToday();
    else if (localToday() !== S.today && !S._reloading) { S._reloading = true; load(); }   // tab left open past midnight
    const wasOpen = !!document.getElementById('rd-picker');
    const day = current();
    el.value = day; el.min = day; el.max = day; el.readOnly = true; el.style.opacity = '.75';
    el.title = 'Reports are for today — or, until 12 noon, the previous working day';
    let pk = document.getElementById('rd-picker');
    if (windowOpenNow()) {
      if (!pk) { pk = document.createElement('div'); pk.id = 'rd-picker'; pk.style.cssText = 'display:flex;flex-direction:column;gap:6px;margin-bottom:8px'; el.parentElement.insertBefore(pk, el); }
      const opt = (d, lbl) => '<label style="display:flex;align-items:center;gap:8px;padding:8px 10px;border:1px solid '+(d===day?'#1F3A5F':'#E2DFD8')+';border-radius:8px;cursor:pointer;font-size:13px;background:'+(d===day?'#EEF3F9':'#fff')+'">'
        + '<input type="radio" name="rd-day" value="'+d+'"'+(d===day?' checked':'')+' style="width:auto;margin:0">'
        + '<span style="flex:1">'+lbl+'</span>'
        + '<span style="font-size:11px;color:'+(S.filed[d]?'#1B5E20':'#8B2020')+'">'+(S.filed[d]?'submitted ✓':'not submitted')+'</span></label>';
      pk.innerHTML = opt(S.prev, 'For <b>'+pretty(S.prev)+'</b> <span style="color:#7A4A00">(until 12 noon)</span>') + opt(S.today, 'For <b>today</b>, '+pretty(S.today));
      pk.querySelectorAll('input').forEach(r => r.onchange = () => { S.choice = r.value; render(); fire(); });
    } else if (pk) pk.remove();
    notice(day);
    if (wasOpen && !windowOpenNow()) fire();   // noon passed while the page was open — back to today
  }

  function notice(day){
    const n = new Date(), mins = n.getHours()*60 + n.getMinutes();
    let el = document.getElementById('day-notice');
    const anchor = document.getElementById('f-date');
    if (!el && anchor) { el = document.createElement('div'); el.id = 'day-notice'; anchor.parentElement.appendChild(el); }
    if (!el) return;
    let msg = '', short = '', bg = '#FFF4E5', fg = '#7A4A00';
    if (windowOpenNow() && day === S.prev) {
      msg = '📅 You’re filing the report for ' + pretty(S.prev) + '. That’s allowed until 12 noon today; after that it counts as missed. Today’s report is separate — file it this evening as usual.';
      short = 'Filing the report for ' + pretty(S.prev) + ' — allowed until 12 noon today';
    } else if (windowOpenNow() && !S.filed[S.prev]) {
      msg = '⚠ Your report for ' + pretty(S.prev) + ' isn’t submitted. You can still file it until 12 noon — choose it above.';
      short = 'Your report for ' + pretty(S.prev) + ' is missing — you can file it until 12 noon';
      bg = '#FDECEA'; fg = '#8B2020';
    } else if (mins >= 19*60 + 30) {
      msg = '⏰ Past 7:30 PM — please submit tonight. If you forget, you can still file today’s report tomorrow until 12 noon; after that it counts as missed.';
      short = 'Past 7:30 PM — please submit tonight (catch-up allowed until 12 noon tomorrow)';
    }
    el.style.cssText = msg ? 'margin-top:8px;padding:9px 11px;border-radius:8px;font-size:12px;line-height:1.45;background:'+bg+';color:'+fg : 'display:none';
    el.textContent = msg;
    const top = document.getElementById('banner');   // dpr.html's top banner
    if (top) {   // only undo what this script set — dpr.html runs its own pre-7:30 countdown
      if (msg) { if (!top.dataset.dn) top.dataset.def = top.textContent; top.dataset.dn = '1'; top.className = 'banner late'; top.textContent = short; }
      else if (top.dataset.dn) { top.className = 'banner'; top.textContent = top.dataset.def || ''; delete top.dataset.dn; }
    }
  }

  async function load(){
    if (!(window.IDS_USER && window.IDS_USER.email)) { S._reloading = false; return; }
    try {
      const r = await fetch(APPS_SCRIPT_URL, {method:'POST', headers:{'Content-Type':'text/plain'},
        body: JSON.stringify({action:'getDailyFormStatus', form: FORM})}).then(x => x.json());
      if (r && r.status === 'ok') {
        const fresh = !S.loaded || S.today !== r.today;
        Object.assign(S, { today: r.today, prev: r.prev, windowOpen: !!r.windowOpen, graceHour: r.graceHour || 12,
          filed: r.filed || {}, edit: r.edit || {}, loaded: true });
        // Default to the missed day while it can still be filed.
        if (fresh) S.choice = (S.windowOpen && !S.filed[S.prev]) ? S.prev : S.today;
      }
    } catch(e) {}
    S._reloading = false;
    render(); fire();
  }
  window.reloadReportDay = load;
  window.lockReportDateToToday = render;   // existing call sites (draft restore, after submit)

  // Tell the server the form was opened — it emails Siddharth (12 noon–5 pm
  // only) if this person missed the previous working day.
  let pinged = false;
  function pingOpened(){
    if (pinged || !(window.IDS_USER && window.IDS_USER.email)) return;
    const h = new Date().getHours(); if (h < 12 || h >= 17) return;
    pinged = true;
    try { fetch(APPS_SCRIPT_URL, {method:'POST', headers:{'Content-Type':'text/plain'}, body: JSON.stringify({action:'dailyFormOpened', form: FORM})}).catch(()=>{}); } catch(e) {}
  }
  function onAuth(){ load(); pingOpened(); }
  document.addEventListener('ids-auth', onAuth);
  if (window.IDS_USER) setTimeout(onAuth, 0);
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', render); else render();
  setInterval(render, 60000);
})();
