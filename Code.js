// ═══════════════════════════════════════════════════════════════
// IDEAFORM DESIGN STUDIO — Apps Script Final
// Sheet structure verified against IDS_Productivity_System_2.xlsx
// Backend only — browser files (auth.js, *.html) are excluded via .claspignore.
// ═══════════════════════════════════════════════════════════════

var SHEET_ID          = '1PH1nJoPmQWS9wixuhw9B7oo0jIkhJw13htzbzlBWffk';  // Productivity sheet
var SCORECARD_SHEET_ID= '198sgwhnp2GY5KYyVITMZAPLEoiEuMrHIFUat5deG3vk'; // IDS Team Scorecard

// Google display name → TEAM tab name mapping
// Add any name mismatches here
var NAME_MAP = {
  'Astha Agrawal'   : 'Astha Inani',
  'Devenjana Patel' : 'Simi',
};

function resolveName(googleName) {
  var n = String(googleName||'').trim();
  return NAME_MAP[n] || n;
}

// ── Visit/Meeting type helpers ────────────────────────────────
var VISIT_TYPES_ARCH = ['Site Visit (Architecture)', 'Site Visit (Interiors)'];
var VISIT_TYPES_MTG  = ['Client Meeting (Regular)', 'Client Meeting'];

function isVisitTask(taskType) {
  var t = String(taskType||'').trim();
  return VISIT_TYPES_ARCH.indexOf(t) > -1 || VISIT_TYPES_MTG.indexOf(t) > -1 ||
         t === 'Site Visit' || t === 'Meeting';   // CRM-created meeting/visit tasks
}

function calcVisitPts(taskType, hours) {
  var t = String(taskType||'').trim();
  var h = parseFloat(hours) || 0;
  if (VISIT_TYPES_ARCH.indexOf(t) > -1 || t === 'Site Visit') return Math.round(h * 2 * 10) / 10;
  if (VISIT_TYPES_MTG.indexOf(t)  > -1 || t === 'Meeting')    return Math.round(h * 1 * 10) / 10;
  return 0;
}

// TASK_TAB / TASK_LOG removed — all tasks now in TASK_ASSIGNMENTS
var SUMMARY_TAB   = 'DAILY_SUMMARY';
var CONFIG_TAB    = 'CONFIG';
var PROJECTS_TAB  = 'PROJECTS';
var ASSIGN_TAB    = 'TASK_ASSIGNMENTS';
var BLOCK_LOG_TAB = 'BLOCK_LOG';   // audit of every block raised + lead disposition
var TEAM_TAB      = 'TEAM';
// TASK_ASSIGNMENTS appended block-workflow columns (1-indexed): X=24 prior status,
// Y=25 original deadline, Z=26 disposition (Pending/Rescheduled/Parked/Parked (Stalled)/Cancelled/Reassigned/Rejected),
// AA=27 park reason (free text for manual parks; auto-set for stalled). Added beyond the 23
// core cols, so existing index reads are unaffected. Review date lives in BLOCK_LOG.
var COL_BLK_PRIOR=24, COL_BLK_ORIGDL=25, COL_BLK_DISPO=26, COL_PARK_REASON=27;  // X,Y,Z,AA
var BILL_REQ_TAB  = 'BILL_REQUESTS';  // billables raised on DPR → approved → CRM raise-bill task
// Only Siddharth may park tasks directly from the dashboard (EPIC I).
var DIRECTOR_EMAILS = { 'sidinani14@gmail.com':1, 'siddharth@ideaform.in':1 };
// Approval form + weekly report are restricted to Siddharth & Astha only.
var MANAGER_EMAILS = { 'sidinani14@gmail.com':1, 'siddharth@ideaform.in':1, 'astha@ideaform.in':1, 'astha.uch@gmail.com':1 };
function isManager(email){ return !!MANAGER_EMAILS[String(email||'').toLowerCase()]; }
// Actions used ONLY by the approval form / weekly report (verified not shared
// with the dashboard) — callable by managers only.
var MANAGER_ONLY = { getWeeklyStats:1, getDeepakWeeklyStats:1, getAmanWeeklyStats:1,
  getPendingTasks:1, getBlockRequests:1, getMeetingApprovals:1, getBillRequests:1,
  submitApprovals:1, approveMeetingLog:1, disposeBillRequest:1, getWeeklyProjectDigest:1, getAllMeetingLogs:1 };
// EPIC K — unified Site Visit / Meeting log → AI-polished → lead-approved cumulative client PDF
var MEETING_LOG_TAB  = 'MEETING_LOG';   // one row per visit/meeting
var DECISION_LOG_TAB = 'DECISION_LOG';  // one row per action item
var LOGS_ROOT_FOLDER = 'IDS Logs';      // Drive root for per-project report folders
function ensureCols(sheet, n){ var m=sheet.getMaxColumns(); if (m < n) sheet.insertColumnsAfter(m, n-m); }
// Insert a new data row at the TOP (row 2, just under the header) so the latest
// entries are visible without scrolling. Use for human-reviewed log tabs only —
// NOT for tabs whose automation reads the just-written row via getLastRow().
function prependRow(sheet, values){
  ensureCols(sheet, values.length);
  sheet.insertRowBefore(2);
  var rng = sheet.getRange(2, 1, 1, sheet.getMaxColumns());
  // a freshly-inserted row inherits the header row's formatting — reset to plain
  rng.setBackground('#FFFFFF').setFontColor('#000000').setFontWeight('normal').setFontStyle('normal');
  sheet.getRange(2, 1, 1, values.length).setValues([values]);
  return 2;
}
function writeBlockLogHeaders(sheet){
  var h=['Block ID','Task ID','Assign Row','Member','Project','Task Type','Raised Date',
         'Reason','Prior Status','Original Deadline','Disposition','Reviewed By','Review Date','Note'];
  sheet.getRange(1,1,1,h.length).setValues([h]).setBackground('#8B2020').setFontColor('#FFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
}
var EXCLUDED_MEMBERS = ['Simi', 'Khushi Agrawal'];  // departed — hidden from all forms & dashboards
var SCORECARD_TAB = 'TEAM_SCORECARD';
var APPROVAL_FORM_URL = 'https://team.ideaformdesignstudio.com/approval.html';
var DAYS_BEFORE_ARCH  = 90;

// ── Helpers ───────────────────────────────────────────────────
function respond(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// CORS-safe respond — use this for doGet/doPost
function respondCORS(obj) {
  var output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function safeRespond(fn) {
  try { return respond(fn()); }
  catch(e) { return respond({error: e.toString()}); }
}

function db() { return SpreadsheetApp.openById(SHEET_ID); }
function sdb() { return SpreadsheetApp.openById(SCORECARD_SHEET_ID); }

function getOrCreate(name, hFn) {
  var s = db().getSheetByName(name);
  if (!s) { s = db().insertSheet(name); if (hFn) hFn(s); }
  else if (s.getLastRow() === 0 && hFn) hFn(s);
  return s;
}

function getOrCreateScorecard(name, hFn) {
  var s = sdb().getSheetByName(name);
  if (!s) { s = sdb().insertSheet(name); if (hFn) hFn(s); }
  else if (s.getLastRow() === 0 && hFn) hFn(s);
  return s;
}

function dateStr(d) {
  return (d ? new Date(d) : new Date()).toISOString().substring(0, 10);
}

// Convert a sheet cell value (Date object or string) to YYYY-MM-DD
// Google Sheets stores dates as Date objects - toISOString() shifts timezone
function cellDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    var y = v.getFullYear();
    var m = String(v.getMonth()+1).padStart(2,'0');
    var d = String(v.getDate()).padStart(2,'0');
    return y+'-'+m+'-'+d;
  }
  // Already a string - take first 10 chars
  return String(v).substring(0,10);
}

function nowStr() {
  return new Date().toLocaleString('en-IN', {
    day:'2-digit', month:'short', year:'numeric',
    hour:'2-digit', minute:'2-digit', hour12:true
  });
}

function mondayOf(date) {
  var d = new Date(date), day = d.getDay();
  d.setDate(d.getDate() + (day === 0 ? -6 : 1 - day));
  d.setHours(0, 0, 0, 0);
  return d;
}

// ════════════════════════════════════════════════════════════════
// ROUTING
// ════════════════════════════════════════════════════════════════

// Handle CORS preflight requests
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
// ACCESS CONTROL — every request must carry a Google ID token from an
// allowlisted Ideaform account. Verified via Google's tokeninfo endpoint
// (audience must match our OAuth client + email on the team list).
// ════════════════════════════════════════════════════════════════
var AUTH_CLIENT_ID = '48052407111-mantkqn708ejp5otfc34nch2ngl8o9ot.apps.googleusercontent.com';
var AUTH_ALLOWED = {
  'ar.deepaksoni@gmail.com':1, 'achal.rathore@ideaform.in':1, 'himanshu.malviya@ideaform.in':1,
  'poorvi.goyal@ideaform.in':1, 'sahil.maurya@ideaform.in':1, 'aaditya.koshti@ideaform.in':1,
  'aman.raghuwanshi@ideaform.in':1, 'bhavesh.bhagwat@ideaform.in':1, 'aashi.agrawal@ideaform.in':1,
  'siddharth@ideaform.in':1, 'sidinani14@gmail.com':1, 'astha@ideaform.in':1, 'astha.uch@gmail.com':1
};
// Single source of truth for access = the TEAM tab's Email column (active rows)
// UNIONED with the hardcoded directors list (fallback so a sheet read failure
// can never lock everyone out). Cached 5 min. Add someone to TEAM → they get in.
function getAllowedEmails() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('allowed_emails');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  var set = {};
  for (var k in AUTH_ALLOWED) set[k] = 1;            // directors fallback
  try {
    var tSheet = db().getSheetByName(TEAM_TAB);       // TEAM cols: …Email(E,4) Active(F,5)
    if (tSheet && tSheet.getLastRow() > 1) {
      var rows = tSheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var email  = String(rows[i][4] || '').trim().toLowerCase();
        var active = String(rows[i][5] || '').trim().toLowerCase();
        if (email && active !== 'no') set[email] = 1;
      }
    }
  } catch (e) {}
  cache.put('allowed_emails', JSON.stringify(set), 300);
  return set;
}
// Returns the verified allowlisted email, or null. Caches results (token is
// short-lived) to avoid an external fetch on every single API call.
function verifyIdToken(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var key = 'auth_' + Utilities.base64EncodeWebSafe(
              Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, token));
  var cached = cache.get(key);
  if (cached) return cached === '-' ? null : cached;
  try {
    var resp = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token),
      { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) { cache.put(key, '-', 300); return null; }
    var info  = JSON.parse(resp.getContentText());
    var email = String(info.email || '').toLowerCase();
    var ok = (info.aud === AUTH_CLIENT_ID) &&
             (String(info.email_verified) !== 'false') &&
             getAllowedEmails()[email];
    if (!ok) { cache.put(key, '-', 300); return null; }
    cache.put(key, email, 1800);  // 30 min (token itself expires ~1 hr)
    return email;
  } catch (err) { return null; }
}
function respondUnauthorized() {
  return respond({ status: 'error', code: 'unauthorized',
                   message: 'Sign in with an authorized Ideaform team account.' });
}
// ── ONE-TIME SETUP ──────────────────────────────────────────────
// Run this ONCE from the Apps Script editor (select authorizeNow → Run) and
// click "Allow" to grant the external-request permission the access check uses.
// Safe to run; it only pings Google's token-info endpoint. After it succeeds,
// the backend access gate can be deployed.
function authorizeNow() {
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=ping',
                            { muteHttpExceptions: true });
  Logger.log('External-request permission granted. tokeninfo HTTP ' + r.getResponseCode());
  return r.getResponseCode();
}
// EPIC K — run ONCE from the editor (select authorizeEpicK → Run, approve the
// Drive + external-request consent screen) so the Meeting Log can write to Drive
// and call the Anthropic API. Also reports whether the API key is set.
function authorizeEpicK() {
  var folder = (function(){ var it=DriveApp.getFoldersByName(LOGS_ROOT_FOLDER); return it.hasNext()?it.next():DriveApp.createFolder(LOGS_ROOT_FOLDER); })();
  UrlFetchApp.fetch('https://api.anthropic.com/v1/models', { muteHttpExceptions:true, headers:{'anthropic-version':'2023-06-01'} });
  var hasKey = !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  Logger.log('EPIC K authorized. Drive folder "'+LOGS_ROOT_FOLDER+'" ready: '+folder.getId()+'. ANTHROPIC_API_KEY set: '+hasKey);
  return { driveFolderId: folder.getId(), anthropicKeySet: hasKey };
}

function doGet(e) {
  // Handle CORS preflight — Apps Script handles this automatically
  // but we ensure JSON content type is always set
  var p = e && e.parameter ? e.parameter : {};
  // Pre-gate scope diagnostic (no sensitive data) — confirms the external_request
  // OAuth scope is authorized before we rely on token verification.
  if (p.action === 'authDiag') {
    try {
      var rr = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=ping', {muteHttpExceptions:true});
      return respond({ok:true, scope:'authorized', code:rr.getResponseCode()});
    } catch (er) { return respond({ok:false, scope:'NOT-authorized', error:String(er)}); }
  }
  // Pre-gate diagnostic: explains exactly why a real token is accepted/rejected.
  if (p.action === 'whoami') {
    var tk = p.idToken || '';
    var out = { hasToken: !!tk };
    if (!tk) return respond(out);
    try {
      var wr = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(tk), {muteHttpExceptions:true});
      out.tokeninfoHttp = wr.getResponseCode();
      var wi = {}; try { wi = JSON.parse(wr.getContentText()); } catch (e) { out.parseError = true; }
      var em = String(wi.email || '').toLowerCase();
      var allow = getAllowedEmails();
      out.email = em;
      out.aud = wi.aud;
      out.audExpected = AUTH_CLIENT_ID;
      out.audMatch = wi.aud === AUTH_CLIENT_ID;
      out.emailVerified = wi.email_verified;
      out.inAllowlist = !!allow[em];
      out.allowedCount = Object.keys(allow).length;
      out.wouldPass = out.audMatch && (String(wi.email_verified) !== 'false') && out.inAllowlist;
    } catch (e) { out.error = String(e); }
    return respond(out);
  }
  var authEmail = verifyIdToken(p.idToken);
  if (!authEmail) return respondUnauthorized();
  var action = p.action || '', member = p.member || '';
  if (MANAGER_ONLY[action] && !isManager(authEmail))
    return respond({ status:'error', code:'forbidden', message:'Restricted to Siddharth & Astha.' });
  // Lightweight access check used by the sign-in gate (auth.js)
  if (action === 'checkAccess') return respond({ allowed: true, email: authEmail });
  if (action === 'getLists')              return safeRespond(getLists);
  if (action === 'getConfig')             return safeRespond(readConfig);
  if (action === 'getPendingTasks')       return safeRespond(getPendingTasks);
  if (action === 'getAllTasks')           return safeRespond(getAllTasks);
  if (action === 'getOpenTasksForMember')    return safeRespond(function() { return getOpenTasksForMember(member); });
  if (action === 'getNotifications')         return safeRespond(function() { return getNotificationsForMember(member); });
  if (action === 'getWeeklyStats')           return safeRespond(function() { return getWeeklyStats(p.weekStart||''); });
  if (action === 'getProjectStats')          return safeRespond(getProjectStats);
  if (action === 'getDeepakVisitSummary')    return safeRespond(function() { return getDeepakVisitSummary(p.weekStart||''); });
  if (action === 'getCalendarData')          return safeRespond(getCalendarData);
  if (action === 'getDeepakWeeklyStats')     return safeRespond(function(){ return getDeepakWeeklyStats(p.weekStart||''); });
  if (action === 'getAmanWeeklyStats')       return safeRespond(function(){ return getAmanWeeklyStats(p.weekStart||''); });
  if (action === 'getLeadsAnalytics')        return safeRespond(function(){ return getLeadsAnalytics(p.month||''); });
  if (action === 'getBlockersThisWeek')      return safeRespond(getBlockersThisWeek);
  if (action === 'getBlockRequests')         return safeRespond(getBlockRequests);
  if (action === 'getBillRequests')          return safeRespond(getBillRequests);
  if (action === 'getMeetingApprovals')      return safeRespond(getMeetingApprovals);
  if (action === 'getMeetingTimeline')       return safeRespond(function(){ return getMeetingTimeline(p.project||''); });
  if (action === 'reconcileStalled')         return safeRespond(reconcileStalledParks);
  if (action === 'getProjectsHealth')        return safeRespond(getProjectsHealth);
  if (action === 'getWeeklyProjectDigest')   return safeRespond(function(){ return getWeeklyProjectDigest(p.weekStart||''); });
  if (action === 'getAllMeetingLogs')        return safeRespond(getAllMeetingLogs);
  if (action === 'getProjectDetail')         return safeRespond(function(){ return getProjectDetail(p.project||''); });
  if (action === 'getWeeklyDiag')            return safeRespond(function(){ return getWeeklyDiag(p.weekStart||''); });
  if (action === 'getOpenLeads')             return safeRespond(function(){ return getOpenLeads(p.member||''); });
  if (action === 'migrateAmanDaily')         return safeRespond(migrateAmanDailyToTabs);
  if (action === 'mergeCrmLog')              return safeRespond(mergeCrmLogTabs);
  if (action === 'testSiddharth')            return respond(testCreateSiddharthTask());
  return respond({status: 'IDS DPR live'});
}

function doPost(e) {
  try {
    var raw  = e && e.postData ? e.postData.contents : '';
    var data = JSON.parse(raw);
    var authEmail = verifyIdToken(data.idToken);
    if (!authEmail) return respondUnauthorized();
    if (MANAGER_ONLY[data.action] && !isManager(authEmail))
      return respond({ status:'error', code:'forbidden', message:'Restricted to Siddharth & Astha.' });
    Logger.log('doPost action: ' + data.action + ' | raw: ' + raw.substring(0,100));

    // GET-style actions sent via POST to bypass CORS
    if (data.action === 'getAllTasks')         return respond(getAllTasks());
    if (data.action === 'getLists')            return respond(getLists());
    if (data.action === 'getConfig')           return respond(readConfig());
    if (data.action === 'getPendingTasks')     return respond(getPendingTasks());
    if (data.action === 'getOpenTasksForMember') return respond(getOpenTasksForMember(data.member||''));
    if (data.action === 'getNotifications')    return respond(getNotificationsForMember(data.member||''));
    if (data.action === 'getWeeklyStats')         return respond(getWeeklyStats(data.weekStart||''));
    if (data.action === 'getDeepakWeeklyStats')   return respond(getDeepakWeeklyStats(data.weekStart||''));
    if (data.action === 'getAmanWeeklyStats')     return respond(getAmanWeeklyStats(data.weekStart||''));
    if (data.action === 'getLeadsAnalytics')      return respond(getLeadsAnalytics(data.month||''));
    if (data.action === 'getBlockersThisWeek')    return respond(getBlockersThisWeek());
    if (data.action === 'getBlockRequests')       return respond(getBlockRequests());
    if (data.action === 'getProjectsHealth')      return respond(getProjectsHealth());
    if (data.action === 'getWeeklyProjectDigest') return respond(getWeeklyProjectDigest(data.weekStart||''));
    if (data.action === 'getAllMeetingLogs')      return respond(getAllMeetingLogs());
    if (data.action === 'getProjectDetail')       return respond(getProjectDetail(data.project||''));
    if (data.action === 'disposeBlock')           return respond(disposeBlock(data));
    if (data.action === 'parkTask')               return respond(parkTask(data, authEmail));
    if (data.action === 'unparkTask')             return respond(unparkTask(data, authEmail));
    if (data.action === 'reconcileStalled')       return respond(reconcileStalledParks());
    if (data.action === 'getBillRequests')        return respond(getBillRequests());
    if (data.action === 'disposeBillRequest')     return respond(disposeBillRequest(data));
    if (data.action === 'getProjectStats')        return respond(getProjectStats());
    if (data.action === 'getDeepakVisitSummary')  return respond(getDeepakVisitSummary(data.weekStart||''));
    if (data.action === 'getSiteIssues')       return respond(getSiteIssues(data.project||''));
    if (data.action === 'resolveIssue')        return respond(resolveIssue(data.issueId||''));
    if (data.action === 'getSiteExecution')    return respond(getSiteExecutionSummary(data.project||''));
    if (data.action === 'submitDPER')          return respond(handleDPERSubmission(data));
    if (data.action === 'getCalendarData')     return respond(getCalendarData());
    if (data.action === 'getDeepakIssues')     return respond(getIssuesByReporter(data.lead||'Deepak Soni', true));
    if (data.action === 'getAmanIssues')       return respond(getIssuesByReporter(data.member||'Aman Raghuwanshi', false));
    if (data.action === 'updateIssueStatus')   return respond(updateIssueStatus(data.issueId||'', data.status||'', data.targetDate||''));
    if (data.action === 'submitAmanCRM')       return respond(submitAmanCRM(data));
    if (data.action === 'logConnections')      return respond(logConnections(data));
    if (data.action === 'submitBillables')     return respond(submitBillables(data));
    if (data.action === 'uploadMeetingPhoto')  return respond(uploadMeetingPhoto(data));
    if (data.action === 'submitMeetingLog')    return respond(submitMeetingLog(data, authEmail));
    if (data.action === 'getMeetingApprovals') return respond(getMeetingApprovals());
    if (data.action === 'approveMeetingLog')   return respond(approveMeetingLog(data, authEmail));
    if (data.action === 'finalizeMeetingLog')  return respond(finalizeMeetingLog(data, authEmail));
    if (data.action === 'deleteMeetingLog')    return respond(deleteMeetingLog(data, authEmail));
    if (data.action === 'getMeetingTimeline')  return respond(getMeetingTimeline(data.project||''));
    if (data.action === 'getRecentLeads')      return respond(getRecentLeads(data.date||''));
    if (data.action === 'getOpenLeads')        return respond(getOpenLeads(data.member||''));

    // Form submission actions
    if (data.action === 'submitApprovals')   return respond(submitApprovals(data));
    if (data.action === 'reassignTask')      return respond(reassignTask(data));
    if (data.action === 'assignTasks')       return respond(assignTasks(data));
    if (data.action === 'markNotifSeen')     return respond(markNotificationSeen(data));
    if (data.action === 'createSelfTask')    return respond(createSelfAssignedTask(data));
    if (data.action === 'createDoneTask')    return respond(createDoneTask(data));
    if (data.action === 'createSiddharthTask') return respond(createSiddharthTask(data));
    // Default: DPR submission
    var cfg = readConfig();
    var vPts = {};
    (cfg.visits || []).forEach(function(v) { vPts[v.duration] = v.pts; });
    // TASK_LOG no longer used for new tasks — Section 3 Done tasks
    // go to TASK_ASSIGNMENTS via createDoneTask action from DPR form
    // writeTaskLog kept for backward compat but not called for new submissions
    writeDailySummary(data);

    // Update assigned task statuses + write both date columns
    if (data['Assigned Statuses']) {
      try {
        var st = JSON.parse(data['Assigned Statuses']);
        updateTaskStatusesFromDPR(st, dateStr(data['Timestamp']));
      } catch(err) { Logger.log('Status update error: ' + err); }
    }

    // Section 3 done tasks → TASK_ASSIGNMENTS (replaces TASK_LOG)
    if (data['Done Tasks']) {
      try {
        var doneTasks = JSON.parse(data['Done Tasks']);
        doneTasks.forEach(function(t) { createDoneTask(t); });
        Logger.log('Created ' + doneTasks.length + ' done tasks in TASK_ASSIGNMENTS');
      } catch(err) { Logger.log('Done tasks error: ' + err); }
    }

    // Section 3 ongoing tasks → self-assigned, deadline = tomorrow
    if (data['Ongoing Tasks']) {
      try {
        var ongoing = JSON.parse(data['Ongoing Tasks']);
        ongoing.forEach(function(t) { createSelfAssignedTask(t); });
        Logger.log('Created ' + ongoing.length + ' self-assigned ongoing tasks');
      } catch(err) { Logger.log('Ongoing tasks error: ' + err); }
    }

    return respond({status: 'ok'});
  } catch(err) {
    return respond({status: 'error', message: err.toString()});
  }
}

// ════════════════════════════════════════════════════════════════
// getLists — team from TEAM tab, projects from PROJECTS tab
//
// TEAM tab cols: Team(A) Role(B) WeeklyTarget(C) DailyTarget(D) Email(E) Active(F)
// PROJECTS tab cols: ID(A) Name(B) Stage(C) Discipline(D) Multiplier(E) Lead(F)
//   Team(G) %Complete(H) OpenTasks(I) DelayedTasks(J) RAGStatus(K) Status(L) Notes(M)
// ════════════════════════════════════════════════════════════════
function getLists() {
  var s = db();

  // Team — two lists:
  // team = Active = Yes only (for DPR form member selector)
  // allMembers = everyone (for Task Assignment form assignee dropdowns)
  var tSheet = s.getSheetByName(TEAM_TAB);
  var team = [], emails = [], allMembers = [], allEmails = [];
  if (tSheet) {
    var tRows = tSheet.getDataRange().getValues();
    for (var i = 1; i < tRows.length; i++) {
      var name   = String(tRows[i][0] || '').trim();
      var email  = String(tRows[i][4] || '').trim();
      var active = String(tRows[i][5] || '').trim().toLowerCase();
      if (!name) continue;
      // Departed — removed from all forms & dashboards
      if (EXCLUDED_MEMBERS.indexOf(name) !== -1) continue;
      // All members regardless of active status
      allMembers.push(name);
      allEmails.push(email);
      // Active only — for DPR form
      if (active !== 'no') { team.push(name); emails.push(email); }
    }
  }

  // Projects — read name, discipline, multiplier
  // Multiplier col E contains a VLOOKUP formula; read computed value
  var pSheet = s.getSheetByName(PROJECTS_TAB);
  var projects = [];
  if (pSheet) {
    var pRows = pSheet.getDataRange().getValues();
    for (var j = 1; j < pRows.length; j++) {
      var pId   = String(pRows[j][0] || '').trim();  // A
      var pName = String(pRows[j][1] || '').trim();  // B
      var pDisc = String(pRows[j][3] || '').trim();  // D
      var pMult = parseFloat(pRows[j][4]) || 1.0;    // E (formula resolves to number)
      var pStat = String(pRows[j][2]  || '').trim(); // C Status (Ongoing/On Hold/Completed)
      if (pName) {
        var pLead = String(pRows[j][5] || '').trim(); // F Lead
        projects.push({ id:pId, name:pName, discipline:pDisc,
                        multiplier:pMult, status:pStat, lead:pLead });
      }
    }
  }

  return { team:team, emails:emails, allMembers:allMembers, projects:projects };
}

// ════════════════════════════════════════════════════════════════
// readConfig — reads all CONFIG tables
//
// CONFIG column positions (verified):
//   A,B,C  = Stage Weights (name, pts, category)
//   E,F    = Discipline Multipliers
//   H-K    = Stage % Weights (not used in script yet)
//   M,N    = Visit / Meeting Points
//   P,Q    = Scoring Weights + Revision Penalties
//   S,T,U  = Approval Leads
// ════════════════════════════════════════════════════════════════
function readConfig() {
  var sheet = db().getSheetByName(CONFIG_TAB);
  if (!sheet) return {error: 'CONFIG not found'};
  var rows = sheet.getDataRange().getValues();

  var stages = [], discs = [], visits = [], leads = [];
  var sw = {points:35, ontime:25, dpr:20, attendance:10, quality:10};
  var rp = {internal:1.5, client:0.75};
  var taskGroups = [];   // [{name, tasks:[labels]}] — from the — SECTION — headers in col A
  var curGroup = null;

  var SKIP_A = ['STAGE WEIGHTS','TASK / STAGE TYPE','TOTAL',''];
  var SKIP_E = ['DISCIPLINE MULTIPLIERS','DISCIPLINE & SCALE',''];
  var SKIP_M = ['VISIT / MEETING POINTS','DURATION',''];
  var SKIP_P = ['SCORING WEIGHTS','METRIC','TOTAL','REVISION PENALTIES','REVISION TYPE',''];
  var SKIP_S = ['APPROVAL LEADS','NAME',''];

  for (var i = 0; i < rows.length; i++) {
    var r  = rows[i];
    var a  = String(r[0]  || '').trim();
    var b  = r[1];
    var c  = String(r[2]  || '').trim();
    var e  = String(r[4]  || '').trim();
    var f  = String(r[5]  || '').replace('x','').trim();
    var m  = String(r[12] || '').trim();
    var n  = r[13];
    var p  = String(r[15] || '').trim();
    var q  = r[16];
    var sv = String(r[18] || '').trim();
    var tv = String(r[19] || '').trim();
    var uv = String(r[20] || '').trim();

    // Section header (— ARCHITECTURE —, — TOWNSHIP —, …) → starts a task-category group
    if (a.startsWith('—') && !a.startsWith('←')) {
      var gname = a.replace(/[—–-]/g, '').trim();
      if (gname && isNaN(parseFloat(gname))) { curGroup = {name: gname, tasks: []}; taskGroups.push(curGroup); }
    }

    if (a && SKIP_A.indexOf(a.toUpperCase()) === -1 && !a.startsWith('—') &&
        !a.startsWith('←') && !isNaN(parseFloat(b))) {
      stages.push({label:a, pts:parseFloat(b), category:c});
      if (curGroup) curGroup.tasks.push(a);
    }

    if (e && SKIP_E.indexOf(e.toUpperCase()) === -1 &&
        !e.startsWith('←') && !isNaN(parseFloat(f)))
      discs.push({label:e, multiplier:parseFloat(f)});

    if (m && SKIP_M.indexOf(m.toUpperCase()) === -1 &&
        !m.startsWith('←') && !isNaN(parseFloat(n)))
      visits.push({duration:m, pts:parseFloat(n)});

    if (p && SKIP_P.indexOf(p.toUpperCase()) === -1 &&
        !p.startsWith('←') && !isNaN(parseFloat(q))) {
      var pL = p.toLowerCase();
      // Scoring weights stored as decimals (0.35) — convert to %
      if      (pL.includes('points'))     sw.points     = parseFloat(q) <= 1 ? parseFloat(q)*100 : parseFloat(q);
      else if (pL.includes('on-time'))    sw.ontime     = parseFloat(q) <= 1 ? parseFloat(q)*100 : parseFloat(q);
      else if (pL.includes('dpr'))        sw.dpr        = parseFloat(q) <= 1 ? parseFloat(q)*100 : parseFloat(q);
      else if (pL.includes('attendance')) sw.attendance = parseFloat(q) <= 1 ? parseFloat(q)*100 : parseFloat(q);
      else if (pL.includes('quality'))    sw.quality    = parseFloat(q) <= 1 ? parseFloat(q)*100 : parseFloat(q);
      else if (pL.includes('internal'))   rp.internal   = parseFloat(q);
      else if (pL.includes('client'))     rp.client     = parseFloat(q);
    }

    if (sv && SKIP_S.indexOf(sv.toUpperCase()) === -1 &&
        !sv.startsWith('←') && tv) {
      leads.push({name:sv, email:tv, scope:uv || 'All projects'});
      Logger.log('Lead found: ' + sv + ' / ' + tv);
    }
  }

  // ── Payment stages (EPIC H) — CONFIG cols V(21)=Discipline, W(22)=Milestone ──
  // Under a "PAYMENT STAGES" marker. Tolerate blank spacer rows between
  // disciplines; skip the marker + the "Discipline"/"Milestone" header row.
  var paymentStages = {}, paymentDisciplines = [];
  for (var pj = 0; pj < rows.length; pj++) {
    var disc = String(rows[pj][21] || '').trim();   // col V
    var mile = String(rows[pj][22] || '').trim();   // col W
    if (!disc || !mile) continue;
    var dU = disc.toUpperCase();
    if (dU === 'PAYMENT STAGES' || dU === 'DISCIPLINE') continue;  // marker / header
    if (!paymentStages[disc]) { paymentStages[disc] = []; paymentDisciplines.push(disc); }
    paymentStages[disc].push(mile);
  }

  Logger.log('readConfig: found ' + leads.length + ' leads: ' + leads.map(function(l){return l.name;}).join(', '));
  return {stages:stages, disciplines:discs, visits:visits,
          leads:leads, scoringWeights:sw, revPenalties:rp, taskGroups:taskGroups,
          paymentStages:paymentStages, paymentDisciplines:paymentDisciplines};
}

// writeTaskLog and writeTaskLogHeaders removed — TASK_LOG tab is obsolete.
// All tasks (planned + unplanned) now written to TASK_ASSIGNMENTS.

// ════════════════════════════════════════════════════════════════
// DAILY SUMMARY
// Cols: A:Date B:Member C:E-mail D:Arrived on Time
//       E:Full List of Deliverables F:Anything Blocking G:Mood (1–5)
// ════════════════════════════════════════════════════════════════
function writeDailySummary(data) {
  var sheet = getOrCreate(SUMMARY_TAB, writeSummaryHeaders);
  // Time in HH:MM 24-hour format from Timestamp
  var ts    = data['Timestamp'] ? new Date(data['Timestamp']) : new Date();
  var hh    = String(ts.getHours()).padStart(2,'0');
  var mm    = String(ts.getMinutes()).padStart(2,'0');
  var timeStr = hh + ':' + mm;
  prependRow(sheet, [
    dateStr(data['Timestamp']),  // A Date
    timeStr,                      // B Time (HH:MM 24h)
    data['Member']            || '',  // C Member
    data['Member Email']      || '',  // D Email
    data['Arrived on Time']   || '',  // E Arrived on Time
    data['Deliverables List'] || '',  // F Deliverables
    data['Blockers']          || '',  // G Blockers
    data['Mood Score']        || '',  // H Mood
  ]);
}

function writeSummaryHeaders(s) {
  var h = ['Date','Time','Member','E-mail','Arrived on Time',
           'Full List of Deliverables','Anything Blocking','Mood (1–5)'];
  var r = s.getRange(1, 1, 1, h.length);
  r.setValues([h]);
  r.setBackground('#1F3A5F'); r.setFontColor('#FFFFFF');
  r.setFontWeight('bold'); r.setFontSize(10);
  s.setFrozenRows(1);
  [90,70,140,180,110,420,300,80]
    .forEach(function(w, i) { s.setColumnWidth(i+1, w); });
}

// ════════════════════════════════════════════════════════════════
// TASK ASSIGNMENTS — write from Task Assignment form
//
// TASK_ASSIGNMENTS cols (23 total — col P inserted for Actual Completion Date):
// A:TaskID B:ProjectID C:ProjectName D:AssignedTo E:Stage F:Discipline
// G:StageBasePts H:DiscMultiplier I:WeightedPoints J:AssignedDate K:Deadline
// L:Area M:DrawingName N:SelfStatus O:SelfStatusDate P:ActualCompletionDate
// Q:LeadApproved R:ApprovedBy S:ApprovalDate T:RevisionTag U:Notes
// V:AssignedBy W:Priority
// ════════════════════════════════════════════════════════════════
function assignTasks(data) {
  var sheet = getOrCreate(ASSIGN_TAB, writeAssignHeaders);
  var tasks = data.tasks || [];
  var today = dateStr();
  var written = 0;
  var stalledSet = stalledProjectSet();

  tasks.forEach(function(t) {
    var tUnits      = parseFloat(t.units) || 1;
    var tWeighted   = Math.round((t.basePts||0) * (data.multiplier||1) * tUnits * 10) / 10;
    sheet.appendRow([
      'T-' + Utilities.getUuid().substring(0,8).toUpperCase(), // A
      data.projectId    || '',        // B
      data.project      || '',        // C
      data.assignedTo   || '',        // D
      t.taskType        || '',        // E
      data.multiplier   || 1,         // F Disc. Multiplier (number)
      t.basePts         || 0,         // G Stage Base Pts
      tUnits,                         // H Units
      tWeighted,                      // I Weighted Pts = F × G × H
      data.dateAssigned || today,     // J
      t.targetDate      || '',        // K
      t.area            || '',        // L
      t.drawing         || '',        // M
      'Not Started',                  // N SelfStatus
      '',                             // O SelfStatusDate
      '',                             // P ActualCompletionDate
      'Pending',                      // Q LeadApproved
      '',                             // R ApprovedBy
      '',                             // S ApprovalDate
      '',                             // T RevisionTag
      t.notes           || '',        // U Notes
      data.assignedBy   || '',        // V AssignedBy
      t.priority        || 'Medium',  // W Priority
    ]);
    parkRowIfStalled(sheet, data.project || '', stalledSet);  // stalled project → auto-park
    written++;
  });
  return {status:'ok', written:written};
}

function writeAssignHeaders(s) {
  var h = ['Task ID','Project ID','Project Name','Assigned To','Stage','Disc. Multiplier',
           'Stage Base Pts','Units','Weighted Points','Assigned Date','Deadline',
           'Area','Drawing Name','Self Status','Self Done Date','Actual Completion Date',
           'Lead Approved','Approved By','Approval Date','Revision Tag','Notes','Assigned by','Priority'];
  var r = s.getRange(1, 1, 1, h.length);
  r.setValues([h]);
  r.setBackground('#1F3A5F'); r.setFontColor('#FFFFFF');
  r.setFontWeight('bold'); r.setFontSize(10);
  s.setFrozenRows(1);
  [90,80,180,140,220,180,90,90,90,100,90,120,160,100,100,100,120,100,90,200,120,80]
    .forEach(function(w, i) { s.setColumnWidth(i+1, w); });
}

// ════════════════════════════════════════════════════════════════
// GET ALL TASKS — for dashboard
// ════════════════════════════════════════════════════════════════
function getAllTasks() {
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) return {tasks:[]};
  var rows = sheet.getDataRange().getValues();
  var tasks = [], today = dateStr();

  // Build project → discipline text map from PROJECTS tab (col B=Name, col D=Discipline)
  var projSheet = db().getSheetByName(PROJECTS_TAB);
  var discMap = {};
  if (projSheet) {
    var pRows = projSheet.getDataRange().getValues();
    for (var pi = 1; pi < pRows.length; pi++) {
      var pn = String(pRows[pi][1]||'').trim();
      var pd = String(pRows[pi][3]||'').trim();
      if (pn) discMap[pn] = pd;
    }
  }

  // Auto-detect column structure based on header row
  // 22-col (old): LeadApproved=P(15) RevTag=S(18) Notes=T(19) AssignedBy=U(20) Priority=V(21)
  // 23-col (new): LeadApproved=Q(16) RevTag=T(19) Notes=U(20) AssignedBy=V(21) Priority=W(22)
  var headers = rows[0] ? rows[0].map(function(h){ return String(h||'').trim(); }) : [];
  var numCols = headers.length;
  var is23col = numCols >= 23 ||
    headers.indexOf('Actual Completion Date') > -1 ||
    (headers[15] && String(headers[15]).toLowerCase().includes('actual'));

  var COL_SELFSTATUS = 13;   // N — same in both
  var COL_TARGET     = 10;   // K — same in both
  var COL_AREA       = 11;   // L — same in both
  var COL_DRAWING    = 12;   // M — same in both
  var COL_STATUSDATE = 14;   // O — same in both (SelfStatusDate or SelfDoneDate)
  var COL_ACTUALDATE = is23col ? 15 : -1;  // P — only in 23-col
  var COL_LEADAPPR   = is23col ? 16 : 15;  // Q(16) new, P(15) old
  var COL_APPROVEDBY = is23col ? 17 : 16;  // R(17) new, Q(16) old
  var COL_REVTAG     = is23col ? 19 : 18;  // T(19) new, S(18) old
  var COL_NOTES      = is23col ? 20 : 19;  // U(20) new, T(19) old
  var COL_ASSIGNEDBY = is23col ? 21 : 20;  // V(21) new, U(20) old
  var COL_PRIORITY   = is23col ? 22 : 21;  // W(22) new, V(21) old

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var selfStatus   = String(r[COL_SELFSTATUS] || 'Not Started').trim();
    var leadApproved = String(r[COL_LEADAPPR]   || 'Pending').trim();
    var target       = cellDate(r[COL_TARGET]);
    var revTag       = String(r[COL_REVTAG] || '').trim();

    // Monday of current week — for completed task weekly filter
    var monOfWeek = cellDate(mondayOf(new Date()));

    var status = 'Upcoming';
    var blockDispo = String(r[COL_BLK_DISPO-1] || '').trim();   // Z disposition
    if (selfStatus === 'Work Not Done') {
      status = 'Work Not Done';  // lead closed it off the assignee (reassigned) — penalty applies
    }
    else if (selfStatus === 'Reassigned') {
      status = 'Hidden';         // original row of an approved-block reassignment — drop it
    }
    else if (selfStatus === 'Parked') {
      status = 'Parked';         // lead parked it — no deadline, off the workload map
    }
    else if (selfStatus === 'Blocked') {
      status = 'Blocked';        // pending lead approval (or kept as a cancelled record)
    }
    else if (selfStatus === 'Done') {
      if      (leadApproved === 'Yes') {
        // Completed — only show if completed this week (Mon onwards)
        var doneOn = (COL_ACTUALDATE > -1 ? cellDate(r[COL_ACTUALDATE]) : '') ||
                     cellDate(r[COL_STATUSDATE]) || today;
        status = doneOn >= monOfWeek ? 'Completed' : 'Hidden';
      }
      else if (leadApproved === 'No')      status = 'Rejected';
      else                                  status = 'Approval Pending'; // Pending
    }
    else if (revTag && selfStatus !== 'Done') status = 'Revision Required';
    else if (selfStatus === 'In Progress')    status = 'Ongoing';
    else if (target && target < today)        status = 'Delayed';

    if (status === 'Hidden') continue; // last week's completed tasks

    var proj = String(r[2] || '');
    tasks.push({
      row              : i+1,
      taskId           : String(r[0]  || ''),
      projectId        : String(r[1]  || ''),
      project          : proj,
      assignedTo       : String(r[3]  || ''),
      taskType         : String(r[4]  || ''),
      multiplier       : parseFloat(r[5]) || 1,       // F Disc. Multiplier (number)
      basePts          : parseFloat(r[6]) || 0,       // G Stage Base Pts
      units            : parseFloat(r[7]) || 1,       // H Units
      weightedPts      : parseFloat(r[8]) || 0,       // I Weighted Pts
      discipline       : discMap[proj]   || '',       // lookup from PROJECTS tab
      dateAssigned     : cellDate(r[9]),
      targetDate       : target,
      area             : String(r[COL_AREA]       || ''),
      drawing          : String(r[COL_DRAWING]    || ''),
      selfStatus       : selfStatus,
      selfStatusDate   : cellDate(r[COL_STATUSDATE]),
      actualCompletion : COL_ACTUALDATE > -1 ? cellDate(r[COL_ACTUALDATE]) : '',
      leadApproved     : leadApproved,
      approvedBy       : String(r[COL_APPROVEDBY] || ''),
      revisionTag      : revTag,
      notes            : String(r[COL_NOTES]      || ''),
      assignedBy       : String(r[COL_ASSIGNEDBY] || ''),
      priority         : String(r[COL_PRIORITY]   || 'Medium'),
      status           : status,
      approvalDate     : cellDate(r[is23col ? 18 : 17]),
      isRejected       : leadApproved === 'No',
      rejectionNote    : leadApproved === 'No' ? String(r[COL_NOTES]||'').trim() : '',
      blockDisposition : blockDispo,                  // ''|Pending|Cancelled|Parked|Parked (Stalled)|…
      parkReason       : String(r[COL_PARK_REASON-1] || ''),   // AA — why parked
      blockedDate      : (status === 'Blocked' || status === 'Parked') ? cellDate(r[COL_STATUSDATE]) : '',
    });
  }
  return {tasks:tasks};
}

// ════════════════════════════════════════════════════════════════
// GET OPEN TASKS FOR MEMBER — for DPR form section 2
// Returns Upcoming + Ongoing + Delayed + Revision Required
// ════════════════════════════════════════════════════════════════
function getOpenTasksForMember(member) {
  var resolved = resolveName(member);
  var all      = getAllTasks();
  var allTasks = all.tasks || [];

  // Debug: log all unique assignees in the sheet
  var allAssignees = {};
  allTasks.forEach(function(t) {
    allAssignees[t.assignedTo] = (allAssignees[t.assignedTo]||0) + 1;
  });
  Logger.log('getOpenTasksForMember called: google="'+member+'" resolved="'+resolved+'"');
  Logger.log('Total tasks from getAllTasks: '+allTasks.length);

  // Show first few tasks with their assignedTo and status for debugging
  var sample = allTasks.slice(0,5).map(function(t){
    return t.assignedTo+'|'+t.status;
  });
  Logger.log('Sample tasks (assignedTo|status): '+JSON.stringify(sample));

  // Show tasks assigned to resolved name
  var myTasks = allTasks.filter(function(t){
    return String(t.assignedTo||'').split(',').map(function(a){return a.trim();}).indexOf(resolved) > -1;
  });
  Logger.log('Tasks with assignedTo="'+resolved+'": '+myTasks.length);
  var myStatuses = myTasks.map(function(t){ return t.status; });
  Logger.log('Their statuses: '+JSON.stringify(myStatuses));

  var yesterday3 = addDaysToStr(dateStr(), -1);
  var tasks = myTasks.filter(function(t) {
    if (['Upcoming','Ongoing','Delayed','Revision Required','Approval Pending'].indexOf(t.status) > -1) return true;
    // Show rejected tasks for last 24h only
    if (t.status === 'Rejected' && t.approvalDate && t.approvalDate >= yesterday3) return true;
    return false;
  });

  Logger.log('Open tasks after status filter: '+tasks.length);
  return {
    tasks: tasks,
    debug: {
      googleName   : member,
      resolvedName : resolved,
      totalTasks   : allTasks.length,
      myTasks      : myTasks.length,
      myStatuses   : myStatuses,
      matchedTasks : tasks.length,
      sampleAssignees: allTasks.slice(0,5).map(function(t){return t.assignedTo;}),
    }
  };
}

// ════════════════════════════════════════════════════════════════
// UPDATE TASK STATUSES FROM DPR — 5-point match
// member + project + taskType + area + drawing
// TASK_ASSIGNMENTS: SelfStatus=N(14 1-idx), SelfDoneDate=O(15)
// ════════════════════════════════════════════════════════════════
function updateTaskStatusesFromDPR(statuses, submissionDate) {
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet || !statuses || !statuses.length) return;
  var rows = sheet.getDataRange().getValues();
  var now  = submissionDate || dateStr();

  statuses.forEach(function(a) {
    if (!a.taskType || !a.project || !a.member) return;
    var newStatus = '';
    if      (a.status === 'Done')    newStatus = 'Done';
    else if (a.status === 'Ongoing') newStatus = 'In Progress';
    else if (a.status === 'Blocked') newStatus = 'Blocked';
    else return; // Not Started — no change needed

    for (var i = 1; i < rows.length; i++) {
      var rM = String(rows[i][3]  || '').trim(); // D AssignedTo
      var rP = String(rows[i][2]  || '').trim(); // C ProjectName
      var rT = String(rows[i][4]  || '').trim(); // E Stage
      var rA = String(rows[i][11] || '').trim(); // L Area
      var rD = String(rows[i][12] || '').trim(); // M DrawingName
      var rS = String(rows[i][13] || '').trim(); // N SelfStatus

      var match = rM === a.member && rP === a.project && rT === a.taskType &&
                  (!a.area    || !rA || rA === a.area) &&
                  (!a.drawing || !rD || rD === a.drawing) &&
                  rS !== 'Done';

      if (match) {
        // ── Block raised: snapshot prior state + log it (don't re-snapshot if already pending) ──
        if (newStatus === 'Blocked' && String(rows[i][COL_BLK_DISPO-1]||'').trim() !== 'Pending') {
          ensureCols(sheet, COL_BLK_DISPO);
          var priorStatus  = rS || 'Not Started';
          var origDeadline = cellDate(rows[i][10]); // K Deadline (keeps running until approved)
          sheet.getRange(i+1, COL_BLK_PRIOR ).setValue(priorStatus);
          sheet.getRange(i+1, COL_BLK_ORIGDL).setValue(origDeadline);
          sheet.getRange(i+1, COL_BLK_DISPO ).setValue('Pending');
          logBlockRaised(String(rows[i][0]||''), i+1, a.member, a.project, a.taskType, now, a.note||'', priorStatus, origDeadline);
        }
        sheet.getRange(i+1, 14).setValue(newStatus); // N SelfStatus
        sheet.getRange(i+1, 15).setValue(now);       // O SelfStatusDate always

        if (newStatus === 'Done') {
          // P ActualCompletionDate
          var actualDate = a.actualCompletionDate || now;
          sheet.getRange(i+1, 16).setValue(actualDate);
          // Re-enter the approval queue (clears a prior rejection so redone work
          // returns to 'Approval Pending' instead of staying rejected)
          sheet.getRange(i+1, 17).setValue('Pending'); // Q LeadApproved

          // If visit/meeting task — overwrite WeightedPoints (col I) with hours-based pts
          var taskType = String(rows[i][4]||'').trim(); // E
          if (isVisitTask(taskType) && a.visitHours) {
            var visitPts = calcVisitPts(taskType, a.visitHours);
            if (visitPts > 0) {
              sheet.getRange(i+1, 9).setValue(visitPts);  // I WeightedPoints
              sheet.getRange(i+1, 7).setValue(a.visitHours); // G StageBasePts = hours
              Logger.log('Visit pts updated: ' + taskType + ' ' + a.visitHours + 'h = ' + visitPts + 'pts');
            }
          }
        }

        // Note
        if (a.note) {
          var existNote = String(sheet.getRange(i+1, 21).getValue() || '').trim(); // U Notes
          var newNote   = now + ': ' + a.note;
          sheet.getRange(i+1, 21).setValue(existNote ? existNote + ' | ' + newNote : newNote);
        }
        Logger.log('Status updated: ' + a.member + ' / ' + a.taskType + ' → ' + newStatus);
        break;
      }
    }
  });
}

// ════════════════════════════════════════════════════════════════
// GET PENDING TASKS — for approval form
// TASK_LOG cols: SubID(A=0) Date(B=1) Member(C=2) Project(D=3) Disc(E=4)
//   TaskType(F=5) Area(G=6) Drawing(H=7) Units(I=8) Pts(J=9)
//   LeadApproved(K=10) ApprovedBy(L=11) ReviewedOn(M=12) Notes(N=13)
// ════════════════════════════════════════════════════════════════
function getPendingTasks() {
  var tasks = [];

  // ── TASK_ASSIGNMENTS — only source: assigned tasks marked Done, not yet approved ──
  // TASK_LOG removed — obsolete. All tasks are now in TASK_ASSIGNMENTS.
  // and marked Done via DPR Section 2. Need separate approval.
  // TASK_ASSIGNMENTS cols (23 total after col P insertion):
  //   TaskID(A=0) ProjectID(B=1) Project(C=2) AssignedTo(D=3)
  //   Stage(E=4) Discipline(F=5) BasePts(G=6) Multiplier(H=7) WeightedPts(I=8)
  //   AssignedDate(J=9) Deadline(K=10) Area(L=11) Drawing(M=12)
  //   SelfStatus(N=13) SelfStatusDate(O=14) ActualCompletionDate(P=15)
  //   LeadApproved(Q=16) ApprovedBy(R=17) ApprovalDate(S=18)
  //   RevisionTag(T=19) Notes(U=20) AssignedBy(V=21) Priority(W=22)
  var asSheet = db().getSheetByName(ASSIGN_TAB);
  if (asSheet) {
    var asRows = asSheet.getDataRange().getValues();
    // Auto-detect column structure (same logic as getAllTasks)
    var asHeaders = asRows[0] ? asRows[0].map(function(h){ return String(h||'').trim(); }) : [];
    var asIs23 = asHeaders.length >= 23 || asHeaders.indexOf('Actual Completion Date') > -1;
    var AS_LEADAPPR = asIs23 ? 16 : 15; // Q(16) new, P(15) old
    var AS_ACTUALDT = asIs23 ? 15 : -1;
    var AS_REVTAG   = asIs23 ? 19 : 18;
    var AS_NOTES    = asIs23 ? 20 : 19;

    for (var j = 1; j < asRows.length; j++) {
      var selfStatus   = String(asRows[j][13]          || '').trim(); // N always
      var leadApproved = String(asRows[j][AS_LEADAPPR]  || '').trim();
      // Show all Done tasks not yet approved or rejected (blank or Pending)
      if (selfStatus !== 'Done') continue;
      if (leadApproved === 'Yes' || leadApproved === 'No') continue;
      var doneDate = (AS_ACTUALDT > -1 ? cellDate(asRows[j][AS_ACTUALDT]) : '') ||
                     cellDate(asRows[j][14]) || dateStr(); // P or O
      tasks.push({
        row          : j + 1,
        source       : 'TASK_ASSIGNMENTS',
        subId        : String(asRows[j][0]  || ''), // A TaskID
        date         : doneDate,
        member       : String(asRows[j][3]  || ''), // D AssignedTo
        project      : String(asRows[j][2]  || ''), // C Project
        discipline   : String(asRows[j][5]  || ''), // F Discipline
        taskType     : String(asRows[j][4]  || ''), // E Stage
        area         : String(asRows[j][11] || ''), // L Area
        drawing      : String(asRows[j][12] || ''), // M Drawing
        units        : asRows[j][8]  || 0,           // I WeightedPts
        pts          : asRows[j][8]  || 0,           // I WeightedPts
        taskAssignRow: j + 1,
        revisionTag  : String(asRows[j][AS_REVTAG] || ''),
        notes        : String(asRows[j][AS_NOTES]  || ''),
      });
    }
  }


  return {tasks: tasks};
}
// findTaskAssignRow removed — was only used by TASK_LOG approval path (now obsolete)

// ════════════════════════════════════════════════════════════════
// BLOCK WORKFLOW — a Blocked task is a request the lead approves/rejects.
// ════════════════════════════════════════════════════════════════
function logBlockRaised(taskId, assignRow, member, project, taskType, raisedDate, reason, priorStatus, origDeadline){
  var sheet = getOrCreate(BLOCK_LOG_TAB, writeBlockLogHeaders);
  sheet.appendRow([ nextId(sheet,'BLK-'), taskId, assignRow, member, project, taskType,
    raisedDate, reason||'', priorStatus||'', origDeadline||'', 'Pending', '', '', '' ]);
}
function updateBlockLog(taskId, dispoLabel, reviewedBy, reviewDate, note){
  var sheet = db().getSheetByName(BLOCK_LOG_TAB);
  if (!sheet || sheet.getLastRow() < 2) return;
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length-1; i >= 1; i--){   // most-recent pending row for this task
    if (String(rows[i][1]||'') === taskId && String(rows[i][10]||'').trim() === 'Pending'){
      sheet.getRange(i+1,11).setValue(dispoLabel);
      sheet.getRange(i+1,12).setValue(reviewedBy||'');
      sheet.getRange(i+1,13).setValue(reviewDate||'');
      sheet.getRange(i+1,14).setValue(note||'');
      break;
    }
  }
}
// Tasks awaiting a lead's block decision — for the approval form
function getBlockRequests(){
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet || sheet.getLastRow() < 2) return {requests:[]};
  var rows = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++){
    if (String(rows[i][13]||'').trim() !== 'Blocked') continue;
    if (String(rows[i][COL_BLK_DISPO-1]||'').trim() !== 'Pending') continue;
    out.push({
      row:i+1, taskId:String(rows[i][0]||''), member:String(rows[i][3]||''),
      project:String(rows[i][2]||''), taskType:String(rows[i][4]||''),
      area:String(rows[i][11]||''), drawing:String(rows[i][12]||''),
      weightedPts:parseFloat(rows[i][8])||0,
      priorStatus:String(rows[i][COL_BLK_PRIOR-1]||''),
      origDeadline:cellDate(rows[i][COL_BLK_ORIGDL-1]),
      deadline:cellDate(rows[i][10]),
      raisedDate:cellDate(rows[i][14]),
      reason:latestNoteText(rows[i][20])
    });
  }
  return {requests:out};
}
// Lead disposition of a block: reschedule | park | cancel | reassign | reject
function disposeBlock(data){
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) return {status:'error', message:'TASK_ASSIGNMENTS not found'};
  var row = parseInt(data.row,10);
  if (!row) return {status:'error', message:'no row'};
  ensureCols(sheet, COL_BLK_DISPO);
  var r = sheet.getRange(row,1,1,Math.max(sheet.getLastColumn(),COL_BLK_DISPO)).getValues()[0];
  var st = String(r[13]||'').trim(), dp = String(r[COL_BLK_DISPO-1]||'').trim();
  // Accept a pending block, OR a parked task being reassigned from the dashboard
  if (!((st === 'Blocked' && dp === 'Pending') || st === 'Parked'))
    return {status:'error', message:'not a pending block or parked task'};
  var dispo = String(data.disposition||'').trim();
  var reviewedBy = data.reviewedBy || '';
  var today = dateStr();
  var priorStatus  = String(r[COL_BLK_PRIOR-1]||'').trim() || 'In Progress';
  var origDeadline = cellDate(r[COL_BLK_ORIGDL-1]);
  var label, note = data.note || '';

  if (dispo === 'reschedule'){
    sheet.getRange(row,14).setValue(priorStatus);
    sheet.getRange(row,11).setValue(data.newDeadline || origDeadline);
    sheet.getRange(row,COL_BLK_DISPO).setValue('Rescheduled'); label='Rescheduled';
  } else if (dispo === 'park'){
    sheet.getRange(row,14).setValue('Parked');
    sheet.getRange(row,11).setValue('');                        // no deadline
    sheet.getRange(row,COL_BLK_DISPO).setValue('Parked'); label='Parked';
  } else if (dispo === 'cancel'){
    sheet.getRange(row,14).setValue('Blocked');                 // keep as record
    sheet.getRange(row,COL_BLK_DISPO).setValue('Cancelled'); label='Cancelled';
  } else if (dispo === 'reassign'){
    if (!data.newAssignee) return {status:'error', message:'newAssignee required'};
    sheet.getRange(row,14).setValue('Reassigned');              // original released (no penalty)
    sheet.getRange(row,COL_BLK_DISPO).setValue('Reassigned'); label='Reassigned';
    var n = r.slice(0,23);
    n[0]=nextId(sheet,'T-'); n[3]=data.newAssignee; n[9]=today;
    n[10]=data.newDeadline || origDeadline; n[13]='Not Started';
    n[14]=''; n[15]=''; n[16]='Pending'; n[17]=''; n[18]=''; n[19]='';
    n[20]='Reassigned from '+String(r[3]||'')+' (approved block)'; n[21]=reviewedBy+' (block reassign)';
    sheet.appendRow(n);
  } else if (dispo === 'reject'){
    sheet.getRange(row,14).setValue(priorStatus);
    sheet.getRange(row,11).setValue(origDeadline);              // original deadline stands
    sheet.getRange(row,COL_BLK_DISPO).setValue('Rejected'); label='Rejected';
  } else {
    return {status:'error', message:'unknown disposition: '+dispo};
  }
  if (note){
    var ex = String(sheet.getRange(row,21).getValue()||'').trim();
    sheet.getRange(row,21).setValue(ex ? ex+' | '+today+' (lead): '+note : today+' (lead): '+note);
  }
  updateBlockLog(String(r[0]||''), label, reviewedBy, today, note);
  return {status:'ok', disposition:label};
}
// Reliability penalty: rejected blocks this week, per member (−1 each)
function rejectedBlocksThisWeek(name){
  var sheet = db().getSheetByName(BLOCK_LOG_TAB);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var mon = dateStr(mondayOf(new Date()));
  var sun = new Date(mondayOf(new Date())); sun.setDate(sun.getDate()+6);
  var sunStr = dateStr(sun);
  var rows = sheet.getDataRange().getValues(), c = 0;
  for (var i = 1; i < rows.length; i++){
    if (String(rows[i][3]||'').trim() !== name) continue;
    if (String(rows[i][10]||'').trim() !== 'Rejected') continue;
    var rd = dateStr(rows[i][12]);
    if (rd >= mon && rd <= sunStr) c++;
  }
  return c;
}
// Per-member block stats for a given week (for the weekly scorecard line).
// Counts only member-RAISED blocks (Pending → a real disposition); direct
// dashboard parks (logged as "Parked (Direct)") and stalled auto-parks are
// excluded so they don't inflate someone's "blocks raised".
function blockStatsForWeek(name, mon, sat){
  var out = {raised:0, approved:0, rejected:0, pending:0};
  var sheet = db().getSheetByName(BLOCK_LOG_TAB);
  if (!sheet || sheet.getLastRow() < 2) return out;
  var APPROVED = {'Rescheduled':1,'Reassigned':1,'Parked':1,'Cancelled':1};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++){
    if (String(rows[i][3]||'').trim() !== name) continue;     // Member
    var rd = cellDate(rows[i][6]);                            // Raised Date
    if (!rd || rd < mon || rd > sat) continue;
    var dispo = String(rows[i][10]||'').trim();               // Disposition
    if (dispo === '' || dispo === 'Pending'){ out.raised++; out.pending++; }
    else if (dispo === 'Rejected'){ out.raised++; out.rejected++; }
    else if (APPROVED[dispo]){ out.raised++; out.approved++; }
    // else: "Parked (Direct)" / other non-member-raised audit rows → skip
  }
  return out;
}

// ════════════════════════════════════════════════════════════════
// EPIC I — Park / un-park a task directly from the dashboard (Siddharth only).
// Mirrors the approval-form park, but needs no prior Block. Snapshots the
// prior status + deadline (X/Y), sets status Parked (N) with no deadline (K),
// disposition Parked (Z) and a free-text reason (AA), and audits to BLOCK_LOG.
// ════════════════════════════════════════════════════════════════
function isDirector(email){ return !!DIRECTOR_EMAILS[String(email||'').toLowerCase()]; }

function parkTask(data, authEmail){
  if (!isDirector(authEmail)) return {status:'error', code:'forbidden', message:'Only Siddharth can park tasks.'};
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) return {status:'error', message:'TASK_ASSIGNMENTS not found'};
  var row = parseInt(data.row,10);
  if (!row) return {status:'error', message:'no row'};
  ensureCols(sheet, COL_PARK_REASON);
  var r = sheet.getRange(row,1,1,Math.max(sheet.getLastColumn(),COL_PARK_REASON)).getValues()[0];
  var st = String(r[13]||'').trim();
  if (st === 'Parked') return {status:'ok', already:true};
  if (['Done','Reassigned','Work Not Done'].indexOf(st) > -1)
    return {status:'error', message:'cannot park a '+st.toLowerCase()+' task'};
  var today = dateStr();
  if (!String(r[COL_BLK_PRIOR-1]||'').trim()){
    sheet.getRange(row, COL_BLK_PRIOR ).setValue(st || 'In Progress');
    sheet.getRange(row, COL_BLK_ORIGDL).setValue(cellDate(r[10]) || '');
  }
  sheet.getRange(row,14).setValue('Parked');                       // N selfStatus
  sheet.getRange(row,11).setValue('');                             // K no deadline
  sheet.getRange(row,COL_BLK_DISPO).setValue('Parked');            // Z disposition
  sheet.getRange(row,COL_PARK_REASON).setValue(data.reason || ''); // AA reason
  var lg = getOrCreate(BLOCK_LOG_TAB, writeBlockLogHeaders);
  lg.appendRow([ nextId(lg,'BLK-'), String(r[0]||''), row, String(r[3]||''),
    String(r[2]||''), String(r[4]||''), today, data.reason||'',
    String(r[COL_BLK_PRIOR-1]||st||''), cellDate(r[COL_BLK_ORIGDL-1])||'',
    'Parked (Direct)', data.by||authEmail, today, data.reason||'' ]);
  return {status:'ok'};
}

function unparkTask(data, authEmail){
  if (!isDirector(authEmail)) return {status:'error', code:'forbidden', message:'Only Siddharth can un-park tasks.'};
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) return {status:'error', message:'TASK_ASSIGNMENTS not found'};
  var row = parseInt(data.row,10);
  if (!row) return {status:'error', message:'no row'};
  ensureCols(sheet, COL_PARK_REASON);
  var r = sheet.getRange(row,1,1,Math.max(sheet.getLastColumn(),COL_PARK_REASON)).getValues()[0];
  if (String(r[13]||'').trim() !== 'Parked') return {status:'error', message:'not parked'};
  var prior = String(r[COL_BLK_PRIOR-1]||'').trim() || 'In Progress';
  sheet.getRange(row,14).setValue(prior);                          // N restore
  sheet.getRange(row,11).setValue(cellDate(r[COL_BLK_ORIGDL-1])||''); // K restore deadline
  sheet.getRange(row,COL_BLK_DISPO).setValue('');                  // Z clear
  sheet.getRange(row,COL_PARK_REASON).setValue('');                // AA clear
  sheet.getRange(row,COL_BLK_PRIOR).setValue('');
  sheet.getRange(row,COL_BLK_ORIGDL).setValue('');
  return {status:'ok', restoredTo:prior};
}

// ════════════════════════════════════════════════════════════════
// EPIC J — "Stalled" project stage. Stalled = paused, no client response.
// Open tasks on a stalled project are auto-parked with disposition
// "Parked (Stalled)"; when the project un-stalls, only those revive.
// ════════════════════════════════════════════════════════════════
function stalledProjectSet(){
  var set = {};
  var pSheet = db().getSheetByName(PROJECTS_TAB);
  if (pSheet && pSheet.getLastRow() > 1){
    var pr = pSheet.getDataRange().getValues();
    for (var i=1;i<pr.length;i++){
      if (String(pr[i][2]||'').trim().toLowerCase() === 'stalled'){
        var nm = String(pr[i][1]||'').trim().toLowerCase();
        if (nm) set[nm] = 1;
      }
    }
  }
  return set;
}
// Park a just-appended task row if its project is currently stalled.
function parkRowIfStalled(sheet, projectName, stalledSet){
  if (!projectName) return false;
  var set = stalledSet || stalledProjectSet();
  if (!set[String(projectName).trim().toLowerCase()]) return false;
  var row = sheet.getLastRow();
  ensureCols(sheet, COL_PARK_REASON);
  var prior = String(sheet.getRange(row,14).getValue()||'').trim() || 'Not Started';
  if (prior === 'Parked') return false;
  sheet.getRange(row,COL_BLK_PRIOR ).setValue(prior);
  sheet.getRange(row,COL_BLK_ORIGDL).setValue(cellDate(sheet.getRange(row,11).getValue())||'');
  sheet.getRange(row,14).setValue('Parked');
  sheet.getRange(row,11).setValue('');
  sheet.getRange(row,COL_BLK_DISPO).setValue('Parked (Stalled)');
  sheet.getRange(row,COL_PARK_REASON).setValue('Project stalled — no client response');
  return true;
}
// Reconcile every task against the current stalled-project list (park new,
// revive un-stalled). Runs from the Monday trigger + callable on demand.
function reconcileStalledParks(){
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet || sheet.getLastRow() < 2) return {status:'ok', parked:0, unparked:0};
  ensureCols(sheet, COL_PARK_REASON);
  var stalled = stalledProjectSet();
  var rows = sheet.getDataRange().getValues();
  var parked=0, unparked=0;
  var SKIP = ['Done','Reassigned','Work Not Done'];
  for (var i=1;i<rows.length;i++){
    var r = rows[i];
    var proj  = String(r[2]||'').trim().toLowerCase();
    var st    = String(r[13]||'').trim();
    var dispo = String(r[COL_BLK_DISPO-1]||'').trim();
    var isStalled = !!stalled[proj];
    if (isStalled && st !== 'Parked' && SKIP.indexOf(st) === -1){
      if (!String(r[COL_BLK_PRIOR-1]||'').trim()){
        sheet.getRange(i+1,COL_BLK_PRIOR ).setValue(st || 'In Progress');
        sheet.getRange(i+1,COL_BLK_ORIGDL).setValue(cellDate(r[10])||'');
      }
      sheet.getRange(i+1,14).setValue('Parked');
      sheet.getRange(i+1,11).setValue('');
      sheet.getRange(i+1,COL_BLK_DISPO).setValue('Parked (Stalled)');
      sheet.getRange(i+1,COL_PARK_REASON).setValue('Project stalled — no client response');
      parked++;
    } else if (!isStalled && st === 'Parked' && dispo === 'Parked (Stalled)'){
      var prior = String(r[COL_BLK_PRIOR-1]||'').trim() || 'In Progress';
      sheet.getRange(i+1,14).setValue(prior);
      sheet.getRange(i+1,11).setValue(cellDate(r[COL_BLK_ORIGDL-1])||'');
      sheet.getRange(i+1,COL_BLK_DISPO).setValue('');
      sheet.getRange(i+1,COL_PARK_REASON).setValue('');
      sheet.getRange(i+1,COL_BLK_PRIOR).setValue('');
      sheet.getRange(i+1,COL_BLK_ORIGDL).setValue('');
      unparked++;
    }
  }
  return {status:'ok', parked:parked, unparked:unparked};
}

// Installable onEdit handler — fires the instant a PROJECTS stage cell (col C)
// is edited, so setting a project to "Stalled" auto-parks its tasks immediately
// (no wait for the Monday sync). Reconcile is idempotent; wrapped so an edit
// never fails. (This is a standalone script, so the trigger must be installed
// once via setupStalledTrigger() — a simple onEdit would not fire here.)
function onProjectsStageEdit(e){
  try {
    if (!e || !e.range) return;
    if (e.range.getSheet().getName() !== PROJECTS_TAB) return;
    if (e.range.getColumn() !== 3) return;     // col C = Stage
    reconcileStalledParks();
  } catch (err) { /* never block the user's edit */ }
}
// Run ONCE from the Apps Script editor (select setupStalledTrigger → Run) to
// install the edit trigger. Removes any prior copy so it's safe to re-run.
function setupStalledTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'onProjectsStageEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onProjectsStageEdit')
    .forSpreadsheet(SHEET_ID).onEdit().create();
  return 'onProjectsStageEdit trigger installed for PROJECTS edits.';
}

// ════════════════════════════════════════════════════════════════
// EPIC G — shared connection logger. DPR + DPER call this once to append
// client connections to CRM_LOG (same schema as the CRM form). The author
// is the signed-in DPR/DPER member. Visibility-only — does not score.
// ════════════════════════════════════════════════════════════════
function appendConnections(member, day, connections, subId){
  if (!member || !connections || !connections.length) return 0;
  var logSheet = getOrCreate(CRM_LOG_TAB, writeCrmLogHeaders);
  var n = 0;
  connections.forEach(function(c){
    if (!c || (!c.project && !c.type && !c.notes)) return;
    prependRow(logSheet, [ nextId(logSheet,'LOG-'), subId||'', day, member,
      'Client Connection', c.type||'', c.project||'', c.notes||'' ]);
    n++;
  });
  return n;
}
function logConnections(data){
  var conns = data.connections;
  if (typeof conns === 'string'){ try { conns = JSON.parse(conns); } catch(e){ conns = []; } }
  var n = appendConnections(data.member||'', dateStr(data.date||''), conns||[]);
  return {status:'ok', written:n};
}

// ════════════════════════════════════════════════════════════════
// EPIC H — billables raised on DPR → BILL_REQUESTS (Pending) → approval
// panel approves → 0-pt "Raise Bill" CRM task. No points / no SLA on that
// task (billing timing is outside CRM's control), so it is created with
// no deadline and 0 weighted pts → never counts as delayed or in output.
// ════════════════════════════════════════════════════════════════
function writeBillReqHeaders(sheet){
  var h = ['Request ID','Date','Raised By','Project','Discipline','Stage',
           'Status','Approved By','Approval Date','Task ID','Note'];
  sheet.getRange(1,1,1,h.length).setValues([h]).setBackground('#1F3A5F').setFontColor('#FFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
}
function submitBillables(data){
  var bills = data.billables;
  if (typeof bills === 'string'){ try { bills = JSON.parse(bills); } catch(e){ bills = []; } }
  if (!bills || !bills.length) return {status:'ok', written:0};
  var sheet = getOrCreate(BILL_REQ_TAB, writeBillReqHeaders);
  var member = data.member || '', today = dateStr(data.date||''), n = 0;
  bills.forEach(function(b){
    if (!b || !b.project || !b.stage) return;
    sheet.appendRow([ nextId(sheet,'BR-'), today, member, b.project,
      b.discipline||'', b.stage, 'Pending', '', '', '', '' ]);
    n++;
  });
  return {status:'ok', written:n};
}
function getBillRequests(){
  var sheet = db().getSheetByName(BILL_REQ_TAB);
  if (!sheet || sheet.getLastRow() < 2) return {requests:[]};
  var rows = sheet.getDataRange().getValues(), out = [];
  for (var i=1;i<rows.length;i++){
    if (String(rows[i][6]||'').trim() !== 'Pending') continue;
    out.push({ row:i+1, reqId:String(rows[i][0]||''), date:cellDate(rows[i][1]),
      raisedBy:String(rows[i][2]||''), project:String(rows[i][3]||''),
      discipline:String(rows[i][4]||''), stage:String(rows[i][5]||'') });
  }
  return {requests:out};
}
function disposeBillRequest(data){
  var sheet = db().getSheetByName(BILL_REQ_TAB);
  if (!sheet) return {status:'error', message:'BILL_REQUESTS not found'};
  var row = parseInt(data.row,10);
  if (!row) return {status:'error', message:'no row'};
  var r = sheet.getRange(row,1,1,11).getValues()[0];
  if (String(r[6]||'').trim() !== 'Pending') return {status:'error', message:'already '+String(r[6]||'')};
  var today = dateStr(), action = String(data.disposition||'').trim();
  if (action === 'reject'){
    sheet.getRange(row,7).setValue('Rejected');
    sheet.getRange(row,8).setValue(data.reviewedBy||'');
    sheet.getRange(row,9).setValue(today);
    if (data.note) sheet.getRange(row,11).setValue(data.note);
    return {status:'ok', disposition:'Rejected'};
  }
  if (action === 'approve'){
    var aSheet = getOrCreate(ASSIGN_TAB, writeAssignHeaders);
    var taskId = 'T-' + Utilities.getUuid().substring(0,8).toUpperCase();
    var note = 'Raise bill — ' + String(r[4]||'') + ' · ' + String(r[5]||'');
    aSheet.appendRow([ taskId, '', String(r[3]||''), 'Aman Raghuwanshi', 'Raise Bill',
      1, 0, 1, 0, today, '', '', '', 'Not Started', '', '', 'Pending', '', '', '',
      note, (data.reviewedBy||'') + ' (bill approval)', 'Medium' ]);
    parkRowIfStalled(aSheet, String(r[3]||''));
    sheet.getRange(row,7).setValue('Approved');
    sheet.getRange(row,8).setValue(data.reviewedBy||'');
    sheet.getRange(row,9).setValue(today);
    sheet.getRange(row,10).setValue(taskId);
    return {status:'ok', disposition:'Approved', taskId:taskId};
  }
  return {status:'error', message:'unknown disposition: '+action};
}

// ════════════════════════════════════════════════════════════════
// EPIC K — Site Visit / Meeting Log → AI polish → lead approval → cumulative client PDF
// ════════════════════════════════════════════════════════════════
function writeMeetingLogHeaders(sheet){
  var h=['Log ID','Date','Time','Type','Project','Logged By','Team Attendees','Client Attendees',
         'Purpose','Body (raw)','Body (polished)','Duration (hrs)','Drive Folder','Photo IDs',
         'Video Links','Status','Approved By','Approval Date','What Changed','Report PDF ID',
         'Frozen Snapshot','Lead Reviewed'];
  sheet.getRange(1,1,1,h.length).setValues([h]).setBackground('#1F3A5F').setFontColor('#FFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
}
function writeDecisionLogHeaders(sheet){
  var h=['Item ID','Log ID','Project','Date','Category','Owner','Task','Deadline','Status'];
  sheet.getRange(1,1,1,h.length).setValues([h]).setBackground('#1F3A5F').setFontColor('#FFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// AI polish via Claude Haiku 4.5. Returns {body, items:{id:text}, whatChanged} or null
// (null → caller falls back to the raw text). Key in Script Properties; never in code.
function aiPolishLog(payload){
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) return null;
  var sys = "You are an editor for Ideaform Design Studio (architecture & interiors). Rewrite the provided "
    + "site-visit / meeting notes into clear, professional, courteous, client-ready English. ONLY improve tone, "
    + "grammar, clarity and structure. NEVER add, remove or change facts, numbers, measurements, names, dates, "
    + "deadlines or technical content. Format the minutes/observations as concise discrete POINTS — one point per "
    + "line, separated by a single newline, with NO leading numbers or bullets (the report numbers them). Keep it "
    + "concise. Return STRICT JSON only — no prose, no markdown fences.";
  var items = payload.items || [];
  var userText = 'Return JSON exactly in this shape: '
    + '{"body":"<polished minutes/observations as separate points, ONE per line, no numbering or bullets>","items":{"<id>":"<polished item text>"},'
    + '"whatChanged":"<one or two neutral sentences on how decisions changed versus the earlier open items below, or an empty string>"}.'
    + '\n\nNOTES:\n' + (payload.body || '(none)')
    + '\n\nACTION ITEMS (id: text):\n' + items.map(function(it){ return it.id+': '+it.text; }).join('\n')
    + (payload.prevContext ? '\n\nEARLIER OPEN DECISIONS (context for whatChanged only — do not rewrite these):\n'+payload.prevContext : '');
  try {
    var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method:'post', contentType:'application/json', muteHttpExceptions:true,
      headers:{ 'x-api-key':key, 'anthropic-version':'2023-06-01' },
      payload: JSON.stringify({ model:'claude-haiku-4-5', max_tokens:1500, system:sys,
                                messages:[{ role:'user', content:userText }] })
    });
    if (resp.getResponseCode() !== 200){ Logger.log('aiPolish HTTP '+resp.getResponseCode()+': '+resp.getContentText().substring(0,200)); return null; }
    var d = JSON.parse(resp.getContentText());
    if (d.stop_reason === 'refusal') return null;
    var txt=''; (d.content||[]).forEach(function(b){ if(b.type==='text') txt+=b.text; });
    var m = txt.match(/\{[\s\S]*\}/); if(!m) return null;
    return JSON.parse(m[0]);
  } catch(e){ Logger.log('aiPolish error: '+e); return null; }
}

// Drive: per-project / per-date folder under "IDS Logs"
function meetingDriveFolder(project, day){
  var root, it = DriveApp.getFoldersByName(LOGS_ROOT_FOLDER);
  root = it.hasNext() ? it.next() : DriveApp.createFolder(LOGS_ROOT_FOLDER);
  function child(parent, name){ var c=parent.getFoldersByName(name); return c.hasNext()?c.next():parent.createFolder(name); }
  return child(child(root, project||'Unfiled'), day||dateStr());
}
// Upload one base64 image to the visit's Drive folder. Called once per photo.
function uploadMeetingPhoto(data){
  try {
    var folder = meetingDriveFolder(data.project||'', dateStr(data.date||''));
    var b64 = String(data.dataUrl||'').replace(/^data:[^,]*,/,'');
    var blob = Utilities.newBlob(Utilities.base64Decode(b64), data.mime||'image/jpeg', data.name||('photo-'+Date.now()+'.jpg'));
    var f = folder.createFile(blob);
    return { status:'ok', fileId:f.getId(), folderId:folder.getId() };
  } catch(e){ return { status:'error', message:String(e) }; }
}

// Main submission. Writes MEETING_LOG (raw + AI-polished), DECISION_LOG items, creates
// IDS-team tasks, registers a CRM_LOG connection, leaves it "Pending" lead approval.
function submitMeetingLog(data, authEmail){
  var s = db();
  var logSheet = getOrCreate(MEETING_LOG_TAB, writeMeetingLogHeaders);
  var decSheet = getOrCreate(DECISION_LOG_TAB, writeDecisionLogHeaders);
  var logId = nextId(logSheet, 'ML-');
  var day = dateStr(data.date || '');
  var project = String(data.project || '').trim();
  var loggedBy = String(data.loggedBy || '').trim();
  var type = (data.type === 'Site Visit') ? 'Site Visit' : 'Meeting';

  // Action plans: data.actions = {Client:[{owner,text,deadline}], IDS:[...], Contractor:[...], Other:[...]}
  var actions = data.actions || {};
  var flat = [];   // {cat, owner, text, deadline, itemId}
  ['Client','IDS','Contractor','Other'].forEach(function(cat){
    (actions[cat]||[]).forEach(function(it){
      var t = String(it.text||'').trim(); if(!t) return;
      // Unique per item — nextId() would return the SAME id for each (rows aren't
      // appended until later), which collapsed every action item to one text.
      flat.push({ cat:cat, owner:String(it.owner||'').trim(), text:t, deadline:String(it.deadline||'').trim(),
                  itemId: 'DEC-'+Utilities.getUuid().substring(0,8).toUpperCase() });
    });
  });

  // AI polish (raw kept; polished used for the client report). Falls back to raw if no key / failure.
  var bodyRaw = String(data.body || '').trim();
  var prev = openDecisionsText(project);
  var polished = aiPolishLog({ body:bodyRaw, items:flat.map(function(f){return {id:f.itemId, text:f.text};}), prevContext:prev });
  var bodyPolished = (polished && polished.body) ? polished.body : bodyRaw;
  var whatChanged  = (polished && polished.whatChanged) ? polished.whatChanged : '';
  if (polished && polished.items){ flat.forEach(function(f){ if(polished.items[f.itemId]) f.textPolished = polished.items[f.itemId]; }); }

  // Photos already uploaded via uploadMeetingPhoto → data.photoIds; folder id passed through
  var photoIds = (data.photoIds||[]).join(',');
  var folderId = data.folderId || '';
  var videoLinks = (data.videoLinks||[]).join(' | ');

  prependRow(logSheet, [ logId, day, String(data.time||''), type, project, loggedBy,
    (data.teamAttendees||[]).join(', '), String(data.clientAttendees||''),
    (data.purpose||[]).join(', '), bodyRaw, bodyPolished, String(data.duration||''),
    folderId, photoIds, videoLinks, 'Draft', '', '', whatChanged, '', '', '' ]);

  // Decision items → DECISION_LOG (store polished text where available); IDS items → tasks
  var aSheet = getOrCreate(ASSIGN_TAB, writeAssignHeaders);
  flat.forEach(function(f){
    decSheet.appendRow([ f.itemId, logId, project, day, f.cat, f.owner, f.textPolished||f.text, f.deadline, 'Open' ]);
    if (f.cat === 'IDS' && f.owner && !data.skipTasks){
      var assignee = resolveAssignee(f.owner, project) || f.owner;
      var taskId = 'T-'+Utilities.getUuid().substring(0,8).toUpperCase();
      aSheet.appendRow([ taskId, '', project, assignee, 'MoM Action', 1, 1, 1, 1, day,
        f.deadline || addDaysToStr(day,3), '', '', 'Not Started', '', '', 'Pending', '', '', '',
        (f.textPolished||f.text)+' [From '+type+' '+day+']', loggedBy+' (MoM)', 'Medium' ]);
      parkRowIfStalled(aSheet, project);
    }
  });

  // Register as a client connection so it shows on the projects dashboard
  appendConnections(loggedBy, day, [{ project:project, type:type, notes:'Logged '+type+(data.clientAttendees?(' with '+data.clientAttendees):'') }], logId);

  // Return the AI-polished content so the author can review/edit it before the
  // shareable PDF is generated (Option A — review-then-generate).
  return { status:'ok', logId:logId, polished: !!polished, bodyPolished: bodyPolished,
    items: flat.map(function(f){ return { id:f.itemId, cat:f.cat, owner:f.owner, text:(f.textPolished||f.text) }; }) };
}

// Finalize an entry after the author has reviewed/edited the polished text:
// apply edits, freeze the immutable snapshot, mark Final, regenerate the
// cumulative PDF, and return its link so the team can share immediately.
function finalizeMeetingLog(data, authEmail){
  var s = db();
  var logSheet = s.getSheetByName(MEETING_LOG_TAB);
  var decSheet = s.getSheetByName(DECISION_LOG_TAB);
  if (!logSheet) return {status:'error', message:'MEETING_LOG not found'};
  var logId = String(data.logId||'').trim();
  if (!logId) return {status:'error', message:'no logId'};
  var rows = logSheet.getDataRange().getValues(), rIdx=-1;
  for (var i=1;i<rows.length;i++){ if(String(rows[i][0]||'')===logId){ rIdx=i; break; } }
  if (rIdx<0) return {status:'error', message:'log not found'};
  var rowNum = rIdx+1;

  // apply edited body
  if (typeof data.bodyPolished === 'string') logSheet.getRange(rowNum,11).setValue(data.bodyPolished);
  // apply edited attendees / purpose / photos (so a missed detail can be added
  // on the review screen without re-filling the whole form)
  function csv(v){ return Array.isArray(v) ? v.join(', ') : String(v||''); }
  if (data.teamAttendees   != null) logSheet.getRange(rowNum,7 ).setValue(csv(data.teamAttendees));
  if (data.clientAttendees != null) logSheet.getRange(rowNum,8 ).setValue(String(data.clientAttendees||''));
  if (data.purpose         != null) logSheet.getRange(rowNum,9 ).setValue(csv(data.purpose));
  if (data.photoIds        != null) logSheet.getRange(rowNum,14).setValue(Array.isArray(data.photoIds)?data.photoIds.join(','):String(data.photoIds||''));
  // apply edited action-item texts
  var edits = data.items || {};
  if (decSheet && decSheet.getLastRow()>1){
    var dr = decSheet.getDataRange().getValues();
    for (var j=1;j<dr.length;j++){ var id=String(dr[j][0]||''); if(edits[id]!=null && String(edits[id]).trim()) decSheet.getRange(j+1,7).setValue(String(edits[id])); }
  }

  // build the entry from current (edited) values, freeze its text snapshot
  var purposeVal = (data.purpose!=null) ? csv(data.purpose) : String(rows[rIdx][8]||'');
  var entry = {
    logId:logId, date:cellDate(rows[rIdx][1]), type:String(rows[rIdx][3]||''),
    purpose:purposeVal, body:String(data.bodyPolished!=null?data.bodyPolished:rows[rIdx][10]||'') };
  var snapshot = buildEntryBodyHTML(entry, decisionsForLog(logId, edits));
  logSheet.getRange(rowNum,21).setValue(snapshot);   // U frozen snapshot
  logSheet.getRange(rowNum,16).setValue('Final');    // P status → shareable
  logSheet.getRange(rowNum,17).setValue(authEmail||''); logSheet.getRange(rowNum,18).setValue(dateStr());

  var project = String(rows[rIdx][4]||'').trim();
  var pdf = generateProjectReportPDF(project);
  if (pdf && pdf.fileId) logSheet.getRange(rowNum,20).setValue(pdf.fileId);
  return { status:'ok', logId:logId, pdfUrl: pdf && pdf.url, pdfId: pdf && pdf.fileId };
}

// Current decision items for a log (id/cat/owner/text), applying any pending edits.
function decisionsForLog(logId, edits){
  var sheet = db().getSheetByName(DECISION_LOG_TAB);
  var by={Client:[],IDS:[],Contractor:[],Other:[]};
  if (!sheet || sheet.getLastRow()<2) return by;
  edits = edits||{};
  var rows = sheet.getDataRange().getValues();
  for (var i=1;i<rows.length;i++){ if(String(rows[i][1]||'')!==logId) continue;
    var id=String(rows[i][0]||''), cat=String(rows[i][4]||'Other');
    (by[cat]||by.Other).push({ owner:String(rows[i][5]||''), text:(edits[id]!=null&&String(edits[id]).trim())?String(edits[id]):String(rows[i][6]||'') });
  }
  return by;
}

// Concatenated open decisions for a project (feeds AI "what changed" + report continuity)
function openDecisionsText(project){
  var sheet = db().getSheetByName(DECISION_LOG_TAB);
  if (!sheet || sheet.getLastRow()<2) return '';
  var rows = sheet.getDataRange().getValues(), out=[], pl=String(project||'').toLowerCase();
  for (var i=1;i<rows.length;i++){
    if (String(rows[i][2]||'').toLowerCase() !== pl) continue;
    if (String(rows[i][8]||'').trim() !== 'Open') continue;
    out.push('- ['+String(rows[i][4]||'')+'] '+String(rows[i][6]||'')+(rows[i][7]?(' (due '+cellDate(rows[i][7])+')'):''));
  }
  return out.slice(-30).join('\n');
}

// Pending lead approvals — for the approval form
function getMeetingApprovals(){
  var sheet = db().getSheetByName(MEETING_LOG_TAB);
  if (!sheet || sheet.getLastRow()<2) return {logs:[]};
  var rows = sheet.getDataRange().getValues(), out=[];
  for (var i=1;i<rows.length;i++){
    if (String(rows[i][15]||'').trim() !== 'Pending') continue;
    out.push({ row:i+1, logId:String(rows[i][0]||''), date:cellDate(rows[i][1]), type:String(rows[i][3]||''),
      project:String(rows[i][4]||''), loggedBy:String(rows[i][5]||''), team:String(rows[i][6]||''),
      clients:String(rows[i][7]||''), purpose:String(rows[i][8]||''),
      bodyRaw:String(rows[i][9]||''), bodyPolished:String(rows[i][10]||''), whatChanged:String(rows[i][18]||'') });
  }
  return {logs:out};
}

// Lead approves (optionally edits the polished text), then the cumulative PDF is generated.
function approveMeetingLog(data, authEmail){
  var sheet = db().getSheetByName(MEETING_LOG_TAB);
  if (!sheet) return {status:'error', message:'MEETING_LOG not found'};
  var row = parseInt(data.row,10); if(!row) return {status:'error', message:'no row'};
  var today = dateStr();
  if (data.disposition === 'reject'){
    sheet.getRange(row,16).setValue('Rejected'); sheet.getRange(row,17).setValue(data.reviewedBy||authEmail||''); sheet.getRange(row,18).setValue(today);
    return {status:'ok', disposition:'Rejected'};
  }
  if (typeof data.bodyPolished === 'string' && data.bodyPolished.trim()) sheet.getRange(row,11).setValue(data.bodyPolished);
  sheet.getRange(row,16).setValue('Approved'); sheet.getRange(row,17).setValue(data.reviewedBy||authEmail||''); sheet.getRange(row,18).setValue(today);
  var project = String(sheet.getRange(row,5).getValue()||'').trim();
  var pdf = generateProjectReportPDF(project);
  if (pdf && pdf.fileId) sheet.getRange(row,20).setValue(pdf.fileId);
  return {status:'ok', disposition:'Approved', pdfUrl: pdf && pdf.url, pdfId: pdf && pdf.fileId };
}

// ── Frozen-snapshot helpers (shared by finalize + report) ────────────────────
function mlEsc(t){ return String(t==null?'':t).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function mlBodyPoints(t){ return String(t||'').split(/\r?\n/).map(function(l){
    return l.replace(/^\s*(?:\d+[.)]|[-•*])\s*/,'').trim(); }).filter(Boolean); }
function mlPlanBlock(title,arr){ if(!arr||!arr.length) return '';
  return '<div class="ap"><div class="ap-t">'+title+'</div><ol>'+arr.map(function(x){
    return '<li>'+(x.owner?'<b>'+mlEsc(x.owner)+':</b> ':'')+mlEsc(x.text)+'</li>'; }).join('')+'</ol></div>'; }
// The immutable body of one entry: purpose + numbered minutes/observations +
// action plans (NO status, NO dates). Header/badge/images are added by the report.
function buildEntryBodyHTML(e, decsBy){
  var pts=mlBodyPoints(e.body);
  var head=(e.type==='Site Visit')?'Key Observations / Issues':'Minutes of Meeting';
  var bodyHtml = pts.length ? '<div class="bh">'+head+'</div><ol class="body">'+pts.map(function(p){return '<li>'+mlEsc(p)+'</li>';}).join('')+'</ol>' : '';
  return (e.purpose?'<div class="meta">Purpose: '+mlEsc(e.purpose)+'</div>':'')
    + bodyHtml
    + mlPlanBlock('Action Plan — Client', decsBy.Client) + mlPlanBlock('Action Plan — IDS Team', decsBy.IDS)
    + mlPlanBlock('Action Plan — Contractors', decsBy.Contractor) + mlPlanBlock('Action Plan — Other', decsBy.Other);
}

// Build the cumulative, newest-first client PDF for a project: cover index of all
// visits/meetings (date + attendees) + each finalized entry, photos inlined as base64.
function generateProjectReportPDF(project){
  var sheet = db().getSheetByName(MEETING_LOG_TAB);
  if (!sheet || sheet.getLastRow()<2) return null;
  var rows = sheet.getDataRange().getValues(), pl=String(project||'').toLowerCase();
  var decSheet = db().getSheetByName(DECISION_LOG_TAB);
  var decRows = decSheet ? decSheet.getDataRange().getValues() : [];
  var entries=[];
  for (var i=1;i<rows.length;i++){
    if (String(rows[i][4]||'').toLowerCase() !== pl) continue;
    var st = String(rows[i][15]||'').trim();
    if (st !== 'Final' && st !== 'Approved') continue;   // shareable entries only
    entries.push({
      logId:String(rows[i][0]||''), date:cellDate(rows[i][1]), type:String(rows[i][3]||''),
      team:String(rows[i][6]||''), clients:String(rows[i][7]||''), purpose:String(rows[i][8]||''),
      body:String(rows[i][10]||''), photoIds:String(rows[i][13]||''), videos:String(rows[i][14]||''),
      snapshot:String(rows[i][20]||'') });   // U — frozen text snapshot (immutable history)
  }
  if (!entries.length) return null;
  entries.sort(function(a,b){ return a.date < b.date ? 1 : (a.date > b.date ? -1 : 0); }); // newest first

  function esc(t){ return mlEsc(t); }
  function decsFor(logId){
    var by={Client:[],IDS:[],Contractor:[],Other:[]};
    for (var j=1;j<decRows.length;j++){ if(String(decRows[j][1]||'')!==logId) continue;
      var cat=String(decRows[j][4]||'Other'); (by[cat]||by.Other).push({owner:String(decRows[j][5]||''),text:String(decRows[j][6]||'')}); }
    return by;
  }

  // Estimate each entry's START page for the cover index: cover = page 1, then per
  // entry 1 text page + ceil(images/2) image pages (2 photos per A4 page).
  var running=2;
  entries.forEach(function(e){
    e.imgIds=e.photoIds.split(',').map(function(s){return s.trim();}).filter(Boolean);
    e.startPage=running; running += 1 + Math.ceil(e.imgIds.length/2);
  });

  // cover index — Date · Type · Attendees · Page (page LAST)
  var index = entries.map(function(e){ return '<tr><td>'+esc(e.date)+'</td><td>'+esc(e.type)+'</td><td>'
      +esc([e.team,e.clients].filter(Boolean).join(', ')||'—')+'</td><td class="pg">'+e.startPage+'</td></tr>'; }).join('');

  var sections = entries.map(function(e,idx){
    var d=decsFor(e.logId);
    var imgPages='';
    for (var k=0;k<e.imgIds.length;k+=2){
      var pair='';
      [e.imgIds[k], e.imgIds[k+1]].forEach(function(id){ if(!id) return;
        try{ var f=DriveApp.getFileById(id); var b=f.getBlob();
          pair += '<img src="data:'+b.getContentType()+';base64,'+Utilities.base64Encode(b.getBytes())+'">'; }catch(er){} });
      if (pair) imgPages += '<div class="imgpage">'+pair+'</div>';
    }
    // Frozen snapshot = the entry's immutable text (built at finalize). Older
    // entries are reproduced exactly as shared; only the header/badge/images are
    // positional. Fallback to a live build for any legacy entry without a snapshot.
    var bodyHtml = e.snapshot || buildEntryBodyHTML(e, d);
    return '<div class="sec'+(idx>0?' brk':'')+'">'
      + '<div class="sh">'+esc(e.type)+' &mdash; '+esc(e.date)+(idx===0?' <span class="latest">Latest</span>':'')+'</div>'
      + bodyHtml
      + (e.videos?'<div class="vids"><b>Videos:</b> '+e.videos.split('|').map(function(v){v=v.trim();return v?'<a href="'+esc(v)+'">'+esc(v)+'</a>':'';}).join(' &nbsp; ')+'</div>':'')
      + imgPages
      + '</div>';
  }).join('');

  var logoImg = (typeof IDS_LOGO_B64!=='undefined' && IDS_LOGO_B64)
    ? '<img class="logo" src="'+IDS_LOGO_B64+'">'
    : '<div class="logotext"><span style="color:#4D4D4F">IDEA</span><span style="color:#F2A03D">FORM</span></div>';

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    + 'body{font-family:Georgia,\'Times New Roman\',serif;color:#2A2A2A;margin:0;padding:34px 40px;font-size:13px;line-height:1.6}'
    + '.cover{min-height:980px;display:flex;flex-direction:column;text-align:center;page-break-after:always}'
    + '.logo{max-width:300px;max-height:120px;margin:46px auto 0;display:block}'
    + '.logotext{font-family:Arial,sans-serif;font-size:40px;font-weight:700;letter-spacing:8px;margin-top:64px}'
    + '.rule{height:2px;background:#F2A03D;width:70px;margin:30px auto}'
    + '.cv-lbl{font-size:13px;letter-spacing:4px;color:#9E9B94;text-transform:uppercase;margin-top:30px}'
    + '.cv-proj{font-size:30px;color:#3F3F41;margin-top:12px}'
    + '.cv-title{font-size:22px;color:#E08A1E;font-style:italic;margin-top:26px}'
    + '.idxwrap{margin-top:auto}'
    + '.idx-h{font-size:13px;letter-spacing:3px;color:#4D4D4F;text-transform:uppercase;border-bottom:1.5px solid #F2A03D;padding-bottom:6px;text-align:left}'
    + 'table.idx{width:100%;border-collapse:collapse;font-size:12.5px}'
    + 'table.idx th{background:#4D4D4F;color:#fff;text-align:left;padding:7px 9px;font-weight:normal}'
    + 'table.idx th.pg,table.idx td.pg{text-align:center;width:48px}'
    + 'table.idx td{border-bottom:1px solid #E7E3DA;padding:7px 9px;text-align:left;vertical-align:top}'
    + '.sec{padding-top:6px}.brk{page-break-before:always}'
    + '.sh{font-size:18px;font-weight:bold;color:#4D4D4F;border-bottom:2px solid #F2A03D;padding-bottom:5px;margin-bottom:8px}'
    + '.latest{font-size:11px;background:#F2A03D;color:#fff;padding:2px 8px;border-radius:9px;vertical-align:middle;font-family:Arial,sans-serif}'
    + '.meta{font-size:12px;color:#6B6860;margin-bottom:10px}'
    + '.bh{font-size:15px;font-weight:bold;color:#4D4D4F;margin:12px 0 4px}'
    + 'ol.body{margin:4px 0 8px;padding-left:24px;font-size:13px;line-height:1.65}ol.body li{margin-bottom:4px}'
    + '.ap{margin:10px 0}.ap-t{font-size:12px;font-weight:bold;color:#E08A1E;text-transform:uppercase;letter-spacing:.5px}'
    + '.ap ol{margin:4px 0 0;padding-left:24px;font-size:13px;line-height:1.6}.ap li{margin-bottom:3px}.st{color:#3B6D11;font-size:11px}'
    + '.imgpage{page-break-before:always;text-align:center}'
    + '.imgpage img{display:block;width:100%;max-height:430px;object-fit:contain;margin:0 auto 16px;border:1px solid #E2DFD8}'
    + '.vids{font-size:12px;margin-top:8px}'
    + '.foot{margin-top:22px;border-top:1px solid #E7E3DA;padding-top:7px;font-size:10px;color:#9E9B94;text-align:center}'
    + '</style></head><body>'
    + '<div class="cover">'+logoImg
    + '<div class="rule"></div>'
    + '<div class="cv-lbl">Project</div><div class="cv-proj">'+esc(project)+'</div>'
    + '<div class="cv-title">Meeting &amp; Site Visit Logs</div>'
    + '<div class="idxwrap"><div class="idx-h">Index of logs</div>'
    + '<table class="idx"><tr><th>Date</th><th>Type</th><th>Attendees</th><th class="pg">Page</th></tr>'+index+'</table></div>'
    + '</div>'
    + sections
    + '<div class="foot">All action plans are subject to follow-up and verification. For clarifications, please contact the project manager. &copy; Ideaform Design Studio &middot; Confidential</div>'
    + '</body></html>';

  var pdf = Utilities.newBlob(html, 'text/html', project+' — IDS Log.html').getAs('application/pdf')
            .setName(project+' — Meeting & Site Visit Log.pdf');
  var folder, it = DriveApp.getFoldersByName(LOGS_ROOT_FOLDER);
  var root = it.hasNext()?it.next():DriveApp.createFolder(LOGS_ROOT_FOLDER);
  var pc = root.getFoldersByName(project||'Unfiled'); folder = pc.hasNext()?pc.next():root.createFolder(project||'Unfiled');
  // stable filename → each submission cleanly REPLACES the single cumulative PDF
  var old = folder.getFilesByName(pdf.getName()); while(old.hasNext()){ old.next().setTrashed(true); }
  var saved = folder.createFile(pdf);
  return { fileId:saved.getId(), url:saved.getUrl() };
}

// Per-project meeting timeline for the projects dashboard drawer
function getMeetingTimeline(project){
  var sheet = db().getSheetByName(MEETING_LOG_TAB);
  if (!sheet || sheet.getLastRow()<2) return {meetings:[]};
  var rows = sheet.getDataRange().getValues(), pl=String(project||'').toLowerCase(), out=[];
  for (var i=1;i<rows.length;i++){ if(String(rows[i][4]||'').toLowerCase()!==pl) continue;
    if(String(rows[i][15]||'').trim()==='Deleted') continue;   // hide deleted logs
    out.push({ logId:String(rows[i][0]||''), date:cellDate(rows[i][1]), type:String(rows[i][3]||''), loggedBy:String(rows[i][5]||''),
      clients:String(rows[i][7]||''), status:String(rows[i][15]||''), approved:String(rows[i][15]||''),
      pdfId:String(rows[i][19]||'') }); }
  out.sort(function(a,b){ return a.date<b.date?1:(a.date>b.date?-1:0); });
  return {meetings:out};
}

// Delete an incorrect log (Siddharth only): mark it Deleted (excluded from the
// report + timeline), void its action items, and regenerate the cumulative PDF.
// Map a signed-in email → TEAM member name (so an author can manage their own logs).
function nameForEmail(email){
  email=String(email||'').toLowerCase(); if(!email) return '';
  var t=db().getSheetByName(TEAM_TAB); if(!t||t.getLastRow()<2) return '';
  var r=t.getDataRange().getValues();
  for(var i=1;i<r.length;i++){ if(String(r[i][4]||'').trim().toLowerCase()===email) return String(r[i][0]||'').trim(); }
  return '';
}
function deleteMeetingLog(data, authEmail){
  var s=db(), sheet=s.getSheetByName(MEETING_LOG_TAB);
  if(!sheet) return {status:'error', message:'MEETING_LOG not found'};
  var logId=String(data.logId||'').trim(); if(!logId) return {status:'error', message:'no logId'};
  var rows=sheet.getDataRange().getValues(), rIdx=-1;
  for(var i=1;i<rows.length;i++){ if(String(rows[i][0]||'')===logId){ rIdx=i; break; } }
  if(rIdx<0) return {status:'error', message:'log not found'};
  // The author of the log, or a director, may delete it.
  var loggedBy=String(rows[rIdx][5]||'').trim();
  if(!isDirector(authEmail) && nameForEmail(authEmail)!==loggedBy)
    return {status:'error', code:'forbidden', message:'Only the author or Siddharth can delete this log.'};
  var project=String(rows[rIdx][4]||'').trim();
  sheet.getRange(rIdx+1,16).setValue('Deleted');
  var dec=s.getSheetByName(DECISION_LOG_TAB);
  if(dec && dec.getLastRow()>1){ var dr=dec.getDataRange().getValues();
    for(var j=1;j<dr.length;j++){ if(String(dr[j][1]||'')===logId) dec.getRange(j+1,9).setValue('Deleted'); } }
  // remove this log's auto-created CRM connection (Submission ID = the logId)
  var crm=s.getSheetByName(CRM_LOG_TAB);
  if(crm && crm.getLastRow()>1){ var cr=crm.getDataRange().getValues();
    for(var k=cr.length-1;k>=1;k--){ if(String(cr[k][1]||'').trim()===logId) crm.deleteRow(k+1); } }
  var pdf=generateProjectReportPDF(project);
  return {status:'ok', pdfUrl: pdf && pdf.url, pdfId: pdf && pdf.fileId };
}

// All non-deleted meeting/site-visit logs (managers only) — for the Logs Manager.
function getAllMeetingLogs(){
  var sheet=db().getSheetByName(MEETING_LOG_TAB);
  if(!sheet || sheet.getLastRow()<2) return {logs:[]};
  var rows=sheet.getDataRange().getValues(), out=[];
  for(var i=1;i<rows.length;i++){ if(String(rows[i][15]||'').trim()==='Deleted') continue;
    out.push({ logId:String(rows[i][0]||''), date:cellDate(rows[i][1]), type:String(rows[i][3]||''),
      project:String(rows[i][4]||''), loggedBy:String(rows[i][5]||''), clients:String(rows[i][7]||''),
      status:String(rows[i][15]||''), pdfId:String(rows[i][19]||'') }); }
  out.sort(function(a,b){ return (b.date||'').localeCompare(a.date||''); });
  return {logs:out};
}

// ONE-TIME cleanup — run from the Apps Script editor (select purgeMockLogsNow → Run).
// Removes ALL meeting/site-visit logs + their derived data (they were mock):
//   • MEETING_LOG + DECISION_LOG rows (and trashes their report PDFs in Drive)
//   • MoM-Action tasks in TASK_ASSIGNMENTS
//   • meeting-derived CRM_LOG connections (Submission ID begins "ML-")
// Does NOT touch SITE_EXECUTION (real DPER site reports) or normal DPR connections.
function purgeMockLogsNow(){
  var s=db(), out={meetingRows:0, decisionRows:0, momTasks:0, connections:0, pdfsTrashed:0};
  var ml=s.getSheetByName(MEETING_LOG_TAB);
  if(ml && ml.getLastRow()>1){
    var mr=ml.getDataRange().getValues();
    for(var i=1;i<mr.length;i++){ var pid=String(mr[i][19]||'').trim(); if(pid){ try{ DriveApp.getFileById(pid).setTrashed(true); out.pdfsTrashed++; }catch(e){} } }
    out.meetingRows=ml.getLastRow()-1; ml.deleteRows(2, ml.getLastRow()-1);
  }
  var dl=s.getSheetByName(DECISION_LOG_TAB);
  if(dl && dl.getLastRow()>1){ out.decisionRows=dl.getLastRow()-1; dl.deleteRows(2, dl.getLastRow()-1); }
  var a=s.getSheetByName(ASSIGN_TAB);
  if(a && a.getLastRow()>1){ var ar=a.getDataRange().getValues();
    for(var i=ar.length-1;i>=1;i--){ if(String(ar[i][4]||'').trim()==='MoM Action'){ a.deleteRow(i+1); out.momTasks++; } } }
  var c=s.getSheetByName(CRM_LOG_TAB);
  if(c && c.getLastRow()>1){ var cr=c.getDataRange().getValues();
    for(var i=cr.length-1;i>=1;i--){ if(String(cr[i][1]||'').trim().indexOf('ML-')===0){ c.deleteRow(i+1); out.connections++; } } }
  Logger.log('purgeMockLogsNow done: '+JSON.stringify(out));
  return out;
}

// ════════════════════════════════════════════════════════════════
// SUBMIT APPROVALS
// TASK_LOG: LeadApproved=K(11) ApprovedBy=L(12) ReviewedOn=M(13) Notes=N(14)
// TASK_ASSIGNMENTS: LeadApproved=P(16) ApprovedBy=Q(17) ApprovalDate=R(18)
//   SelfStatus=N(14) SelfDoneDate=O(15) RevisionTag=S(19) Notes=T(20)
// ════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════
// reassignTask — lead marks a delayed task "Work Not Done" off the
// original assignee (the −2 reliability penalty sticks to them) and
// re-creates it fresh for a new assignee at full points.
// ════════════════════════════════════════════════════════════════
function reassignTask(data) {
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) return {status:'error', message:'TASK_ASSIGNMENTS not found'};
  var row = parseInt(data.row);
  if (!row || row < 2) return {status:'error', message:'Invalid row'};
  var newAssignee = String(data.newAssignee||'').trim();
  if (!newAssignee) return {status:'error', message:'No new assignee'};
  var leadName = String(data.leadName||'lead').trim();
  var today = dateStr();

  var lastCol = sheet.getLastColumn();
  var orig = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  var origAssignee = String(orig[3]||'').trim();

  // 1. Close the original off the assignee (penalty stays with them)
  sheet.getRange(row, 14).setValue('Work Not Done');  // N SelfStatus
  var existNotes = String(orig[20]||'').trim();        // U Notes
  sheet.getRange(row, 21).setValue(
    (existNotes ? existNotes + ' | ' : '') +
    'Reassigned to ' + newAssignee + ' by ' + leadName + ' on ' + today + ' (delay)');

  // 2. Fresh copy for the new assignee at full points
  var newId = 'T-' + Utilities.getUuid().substring(0,8).toUpperCase();
  var n = orig.slice();
  n[0]  = newId;            // A TaskID
  n[3]  = newAssignee;      // D AssignedTo
  n[9]  = today;            // J AssignedDate (K Deadline kept from original)
  n[13] = 'Not Started';    // N SelfStatus
  n[14] = '';               // O SelfStatusDate
  if (n.length > 15) n[15] = ''; // P ActualCompletion
  n[16] = 'Pending';        // Q LeadApproved
  n[17] = '';               // R ApprovedBy
  n[18] = '';               // S ApprovalDate
  n[19] = '';               // T RevisionTag
  n[20] = 'Reassigned from ' + origAssignee + ' (delay)'; // U Notes
  n[21] = leadName + ' (reassigned)'; // V AssignedBy
  sheet.appendRow(n);

  Logger.log('Reassigned row '+row+' from '+origAssignee+' → '+newAssignee+' ('+newId+')');
  return {status:'ok', newTaskId:newId, original:origAssignee, newAssignee:newAssignee};
}

function submitApprovals(data) {
  var s        = db();
  var taskAssn = s.getSheetByName(ASSIGN_TAB);

  var approvals = data.approvals || [];
  var approved = 0, rejected = 0;
  var now = nowStr(), today = dateStr();

  approvals.forEach(function(a) {
    var row    = parseInt(a.row);
    if (!row || row < 2) return;

    // ── Update TASK_ASSIGNMENTS ───────────────────────────────
    // source is always TASK_ASSIGNMENTS (TASK_LOG removed)
    var ar = row;

    if (taskAssn && ar >= 2) {
      // Auto-detect column structure for write-back
      var apprHeaders = taskAssn.getRange(1,1,1,taskAssn.getLastColumn()).getValues()[0];
      var apprIs23 = apprHeaders.length >= 23 || apprHeaders.indexOf('Actual Completion Date') > -1;
      var WR_LEADAPPR = apprIs23 ? 17 : 16; // 1-indexed: Q=17 new, P=16 old
      var WR_APPBY    = apprIs23 ? 18 : 17;
      var WR_APPDTCOL = apprIs23 ? 19 : 18;
      var WR_SELFSTAT = 14; // N always 1-indexed=14
      var WR_STATUSDT = 15; // O always 1-indexed=15
      var WR_ACTUALDT = apprIs23 ? 16 : -1; // P only in 23-col, 1-indexed=16
      var WR_REVTAG   = apprIs23 ? 20 : 19;
      var WR_NOTES    = apprIs23 ? 21 : 20;

      taskAssn.getRange(ar, WR_LEADAPPR).setValue(a.decision === 'Yes' ? 'Yes' : 'No');
      taskAssn.getRange(ar, WR_APPBY).setValue(a.leadName);
      taskAssn.getRange(ar, WR_APPDTCOL).setValue(today);

      // If lead corrected task type or points — update task row
      if (a.correctedTaskType && a.correctedTaskType !== '') {
        taskAssn.getRange(ar, 5).setValue(a.correctedTaskType);  // E Stage (same in both)
      }
      if (a.correctedPts && parseFloat(a.correctedPts) > 0) {
        taskAssn.getRange(ar, 9).setValue(parseFloat(a.correctedPts)); // I WeightedPoints
      }

      if (a.decision === 'No') {
        // Rejection — reopen task + auto-increment revision tag
        taskAssn.getRange(ar, WR_SELFSTAT).setValue('In Progress');
        taskAssn.getRange(ar, WR_STATUSDT).setValue('');
        if (WR_ACTUALDT > 0) taskAssn.getRange(ar, WR_ACTUALDT).setValue('');
        var existing = String(taskAssn.getRange(ar, WR_REVTAG).getValue() || '').trim();
        var revNum   = existing ? parseInt(existing.replace('R','')) + 1 : 1;
        taskAssn.getRange(ar, WR_REVTAG).setValue('R' + revNum);
        var existNotes = String(taskAssn.getRange(ar, WR_NOTES).getValue() || '').trim();
        taskAssn.getRange(ar, WR_NOTES).setValue(
          (existNotes ? existNotes + ' | ' : '') +
          'Rejected ' + today + (a.note ? ' — ' + a.note : '')
        );
        // Write notification to NOTIFICATIONS tab
        var notifData = {
          project          : a.project      || String(taskAssn.getRange(ar, 3).getValue()||''),
          taskType         : a.taskType     || String(taskAssn.getRange(ar, 5).getValue()||''),
          correctedTaskType: a.correctedTaskType || '',
          originalPts      : taskAssn.getRange(ar, 9).getValue() || 0,
          correctedPts     : a.correctedPts || '',
          note             : a.note         || '',
          leadName         : a.leadName     || '',
        };
        var memberName = String(taskAssn.getRange(ar, 4).getValue()||'');
        // Notification removed — rejections shown in DPR form directly
      }
    }
    if (a.decision === 'Yes') approved++; else rejected++;
  });

  // ── Resolve linked SITE_ISSUES when task approved ──────────
  if (approved > 0) {
    var issSheet = db().getSheetByName(SITE_ISSUES_TAB);
    if (issSheet && issSheet.getLastRow() > 1) {
      var issRows = issSheet.getDataRange().getValues();
      for (var ii=1; ii<issRows.length; ii++) {
        var issTaskId = String(issRows[ii][11]||'').trim(); // L TaskID
        var issStatus = String(issRows[ii][10]||'').trim(); // K Status
        if (!issTaskId || issStatus === 'Resolved') continue;
        approvals.forEach(function(a) {
          if (a.taskId === issTaskId && a.decision === 'Yes') {
            issSheet.getRange(ii+1, 11).setValue('Resolved');
            issSheet.getRange(ii+1, 14).setValue(dateStr());
            Logger.log('SITE_ISSUES resolved: ' + issTaskId);
          }
        });
      }
    }
  }

  return {status:'ok', approved:approved, rejected:rejected};
}

// ════════════════════════════════════════════════════════════════
// WEEKLY SCORECARD — runs every Monday 6 AM
//
// TEAM_SCORECARD cols: Member(A) Role(B) WeeklyTarget(C) ApprovedPts(D)
//   OnTimeRate(E) RevisionRate(F) DPRStreak(G) OverdueTasks(H)
//   MonthlyPts(I) MonthlyBenchmark(J) MonthlyScore%(K) Status(L)
//
// TASK_LOG (after Area/Drawing added):
//   Date=B(1) Member=C(2) Pts=J(9) LeadApproved=K(10)
// TASK_ASSIGNMENTS:
//   AssignedTo=D(3) ProjectName=C(2) Deadline=K(10)
//   SelfStatus=N(13) SelfDoneDate=O(14) LeadApproved=P(15)
//   AssignedDate=J(9) RevisionTag=S(18)
// DAILY_SUMMARY:
//   Date=A(0) Member=B(1) ArrivedOnTime=D(3)
// ════════════════════════════════════════════════════════════════
function calculateWeeklyScorecard() {
  Logger.log('=== Weekly Scorecard: ' + new Date().toISOString() + ' ===');

  var s   = db();
  var cfg = readConfig();
  var sw  = cfg.scoringWeights;
  var rp  = cfg.revPenalties;

  // Previous Monday → Sunday
  var mon = mondayOf(new Date()); mon.setDate(mon.getDate() - 7);
  var sun = new Date(mon); sun.setDate(sun.getDate() + 6);
  var monStr = dateStr(mon), sunStr = dateStr(sun);
  var monthStart    = new Date(mon.getFullYear(), mon.getMonth(), 1);
  var monthStartStr = dateStr(monthStart);
  var today         = dateStr();

  Logger.log('Week: ' + monStr + ' to ' + sunStr);

  // Load all data once
  var asSheet  = s.getSheetByName(ASSIGN_TAB);
  var sumSheet = s.getSheetByName(SUMMARY_TAB);
  var tSheet   = s.getSheetByName(TEAM_TAB);
  if (!tSheet) { Logger.log('TEAM tab not found'); return; }

  var asRows  = asSheet  ? asSheet.getDataRange().getValues()  : [];
  var sumRows = sumSheet ? sumSheet.getDataRange().getValues() : [];
  var tRows   = tSheet.getDataRange().getValues();

  // Auto-detect TASK_ASSIGNMENTS column structure
  var asHeaders = asRows[0] ? asRows[0].map(function(h){ return String(h||'').trim(); }) : [];
  var scIs23    = asHeaders.length >= 23 || asHeaders.indexOf('Actual Completion Date') > -1;
  var SC_LEADAPPR = scIs23 ? 16 : 15;
  var SC_APPDATE  = scIs23 ? 18 : 17;
  var SC_DONEDATE = scIs23 ? 15 : 14;
  var SC_STATUSDT = 14;
  var SC_DEADLINE = 10;
  var SC_REVTAG   = scIs23 ? 19 : 18;
  var SC_ASSIGNDT = 9;
  Logger.log('SC col structure: '+(scIs23?'23-col':'22-col'));

  var scorecard = getOrCreateScorecard(SCORECARD_TAB, writeScorecardHeaders);
  var lastScRow = scorecard.getLastRow();
  if (lastScRow > 1) scorecard.getRange(2, 1, lastScRow-1, 13).clearContent();

  for (var ti = 1; ti < tRows.length; ti++) {
    var name   = String(tRows[ti][0] || '').trim();
    var role   = String(tRows[ti][1] || '').trim();
    var wkTgt  = parseFloat(tRows[ti][2]) || 50;
    var active = String(tRows[ti][5] || '').trim().toLowerCase();
    if (!name || active === 'no') continue;

    var approvedPts = 0;

    // Source B: TASK_ASSIGNMENTS — filter by SelfStatusDate so work done in the week
    // counts even if lead approves after the weekend
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var selfDoneDate = cellDate(r[SC_DONEDATE]) || cellDate(r[SC_STATUSDT]); // P ActualCompletion → O SelfStatus fallback
      if (String(r[3]           || '').trim() === name &&
          String(r[SC_LEADAPPR] || '').trim() === 'Yes' &&
          selfDoneDate >= monStr && selfDoneDate <= sunStr)
        approvedPts += parseFloat(r[8]) || 0; // I WeightedPoints
    });

    // 2. On-time rate
    // TASK_ASSIGNMENTS: AssignedTo=D(3) SelfStatus=N(13) SelfDoneDate=O(14) Deadline=K(10)
    var tasksDone = 0, onTime = 0;
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var rDone = cellDate(r[SC_DONEDATE]) || cellDate(r[SC_STATUSDT]);
      if (String(r[3] || '').trim() === name &&
          String(r[13] || '').trim() === 'Done' &&
          rDone >= monStr && rDone <= sunStr) {
        tasksDone++;
        var rDead = cellDate(r[10]);
        if (rDead && rDone <= rDead) onTime++;
      }
    });
    var onTimeRate = tasksDone > 0 ? Math.round(onTime / tasksDone * 100) : 100;

    // 3. Revision rate this month
    // TASK_ASSIGNMENTS: LeadApproved=P(15) AssignedDate=J(9) RevisionTag=S(18)
    var approvedTasks = 0, revisions = 0;
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var rDate = String(r[9] || '').substring(0, 10);
      if (String(r[3] || '').trim() === name && rDate >= monthStartStr) {
        if (String(r[SC_LEADAPPR] || '').trim() === 'Yes') approvedTasks++;
        if (String(r[SC_REVTAG]   || '').trim()) revisions++;
      }
    });
    var revisionRate = approvedTasks > 0 ? Math.round(revisions / approvedTasks * 100) : 0;
    var qualityScore = Math.max(0, 10 - revisions * rp.internal);

    // 4. DPR streak + attendance
    // DAILY_SUMMARY: Date=A(0) Time=B(1) Member=C(2) Email=D(3) ArrivedOnTime=E(4)
    var memberSum = [];
    sumRows.forEach(function(r, i) {
      if (i === 0) return;
      var rDate = String(r[0] || '').substring(0, 10);
      if (String(r[2] || '').trim() === name && rDate >= monStr && rDate <= sunStr)
        memberSum.push({date:rDate, ontime:String(r[4] || '').trim()});
    });

    // Working days Mon–Sat in the week
    var workDays = 0;
    for (var d = new Date(mon); d <= sun; d.setDate(d.getDate() + 1))
      if (d.getDay() !== 0) workDays++;

    // Consecutive streak back from last working day
    var dprDates = memberSum.map(function(r) { return r.date; });
    var streak = 0, chk = new Date(sun);
    while (chk >= mon) {
      if (chk.getDay() !== 0) {
        if (dprDates.indexOf(dateStr(chk)) > -1) streak++;
        else break;
      }
      chk.setDate(chk.getDate() - 1);
    }

    var dprDays    = memberSum.length;
    var onTimeDays = memberSum.filter(function(r) { return r.ontime === 'Yes'; }).length;
    var dprScore   = workDays > 0 ? Math.round(dprDays    / workDays * 100) : 0;
    var attScore   = workDays > 0 ? Math.round(onTimeDays / workDays * 100) : 0;

    // 5. Monthly pts — same two sources
    var monthlyPts = 0;

    // TASK_LOG source removed — all tasks now in TASK_ASSIGNMENTS

    // Source B: TASK_ASSIGNMENTS — use ActualCompletion → SelfStatusDate for month attribution
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var selfDoneDate = cellDate(r[SC_DONEDATE]) || cellDate(r[SC_STATUSDT]);
      if (String(r[3]           || '').trim() === name &&
          String(r[SC_LEADAPPR] || '').trim() === 'Yes' &&
          selfDoneDate >= monthStartStr)
        monthlyPts += parseFloat(r[8]) || 0;
    });
    var monthlyBenchmark = wkTgt * 4;

    // 6. Overdue tasks (live)
    var overdue = asRows.filter(function(r, i) {
      if (i === 0) return false;
      var rDead = cellDate(r[SC_DEADLINE]);
      return String(r[3]              || '').trim() === name &&
             rDead && rDead < today &&
             String(r[13]             || '').trim() !== 'Done' &&
             String(r[SC_LEADAPPR]    || '').trim() !== 'Yes';
    }).length;

    // 7. Composite score
    var ptScore   = Math.min(100, Math.round(approvedPts / wkTgt * 100));
    var composite = Math.round(
      (ptScore        * sw.points     / 100) +
      (onTimeRate     * sw.ontime     / 100) +
      (dprScore       * sw.dpr        / 100) +
      (attScore       * sw.attendance / 100) +
      (qualityScore * 10 * sw.quality / 100)
    );
    composite = Math.min(100, Math.max(0, composite));
    var status = composite >= 85 ? 'On Track'
               : composite >= 70 ? 'Needs Attention'
               : composite >= 50 ? 'At Risk' : 'Lagging';

    // Write to TEAM_SCORECARD (append row — historical)
    scorecard.appendRow([
      name,
      role,
      wkTgt,
      Math.round(approvedPts * 10) / 10,
      onTimeRate + '%',
      revisionRate + '%',
      streak,
      overdue,
      Math.round(monthlyPts * 10) / 10,
      monthlyBenchmark,
      Math.round(monthlyPts / monthlyBenchmark * 100) + '%',
      status,
      monStr + ' → ' + sunStr,
    ]);

    // Write to individual member tab
    writeMemberTab(sdb(), name, {
      week:monStr, role:role, wkTgt:wkTgt,
      approvedPts:Math.round(approvedPts*10)/10,
      onTimeRate:onTimeRate, revisionRate:revisionRate,
      streak:streak, overdue:overdue,
      monthlyPts:Math.round(monthlyPts*10)/10,
      monthlyBenchmark:monthlyBenchmark,
      composite:composite, status:status,
    });

    Logger.log(name + ': ' + composite + '% (' + status + ') pts=' + Math.round(approvedPts*10)/10 + '/' + wkTgt);
  }

  Logger.log('=== Scorecard complete ===');
}

function writeScorecardHeaders(s) {
  var h = ['Member','Role','Weekly Pts Target','Approved Pts This Week',
           'On-Time Rate','Revision Rate','DPR Streak (days)','Overdue Tasks',
           'Monthly Pts','Monthly Benchmark','Monthly Score %','Status','Week Range'];
  var r = s.getRange(1, 1, 1, h.length);
  r.setValues([h]);
  r.setBackground('#1F3A5F'); r.setFontColor('#FFFFFF');
  r.setFontWeight('bold'); r.setFontSize(10);
  s.setFrozenRows(1);
  [140,120,120,140,90,90,110,90,90,120,110,100,160]
    .forEach(function(w, i) { s.setColumnWidth(i+1, w); });
}

function writeMemberTab(s, name, data) {
  var parts   = name.trim().split(' ');
  var tabName = (parts[0] + (parts.length > 1 ? ' ' + parts[1].charAt(0) + '.' : '')).substring(0, 30);
  var sheet   = s.getSheetByName(tabName);

  if (!sheet) {
    sheet = s.insertSheet(tabName);
    sheet.getRange(1,1).setValue(name).setFontWeight('bold').setFontSize(13).setFontColor('#1F3A5F');
    sheet.getRange(2,1).setValue(data.role).setFontSize(10).setFontColor('#888888');
    sheet.getRange(4,1).setValue('WEEKLY SCORES')
      .setBackground('#1F3A5F').setFontColor('#FFFFFF').setFontWeight('bold');
    var wh = ['Week Of','Approved Pts','Target','On-Time %','Revision %','DPR Streak','Overdue','Score %','Status'];
    var wr = sheet.getRange(5, 1, 1, wh.length);
    wr.setValues([wh]); wr.setBackground('#E8EDF4').setFontWeight('bold').setFontSize(9);
    sheet.setFrozenRows(5);
    sheet.getRange(4,11).setValue('MONTHLY SUMMARY')
      .setBackground('#1F3A5F').setFontColor('#FFFFFF').setFontWeight('bold');
    var mh = ['Month','Monthly Pts','Benchmark','Score %'];
    sheet.getRange(5, 11, 1, mh.length).setValues([mh])
      .setBackground('#E8EDF4').setFontWeight('bold').setFontSize(9);
  }

  sheet.appendRow([
    data.week, data.approvedPts, data.wkTgt,
    data.onTimeRate + '%', data.revisionRate + '%',
    data.streak, data.overdue, data.composite + '%', data.status,
  ]);

  var newRow = sheet.getLastRow();
  var colors = {'On Track':'#EAF3EA','Needs Attention':'#FDF3E3',
                'At Risk':'#FDEAEA','Lagging':'#FDEAEA'};
  sheet.getRange(newRow, 9).setBackground(colors[data.status] || '#FFFFFF');

  // Monthly summary — update or append
  var monthStr = data.week.substring(0, 7);
  var vals = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 5; i < vals.length; i++) {
    if (String(vals[i][10] || '').substring(0, 7) === monthStr) {
      found = true;
      sheet.getRange(i+1, 12).setValue(data.monthlyPts);
      sheet.getRange(i+1, 14).setValue(
        Math.round(data.monthlyPts / data.monthlyBenchmark * 100) + '%'
      );
      break;
    }
  }
  if (!found) {
    sheet.appendRow(
      Array(10).fill('').concat([
        monthStr + '-01',
        data.monthlyPts,
        data.monthlyBenchmark,
        Math.round(data.monthlyPts / data.monthlyBenchmark * 100) + '%'
      ])
    );
  }
}

// ════════════════════════════════════════════════════════════════
// 8 PM EMAIL — daily approval reminder
// ════════════════════════════════════════════════════════════════
function filterTasksForLead(tasks, scope) {
  if (!scope || scope.trim().toLowerCase() === 'all projects') return tasks;
  var allowed = scope.split(',').map(function(s) { return s.trim().toLowerCase(); });
  return tasks.filter(function(t) {
    return allowed.indexOf(t.project.trim().toLowerCase()) > -1;
  });
}

function sendDailyApprovalEmail() {
  var cfg   = readConfig();
  var leads = cfg.leads || [];
  if (!leads.length) { Logger.log('No leads in CONFIG'); return; }

  var pending = getPendingTasks();
  var tasks   = pending.tasks || [];
  if (!tasks.length) { Logger.log('No pending tasks — skipping email'); return; }

  var today = new Date().toLocaleDateString('en-IN',
    {weekday:'long', day:'numeric', month:'long'});

  leads.forEach(function(lead) {
    if (!lead.email) return;
    var lt = filterTasksForLead(tasks, lead.scope);
    if (!lt.length) { Logger.log('No tasks for ' + lead.name); return; }

    var byDate = {};
    lt.forEach(function(t) {
      if (!byDate[t.date]) byDate[t.date] = {};
      if (!byDate[t.date][t.member]) byDate[t.date][t.member] = [];
      byDate[t.date][t.member].push(t);
    });

    var subject = 'IDS · ' + lt.length + ' task' + (lt.length !== 1 ? 's' : '') +
                  ' pending approval — ' + today;
    GmailApp.sendEmail(lead.email, subject, '', {htmlBody: buildEmailBody(lt, byDate, today)});
    Logger.log('Email sent → ' + lead.name + ' (' + lt.length + ' tasks)');
  });
}

function buildEmailBody(tasks, byDate, today) {
  var totalPts = tasks.reduce(function(s, t) { return s + (parseFloat(t.pts) || 0); }, 0);
  var h = '<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto">';
  h += '<div style="background:#1F3A5F;padding:20px 24px;border-radius:8px 8px 0 0">';
  h += '<p style="color:#fff;font-size:16px;font-weight:bold;margin:0">Ideaform Design Studio</p>';
  h += '<p style="color:rgba(255,255,255,.55);font-size:12px;margin:4px 0 0">Pending Approvals · ' + today + '</p></div>';
  h += '<div style="background:#162d4a;padding:12px 24px;display:flex;gap:32px">';
  h += '<div style="text-align:center"><p style="color:#fff;font-size:20px;font-weight:bold;margin:0">' + tasks.length + '</p><p style="color:rgba(255,255,255,.5);font-size:10px;margin:2px 0 0;text-transform:uppercase">Tasks pending</p></div>';
  h += '<div style="text-align:center"><p style="color:#fff;font-size:20px;font-weight:bold;margin:0">' + Object.keys(byDate).length + '</p><p style="color:rgba(255,255,255,.5);font-size:10px;margin:2px 0 0;text-transform:uppercase">Days outstanding</p></div>';
  h += '<div style="text-align:center"><p style="color:#fff;font-size:20px;font-weight:bold;margin:0">' + totalPts.toFixed(1) + '</p><p style="color:rgba(255,255,255,.5);font-size:10px;margin:2px 0 0;text-transform:uppercase">Pts unconfirmed</p></div></div>';
  h += '<div style="background:#fff;padding:20px 24px;border:1px solid #E2DFD8;border-top:none">';

  Object.keys(byDate).sort().reverse().forEach(function(date) {
    var isToday = date === dateStr();
    h += '<p style="font-size:11px;font-weight:bold;color:#9E9B94;text-transform:uppercase;margin:0 0 8px">' + date + (isToday ? ' — Today' : '') + '</p>';
    Object.keys(byDate[date]).forEach(function(member) {
      var mt   = byDate[date][member];
      var mPts = mt.reduce(function(s, t) { return s + (parseFloat(t.pts) || 0); }, 0);
      h += '<div style="margin-bottom:10px;padding:12px 14px;background:#F9F8F5;border-radius:6px;border:1px solid #E2DFD8">';
      h += '<div style="display:flex;justify-content:space-between;margin-bottom:8px">';
      h += '<p style="font-size:13px;font-weight:bold;color:#1F3A5F;margin:0">' + member + '</p>';
      h += '<p style="font-size:11px;color:#1F3A5F;font-weight:bold;margin:0">' + mPts.toFixed(1) + ' pts</p></div>';
      mt.forEach(function(t) {
        h += '<div style="display:flex;justify-content:space-between;margin:4px 0">';
        h += '<p style="font-size:12px;color:#6B6860;margin:0">· ' + t.taskType +
             (t.area    ? ' — ' + t.area    : '') +
             (t.drawing ? ' (' + t.drawing + ')' : '') + '</p>';
        h += '<p style="font-size:12px;color:#1F3A5F;font-weight:bold;margin:0">' + t.pts + 'p</p></div>';
      });
      h += '</div>';
    });
  });

  h += '<div style="margin-top:20px;text-align:center">';
  h += '<a href="' + APPROVAL_FORM_URL + '" style="display:inline-block;padding:13px 36px;background:#1F3A5F;color:#fff;text-decoration:none;border-radius:6px;font-size:14px;font-weight:bold">Review &amp; Approve Tasks →</a></div>';
  h += '<p style="font-size:11px;color:#9E9B94;margin:16px 0 0;text-align:center">Ideaform Design Studio · Daily Progress System</p>';
  h += '</div></div>';
  return h;
}

// ════════════════════════════════════════════════════════════════
// ARCHIVING — monthly, 1st of month 2 AM
//
// TASK_LOG:        Date=B(idx 1), LeadApproved=K(idx 10) — archive if >90d & decided
// TASK_ASSIGNMENTS: AssignedDate=J(idx 9), LeadApproved=P(idx 15) — archive if >90d & decided
// DAILY_SUMMARY:   Date=A(idx 0) — archive if >90d
// ════════════════════════════════════════════════════════════════
function archiveSheet(srcName, archName, dateIdx, approvedIdx) {
  var s   = db();
  var src = s.getSheetByName(srcName);
  if (!src || src.getLastRow() < 2) { Logger.log('Nothing to archive from ' + srcName); return; }

  var arch = s.getSheetByName(archName);
  if (!arch) {
    arch = s.insertSheet(archName);
    var hRow = src.getRange(1, 1, 1, src.getLastColumn()).getValues();
    arch.getRange(1, 1, 1, hRow[0].length).setValues(hRow);
    arch.getRange(1, 1, 1, hRow[0].length)
      .setBackground('#1F3A5F').setFontColor('#FFFFFF').setFontWeight('bold');
    arch.setFrozenRows(1);
  }

  var cutoff = new Date(); cutoff.setDate(cutoff.getDate() - DAYS_BEFORE_ARCH);
  var cutStr = dateStr(cutoff);
  var lastRow = src.getLastRow();
  var data = src.getRange(2, 1, lastRow-1, src.getLastColumn()).getValues();
  var toArch = [], toDel = [];

  data.forEach(function(row, i) {
    var rDate = cellDate(row[dateIdx]);
    var rAppr = approvedIdx >= 0 ? String(row[approvedIdx] || '').trim() : '';
    // Only archive completed AND approved tasks (LeadApproved=Yes)
    // Ongoing, unapproved, or delayed tasks stay regardless of age
    var ok    = rDate && rDate <= cutStr &&
                (approvedIdx < 0 || rAppr === 'Yes');
    if (ok) { toArch.push(row); toDel.push(i + 2); }
  });

  if (!toArch.length) { Logger.log('Nothing old enough in ' + srcName); return; }
  var al = arch.getLastRow();
  arch.getRange(al+1, 1, toArch.length, toArch[0].length).setValues(toArch);
  toDel.reverse().forEach(function(r) { src.deleteRow(r); });
  Logger.log('Archived ' + toArch.length + ' rows from ' + srcName + ' → ' + archName);
}

function runMonthlyArchive() {
  Logger.log('=== Monthly Archive: ' + new Date().toISOString() + ' ===');
  archiveSheet(ASSIGN_TAB,  'ARCHIVE_ASSIGNMENTS',  9, 16); // AssignedDate=J(9), LeadApproved=Q(16)
  archiveSheet(SUMMARY_TAB, 'ARCHIVE_SUMMARY',      0, -1); // Date=A(0), no approval filter
  Logger.log('=== Archive complete ===');
}

// ════════════════════════════════════════════════════════════════
// TRIGGERS — set up manually in Apps Script console:
//
// Function                  | Type        | Schedule
// sendDailyApprovalEmail    | Day timer   | 8pm–9pm daily
// calculateWeeklyScorecard  | Week timer  | Monday 6am–7am
// runMonthlyArchive         | Month timer | Day 1, 2am–3am
// ════════════════════════════════════════════════════════════════


// ════════════════════════════════════════════════════════════════
// SITE VISITS TAB — builds project-wise visit summary
// Reads from TASK_LOG — visit rows identified by Task Type containing
// "site visit", "client meeting", or "visit" (case-insensitive)
//
// TASK_LOG cols: Date=B(1) Member=C(2) Project=D(3) Discipline=E(4)
//   TaskType=F(5) Area=G(6) Drawing=H(7) Units=I(8) Pts=J(9)
//   LeadApproved=K(10)
//
// Runs daily at 7 AM via trigger
// ════════════════════════════════════════════════════════════════

var SITE_VISITS_TAB = 'SITE_VISITS';

// Health thresholds (days)
var VISIT_GREEN  = 7;
var VISIT_AMBER  = 14;
var VISIT_ORANGE = 30;
var MEETING_GREEN = 7;
var MEETING_AMBER = 30;

function buildSiteVisitsTab() {
  Logger.log('=== Building Site Visits Tab: ' + new Date().toISOString() + ' ===');

  var s        = db();
  var asSheet  = s.getSheetByName(ASSIGN_TAB);
  var projSheet= s.getSheetByName(PROJECTS_TAB);

  if (!asSheet)   { Logger.log('TASK_ASSIGNMENTS not found'); return; }
  if (!projSheet) { Logger.log('PROJECTS tab not found'); return; }

  // ── Load visit rows from TASK_ASSIGNMENTS (Done + Approved) ──
  var asRows  = asSheet.getDataRange().getValues();
  var asHdrs  = asRows[0] ? asRows[0].map(function(h){return String(h||'').trim();}) : [];
  var aIs23   = asHdrs.length >= 23 || asHdrs.indexOf('Actual Completion Date') > -1;
  var A_LAPPR = aIs23 ? 16 : 15;
  var A_ACTDT = aIs23 ? 15 : -1;
  var A_STATDT= 14;

  var visits  = [];
  var today   = new Date(); today.setHours(0,0,0,0);

  for (var i = 1; i < asRows.length; i++) {
    var r        = asRows[i];
    var taskType = String(r[4] || '').trim(); // E Stage/TaskType
    var ttLow    = taskType.toLowerCase();
    var isVisit  = ttLow.indexOf('site visit') > -1 || ttLow.indexOf('site supervision') > -1;
    var isMeeting= ttLow.indexOf('client meeting') > -1 || ttLow.indexOf('client + site') > -1;
    if (!isVisit && !isMeeting) continue;

    // Only include Done + Approved visits for history
    var selfStatus   = String(r[13]         || '').trim(); // N
    var leadApproved = String(r[A_LAPPR]    || '').trim(); // Q
    if (selfStatus !== 'Done' || leadApproved !== 'Yes') continue;

    var rawDate  = (A_ACTDT > -1 ? r[A_ACTDT] : null) || r[A_STATDT]; // P or O
    var visitDate= rawDate instanceof Date ? rawDate : new Date(rawDate);
    if (isNaN(visitDate.getTime())) continue;
    visitDate.setHours(0,0,0,0);

    visits.push({
      date      : visitDate,
      dateStr   : cellDate(rawDate),
      project   : String(r[2] || '').trim(),  // C ProjectName
      member    : String(r[3] || '').trim(),  // D AssignedTo
      taskType  : taskType,
      discipline: '',                          // not stored in new structure
      units     : parseFloat(r[7]) || 1,      // H Units
      pts       : parseFloat(r[8]) || 0,      // I WeightedPts
      notes     : String(r[20] || '').trim(), // U Notes
      isVisit   : isVisit,
      isMeeting : isMeeting,
    });
  }

  Logger.log('Total visit rows found: ' + visits.length);

  // ── Load project list ─────────────────────────────────────────
  var pRows    = projSheet.getDataRange().getValues();
  var projects = [];
  for (var j = 1; j < pRows.length; j++) {
    var pName = String(pRows[j][1] || '').trim();
    var pStat = String(pRows[j][11]|| '').trim();
    if (pName) projects.push({name:pName, status:pStat});
  }

  // ── Group visits by project ───────────────────────────────────
  var byProject = {};
  visits.forEach(function(v) {
    if (!v.project) return;
    if (!byProject[v.project]) byProject[v.project] = {siteVisits:[], meetings:[]};
    if (v.isVisit)   byProject[v.project].siteVisits.push(v);
    if (v.isMeeting) byProject[v.project].meetings.push(v);
  });

  // ── Get or create SITE_VISITS tab ────────────────────────────
  var sheet = s.getSheetByName(SITE_VISITS_TAB);
  if (sheet) sheet.clearContents();
  else sheet = s.insertSheet(SITE_VISITS_TAB);

  // ── SECTION 1: Project Summary ────────────────────────────────
  var ACCENT = '#1F3A5F', WHITE = '#FFFFFF', SUBHD = '#E8EDF4';

  // Title
  sheet.getRange(1,1).setValue('IDEAFORM DESIGN STUDIO — SITE VISITS & CLIENT MEETINGS');
  sheet.getRange(1,1).setFontWeight('bold').setFontSize(13).setFontColor(ACCENT);
  sheet.getRange(1,9).setValue('Last updated: ' + nowStr());
  sheet.getRange(1,9).setFontSize(9).setFontColor('#888888').setHorizontalAlignment('right');

  sheet.getRange(2,1).setValue('As of ' + new Date().toLocaleDateString('en-IN',
    {weekday:'long', day:'numeric', month:'long', year:'numeric'}));
  sheet.getRange(2,1).setFontSize(10).setFontColor('#888888');

  // Summary headers (row 4)
  var sumHeaders = [
    'Project','Status',
    'Last Site Visit','Days Ago',
    'Last Client Meeting','Days Ago',
    'Last Attendee',
    'Total Visits','Total Meetings',
    'Avg Visit Gap','Visit Health','Meeting Health'
  ];
  var hRange = sheet.getRange(4, 1, 1, sumHeaders.length);
  hRange.setValues([sumHeaders]);
  hRange.setBackground(ACCENT).setFontColor(WHITE).setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(4);

  // Column widths
  var widths = [200,90,110,70,130,70,140,90,110,100,110,120];
  widths.forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  sheet.setRowHeight(4, 28);

  // ── Write project summary rows ────────────────────────────────
  var dataRows = [];
  var healthCols = []; // track rows needing health colour

  projects.forEach(function(proj) {
    var pData = byProject[proj.name] || {siteVisits:[], meetings:[]};

    // Site visits
    var siteVisits = pData.siteVisits.sort(function(a,b){ return b.date-a.date; });
    var lastVisit  = siteVisits.length > 0 ? siteVisits[0] : null;
    var lastVisitStr = lastVisit ? lastVisit.dateStr : '';
    var visitDaysAgo = lastVisit ?
      Math.floor((today - lastVisit.date) / 86400000) : null;

    // Average visit gap
    var avgGap = '';
    if (siteVisits.length > 1) {
      var sortedAsc = siteVisits.slice().reverse();
      var totalGap = 0;
      for (var k = 1; k < sortedAsc.length; k++) {
        totalGap += (sortedAsc[k].date - sortedAsc[k-1].date) / 86400000;
      }
      avgGap = Math.round(totalGap / (sortedAsc.length - 1));
    }

    // Client meetings
    var meetings    = pData.meetings.sort(function(a,b){ return b.date-a.date; });
    var lastMeeting = meetings.length > 0 ? meetings[0] : null;
    var lastMeetStr = lastMeeting ? lastMeeting.dateStr : '';
    var meetDaysAgo = lastMeeting ?
      Math.floor((today - lastMeeting.date) / 86400000) : null;

    // Health status
    var visitHealth, meetHealth;
    if (!lastVisit) {
      visitHealth = 'No visits yet';
    } else if (visitDaysAgo <= VISIT_GREEN) {
      visitHealth = 'Green';
    } else if (visitDaysAgo <= VISIT_AMBER) {
      visitHealth = 'Amber';
    } else if (visitDaysAgo <= VISIT_ORANGE) {
      visitHealth = 'Orange';
    } else {
      visitHealth = 'Red';
    }

    if (!lastMeeting) {
      meetHealth = 'No meetings yet';
    } else if (meetDaysAgo <= MEETING_GREEN) {
      meetHealth = 'Green';
    } else if (meetDaysAgo <= MEETING_AMBER) {
      meetHealth = 'Amber';
    } else {
      meetHealth = 'Red';
    }

    dataRows.push([
      proj.name,
      proj.status,
      lastVisitStr,
      visitDaysAgo !== null ? visitDaysAgo : '—',
      lastMeetStr,
      meetDaysAgo !== null ? meetDaysAgo : '—',
      lastVisit ? lastVisit.member : (lastMeeting ? lastMeeting.member : '—'),
      siteVisits.length,
      meetings.length,
      avgGap !== '' ? avgGap + ' days' : (siteVisits.length === 1 ? 'Only 1 visit' : '—'),
      visitHealth,
      meetHealth,
    ]);
  });

  if (dataRows.length > 0) {
    var dataRange = sheet.getRange(5, 1, dataRows.length, 12);
    dataRange.setValues(dataRows);

    // Format date columns
    sheet.getRange(5, 3, dataRows.length, 1).setNumberFormat('dd-mmm-yyyy');
    sheet.getRange(5, 5, dataRows.length, 1).setNumberFormat('dd-mmm-yyyy');

    // Alternating row colours
    for (var ri = 0; ri < dataRows.length; ri++) {
      var bg = ri % 2 === 0 ? '#FFFFFF' : '#F9F8F5';
      sheet.getRange(ri+5, 1, 1, 12).setBackground(bg);
      sheet.getRange(ri+5, 1, 1, 12).setFontSize(10);
    }

    // Conditional formatting for health columns (K=11, L=12)
    var healthMap = {
      'Green'          : {bg:'#EAF3EA', text:'#2D6A2D'},
      'Amber'          : {bg:'#FDF3E3', text:'#7A4F0A'},
      'Orange'         : {bg:'#FEF0E0', text:'#8B4513'},
      'Red'            : {bg:'#FDEAEA', text:'#8B2020'},
      'No visits yet'  : {bg:'#F1F1F1', text:'#888888'},
      'No meetings yet': {bg:'#F1F1F1', text:'#888888'},
    };

    for (var ri2 = 0; ri2 < dataRows.length; ri2++) {
      var vH = dataRows[ri2][10]; // Visit health
      var mH = dataRows[ri2][11]; // Meeting health
      var sheetRow = ri2 + 5;

      if (healthMap[vH]) {
        sheet.getRange(sheetRow, 11).setBackground(healthMap[vH].bg)
          .setFontColor(healthMap[vH].text).setFontWeight('bold');
      }
      if (healthMap[mH]) {
        sheet.getRange(sheetRow, 12).setBackground(healthMap[mH].bg)
          .setFontColor(healthMap[mH].text).setFontWeight('bold');
      }

      // Centre align Days Ago and count cols
      sheet.getRange(sheetRow, 4).setHorizontalAlignment('center');
      sheet.getRange(sheetRow, 6).setHorizontalAlignment('center');
      sheet.getRange(sheetRow, 8).setHorizontalAlignment('center');
      sheet.getRange(sheetRow, 9).setHorizontalAlignment('center');
      sheet.getRange(sheetRow, 10).setHorizontalAlignment('center');
    }
  }

  // ── Summary stats row ─────────────────────────────────────────
  var summaryRow = 5 + dataRows.length + 1;
  var totalVisits   = visits.filter(function(v){ return v.isVisit; }).length;
  var totalMeetings = visits.filter(function(v){ return v.isMeeting; }).length;
  var activeProjs   = dataRows.filter(function(r){ return r[7] > 0 || r[8] > 0; }).length;

  sheet.getRange(summaryRow, 1).setValue('TOTALS');
  sheet.getRange(summaryRow, 1, 1, 12).setBackground(SUBHD).setFontWeight('bold').setFontSize(10);
  sheet.getRange(summaryRow, 8).setValue(totalVisits);
  sheet.getRange(summaryRow, 9).setValue(totalMeetings);
  sheet.getRange(summaryRow, 7).setValue(activeProjs + ' projects with activity');

  // ── SECTION 2: Full Visit Log ─────────────────────────────────
  var logStart = summaryRow + 3;

  sheet.getRange(logStart, 1).setValue('FULL VISIT LOG');
  sheet.getRange(logStart, 1).setFontWeight('bold').setFontSize(12).setFontColor(ACCENT);
  sheet.getRange(logStart, 9).setValue('All site visits and client meetings from TASK_ASSIGNMENTS');
  sheet.getRange(logStart, 9).setFontSize(9).setFontColor('#888888');

  var logHeaders = ['Date','Project','Member','Visit Type','Duration / Type','Points','Notes'];
  var lhRange = sheet.getRange(logStart+1, 1, 1, logHeaders.length);
  lhRange.setValues([logHeaders]);
  lhRange.setBackground(ACCENT).setFontColor(WHITE).setFontWeight('bold').setFontSize(10);
  sheet.setRowHeight(logStart+1, 24);

  // Sort all visits by date descending
  var allVisitsSorted = visits.slice().sort(function(a,b){ return b.date-a.date; });

  var logRows = allVisitsSorted.map(function(v) {
    // Parse duration from taskType if present (e.g. "Site visit · 2–3 hours")
    var parts    = v.taskType.split('·');
    var vType    = parts[0].trim();
    var duration = parts.length > 1 ? parts[1].trim() : '';
    return [
      v.dateStr,
      v.project,
      v.member,
      vType,
      duration,
      v.pts,
      v.notes,
    ];
  });

  if (logRows.length > 0) {
    var logRange = sheet.getRange(logStart+2, 1, logRows.length, 7);
    logRange.setValues(logRows);
    logRange.setFontSize(10);

    // Alternating rows
    for (var li = 0; li < logRows.length; li++) {
      var bg2 = li % 2 === 0 ? '#FFFFFF' : '#F9F8F5';
      sheet.getRange(logStart+2+li, 1, 1, 7).setBackground(bg2);
    }

    // Colour visit type column by type
    for (var li2 = 0; li2 < logRows.length; li2++) {
      var vt = String(logRows[li2][3]).toLowerCase();
      var typeCell = sheet.getRange(logStart+2+li2, 4);
      if (vt.includes('site visit') || vt.includes('site supervision')) {
        typeCell.setBackground('#E8EDF4').setFontColor(ACCENT);
      } else if (vt.includes('client meeting')) {
        typeCell.setBackground('#EAF3EA').setFontColor('#2D6A2D');
      } else if (vt.includes('client meeting + site') || vt.includes('client + site')) {
        typeCell.setBackground('#FDF3E3').setFontColor('#7A4F0A');
      }
    }
  } else {
    sheet.getRange(logStart+2, 1).setValue('No approved visit tasks found in TASK_ASSIGNMENTS yet.');
    sheet.getRange(logStart+2, 1).setFontColor('#888888').setFontStyle('italic');
  }

  Logger.log('SITE_VISITS tab built: ' + dataRows.length + ' projects, ' +
             totalVisits + ' site visits, ' + totalMeetings + ' client meetings');
}

// ════════════════════════════════════════════════════════════════
// calculateCurrentWeekScorecard
// Same as calculateWeeklyScorecard but uses the CURRENT running week
// (Monday of this week → today)
// Run this manually anytime to see live scores mid-week
// ════════════════════════════════════════════════════════════════
function calculateCurrentWeekScorecard() {
  Logger.log('=== Current Week Scorecard: ' + new Date().toISOString() + ' ===');

  var s   = db();
  var cfg = readConfig();
  var sw  = cfg.scoringWeights;
  var rp  = cfg.revPenalties;

  // THIS week: Monday of current week → today
  var mon    = mondayOf(new Date()); // Monday of current week
  var today  = new Date(); today.setHours(23,59,59,0);
  var monStr = dateStr(mon);
  var sunStr = dateStr(today); // up to today, not end of week
  var monthStart    = new Date(mon.getFullYear(), mon.getMonth(), 1);
  var monthStartStr = dateStr(monthStart);
  var todayStr2     = dateStr(new Date());

  Logger.log('Current week: ' + monStr + ' to ' + sunStr);

  // Load all data
  var asSheet  = s.getSheetByName(ASSIGN_TAB);
  var sumSheet = s.getSheetByName(SUMMARY_TAB);
  var tSheet   = s.getSheetByName(TEAM_TAB);
  if (!tSheet) { Logger.log('TEAM tab not found'); return; }

  var asRows  = asSheet  ? asSheet.getDataRange().getValues()  : [];
  var sumRows = sumSheet ? sumSheet.getDataRange().getValues() : [];
  var tRows   = tSheet.getDataRange().getValues();

  // Auto-detect TASK_ASSIGNMENTS column structure
  var asHeaders = asRows[0] ? asRows[0].map(function(h){ return String(h||'').trim(); }) : [];
  var scIs23    = asHeaders.length >= 23 || asHeaders.indexOf('Actual Completion Date') > -1;
  var SC_LEADAPPR = scIs23 ? 16 : 15;
  var SC_APPDATE  = scIs23 ? 18 : 17;
  var SC_DONEDATE = scIs23 ? 15 : 14;
  var SC_STATUSDT = 14;
  var SC_DEADLINE = 10;
  var SC_REVTAG   = scIs23 ? 19 : 18;
  var SC_ASSIGNDT = 9;
  Logger.log('SC col structure: '+(scIs23?'23-col':'22-col'));

  // Write to IDS Team Scorecard sheet — clear and rewrite (live snapshot)
  var scorecard = getOrCreateScorecard(SCORECARD_TAB, writeScorecardHeaders);
  var lastScRow = scorecard.getLastRow();
  if (lastScRow > 1) scorecard.getRange(2, 1, lastScRow-1, 13).clearContent();


  for (var ti = 1; ti < tRows.length; ti++) {
    var name   = String(tRows[ti][0] || '').trim();
    var role   = String(tRows[ti][1] || '').trim();
    var wkTgt  = parseFloat(tRows[ti][2]) || 50;
    var active = String(tRows[ti][5] || '').trim().toLowerCase();
    if (!name || active === 'no') continue;

    // Approved pts this week — TASK_LOG + TASK_ASSIGNMENTS
    // Filter by SelfStatusDate (when work was done) so weekend approvals
    // still count for the week the work was completed
    var approvedPts = 0;

    // TASK_LOG source removed

    // Source B: TASK_ASSIGNMENTS — filter by SelfStatusDate (O=14) so work done
    // Mon-Sat counts even if lead approves on Sunday or next week
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var selfDoneDate = cellDate(r[SC_DONEDATE]) || cellDate(r[SC_STATUSDT]); // P ActualCompletion → O SelfStatus fallback
      if (String(r[3]           || '').trim() === name &&
          String(r[SC_LEADAPPR] || '').trim() === 'Yes' &&
          selfDoneDate >= monStr && selfDoneDate <= sunStr)
        approvedPts += parseFloat(r[8]) || 0; // I WeightedPoints
    });

    // On-time rate
    var tasksDone = 0, onTime = 0;
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var rDone = cellDate(r[SC_DONEDATE]) || cellDate(r[SC_STATUSDT]);
      if (String(r[3]  || '').trim() === name &&
          String(r[13] || '').trim() === 'Done' &&
          rDone >= monStr && rDone <= sunStr) {
        tasksDone++;
        var rDead = cellDate(r[SC_DEADLINE]);
        if (rDead && rDone <= rDead) onTime++;
      }
    });
    var onTimeRate = tasksDone > 0 ? Math.round(onTime / tasksDone * 100) : 100;

    // Revision rate this month
    var approvedTasks = 0, revisions = 0;
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var rDate = cellDate(r[9]);
      if (String(r[3] || '').trim() === name && rDate >= monthStartStr) {
        if (String(r[SC_LEADAPPR] || '').trim() === 'Yes') approvedTasks++;
        if (String(r[SC_REVTAG]   || '').trim()) revisions++;
      }
    });
    var revisionRate = approvedTasks > 0 ? Math.round(revisions / approvedTasks * 100) : 0;
    var qualityScore = Math.max(0, 10 - revisions * rp.internal);

    // DPR streak + attendance — this week only
    // DAILY_SUMMARY: Date=A(0) Time=B(1) Member=C(2) Email=D(3) ArrivedOnTime=E(4)
    var memberSum = [];
    sumRows.forEach(function(r, i) {
      if (i === 0) return;
      var rDate = cellDate(r[0]);
      if (String(r[2] || '').trim() === name && rDate >= monStr && rDate <= sunStr)
        memberSum.push({date:rDate, ontime:String(r[4]||'').trim()});
    });

    // Working days from Monday to today
    var workDays = 0;
    for (var d = new Date(mon); d <= today; d.setDate(d.getDate() + 1))
      if (d.getDay() !== 0) workDays++;

    // Streak
    var dprDates = memberSum.map(function(r){ return r.date; });
    var streak = 0, chk = new Date(today);
    while (cellDate(chk) >= monStr) {
      if (chk.getDay() !== 0) {
        if (dprDates.indexOf(cellDate(chk)) > -1) streak++;
        else break;
      }
      chk.setDate(chk.getDate() - 1);
    }

    var dprDays    = memberSum.length;
    var onTimeDays = memberSum.filter(function(r){ return r.ontime === 'Yes'; }).length;
    var dprScore   = workDays > 0 ? Math.round(dprDays    / workDays * 100) : 0;
    var attScore   = workDays > 0 ? Math.round(onTimeDays / workDays * 100) : 0;

    // Monthly pts
    var monthlyPts = 0;
    // TASK_LOG source removed
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var selfDoneDate = cellDate(r[SC_DONEDATE]) || cellDate(r[SC_STATUSDT]);
      if (String(r[3]           || '').trim() === name &&
          String(r[SC_LEADAPPR] || '').trim() === 'Yes' &&
          selfDoneDate >= monthStartStr)
        monthlyPts += parseFloat(r[8]) || 0;
    });
    var monthlyBenchmark = wkTgt * 4;

    // Overdue tasks
    var overdue = asRows.filter(function(r, i) {
      if (i === 0) return false;
      var rDead = cellDate(r[10]);
      return String(r[3]  || '').trim() === name &&
             rDead && rDead < todayStr2 &&
             String(r[13] || '').trim() !== 'Done' &&
             String(r[15] || '').trim() !== 'Yes';
    }).length;

    // Composite score
    var ptScore   = Math.min(100, Math.round(approvedPts / wkTgt * 100));
    var composite = Math.round(
      (ptScore        * sw.points     / 100) +
      (onTimeRate     * sw.ontime     / 100) +
      (dprScore       * sw.dpr        / 100) +
      (attScore       * sw.attendance / 100) +
      (qualityScore * 10 * sw.quality / 100)
    );
    composite = Math.min(100, Math.max(0, composite));
    var status = composite >= 85 ? 'On Track'
               : composite >= 70 ? 'Needs Attention'
               : composite >= 50 ? 'At Risk' : 'Lagging';

    scorecard.appendRow([
      name, role, wkTgt,
      Math.round(approvedPts * 10) / 10,
      onTimeRate + '%',
      revisionRate + '%',
      streak,
      overdue,
      Math.round(monthlyPts * 10) / 10,
      monthlyBenchmark,
      Math.round(monthlyPts / monthlyBenchmark * 100) + '%',
      status,
      'Live: ' + monStr + ' → ' + todayStr2,
    ]);

    writeMemberTab(sdb(), name, {
      week: monStr + ' (live)', role:role, wkTgt:wkTgt,
      approvedPts: Math.round(approvedPts*10)/10,
      onTimeRate:onTimeRate, revisionRate:revisionRate,
      streak:streak, overdue:overdue,
      monthlyPts: Math.round(monthlyPts*10)/10,
      monthlyBenchmark:monthlyBenchmark,
      composite:composite, status:status,
    });

    Logger.log(name + ': ' + composite + '% (' + status + ') | pts=' +
               Math.round(approvedPts*10)/10 + '/' + wkTgt +
               ' | dprDays=' + dprDays + '/' + workDays +
               ' | streak=' + streak + ' | onTime=' + onTimeDays);
  }

  Logger.log('=== Current Week Scorecard complete ===');
}

// ════════════════════════════════════════════════════════════════
// TRIGGER: calculateWeeklyScorecard — runs SUNDAY 6am (not Monday)
// This gives you scores by Sunday evening to build the report
// Change trigger: Week timer → Sunday → 6am–7am
// (Previously set to Monday — update this in Apps Script triggers)
//
// TRIGGER: buildSiteVisitsTab — runs daily at 7 AM
// Add in Apps Script: Time-driven → Day timer → 7am–8am
// Also run manually anytime: select buildSiteVisitsTab → Run
// ════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════
// VISIT PLANNER SYSTEM
// Tabs: VISIT_PLANNER, VISIT_SCHEDULE, PROJECT_HEALTH
//
// VISIT_PLANNER (you fill once):
//   Project(A) VisitType(B) Assignee(C) Frequency(D)
//   FixedDay(E) FixedTime(F) Active(G) Notes(H)
//
// VISIT_SCHEDULE (script writes daily):
//   Project(A) VisitType(B) Assignee(C) LastVisitDate(D)
//   LastVisitBy(E) NextVisitDate(F) Status(G) TaskID(H)
//   DaysUntil(I) Priority(J)
//
// PROJECT_HEALTH (script writes daily):
//   Project(A) Lead(B) ActiveVisitTypes(C)
//   LastSiteVisitArch(D) DaysSinceArch(E)
//   LastSiteVisitID(F) DaysSinceID(G)
//   LastClientMeeting(H) DaysSinceMeeting(I)
//   TotalVisitsMonth(J) TotalMeetingsMonth(K)
//   NextPlannedVisit(L) NextPlannedMeeting(M)
//   MissedVisits30Days(N)
//   ArchHealth(O) IDHealth(P) MeetingHealth(Q) OverallHealth(R)
// ════════════════════════════════════════════════════════════════

var PLANNER_TAB      = 'VISIT_PLANNER';
var SITE_EXEC_TAB    = 'SITE_EXECUTION';
var SITE_ISSUES_TAB  = 'SITE_ISSUES';
var SITE_WA_TAB      = 'SITE_WA_MESSAGES';
var VSCHED_TAB   = 'VISIT_SCHEDULE';
var PHEALTH_TAB  = 'PROJECT_HEALTH';

var VISIT_FREQ_DAYS = {'Weekly':7,'Fortnightly':14,'Monthly':30};
var DAY_INDEX = {
  'Sunday':0,'Monday':1,'Tuesday':2,'Wednesday':3,
  'Thursday':4,'Friday':5,'Saturday':6
};

// Visit type constants — must match CONFIG and DPR dropdown exactly
var VT_ARCH    = 'Site Visit (Architecture)';
var VT_ID      = 'Site Visit (Interiors)';
var VT_MEETING = 'Client Meeting (Regular)';
var VT_SPECIAL = 'Client Meeting (Special)';

// Health thresholds by frequency
var HEALTH_THRESH = {
  'Weekly'      : {green:7,  amber:14},
  'Fortnightly' : {green:14, amber:28},
  'Monthly'     : {green:30, amber:45},
};
var MEETING_THRESH = {green:7, amber:21};

// ── Low-level helpers ─────────────────────────────────────────
function isVisitType(taskType) {
  var t = String(taskType||'').toLowerCase();
  return t.includes('site visit') || t.includes('client meeting');
}

function normaliseVisitType(taskType) {
  var t = String(taskType||'').toLowerCase();
  if (t.includes('architecture'))     return VT_ARCH;
  if (t.includes('interior'))         return VT_ID;
  if (t.includes('special'))          return VT_SPECIAL;
  if (t.includes('client meeting'))   return VT_MEETING;
  if (t.includes('site visit'))       return VT_ARCH; // default site visit to arch
  return null;
}

function addDaysToStr(dateStr, days) {
  if (!dateStr) return '';
  var d = new Date(dateStr);
  if (isNaN(d.getTime())) return '';
  d.setDate(d.getDate() + days);
  return cellDate(d);
}

function daysDiff(fromStr, toStr) {
  if (!fromStr || !toStr) return null;
  var a = new Date(fromStr), b = new Date(toStr);
  if (isNaN(a)||isNaN(b)) return null;
  return Math.round((b-a)/86400000);
}

function todayStr() { return cellDate(new Date()); }

function nextOccurrenceOfDay(dayName) {
  var target = DAY_INDEX[dayName];
  if (target === undefined) return '';
  var d = new Date(); d.setHours(0,0,0,0);
  var diff = (target - d.getDay() + 7) % 7;
  if (diff === 0) diff = 7;
  d.setDate(d.getDate() + diff);
  return cellDate(d);
}

function getOrMakeTab(name, setupFn) {
  var s = db(), sheet = s.getSheetByName(name);
  if (!sheet) { sheet = s.insertSheet(name); if (setupFn) setupFn(sheet); }
  return sheet;
}

function styleHeader(range) {
  range.setBackground('#1F3A5F').setFontColor('#FFFFFF')
       .setFontWeight('bold').setFontSize(10);
  return range;
}

function healthColour(sheet, row, col, status) {
  var colours = {
    'Green':  {bg:'#EAF3EA', fg:'#2D6A2D'},
    'Amber':  {bg:'#FDF3E3', fg:'#7A4F0A'},
    'Red':    {bg:'#FDEAEA', fg:'#8B2020'},
    'Grey':   {bg:'#F1F1F1', fg:'#888888'},
    'Missed': {bg:'#FDEAEA', fg:'#8B2020'},
    'Overdue':{bg:'#FDF3E3', fg:'#7A4F0A'},
  };
  var c = colours[status] || {bg:'#FFFFFF', fg:'#000000'};
  sheet.getRange(row, col)
    .setBackground(c.bg).setFontColor(c.fg).setFontWeight('bold');
}

function altRow(sheet, row, cols) {
  var bg = (row % 2 === 0) ? '#F9F8F5' : '#FFFFFF';
  sheet.getRange(row, 1, 1, cols).setBackground(bg).setFontSize(10);
}

// ── Tab setup ─────────────────────────────────────────────────
function setupPlannerTab(sheet) {
  sheet.getRange(1,1).setValue(
    'VISIT PLANNER — Add one row per project per visit type. ' +
    'Script reads this daily to generate schedule and tasks.')
    .setFontWeight('bold').setFontSize(11).setFontColor('#1F3A5F');
  sheet.getRange(1,1,1,8).merge();
  var h = ['Project','Visit Type','Assignee','Frequency',
           'Fixed Day','Fixed Time','Active','Last Site Visit Date','Notes'];
  var hr = sheet.getRange(2,1,1,h.length);
  hr.setValues([h]); styleHeader(hr); sheet.setFrozenRows(2);
  sheet.getRange(3,1).setValue(
    'Visit Types: Site Visit (Architecture) | Site Visit (Interiors) | ' +
    'Client Meeting (Regular)    Frequency: Weekly | Fortnightly | Monthly | None    ' +
    'Col H: Fill last actual visit date for history before system go-live')
    .setFontColor('#888888').setFontStyle('italic').setFontSize(9);
  sheet.getRange(3,1,1,9).merge();
  [200,180,140,100,100,90,70,130,220]
    .forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  sheet.setRowHeight(1,28); sheet.setRowHeight(3,20);
}

function setupScheduleTab(sheet) {
  sheet.getRange(1,1).setValue(
    'VISIT SCHEDULE — Auto-generated daily at 7 AM. ' +
    'You may override Next Visit Date (col F) manually if needed.')
    .setFontWeight('bold').setFontSize(11).setFontColor('#1F3A5F');
  sheet.getRange(1,1,1,10).merge();
  var h = ['Project','Visit Type','Assignee','Last Visit Date','Last Visit By',
           'Next Visit Date','Status','Task ID','Days Until / Overdue','Priority'];
  var hr = sheet.getRange(2,1,1,h.length);
  hr.setValues([h]); styleHeader(hr); sheet.setFrozenRows(2);
  [200,180,140,120,140,120,90,120,120,80]
    .forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  sheet.setRowHeight(1,28);
}

function setupHealthTab(sheet) {
  sheet.getRange(1,1).setValue(
    'PROJECT HEALTH — Auto-refreshed daily. Read only. ' +
    'Shows visit frequency, recency, and health per project.')
    .setFontWeight('bold').setFontSize(11).setFontColor('#1F3A5F');
  sheet.getRange(1,1,1,18).merge();
  var h = [
    'Project','Lead','Active Visit Types',
    'Last Site Visit (Arch)','Days Since',
    'Last Site Visit (ID)','Days Since',
    'Last Client Meeting','Days Since',
    'Visits This Month','Meetings This Month',
    'Next Planned Visit','Next Planned Meeting',
    'Missed Visits (30d)',
    'Arch Health','ID Health','Meeting Health','Overall Health'
  ];
  var hr = sheet.getRange(2,1,1,h.length);
  hr.setValues([h]); styleHeader(hr); sheet.setFrozenRows(2);
  [200,130,160,130,80,130,80,130,80,100,110,130,130,100,90,80,100,110]
    .forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  sheet.setRowHeight(1,28);
}

// ── Load VISIT_PLANNER cadence ────────────────────────────────
function loadCadence(planSheet) {
  var rows = planSheet.getLastRow() > 3
    ? planSheet.getRange(4,1,planSheet.getLastRow()-3,9).getValues()
    : [];
  var cadence = []; // array of cadence entries
  rows.forEach(function(r) {
    var proj   = String(r[0]||'').trim();
    var vType  = String(r[1]||'').trim();
    var active = String(r[6]||'Yes').trim().toLowerCase();
    if (!proj || !vType || active === 'no') return;
    cadence.push({
      project      : proj,
      visitType    : vType,
      assignee     : String(r[2]||'').trim(),
      frequency    : String(r[3]||'').trim(),
      fixedDay     : String(r[4]||'').trim(),
      fixedTime    : String(r[5]||'').trim(),
      // Col G = Active (already checked above)
      manualLastDate: cellDate(r[7]),        // H Last Site Visit Date (manual history)
      notes        : String(r[8]||'').trim(), // I Notes
    });
  });
  return cadence;
}

// ── Load visit history from TASK_LOG ─────────────────────────
// Returns: {project: {visitType: [{date, member, pts}]}}
function loadVisitHistory() {
  // Read from TASK_ASSIGNMENTS — completed (Done + Approved) visit tasks
  var sheet = db().getSheetByName(ASSIGN_TAB);
  var history = {};
  if (!sheet) return history;
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0] ? rows[0].map(function(h){ return String(h||'').trim(); }) : [];
  var is23 = headers.length >= 23;

  for (var i = 1; i < rows.length; i++) {
    var r        = rows[i];
    var proj     = String(r[2]||'').trim();  // C ProjectName
    var tType    = String(r[4]||'').trim();  // E TaskType
    var member   = String(r[3]||'').trim();  // D AssignedTo
    var pts      = parseFloat(r[8])||0;      // I WeightedPts
    var selfStat = String(r[13]||'').trim(); // N SelfStatus
    var approved = String(r[is23?16:15]||'').trim(); // Q LeadApproved
    if (!proj || !isVisitType(tType)) continue;
    if (selfStat !== 'Done' || approved !== 'Yes') continue;
    var dStr = cellDate(r[15]) || cellDate(r[14]) || cellDate(r[is23?18:17]) || '';
    if (!dStr) continue;
    var vType = normaliseVisitType(tType) || tType;
    if (!history[proj]) history[proj] = {};
    if (!history[proj][vType]) history[proj][vType] = [];
    history[proj][vType].push({date:dStr, member:member, pts:pts});
  }
  Object.keys(history).forEach(function(p) {
    Object.keys(history[p]).forEach(function(vt) {
      history[p][vt].sort(function(a,b){ return b.date.localeCompare(a.date); });
    });
  });
  Logger.log('loadVisitHistory: '+Object.keys(history).length+' projects with history');
  return history;
}

// ── Load open visit tasks from TASK_ASSIGNMENTS ───────────────
// Returns: {project+type key: {taskId, deadline}}
function loadOpenVisitTasks() {
  var sheet = db().getSheetByName(ASSIGN_TAB);
  var open  = {};
  if (!sheet) return open;
  var rows = sheet.getDataRange().getValues();
  // Auto-detect 22-col vs 23-col (ActualCompletionDate at P inserts before LeadApproved)
  var headers = rows[0] ? rows[0].map(function(h){ return String(h||'').trim(); }) : [];
  var is23    = headers.length >= 23 || headers.indexOf('Actual Completion Date') > -1;
  var COL_LAPPR = is23 ? 16 : 15; // Q(16) new / P(15) old

  for (var i = 1; i < rows.length; i++) {
    var r       = rows[i];
    var taskId  = String(r[0]||'').trim();           // A
    var proj    = String(r[2]||'').trim();           // C
    var tType   = String(r[4]||'').trim();           // E
    var deadline= cellDate(r[10]);                   // K
    var status  = String(r[13]||'').trim();          // N SelfStatus
    var appr    = String(r[COL_LAPPR]||'').trim();   // Q LeadApproved
    if (!proj || !deadline || !isVisitType(tType)) continue;
    if (status === 'Done' && appr === 'Yes') continue; // completed — skip
    // Use raw task type as key — consistent with pushVisitTasks
    var assigneeVal = String(r[3]||'').trim(); // D AssignedTo
    var key     = proj + '||' + tType + '||' + assigneeVal;
    var keyBase = proj + '||' + tType;
    if (!open[keyBase]) open[keyBase] = {taskId:taskId, deadline:deadline, assignee:assigneeVal};
    if (!open[key] || deadline < open[key].deadline) {
      open[key] = {taskId:taskId, deadline:deadline, row:i+1};
    }
  }
  return open;
}

// ── Calculate next visit date for one cadence entry ───────────
function calcNextVisitDate(entry, lastVisitDate, openTasks) {
  // Check open task for this specific assignee first, then any assignee
  var assigneeStr = String(entry.assignee||'').split(',')[0].trim();
  // Keys use raw visitType — consistent with loadOpenVisitTasks
  var key     = entry.project + '||' + entry.visitType + '||' + assigneeStr;
  var keyAny  = entry.project + '||' + entry.visitType;

  // Priority 1: existing open task in TASK_ASSIGNMENTS
  var openKey = openTasks[key] ? key : (openTasks[keyAny] ? keyAny : null);
  if (openKey) return {
    date   : openTasks[openKey].deadline,
    taskId : openTasks[openKey].taskId,
    source : 'task',
  };

  var freq    = entry.frequency;
  var freqDays= VISIT_FREQ_DAYS[freq];
  var today   = todayStr();

  // Priority 2: Weekly with fixed day → always next occurrence of that day
  if (freq === 'Weekly' && entry.fixedDay) {
    return {date: nextOccurrenceOfDay(entry.fixedDay), taskId:'', source:'cadence'};
  }

  // Priority 3: Interval-based → last actual visit + frequency
  if (freqDays && lastVisitDate) {
    return {date: addDaysToStr(lastVisitDate, freqDays), taskId:'', source:'cadence'};
  }

  // Priority 4: No visits yet → start from today + frequency
  if (freqDays) {
    return {date: addDaysToStr(today, freqDays), taskId:'', source:'cadence'};
  }

  return {date:'', taskId:'', source:'none'};
}

// ── Build VISIT_SCHEDULE tab ──────────────────────────────────
function buildVisitSchedule(cadence, history, openTasks) {
  var sheet   = getOrMakeTab(VSCHED_TAB, setupScheduleTab);
  var today   = todayStr();
  var dataRows= [];

  cadence.forEach(function(entry) {
    var hist = (history[entry.project] || {})[entry.visitType] || [];
    var lastVisit   = hist.length > 0 ? hist[0] : null;
    // Use TASK_LOG history if available; fall back to manual date from VISIT_PLANNER col H
    var lastDate    = lastVisit ? lastVisit.date : (entry.manualLastDate || '');
    var lastMember  = lastVisit ? lastVisit.member : (entry.manualLastDate ? 'Pre-system record' : '');
    var nextInfo    = calcNextVisitDate(entry, lastDate, openTasks);
    var nextDate    = nextInfo.date;
    var taskId      = nextInfo.taskId;

    // Status
    var status = 'Scheduled';
    var daysUntil = nextDate ? daysDiff(today, nextDate) : null;
    if (!nextDate) {
      status = 'Not scheduled';
    } else if (daysUntil < -VISIT_FREQ_DAYS[entry.frequency]) {
      status = 'Missed';
    } else if (daysUntil < 0) {
      status = 'Overdue';
    }

    var priority = (status === 'Missed' || status === 'Overdue') ? 'High' : 'Medium';

    dataRows.push([
      entry.project,
      entry.visitType,
      entry.assignee,
      lastDate,
      lastMember,
      nextDate,
      status,
      taskId,
      daysUntil !== null ? daysUntil : '—',
      priority,
    ]);
  });

  // Write all rows
  var startRow = 3;
  // Clear old data
  var lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow-startRow+1, 10).clearContent()
      .setBackground('#FFFFFF').setFontColor('#000000').setFontWeight('normal');
  }
  if (dataRows.length === 0) { Logger.log('No cadence rows to write'); return; }

  var range = sheet.getRange(startRow, 1, dataRows.length, 10);
  range.setValues(dataRows);

  // Format date cols
  sheet.getRange(startRow,4,dataRows.length,1).setNumberFormat('dd-mmm-yyyy');
  sheet.getRange(startRow,6,dataRows.length,1).setNumberFormat('dd-mmm-yyyy');

  // Style each row
  dataRows.forEach(function(row, i) {
    var r      = startRow + i;
    var status = row[6];
    altRow(sheet, r, 10);
    // Status cell colour
    var statusCol = 7;
    if      (status === 'Scheduled')     { sheet.getRange(r,statusCol).setBackground('#EAF3EA').setFontColor('#2D6A2D'); }
    else if (status === 'Overdue')       { sheet.getRange(r,statusCol).setBackground('#FDF3E3').setFontColor('#7A4F0A').setFontWeight('bold'); }
    else if (status === 'Missed')        { sheet.getRange(r,statusCol).setBackground('#FDEAEA').setFontColor('#8B2020').setFontWeight('bold'); }
    else if (status === 'Not scheduled') { sheet.getRange(r,statusCol).setBackground('#F1F1F1').setFontColor('#888888'); }
    // Days until — colour based on urgency
    var days = row[8];
    if (typeof days === 'number') {
      if      (days < 0)  { sheet.getRange(r,9).setFontColor('#8B2020').setFontWeight('bold'); }
      else if (days <= 3) { sheet.getRange(r,9).setFontColor('#7A4F0A').setFontWeight('bold'); }
      else                { sheet.getRange(r,9).setFontColor('#2D6A2D'); }
    }
  });

  Logger.log('VISIT_SCHEDULE written: '+dataRows.length+' rows');
  return dataRows;
}

// ── Push visit tasks to TASK_ASSIGNMENTS (3 days before) ──────
function pushVisitTasks(cadence, history, openTasks, schedRows) {
  var asSheet  = getOrMakeTab(ASSIGN_TAB, writeAssignHeaders);
  var projSheet= db().getSheetByName(PROJECTS_TAB);
  var teamSheet= db().getSheetByName(TEAM_TAB);
  var today    = todayStr();
  // Assign through end of current week (Saturday) so full week visible on Monday
  var todayDate = new Date();
  var dow       = todayDate.getDay(); // 0=Sun,1=Mon...6=Sat
  var daysToSat = dow === 0 ? 6 : 6 - dow; // days until Saturday
  var threshold = addDaysToStr(today, daysToSat);
  var pushed   = 0;

  // Build project → discipline/multiplier map
  var projMap = {};
  if (projSheet) {
    var pRows = projSheet.getDataRange().getValues();
    for (var pi=1; pi<pRows.length; pi++) {
      var pn = String(pRows[pi][1]||'').trim();
      if (pn) projMap[pn] = {disc:String(pRows[pi][3]||''), mult:parseFloat(pRows[pi][4])||1.0};
    }
  }

  // Build name → email map for notifications
  var emailMap = {};
  if (teamSheet) {
    var tRows = teamSheet.getDataRange().getValues();
    for (var ti=1; ti<tRows.length; ti++) {
      var tn = String(tRows[ti][0]||'').trim();
      var te = String(tRows[ti][4]||'').trim();
      if (tn && te) emailMap[tn] = te;
    }
  }

  cadence.forEach(function(entry) {
    var hist     = (history[entry.project]||{})[entry.visitType]||[];
    var lastDate = hist.length > 0 ? hist[0].date : '';
    var nextInfo = calcNextVisitDate(entry, lastDate, openTasks);
    var nextDate = nextInfo.date;

    if (!nextDate) return;
    if (nextDate > threshold) return;

    var isOverdue = nextDate < today;
    var pd        = projMap[entry.project] || {disc:'',mult:1.0};

    // Split comma-separated assignees and create one task per person
    var rawAssignees = String(entry.assignee||'');
    var parts        = rawAssignees.split(',');
    var ai;
    for (ai = 0; ai < parts.length; ai++) {
      var assignee = parts[ai].trim();
      if (!assignee) continue;

      // Per-assignee dedup key
      var key = entry.project + '||' + entry.visitType + '||' + assignee;
      if (openTasks[key]) continue;

      var newId   = 'T-'+Utilities.getUuid().substring(0,8).toUpperCase();
      var isVisit = VISIT_TYPES_ARCH.indexOf(entry.visitType) > -1 ||
                    (entry.visitType||'').toLowerCase().indexOf('site visit') > -1;
      asSheet.appendRow([
        newId,                    // A TaskID
        '',                       // B ProjectID
        entry.project,            // C ProjectName
        assignee,                 // D AssignedTo
        entry.visitType,          // E Stage/TaskType
        pd.mult,                  // F Disc. Multiplier (number)
        0,                        // G Stage Base Pts (0 — visit pts based on hours)
        1,                        // H Units (always 1 visit unit)
        0,                        // I Weighted Pts (0 placeholder — updated when hours entered on DPR)
        today,                    // J AssignedDate
        nextDate,                 // K Deadline
        '',                       // L Area
        '',                       // M Drawing
        'Not Started',            // N SelfStatus
        '',                       // O SelfStatusDate
        '',                       // P ActualCompletionDate
        'Pending',                // Q LeadApproved
        '',                       // R ApprovedBy
        '',                       // S ApprovalDate
        '',                       // T RevisionTag
        (isOverdue ? 'OVERDUE — was due ' + nextDate + '. ' : '') + (entry.notes||''), // U Notes
        'Auto-scheduled',         // V AssignedBy
        isOverdue ? 'High' : 'Medium', // W Priority
      ]);

      pushed++;
      Logger.log('Visit task: '+entry.project+' / '+entry.visitType+' / '+assignee+' due '+nextDate);

      if (emailMap[assignee]) {
        sendVisitNotification(entry, nextDate, hist, isOverdue, emailMap[assignee]);
      }
    } // end for assignees
  }); // end cadence.forEach

  Logger.log('Pushed '+pushed+' visit tasks');
  return pushed;
}

// ── Flag missed visits in VISIT_SCHEDULE ─────────────────────
function flagMissedVisits(schedSheet, cadence, history, openTasks) {
  var asSheet = db().getSheetByName(ASSIGN_TAB);
  var today   = todayStr();
  var missed  = 0;

  cadence.forEach(function(entry) {
    var key      = entry.project + '||' + entry.visitType;
    var freqDays = VISIT_FREQ_DAYS[entry.frequency] || 14;
    var hist     = (history[entry.project]||{})[entry.visitType]||[];
    var lastDate = hist.length > 0 ? hist[0].date : '';
    var nextInfo = calcNextVisitDate(entry, lastDate, openTasks);
    var nextDate = nextInfo.date;
    if (!nextDate) return;

    // Missed = next date passed by more than 1 full frequency cycle
    var daysOverdue = daysDiff(nextDate, today);
    if (!daysOverdue || daysOverdue <= freqDays) return;

    missed++;
    Logger.log('MISSED: '+entry.project+' / '+entry.visitType+
               ' was due '+nextDate+', '+daysOverdue+'d overdue');

    // Create overdue task if no open task exists
    if (!openTasks[key] && asSheet) {
      var pd    = {};
      var proj2 = db().getSheetByName(PROJECTS_TAB);
      if (proj2) {
        var pRows = proj2.getDataRange().getValues();
        for (var pi=1; pi<pRows.length; pi++) {
          if (String(pRows[pi][1]||'').trim()===entry.project) {
            pd={disc:String(pRows[pi][3]||''),mult:parseFloat(pRows[pi][4])||1.0};
            break;
          }
        }
      }
      var newId='T-'+Utilities.getUuid().substring(0,8).toUpperCase();
      asSheet.appendRow([
        newId,                    // A TaskID
        '',                       // B ProjectID
        entry.project,            // C ProjectName
        entry.assignee,           // D AssignedTo
        entry.visitType,          // E Stage/TaskType
        pd.mult||1,               // F Disc. Multiplier (number)
        0,                        // G Stage Base Pts (0 — visit pts based on hours)
        1,                        // H Units (always 1 visit unit)
        0,                        // I Weighted Pts (0 placeholder)
        today,                    // J AssignedDate
        today,                    // K Deadline = today (overdue)
        '',                       // L Area
        '',                       // M Drawing
        'Not Started',            // N SelfStatus
        '',                       // O SelfStatusDate
        '',                       // P ActualCompletionDate
        'Pending',                // Q LeadApproved
        '',                       // R ApprovedBy
        '',                       // S ApprovalDate
        '',                       // T RevisionTag
        'MISSED VISIT — was due '+nextDate+'. '+(entry.notes||''), // U Notes
        'Auto-scheduled',         // V AssignedBy
        'High',                   // W Priority
      ]);
    }
  });

  Logger.log('Missed visits flagged: '+missed);
}

// ── Build PROJECT_HEALTH tab ──────────────────────────────────
function buildProjectHealth(cadence, history) {
  var sheet     = getOrMakeTab(PHEALTH_TAB, setupHealthTab);
  var projSheet = db().getSheetByName(PROJECTS_TAB);
  var today     = todayStr();
  var monthStart= cellDate(new Date(
    new Date().getFullYear(), new Date().getMonth(), 1));

  // Get all projects
  var projects = [];
  if (projSheet) {
    var pRows = projSheet.getDataRange().getValues();
    for (var pi=1; pi<pRows.length; pi++) {
      var pn = String(pRows[pi][1]||'').trim();
      var pl = String(pRows[pi][5]||'').trim(); // col F = Lead
      if (pn) projects.push({name:pn, lead:pl});
    }
  }

  // Build cadence lookup: project → [entries]
  var cadByProj = {};
  cadence.forEach(function(e) {
    if (!cadByProj[e.project]) cadByProj[e.project] = [];
    cadByProj[e.project].push(e);
  });

  // Build next visit map from VISIT_SCHEDULE
  var schedSheet  = db().getSheetByName(VSCHED_TAB);
  var nextVisitMap= {}; // project+type → next date
  if (schedSheet && schedSheet.getLastRow() >= 3) {
    var sRows=schedSheet.getRange(3,1,schedSheet.getLastRow()-2,6).getValues();
    sRows.forEach(function(r) {
      var p=String(r[0]||'').trim(), vt=String(r[1]||'').trim();
      var nd=cellDate(r[5]);
      if (p && vt && nd) nextVisitMap[p+'||'+vt]=nd;
    });
  }

  var dataRows = [];

  projects.forEach(function(proj) {
    var pHist     = history[proj.name] || {};
    var pCadence  = cadByProj[proj.name] || [];
    var activeTypes= pCadence.map(function(e){ return e.visitType; }).join(', ');

    // Last visits by type
    var archVisits= pHist[VT_ARCH]    || [];
    var idVisits  = pHist[VT_ID]      || [];
    var mtVisits  = pHist[VT_MEETING] || [];
    var spVisits  = pHist[VT_SPECIAL] || [];
    var allMeetings= mtVisits.concat(spVisits).sort(function(a,b){
      return b.date.localeCompare(a.date);
    });

    var lastArch = archVisits.length>0 ? archVisits[0].date : '';
    var lastID   = idVisits.length>0   ? idVisits[0].date   : '';
    var lastMeet = allMeetings.length>0 ? allMeetings[0].date: '';

    var daySinceArch = lastArch ? daysDiff(lastArch,today) : null;
    var daySinceID   = lastID   ? daysDiff(lastID,today)   : null;
    var daySinceMeet = lastMeet ? daysDiff(lastMeet,today) : null;

    // This month counts
    var visitsMonth  = archVisits.filter(function(v){ return v.date>=monthStart; }).length
                     + idVisits.filter(function(v){ return v.date>=monthStart; }).length;
    var meetingsMonth= allMeetings.filter(function(v){ return v.date>=monthStart; }).length;

    // Next planned
    var nextVisit  = nextVisitMap[proj.name+'||'+VT_ARCH]
                  || nextVisitMap[proj.name+'||'+VT_ID] || '';
    var nextMeeting= nextVisitMap[proj.name+'||'+VT_MEETING]
                  || nextVisitMap[proj.name+'||'+VT_SPECIAL] || '';

    // Missed visits in last 30 days
    var missed30 = 0;
    pCadence.forEach(function(e) {
      var freqDays = VISIT_FREQ_DAYS[e.frequency]||14;
      var hist2 = pHist[e.visitType]||[];
      // Count expected visits in last 30 days vs actual
      var expected = Math.floor(30/freqDays);
      var actual   = hist2.filter(function(v){
        return daysDiff(v.date,today)<=30 && daysDiff(v.date,today)>=0;
      }).length;
      missed30 += Math.max(0, expected-actual);
    });

    // Health per type — find cadence entry to get frequency
    function getFreq(vType) {
      var e = pCadence.find(function(c){ return c.visitType===vType; });
      return e ? e.frequency : null;
    }

    function calcHealth(daysSince, freq, threshMap) {
      if (daysSince === null) return 'Grey';
      var th = threshMap || HEALTH_THRESH[freq] || HEALTH_THRESH['Fortnightly'];
      if (daysSince <= th.green) return 'Green';
      if (daysSince <= th.amber) return 'Amber';
      return 'Red';
    }

    var archFreq   = getFreq(VT_ARCH);
    var idFreq     = getFreq(VT_ID);
    var archHealth = calcHealth(daySinceArch, archFreq, HEALTH_THRESH[archFreq]);
    var idHealth   = calcHealth(daySinceID,   idFreq,   HEALTH_THRESH[idFreq]);
    var meetHealth = calcHealth(daySinceMeet, null,     MEETING_THRESH);

    // Overall = worst
    var healthOrder= {'Red':3,'Amber':2,'Green':1,'Grey':0};
    var allHealths = [archHealth,idHealth,meetHealth];
    var overall    = allHealths.reduce(function(worst,h){
      return (healthOrder[h]||0) > (healthOrder[worst]||0) ? h : worst;
    },'Grey');

    dataRows.push([
      proj.name, proj.lead, activeTypes||'—',
      lastArch,  daySinceArch!==null ? daySinceArch : '—',
      lastID,    daySinceID!==null   ? daySinceID   : '—',
      lastMeet,  daySinceMeet!==null ? daySinceMeet : '—',
      visitsMonth, meetingsMonth,
      nextVisit, nextMeeting,
      missed30,
      archHealth, idHealth, meetHealth, overall,
    ]);
  });

  // Write data
  var startRow=3;
  var lastRow=sheet.getLastRow();
  if (lastRow>=startRow)
    sheet.getRange(startRow,1,lastRow-startRow+1,18)
      .clearContent().setBackground('#FFFFFF')
      .setFontColor('#000000').setFontWeight('normal');

  if (dataRows.length===0) return;

  sheet.getRange(startRow,1,dataRows.length,18).setValues(dataRows);

  // Format date cols
  [4,6,8].forEach(function(col){
    sheet.getRange(startRow,col,dataRows.length,1).setNumberFormat('dd-mmm-yyyy');
  });

  // Style rows and health cells
  dataRows.forEach(function(row,i) {
    var r=startRow+i;
    altRow(sheet,r,18);
    // Health cols: O=15, P=16, Q=17, R=18
    healthColour(sheet,r,15,row[14]);
    healthColour(sheet,r,16,row[15]);
    healthColour(sheet,r,17,row[16]);
    healthColour(sheet,r,18,row[17]);
    // Missed visits col: red if >0
    if (parseInt(row[13])>0) {
      sheet.getRange(r,14).setBackground('#FDEAEA').setFontColor('#8B2020').setFontWeight('bold');
    }
  });

  // Timestamp
  sheet.getRange(1,16).setValue('Updated: '+nowStr())
    .setFontSize(9).setFontColor('#888888').setHorizontalAlignment('right');

  Logger.log('PROJECT_HEALTH written: '+dataRows.length+' projects');
}

// ── Send email notification to assignee ───────────────────────
function sendVisitNotification(entry, nextDate, hist, isOverdue, email) {
  var lastVisit  = hist.length>0 ? hist[0] : null;
  var lastDateStr= lastVisit ? lastVisit.date+' — '+lastVisit.member : 'No previous visit on record';
  var daysSinceLast = lastVisit ? daysDiff(lastVisit.date, todayStr()) : null;

  var subject = '[IDS] '+(isOverdue?'OVERDUE: ':'')+
    entry.visitType+' due — '+entry.project+' — '+nextDate;

  var body = '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto">';
  body += '<div style="background:#1F3A5F;padding:18px 24px;border-radius:8px 8px 0 0">';
  body += '<p style="color:#fff;font-size:15px;font-weight:bold;margin:0">Ideaform Design Studio</p>';
  body += '<p style="color:rgba(255,255,255,.55);font-size:11px;margin:4px 0 0">Visit Reminder</p>';
  body += '</div>';
  body += '<div style="background:#fff;padding:20px 24px;border:1px solid #E2DFD8;border-top:none;border-radius:0 0 8px 8px">';
  if (isOverdue) {
    body += '<div style="background:#FDEAEA;border:1px solid #f0a0a0;border-radius:6px;padding:10px 14px;margin-bottom:16px">';
    body += '<p style="color:#8B2020;font-weight:bold;margin:0">⚠ OVERDUE VISIT</p></div>';
  }
  body += '<table style="width:100%;border-collapse:collapse;font-size:13px">';
  var rows2 = [
    ['Project',      entry.project],
    ['Visit type',   entry.visitType],
    ['Due date',     nextDate+(entry.fixedTime?' at '+entry.fixedTime:'')],
    ['Assigned to',  entry.assignee],
    ['Frequency',    entry.frequency],
    ['Last visit',   lastDateStr+(daysSinceLast!==null?' ('+daysSinceLast+' days ago)':'')],
  ];
  if (entry.notes) rows2.push(['Notes', entry.notes]);
  rows2.forEach(function(rv) {
    body += '<tr><td style="padding:7px 0;color:#6B6860;width:110px">'+rv[0]+'</td>';
    body += '<td style="padding:7px 0;font-weight:500;color:#1A1917">'+rv[1]+'</td></tr>';
  });
  body += '</table>';
  body += '<p style="font-size:11px;color:#9E9B94;margin:20px 0 0">Mark this visit done in your DPR form when completed. Points are calculated based on visit duration.</p>';
  body += '</div></div>';

  try {
    GmailApp.sendEmail(email, subject, '', {htmlBody:body});
    Logger.log('Email sent to '+entry.assignee+' ('+email+')');
  } catch(e) {
    Logger.log('Email failed for '+entry.assignee+': '+e);
  }
}

// ═══════════════════════════════════════════════════════════════
// MASTER FUNCTION — runs daily at 7 AM
// ═══════════════════════════════════════════════════════════════
// Alias — existing time-based trigger was created with this name
function runVisitScheduler() { syncVisitSchedule(); }

function syncVisitSchedule() {
  Logger.log('=== syncVisitSchedule: '+new Date().toISOString()+' ===');

  // Reconcile stalled-project parks every run (park new, revive un-stalled).
  try { var rc = reconcileStalledParks(); Logger.log('reconcileStalledParks: '+JSON.stringify(rc)); }
  catch(e){ Logger.log('reconcileStalledParks error: '+e); }

  var planSheet = getOrMakeTab(PLANNER_TAB, setupPlannerTab);

  // 1. Load all data
  var cadence   = loadCadence(planSheet);
  var history   = loadVisitHistory();
  var openTasks = loadOpenVisitTasks();

  // 1b. Load manual next dates from VISIT_SCHEDULE col F (user-editable)
  var schedSheet2 = db().getSheetByName(VSCHED_TAB);
  var manualDates = {};
  if (schedSheet2 && schedSheet2.getLastRow() > 2) {
    var sRows = schedSheet2.getDataRange().getValues();
    for (var si=2; si<sRows.length; si++) { // data starts row 3 (index 2)
      var sPrj  = String(sRows[si][0]||'').trim(); // A Project
      var sType = String(sRows[si][1]||'').trim(); // B Visit Type
      var sNext = cellDate(sRows[si][5]);           // F Next Visit Date (manual)
      if (sPrj && sType && sNext) {
        manualDates[sPrj + '||' + sType] = sNext;
      }
    }
  }
  // Inject manual next dates into cadence entries
  cadence.forEach(function(e) {
    var mk = e.project + '||' + e.visitType;
    if (manualDates[mk]) e.manualNextDate = manualDates[mk];
  });

  Logger.log('Cadence entries: '+cadence.length);

  if (cadence.length === 0) {
    Logger.log('No active cadence rows in VISIT_PLANNER — add projects to get started');
    return;
  }

  // 2. Build VISIT_SCHEDULE tab
  buildVisitSchedule(cadence, history, openTasks);

  // 3. Push tasks for visits due within 3 days (also sends emails)
  pushVisitTasks(cadence, history, openTasks);

  // 4. Flag missed visits + create overdue tasks
  var schedSheet = db().getSheetByName(VSCHED_TAB);
  flagMissedVisits(schedSheet, cadence, history, openTasks);

  // 5. Build PROJECT_HEALTH tab
  buildProjectHealth(cadence, history);

  Logger.log('=== syncVisitSchedule complete ===');
}

// ════════════════════════════════════════════════════════════════
// TRIGGER TO ADD:
// syncVisitSchedule → Time-driven → Day timer → 7am–8am daily
//
// TO RUN MANUALLY:
// Select syncVisitSchedule in function dropdown → Run
// Check Execution log — tabs created/updated within 30 seconds
//
// HOW TO SET UP VISIT_PLANNER:
// 1. Run syncVisitSchedule once → VISIT_PLANNER tab created
// 2. Add one row per project per visit type starting row 4:
//    Col A: Project name (must match PROJECTS tab exactly)
//    Col B: Visit Type (Site Visit (Architecture) | Site Visit (Interiors) | Client Meeting (Regular))
//    Col C: Assignee (must match TEAM tab exactly)
//    Col D: Frequency (Weekly | Fortnightly | Monthly)
//    Col E: Fixed Day (Monday-Saturday — only needed for Weekly)
//    Col F: Fixed Time (e.g. 11:00 AM — for client meetings)
//    Col G: Active (Yes / No)
//    Col H: Notes
// 3. Run syncVisitSchedule again → VISIT_SCHEDULE and PROJECT_HEALTH populated
// ════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════
// NOTIFICATIONS SYSTEM
// Tab: NOTIFICATIONS
// Written when lead rejects or corrects a task
// Read by DPR form on member select — shown in Section 0
//
// Cols: NotifID(A) Member(B) Project(C) TaskType(D) TaskTypeCorrected(E)
//   PtsOriginal(F) PtsCorrected(G) RejectionNote(H) LeadName(I)
//   Date(J) Seen(K) Source(L) TaskRow(M)
// ════════════════════════════════════════════════════════════════

var NOTIF_TAB = 'NOTIFICATIONS';

function setupNotifTab(sheet) {
  var h = ['Notif ID','Member','Project','Task Type','Task Type (Corrected)',
           'Pts (Original)','Pts (Corrected)','Rejection Note','Lead Name',
           'Date','Seen','Source','Task Row'];
  var r = sheet.getRange(1,1,1,h.length);
  r.setValues([h]);
  r.setBackground('#1F3A5F').setFontColor('#FFFFFF')
   .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  [90,140,180,200,200,90,90,240,120,100,70,120,80]
    .forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
}

function writeNotification(approval, member, taskRow, source) {
  var sheet = getOrCreate(NOTIF_TAB, setupNotifTab);
  var nid   = 'N-'+Utilities.getUuid().substring(0,8).toUpperCase();
  sheet.appendRow([
    nid,
    member,
    approval.project     || '',
    approval.taskType    || '',
    approval.correctedTaskType || '',
    approval.originalPts || '',
    approval.correctedPts|| '',
    approval.note        || '',
    approval.leadName    || '',
    dateStr(),
    'No',   // Seen
    source  || 'TASK_ASSIGNMENTS',
    taskRow || '',
  ]);
}

// GET notifications for a member — called from DPR form
function getNotificationsForMember(member) {
  member = resolveName(member);
  var sheet = db().getSheetByName(NOTIF_TAB);
  if (!sheet || sheet.getLastRow() < 2) return {notifications:[]};
  var rows  = sheet.getDataRange().getValues();
  var notifs= [];
  for (var i=1; i<rows.length; i++) {
    var rMember = String(rows[i][1]||'').trim();
    var seen    = String(rows[i][10]||'').trim().toLowerCase();
    if (rMember === member && seen !== 'yes') {
      notifs.push({
        row              : i+1,
        notifId          : String(rows[i][0]||''),
        member           : rMember,
        project          : String(rows[i][2]||''),
        taskType         : String(rows[i][3]||''),
        taskTypeCorrected: String(rows[i][4]||''),
        ptsOriginal      : rows[i][5]||'',
        ptsCorrected     : rows[i][6]||'',
        note             : String(rows[i][7]||''),
        leadName         : String(rows[i][8]||''),
        date             : String(rows[i][9]||'').substring(0,10),
        seen             : seen,
        source           : String(rows[i][11]||''),
      });
    }
  }
  return {notifications: notifs};
}

// Mark notification as seen — called from DPR form
function markNotificationSeen(data) {
  var sheet = db().getSheetByName(NOTIF_TAB);
  if (!sheet) return {status:'error'};
  var rows  = data.rows || [];
  rows.forEach(function(row) {
    var r = parseInt(row);
    if (r >= 2) sheet.getRange(r, 11).setValue('Yes'); // K Seen
  });
  return {status:'ok', marked: rows.length};
}


// ── Section 3 Done task → TASK_ASSIGNMENTS (replaces TASK_LOG) ───────────
function createDoneTask(data) {
  var sheet   = getOrCreate(ASSIGN_TAB, writeAssignHeaders);
  var s       = db();
  var today   = dateStr();
  var now     = new Date().toTimeString().substring(0,5); // HH:MM

  // Get project discipline + multiplier
  var projSheet = s.getSheetByName(PROJECTS_TAB);
  var disc = '', mult = 1.0;
  if (projSheet) {
    var pRows = projSheet.getDataRange().getValues();
    for (var i=1; i<pRows.length; i++) {
      if (String(pRows[i][1]||'').trim() === data.project ||
          String(pRows[i][0]||'').trim() === data.project) {
        disc = String(pRows[i][3]||'').trim();
        mult = parseFloat(pRows[i][4]) || 1.0;
        break;
      }
    }
  }

  // Calculate points — visit tasks use hours×rate, others use basePts×mult×units
  var basePts = parseFloat(data.basePts) || 0;
  var units   = parseFloat(data.units)   || 1;
  var weightedPts;
  if (isVisitTask(data.taskType) && data.visitHours) {
    weightedPts = calcVisitPts(data.taskType, data.visitHours);
    basePts     = parseFloat(data.visitHours) || basePts;
    units       = 1; // visits always 1 unit
  } else {
    weightedPts = Math.round(basePts * mult * units * 10) / 10;
  }

  var newId = 'T-'+Utilities.getUuid().substring(0,8).toUpperCase();
  var member = data.member || '';

  sheet.appendRow([
    newId,
    '',
    data.project      || '',   // C ProjectName
    member,                    // D AssignedTo (self)
    data.taskType     || '',   // E Stage
    mult,                      // F Disc. Multiplier (number)
    basePts,                   // G Stage Base Pts
    units,                     // H Units
    weightedPts,               // I Weighted Pts = F × G × H
    data.date         || today, // J AssignedDate
    today,                     // K Deadline = today (Done same day)
    data.area         || '',   // L Area
    data.drawing      || '',   // M Drawing
    'Done',                    // N SelfStatus
    today,                     // O SelfStatusDate
    today,                     // P ActualCompletionDate
    'Pending',                 // Q LeadApproved
    '',                        // R ApprovedBy
    '',                        // S ApprovalDate
    '',                        // T RevisionTag
    'Unplanned task — self logged via DPR', // U Notes
    member,                    // V AssignedBy (self)
    'Medium',                  // W Priority
  ]);

  Logger.log('Done task created: '+member+' / '+data.taskType+' = '+weightedPts+'pts → '+newId);
  return {status:'ok', taskId:newId};
}

// Handle self-created unplanned ongoing task — write to TASK_ASSIGNMENTS
function createSelfAssignedTask(data) {
  var sheet    = getOrCreate(ASSIGN_TAB, writeAssignHeaders);
  var tomorrow = addDaysToStr(dateStr(), 1);
  var s        = db();

  // Get project discipline + multiplier
  var projSheet = s.getSheetByName(PROJECTS_TAB);
  var disc = '', mult = 1.0;
  if (projSheet) {
    var pRows = projSheet.getDataRange().getValues();
    for (var i=1; i<pRows.length; i++) {
      if (String(pRows[i][1]||'').trim() === data.project) {
        disc = String(pRows[i][3]||'').trim();
        mult = parseFloat(pRows[i][4]) || 1.0;
        break;
      }
    }
  }

  var basePts    = parseFloat(data.basePts)  || 0;
  var units      = parseFloat(data.units)    || 1;
  var weightedPts= parseFloat((basePts * mult * units).toFixed(2));
  var newId = 'T-'+Utilities.getUuid().substring(0,8).toUpperCase();

  sheet.appendRow([
    newId,
    '',
    data.project      || '',   // C
    data.member       || '',   // D AssignedTo (self)
    data.taskType     || '',   // E Stage
    mult,                      // F Disc. Multiplier (number)
    basePts,                   // G Stage Base Pts
    units,                     // H Units
    weightedPts,               // I Weighted Pts = F × G × H
    data.date         || dateStr(), // J AssignedDate
    tomorrow,                  // K Deadline = tomorrow
    data.area         || '',   // L Area
    data.drawing      || '',   // M Drawing
    'In Progress',             // N SelfStatus
    dateStr(),                 // O SelfStatusDate
    '',                        // P ActualCompletionDate
    'Pending',                 // Q LeadApproved
    '',                        // R ApprovedBy
    '',                        // S ApprovalDate
    '',                        // T RevisionTag
    'Unplanned — self assigned via DPR', // U Notes
    data.member       || '',   // V AssignedBy (self)
    'Medium',                  // W Priority
  ]);
  parkRowIfStalled(sheet, data.project || '');  // stalled project → auto-park

  Logger.log('Self-assigned task created: '+data.member+' / '+data.taskType+' → '+newId);
  return {status:'ok', taskId:newId};
}

// Handle Siddharth task creation from DPER form — Deepak assigns pending discussions
// ════════════════════════════════════════════════════════════════
// DEBUG — run testCreateSiddharthTask() directly from the editor
// to verify createSiddharthTask works independently of the form
// ════════════════════════════════════════════════════════════════
function testCreateSiddharthTask() {
  var result;
  try {
    result = createSiddharthTask({
      project    : 'TEST-DEBUG',
      description: 'Test pending discussion created by testCreateSiddharthTask() at ' + new Date().toLocaleString(),
      assignedBy : 'Deepak Soni',
      date       : dateStr(),
    });
    Logger.log('testCreateSiddharthTask SUCCESS: ' + JSON.stringify(result));
    // Write outcome to spreadsheet so you can see it without logs
    var s = db().getSheetByName(ASSIGN_TAB);
    if (s) {
      var last = s.getLastRow();
      s.getRange(last, 1, 1, 23).setBackground('#C8E6C9'); // green highlight = test row
    }
  } catch(e) {
    Logger.log('testCreateSiddharthTask FAILED: ' + String(e));
    result = {status:'error', message: String(e)};
    // Write a visible error marker to spreadsheet
    try {
      var errSheet = db().getSheetByName('DEBUG_LOG') || db().insertSheet('DEBUG_LOG');
      errSheet.appendRow([new Date(), 'createSiddharthTask ERROR', String(e)]);
    } catch(e2) {}
  }
  return result;
}

function createSiddharthTask(data) {
  Logger.log('createSiddharthTask called: project=' + data.project + ' desc=' + data.description);
  var sheet    = getOrCreate(ASSIGN_TAB, writeAssignHeaders);
  var today    = dateStr();
  var deadline = addDaysToStr(today, 1); // Next day deadline
  var s        = db();

  // Get project discipline + multiplier from PROJECTS sheet
  var mult = 1.0;
  var projSheet = s.getSheetByName(PROJECTS_TAB);
  if (projSheet) {
    var pRows = projSheet.getDataRange().getValues();
    for (var i=1; i<pRows.length; i++) {
      if (String(pRows[i][1]||'').trim() === data.project ||
          String(pRows[i][0]||'').trim() === data.project) {
        mult = parseFloat(pRows[i][4]) || 1.0;
        break;
      }
    }
  }

  var newId = 'T-'+Utilities.getUuid().substring(0,8).toUpperCase();
  var basePts = 0; // Pending discussion items don't have points yet
  var weightedPts = 0;

  sheet.appendRow([
    newId,                       // A TaskID
    '',                          // B ProjectID
    data.project      || '',     // C ProjectName
    'Siddharth Inani',           // D AssignedTo
    'Pending Discussion',        // E Stage/TaskType
    mult,                        // F Disc. Multiplier
    basePts,                     // G Stage Base Pts
    1,                           // H Units
    weightedPts,                 // I Weighted Pts
    data.date || today,          // J AssignedDate
    deadline,                    // K Deadline
    '',                          // L Area
    '',                          // M DrawingName
    'Not Started',               // N SelfStatus
    '',                          // O SelfStatusDate
    '',                          // P ActualCompletionDate
    'Pending',                   // Q LeadApproved
    '',                          // R ApprovedBy
    '',                          // S ApprovalDate
    '',                          // T RevisionTag
    data.description || '',      // U Notes (the pending discussion items)
    data.assignedBy || 'Deepak Soni', // V AssignedBy
    'High',                      // W Priority
  ]);

  Logger.log('Siddharth task created for '+data.project+': '+newId);
  return {status:'ok', taskId:newId};
}

// ════════════════════════════════════════════════════════════════
// getWeeklyStatsForReport — called by weekly report HTML
// Returns per-member stats for a given week (Mon-Sat)
// action=getWeeklyStats&weekStart=2026-05-11
// ════════════════════════════════════════════════════════════════
function getWeeklyStats(weekStart) {
  var s   = db();
  var mon = weekStart || dateStr(mondayOf(new Date()));
  var sat = addDaysToStr(mon, 5);
  // Completed/approved work counts through TODAY when the week has already ended,
  // so tasks done in the week but approved a day or two later (after Saturday)
  // are still included. Punctuality/hours stay on the Mon–Sat window.
  var today = dateStr();
  var doneUpper = (today > sat) ? today : sat;

  var asSheet  = s.getSheetByName(ASSIGN_TAB);
  var sumSheet = s.getSheetByName(SUMMARY_TAB);
  var tSheet   = s.getSheetByName(TEAM_TAB);
  if (!tSheet) return {error:'TEAM tab not found'};

  var asRows  = asSheet  ? asSheet.getDataRange().getValues()  : [];
  var sumRows = sumSheet ? sumSheet.getDataRange().getValues() : [];
  var tRows   = tSheet.getDataRange().getValues();

  // Auto-detect column structure
  var asHeaders   = asRows[0] ? asRows[0].map(function(h){ return String(h||'').trim(); }) : [];
  var is23        = asHeaders.length >= 23 || asHeaders.indexOf('Actual Completion Date') > -1;
  var COL_LEADAPPR= is23 ? 16 : 15;
  var COL_STATUSDT= 14;
  var COL_ACTUALDT= is23 ? 15 : -1; // P ActualCompletionDate (23-col only)
  var COL_DEADLINE= 10;
  var COL_REVTAG  = is23 ? 19 : 18;
  var COL_ASSIGNDT= 9;

  var results = [];

  for (var ti = 1; ti < tRows.length; ti++) {
    var name   = String(tRows[ti][0] || '').trim();
    var role   = String(tRows[ti][1] || '').trim();
    var wkTgt  = parseFloat(tRows[ti][2]) || 50;
    var active = String(tRows[ti][5] || '').trim().toLowerCase();
    if (!name || active === 'no') continue;

    // Tasks assigned this week (AssignedDate in Mon-Sat)
    var tasksAssigned = 0, assignedPts = 0;
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var aDate = cellDate(r[COL_ASSIGNDT]);
      if (String(r[3]||'').trim() === name && aDate >= mon && aDate <= sat) {
        tasksAssigned++;
        assignedPts += parseFloat(r[8]) || 0;
      }
    });

    // Tasks completed this week — SOURCE A: TASK_ASSIGNMENTS (assigned tasks)
    var completedOnTime = 0, completedLate = 0;
    var completedTasks = [];

    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var doneDate   = (COL_ACTUALDT > -1 ? cellDate(r[COL_ACTUALDT]) : '') || cellDate(r[COL_STATUSDT]);
      var isApproved = String(r[COL_LEADAPPR]||'').trim() === 'Yes';
      if (String(r[3]||'').trim() === name &&
          String(r[13]||'').trim() === 'Done' &&
          isApproved &&
          doneDate >= mon && doneDate <= doneUpper) {
        var deadline = cellDate(r[COL_DEADLINE]);
        var onTime   = deadline && doneDate <= deadline;
        if (onTime) completedOnTime++; else completedLate++;
        completedTasks.push({
          project  : String(r[2]  ||'').trim(),
          taskType : String(r[4]  ||'').trim(),
          area     : String(r[11] ||'').trim(),
          drawing  : String(r[12] ||'').trim(),
          doneDate : doneDate,
          deadline : deadline || '',
          pts      : parseFloat(r[8]) || 0,
          onTime   : onTime,
          source   : 'assigned',
        });
      }
    });

    // SOURCE B: TASK_LOG (unplanned DPR work — also approved)
    // TASK_LOG: Date=B(1) Member=C(2) Project=D(3) Disc=E(4) TaskType=F(5)
    //   Area=G(6) Drawing=H(7) Pts=J(9) LeadApproved=K(10)
    // TASK_LOG source removed

    // Sort by project then done date
    completedTasks.sort(function(a,b){
      return a.project.localeCompare(b.project) || a.doneDate.localeCompare(b.doneDate);
    });

    // Approved pts this week (ActualCompletion → SelfStatusDate filter)
    var approvedPts = 0;
    // TASK_LOG source removed
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var doneDate = (COL_ACTUALDT > -1 ? cellDate(r[COL_ACTUALDT]) : '') || cellDate(r[COL_STATUSDT]);
      if (String(r[3]||'').trim() === name &&
          String(r[COL_LEADAPPR]||'').trim() === 'Yes' &&
          doneDate >= mon && doneDate <= doneUpper)
        approvedPts += parseFloat(r[8]) || 0;
    });

    // DPR days filed this week
    // DAILY_SUMMARY cols: Date=A(0) Time=B(1) Member=C(2) Email=D(3) ArrivedOnTime=E(4)
    var dprDays = 0;
    sumRows.forEach(function(r, i) {
      if (i === 0) return;
      var rDate = cellDate(r[0]); // col A = Date
      var rName = String(r[2]||'').trim(); // col C = Member name
      if (rName === name && rDate >= mon && rDate <= sat)
        dprDays++;
    });

    // Reliability — delays this week. Late (deadline in week, passed, not done
    // on time) = −1; "Work Not Done" (reassigned by a lead) = −2.
    var lateCount = 0, wndCount = 0;
    var nowD = dateStr();
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      if (String(r[3]||'').trim() !== name) return;
      var dl = cellDate(r[COL_DEADLINE]);
      if (!dl || dl < mon || dl > sat) return;       // deadline must fall in this week
      var selfStat = String(r[13]||'').trim();         // N SelfStatus
      if (selfStat === 'Work Not Done') { wndCount++; return; }
      if (dl >= nowD) return;                          // not due yet — not late
      var doneDate = (COL_ACTUALDT > -1 ? cellDate(r[COL_ACTUALDT]) : '') || cellDate(r[COL_STATUSDT]);
      var onTime   = String(r[COL_LEADAPPR]||'').trim() === 'Yes' && doneDate && doneDate <= dl;
      if (!onTime) lateCount++;
    });
    var rejBlk = rejectedBlocksThisWeek(name);   // −1 each (gaming a block)
    var reliabilityPenalty = lateCount * 1 + wndCount * 2 + rejBlk * 1;

    results.push({
      name           : name,
      role           : role,
      weeklyTarget   : wkTgt,
      tasksAssigned  : tasksAssigned,
      assignedPts    : Math.round(assignedPts * 10) / 10,
      completedOnTime: completedOnTime,
      completedLate  : completedLate,
      approvedPts    : Math.round(approvedPts * 10) / 10,
      dprDays        : dprDays,
      lateCount      : lateCount,
      workNotDone    : wndCount,
      rejectedBlocks : rejBlk,
      blockStats     : blockStatsForWeek(name, mon, sat),
      reliabilityPenalty: reliabilityPenalty,
      completedTasks : completedTasks,
    });
  }

  return {stats: results, weekStart: mon, weekEnd: sat};
}

// ════════════════════════════════════════════════════════════════
// getProjectStats — for projects dashboard
// Returns per-project task stats, team members, visit schedule
// ════════════════════════════════════════════════════════════════
function getProjectStats() {
  var s         = db();
  var asSheet   = s.getSheetByName(ASSIGN_TAB);
  var projSheet = s.getSheetByName(PROJECTS_TAB);
  var vsSheet   = s.getSheetByName(VSCHED_TAB);
  var teamSheet = s.getSheetByName(TEAM_TAB);

  if (!asSheet || !projSheet) return {projects:[]};

  var asRows   = asSheet.getDataRange().getValues();
  var projRows = projSheet.getDataRange().getValues();
  var vsRows   = vsSheet   ? vsSheet.getDataRange().getValues()   : [];
  var teamRows = teamSheet ? teamSheet.getDataRange().getValues() : [];

  // Auto-detect columns
  var headers  = asRows[0] ? asRows[0].map(function(h){return String(h||'').trim();}) : [];
  var is23     = headers.length >= 23 || headers.indexOf('Actual Completion Date') > -1;
  var COL_LAPPR= is23 ? 16 : 15;
  var COL_SSTAT    = 13;
  var COL_SDATE    = 14;  // O SelfStatusDate
  var COL_ACTUALDT = is23 ? 15 : -1; // P ActualCompletionDate
  var COL_DEAD     = 10;

  var today = dateStr();

  // Build project map from PROJECTS tab
  // Cols: ProjectID(A) Name(B) Status(C) Discipline(D) Multiplier(E) Lead(F)
  var projMap = {};
  for (var pi = 1; pi < projRows.length; pi++) {
    var pid  = String(projRows[pi][0]||'').trim();
    var pname= String(projRows[pi][1]||'').trim();
    if (!pname) continue;
    projMap[pname] = {
      id        : pid,
      name      : pname,
      status    : String(projRows[pi][2]||'').trim(),
      discipline: String(projRows[pi][3]||'').trim(),
      lead      : String(projRows[pi][5]||'').trim(),
    };
  }

  // Build task stats per project
  var stats = {}; // key = project name
  for (var ai = 1; ai < asRows.length; ai++) {
    var r       = asRows[ai];
    var proj    = String(r[2]||'').trim();
    var assignee= String(r[3]||'').trim();
    var taskType= String(r[4]||'').trim();
    var area    = String(r[11]||'').trim();
    var drawing = String(r[12]||'').trim();
    var sStatus = String(r[COL_SSTAT]||'').trim();
    var lApproved=String(r[COL_LAPPR]||'').trim();
    var deadline= cellDate(r[COL_DEAD]);
    var sDate   = (COL_ACTUALDT > -1 ? cellDate(r[COL_ACTUALDT]) : '') || cellDate(r[COL_SDATE]);
    var pts     = parseFloat(r[8])||0;
    var priority= String(r[22]||'').trim();

    if (!proj) continue;

    if (!stats[proj]) {
      stats[proj] = {
        name           : proj,
        tasks          : [],
        totalTasks     : 0,
        completedTasks : 0,
        onTimeTasks    : 0,
        lateTasks      : 0,
        inProgressTasks: 0,
        notStartedTasks: 0,
        delayedTasks   : 0,
        approvalPending: 0,
        revisionTasks  : 0,
        totalPts       : 0,
        approvedPts    : 0,
        pendingPts     : 0,
        members        : {},
      };
    }

    var st = stats[proj];
    st.totalTasks++;
    st.totalPts += pts;

    // Determine display status
    var dispStatus;
    if (sStatus === 'Done' && lApproved === 'Yes') {
      dispStatus = 'Completed';
      st.completedTasks++;
      st.approvedPts += pts;
      if (deadline && sDate && sDate <= deadline) st.onTimeTasks++;
      else if (deadline && sDate) st.lateTasks++;
    } else if (sStatus === 'Done' && lApproved !== 'Yes' && lApproved !== 'No') {
      dispStatus = 'Approval Pending';
      st.approvalPending++;
      st.pendingPts += pts;
    } else if (sStatus === 'Done' && lApproved === 'No') {
      dispStatus = 'Revision Required';
      st.revisionTasks++;
    } else if (sStatus === 'In Progress') {
      dispStatus = 'In Progress';
      st.inProgressTasks++;
      st.pendingPts += pts;
      if (deadline && deadline < today) { dispStatus = 'Delayed'; st.delayedTasks++; }
    } else if (deadline && deadline < today && sStatus !== 'Done') {
      dispStatus = 'Delayed';
      st.delayedTasks++;
      st.pendingPts += pts;
    } else {
      dispStatus = 'Not Started';
      st.notStartedTasks++;
      st.pendingPts += pts;
    }

    // Add task to list
    st.tasks.push({
      taskId  : String(r[0]||''),
      taskType: taskType,
      area    : area,
      drawing : drawing,
      assignee: assignee,
      deadline: deadline,
      status  : dispStatus,
      pts     : pts,
      priority: priority,
    });

    // Track members
    assignee.split(',').map(function(a){return a.trim();}).forEach(function(a){
      if (a && a !== 'Unassigned') st.members[a] = (st.members[a]||0) + 1;
    });
  }

  // Build visit schedule per project
  var visitMap = {}; // key = project name
  for (var vi = 1; vi < vsRows.length; vi++) {
    var vr      = vsRows[vi];
    var vproj   = String(vr[0]||'').trim(); // A Project
    var vtype   = String(vr[1]||'').trim(); // B VisitType
    var vassign = String(vr[2]||'').trim(); // C Assignee
    var vlast   = cellDate(vr[3]);          // D LastVisitDate
    var vnext   = cellDate(vr[5]);          // F NextVisitDate
    var vstatus = String(vr[6]||'').trim(); // G Status
    if (!vproj) continue;
    if (!visitMap[vproj]) visitMap[vproj] = [];
    visitMap[vproj].push({
      type    : vtype,
      assignee: vassign,
      lastDate: vlast,
      nextDate: vnext,
      status  : vstatus,
    });
  }

  // Combine and return
  var result = [];
  Object.keys(stats).forEach(function(pname) {
    var st   = stats[pname];
    var info = projMap[pname] || {};
    result.push({
      name           : pname,
      projectId      : info.id      || '',
      status         : info.status  || 'Ongoing',
      discipline     : info.discipline || '',
      lead           : info.lead    || '',
      totalTasks     : st.totalTasks,
      completedTasks : st.completedTasks,
      onTimeTasks    : st.onTimeTasks,
      lateTasks      : st.lateTasks,
      inProgressTasks: st.inProgressTasks,
      notStartedTasks: st.notStartedTasks,
      delayedTasks   : st.delayedTasks,
      approvalPending: st.approvalPending,
      revisionTasks  : st.revisionTasks,
      totalPts       : Math.round(st.totalPts  * 10) / 10,
      approvedPts    : Math.round(st.approvedPts * 10) / 10,
      pendingPts     : Math.round(st.pendingPts * 10) / 10,
      completionPct  : st.totalTasks > 0 ? Math.round(st.completedTasks/st.totalTasks*100) : 0,
      onTimePct      : (st.completedTasks+st.lateTasks) > 0
                        ? Math.round(st.onTimeTasks/(st.completedTasks+st.lateTasks)*100) : 0,
      members        : Object.keys(st.members),
      tasks          : st.tasks,
      visits         : visitMap[pname] || [],
    });
  });

  // Sort by delayed first, then in progress, then name
  result.sort(function(a,b){
    if (a.delayedTasks > 0 && b.delayedTasks === 0) return -1;
    if (b.delayedTasks > 0 && a.delayedTasks === 0) return 1;
    return a.name.localeCompare(b.name);
  });

  return {projects: result};
}

// ════════════════════════════════════════════════════════════════
// getCalendarData — for projects dashboard calendar view
// Returns visit/meeting blocks for the current week
// ════════════════════════════════════════════════════════════════
function getCalendarData() {
  var s       = db();
  var asSheet = s.getSheetByName(ASSIGN_TAB);
  if (!asSheet) return {events:[]};

  var rows    = asSheet.getDataRange().getValues();
  var headers = rows[0] ? rows[0].map(function(h){return String(h||'').trim();}) : [];
  var is23    = headers.length >= 23 || headers.indexOf('Actual Completion Date') > -1;
  var COL_LAPPR = is23 ? 16 : 15;

  var today   = dateStr();
  var mon     = mondayOf(new Date());
  var monStr  = dateStr(mon);
  var sunStr  = addDaysToStr(monStr, 6);
  var events  = [];

  for (var i = 1; i < rows.length; i++) {
    var r        = rows[i];
    var taskType = String(r[4]||'').trim();
    if (!isVisitTask(taskType)) continue;

    var assignee = String(r[3]||'').trim();
    var project  = String(r[2]||'').trim();
    var deadline = cellDate(r[10]);
    var timeStr  = String(r[12]||'').trim(); // M = planned time HH:MM
    var sStatus  = String(r[13]||'').trim();
    var lApproved= String(r[COL_LAPPR]||'').trim();
    var pts      = parseFloat(r[8])||0;

    if (!deadline || deadline < monStr || deadline > sunStr) continue;
    if (sStatus === 'Done' && lApproved === 'Yes') continue; // already completed

    // Parse time — default 10:00 if missing
    var hour = 10, minute = 0;
    if (timeStr && timeStr.indexOf(':') > -1) {
      var parts = timeStr.split(':');
      hour   = parseInt(parts[0]) || 10;
      minute = parseInt(parts[1]) || 0;
    }

    // Duration: Site Visit = 2 hours, Client Meeting = 1 hour
    var duration = (VISIT_TYPES_ARCH.indexOf(taskType) > -1) ? 2 : 1;

    events.push({
      project : project,
      taskType: taskType,
      assignee: assignee,
      date    : deadline,
      hour    : hour,
      minute  : minute,
      duration: duration,
      status  : sStatus,
      pts     : pts,
    });
  }

  return {events: events};
}


// ── handleDPERSubmission — main DPER submission handler ───────
function handleDPERSubmission(data) {
  try {
    var subId     = writeSiteExecution(data);
    var issues    = [];
    try { issues = data.issues ? JSON.parse(data.issues) : []; } catch(e){}
    writeSiteIssues(subId, data.date||dateStr(), data.project||'',
                    issues, data.onTrack||'Yes');

    // Create Siddharth task if pending discussion text was entered at project level
    Logger.log('siddharthPending received: [' + data.siddharthPending + '] for project: ' + data.project);
    var sidPending = data.siddharthPending ? String(data.siddharthPending).trim() : '';
    var siddharthTaskId = null;
    var siddharthError  = null;
    if (sidPending) {
      try {
        var taskResult = createSiddharthTask({
          project    : data.project   || '',
          description: sidPending,
          assignedBy : data.lead      || 'Deepak Soni',
          date       : data.date      || dateStr(),
        });
        siddharthTaskId = taskResult.taskId || null;
        Logger.log('Siddharth task created: ' + JSON.stringify(taskResult));
      } catch(sidErr) {
        siddharthError = String(sidErr);
        Logger.log('ERROR in createSiddharthTask: ' + siddharthError);
      }
    }

    // Write WhatsApp message to SITE_WA_MESSAGES tab
    if (data.whatsappMsg) {
      var waSheet = getOrCreate(SITE_WA_TAB, writeSiteWAHeaders);
      var nonDesignIssues = issues.filter(function(iss) {
        return iss.issueType !== 'Design';
      });
      waSheet.appendRow([
        subId,
        data.date   || dateStr(),
        data.project|| '',
        data.lead   || '',
        data.onTrack|| 'Yes',
        nonDesignIssues.length,
        data.whatsappMsg,
      ]);
    }

    // EPIC K — a DPER site visit auto-publishes a client Site Visit Log (no second form).
    // Tasks are already created by writeSiteIssues above, so skipTasks=true here.
    var meetingLogId = null;
    try {
      if (String(data.siteVisit||'').toLowerCase() === 'yes') {
        var mlActions = { Client:[], IDS:[], Contractor:[], Other:[] };
        issues.forEach(function(iss){
          var item = { owner:String(iss.assignedTo||''), text:String(iss.description||''), deadline:cellDate(iss.targetDate)||'' };
          if (!item.text) return;
          if ((iss.issueType||'')==='Design') mlActions.IDS.push(item); else mlActions.Contractor.push(item);
        });
        if (data.clientConcerns) mlActions.Client.push({ owner:'', text:String(data.clientConcerns), deadline:'' });
        if (data.blocking)       mlActions.Other.push({ owner:'', text:String(data.blocking), deadline:'' });
        var mlRes = submitMeetingLog({
          type:'Site Visit', project:data.project||'', date:data.date||dateStr(),
          time:data.time||'', loggedBy:data.lead||'Deepak Soni',
          teamAttendees:(data.lead?[data.lead]:[]), clientAttendees:'',
          purpose:['Progress review','Quality check'],
          body:String(data.worksToday||''), actions:mlActions,
          duration:'', photoIds:[], videoLinks:[], skipTasks:true
        }, '');
        meetingLogId = mlRes && mlRes.logId;
      }
    } catch(mlErr){ Logger.log('DPER→meetingLog error: '+mlErr); }

    Logger.log('DPER submitted: ' + subId + ' / ' + (data.project||''));
    return {
      status           : 'ok',
      subId            : subId,
      meetingLogId     : meetingLogId,
      siddharthTaskId  : siddharthTaskId,
      siddharthError   : siddharthError,
    };
  } catch(err) {
    Logger.log('DPER error: ' + String(err));
    return {status:'error', message:String(err)};
  }
}

// ════════════════════════════════════════════════════════════════
// SITE EXECUTION — DPER Form functions
// ════════════════════════════════════════════════════════════════

function writeSiteExecutionHeaders(sheet) {
  var h = ['Submission ID','Date','Time','Project Name','Execution Lead',
    'Site Visit Done','Current Stage','Works Completed Today',
    '% Works Completed as per Plan','% Overall Project Completion',
    'Work Planned for Tomorrow','On Track Status','Delay/Risk Reason',
    'Idle Time Observed','Idle Time Reason','Materials Required (3-5 days)',
    'Material Delays Impacting Work','Client Updated Today',
    'Client Concerns/Inputs','Blocking Tomorrow',
    'Decisions Pending from Siddharth','Additional Remarks'];
  sheet.getRange(1,1,1,h.length).setValues([h])
    .setBackground('#1F3A5F').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  var widths=[160,100,80,200,160,100,220,280,100,100,280,120,220,100,180,220,220,100,220,220,220,220];
  widths.forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
}

function writeSiteWAHeaders(sheet) {
  var h = ['Submission ID','Date','Project Name','Execution Lead',
            'On Track Status','Issues Count','WhatsApp Message'];
  sheet.getRange(1,1,1,h.length).setValues([h])
    .setBackground('#1B5E20').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  var widths = [140,100,200,160,140,100,600];
  widths.forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  // Set col G to wrap text
  sheet.getRange(1,7,1000,1).setWrap(true);
}

function writeSiteIssuesHeaders(sheet) {
  var h = ['Issue ID','Submission ID','Date','Project Name','Issue #',
    'Issue Type','Description','Assigned To','Priority (1=Low 3=High)',
    'Target Date','Status','Task ID (TASK_ASSIGNMENTS)',
    'Resolution Notes','Resolved Date','Reported By'];
  sheet.getRange(1,1,1,h.length).setValues([h])
    .setBackground('#8B2020').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  var widths=[140,140,100,200,70,160,340,160,120,110,120,180,220,110,160];
  widths.forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
}

// ── Generate sequential ID ─────────────────────────────────
function nextId(sheet, prefix) {
  var last = sheet.getLastRow();
  if (last < 2) return prefix + '001';
  var vals = sheet.getRange(2,1,last-1,1).getValues();
  var nums = vals.map(function(r){
    var n = parseInt(String(r[0]).replace(prefix,''));
    return isNaN(n) ? 0 : n;
  });
  var max = nums.reduce(function(a,b){ return Math.max(a,b); }, 0);
  return prefix + String(max+1).padStart(3,'0');
}

// ── writeSiteExecution ─────────────────────────────────────
function writeSiteExecution(data) {
  var sheet = getOrCreate(SITE_EXEC_TAB, writeSiteExecutionHeaders);
  var subId = nextId(sheet, 'SE-');
  var today = dateStr();
  var now   = new Date().toTimeString().substring(0,5);

  prependRow(sheet, [
    subId,
    data.date         || today,
    data.time         || now,
    data.project      || '',
    data.lead         || '',
    data.siteVisit    || 'No',
    data.stage        || '',
    data.worksToday   || '',
    data.pctPlan      || 0,
    data.pctOverall   || 0,
    data.workTomorrow || '',
    data.onTrack      || 'Yes',
    data.delayReason  || '',
    data.idleTime     || 'No',
    data.idleReason   || '',
    data.materials    || '',
    data.materialDelays|| '',
    data.clientUpdated|| 'No',
    data.clientConcerns|| '',
    data.blocking          || '',
    data.siddharthPending  || '',   // col U: Decisions Pending from Siddharth
    data.remarks           || '',
  ]);
  return subId;
}

// ── writeSiteIssues ────────────────────────────────────────
function writeSiteIssues(subId, date, project, issues, onTrack, reportedBy) {
  if (!issues || issues.length === 0) return [];
  var sheet  = getOrCreate(SITE_ISSUES_TAB, writeSiteIssuesHeaders);
  var asSheet= db().getSheetByName(ASSIGN_TAB);
  var taskIds= [];
  var reporter = reportedBy || 'Deepak Soni';

  issues.forEach(function(iss, idx) {
    var issId  = nextId(sheet, 'ISS-');
    var taskId = '';
    var issProject = iss.project || project;  // each issue carries its own project (CRM); DPER uses shared param

    // Auto-create task in TASK_ASSIGNMENTS for all issue types
    if (asSheet && iss.assignedTo && iss.description) {
      taskId = createIssueTask(iss, issProject, date, onTrack);
    }

    prependRow(sheet, [
      issId,
      subId,
      date,
      issProject,
      idx + 1,
      iss.issueType    || '',
      iss.description  || '',
      iss.assignedTo   || '',
      iss.priority     || 2,
      iss.targetDate   || '',
      'Open',
      taskId,
      '',
      '',
      iss.reportedBy   || reporter,  // col O: Reported By
    ]);
    taskIds.push(taskId);
  });
  return taskIds;
}

// ── createIssueTask ────────────────────────────────────────
function getProjectLead(projectName) {
  // Look up project lead from PROJECTS tab col F (index 5)
  var projSheet = db().getSheetByName(PROJECTS_TAB);
  if (!projSheet) return '';
  var rows = projSheet.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    var pName = String(rows[i][1]||'').trim(); // col B = project name
    if (pName === projectName) return String(rows[i][5]||'').trim(); // col F = lead
  }
  return '';
}

// Generic role names that should be replaced with actual person
var GENERIC_ROLES = ['Design Lead','Execution Lead','Client','Contractor',
                     'Vendor','PMC','MEP','Issue resolved',''];

function resolveAssignee(assignedTo, projectName) {
  var name = String(assignedTo||'').trim();
  // If it's a generic role or blank — use project lead from col F
  if (GENERIC_ROLES.indexOf(name) > -1) {
    return getProjectLead(projectName) || name;
  }
  return name; // Manual selection from Deepak — use as-is
}

function createIssueTask(iss, project, date, onTrack) {
  // Create tasks for Design issues AND CRM design deliverables
  if ((iss.issueType||'') !== 'Design' && (iss.kind||'') !== 'Deliverable') {
    Logger.log('Skipping non-design issue task: ' + iss.issueType);
    return '';
  }

  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) return '';
  var newId = 'T-' + Utilities.getUuid().substring(0,8).toUpperCase();
  var today = dateStr();

  // Resolve assignee — use manual selection if valid, else project lead from col F
  var assignee = resolveAssignee(iss.assignedTo, project);
  if (!assignee) {
    Logger.log('No assignee found for issue in project: ' + project);
    return '';
  }

  var priority = 'Medium'; // Default — site issues are medium priority
  if (onTrack === 'At Risk' || onTrack === 'Delayed') priority = 'High';

  var srcLabel = (iss.reportedBy === 'Aman Raghuwanshi') ? 'CRM' : 'DPER';
  var taskType = (iss.kind === 'Deliverable') ? 'Design Deliverable — CRM' : 'Site Issue — Design';
  var notes    = iss.description + ' [From ' + srcLabel + ': ' + project + ', ' + date + ']';

  sheet.appendRow([
    newId,                    // A TaskID
    '',                       // B ProjectID
    project,                  // C ProjectName
    assignee,                 // D AssignedTo
    taskType,                 // E Stage
    1,                        // F Disc. Multiplier (1.0 for issues)
    1,                        // G Stage Base Pts (1pt per issue)
    1,                        // H Units
    1,                        // I Weighted Pts = 1 × 1 × 1
    today,                    // J AssignedDate
    iss.targetDate || addDaysToStr(today,2), // K Deadline
    '',                       // L Area
    '',                       // M Drawing
    'Not Started',            // N SelfStatus
    '',                       // O SelfStatusDate
    '',                       // P ActualCompletionDate
    'Pending',                // Q LeadApproved
    '',                       // R ApprovedBy
    '',                       // S ApprovalDate
    '',                       // T RevisionTag
    notes,                    // U Notes
    srcLabel + ' — ' + (iss.reportedBy||'Execution Lead'), // V AssignedBy
    priority,                 // W Priority
  ]);
  parkRowIfStalled(sheet, project || '');  // stalled project → auto-park

  Logger.log('Issue task created: ' + newId + ' → ' + iss.assignedTo + ' / ' + taskType);
  return newId;
}

// ── writeSiteDecisions ─────────────────────────────────────
function writeSiteDecisions(subId, date, project, decisions, onTrack, lead) {
  if (!decisions || decisions.length === 0) return;
  var sheet  = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) return;
  var today  = dateStr();
  var priority = (onTrack === 'At Risk' || onTrack === 'Delayed') ? 'High' : 'Medium';

  decisions.forEach(function(dec) {
    if (!dec.description) return;
    var newId = 'T-' + Utilities.getUuid().substring(0,8).toUpperCase();
    sheet.appendRow([
      newId,
      '',
      project,
      'Siddharth Inani',
      'Decision Pending',
      1,             // F Disc. Multiplier
      0,             // G Stage Base Pts
      1,             // H Units
      0,             // I Weighted Pts (0 — decisions are tracking tasks, not scored)
      today,
      addDaysToStr(today, 1), // Deadline = tomorrow (site is waiting)
      '',
      '',
      'Not Started',
      '', '',
      'Pending',
      '', '', '',
      dec.description + ' [From DPER: ' + project + ', ' + date + ', Lead: ' + lead + ']',
      lead,
      priority,
    ]);
    Logger.log('Decision task created for Siddharth: ' + newId);
  });
}

// ── getSiteIssues — for projects dashboard ─────────────────
function getSiteIssues(project) {
  var sheet = db().getSheetByName(SITE_ISSUES_TAB);
  if (!sheet) return {issues:[]};
  var rows  = sheet.getDataRange().getValues();
  var issues= [];
  for (var i=1; i<rows.length; i++) {
    var r    = rows[i];
    var proj = String(r[3]||'').trim();
    if (project && proj !== project) continue;
    var issStatus  = String(r[10]||'').trim();
    var resolvedDt = String(r[13]||'').substring(0,10); // N Resolved Date
    var sevenDaysAgo = addDaysToStr(dateStr(), -7);
    // Skip resolved issues older than 7 days
    if (issStatus === 'Resolved' && resolvedDt < sevenDaysAgo) continue;
    issues.push({
      issueId   : String(r[0]||''),
      subId     : String(r[1]||''),
      date      : String(r[2]||'').substring(0,10),
      project   : proj,
      issueNum  : r[4]||1,
      issueType : String(r[5]||''),
      description:String(r[6]||''),
      assignedTo: String(r[7]||''),
      priority  : r[8]||2,
      targetDate: String(r[9]||''),
      status    : String(r[10]||''),
      taskId    : String(r[11]||''),
    });
  }
  return {issues: issues};
}

// ════════════════════════════════════════════════════════════════
// getBlockersThisWeek — everything blocking work right now, collected from
// all three forms, for the current week (Mon–Sun):
//   • Tasks explicitly marked "Blocked" in the DPR (active, any date)
//   • DPR daily "Anything blocking you?"  (DAILY_SUMMARY col G)
//   • DPER "Blocking Tomorrow"            (SITE_EXECUTION col T)
// ════════════════════════════════════════════════════════════════
function latestNoteText(n) {
  if (!n) return '';
  var last = String(n).split('|').pop().trim();
  return last.replace(/^\d{4}-\d{2}-\d{2}:\s*/, '');
}
function getBlockersThisWeek() {
  var s = db();
  var mon = mondayOf(new Date());
  var monStr = dateStr(mon);
  var sun = new Date(mon); sun.setDate(sun.getDate() + 6);
  var sunStr = dateStr(sun);
  function inWeek(d) { return d && d >= monStr && d <= sunStr; }
  var out = [];

  // 1. Explicitly blocked TASKS (active until unblocked)
  var aSheet = s.getSheetByName(ASSIGN_TAB);
  if (aSheet && aSheet.getLastRow() > 1) {
    var ar = aSheet.getDataRange().getValues();
    for (var i = 1; i < ar.length; i++) {
      if (String(ar[i][13] || '').trim() === 'Blocked') {       // N SelfStatus
        out.push({ type:'Task', who:String(ar[i][3]||''), project:String(ar[i][2]||''),
          detail:String(ar[i][4]||''), reason:latestNoteText(ar[i][20]) || 'No reason given',
          date:dateStr(ar[i][14]) });                            // O SelfStatusDate
      }
    }
  }
  // 2. DPR daily blockers this week
  var dSheet = s.getSheetByName(SUMMARY_TAB);
  if (dSheet && dSheet.getLastRow() > 1) {
    var dr = dSheet.getDataRange().getValues();
    for (var j = 1; j < dr.length; j++) {
      var bl = String(dr[j][6] || '').trim();                    // G Anything Blocking
      var dt = dateStr(dr[j][0]);                                // A Date
      if (bl && inWeek(dt)) out.push({ type:'Daily', who:String(dr[j][2]||''), project:'',
        detail:'Daily report', reason:bl, date:dt });
    }
  }
  // 3. DPER site blockers this week
  var eSheet = s.getSheetByName(SITE_EXEC_TAB);
  if (eSheet && eSheet.getLastRow() > 1) {
    var er = eSheet.getDataRange().getValues();
    for (var k = 1; k < er.length; k++) {
      var sb = String(er[k][19] || '').trim();                   // T Blocking Tomorrow
      var edt = dateStr(er[k][1]);                               // B Date
      if (sb && inWeek(edt)) out.push({ type:'Site', who:String(er[k][4]||''),
        project:String(er[k][3]||''), detail:'Site execution', reason:sb, date:edt });
    }
  }
  out.sort(function(a, b){ return String(b.date||'').localeCompare(String(a.date||'')); });
  return { weekStart:monStr, weekEnd:sunStr, count:out.length, blockers:out };
}

// ════════════════════════════════════════════════════════════════
// PROJECTS HEALTH — "which projects need attention today" (EPIC B)
// Stages (PROJECTS col C): Briefing · Arch Design · Arch WDs · Civil Execution ·
//   Interior Design · Interior WDs · Interiors Execution · Finishing/Handover ·
//   On hold · New Lead · Completed · Dead
// No money amounts are returned — bill STATUS + dates only.
// ════════════════════════════════════════════════════════════════
var DESIGN_STAGES = ['Briefing','Arch Design','Arch WDs','Interior Design','Interior WDs'];
function pjDaysAgo(d, today){ return d ? Math.round((new Date(today)-new Date(d))/86400000) : 9999; }
function engagementOverdue(stage, p, today){
  if (stage === 'Civil Execution')      return pjDaysAgo(p.lastVisit, today) > 15;
  if (stage === 'Interiors Execution')  return pjDaysAgo(p.lastVisit, today) > 7;
  if (stage === 'Finishing/Handover')   return pjDaysAgo(p.lastConn,  today) > 15;
  if (DESIGN_STAGES.indexOf(stage) > -1){
    var last = [p.lastConn, p.lastVisit, p.lastMeeting].filter(Boolean).sort().pop() || '';
    return pjDaysAgo(last, today) > 7;
  }
  return false;
}
function getProjectsHealth(){
  var s = db(), today = dateStr();
  var mon = dateStr(mondayOf(new Date())), sun = addDaysToStr(mon, 6);

  var pSheet = s.getSheetByName(PROJECTS_TAB);
  var pr = pSheet ? pSheet.getDataRange().getValues() : [[]];
  var P = {}, order = [];
  for (var i=1;i<pr.length;i++){
    var name = String(pr[i][1]||'').trim(); if(!name) continue;
    var key = name.toLowerCase();
    P[key] = { name:name, stage:String(pr[i][2]||'').trim(), discipline:String(pr[i][3]||''),
      lead:String(pr[i][5]||''), client:String(pr[i][6]||''),
      openTasks:0, tasksThisWeek:0, delayedTasks:0, openIssues:0,
      lastTask:'', lastVisit:'', lastConn:'', lastMeeting:'', connBy:'',
      lastBillRaised:'', lastPayCleared:'', billOverdue:false, dperPct:'' };
    order.push(key);
  }
  function newer(o,f,d){ if(d && o[f] < d) o[f]=d; }

  var aSheet = s.getSheetByName(ASSIGN_TAB);
  if(aSheet){ var ar=aSheet.getDataRange().getValues();
    for(var j=1;j<ar.length;j++){ var pk=String(ar[j][2]||'').trim().toLowerCase(); if(!P[pk])continue;
      var ss=String(ar[j][13]||'').trim(), ssd=cellDate(ar[j][14]), td=cellDate(ar[j][10]);
      var dispo=String(ar[j][COL_BLK_DISPO-1]||'').trim();
      if(ssd) newer(P[pk],'lastTask',ssd);
      var inactive=(ss==='Blocked'&&dispo==='Cancelled')||ss==='Parked'||ss==='Reassigned'||ss==='Work Not Done';
      if(ss!=='Done' && !inactive){ P[pk].openTasks++; if(td && td<today) P[pk].delayedTasks++; }
      if(ssd>=mon && ssd<=sun) P[pk].tasksThisWeek++;
    }
  }
  var iSheet=s.getSheetByName(SITE_ISSUES_TAB);
  if(iSheet){ var ir=iSheet.getDataRange().getValues();
    for(var k=1;k<ir.length;k++){ var pk=String(ir[k][3]||'').trim().toLowerCase(); if(!P[pk])continue;
      if(String(ir[k][10]||'').trim()!=='Resolved') P[pk].openIssues++; }
  }
  var cSheet=s.getSheetByName(CRM_LOG_TAB);
  if(cSheet){ var cr=cSheet.getDataRange().getValues();
    for(var m=1;m<cr.length;m++){ var pk=String(cr[m][6]||'').trim().toLowerCase(); if(!P[pk])continue;
      if(String(cr[m][4]||'').trim()!=='Client Connection') continue;
      var cd=cellDate(cr[m][2]), typ=String(cr[m][5]||'');
      if(cd && cd>P[pk].lastConn){ P[pk].lastConn=cd; P[pk].connBy=String(cr[m][3]||''); }
      if(/meet|visit/i.test(typ)) newer(P[pk],'lastMeeting',cd);
    }
  }
  var eSheet=s.getSheetByName(SITE_EXEC_TAB);
  if(eSheet){ var er=eSheet.getDataRange().getValues();
    for(var n=1;n<er.length;n++){ var pk=String(er[n][3]||'').trim().toLowerCase(); if(!P[pk])continue;
      var ed=cellDate(er[n][1]); if(ed) newer(P[pk],'lastVisit',ed);
      var pct=String(er[n][9]||'').trim(); if(pct) P[pk].dperPct=pct; }
  }
  var bSheet=s.getSheetByName(BILLING_TAB);
  if(bSheet){ var br=bSheet.getDataRange().getValues();
    for(var q=1;q<br.length;q++){ var pk=String(br[q][2]||'').trim().toLowerCase(); if(!P[pk])continue;
      var bd=cellDate(br[q][3]), rd=cellDate(br[q][6]);
      var amt=parseFloat(br[q][4])||0, rec=parseFloat(br[q][5])||0, st=String(br[q][8]||'').trim().toLowerCase();
      if(bd) newer(P[pk],'lastBillRaised',bd);
      if(rd) newer(P[pk],'lastPayCleared',rd);
      var unpaid = rec < amt && !/clear|paid|received|complete/.test(st);
      if(unpaid && bd && pjDaysAgo(bd,today)>30) P[pk].billOverdue=true;
    }
  }

  var cards=[], newLeadKeys=[];
  order.forEach(function(key){
    var p=P[key]; if(p.stage==='Dead') return;
    if(p.stage==='New Lead'){ newLeadKeys.push(p); return; }
    var la=[p.lastTask,p.lastVisit,p.lastConn].filter(Boolean).sort().pop()||'';
    p.lastActivity=la; p.lastActivityDays = la?pjDaysAgo(la,today):null;
    p.stalled = p.stage==='Stalled';
    var inactive = p.stage==='On hold'||p.stage==='Completed'||p.stalled;
    p.neglected = !inactive && (!la || pjDaysAgo(la,today)>7);
    p.engagementOverdue = inactive ? false : engagementOverdue(p.stage, p, today);
    p.paymentOverdue = p.billOverdue;
    p.behind = p.delayedTasks>0;
    var rag='green';
    if(inactive) rag='grey';
    else if(p.billOverdue || p.engagementOverdue || (la && pjDaysAgo(la,today)>14)) rag='red';
    else if(p.neglected || p.behind || p.openIssues>0) rag='amber';
    p.rag=rag;
    p.billStatus = p.billOverdue ? 'overdue'
      : (p.lastPayCleared && (!p.lastBillRaised || p.lastPayCleared>=p.lastBillRaised) ? 'cleared'
      : (p.lastBillRaised ? 'due' : '—'));
    cards.push(p);
  });
  var rank={red:0,amber:1,green:2,grey:3};
  cards.sort(function(a,b){ return (rank[a.rag]-rank[b.rag]) || (b.openIssues-a.openIssues) ||
    ((b.lastActivityDays||0)-(a.lastActivityDays||0)); });

  // New Leads pipeline — status/last-connection from LEADS, joined by client/project name
  var leadIdx={};
  var lSheet=s.getSheetByName(LEADS_TAB);
  if(lSheet){ var lr=lSheet.getDataRange().getValues();
    for(var x=1;x<lr.length;x++){ var nm=String(lr[x][1]||'').trim().toLowerCase(); if(!nm)continue;
      leadIdx[nm]={ status:String(lr[x][5]||''), lastContacted:cellDate(lr[x][8]), by:String(lr[x][6]||'') }; }
  }
  var newLeads = newLeadKeys.map(function(p){
    var L = leadIdx[p.name.toLowerCase()] || leadIdx[(p.client||'').toLowerCase()] || {};
    return { project:p.name, status:L.status||'—', lastConnection:L.lastContacted||p.lastConn||'',
      connBy:L.by||p.connBy||'', tasks:p.openTasks };
  });

  return { today:today, weekStart:mon, projects:cards, newLeads:newLeads, counts:{
      total:cards.length,
      neglected:cards.filter(c=>c.neglected).length,
      paymentOverdue:cards.filter(c=>c.paymentOverdue).length,
      engagementOverdue:cards.filter(c=>c.engagementOverdue).length,
      behind:cards.filter(c=>c.behind).length }};
}

// Per-project drawer detail
function getProjectDetail(project){
  var s=db(), pj=String(project||'').trim(), pjl=pj.toLowerCase(), today=dateStr();
  var mon=dateStr(mondayOf(new Date())), sun=addDaysToStr(mon,6);
  var out={ project:pj, openIssues:[], actionItems:[], tasks:[], siteVisits:[], connections:[],
    otherActivities:[], billables:[], lastBillRaised:'', lastPayCleared:'', dperPct:'' };

  // Action items raised in meeting / site-visit logs (DECISION_LOG) — shown under Issues.
  var dlSheet=s.getSheetByName(DECISION_LOG_TAB);
  if(dlSheet && dlSheet.getLastRow()>1){ var dl=dlSheet.getDataRange().getValues();
    for(var di=1;di<dl.length;di++){ if(String(dl[di][2]||'').trim().toLowerCase()!==pjl)continue;
      out.actionItems.push({ category:String(dl[di][4]||''), owner:String(dl[di][5]||''),
        text:String(dl[di][6]||''), date:cellDate(dl[di][3]), status:String(dl[di][8]||'Open') }); }
    out.actionItems.sort((a,b)=>(b.date||'').localeCompare(a.date||''));
  }

  var iSheet=s.getSheetByName(SITE_ISSUES_TAB);
  if(iSheet){ var ir=iSheet.getDataRange().getValues();
    for(var k=1;k<ir.length;k++){ if(String(ir[k][3]||'').trim().toLowerCase()!==pjl)continue;
      var st=String(ir[k][10]||'').trim(); if(st==='Resolved')continue;
      out.openIssues.push({ desc:String(ir[k][6]||''), type:String(ir[k][5]||''), status:st,
        assignedTo:String(ir[k][7]||''), targetDate:cellDate(ir[k][9]), date:cellDate(ir[k][2]) }); }
    var iord={Open:0,Escalated:1}; out.openIssues.sort((a,b)=>(iord[a.status]??2)-(iord[b.status]??2));
  }
  var aSheet=s.getSheetByName(ASSIGN_TAB);
  if(aSheet){ var ar=aSheet.getDataRange().getValues();
    for(var j=1;j<ar.length;j++){ if(String(ar[j][2]||'').trim().toLowerCase()!==pjl)continue;
      var ss=String(ar[j][13]||'').trim(), ssd=cellDate(ar[j][14]), tt=String(ar[j][4]||'');
      var dispo=String(ar[j][COL_BLK_DISPO-1]||'').trim();
      if((ss==='Blocked'&&dispo==='Cancelled')||ss==='Reassigned') continue; // exclude approved-blocked
      var done = ss==='Done';
      if(done && !(ssd>=mon && ssd<=sun)) continue;  // only this week's done
      // visit task → hours from col G when done
      if(isVisitTask(tt) && done){ out.siteVisits.push({ who:String(ar[j][3]||''), hours:parseFloat(ar[j][6])||'', date:ssd, type:tt }); }
      out.tasks.push({ taskType:tt, assignedTo:String(ar[j][3]||''), selfStatus:ss, date:ssd, deadline:cellDate(ar[j][10]) });
    }
  }
  var eSheet=s.getSheetByName(SITE_EXEC_TAB);
  if(eSheet){ var er=eSheet.getDataRange().getValues();
    for(var n=1;n<er.length;n++){ if(String(er[n][3]||'').trim().toLowerCase()!==pjl)continue;
      var ed=cellDate(er[n][1]); out.siteVisits.push({ who:String(er[n][4]||''), hours:'', date:ed, type:'Site report (DPER)' });
      var pct=String(er[n][9]||'').trim(); if(pct) out.dperPct=pct; }
  }
  out.siteVisits.sort((a,b)=>(b.date||'').localeCompare(a.date||''));
  var cSheet=s.getSheetByName(CRM_LOG_TAB);
  if(cSheet){ var cr=cSheet.getDataRange().getValues();
    for(var m=1;m<cr.length;m++){ if(String(cr[m][6]||'').trim().toLowerCase()!==pjl)continue;
      var cat=String(cr[m][4]||'').trim();
      var rec={ date:cellDate(cr[m][2]), by:String(cr[m][3]||''), type:String(cr[m][5]||''), notes:String(cr[m][7]||'') };
      if(cat==='Client Connection') out.connections.push(rec);
      else out.otherActivities.push(rec); }   // Vendor / Site issues / TnCP / BNI etc.
    out.connections.sort((a,b)=>(b.date||'').localeCompare(a.date||''));
    out.otherActivities.sort((a,b)=>(b.date||'').localeCompare(a.date||''));
  }
  var bSheet=s.getSheetByName(BILLING_TAB);
  if(bSheet){ var br=bSheet.getDataRange().getValues();
    for(var q=1;q<br.length;q++){ if(String(br[q][2]||'').trim().toLowerCase()!==pjl)continue;
      var bd=cellDate(br[q][3]), rd=cellDate(br[q][6]);
      if(bd && bd>out.lastBillRaised) out.lastBillRaised=bd;
      if(rd && rd>out.lastPayCleared) out.lastPayCleared=rd; }
  }
  // Billables raised on DPR (BILL_REQUESTS) — stage + status only, NO amounts.
  var brq=s.getSheetByName(BILL_REQ_TAB);
  if(brq && brq.getLastRow()>1){ var qr=brq.getDataRange().getValues();
    for(var bi=1;bi<qr.length;bi++){ if(String(qr[bi][3]||'').trim().toLowerCase()!==pjl)continue;
      out.billables.push({ date:cellDate(qr[bi][1]), raisedBy:String(qr[bi][2]||''),
        discipline:String(qr[bi][4]||''), stage:String(qr[bi][5]||''), status:String(qr[bi][6]||'') }); }
    out.billables.sort((a,b)=>(b.date||'').localeCompare(a.date||''));
  }
  return out;
}

// ════════════════════════════════════════════════════════════════
// getWeeklyProjectDigest — per-project weekly summary (managers only).
// Buckets every source by project for the Mon–Sat week. "Active" = at least
// one item this week (its own page); the rest are listed name/stage/last-seen.
// NO money amounts — billing shows events only.
// ════════════════════════════════════════════════════════════════
function getWeeklyProjectDigest(weekStart){
  var s=db();
  var mon=weekStart||dateStr(mondayOf(new Date())), sat=addDaysToStr(mon,5);
  function inWk(d){ return d && d>=mon && d<=sat; }
  var P={}, order=[];
  var pSheet=s.getSheetByName(PROJECTS_TAB);
  if(pSheet){ var pr=pSheet.getDataRange().getValues();
    for(var i=1;i<pr.length;i++){ var nm=String(pr[i][1]||'').trim(); if(!nm) continue; var k=nm.toLowerCase();
      if(!P[k]){ P[k]={name:nm, stage:String(pr[i][2]||'').trim(), lastActivity:'',
        meetings:[], tasksDone:[], connections:[], others:[], billing:[], visits:[], actions:[], issues:[]}; order.push(k); } }
  }
  function touch(k,d){ if(P[k]&&d&&(!P[k].lastActivity||d>P[k].lastActivity)) P[k].lastActivity=d; }

  var ml=s.getSheetByName(MEETING_LOG_TAB);
  if(ml){ var mr=ml.getDataRange().getValues();
    for(var i=1;i<mr.length;i++){ if(String(mr[i][15]||'').trim()==='Deleted')continue;
      var k=String(mr[i][4]||'').trim().toLowerCase(); if(!P[k])continue; var d=cellDate(mr[i][1]); touch(k,d);
      if(inWk(d)) P[k].meetings.push({date:d, type:String(mr[i][3]||''), who:[String(mr[i][6]||''),String(mr[i][7]||'')].filter(Boolean).join(', ')}); } }
  var a=s.getSheetByName(ASSIGN_TAB);
  if(a){ var ar=a.getDataRange().getValues();
    for(var i=1;i<ar.length;i++){ var k=String(ar[i][2]||'').trim().toLowerCase(); if(!P[k])continue;
      var ss=String(ar[i][13]||'').trim(), sd=cellDate(ar[i][14]); touch(k,sd);
      if(ss==='Done'&&inWk(sd)) P[k].tasksDone.push({date:sd, who:String(ar[i][3]||''), task:String(ar[i][4]||'')}); } }
  var c=s.getSheetByName(CRM_LOG_TAB);
  if(c){ var cr=c.getDataRange().getValues();
    for(var i=1;i<cr.length;i++){ var k=String(cr[i][6]||'').trim().toLowerCase(); if(!P[k])continue;
      var d=cellDate(cr[i][2]); touch(k,d); if(!inWk(d))continue; var rec={date:d, by:String(cr[i][3]||''), type:String(cr[i][5]||''), notes:String(cr[i][7]||'')};
      if(String(cr[i][4]||'').trim()==='Client Connection') P[k].connections.push(rec); else P[k].others.push(rec); } }
  var b=s.getSheetByName(BILLING_TAB);
  if(b){ var br=b.getDataRange().getValues();
    for(var i=1;i<br.length;i++){ var k=String(br[i][2]||'').trim().toLowerCase(); if(!P[k])continue;
      var bd=cellDate(br[i][3]), rd=cellDate(br[i][6]); touch(k,bd); touch(k,rd);
      if(inWk(bd)) P[k].billing.push({date:bd, event:'Bill raised'});
      if(inWk(rd)) P[k].billing.push({date:rd, event:'Payment received'}); } }
  var brq=s.getSheetByName(BILL_REQ_TAB);
  if(brq){ var qr=brq.getDataRange().getValues();
    for(var i=1;i<qr.length;i++){ var k=String(qr[i][3]||'').trim().toLowerCase(); if(!P[k])continue;
      var d=cellDate(qr[i][1]); touch(k,d); if(inWk(d)) P[k].billing.push({date:d, event:'Billable reached — '+String(qr[i][4]||'')+' · '+String(qr[i][5]||'')}); } }
  var e=s.getSheetByName(SITE_EXEC_TAB);
  if(e){ var er=e.getDataRange().getValues();
    for(var i=1;i<er.length;i++){ var k=String(er[i][3]||'').trim().toLowerCase(); if(!P[k])continue;
      var d=cellDate(er[i][1]); touch(k,d); if(inWk(d)) P[k].visits.push({date:d, who:String(er[i][4]||'')}); } }
  var dl=s.getSheetByName(DECISION_LOG_TAB);
  if(dl){ var dr=dl.getDataRange().getValues();
    for(var i=1;i<dr.length;i++){ if(String(dr[i][8]||'').trim()==='Deleted')continue;
      var k=String(dr[i][2]||'').trim().toLowerCase(); if(!P[k])continue; var d=cellDate(dr[i][3]); touch(k,d);
      if(inWk(d)) P[k].actions.push({date:d, cat:String(dr[i][4]||''), owner:String(dr[i][5]||''), text:String(dr[i][6]||'')}); } }
  var iss=s.getSheetByName(SITE_ISSUES_TAB);
  if(iss){ var ir=iss.getDataRange().getValues();
    for(var i=1;i<ir.length;i++){ var k=String(ir[i][3]||'').trim().toLowerCase(); if(!P[k])continue;
      var d=cellDate(ir[i][2]); touch(k,d); if(inWk(d)) P[k].issues.push({date:d, desc:String(ir[i][6]||''), status:String(ir[i][10]||'')}); } }

  function cnt(p){ return p.meetings.length+p.tasksDone.length+p.connections.length+p.others.length+p.billing.length+p.visits.length+p.actions.length+p.issues.length; }
  var active=[], inactive=[];
  order.forEach(function(k){ var p=P[k]; if(p.stage==='Dead') return;
    if(cnt(p)>0) active.push(p); else inactive.push({name:p.name, stage:p.stage||'—', lastActivity:p.lastActivity||''}); });
  active.sort(function(x,y){ return cnt(y)-cnt(x); });
  inactive.sort(function(x,y){ return (x.lastActivity||'').localeCompare(y.lastActivity||''); });
  return {weekStart:mon, weekEnd:sat, active:active, inactive:inactive, counts:{active:active.length, inactive:inactive.length}};
}

// ── getSiteExecutionSummary — per project ──────────────────
function getSiteExecutionSummary(project) {
  var sheet = db().getSheetByName(SITE_EXEC_TAB);
  if (!sheet || sheet.getLastRow() < 2) return {entries:[]};
  var rows  = sheet.getDataRange().getValues();
  var entries = [];
  for (var i=1; i<rows.length; i++) {
    var r    = rows[i];
    var proj = String(r[3]||'').trim();
    if (project && proj !== project) continue;
    entries.push({
      subId      : String(r[0]||''),
      date       : String(r[1]||'').substring(0,10), // B=date
      project    : proj,
      lead       : String(r[4]||''),
      siteVisit  : String(r[5]||''),
      stage      : String(r[6]||''),
      worksToday : String(r[7]||''),
      pctPlan    : r[8]||0,
      pctOverall : r[9]||0,
      onTrack    : String(r[11]||''),
      delayReason: String(r[12]||''),
      blocking   : String(r[19]||''),
    });
  }
  // Return last 30 entries sorted recent first
  entries.sort(function(a,b){ return b.date.localeCompare(a.date); });
  return {entries: entries.slice(0,30)};
}

// ── migrateDPERData — run once from editor ─────────────────
// migrateDPERData removed — migration complete

// ── resolveIssue — called from Projects Dashboard ─────────
function resolveIssue(issueId) {
  var sheet = db().getSheetByName(SITE_ISSUES_TAB);
  if (!sheet) return {status:'error', message:'SITE_ISSUES tab not found'};
  var rows  = sheet.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (String(rows[i][0]||'').trim() === issueId) {
      sheet.getRange(i+1, 11).setValue('Resolved'); // K Status
      sheet.getRange(i+1, 14).setValue(dateStr());  // N Resolved Date
      Logger.log('Issue resolved manually: ' + issueId);
      return {status:'ok', issueId:issueId, resolvedDate:dateStr()};
    }
  }
  return {status:'error', message:'Issue not found: ' + issueId};
}

// ════════════════════════════════════════════════════════════════
// TRIGGER SETUP — run once manually from Apps Script editor
// ════════════════════════════════════════════════════════════════

// Run this ONCE manually to set up the Monday 8AM trigger
function setupMondayTrigger() {
  // Delete existing visit scheduling triggers
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'syncVisitSchedule' ||
        t.getHandlerFunction() === 'pushVisitTasks') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Deleted trigger: ' + t.getHandlerFunction());
    }
  });

  // Create new Monday 8AM trigger
  ScriptApp.newTrigger('syncVisitSchedule')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  Logger.log('Monday 8AM trigger created for syncVisitSchedule');
}

// Run this to see all current triggers
function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    Logger.log(t.getHandlerFunction() + ' — ' + t.getEventType());
  });
}

// ════════════════════════════════════════════════════════════════
// getDeepakVisitSummary — weekly site visit coverage for Deepak
// Returns per-project visit status for the current or given week
// Sources: VISIT_PLANNER (assigned projects) + SITE_EXECUTION (DPER)
//          + TASK_ASSIGNMENTS (completed visit tasks)
// ════════════════════════════════════════════════════════════════
function getDeepakVisitSummary(weekStart) {
  var mon = weekStart || dateStr(mondayOf(new Date()));
  var sat = addDaysToStr(mon, 5);
  var today = dateStr();

  // ── 1. Deepak's assigned projects from VISIT_PLANNER ─────────
  var planSheet = db().getSheetByName(PLANNER_TAB);
  var assignedProjects = []; // projects Deepak is supposed to visit weekly
  if (planSheet && planSheet.getLastRow() > 3) {
    var pRows = planSheet.getRange(4, 1, planSheet.getLastRow()-3, 9).getValues();
    pRows.forEach(function(r) {
      var proj     = String(r[0]||'').trim();
      var active   = String(r[6]||'Yes').trim().toLowerCase();
      var assignee = String(r[2]||'').trim();
      if (!proj || active === 'no') return;
      var assignees = assignee.split(',').map(function(a){ return a.trim(); });
      if (assignees.indexOf('Deepak Soni') > -1 &&
          assignedProjects.indexOf(proj) === -1) {
        assignedProjects.push(proj);
      }
    });
  }

  // ── 2. Visits recorded via DPER (SITE_EXECUTION) ─────────────
  // Cols: B=Date(1) D=ProjectName(3) E=Lead(4) F=SiteVisitDone(5)
  var execSheet = db().getSheetByName(SITE_EXEC_TAB);
  var visitMap = {}; // project → [{date, siteVisit, source}]
  if (execSheet && execSheet.getLastRow() > 1) {
    var eRows = execSheet.getDataRange().getValues();
    for (var i = 1; i < eRows.length; i++) {
      var eDate    = cellDate(eRows[i][1]);               // B Date
      var eProj    = String(eRows[i][3]||'').trim();      // D ProjectName
      var eLead    = String(eRows[i][4]||'').trim();      // E Lead
      var eVisited = String(eRows[i][5]||'').trim();      // F SiteVisitDone
      if (eLead !== 'Deepak Soni') continue;
      if (!eDate || eDate < mon || eDate > sat) continue;
      if (!visitMap[eProj]) visitMap[eProj] = [];
      visitMap[eProj].push({date:eDate, siteVisit:eVisited, source:'DPER'});
    }
  }

  // ── 3. Completed visit tasks in TASK_ASSIGNMENTS ──────────────
  var asSheet = db().getSheetByName(ASSIGN_TAB);
  if (asSheet) {
    var asRows  = asSheet.getDataRange().getValues();
    var aHdrs   = asRows[0] ? asRows[0].map(function(h){return String(h||'').trim();}) : [];
    var aIs23   = aHdrs.length >= 23 || aHdrs.indexOf('Actual Completion Date') > -1;
    var A_ACTDT = aIs23 ? 15 : -1;
    var A_STATDT= 14;

    for (var j = 1; j < asRows.length; j++) {
      var ar = asRows[j];
      if (String(ar[3]||'').trim() !== 'Deepak Soni') continue;
      if (!isVisitTask(String(ar[4]||'').trim())) continue;
      if (String(ar[13]||'').trim() !== 'Done') continue;
      var doneDate = (A_ACTDT > -1 ? cellDate(ar[A_ACTDT]) : '') || cellDate(ar[A_STATDT]);
      if (!doneDate || doneDate < mon || doneDate > sat) continue;
      var aProj = String(ar[2]||'').trim();
      if (!visitMap[aProj]) visitMap[aProj] = [];
      visitMap[aProj].push({date:doneDate, siteVisit:'Yes', source:'task'});
    }
  }

  // ── 4. Build per-project result ───────────────────────────────
  var projects = assignedProjects.map(function(proj) {
    var visits     = visitMap[proj] || [];
    var siteVisits = visits.filter(function(v){ return v.siteVisit==='Yes'; });
    var dperDays   = visits.filter(function(v){ return v.source==='DPER'; }).length;
    return {
      project    : proj,
      visited    : siteVisits.length > 0,
      visitDates : siteVisits.map(function(v){ return v.date; }).filter(function(d,i,a){ return a.indexOf(d)===i; }).sort(),
      dperDays   : dperDays,
      unplanned  : false,
    };
  });

  // Also surface unplanned projects visited this week
  Object.keys(visitMap).forEach(function(proj) {
    if (assignedProjects.indexOf(proj) > -1) return;
    var sv = visitMap[proj].filter(function(v){ return v.siteVisit==='Yes'; });
    if (sv.length > 0) {
      projects.push({
        project    : proj,
        visited    : true,
        visitDates : sv.map(function(v){ return v.date; }).filter(function(d,i,a){ return a.indexOf(d)===i; }).sort(),
        dperDays   : visitMap[proj].filter(function(v){ return v.source==='DPER'; }).length,
        unplanned  : true,
      });
    }
  });

  var visited = projects.filter(function(p){ return p.visited; }).length;
  var total   = assignedProjects.length;

  return {
    weekStart   : mon,
    weekEnd     : sat,
    projects    : projects,
    visitedCount: visited,
    totalCount  : total,
    coveragePct : total > 0 ? Math.round(visited / total * 100) : 0,
  };
}

// ════════════════════════════════════════════════════════════════
// getDeepakWeeklyStats — Deepak Soni weekly scorecard
// action=getDeepakWeeklyStats&weekStart=2026-06-01
//
// Scoring (total /100):
//   Site Visits      /20 — active sites visited ≥1× / total sites
//   Client Comm      /10 — active sites with clientUpdated=Yes ≥1× / total
//   Task Completion  /20 — (assigned TO him + assigned BY him to team)
//                          tasks completed on time this week / total due
//   DPER Consistency /15 — days with ≥1 DPER submission / 6 × 15
//   Punctuality      /20 — same base as team (biometric)
//   Hours            /15 — same base as team (biometric)
// ════════════════════════════════════════════════════════════════
function getDeepakWeeklyStats(weekStart) {
  var s   = db();
  var mon = weekStart || dateStr(mondayOf(new Date()));
  var sat = addDaysToStr(mon, 5);

  // ── 1. Read active projects from CONFIG tab ───────────────
  // Scans ALL columns for the "DEEPAK ACTIVE PROJECTS" header (it lives in
  // col T, not col A), then reads names down that same column until blank.
  var activeProjects = [];
  var configSheet = s.getSheetByName(CONFIG_TAB);
  if (configSheet) {
    var cRows = configSheet.getDataRange().getValues();
    var hdrRow = -1, hdrCol = -1;
    for (var hr = 0; hr < cRows.length && hdrRow === -1; hr++) {
      for (var hc = 0; hc < cRows[hr].length; hc++) {
        if (String(cRows[hr][hc] || '').trim().toUpperCase() === 'DEEPAK ACTIVE PROJECTS') {
          hdrRow = hr; hdrCol = hc; break;
        }
      }
    }
    if (hdrRow > -1) {
      for (var ci = hdrRow + 1; ci < cRows.length; ci++) {
        var cell = String((cRows[ci] || [])[hdrCol] || '').trim();
        if (!cell) break;
        activeProjects.push(cell);
      }
    }
  }
  var totalSites = activeProjects.length;

  // ── 2. Scan SITE_EXECUTION for this week ─────────────────
  // Cols: B=Date(1) D=ProjectName(3) E=Lead(4) F=SiteVisitDone(5) R=ClientUpdated(17)
  var visitedSet = {};   // project → true if visited ≥1
  var clientSet  = {};   // project → true if clientUpdated=Yes ≥1
  var dperDaysSet= {};   // date → true (days on which ≥1 DPER was submitted)
  var projectDays= {};   // project → set of dates (for per-project DPER days)

  var execSheet = s.getSheetByName(SITE_EXEC_TAB);
  if (execSheet && execSheet.getLastRow() > 1) {
    var eRows = execSheet.getDataRange().getValues();
    for (var i = 1; i < eRows.length; i++) {
      var eDate   = cellDate(eRows[i][1]);
      var eProj   = String(eRows[i][3] || '').trim().toLowerCase();
      var eLead   = String(eRows[i][4] || '').trim().toLowerCase();
      var eVisit  = String(eRows[i][5] || '').trim().toLowerCase();
      var eClient = String(eRows[i][17]|| '').trim().toLowerCase();
      if (eLead.indexOf('deepak') === -1) continue;   // tolerate "Deepak"/"Deepak Soni"/case
      if (!eDate || eDate < mon || eDate > sat) continue;

      dperDaysSet[eDate] = true;
      if (!projectDays[eProj]) projectDays[eProj] = {};
      projectDays[eProj][eDate] = true;
      if (eVisit  === 'yes') visitedSet[eProj] = true;
      if (eClient === 'yes') clientSet[eProj]  = true;
    }
  }

  // ── 3. Per-project breakdown (match CONFIG names case-insensitively) ──
  var perProject = activeProjects.map(function(proj) {
    var k = proj.toLowerCase();
    return {
      project       : proj,
      visited       : !!visitedSet[k],
      clientUpdated : !!clientSet[k],
      dperDays      : projectDays[k] ? Object.keys(projectDays[k]).length : 0,
    };
  });

  var visitedCount = perProject.filter(function(p){ return p.visited; }).length;
  var clientCount  = perProject.filter(function(p){ return p.clientUpdated; }).length;
  var dperDaysCount= Object.keys(dperDaysSet).length; // unique days with ≥1 submission

  // ── 4. Task Completion from TASK_ASSIGNMENTS ─────────────
  // Pool: tasks assigned TO Deepak OR assigned BY Deepak to others
  // Assigned this week: AssignedDate in Mon–Sat
  // Completed same week: Done + LeadApproved=Yes + doneDate in Mon–Sat
  var asSheet = s.getSheetByName(ASSIGN_TAB);
  var tasksAssignedThisWeek = 0, tasksCompletedThisWeek = 0;
  var taskDetails = [];
  var dLate = 0, dWnd = 0;   // reliability: late/overdue and Work-Not-Done tasks

  if (asSheet && asSheet.getLastRow() > 1) {
    var asRows   = asSheet.getDataRange().getValues();
    var asHdrs   = asRows[0] ? asRows[0].map(function(h){ return String(h||'').trim(); }) : [];
    var is23     = asHdrs.length >= 23 || asHdrs.indexOf('Actual Completion Date') > -1;
    var COL_ACTDT  = is23 ? 15 : -1;
    var COL_STATDT = 14;
    var COL_LEADAP = is23 ? 16 : 15;
    var COL_ASSIGNDT = 9; // col J AssignedDate

    for (var j = 1; j < asRows.length; j++) {
      var ar         = asRows[j];
      var assignedTo = String(ar[3]  || '').trim();
      var assignedBy = String(ar[21] || '').trim(); // col V AssignedBy
      var assignedDt = cellDate(ar[COL_ASSIGNDT]);
      var doneDate   = (COL_ACTDT > -1 ? cellDate(ar[COL_ACTDT]) : '') || cellDate(ar[COL_STATDT]);
      var isApproved = String(ar[COL_LEADAP] || '').trim() === 'Yes';
      var selfStatus = String(ar[13] || '').trim();

      // Pool: assigned TO Deepak, OR assigned BY Deepak to someone else
      var inPool = (assignedTo === 'Deepak Soni') ||
                   (assignedBy === 'Deepak Soni' && assignedTo !== 'Deepak Soni');
      if (!inPool) continue;

      // Must have been assigned this week
      if (!assignedDt || assignedDt < mon || assignedDt > sat) continue;
      tasksAssignedThisWeek++;

      // Completed within the same week
      if (selfStatus === 'Done' && isApproved && doneDate && doneDate >= mon && doneDate <= (dateStr() > sat ? dateStr() : sat)) {
        tasksCompletedThisWeek++;
        taskDetails.push({
          project   : String(ar[2] || '').trim(),
          taskType  : String(ar[4] || '').trim(),
          assignedTo: assignedTo,
          assignedDt: assignedDt,
          doneDate  : doneDate,
        });
      }
    }

    // Reliability — Deepak's OWN tasks with a deadline in this week (same rule
    // as the team): late/overdue = −1, "Work Not Done" (reassigned) = −2.
    var nowD = dateStr();
    for (var k = 1; k < asRows.length; k++) {
      var rr = asRows[k];
      if (String(rr[3] || '').trim() !== 'Deepak Soni') continue;
      var dl = cellDate(rr[10]); // K Deadline
      if (!dl || dl < mon || dl > sat) continue;
      var ss = String(rr[13] || '').trim();
      if (ss === 'Work Not Done') { dWnd++; continue; }
      if (dl >= nowD) continue;                       // not yet due — not late
      var dd = (COL_ACTDT > -1 ? cellDate(rr[COL_ACTDT]) : '') || cellDate(rr[COL_STATDT]);
      var ot = String(rr[COL_LEADAP] || '').trim() === 'Yes' && dd && dd <= dl;
      if (!ot) dLate++;
    }
  }

  // ── 5. Biometric from DAILY_SUMMARY ──────────────────────
  var sumSheet    = s.getSheetByName(SUMMARY_TAB);
  var daysPresent = 0, lateCount = 0;
  var DEEPAK_THR  = '09:10';

  if (sumSheet && sumSheet.getLastRow() > 1) {
    var sRows = sumSheet.getDataRange().getValues();
    for (var si = 1; si < sRows.length; si++) {
      var rDate = cellDate(sRows[si][0]);
      var rName = String(sRows[si][2] || '').trim();
      if (rName !== 'Deepak Soni') continue;
      if (!rDate || rDate < mon || rDate > sat) continue;
      daysPresent++;
      var arrTime = String(sRows[si][1] || '').trim();
      if (arrTime && arrTime > DEEPAK_THR) lateCount++;
    }
  }

  var absentDays = 6 - daysPresent;

  // ── 6. Score calculation ──────────────────────────────────
  // Site Visits /20
  var s_visit  = totalSites > 0
    ? Math.round(visitedCount / totalSites * 20 * 10) / 10 : 0;

  // Client Communication /10
  var s_client = totalSites > 0
    ? Math.round(clientCount / totalSites * 10 * 10) / 10 : 0;

  // Task Completion /20 — completed same week / assigned this week
  // 0 assigned → full marks (nothing to complete)
  var s_tasks  = tasksAssignedThisWeek > 0
    ? Math.round(tasksCompletedThisWeek / tasksAssignedThisWeek * 20 * 10) / 10
    : 20;

  // DPER Consistency /15 — days with ≥1 submission / 6 (mirrors team DPR formula)
  var s_dper   = Math.min(15, Math.round(dperDaysCount / 6 * 15 * 10) / 10);

  // Punctuality /15 — same base as team (punctBase already on a /15 scale)
  var punctBase   = Math.max(0, 15 - lateCount * 2 - absentDays * 2);
  var s_punct     = Math.min(15, Math.round(punctBase * 10) / 10);

  // Hours /10 — same base as team (hrsBase already on a /10 scale)
  var hrsBase  = Math.max(0, Math.round((daysPresent / 6 * 10 - absentDays * 1.5) * 10) / 10);
  var s_hrs    = Math.min(10, Math.round(hrsBase * 10) / 10);

  // Reliability /10 — −1 per late/overdue task, −2 per Work Not Done
  var reliabilityPenalty = dLate * 1 + dWnd * 2;
  var s_reliability = Math.max(0, Math.round((10 - reliabilityPenalty) * 10) / 10);

  var total = Math.round((s_visit + s_client + s_tasks + s_dper + s_punct + s_hrs + s_reliability) * 10) / 10;

  return {
    weekStart      : mon,
    weekEnd        : sat,
    totalSites     : totalSites,
    visitedCount   : visitedCount,
    clientCount    : clientCount,
    dperDaysCount  : dperDaysCount,
    daysPresent    : daysPresent,
    lateCount      : lateCount,
    absentDays     : absentDays,
    tasksAssigned  : tasksAssignedThisWeek,
    tasksCompleted : tasksCompletedThisWeek,
    taskDetails    : taskDetails,
    perProject     : perProject,
    lateTaskCount  : dLate,
    workNotDone    : dWnd,
    reliabilityPenalty : reliabilityPenalty,
    scores: {
      s_visit  : s_visit,
      s_client : s_client,
      s_tasks  : s_tasks,
      s_dper   : s_dper,
      s_punct  : s_punct,
      s_hrs    : s_hrs,
      s_reliability : s_reliability,
      total    : total,
    },
  };
}

// ════════════════════════════════════════════════════════════════
// getAmanWeeklyStats — CRM (Aman) weekly scorecard /100
//   Output /50:
//     Client Connection Coverage /20 — ongoing projects connected ÷ total
//     Lead Management            /15 — leads 24hr-contacted ÷ new leads
//     Revenue Collection         /15 — min(collected÷billsRaised, 70%)/70%
//   DPR Consistency  /15 — AMAN_DAILY submission days ÷ 6
//   Punctuality      /20 — biometric (same base as team)
//   Hours            /15 — biometric (same base as team)
// ════════════════════════════════════════════════════════════════
function getAmanWeeklyStats(weekStart, member) {
  var s   = db();
  var mon = weekStart || dateStr(mondayOf(new Date()));
  var sat = addDaysToStr(mon, 5);
  var who = member || 'Aman Raghuwanshi';
  var whoLC = who.toLowerCase();   // CRM_LOG stores "aman raghuwanshi" (lowercase) — match case-insensitively
  var COLLECTION_TARGET = 0.70;  // 70% of bills raised cleared = full marks

  // ── 1. Ongoing projects (denominator) — active design/WD/execution stages only ──
  // Excludes New Lead, On hold, Completed, Dead (and legacy closed/cancelled/proposal).
  var EXCL = ['completed','closed','dead','cancelled','proposal','new lead','hold','on hold','finishing','handover','stalled'];
  function isExcluded(stat){
    stat = String(stat||'').toLowerCase();
    for (var i=0;i<EXCL.length;i++){ if (stat.indexOf(EXCL[i]) > -1) return true; }
    return false;
  }
  var ongoing = [];
  var projSheet = s.getSheetByName(PROJECTS_TAB);
  if (projSheet) {
    var pRows = projSheet.getDataRange().getValues();
    for (var pi = 1; pi < pRows.length; pi++) {
      var pName = String(pRows[pi][1]||'').trim();
      if (!pName || isExcluded(pRows[pi][2])) continue;
      ongoing.push(pName);
    }
  }
  var totalOngoing = ongoing.length;

  // ── 2a. DPR days — distinct days Aman filed a CRM report (CRM_LOG, by member) ──
  // (AMAN_DAILY is deprecated — merged into CRM_LOG.)
  var dprDaysSet = {};     // date → true
  var dLogSheet = s.getSheetByName(CRM_LOG_TAB);
  if (dLogSheet && dLogSheet.getLastRow() > 1) {
    var dRows = dLogSheet.getDataRange().getValues();
    for (var di = 1; di < dRows.length; di++) {
      var dDate = cellDate(dRows[di][2]);          // C Date
      var dMem  = String(dRows[di][3]||'').trim().toLowerCase(); // D Member
      if (dMem !== whoLC) continue;
      if (!dDate || dDate < mon || dDate > sat) continue;
      dprDaysSet[dDate] = true;
    }
  }

  // ── 2b. Client connections this week — CRM_LOG (Category = Client Connection) ──
  var connected = {};      // project lower → true
  var visitsDone = 0, meetingsDone = 0;  // derived from contact type
  var logSheet = s.getSheetByName(CRM_LOG_TAB);
  if (logSheet && logSheet.getLastRow() > 1) {
    var cRows = logSheet.getDataRange().getValues();
    for (var ci = 1; ci < cRows.length; ci++) {
      var cDate = cellDate(cRows[ci][2]);          // C Date
      var cMem  = String(cRows[ci][3]||'').trim().toLowerCase(); // D Member
      var cCat  = String(cRows[ci][4]||'').trim(); // E Category
      if (cMem !== whoLC || cCat !== 'Client Connection') continue;
      if (!cDate || cDate < mon || cDate > sat) continue;
      var pr = String(cRows[ci][6]||'').trim();    // G Project
      if (pr) connected[pr.toLowerCase()] = true;
      var ty = String(cRows[ci][5]||'');           // F Type / Activity
      if (ty.indexOf('Site Visit') > -1)   visitsDone++;
      else if (ty.indexOf('Meeting') > -1) meetingsDone++;
    }
  }

  // A meeting / site visit logged by CRM this week also counts as a connection
  var asSheet = s.getSheetByName(ASSIGN_TAB);
  if (asSheet && asSheet.getLastRow() > 1) {
    var aRows = asSheet.getDataRange().getValues();
    for (var ai = 1; ai < aRows.length; ai++) {
      var aType = String(aRows[ai][4]  || '').trim();   // E Type
      var aProj = String(aRows[ai][2]  || '').trim();   // C ProjectName
      var aBy   = String(aRows[ai][21] || '').trim();   // V AssignedBy
      var aDate = cellDate(aRows[ai][9]);               // J AssignedDate
      if (aBy.indexOf('CRM') !== 0) continue;           // created via CRM form
      if (aType !== 'Meeting' && aType !== 'Site Visit') continue;
      if (!aDate || aDate < mon || aDate > sat) continue;
      if (aProj) connected[aProj.toLowerCase()] = true;
    }
  }

  // A monthly client feedback collected this week also counts as a connection
  var fbSheet = s.getSheetByName(FEEDBACK_TAB);
  if (fbSheet && fbSheet.getLastRow() > 1) {
    var fRows = fbSheet.getDataRange().getValues();
    for (var fi = 1; fi < fRows.length; fi++) {
      var fDate = cellDate(fRows[fi][2]);               // C Date Recorded
      var fProj = String(fRows[fi][4] || '').trim();    // E Project
      if (String(fRows[fi][3]||'').trim().toLowerCase() !== whoLC) continue; // D Recorded By
      if (!fDate || fDate < mon || fDate > sat) continue;
      if (fProj) connected[fProj.toLowerCase()] = true;
    }
  }

  var perProject = ongoing.map(function(p){
    return { project:p, connected: !!connected[p.toLowerCase()] };
  });
  var connectedCount = perProject.filter(function(p){ return p.connected; }).length;
  var dprDaysCount   = Object.keys(dprDaysSet).length;

  // ── 3. Lead pipeline (LEADS) ──
  // Score = open leads worked this week ÷ leads that needed attention.
  // "needed attention" = open leads (Not Contacted / Contacted over call)
  // plus any lead touched this week (Last Contacted in week). Untouched open
  // leads drag the score down; if no leads need attention at all → N/A.
  var OPEN_LEAD = ['', 'Not Contacted', 'Contacted over call'];
  var leadWorked = 0, leadOpenUntouched = 0;
  var leadsThisWeek = 0, leads24 = 0;   // kept for display only
  var leadSlaMissed = 0;                // reliability: 24hr first-contact misses
  var nowD = dateStr();
  var leadsSheet = s.getSheetByName(LEADS_TAB);
  if (leadsSheet && leadsSheet.getLastRow() > 1) {
    var lRows = leadsSheet.getDataRange().getValues();
    for (var li = 1; li < lRows.length; li++) {
      var lStatus  = String(lRows[li][5] || '').trim();   // F Lead Status
      var lLast    = cellDate(lRows[li][8]);              // I Last Contacted
      var lCreated = cellDate(lRows[li][7]);              // H Lead Creation Date
      var isOpen   = OPEN_LEAD.indexOf(lStatus) > -1;
      var touched  = lLast && lLast >= mon && lLast <= sat;
      if (touched) leadWorked++;
      else if (isOpen) leadOpenUntouched++;
      if (lCreated && lCreated >= mon && lCreated <= sat) {
        leadsThisWeek++;
        var m24 = String(lRows[li][12]||'').trim();        // M 24hr Contact Done
        if (m24 === 'Yes') leads24++;
        else if (m24 === 'No') leadSlaMissed++;            // explicitly missed
        else if (lCreated < nowD) leadSlaMissed++;         // Pending/blank past the 24h window
      }
    }
  }
  var leadBase = leadWorked + leadOpenUntouched;

  // ── 4. Collection vs total outstanding (BILLING, cumulative) ──
  // Collection rate = total collected ÷ total billed (all time). Old unpaid
  // bills lower it — no free marks for a quiet week. Nothing billed → N/A.
  var totalBilled = 0, totalCollected = 0;
  var billSheet = s.getSheetByName(BILLING_TAB);
  if (billSheet && billSheet.getLastRow() > 1) {
    var bRows = billSheet.getDataRange().getValues();
    for (var bi = 1; bi < bRows.length; bi++) {
      totalBilled    += parseFloat(bRows[bi][4]) || 0;   // E Bill Amount
      totalCollected += parseFloat(bRows[bi][5]) || 0;   // F Amount Received
    }
  }

  // ── 5. Biometric from DAILY_SUMMARY (same as team) ──
  var sumSheet = s.getSheetByName(SUMMARY_TAB);
  var daysPresent = 0, lateCount = 0;
  var THR = '09:10';
  if (sumSheet && sumSheet.getLastRow() > 1) {
    var smRows = sumSheet.getDataRange().getValues();
    for (var mi = 1; mi < smRows.length; mi++) {
      var rDate = cellDate(smRows[mi][0]);
      var rName = String(smRows[mi][2]||'').trim().toLowerCase();
      if (rName !== whoLC) continue;
      if (!rDate || rDate < mon || rDate > sat) continue;
      daysPresent++;
      var arr = String(smRows[mi][1]||'').trim();
      if (arr && arr > THR) lateCount++;
    }
  }
  var absentDays = 6 - daysPresent;

  // ── 6. Scores ──
  function r1(n){ return Math.round(n*10)/10; }

  // Output components — each is "active" only if there's something to measure.
  // N/A components are excluded and the output is rescaled across active ones.
  var clientActive = totalOngoing > 0;
  var leadActive   = leadBase > 0;
  var revActive    = totalBilled > 0;

  var s_client = clientActive ? r1(connectedCount/totalOngoing*20) : 0;
  var s_lead   = leadActive   ? r1(leadWorked/leadBase*15) : 0;
  var collRatio= revActive    ? totalCollected/totalBilled : null;
  var s_rev    = revActive
                 ? r1(Math.min(collRatio, COLLECTION_TARGET)/COLLECTION_TARGET*15) : 0;

  // Output /50 = client-connection coverage of ONGOING projects only (lead-mgmt &
  // collection are shown for context but do not drive output this week — per spec).
  var output      = totalOngoing > 0 ? r1(connectedCount/totalOngoing*50) : 50;
  var missedProjects = perProject.filter(function(p){ return !p.connected; }).map(function(p){ return p.project; });

  var s_dpr    = Math.min(15, r1(dprDaysCount/6*15));
  var punctBase= Math.max(0, 15 - lateCount*2 - absentDays*2);
  var s_punct  = Math.min(15, r1(punctBase));          // /15 (base already on /15 scale)
  var hrsBase  = Math.max(0, r1(daysPresent/6*10 - absentDays*1.5));
  var s_hrs    = Math.min(10, r1(hrsBase));            // /10 (base already on /10 scale)

  // Reliability /10 — 24hr lead first-contact SLA: −1 per missed lead
  var reliabilityPenalty = leadSlaMissed * 1;
  var s_reliability = Math.max(0, r1(10 - reliabilityPenalty));

  var total    = r1(output + s_dpr + s_punct + s_hrs + s_reliability);

  return {
    member        : who,
    weekStart     : mon,
    weekEnd       : sat,
    totalOngoing  : totalOngoing,
    connectedCount: connectedCount,
    visitsDone    : visitsDone,
    meetingsDone  : meetingsDone,
    leadsThisWeek : leadsThisWeek,
    leads24       : leads24,
    leadSlaMissed : leadSlaMissed,
    leadWorked    : leadWorked,
    leadBase      : leadBase,
    collectionPct : (collRatio === null) ? null : Math.round(collRatio*100),
    collectionTarget: COLLECTION_TARGET*100,
    clientActive  : clientActive,
    leadActive    : leadActive,
    revActive     : revActive,
    dprDaysCount  : dprDaysCount,
    daysPresent   : daysPresent,
    lateCount     : lateCount,
    absentDays    : absentDays,
    reliabilityPenalty : reliabilityPenalty,
    perProject    : perProject,
    missedProjects: missedProjects,
    scores: {
      s_client : s_client,
      s_lead   : s_lead,
      s_rev    : s_rev,
      output   : output,
      s_dpr    : s_dpr,
      s_punct  : s_punct,
      s_hrs    : s_hrs,
      s_reliability : s_reliability,
      total    : total,
    },
  };
}

// ════════════════════════════════════════════════════════════════
// getWeeklyDiag — read-only dump of what the Deepak/Aman weekly functions see,
// for debugging a "showing 0" week. action=getWeeklyDiag&weekStart=YYYY-MM-DD
// ════════════════════════════════════════════════════════════════
function getWeeklyDiag(weekStart){
  var s = db();
  var mon = weekStart || dateStr(mondayOf(new Date()));
  var sat = addDaysToStr(mon, 5);
  var out = { week:[mon,sat], deepak:{}, aman:{} };

  // Deepak active projects — scan ALL columns of CONFIG for the header
  var cfg = s.getSheetByName(CONFIG_TAB), active = [], hdrAt = null;
  if (cfg){ var cr=cfg.getDataRange().getValues();
    found: for (var i=0;i<cr.length;i++){ for (var cj=0;cj<cr[i].length;cj++){
      if (String(cr[i][cj]||'').trim().toUpperCase()==='DEEPAK ACTIVE PROJECTS'){
        hdrAt = {row:i+1, col:cj+1};
        for (var rr=i+1;rr<cr.length;rr++){ var v=String(cr[rr][cj]||'').trim(); if(!v) break; active.push(v); }
        break found;
      } } } }
  out.deepak.headerFoundAt = hdrAt;
  out.deepak.activeProjects = active;
  var ex = s.getSheetByName(SITE_EXEC_TAB), erows = [];
  if (ex){ var er=ex.getDataRange().getValues();
    for (var j=1;j<er.length;j++){ var d=cellDate(er[j][1]); if(!d||d<mon||d>sat) continue;
      erows.push({date:d, project:String(er[j][3]||''), lead:String(er[j][4]||''), visit:String(er[j][5]||''), client:String(er[j][17]||'')}); } }
  out.deepak.execRowsThisWeek = erows;

  // Aman — show ALL CRM_LOG rows this week (raw member values), AMAN_DAILY rows, and project stages
  var lg = s.getSheetByName(CRM_LOG_TAB), conns = [];
  if (lg){ var lr=lg.getDataRange().getValues();
    for (var m=1;m<lr.length;m++){ var dt=cellDate(lr[m][2]); if(!dt||dt<mon||dt>sat) continue;
      conns.push({date:dt, member:String(lr[m][3]||''), category:String(lr[m][4]||''), project:String(lr[m][6]||'')}); } }
  out.aman.crmLogThisWeek = conns;
  var ad = s.getSheetByName(AMAN_DAILY_TAB), adr = [];
  if (ad){ var arr=ad.getDataRange().getValues();
    for (var a=1;a<arr.length;a++){ var dd=cellDate(arr[a][1]); if(!dd||dd<mon||dd>sat) continue;
      adr.push({date:dd, member:String(arr[a][3]||'')}); } }
  out.aman.amanDailyThisWeek = adr;
  var pj = s.getSheetByName(PROJECTS_TAB), proj = [];
  if (pj){ var pr=pj.getDataRange().getValues();
    for (var k=1;k<pr.length;k++){ var nm=String(pr[k][1]||'').trim(); if(!nm) continue;
      proj.push({name:nm, stage:String(pr[k][2]||'')}); } }
  out.aman.projects = proj;
  return out;
}

// ════════════════════════════════════════════════════════════════
// fixExistingSheetData — run once to clean up TASK_ASSIGNMENTS
//
// For every non-approved row:
//   1. Looks up correct Disc.Multiplier from PROJECTS tab (col E)
//   2. Looks up correct Stage Base Pts from CONFIG tab (col A→B)
//   3. Detects and fixes rows where F/G were swapped during migration
//   4. Recalculates WeightedPts = Disc.Mult × BasePts × Units
//   5. Sets Decision Pending tasks to 1 pt (so they appear on dashboard)
//   6. Fixes visit tasks: F=project mult, G=hours (keep), H=1,
//      I=hours×rate (2 for site visits, 1 for meetings)
//
// Approved tasks (LeadApproved=Yes) are NEVER touched — they are final.
// ════════════════════════════════════════════════════════════════
function fixExistingSheetData() {
  var s     = db();
  var sheet = s.getSheetByName(ASSIGN_TAB);
  if (!sheet) { Logger.log('TASK_ASSIGNMENTS not found'); return; }

  // ── Build CONFIG stage → basePts lookup ──────────────────────
  var configSheet = s.getSheetByName(CONFIG_TAB);
  var stagePtsMap = {}; // stage label → basePts
  if (configSheet) {
    var cRows = configSheet.getDataRange().getValues();
    var SKIP_A = ['STAGE WEIGHTS','TASK / STAGE TYPE','TOTAL',''];
    for (var ci = 0; ci < cRows.length; ci++) {
      var stg = String(cRows[ci][0]||'').trim();
      var pts = parseFloat(cRows[ci][1]);
      if (stg && SKIP_A.indexOf(stg.toUpperCase()) === -1 &&
          !stg.startsWith('—') && !isNaN(pts) && pts > 0) {
        stagePtsMap[stg] = pts;
      }
    }
  }
  Logger.log('Loaded ' + Object.keys(stagePtsMap).length + ' stage pts from CONFIG');

  // ── Build PROJECTS project → multiplier lookup ────────────────
  var projSheet  = s.getSheetByName(PROJECTS_TAB);
  var projMultMap = {}; // project name → multiplier
  if (projSheet) {
    var pRows = projSheet.getDataRange().getValues();
    for (var pi = 1; pi < pRows.length; pi++) {
      var pn = String(pRows[pi][1]||'').trim();
      var pm = parseFloat(pRows[pi][4]) || 1.0;
      if (pn) projMultMap[pn] = pm;
    }
  }
  Logger.log('Loaded ' + Object.keys(projMultMap).length + ' project multipliers');

  // ── Process every data row ────────────────────────────────────
  var headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var is23     = headers.length >= 23 || headers.indexOf('Actual Completion Date') > -1;
  var COL_LAPPR = is23 ? 16 : 15; // Q LeadApproved (0-based index)

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('No data rows'); return; }

  var data    = sheet.getRange(2, 1, lastRow-1, 23).getValues();
  var fixed   = 0, skipped = 0;

  var swapped = 0, corrected = 0, approvedSkipped = 0, unknownSkipped = 0;

  for (var i = 0; i < data.length; i++) {
    var row          = data[i];
    var project      = String(row[2]  || '').trim(); // C ProjectName
    var taskType     = String(row[4]  || '').trim(); // E Stage
    var colF         = parseFloat(row[5]);           // F Disc.Mult (current)
    var colG         = parseFloat(row[6]);           // G BasePts (current)
    var colH         = parseFloat(row[7]) || 1;      // H Units
    var colI         = parseFloat(row[8]) || 0;      // I WeightedPts
    if (isNaN(colF)) colF = 0;
    if (isNaN(colG)) colG = 0;
    var leadApproved = String(row[COL_LAPPR]||'').trim();

    // ── RULE: Never touch approved tasks ────────────────────────
    // Lead may have manually corrected pts during approval — those are final
    if (leadApproved === 'Yes') { approvedSkipped++; continue; }

    var isVisit    = isVisitTask(taskType);
    var isDecision = (taskType === 'Decision Pending');

    // ── Look up authoritative values ─────────────────────────────
    var multLookup = projMultMap[project];   // undefined if project not in PROJECTS
    var baseLookup = stagePtsMap[taskType];  // undefined if stage not in CONFIG

    var multFound = (multLookup !== undefined);
    var baseFound = (baseLookup !== undefined);

    // Conservative defaults: if lookup fails, keep existing value
    var correctMult = multFound ? multLookup : colF;
    var correctBase = baseFound ? baseLookup : colG;

    var newF = colF, newG = colG, newH = colH, newI = colI;
    var action = '';

    // ── Decision Pending: always 1 pt ───────────────────────────
    if (isDecision) {
      newF = 1; newG = 1; newH = 1; newI = 1;
      action = 'decision-default';

    // ── Visit / Meeting tasks ─────────────────────────────────────
    } else if (isVisit) {
      if (multFound) newF = correctMult; // verified multiplier
      newH = 1; // visits always 1 unit
      // G = hours stored by updateTaskStatusesFromDPR — keep it
      // Only recalculate WeightedPts if we have hours (G > 0)
      if (colG > 0) {
        newI = calcVisitPts(taskType, colG); // hours × rate (2 or 1)
      }
      action = 'visit';

    // ── Regular tasks — both lookups succeeded ────────────────────
    } else if (multFound && baseFound) {
      // Detect F/G swap: F has basePts value, G has multiplier value
      var fMatchesBase = Math.abs(colF - correctBase) < 0.01;
      var gMatchesMult = Math.abs(colG - correctMult) < 0.01;

      if (fMatchesBase && gMatchesMult && Math.abs(correctBase - correctMult) > 0.01) {
        // Clear swap detected (values are numerically different so unambiguous)
        newF = correctMult;
        newG = correctBase;
        action = 'swap-fixed';
        Logger.log('SWAP row ' + (i+2) + ': ' + project + ' | ' + taskType +
          ' | F:' + colF + '→' + newF + ' G:' + colG + '→' + newG);
      } else {
        // Verify against lookup — correct any mismatch
        newF = correctMult;
        newG = correctBase;
        action = 'verified';
      }
      // Recalculate WeightedPts from authoritative F × G × H
      newI = Math.round(newF * newG * colH * 10) / 10;

    // ── Regular tasks — only multiplier lookup succeeded ──────────
    } else if (multFound && !baseFound) {
      // Can verify multiplier but not basePts — fix F only, keep G
      newF = correctMult;
      // Recalculate I using fixed F and existing G
      if (colG > 0) newI = Math.round(newF * colG * colH * 10) / 10;
      action = 'mult-only';

    // ── Regular tasks — only basePts lookup succeeded ─────────────
    } else if (!multFound && baseFound) {
      // Can verify basePts but not multiplier — fix G only, keep F
      newG = correctBase;
      if (colF > 0) newI = Math.round(colF * newG * colH * 10) / 10;
      action = 'base-only';

    // ── Neither lookup succeeded — leave completely unchanged ─────
    } else {
      unknownSkipped++;
      continue;
    }

    // Write only if something actually changed
    var changed = (Math.abs(newF - colF) > 0.001) ||
                  (Math.abs(newG - colG) > 0.001) ||
                  (Math.abs(newH - colH) > 0.001) ||
                  (Math.abs(newI - colI) > 0.001);

    if (changed) {
      sheet.getRange(i+2, 6, 1, 4).setValues([[newF, newG, newH, newI]]);
      if (action === 'swap-fixed') swapped++; else corrected++;
      if ((swapped + corrected) % 50 === 0) SpreadsheetApp.flush();
    } else {
      skipped++; // Already correct — no write needed
    }
  }

  SpreadsheetApp.flush();
  Logger.log('=== fixExistingSheetData complete ===');
  Logger.log('Swaps fixed: ' + swapped);
  Logger.log('Values corrected: ' + corrected);
  Logger.log('Already correct (no change): ' + skipped);
  Logger.log('Approved (skipped): ' + approvedSkipped);
  Logger.log('Unknown stage/project (skipped): ' + unknownSkipped);
  return {status:'ok', swapped:swapped, corrected:corrected,
          alreadyCorrect:skipped, approvedSkipped:approvedSkipped,
          unknownSkipped:unknownSkipped};
}

// ════════════════════════════════════════════════════════════════
// migrateToNewColumnStructure — one-time migration
//
// Converts TASK_ASSIGNMENTS from old structure to new:
//   OLD: F=DisciplineText  G=BasePts  H=Multiplier  I=WeightedPts
//   NEW: F=DiscMultiplier  G=BasePts  H=Units        I=WeightedPts
//
// Logic per row:
//   F(new) = H(old) — multiplier already stored in H, move to F
//   G(new) = G(old) — basePts unchanged
//   H(new) = derived units = round(WeightedPts / (BasePts × Mult))
//             visit tasks → units = 1 always
//             bad division → units = 1 default
//   I(new) = I(old) — keep approved WeightedPts as-is (source of truth)
//
// Also updates header row and applies correct number formatting.
// ════════════════════════════════════════════════════════════════
function migrateToNewColumnStructure() {
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) { Logger.log('TASK_ASSIGNMENTS not found'); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('No data rows'); return; }

  Logger.log('=== migrateToNewColumnStructure: ' + (lastRow-1) + ' rows ===');

  var data      = sheet.getRange(2, 1, lastRow-1, 23).getValues();
  var migrated  = 0;

  for (var i = 0; i < data.length; i++) {
    var row         = data[i];
    var taskType    = String(row[4]||'').trim();   // E
    var discOrMult  = row[5];                      // F — currently discipline text OR already a number
    var basePts     = parseFloat(row[6]) || 0;     // G
    var oldH        = row[7];                      // H — currently multiplier number
    var weightedPts = parseFloat(row[8]) || 0;     // I

    // Determine the multiplier value:
    // If col F already holds a number (e.g. from a partial migration), use it.
    // Otherwise col F holds discipline text → use col H for the multiplier.
    var mult;
    var fIsNumber = (typeof discOrMult === 'number') ||
                    (!isNaN(parseFloat(discOrMult)) && String(discOrMult).trim() !== '');
    if (fIsNumber) {
      mult = parseFloat(discOrMult) || 1;
    } else {
      mult = parseFloat(oldH) || 1;
    }

    // Derive units
    var units;
    if (isVisitTask(taskType)) {
      units = 1; // visits: always 1 unit, pts based on hours
    } else if (basePts > 0 && mult > 0) {
      var derived = weightedPts / (basePts * mult);
      units = Math.round(derived);
      if (units < 1 || isNaN(units)) units = 1;
    } else {
      units = 1;
    }

    // Write: F=mult, G=basePts(unchanged), H=units, I=weightedPts(unchanged)
    sheet.getRange(i+2, 6, 1, 4).setValues([[mult, basePts, units, weightedPts]]);
    migrated++;

    if (migrated % 50 === 0) SpreadsheetApp.flush();
  }

  SpreadsheetApp.flush();

  // Update header row labels
  sheet.getRange(1, 6).setValue('Disc. Multiplier');
  sheet.getRange(1, 8).setValue('Units');

  // Apply correct number formatting
  var dataRows = lastRow - 1;
  sheet.getRange(2, 6, dataRows, 1).setNumberFormat('0.0');   // F Multiplier
  sheet.getRange(2, 7, dataRows, 1).setNumberFormat('0.00');  // G BasePts
  sheet.getRange(2, 8, dataRows, 1).setNumberFormat('0');     // H Units (integer)
  sheet.getRange(2, 9, dataRows, 1).setNumberFormat('0.00');  // I WeightedPts

  Logger.log('=== Migration done: ' + migrated + ' rows updated ===');
  return {status:'ok', migrated: migrated};
}

// ════════════════════════════════════════════════════════════════
// fixShiftedTaskColumns — one-time repair for column shift bug
//
// BEFORE RUNNING THIS:
//   1. Open TASK_ASSIGNMENTS in Google Sheets
//   2. Find the "Deadline" column (currently showing right after Stage)
//   3. Move it back to after the "Assigned Date" column
//      (right-click column header → Move column right, repeat until
//       it sits between Assigned Date and Area)
//   4. Then run this function from Apps Script editor
//
// WHAT IT DOES:
//   Detects rows where col K (Deadline) contains a discipline string
//   instead of a date — these rows had their values shifted left by
//   one position when the Deadline column was moved. Rotates them back.
//
//   Affected rows: [F,G,H,I,J,K] = [basePts,mult,weighted,assignedDate,deadline,disc]
//   Fixed to:      [F,G,H,I,J,K] = [disc,basePts,mult,weighted,assignedDate,deadline]
// ════════════════════════════════════════════════════════════════
function fixShiftedTaskColumns() {
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) { Logger.log('TASK_ASSIGNMENTS not found'); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('No data rows found'); return; }

  Logger.log('=== fixShiftedTaskColumns: scanning ' + (lastRow-1) + ' rows ===');

  // Read all rows at once for speed
  var data = sheet.getRange(2, 1, lastRow - 1, 23).getValues();
  var fixed = 0, skipped = 0;

  for (var i = 0; i < data.length; i++) {
    var row   = data[i];
    var colK  = row[10]; // K = Deadline (0-based index 10)

    // Detect corrupted row: col K should hold a date after move-back.
    // If it holds a non-empty string that doesn't look like a date → corrupted.
    var isDate  = (colK instanceof Date) ||
                  (typeof colK === 'string' && /^\d{4}-\d{2}-\d{2}/.test(colK)) ||
                  (typeof colK === 'number' && colK > 10000); // Sheets date serial
    var isEmpty = (colK === null || colK === undefined || String(colK).trim() === '');

    if (isDate || isEmpty) { skipped++; continue; }

    // col K is a text string (discipline) → this row is corrupted
    // Current: [F=basePts, G=mult, H=weighted, I=assignedDate, J=deadline, K=disc]
    // Correct: [F=disc,    G=basePts, H=mult,  I=weighted,  J=assignedDate, K=deadline]
    var f = row[5], g = row[6], h = row[7], ii = row[8], j = row[9], k = row[10];

    sheet.getRange(i + 2, 6, 1, 6).setValues([[k, f, g, h, ii, j]]);

    Logger.log('Fixed row ' + (i+2) +
      ' | proj=' + row[2] +
      ' | assignee=' + row[3] +
      ' | disc=' + k +
      ' | deadline=' + j);
    fixed++;

    // Flush every 50 writes to avoid timeout
    if (fixed % 50 === 0) SpreadsheetApp.flush();
  }

  SpreadsheetApp.flush();
  Logger.log('=== Done: ' + fixed + ' rows fixed, ' + skipped + ' skipped ===');
  return {status: 'ok', fixed: fixed, skipped: skipped};
}

// ════════════════════════════════════════════════════════════════
// fixColumnFormattingAfterShift — run AFTER fixShiftedTaskColumns()
// Restores correct number / date formats to TASK_ASSIGNMENTS columns
// that inherited wrong formatting during the Deadline column move.
//
// Fixes:
//   Col F  (Discipline)     → plain text / no format
//   Col G  (Stage Base Pts) → number  "0.00"
//   Col H  (Disc.Multiplier)→ number  "0.0"
//   Col I  (Weighted Pts)   → number  "0.00"  ← was date-formatted
//   Col J  (Assigned Date)  → date    "yyyy-mm-dd"
//   Col K  (Deadline)       → date    "yyyy-mm-dd"  ← was DD/MM/YYYY
// ════════════════════════════════════════════════════════════════
function fixColumnFormattingAfterShift() {
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) { Logger.log('TASK_ASSIGNMENTS not found'); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('No data rows'); return; }

  var dataRows = lastRow - 1; // number of data rows (exclude header)

  // Col F (6) — Discipline: plain text, no format that could distort the string
  sheet.getRange(2, 6, dataRows, 1).setNumberFormat('@STRING@');

  // Col G (7) — Stage Base Pts: number with 2 dp
  sheet.getRange(2, 7, dataRows, 1).setNumberFormat('0.00');

  // Col H (8) — Disc. Multiplier: number with 1 dp
  sheet.getRange(2, 8, dataRows, 1).setNumberFormat('0.0');

  // Col I (9) — Weighted Points: number with 2 dp
  // This is the main problem column — was showing 1.5 as "1899-12-31" etc.
  sheet.getRange(2, 9, dataRows, 1).setNumberFormat('0.00');

  // Col J (10) — Assigned Date: ISO date format
  sheet.getRange(2, 10, dataRows, 1).setNumberFormat('yyyy-mm-dd');

  // Col K (11) — Deadline: ISO date format (was showing as DD/MM/YYYY)
  sheet.getRange(2, 11, dataRows, 1).setNumberFormat('yyyy-mm-dd');

  SpreadsheetApp.flush();
  Logger.log('Column formatting fixed for cols F–K (rows 2–' + lastRow + ')');
  return {status: 'ok', rows: dataRows};
}

// Run this manually to verify sheet access
function testSheetAccess() {
  try {
    var s = db();
    Logger.log('Sheet name: ' + s.getName());
    Logger.log('Sheet tabs: ' + s.getSheets().map(function(sh){return sh.getName();}).join(', '));
    var team = s.getSheetByName(TEAM_TAB);
    Logger.log('TEAM tab rows: ' + (team ? team.getLastRow() : 'NOT FOUND'));
    var proj = s.getSheetByName(PROJECTS_TAB);
    Logger.log('PROJECTS tab rows: ' + (proj ? proj.getLastRow() : 'NOT FOUND'));
  } catch(e) {
    Logger.log('ERROR: ' + e.toString());
  }
}

function testLists() {
  var result = getLists();
  Logger.log('TEAM count: ' + result.team.length);
  Logger.log('PROJECTS count: ' + result.projects.length);
  Logger.log('First 3 team: ' + result.team.slice(0,3).join(', '));
  Logger.log('First 3 projects: ' + result.projects.slice(0,3).map(function(p){return p.name;}).join(', '));
}

// ════════════════════════════════════════════════════════════════
// AMAN CRM — Sheet tabs + header writers
// ════════════════════════════════════════════════════════════════
var AMAN_DAILY_TAB = 'AMAN_DAILY';
var LEADS_TAB      = 'LEADS';
var FEEDBACK_TAB   = 'FEEDBACK';
var BILLING_TAB    = 'BILLING';
var CRM_LOG_TAB    = 'CRM_LOG';   // client connections + other activities (one row each)

function writeAmanDailyHeaders(sheet) {
  // Daily index/summary only. Detail rows now live in dedicated tabs:
  //   Client connections → CLIENT_CONNECTIONS (one row each)
  //   Other activities   → ACTIVITIES (one row each)
  //   Leads→LEADS, Finance→BILLING, Issues→SITE_ISSUES, Feedback→FEEDBACK,
  //   Agendas→TASK_ASSIGNMENTS.
  var h = [
    'Submission ID','Date','Time','Member',
    'Vendor Coordination','Site Issues Addressed','TnCP Coordination','BNI Activity',
  ];
  sheet.getRange(1,1,1,h.length).setValues([h])
    .setBackground('#005F73').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
}

function writeCrmLogHeaders(sheet) {
  // Combined client connections + other activities, one row per entry.
  var h = ['Log ID','Submission ID','Date','Member','Category','Type / Activity','Project','Notes / Discussion'];
  sheet.getRange(1,1,1,h.length).setValues([h])
    .setBackground('#005F73').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  var widths=[120,130,100,150,150,180,200,340];
  widths.forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
}

function writeBillingHeaders(sheet) {
  var h = [
    'Bill ID','Invoice No.','Project','Bill Date','Bill Amount (Rs)',
    'Amount Received (Rs)','Received Date','Last Follow-up Date',
    'Status','Submission ID',
  ];
  sheet.getRange(1,1,1,h.length).setValues([h])
    .setBackground('#7A5000').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  var widths=[120,130,200,110,140,150,130,150,110,130];
  widths.forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
}

function writeFeedbackHeaders(sheet) {
  // Mirrors the Monthly Feedback Google Form
  var h = [
    'Feedback ID','Submission ID','Date Recorded','Recorded By',
    'Project / Client','Date of Visit/Meeting',
    'Overall Satisfaction (1-10)','Agenda Communicated',
    'Design & Functionality (1-5)','Communication (1-5)',
    'Problem Resolution (1-5)','Responsiveness (1-5)',
    'Quality of Work (1-5)','Professionalism (1-5)',
    'Meeting On Time','Recommend / NPS (1-10)',
    'Additional Comments','Appraisal of Person','Referrals',
  ];
  sheet.getRange(1,1,1,h.length).setValues([h])
    .setBackground('#6A1B9A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  var widths=[120,130,120,150,200,140,110,120,130,120,130,120,120,120,110,130,300,220,260];
  widths.forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
}

var LEADS_HEADERS = [
  'Lead ID','Client Name','Contact No.','Referred By',
  'Validation Check','Lead Status','Contacted By',
  'Lead Creation Date','Last Contacted','Lead Manager',
  'Remarks','Lost Reasons','24hr Contact Done',
  'Lead Source','Briefing Date','Proposal Date',
];
function writeLeadsHeaders(sheet) {
  // Matches LMS structure from IDS Google Form (+ source & SLA dates)
  var h = LEADS_HEADERS;
  sheet.getRange(1,1,1,h.length).setValues([h])
    .setBackground('#1565C0').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  var widths=[120,200,140,200,120,200,160,130,130,160,300,200,130,150,120,120];
  widths.forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
}

// Non-destructive: adds the new N/O/P headers to an existing LEADS tab.
function migrateLeadsColumns(sheet) {
  if (!sheet) return;
  var lastCol = sheet.getLastColumn();
  if (lastCol >= LEADS_HEADERS.length) return;  // already migrated
  var have = lastCol > 0
    ? sheet.getRange(1,1,1,lastCol).getValues()[0].map(function(x){ return String(x||'').trim(); })
    : [];
  for (var c = 0; c < LEADS_HEADERS.length; c++) {
    if (have[c] !== LEADS_HEADERS[c]) {
      sheet.getRange(1, c+1).setValue(LEADS_HEADERS[c])
        .setBackground('#1565C0').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
    }
  }
}

// ════════════════════════════════════════════════════════════════
// getLeadsAnalytics — funnel, conversion, sources, lost reasons & SLA
// action=getLeadsAnalytics&month=2026-06  (month optional → current)
// Goal: never miss a lead (24h/48h SLA radar) + raise conversion rate.
// ════════════════════════════════════════════════════════════════
var MONTH_NAMES = ['January','February','March','April','May','June',
                   'July','August','September','October','November','December'];
function diffDays(a, b){ // a,b = 'YYYY-MM-DD' strings
  if (!a || !b) return null;
  return Math.round((new Date(b).getTime() - new Date(a).getTime()) / 86400000);
}
function getLeadsAnalytics(month) {
  var sheet = db().getSheetByName(LEADS_TAB);
  if (!sheet) return {error:'LEADS tab not found'};
  migrateLeadsColumns(sheet);

  var today = dateStr();
  var curMonth = month && /^\d{4}-\d{2}$/.test(month) ? month : today.substring(0,7);
  var yy = parseInt(curMonth.substring(0,4),10), mm = parseInt(curMonth.substring(5,7),10);
  var monthLabel = MONTH_NAMES[mm-1] + ' ' + yy;

  var rows = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [[]];

  // Accumulators
  var monthLeads = [];      // leads created in the selected month
  var bySource = {}, byManager = {}, lostByReason = {};
  var funnel = {contacted:0, briefing:0, proposal:0, converted:0, lost:0, invalid:0};
  var sla24 = {met:0, missed:0, pending:0};
  var sla48 = {met:0, late:0, breaching:0, pending:0};
  var needsAttention = [];
  var trend = {}; // 'YYYY-MM' → {created, converted}

  // Last 6 months keys
  var trendKeys = [];
  for (var k = 5; k >= 0; k--) {
    var d = new Date(yy, mm-1-k, 1);
    var key = d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2);
    trendKeys.push(key);
    trend[key] = {label: MONTH_NAMES[d.getMonth()].substring(0,3)+' '+String(d.getFullYear()).substring(2),
                  month:key, created:0, converted:0};
  }

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!String(r[0]||'').trim() && !String(r[1]||'').trim()) continue;
    var name    = String(r[1]||'').trim();
    var status  = String(r[5]||'').trim();
    var manager = String(r[9]||'').trim() || 'Unassigned';
    var lostRsn = String(r[11]||'').trim() || 'Unspecified';
    var c24     = String(r[12]||'').trim();
    var source  = String(r[13]||'').trim() || 'Unspecified';
    var created = cellDate(r[7]);
    var briefing= cellDate(r[14]);
    var proposal= cellDate(r[15]);
    var isInvalid = status === 'Invalid' || String(r[4]||'').trim() === 'Invalid';

    // 6-month trend
    if (created) {
      var cm = created.substring(0,7);
      if (trend[cm]) { trend[cm].created++; if (status === 'Lead Converted') trend[cm].converted++; }
    }

    // ── Needs-attention radar (ALL open leads, any month) ──
    var ageC = created ? diffDays(created, today) : null;
    if (!isInvalid && status !== 'Lost' && status !== 'Lead Converted') {
      if (status === 'Not Contacted' && c24 !== 'Yes' && created && created < today) {
        needsAttention.push({leadId:String(r[0]||''), name:name, manager:manager,
          issue:'No first contact', ageDays:ageC, status:status, sla:'24h'});
      } else if (briefing && !proposal) {
        var ageB = diffDays(briefing, today);
        if (ageB !== null && ageB > 2) needsAttention.push({leadId:String(r[0]||''), name:name,
          manager:manager, issue:'Briefing '+ageB+'d ago, no proposal', ageDays:ageB, status:status, sla:'48h'});
      }
    }

    // ── Per-month metrics ──
    if (created && created.substring(0,7) === curMonth) {
      monthLeads.push(r);
      bySource[source]   = (bySource[source]||0)+1;
      byManager[manager] = (byManager[manager]||0)+1;

      if (isInvalid) funnel.invalid++;
      else {
        if (status === 'Lost') { funnel.lost++; lostByReason[lostRsn] = (lostByReason[lostRsn]||0)+1; }
        if (['Contacted over call','Briefing Meeting Done','Design Proposal Shared','Fee Proposal Shared','Lead Converted'].indexOf(status) > -1) funnel.contacted++;
        if (leadReachedBriefing(status) || briefing) funnel.briefing++;
        if (leadReachedProposal(status) || proposal) funnel.proposal++;
        if (status === 'Lead Converted') funnel.converted++;
      }

      // 24h SLA
      if (c24 === 'Yes') sla24.met++;
      else if (c24 === 'No') sla24.missed++;
      else if (created < today) sla24.missed++;   // Pending past the window
      else sla24.pending++;

      // 48h proposal SLA (only leads that had a briefing)
      if (briefing) {
        if (proposal) { (diffDays(briefing, proposal) <= 2) ? sla48.met++ : sla48.late++; }
        else if (diffDays(briefing, today) > 2) sla48.breaching++;
        else sla48.pending++;
      }
    }
  }

  var totalMonth = monthLeads.length;
  var validMonth = totalMonth - funnel.invalid;
  var conversionRate = validMonth > 0 ? Math.round(funnel.converted / validMonth * 1000)/10 : 0;

  needsAttention.sort(function(a,b){ return (b.ageDays||0) - (a.ageDays||0); });
  var trendArr = trendKeys.map(function(key){
    var t = trend[key];
    t.rate = t.created > 0 ? Math.round(t.converted / t.created * 1000)/10 : 0;
    return t;
  });

  function toArr(obj){ return Object.keys(obj).map(function(k){ return {label:k, count:obj[k]}; })
                              .sort(function(a,b){ return b.count - a.count; }); }

  return {
    month: curMonth, monthLabel: monthLabel,
    totalLeads: totalMonth, validLeads: validMonth,
    conversionRate: conversionRate,
    funnel: funnel,
    bySource: toArr(bySource),
    byManager: toArr(byManager),
    lostByReason: toArr(lostByReason),
    sla24: sla24, sla48: sla48,
    needsAttention: needsAttention,
    needsAttentionCount: needsAttention.length,
    trend: trendArr,
  };
}

// ════════════════════════════════════════════════════════════════
// getIssuesByReporter — used by both Deepak DPER and Aman CRM forms
// Reads SITE_ISSUES col O (index 14) for Reported By
// ════════════════════════════════════════════════════════════════
function getIssuesByReporter(member, useFallback) {
  var sheet = db().getSheetByName(SITE_ISSUES_TAB);
  if (!sheet) return {issues:[]};
  var rows = sheet.getDataRange().getValues();
  var issues = [];
  var sevenAgo = addDaysToStr(dateStr(), -7);

  for (var i = 1; i < rows.length; i++) {
    var r          = rows[i];
    var reportedBy = String(r[14] || '').trim();  // col O (0-indexed 14)
    var assignedTo = String(r[7] || '').trim();

    // Strict match on Reported By. The assignedTo fallback is ONLY used for
    // Deepak's legacy DPER rows that predate the "Reported By" column (blank col O).
    // CRM (Aman) passes useFallback=false so it never picks up DPER-sourced issues.
    var isMatch    = (reportedBy === member) ||
                     (useFallback && reportedBy === '' && assignedTo === member);
    if (member && !isMatch) continue;

    var issStatus  = String(r[10] || '').trim();
    var resolvedDt = String(r[13] || '').substring(0, 10);
    // Blocked / Void / Invalid / Cancelled issues are erased from the open list
    if (['Blocked','Void','Invalid','Cancelled'].indexOf(issStatus) > -1) continue;
    if (issStatus === 'Resolved' && resolvedDt < sevenAgo) continue;

    issues.push({
      issueId    : String(r[0]  || ''),
      date       : String(r[2]  || '').substring(0, 10),
      project    : String(r[3]  || ''),
      issueType  : String(r[5]  || ''),
      description: String(r[6]  || ''),
      assignedTo : assignedTo,
      priority   : r[8]  || 2,
      targetDate : String(r[9]  || ''),
      status     : issStatus,
      reportedBy : reportedBy,
    });
  }
  return {issues: issues};
}

// ════════════════════════════════════════════════════════════════
// updateIssueStatus — called from DPER & CRM forms for issue actions
// ════════════════════════════════════════════════════════════════
function updateIssueStatus(issueId, status, targetDate) {
  var sheet = db().getSheetByName(SITE_ISSUES_TAB);
  if (!sheet) return {status:'error', message:'SITE_ISSUES not found'};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() !== issueId) continue;
    sheet.getRange(i+1, 11).setValue(status);                          // K Status
    if (targetDate) sheet.getRange(i+1, 10).setValue(targetDate);      // J Target Date
    if (status === 'Resolved') sheet.getRange(i+1, 14).setValue(dateStr()); // N Resolved Date
    Logger.log('Issue updated: ' + issueId + ' → ' + status);
    return {status:'ok', issueId:issueId};
  }
  return {status:'error', message:'Issue not found: ' + issueId};
}

// ════════════════════════════════════════════════════════════════
// getRecentLeads — returns leads added on a given date with
// 24hr contact still pending (col M = 'Pending')
// ════════════════════════════════════════════════════════════════
function getRecentLeads(date) {
  var targetDate = date || addDaysToStr(dateStr(), -1);
  var sheet = db().getSheetByName(LEADS_TAB);
  if (!sheet || sheet.getLastRow() < 2) return {leads:[]};
  var rows = sheet.getDataRange().getValues();
  var leads = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var rowDate   = String(r[7]  || '').substring(0, 10); // H Lead Creation Date
    var contact24 = String(r[12] || '').trim();            // M 24hr Contact Done
    if (rowDate !== targetDate) continue;
    if (contact24 !== 'Pending') continue;
    leads.push({
      leadId     : String(r[0] || ''),
      clientName : String(r[1] || ''),
      contactNo  : String(r[2] || ''),
      referredBy : String(r[3] || ''),
      leadStatus : String(r[5] || ''),
      leadManager: String(r[9] || ''),
      date       : rowDate,
      rowNum     : i + 1,
    });
  }
  return {leads: leads};
}

// ════════════════════════════════════════════════════════════════
// getOpenLeads — leads still needing a status decision (early stages).
// Shown in the CRM form's "Open Leads" panel every day until the lead is
// advanced to Briefing Meeting Done (→ promoted to a project) or marked
// Lost/Invalid. Lets the CRM mark: Not Contacted / Contacted / Briefing /
// Rejected on any open lead, not just yesterday's.
// ════════════════════════════════════════════════════════════════
function getOpenLeads(member) {
  var sheet = db().getSheetByName(LEADS_TAB);
  if (!sheet || sheet.getLastRow() < 2) return {leads:[]};
  var rows = sheet.getDataRange().getValues();
  var OPEN = ['', 'Not Contacted', 'Contacted over call'];
  var leads = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var status = String(r[5] || '').trim();   // F Lead Status
    if (OPEN.indexOf(status) === -1) continue;
    leads.push({
      leadId     : String(r[0] || ''),
      clientName : String(r[1] || ''),
      contactNo  : String(r[2] || ''),
      referredBy : String(r[3] || ''),
      leadStatus : status || 'Not Contacted',
      leadManager: String(r[9] || ''),
      date       : String(r[7] || '').substring(0, 10),  // H Lead Creation Date
      rowNum     : i + 1,
    });
  }
  return {leads: leads};
}

// ════════════════════════════════════════════════════════════════
// submitAmanCRM — main CRM daily form submission handler
// Writes: AMAN_DAILY row, LEADS rows (new), LEADS updates (followups),
//         SITE_ISSUES rows (new issues)
// ════════════════════════════════════════════════════════════════
function submitAmanCRM(data) {
  try {
    var today   = data.date || dateStr();
    var member  = data.member || 'Aman Raghuwanshi';
    var now     = data.time  || new Date().toTimeString().substring(0, 5);

    // ── Parse JSON fields ──────────────────────────────────────
    var newLeads    = [];
    var leadFollows = [];
    var contacts    = [];
    var agendas     = [];
    var finance     = {};
    var activities  = {};
    var newIssues   = [];
    var feedback    = [];
    var newProjects = [];
    try { newLeads    = JSON.parse(data.newLeads     || '[]'); } catch(e){}
    try { leadFollows = JSON.parse(data.leadFollowups|| '[]'); } catch(e){}
    try { contacts    = JSON.parse(data.contacts     || '[]'); } catch(e){}
    try { agendas     = JSON.parse(data.agendas      || '[]'); } catch(e){}
    try { finance     = JSON.parse(data.finance      || '{}'); } catch(e){}
    try { activities  = JSON.parse(data.activities    || '{}'); } catch(e){}
    try { newIssues   = JSON.parse(data.newIssues    || '[]'); } catch(e){}
    try { feedback    = JSON.parse(data.feedback     || '[]'); } catch(e){}
    try { newProjects = JSON.parse(data.newProjects  || '[]'); } catch(e){}

    // ── 0. Create any new projects CRM added (PROJECTS tab) ───
    if (newProjects.length > 0) createCrmProjects(newProjects);

    // Tag each issue with the reporter so SITE_ISSUES col O is correct
    newIssues.forEach(function(iss){ iss.reportedBy = iss.reportedBy || member; });

    var vendor    = activities.vendor     || {on:'No', entries:[]};
    var siteIss   = activities.siteIssues || {on:'No', entries:[]};
    var tncp      = activities.tncp       || {on:'No', entries:[]};
    var bni       = activities.bni        || {on:'No', entries:[]};

    // ── 1. Write AMAN_DAILY index row (summary only) ──────────
    var dailySheet = getOrCreate(AMAN_DAILY_TAB, writeAmanDailyHeaders);
    var subId      = nextId(dailySheet, 'CRM-');
    prependRow(dailySheet, [
      subId, today, now, member,
      vendor.on || 'No', siteIss.on || 'No', tncp.on || 'No', bni.on || 'No',
    ]);
    Logger.log('AMAN_DAILY written: ' + subId);

    // ── 1a. Client connections + activities → CRM_LOG (one row each) ──
    var logSheet = getOrCreate(CRM_LOG_TAB, writeCrmLogHeaders);
    contacts.forEach(function(c) {
      if (!c.project && !c.type && !c.notes) return;
      logSheet.appendRow([
        nextId(logSheet, 'LOG-'), subId, today, member,
        'Client Connection', c.type || '', c.project || '', c.notes || '',
      ]);
    });
    function writeActRows(label, act) {
      (act.entries || []).forEach(function(e) {
        if (!e.project && !e.notes) return;
        prependRow(logSheet, [ nextId(logSheet, 'LOG-'), subId, today, member, 'Other Activity', label, e.project || '', e.notes || '' ]);
      });
    }
    if (vendor.on  === 'Yes') writeActRows('Vendor Coordination', vendor);
    if (siteIss.on === 'Yes') writeActRows('Site Issues Addressed', siteIss);
    if (tncp.on    === 'Yes') writeActRows('TnCP Coordination', tncp);
    if (bni.on     === 'Yes') prependRow(logSheet, [ nextId(logSheet, 'LOG-'), subId, today, member, 'Other Activity', 'BNI Activity', '', '' ]);

    // ── 2. Write new leads to LEADS sheet ─────────────────────
    var leadsSheet = getOrCreate(LEADS_TAB, writeLeadsHeaders);
    migrateLeadsColumns(leadsSheet);
    newLeads.forEach(function(lead) {
      if (!lead.clientName) return;
      var leadId = nextId(leadsSheet, 'LEAD-');
      var st = lead.leadStatus || 'Not Contacted';
      prependRow(leadsSheet, [
        leadId,
        lead.clientName  || '',
        lead.contactNo   || '',
        lead.referredBy  || '',
        lead.validation  || 'Not checked',
        st,
        member,                    // Contacted By = Aman
        today,                     // Lead Creation Date
        st === 'Not Contacted' ? '' : today,  // Last Contacted
        lead.leadManager || '',
        lead.remarks     || '',
        lead.lostReasons || '',
        st === 'Not Contacted' ? 'Pending' : 'Yes',  // 24hr Contact Done
        lead.leadSource  || '',                       // N Lead Source
        leadReachedBriefing(st) ? today : '',         // O Briefing Date
        leadReachedProposal(st) ? today : '',         // P Proposal Date
      ]);
      // Promote serious leads (briefing done or further) to PROJECTS tab
      if (isPromoteLeadStage(st)) promoteLeadToProject(lead.clientName, member);
    });
    Logger.log('Leads written: ' + newLeads.length);

    // ── 3. Update 24hr follow-up status on existing leads ─────
    if (leadFollows.length > 0) {
      var leadsData = leadsSheet.getDataRange().getValues();
      leadFollows.forEach(function(f) {
        if (!f.leadId) return;
        for (var i = 1; i < leadsData.length; i++) {
          if (String(leadsData[i][0] || '') !== f.leadId) continue;
          leadsSheet.getRange(i+1, 13).setValue(f.contacted ? 'Yes' : 'No'); // M 24hr Contact Done
          if (f.contacted) {
            leadsSheet.getRange(i+1, 6).setValue(f.newStatus || '');   // F Lead Status
            leadsSheet.getRange(i+1, 9).setValue(today);               // I Last Contacted
            if (f.newStatus === 'Lost' && f.lostReasons)
              leadsSheet.getRange(i+1, 12).setValue(f.lostReasons);    // L Lost Reasons
            // Stamp SLA dates when the stage is reached (only if not already set)
            if (leadReachedBriefing(f.newStatus) && !String(leadsData[i][14]||'').trim())
              leadsSheet.getRange(i+1, 15).setValue(today);            // O Briefing Date
            if (leadReachedProposal(f.newStatus) && !String(leadsData[i][15]||'').trim())
              leadsSheet.getRange(i+1, 16).setValue(today);            // P Proposal Date
            // Promote to PROJECTS once it reaches briefing-done or further
            if (isPromoteLeadStage(f.newStatus)) promoteLeadToProject(String(leadsData[i][1]||''), member);
          }
          break;
        }
      });
    }

    // ── 4. New issues vs design deliverables ──────────────────
    // Issues → SITE_ISSUES (show in CRM open-issues panel).
    // Deliverables → TASK_ASSIGNMENTS only (NOT the open-issues panel).
    if (newIssues.length > 0) {
      var realIssues = [], deliverables = [];
      newIssues.forEach(function(iss){
        if (iss.kind === 'Deliverable') deliverables.push(iss); else realIssues.push(iss);
      });
      if (realIssues.length) writeSiteIssues(subId, today, '', realIssues, 'Yes', member);
      var asD = db().getSheetByName(ASSIGN_TAB);
      if (asD) deliverables.forEach(function(d){
        if (d.assignedTo && d.description) createIssueTask(d, d.project || '', today, 'Yes');
      });
    }

    // ── 5. Write monthly client feedback to FEEDBACK ──────────
    if (feedback.length > 0) {
      var fbSheet = getOrCreate(FEEDBACK_TAB, writeFeedbackHeaders);
      feedback.forEach(function(f) {
        if (!f.project && !f.comments && !f.overall) return;
        var fbId = nextId(fbSheet, 'FB-');
        fbSheet.appendRow([
          fbId, subId, today, member,
          f.project        || '',
          f.visitDate      || '',
          f.overall        || '',
          f.agenda         || '',
          f.design         || '',
          f.communication  || '',
          f.problemRes     || '',
          f.responsiveness || '',
          f.qualityOfWork  || '',
          f.professionalism|| '',
          f.duration       || '',
          f.recommend      || '',
          f.comments       || '',
          f.appraisal      || '',
          f.referrals      || '',
        ]);
      });
      Logger.log('Feedback written: ' + feedback.length);
    }

    // ── 6. Update BILLING tab (bills, payments, follow-ups) ───
    writeBilling(subId, today, finance);

    // ── 7. Auto-create Site Visit / Meeting tasks for attendees ──
    createMeetingTasks(agendas, member, today);

    Logger.log('submitAmanCRM complete: ' + subId);
    return {status:'ok', subId:subId};

  } catch(err) {
    Logger.log('submitAmanCRM error: ' + String(err));
    return {status:'error', message:String(err)};
  }
}

// ════════════════════════════════════════════════════════════════
// migrateAmanDailyToTabs — one-time, non-destructive migration of the
// old JSON-blob AMAN_DAILY into the new split format.
//   - Backs up the original tab (renamed AMAN_DAILY_OLD_<timestamp>)
//   - Rebuilds a slim AMAN_DAILY (8-col index)
//   - Fills CLIENT_CONNECTIONS (row per contact) + ACTIVITIES (row per entry)
// Header-name based, so it works regardless of the old column layout.
// Safe to re-run: aborts if AMAN_DAILY is already in the new (no-JSON) format.
// ════════════════════════════════════════════════════════════════
function migrateAmanDailyToTabs() {
  var s   = db();
  var old = s.getSheetByName(AMAN_DAILY_TAB);
  if (!old || old.getLastRow() < 2) return {status:'nothing to migrate (no AMAN_DAILY data)'};

  var rows = old.getDataRange().getValues();
  var hdr  = rows[0].map(function(h){ return String(h||'').trim().toLowerCase(); });
  function colIdx(sub){ for (var i=0;i<hdr.length;i++){ if (hdr[i].indexOf(sub.toLowerCase())>-1) return i; } return -1; }

  var iContacts = colIdx('client contacts');
  var iVendorOn = colIdx('vendor coordination'), iVendorE = colIdx('vendor entries');
  var iSiteOn   = colIdx('site issues addressed'), iSiteE = colIdx('site issue entries');
  var iTncpOn   = colIdx('tncp coordination'),   iTncpE  = colIdx('tncp entries');
  var iBniOn    = colIdx('bni activity');

  // Already migrated? (new format has no JSON columns)
  if (iContacts < 0 && iVendorE < 0 && iSiteE < 0 && iTncpE < 0)
    return {status:'AMAN_DAILY already in new format — nothing to migrate'};

  function pj(v){ try { var x = JSON.parse(v); return (x && x.length) ? x : []; } catch(e){ return []; } }
  function dstr(v){ return (v instanceof Date) ? dateStr(v) : String(v||'').substring(0,10); }
  function pad(n){ return ('000'+n).slice(-3); }

  var dailyOut=[], logOut=[];
  var logN=0;
  for (var r=1; r<rows.length; r++){
    var row = rows[r];
    var subId = String(row[0]||'');
    if (subId.indexOf('CRM-') !== 0) continue;
    var date = dstr(row[1]), time = String(row[2]||''), member = String(row[3]||'');
    var vOn = iVendorOn>-1 ? String(row[iVendorOn]||'No') : 'No';
    var sOn = iSiteOn>-1   ? String(row[iSiteOn]||'No')   : 'No';
    var tOn = iTncpOn>-1   ? String(row[iTncpOn]||'No')   : 'No';
    var bOn = iBniOn>-1    ? String(row[iBniOn]||'No')    : 'No';

    dailyOut.push([subId, date, time, member, vOn, sOn, tOn, bOn]);

    if (iContacts>-1) pj(row[iContacts]).forEach(function(c){
      logOut.push(['LOG-'+pad(++logN), subId, date, member, 'Client Connection', c.type||'', c.project||'', c.notes||'']);
    });
    function pushAct(label, idx){
      if (idx<0) return;
      pj(row[idx]).forEach(function(e){
        logOut.push(['LOG-'+pad(++logN), subId, date, member, 'Other Activity', label, e.project||'', e.notes||'']);
      });
    }
    pushAct('Vendor Coordination', iVendorE);
    pushAct('Site Issues Addressed', iSiteE);
    pushAct('TnCP Coordination', iTncpE);
    if (String(bOn).toLowerCase()==='yes') logOut.push(['LOG-'+pad(++logN), subId, date, member, 'Other Activity', 'BNI Activity', '', '']);
  }

  // Back up the original tab (renamed — nothing deleted)
  var backupName = 'AMAN_DAILY_OLD_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'Asia/Kolkata', 'yyyyMMdd_HHmm');
  old.setName(backupName);

  // Fresh slim AMAN_DAILY + combined CRM_LOG
  var daily = getOrCreate(AMAN_DAILY_TAB, writeAmanDailyHeaders);
  var log   = getOrCreate(CRM_LOG_TAB, writeCrmLogHeaders);
  if (dailyOut.length) daily.getRange(daily.getLastRow()+1, 1, dailyOut.length, 8).setValues(dailyOut);
  if (logOut.length)   log.getRange(log.getLastRow()+1, 1, logOut.length, 8).setValues(logOut);

  Logger.log('Migration: daily='+dailyOut.length+' log='+logOut.length+' backup='+backupName);
  return {status:'ok', dailyRows:dailyOut.length, logRows:logOut.length, backup:backupName};
}

// ════════════════════════════════════════════════════════════════
// mergeCrmLogTabs — one-time, non-destructive merge of the existing
// CLIENT_CONNECTIONS + ACTIVITIES tabs into the combined CRM_LOG tab.
// Old tabs are renamed (backed up), not deleted. Idempotent.
// ════════════════════════════════════════════════════════════════
function mergeCrmLogTabs() {
  var s    = db();
  var conn = s.getSheetByName('CLIENT_CONNECTIONS');
  var act  = s.getSheetByName('ACTIVITIES');
  if (!conn && !act) return {status:'no CLIENT_CONNECTIONS / ACTIVITIES tabs to merge'};

  function dstr(v){ return (v instanceof Date) ? dateStr(v) : String(v||'').substring(0,10); }
  function pad(n){ return ('000'+n).slice(-3); }

  var log = getOrCreate(CRM_LOG_TAB, writeCrmLogHeaders);
  var out = [];
  var n   = Math.max(0, log.getLastRow() - 1);   // continue numbering

  if (conn && conn.getLastRow() > 1) {            // cols: A id B sub C date D member E project F type G notes
    var cr = conn.getDataRange().getValues();
    for (var i=1;i<cr.length;i++){ var r=cr[i]; if(!String(r[1]||'')) continue;
      out.push(['LOG-'+pad(++n), String(r[1]||''), dstr(r[2]), String(r[3]||''), 'Client Connection', String(r[5]||''), String(r[4]||''), String(r[6]||'')]); }
  }
  if (act && act.getLastRow() > 1) {              // cols: A id B sub C date D member E activity F project G notes
    var ar = act.getDataRange().getValues();
    for (var j=1;j<ar.length;j++){ var r2=ar[j]; if(!String(r2[1]||'')) continue;
      out.push(['LOG-'+pad(++n), String(r2[1]||''), dstr(r2[2]), String(r2[3]||''), 'Other Activity', String(r2[4]||''), String(r2[5]||''), String(r2[6]||'')]); }
  }
  if (out.length) log.getRange(log.getLastRow()+1, 1, out.length, 8).setValues(out);

  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'Asia/Kolkata', 'yyyyMMdd_HHmm');
  var backups = [];
  if (conn) { conn.setName('CLIENT_CONNECTIONS_OLD_'+ts); backups.push('CLIENT_CONNECTIONS_OLD_'+ts); }
  if (act)  { act.setName('ACTIVITIES_OLD_'+ts);          backups.push('ACTIVITIES_OLD_'+ts); }

  Logger.log('Merged into CRM_LOG: '+out.length+' rows; backups: '+backups.join(', '));
  return {status:'ok', merged:out.length, backups:backups};
}

// ════════════════════════════════════════════════════════════════
// Lead → Project promotion. Once a lead reaches "Briefing Meeting Done"
// or a further stage, add it to the PROJECTS tab as a "New Lead" so it
// enters the pipeline. De-duplicated by project name.
//   PROJECTS cols: A ID, B Name, C Status, D Discipline, E Multiplier, F Lead
// ════════════════════════════════════════════════════════════════
var CRM_PROMOTE_STAGES = ['Briefing Meeting Done','Design Proposal Shared',
                          'Fee Proposal Shared','Lead Converted'];
function isPromoteLeadStage(status) {
  return CRM_PROMOTE_STAGES.indexOf(String(status||'').trim()) > -1;
}
// SLA stage helpers — a briefing is "done" at briefing or any later stage;
// a proposal is "shared" at proposal stages or conversion.
var LEAD_BRIEFING_STAGES = ['Briefing Meeting Done','Design Proposal Shared','Fee Proposal Shared','Lead Converted'];
var LEAD_PROPOSAL_STAGES = ['Design Proposal Shared','Fee Proposal Shared','Lead Converted'];
function leadReachedBriefing(status){ return LEAD_BRIEFING_STAGES.indexOf(String(status||'').trim()) > -1; }
function leadReachedProposal(status){ return LEAD_PROPOSAL_STAGES.indexOf(String(status||'').trim()) > -1; }
// CRM adds projects not yet in the PROJECTS tab (name + type + stage + client).
// Persists them so Aman can raise issues/deliverables without Siddharth's help.
//   PROJECTS cols: A ID, B Name, C Status(stage), D Discipline(type), E Mult, F Lead, G Client
function createCrmProjects(list) {
  if (!list || !list.length) return;
  var projSheet = db().getSheetByName(PROJECTS_TAB);
  if (!projSheet) return;
  var rows = projSheet.getDataRange().getValues();
  // Ensure a "Client" header exists in col G
  if (rows.length && String(rows[0][6]||'').trim() === '') projSheet.getRange(1,7).setValue('Client');
  var existing = {};
  for (var i = 1; i < rows.length; i++) existing[String(rows[i][1]||'').trim().toLowerCase()] = true;
  list.forEach(function(np){
    var nm = String(np.name||'').trim();
    if (!nm || existing[nm.toLowerCase()]) return;
    var pid = nextId(projSheet, 'CP-');
    projSheet.appendRow([ pid, nm, np.stage || 'Ongoing', np.type || '', '', '', np.client || '' ]);
    existing[nm.toLowerCase()] = true;
    Logger.log('CRM added project: ' + nm + ' (' + pid + ')');
  });
}

function promoteLeadToProject(clientName, member) {
  var nm = String(clientName||'').trim();
  if (!nm) return false;
  var projSheet = db().getSheetByName(PROJECTS_TAB);
  if (!projSheet) return false;
  var rows = projSheet.getDataRange().getValues();
  var key  = nm.toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]||'').trim().toLowerCase() === key) return false; // already in PROJECTS
  }
  var pid = nextId(projSheet, 'NL-');
  projSheet.appendRow([ pid, nm, 'New Lead', '', '', '' ]);
  Logger.log('Lead promoted to PROJECTS: ' + nm + ' (' + pid + ')');
  return true;
}

// ════════════════════════════════════════════════════════════════
// writeBilling — BILLING tab. Payments match a bill by INVOICE NO. first,
// else fall back to the oldest unpaid bill of the same project. Follow-up
// dates stamped on unpaid bills of the followed-up projects.
//   Cols: A BillID, B InvoiceNo, C Project, D BillDate, E BillAmt,
//         F AmtRecd, G RecdDate, H LastFollowup, I Status, J SubmissionID
// ════════════════════════════════════════════════════════════════
function writeBilling(subId, today, finance) {
  finance = finance || {};
  var bills      = finance.bills || [];
  var payments   = (finance.paymentReceived === 'Yes') ? (finance.payments || []) : [];
  var fupProj    = (finance.followupsDone === 'Yes')   ? (finance.followupProjects || []) : [];
  if (finance.billsRaised !== 'Yes') bills = [];

  if (!bills.length && !payments.length && !fupProj.length) return;

  var sheet = getOrCreate(BILLING_TAB, writeBillingHeaders);

  // 1) Append a row for each bill raised today
  bills.forEach(function(b) {
    if (!b.project && !b.amount && !b.invoice) return;
    var billId = nextId(sheet, 'BILL-');
    prependRow(sheet, [
      billId, b.invoice || '', b.project || '', today, parseFloat(b.amount) || 0,
      '', '', '', 'Pending', subId,
    ]);
  });

  // Re-read after appends so matching sees today's bills too
  var data = sheet.getDataRange().getValues();  // row0 = header
  function norm(s){ return String(s||'').trim().toLowerCase(); }

  // 2) Attach each payment — match by Invoice No. first, else oldest unpaid of project
  payments.forEach(function(p) {
    var amt = parseFloat(p.amount) || 0;
    if (!p.project && !amt && !p.invoice) return;
    var targetRow = -1;
    if (p.invoice) {                                          // (a) exact invoice match
      for (var i = 1; i < data.length; i++) {
        if (norm(data[i][1]) !== norm(p.invoice)) continue;   // B Invoice No.
        if (String(data[i][8]) === 'Paid') continue;          // I Status
        targetRow = i; break;
      }
    }
    if (targetRow < 0) {                                       // (b) oldest unpaid of project
      for (var j = 1; j < data.length; j++) {
        if (norm(data[j][2]) !== norm(p.project)) continue;   // C Project
        if (String(data[j][8]) === 'Paid') continue;          // I Status
        targetRow = j; break;
      }
    }
    if (targetRow > -1) {
      var billAmt  = parseFloat(data[targetRow][4]) || 0;     // E Bill Amount
      var already  = parseFloat(data[targetRow][5]) || 0;     // F Amount Received
      var recd     = already + amt;
      var status   = (billAmt > 0 && recd >= billAmt) ? 'Paid' : 'Partial';
      sheet.getRange(targetRow+1, 6).setValue(recd);          // F Amount Received
      sheet.getRange(targetRow+1, 7).setValue(today);         // G Received Date
      sheet.getRange(targetRow+1, 9).setValue(status);        // I Status
      data[targetRow][5] = recd; data[targetRow][8] = status;
    } else {
      // Payment with no matching bill — record as its own row
      var pid = nextId(sheet, 'BILL-');
      prependRow(sheet, [ pid, p.invoice || '', p.project || '', '', '', amt, today, '', 'Payment (no bill)', subId ]);
    }
  });

  // 3) Stamp last follow-up date on unpaid bills of followed-up projects
  if (fupProj.length) {
    data = sheet.getDataRange().getValues();
    fupProj.forEach(function(proj) {
      for (var i = 1; i < data.length; i++) {
        if (norm(data[i][2]) !== norm(proj)) continue;        // C Project
        if (String(data[i][8]) === 'Paid') continue;          // I Status
        sheet.getRange(i+1, 8).setValue(today);               // H Last Follow-up Date
      }
    });
  }
  Logger.log('BILLING updated — bills:'+bills.length+' payments:'+payments.length+' followups:'+fupProj.length);
}

// ════════════════════════════════════════════════════════════════
// createMeetingTasks — for each tomorrow's meeting/site-visit agenda,
// create a TASK_ASSIGNMENTS row for every team-member attendee.
// Default = 1 weighted point; overwritten later when the member logs
// hours in their DPR (weekly scoring recalculates points = hours).
// ════════════════════════════════════════════════════════════════
function createMeetingTasks(agendas, member, today) {
  if (!agendas || !agendas.length) return;
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) return;
  var count = 0;
  agendas.forEach(function(ag) {
    var attendees = ag.teamAttendees || [];
    if (!attendees.length) return;
    var type = (ag.type === 'Site Visit') ? 'Site Visit' : 'Meeting';
    attendees.forEach(function(person) {
      if (!person) return;
      var newId = 'T-' + Utilities.getUuid().substring(0,8).toUpperCase();
      var notes = (type + ' — ' + (ag.project||'') +
                   (ag.time ? ' @ ' + ag.time : '') +
                   (ag.agenda ? '. Agenda: ' + ag.agenda : '') +
                   ' [CRM auto; points = hours logged in DPR — Site Visit ×2/hr, Meeting ×1/hr]');
      sheet.appendRow([
        newId,                        // A TaskID
        '',                           // B ProjectID
        ag.project || '',             // C ProjectName
        person,                       // D AssignedTo
        type,                         // E Stage/Type
        1,                            // F Disc Multiplier
        0,                            // G Base Pts (0 — visit pts come from hours)
        1,                            // H Units
        0,                            // I Weighted Pts (0 until Done w/ hours → hours × rate)
        today,                        // J AssignedDate
        ag.date || today,             // K Deadline (meeting/visit date)
        '', '',                       // L Area, M Drawing
        'Not Started',                // N SelfStatus
        '', '',                       // O SelfStatusDate, P ActualCompletion
        'Pending',                    // Q LeadApproved
        '', '', '',                   // R ApprovedBy, S ApprovalDate, T RevisionTag
        notes,                        // U Notes
        'CRM — ' + member,            // V AssignedBy
        'Medium',                     // W Priority
      ]);
      count++;
    });
  });
  Logger.log('Meeting/visit tasks created: ' + count);
}
