// ═══════════════════════════════════════════════════════════════
// IDEAFORM DESIGN STUDIO — Apps Script Final
// Sheet structure verified against IDS_Productivity_System_2.xlsx
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
  return VISIT_TYPES_ARCH.indexOf(t) > -1 || VISIT_TYPES_MTG.indexOf(t) > -1;
}

function calcVisitPts(taskType, hours) {
  var t = String(taskType||'').trim();
  var h = parseFloat(hours) || 0;
  if (VISIT_TYPES_ARCH.indexOf(t) > -1) return Math.round(h * 2 * 10) / 10;
  if (VISIT_TYPES_MTG.indexOf(t)  > -1) return Math.round(h * 1 * 10) / 10;
  return 0;
}

var TASK_TAB      = 'TASK_LOG';
var SUMMARY_TAB   = 'DAILY_SUMMARY';
var CONFIG_TAB    = 'CONFIG';
var PROJECTS_TAB  = 'PROJECTS';
var ASSIGN_TAB    = 'TASK_ASSIGNMENTS';
var TEAM_TAB      = 'TEAM';
var SCORECARD_TAB = 'TEAM_SCORECARD';
var APPROVAL_FORM_URL = 'https://sidinani14.github.io/dpr/approval.html';
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

function doGet(e) {
  // Handle CORS preflight — Apps Script handles this automatically
  // but we ensure JSON content type is always set
  var p = e && e.parameter ? e.parameter : {};
  var action = p.action || '', member = p.member || '';
  if (action === 'getLists')              return safeRespond(getLists);
  if (action === 'getConfig')             return safeRespond(readConfig);
  if (action === 'getPendingTasks')       return safeRespond(getPendingTasks);
  if (action === 'getAllTasks')           return safeRespond(getAllTasks);
  if (action === 'getOpenTasksForMember')    return safeRespond(function() { return getOpenTasksForMember(member); });
  if (action === 'getNotifications')         return safeRespond(function() { return getNotificationsForMember(member); });
  if (action === 'getWeeklyStats')           return safeRespond(function() { return getWeeklyStats(p.weekStart||''); });
  if (action === 'getProjectStats')          return safeRespond(getProjectStats);
  if (action === 'getCalendarData')          return safeRespond(getCalendarData);
  return respond({status: 'IDS DPR live'});
}

function doPost(e) {
  try {
    var raw  = e && e.postData ? e.postData.contents : '';
    var data = JSON.parse(raw);
    Logger.log('doPost action: ' + data.action + ' | raw: ' + raw.substring(0,100));

    // GET-style actions sent via POST to bypass CORS
    if (data.action === 'getAllTasks')         return respond(getAllTasks());
    if (data.action === 'getLists')            return respond(getLists());
    if (data.action === 'getConfig')           return respond(readConfig());
    if (data.action === 'getPendingTasks')     return respond(getPendingTasks());
    if (data.action === 'getOpenTasksForMember') return respond(getOpenTasksForMember(data.member||''));
    if (data.action === 'getNotifications')    return respond(getNotificationsForMember(data.member||''));
    if (data.action === 'getWeeklyStats')      return respond(getWeeklyStats(data.weekStart||''));
    if (data.action === 'getProjectStats')     return respond(getProjectStats());
    if (data.action === 'getSiteIssues')       return respond(getSiteIssues(data.project||''));
    if (data.action === 'resolveIssue')        return respond(resolveIssue(data.issueId||''));
    if (data.action === 'getSiteExecution')    return respond(getSiteExecutionSummary(data.project||''));
    if (data.action === 'submitDPER')          return respond(handleDPERSubmission(data));
    if (data.action === 'getCalendarData')     return respond(getCalendarData());

    // Form submission actions
    if (data.action === 'submitApprovals')   return respond(submitApprovals(data));
    if (data.action === 'assignTasks')       return respond(assignTasks(data));
    if (data.action === 'markNotifSeen')     return respond(markNotificationSeen(data));
    if (data.action === 'createSelfTask')    return respond(createSelfAssignedTask(data));
    if (data.action === 'createDoneTask')    return respond(createDoneTask(data));
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

    if (a && SKIP_A.indexOf(a.toUpperCase()) === -1 && !a.startsWith('—') &&
        !a.startsWith('←') && !isNaN(parseFloat(b)))
      stages.push({label:a, pts:parseFloat(b), category:c});

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

  Logger.log('readConfig: found ' + leads.length + ' leads: ' + leads.map(function(l){return l.name;}).join(', '));
  return {stages:stages, disciplines:discs, visits:visits,
          leads:leads, scoringWeights:sw, revPenalties:rp};
}

// ════════════════════════════════════════════════════════════════
// TASK LOG — write DPR unplanned tasks
//
// TASK_LOG cols (14 total):
// A:SubID B:Date C:Member D:Project E:Discipline F:TaskType
// G:Area H:DrawingName I:Units J:WeightedPts K:LeadApproved
// L:ApprovedBy M:ReviewedOn N:Notes
// ════════════════════════════════════════════════════════════════
function writeTaskLog(data, visitPts) {
  var sheet = getOrCreate(TASK_TAB, writeTaskLogHeaders);
  var subId = Utilities.getUuid().substring(0, 8).toUpperCase();
  var date  = dateStr(data['Timestamp']);
  var member= data['Member'] || '';

  for (var p = 1; p <= 3; p++) {
    var project = data['P' + p + ' Name'];
    if (!project) continue;
    var disc       = data['P' + p + ' Discipline'] || '';
    var tasksRaw   = data['P' + p + ' Tasks']       || '';
    var visitRaw   = data['P' + p + ' Visit']       || 'No';
    var visitNotes = data['P' + p + ' Visit Notes'] || '';

    // Parse "TaskType [Area:X] [Drawing:Y] ×N = P pts"
    var parts = tasksRaw ? tasksRaw.split(' | ') : [];
    parts.forEach(function(part) {
      part = part.trim();
      var taskType = '', area = '', drawing = '', units = 0, pts = 0;
      var aM = part.match(/\[Area:\s*([^\]]+)\]/);
      var dM = part.match(/\[Drawing:\s*([^\]]+)\]/);
      if (aM) { area    = aM[1].trim(); part = part.replace(aM[0], '').trim(); }
      if (dM) { drawing = dM[1].trim(); part = part.replace(dM[0], '').trim(); }
      var xI = part.lastIndexOf('×'), eI = part.lastIndexOf('=');
      if (xI > -1 && eI > -1) {
        taskType = part.substring(0, xI).trim();
        units    = parseFloat(part.substring(xI+1, eI).trim()) || 0;
        pts      = parseFloat(part.substring(eI+1).replace('pts','').trim()) || 0;
      } else { taskType = part; }
      sheet.appendRow([subId, date, member, project, disc, taskType,
                        area, drawing, units, pts, 'Pending', '', '', '']);
    });

    // Visit row
    if (visitRaw && visitRaw !== 'No') {
      var vp   = visitRaw.split('·');
      var vTyp = vp[0] ? vp[0].trim() : visitRaw;
      var vDur = vp[1] ? vp[1].trim() : '';
      var vPts = visitPts[vDur] !== undefined ? visitPts[vDur] : 1.0;
      sheet.appendRow([subId, date, member, project, disc, vTyp,
                        '', '', 1, vPts, 'Pending', '', '', visitNotes]);
    }
  }
  var lr = sheet.getLastRow();
  if (lr > 1) sheet.getRange(2, 10, lr-1, 1).setNumberFormat('0.00');
}

function writeTaskLogHeaders(s) {
  var h = ['Submission ID','Date','Member','Project','Discipline','Task Type',
           'Area','Drawing Name','Units','Weighted Pts','Lead Approved',
           'Approved By','Reviewed on','Notes'];
  var r = s.getRange(1, 1, 1, h.length);
  r.setValues([h]);
  r.setBackground('#1F3A5F'); r.setFontColor('#FFFFFF');
  r.setFontWeight('bold'); r.setFontSize(10);
  s.setFrozenRows(1);
  [100,90,140,200,180,240,120,160,60,100,110,130,140,200]
    .forEach(function(w, i) { s.setColumnWidth(i+1, w); });
}

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
  sheet.appendRow([
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
  var h = ['Date','Member','E-mail','Arrived on Time',
           'Full List of Deliverables','Anything Blocking','Mood (1–5)'];
  var r = s.getRange(1, 1, 1, h.length);
  r.setValues([h]);
  r.setBackground('#1F3A5F'); r.setFontColor('#FFFFFF');
  r.setFontWeight('bold'); r.setFontSize(10);
  s.setFrozenRows(1);
  [90,140,180,110,420,300,80]
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

  tasks.forEach(function(t) {
    sheet.appendRow([
      'T-' + Utilities.getUuid().substring(0,8).toUpperCase(), // A
      data.projectId    || '',        // B
      data.project      || '',        // C
      data.assignedTo   || '',        // D
      t.taskType        || '',        // E
      data.discipline   || '',        // F
      t.basePts         || 0,         // G
      data.multiplier   || 1,         // H
      t.weightedPts     || 0,         // I
      data.dateAssigned || today,     // J
      t.targetDate      || '',        // K
      t.area            || '',        // L
      t.drawing         || '',        // M
      'Not Started',                  // N SelfStatus
      '',                             // O SelfStatusDate
      '',                             // P ActualCompletionDate (NEW)
      'Pending',                      // Q LeadApproved
      '',                             // R ApprovedBy
      '',                             // S ApprovalDate
      '',                             // T RevisionTag
      t.notes           || '',        // U Notes
      data.assignedBy   || '',        // V AssignedBy
      t.priority        || 'Medium',  // W Priority
    ]);
    written++;
  });
  return {status:'ok', written:written};
}

function writeAssignHeaders(s) {
  var h = ['Task ID','Project ID','Project Name','Assigned To','Stage','Discipline',
           'Stage Base Pts','Disc. Multiplier','Weighted Points','Assigned Date','Deadline',
           'Area','Drawing Name','Self Status','Self Done Date','Lead Approved',
           'Approved By','Approval Date','Revision Tag','Notes','Assigned by','Priority'];
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
    if (selfStatus === 'Done') {
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

    tasks.push({
      row              : i+1,
      taskId           : String(r[0]  || ''),
      projectId        : String(r[1]  || ''),
      project          : String(r[2]  || ''),
      assignedTo       : String(r[3]  || ''),
      taskType         : String(r[4]  || ''),
      discipline       : String(r[5]  || ''),
      basePts          : r[6]  || 0,
      multiplier       : r[7]  || 1,
      weightedPts      : r[8]  || 0,
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
    else if (a.status === 'Blocked') newStatus = 'In Progress';
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
        sheet.getRange(i+1, 14).setValue(newStatus); // N SelfStatus
        sheet.getRange(i+1, 15).setValue(now);       // O SelfStatusDate always

        if (newStatus === 'Done') {
          // P ActualCompletionDate
          var actualDate = a.actualCompletionDate || now;
          sheet.getRange(i+1, 16).setValue(actualDate);

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

  // ── SOURCE 1: TASK_LOG — unplanned/additional work from DPR Section 3 ──
  // TASK_LOG cols: SubID(A=0) Date(B=1) Member(C=2) Project(D=3) Disc(E=4)
  //   TaskType(F=5) Area(G=6) Drawing(H=7) Units(I=8) Pts(J=9) LeadApproved(K=10)
  var tlSheet = db().getSheetByName(TASK_TAB);
  if (tlSheet) {
    var tlRows = tlSheet.getDataRange().getValues();
    for (var i = 1; i < tlRows.length; i++) {
      var status = String(tlRows[i][10] || '').trim(); // K LeadApproved
      // Show tasks where LeadApproved is blank or Pending (not Yes/No)
      if (status === 'Yes' || status === 'No') continue;
      var member   = String(tlRows[i][2] || '');
      var project  = String(tlRows[i][3] || '');
      var taskType = String(tlRows[i][5] || '');
      var area     = String(tlRows[i][6] || '');
      var drawing  = String(tlRows[i][7] || '');
      var arRow    = findTaskAssignRow(member, project, taskType, area, drawing);
      tasks.push({
        row          : i + 1,
        source       : 'TASK_LOG',
        subId        : String(tlRows[i][0] || ''),
        date         : String(tlRows[i][1] || '').substring(0, 10),
        member       : member,
        project      : project,
        discipline   : String(tlRows[i][4] || ''),
        taskType     : taskType,
        area         : area,
        drawing      : drawing,
        units        : tlRows[i][8]  || 0,
        pts          : tlRows[i][9]  || 0,
        taskAssignRow: arRow,
      });
    }
  }

  // ── SOURCE 2: TASK_ASSIGNMENTS — assigned tasks marked Done by member ──
  // These never appear in TASK_LOG — they are assigned via tasks form
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
function findTaskAssignRow(member, project, taskType, area, drawing) {
  var sheet = db().getSheetByName(ASSIGN_TAB);
  if (!sheet) return null;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][3]  || '').trim() === member   &&  // D
        String(rows[i][2]  || '').trim() === project  &&  // C
        String(rows[i][4]  || '').trim() === taskType &&  // E
        String(rows[i][16] || '').trim() === 'Pending') { // Q LeadApproved (shifted)
      return i + 1;
    }
  }
  return null;
}

// ════════════════════════════════════════════════════════════════
// SUBMIT APPROVALS
// TASK_LOG: LeadApproved=K(11) ApprovedBy=L(12) ReviewedOn=M(13) Notes=N(14)
// TASK_ASSIGNMENTS: LeadApproved=P(16) ApprovedBy=Q(17) ApprovalDate=R(18)
//   SelfStatus=N(14) SelfDoneDate=O(15) RevisionTag=S(19) Notes=T(20)
// ════════════════════════════════════════════════════════════════
function submitApprovals(data) {
  var s        = db();
  var taskLog  = s.getSheetByName(TASK_TAB);
  var taskAssn = s.getSheetByName(ASSIGN_TAB);
  if (!taskLog) return {status:'error', message:'TASK_LOG not found'};

  var approvals = data.approvals || [];
  var approved = 0, rejected = 0;
  var now = nowStr(), today = dateStr();

  approvals.forEach(function(a) {
    var row    = parseInt(a.row);
    var source = a.source || 'TASK_LOG'; // which tab this came from
    if (!row || row < 2) return;

    if (source === 'TASK_LOG') {
      // ── Update TASK_LOG (1-indexed cols) ────────────────────
      taskLog.getRange(row, 11).setValue(a.decision);   // K LeadApproved
      taskLog.getRange(row, 12).setValue(a.leadName);   // L ApprovedBy
      taskLog.getRange(row, 13).setValue(now);           // M ReviewedOn
      taskLog.getRange(row, 14).setValue(a.note || ''); // N Notes
      // If correcting task type or points in TASK_LOG
      if (a.correctedTaskType) taskLog.getRange(row, 6).setValue(a.correctedTaskType); // F TaskType
      if (a.correctedPts)      taskLog.getRange(row, 10).setValue(parseFloat(a.correctedPts)); // J WeightedPts
      // Write notification on rejection
      if (a.decision === 'No') {
        var tlMember = String(taskLog.getRange(row, 3).getValue()||'');
        var tlNotif  = {
          project:'', taskType:String(taskLog.getRange(row,6).getValue()||''),
          correctedTaskType:a.correctedTaskType||'',
          originalPts:taskLog.getRange(row,10).getValue()||0,
          correctedPts:a.correctedPts||'',
          note:a.note||'', leadName:a.leadName||'',
        };
        // Notification removed — rejections shown in DPR form directly
      }
    }

    // ── Update TASK_ASSIGNMENTS (both sources need this) ──────
    // For TASK_LOG source: taskAssignRow is found by matching member+project+task
    // For TASK_ASSIGNMENTS source: row IS the taskAssignRow directly
    var ar = source === 'TASK_ASSIGNMENTS' ? row : parseInt(a.taskAssignRow || 0);

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
  var tlSheet  = s.getSheetByName(TASK_TAB);
  var asSheet  = s.getSheetByName(ASSIGN_TAB);
  var sumSheet = s.getSheetByName(SUMMARY_TAB);
  var tSheet   = s.getSheetByName(TEAM_TAB);
  if (!tSheet) { Logger.log('TEAM tab not found'); return; }

  var tlRows  = tlSheet  ? tlSheet.getDataRange().getValues()  : [];
  var asRows  = asSheet  ? asSheet.getDataRange().getValues()  : [];
  var sumRows = sumSheet ? sumSheet.getDataRange().getValues() : [];
  var tRows   = tSheet.getDataRange().getValues();

  // Auto-detect TASK_ASSIGNMENTS column structure (22-col vs 23-col)
  var asHeaders = asRows[0] ? asRows[0].map(function(h){ return String(h||'').trim(); }) : [];
  var scIs23    = asHeaders.length >= 23 || asHeaders.indexOf('Actual Completion Date') > -1;
  var SC_LEADAPPR = scIs23 ? 16 : 15; // Q(16) new / P(15) old
  var SC_APPDATE  = scIs23 ? 18 : 17; // S(18) new / R(17) old
  var SC_DONEDATE = scIs23 ? 15 : 14; // P(15) new / O(14) old
  var SC_STATUSDT = 14;               // O always
  var SC_DEADLINE = 10;               // K always
  var SC_REVTAG   = scIs23 ? 19 : 18; // T(19) new / S(18) old
  var SC_ASSIGNDT = 9;                // J always
  Logger.log('SC col structure: '+(scIs23?'23-col':'22-col')+' | LeadApproved='+SC_LEADAPPR+' | ApprovalDate='+SC_APPDATE);

  // Write to IDS Team Scorecard sheet — clear and rewrite (live snapshot)
  var scorecard = getOrCreateScorecard(SCORECARD_TAB, writeScorecardHeaders);
  var lastScRow = scorecard.getLastRow();
  if (lastScRow > 1) scorecard.getRange(2, 1, lastScRow-1, 13).clearContent();

  for (var ti = 1; ti < tRows.length; ti++) {
    var name   = String(tRows[ti][0] || '').trim();
    var role   = String(tRows[ti][1] || '').trim();
    var wkTgt  = parseFloat(tRows[ti][2]) || 50; // C WeeklyTarget (=D*6 formula)
    var active = String(tRows[ti][5] || '').trim().toLowerCase();
    if (!name || active === 'no') continue;

    // 1. Approved pts this week — two sources:
    // SOURCE A: TASK_LOG (unplanned DPR work)
    //   Member=C(2) LeadApproved=K(10) Date=B(1) WeightedPts=J(9)
    // SOURCE B: TASK_ASSIGNMENTS (assigned tasks marked done + approved)
    //   AssignedTo=D(3) LeadApproved=P(15) ApprovalDate=R(17) WeightedPts=I(8)
    var approvedPts = 0;

    // TASK_LOG source removed — all tasks now in TASK_ASSIGNMENTS

    // Source B: TASK_ASSIGNMENTS — filter by SelfStatusDate so work done in the week
    // counts even if lead approves after the weekend
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var selfDoneDate = cellDate(r[SC_STATUSDT]); // O SelfStatusDate
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
    // DAILY_SUMMARY: Date=A(0) Member=B(1) ArrivedOnTime=D(3)
    var memberSum = [];
    sumRows.forEach(function(r, i) {
      if (i === 0) return;
      var rDate = String(r[0] || '').substring(0, 10);
      if (String(r[1] || '').trim() === name && rDate >= monStr && rDate <= sunStr)
        memberSum.push({date:rDate, ontime:String(r[3] || '').trim()});
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

    // Source B: TASK_ASSIGNMENTS — use SelfStatusDate for month attribution
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var selfDoneDate = cellDate(r[SC_STATUSDT]);
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
  archiveSheet(TASK_TAB,    'ARCHIVE',              1, 10); // Date=B(1), LeadApproved=K(10)
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
  var tlSheet  = s.getSheetByName(TASK_TAB);
  var projSheet= s.getSheetByName(PROJECTS_TAB);

  if (!tlSheet)   { Logger.log('TASK_LOG not found'); return; }
  if (!projSheet) { Logger.log('PROJECTS tab not found'); return; }

  // ── Load TASK_LOG visit rows ──────────────────────────────────
  var tlRows  = tlSheet.getDataRange().getValues();
  var visits  = [];   // all visit entries
  var today   = new Date(); today.setHours(0,0,0,0);

  for (var i = 1; i < tlRows.length; i++) {
    var r        = tlRows[i];
    var taskType = String(r[5] || '').trim().toLowerCase();
    var isVisit  = taskType.includes('site visit') || taskType.includes('site supervision');
    var isMeeting= taskType.includes('client meeting') || taskType.includes('client + site') ||
                   taskType.includes('client meeting + site');
    if (!isVisit && !isMeeting) continue;

    var rawDate  = r[1];
    var visitDate= rawDate instanceof Date ? rawDate : new Date(rawDate);
    if (isNaN(visitDate.getTime())) continue;
    visitDate.setHours(0,0,0,0);

    visits.push({
      date      : visitDate,
      dateStr   : cellDate(r[1]),
      project   : String(r[3] || '').trim(),
      member    : String(r[2] || '').trim(),
      taskType  : String(r[5] || '').trim(),
      discipline: String(r[4] || '').trim(),
      units     : r[8] || 1,
      pts       : r[9] || 0,
      notes     : String(r[13] || '').trim(),   // N Notes col
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
  sheet.getRange(logStart, 9).setValue('All site visits and client meetings from TASK_LOG');
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
    sheet.getRange(logStart+2, 1).setValue('No visit data found in TASK_LOG yet.');
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
  var tlSheet  = s.getSheetByName(TASK_TAB);
  var asSheet  = s.getSheetByName(ASSIGN_TAB);
  var sumSheet = s.getSheetByName(SUMMARY_TAB);
  var tSheet   = s.getSheetByName(TEAM_TAB);
  if (!tSheet) { Logger.log('TEAM tab not found'); return; }

  var tlRows  = tlSheet  ? tlSheet.getDataRange().getValues()  : [];
  var asRows  = asSheet  ? asSheet.getDataRange().getValues()  : [];
  var sumRows = sumSheet ? sumSheet.getDataRange().getValues() : [];
  var tRows   = tSheet.getDataRange().getValues();

  // Auto-detect TASK_ASSIGNMENTS column structure (22-col vs 23-col)
  var asHeaders = asRows[0] ? asRows[0].map(function(h){ return String(h||'').trim(); }) : [];
  var scIs23    = asHeaders.length >= 23 || asHeaders.indexOf('Actual Completion Date') > -1;
  var SC_LEADAPPR = scIs23 ? 16 : 15; // Q(16) new / P(15) old
  var SC_APPDATE  = scIs23 ? 18 : 17; // S(18) new / R(17) old
  var SC_DONEDATE = scIs23 ? 15 : 14; // P(15) new / O(14) old
  var SC_STATUSDT = 14;               // O always
  var SC_DEADLINE = 10;               // K always
  var SC_REVTAG   = scIs23 ? 19 : 18; // T(19) new / S(18) old
  var SC_ASSIGNDT = 9;                // J always
  Logger.log('SC col structure: '+(scIs23?'23-col':'22-col')+' | LeadApproved='+SC_LEADAPPR+' | ApprovalDate='+SC_APPDATE);

  // Write to IDS Team Scorecard sheet — clear and rewrite (live snapshot)
  var scorecard = getOrCreateScorecard(SCORECARD_TAB, writeScorecardHeaders);
  var lastScRow = scorecard.getLastRow();
  if (lastScRow > 1) scorecard.getRange(2, 1, lastScRow-1, 13).clearContent();

  // Remove any existing rows for current week (avoid duplicates on re-run)
  var lastRow = scorecard.getLastRow();
  if (lastRow > 1) {
    var scRows = scorecard.getRange(2, 1, lastRow-1, 1).getValues();
    var toDelete = [];
    scRows.forEach(function(r, i) {
      // Check if this row is from the current week — col A = Week Of date
      // For current week rows we stored monStr
      // Actually scorecard stores member in col A, not week
      // Week is not stored in TEAM_SCORECARD — skip dedup, just append
    });
  }

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
      var selfDoneDate = cellDate(r[SC_STATUSDT]); // O SelfStatusDate — when team marked done
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
    var memberSum = [];
    sumRows.forEach(function(r, i) {
      if (i === 0) return;
      var rDate = cellDate(r[0]);
      if (String(r[1] || '').trim() === name && rDate >= monStr && rDate <= sunStr)
        memberSum.push({date:rDate, ontime:String(r[3]||'').trim()});
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
      var selfDoneDate = cellDate(r[SC_STATUSDT]);
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
  for (var i = 1; i < rows.length; i++) {
    var r       = rows[i];
    var taskId  = String(r[0]||'').trim();   // A
    var proj    = String(r[2]||'').trim();   // C
    var tType   = String(r[4]||'').trim();   // E
    var deadline= cellDate(r[10]);           // K
    var status  = String(r[13]||'').trim();  // N SelfStatus
    var appr    = String(r[15]||'').trim();  // P LeadApproved
    if (!proj || !deadline || !isVisitType(tType)) continue;
    if (status === 'Done' && appr === 'Yes') continue; // completed
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
        pd.disc,                  // F Discipline
        isVisit ? 0 : 0,          // G BasePts (duration-based)
        pd.mult,                  // H Multiplier
        isVisit ? 3 : 1,          // I WeightedPts placeholder (Visit=3, Meeting=1)
        today,                    // J AssignedDate
        nextDate,                 // K Deadline
        '',                       // L Area
        '',                       // M Drawing — blank for auto-scheduled visits
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
        newId,'',entry.project,entry.assignee,entry.visitType,
        pd.disc||'',0,pd.mult||1,0,today,today,'','',
        'Not Started','','Pending','','','',
        'MISSED VISIT — was due '+nextDate+'. '+entry.notes,
        'Auto-scheduled','High',
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
function syncVisitSchedule() {
  Logger.log('=== syncVisitSchedule: '+new Date().toISOString()+' ===');

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

  // Calculate points — visit tasks use hours×rate, others use basePts×mult
  var basePts = parseFloat(data.basePts) || 0;
  var weightedPts;
  if (isVisitTask(data.taskType) && data.visitHours) {
    weightedPts = calcVisitPts(data.taskType, data.visitHours);
    basePts     = parseFloat(data.visitHours) || basePts;
  } else {
    weightedPts = Math.round(basePts * mult * 10) / 10;
  }

  var newId = 'T-'+Utilities.getUuid().substring(0,8).toUpperCase();
  var member = data.member || '';

  sheet.appendRow([
    newId,
    '',
    data.project      || '',   // C ProjectName
    member,                    // D AssignedTo (self)
    data.taskType     || '',   // E Stage
    disc,                      // F Discipline
    basePts,                   // G StageBasePts
    mult,                      // H Multiplier
    weightedPts,               // I WeightedPoints
    data.date         || today, // J AssignedDate
    today,                     // K Deadline = today (Done same day)
    data.area         || '',   // L Area
    data.drawing      || '',   // M Drawing
    'Done',                    // N SelfStatus
    today,                     // O SelfStatusDate
    today,                     // P ActualCompletionDate
    'Pending',                 // Q LeadApproved
    '',                            // R ApprovedBy
    '',                            // S ApprovalDate
    '',                            // T RevisionTag
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
  var weightedPts= parseFloat((basePts * mult).toFixed(2));
  var newId = 'T-'+Utilities.getUuid().substring(0,8).toUpperCase();

  sheet.appendRow([
    newId,
    '',
    data.project      || '',   // C
    data.member       || '',   // D AssignedTo (self)
    data.taskType     || '',   // E Stage
    disc,                      // F Discipline
    basePts,                   // G BasePts
    mult,                      // H Multiplier
    weightedPts,               // I WeightedPts
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

  Logger.log('Self-assigned task created: '+data.member+' / '+data.taskType+' → '+newId);
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

  var tlSheet  = s.getSheetByName(TASK_TAB);
  var asSheet  = s.getSheetByName(ASSIGN_TAB);
  var sumSheet = s.getSheetByName(SUMMARY_TAB);
  var tSheet   = s.getSheetByName(TEAM_TAB);
  if (!tSheet) return {error:'TEAM tab not found'};

  var tlRows  = tlSheet  ? tlSheet.getDataRange().getValues()  : [];
  var asRows  = asSheet  ? asSheet.getDataRange().getValues()  : [];
  var sumRows = sumSheet ? sumSheet.getDataRange().getValues() : [];
  var tRows   = tSheet.getDataRange().getValues();

  // Auto-detect column structure
  var asHeaders   = asRows[0] ? asRows[0].map(function(h){ return String(h||'').trim(); }) : [];
  var is23        = asHeaders.length >= 23 || asHeaders.indexOf('Actual Completion Date') > -1;
  var COL_LEADAPPR= is23 ? 16 : 15;
  var COL_STATUSDT= 14;
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
      var doneDate   = cellDate(r[COL_STATUSDT]);
      var isApproved = String(r[COL_LEADAPPR]||'').trim() === 'Yes';
      if (String(r[3]||'').trim() === name &&
          String(r[13]||'').trim() === 'Done' &&
          isApproved &&
          doneDate >= mon && doneDate <= sat) {
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

    // Approved pts this week (SelfStatusDate filter)
    var approvedPts = 0;
    // TASK_LOG source removed
    asRows.forEach(function(r, i) {
      if (i === 0) return;
      var doneDate = cellDate(r[COL_STATUSDT]);
      if (String(r[3]||'').trim() === name &&
          String(r[COL_LEADAPPR]||'').trim() === 'Yes' &&
          doneDate >= mon && doneDate <= sat)
        approvedPts += parseFloat(r[8]) || 0;
    });

    // DPR days filed this week
    // DAILY_SUMMARY cols: Date(A=0) Time(B=1) Member(C=2) Email(D=3)
    var dprDays = 0;
    sumRows.forEach(function(r, i) {
      if (i === 0) return;
      var rDate = cellDate(r[0]); // col A = Date (Date object)
      var rName = String(r[2]||'').trim(); // col C = Member name
      if (rName === name && rDate >= mon && rDate <= sat)
        dprDays++;
    });
    // Mon 18 May — everyone gets credit (form was down)
    if (mon === '2026-05-18' && dprDays === 0) dprDays = 1;
    else if (mon === '2026-05-18') dprDays = Math.min(dprDays + 1, 6);

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
  var COL_SSTAT= 13;
  var COL_SDATE= 14;
  var COL_DEAD = 10;

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
    var sDate   = cellDate(r[COL_SDATE]);
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
    var decisions = [];
    try { issues    = data.issues       ? JSON.parse(data.issues)       : []; } catch(e){}
    try { decisions = data.decisionItems ? JSON.parse(data.decisionItems): []; } catch(e){}
    writeSiteIssues(subId, data.date||dateStr(), data.project||'',
                    issues, data.onTrack||'Yes');
    writeSiteDecisions(subId, data.date||dateStr(), data.project||'',
                       decisions, data.onTrack||'Yes', data.lead||'');

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

    Logger.log('DPER submitted: ' + subId + ' / ' + (data.project||''));
    return {status:'ok', subId:subId};
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
    'Resolution Notes','Resolved Date'];
  sheet.getRange(1,1,1,h.length).setValues([h])
    .setBackground('#8B2020').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  var widths=[140,140,100,200,70,160,340,160,120,110,120,180,220,110];
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

  sheet.appendRow([
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
    data.blocking     || '',
    data.decisions    || '',
    data.remarks      || '',
  ]);
  return subId;
}

// ── writeSiteIssues ────────────────────────────────────────
function writeSiteIssues(subId, date, project, issues, onTrack) {
  if (!issues || issues.length === 0) return [];
  var sheet  = getOrCreate(SITE_ISSUES_TAB, writeSiteIssuesHeaders);
  var asSheet= db().getSheetByName(ASSIGN_TAB);
  var taskIds= [];

  issues.forEach(function(iss, idx) {
    var issId  = nextId(sheet, 'ISS-');
    var taskId = '';

    // Auto-create task in TASK_ASSIGNMENTS for all issue types
    if (asSheet && iss.assignedTo && iss.description) {
      taskId = createIssueTask(iss, project, date, onTrack);
    }

    sheet.appendRow([
      issId,
      subId,
      date,
      project,
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
  // Only create tasks for Design issues
  if ((iss.issueType||'') !== 'Design') {
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

  var taskType = 'Site Issue — Design';
  var notes    = iss.description + ' [From DPER: ' + project + ', ' + date + ']';

  sheet.appendRow([
    newId,                    // A TaskID
    '',                       // B ProjectID
    project,                  // C ProjectName
    assignee,                  // D AssignedTo (resolved from col F lead)
    taskType,                 // E Stage
    '',                       // F Discipline
    1,                        // G BasePts
    1,                        // H Multiplier
    1,                        // I WeightedPts (1pt for site issue resolution effort)
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
    'DPER — ' + (iss.reportedBy||'Execution Lead'), // V AssignedBy
    priority,                 // W Priority
  ]);

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
      '',
      0, 1, 0,
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
      date       : String(r[2-1]||'').substring(0,10), // B=date
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
