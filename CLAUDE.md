# Ideaform Design Studio — Productivity System

## Visit planner (2026-08) — VISIT_PLANNER vs VISIT_SCHEDULE
- **VISIT_PLANNER** = the input you edit by hand: project, visit type,
  assignee, frequency, fixed day, last-visit date, Active Yes/No.
- **VISIT_SCHEDULE** = fully auto-generated output, rebuilt from scratch
  every `syncVisitSchedule()` run (daily 7 AM trigger) — next visit date,
  status, days overdue, linked Task ID. Col F (Next Visit Date) is the only
  user-editable manual override; everything else is overwritten each run.
  Don't delete this tab expecting it to "reset" anything meaningful — it's
  derived, not a source of truth, and just regenerates.
- `loadVisitHistory()` only counts **approved** (`LeadApproved==='Yes'`)
  Done visit tasks as "last visit" — briefly relaxed to also count Pending
  ones (so approval lag wouldn't blind the planner), reverted same day per
  explicit instruction: approve visits promptly, don't rely on unapproved
  self-reports for scheduling.
- `calcNextVisitDate`'s Priority 1 treats ANY existing open task for a
  project+type+assignee as "already scheduled" and just echoes its
  deadline — meaning stale open visit tasks silently block a corrected
  VISIT_PLANNER cadence from ever taking effect until those stale tasks are
  cleared. `cleanupPendingVisitTasksAndResync()` (one-off, run manually from
  the Apps Script editor, not wired to any route) handles this: deletes
  only `Not Started`/blank-status visit tasks (never anything with
  progress or a resolution), then re-runs the full sync.
- `notifyMissedVisits()` emails Siddharth/Astha a digest whenever
  `syncVisitSchedule` finds a missed visit (>1 full frequency cycle
  overdue) — fully automatic via the daily trigger, no email on a clean day.

## Reliability fixes (2026-08) — read before touching submission paths
Root-caused two backend bugs behind reported "sometimes access denied" +
"task submissions not reflected, no score":
1. `withLock()` used to proceed WITHOUT the lock if `waitLock()` timed out —
   under real contention (several people submitting DPR near the 7:30
   deadline) this ran unprotected concurrent writes against the same
   TASK_ASSIGNMENTS rows. Now throws instead — every `withLock` caller must
   expect it can throw, and that error should reach the client, not be
   swallowed.
2. `verifyIdToken()` cached ANY non-200 tokeninfo response as a hard denial
   for 5 min — a transient network blip got remembered as "access denied"
   for that token for 5 min regardless of real authorization. Only genuine
   negative verdicts (bad aud/unverified email/not on allowlist) are cached now.
3. The default DPR submission handler ran its whole write batch unlocked with
   each step's errors only `Logger.log`'d — one bad item in a forEach loop
   silently aborted the rest of that batch while the client was told 'ok'
   regardless. Now one locked, atomic batch, per-item try/catch, real errors
   returned as `{status:'partial', errors:[...]}`. Same pattern applied to
   `updateTaskStatusesFromDPR` (now returns `{updated, errors}`, not void) and
   `submitDPER`/`submitAmanCRM` are now wrapped in `withLock` too.
**Any new score-affecting write path must**: (a) go through `withLock`, (b)
isolate per-item failures inside a loop with try/catch rather than letting
one throw kill the batch, (c) return real errors instead of a blanket 'ok'.
**Any new frontend submit handler must** parse the JSON response and surface
`status:'error'`/`'partial'` to the user (a visible warning, not a silent
`.catch(()=>{})`) — dpr.html/DPER.html/CRM.html/social.html all do this now
for their score-relevant POSTs; match that pattern for new forms.

**Follow-up same day — backdating late-filed work (Report Date)**: DPR had no
way to date a self-logged task or field-work entry (site visit/meeting) other
than "today" — filing a day late silently misattributed the points to the
wrong week. DPER/CRM already had an editable `#f-date` "Report Date" field;
DPR now has one too (Section 1, defaults today), threaded into Section 3
(self-logged Done tasks) and Section 4 (Field Work). But the deeper bug was in
`createDoneTask` itself: it only honored a backdated `data.date` for the
AssignedDate column — Deadline/SelfStatusDate/ActualCompletionDate were
hardcoded to `today` regardless, and **SelfStatusDate is the sole source of
truth for which week a Done task's points land in**. That silently defeated
DPER's and CRM's report-date fields too, not just DPR's missing one. Fixed in
both the new-row and matched-existing-row branches — any caller passing
`data.date` now gets it applied consistently across all date columns.

**Follow-up same day — no backdating pre-assigned task completion**: the
"Actual Completion Date" picker shown when marking a pre-assigned task Done
in dpr.html was removed by design — it let people backdate to dodge a
delayed-task flag, and unlike site visits (verifiable via visit log +
attendance) there's no independent way to check when a design task actually
finished. Completion date is now always the real submission date (backend
already fell back to `now` when no override is sent — zero backend change
needed, just stopped sending a value). The Report Date field above is a
*different* thing — it's for self-logged NEW work/field-work only, kept
editable because those ARE independently verifiable.

**Follow-up same day — visit planner blind to pending-approval visits**:
`loadVisitHistory()` (feeds VISIT_SCHEDULE/PROJECT_HEALTH/missed-visit
flagging) required `LeadApproved==='Yes'` to count a visit — but every visit
logged via DPR/DPER/CRM is created `Pending`, not auto-approved. Any lag
approving tasks meant the planner silently didn't know a visit had happened.
Fixed to count any self-reported Done visit unless explicitly `Rejected`.

**Follow-up same day — "Unexpected token '<'" on dashboard/form loads**: this
is `res.json()` choking on an HTML response. Every path in this script goes
through `respond()`, which only ever emits valid JSON — HTML can only mean
Google's platform served an error page *before* the script ran, i.e. Apps
Script's execution quota. The account is a **personal Gmail account, not
Workspace** — a much lower daily quota tier. Mitigated (not eliminated — a
hard account-tier ceiling can't be fixed from code) as **@278**: `getAllTasks`/
`getLists` (whole-team, non-user-specific) cache their response for 30s via
`cachedSafeRespond()` so concurrent dashboard loads collapse into one sheet
read instead of N; `verifyIdToken`'s positive cache stretched 30min→55min to
roughly halve tokeninfo round-trips. If this recurs, the only real diagnostic
is the account owner checking Apps Script editor → Executions for a
quota-exceeded failure — don't keep guessing in code without that signal.

## Who I am
- Studio: Ideaform Design Studio, Bhopal
- Owner: Siddharth Inani (sidinani14@gmail.com)
- Team: ~11 members

## Key URLs
- Apps Script (PROD deployment all frontends use): https://script.google.com/macros/s/AKfycbyyHnaAJ35ry0ot5HrWfz-IGaiM5algp4LLbTU8YtUwb9cvGhbCXkQZOYgMnY2CQ-Y5vA/exec
  (scriptId 18W1_ciQs1QW1EeItd-r55Y7mxs_rZR28SwNHwSdk1OYlzGZxgW-P85xv; redeploy with `clasp deploy -i <that id>`. Hit the 200-version limit before — delete old versions in the editor if deploy fails. See memory dpr-apps-script-deployment.)
- Productivity Sheet ID: 1PH1nJoPmQWS9wixuhw9B7oo0jIkhJw13htzbzlBWffk
- Scorecard Sheet ID: 198sgwhnp2GY5KYyVITMZAPLEoiEuMrHIFUat5deG3vk
- Live site (custom domain): https://team.ideaformdesignstudio.com (served from the GitHub Pages repo ROOT — files are dashboard.html/projects.html/tasks.html/etc., NOT under /DPR/). `.nojekyll` present so Pages serves statically.

## Files in this repo (current names — served at repo root)
- Code.js — Apps Script backend (push via `clasp push`; browser JS excluded by .claspignore)
- index.html — workspace home (tile grid, TOOLS array)
- dashboard.html · projects.html · myprojects.html (projects you lead) · weeklyreport.html · weeklydigest.html
- dpr.html · DPER.html · CRM.html · meetlog.html · approval.html · leads.html · 3m.html
- tasks.html — **Bulk Task Planning** (the "Plan Tasks" form; was assign.html, now deleted)
- searchselect.js · auth.js — shared browser components (not pushed to Apps Script)

## Sheet structure — TASK_ASSIGNMENTS (23 cols A–W)
- A: TaskID, B: ProjectID, C: ProjectName, D: AssignedTo
- E: Stage/TaskType, F: Discipline, G: BasePts, H: Multiplier, I: WeightedPts
- J: AssignedDate, K: Deadline, L: Description (was Area), M: DrawingName (legacy; new tasks blank)
- N: SelfStatus, O: SelfStatusDate, P: ActualCompletionDate
- Q: LeadApproved, R: ApprovedBy, S: ApprovalDate
- T: RevisionTag, U: Notes, V: AssignedBy, W: Priority (no longer set/shown — deadline is the priority)
- Writes go through assignTasks (batch setValues) / bulkAssignTasks, all wrapped in
  withLock (LockService) so concurrent/bulk submits don't drop rows.

## Field work / attendance capture (2026-07-11) — Part A
- DPR form (dpr.html) **Section 4 "Field work today"**: flat, repeatable block —
  each row = type (Site Visit / Meeting / **Material Selection**) + project +
  **start time + end time** + notes. Duration = end−start drives visit points
  (Site Visit & Material Selection ×2/hr, Meeting ×1/hr); each row still creates a
  Done visit task (createDoneTask, hours-based pts) so scoring is unchanged.
- Replaced (removed) the old per-project "Site visit or client meeting" toggle
  (p-vtype/p-vdur/p-vnotes + toggleVisit) to avoid double-capture.
- Material Selection added to isVisitTask/calcVisitPts (Code.js, ×2) + isVisitType/
  isVisitRate2 (dpr.html) + weekly project-digest reverse-rate (wp/2).
- New **FIELD_WORK** sheet (writeFieldWork in doPost, from data['Field Work']):
  A FW ID · B Date · C Member · D Email · E Type · F Project · G Start · H End ·
  I Engaged Hrs · J Notes · K Source · L Approved(Pending). Feeds the upcoming
  attendance import (Part B): per member+date, first start→last end = the field-day
  window (hours score); engaged sum = points. Approved gates it into scores.
- Assigned pre-scheduled visit-task hours (atc-visit-hrs) + unplanned visit-task
  rows (t-visit-hrs) still use plain hours and do NOT write FIELD_WORK — only Sec 4.

## Recent changes (2026-07)
- tasks.html = Bulk Task Planning (assign.html deleted). Auto-saves a draft to
  localStorage + server (PLAN_DRAFTS tab, savePlanDraft/getPlanDraft) so a failed
  submit never loses the form; "Restore/Discard" banner on reopen.
- Weekly targets now read from TEAM tab col C (WeeklyTarget) via getLists.targets;
  dashboard merges them over its hardcoded defaults (Deepak = 72, not 30).
- DPER: Site Visit/Meeting + hours per project (points via calcVisitPts, DPR rate);
  "Unplanned Work" (self-logged Done tasks); per-project "Tasks to Assign" + "Your
  Open Tasks" (status update via updateTaskStatuses). Issues drop the Design type
  and priority. Weekly report has a "How Deepak Spent His Week" activity panel.
- DPER + CRM both show "Your Open Tasks" (getOpenTasksForMember + updateTaskStatuses).
- Project Digest: a visit counts only when SiteVisitDone=Yes (col F); shows visit
  hours + a "Site progress" narrative (SITE_EXECUTION col H); visit/meeting tasks
  de-duped out of "Tasks completed".
- searchselect.js: type-ahead — focus a dropdown and start typing to filter.
- Site-issue priority removed from DPER/CRM forms + open-issues display (sorted by
  target date now); task priority removed from Plan Tasks + DPER + dashboard.

## Sheet structure — SITE_EXECUTION (22 cols A–V)
- A: SubmissionID, B: Date, C: Time, D: ProjectName, E: Execution Lead
- F: SiteVisitDone, G: CurrentStage, H: WorksToday, I: %PlanCompletion
- J: %OverallCompletion, K: WorkTomorrow, L: OnTrackStatus, M: DelayReason
- N: IdleTime (always No), O: IdleReason (always blank)
- P: MaterialsRequired, Q: MaterialDelays, R: ClientUpdated
- S: ClientConcerns, T: BlockingTomorrow, U: DecisionsPendingSiddharth, V: AdditionalRemarks

## EPIC G/H/I/J (built 2026-06-25)
- TASK_ASSIGNMENTS appended cols: X(24) prior status · Y(25) original deadline ·
  Z(26) disposition (…/Parked/Parked (Stalled)) · AA(27) park reason (free text).
- Park: `parkTask`/`unparkTask` (POST) — Siddharth-only (DIRECTOR_EMAILS, verified
  from idToken). Direct park from dashboard, no prior block needed.
- Stalled stage (PROJECTS col C = "Stalled"): open tasks auto-park as
  "Parked (Stalled)"; un-stalling revives only those. `reconcileStalledParks()`
  runs in syncVisitSchedule + callable (action=reconcileStalled). New tasks on a
  stalled project park on creation via parkRowIfStalled().
- Connections: DPR + DPER post `logConnections {member, connections[]}` →
  appendConnections() writes CRM_LOG (Client Connection rows, author=submitter).
  Visibility-only (not scored). Surfaces on projects dashboard (read path existed).
- Billables: DPR posts `submitBillables {member, billables[{project,discipline,stage}]}`
  → BILL_REQUESTS tab (Pending). Approval form panel (getBillRequests +
  disposeBillRequest approve/reject); approve → 0-pt no-deadline "Raise Bill" task
  for Aman (never scored as delayed). Payment stages read from CONFIG cols V(Discipline)/
  W(Milestone) under a "PAYMENT STAGES" marker → readConfig.paymentStages/paymentDisciplines.

## Senior-staff visit/meeting multiplier (built 2026-08)
- Design task points already reflect a higher target for senior staff; visit/
  meeting points (`calcVisitPts`, hours × rate) never did — same rate for
  everyone regardless of target. Fixed for anyone with a daily target > 8 pts
  (weeklyTarget/6): `seniorityMultiplier(member)` reads their live TEAM-tab
  WeeklyTarget and returns `dailyTarget/8` (e.g. 72/wk → 12/day → 1.5x;
  60/wk → 10/day → 1.25x); everyone else gets 1.0 (no change). Fully
  data-driven off the TEAM tab, not a hardcoded name list — tracks whoever
  actually has an elevated target.
- `getProjectMultiplier(projectNameOrId)` — project discipline multiplier
  lookup, only stacks on top of the seniority multiplier (i.e. only applies
  to the senior cohort) — everyone else's visits stay pure hours × rate,
  unaffected. `calcVisitPts(taskType, hours, member, projectMult)` — the last
  two args are optional and default to no-op, so any call site that doesn't
  pass them behaves exactly as before.
- Wired into the 3 live scoring paths: `createDoneTask` (DPR Field Work
  Section 4, CRM visit/meeting tasks, Social Media documentation-as-site-visit),
  `updateTaskStatusesFromDPR` (marking a pre-assigned visit/meeting task Done —
  project multiplier read from the row's own col F, already stored at
  assignment), and the DPER visit-logging path (`getProjectMultiplier` lookup).
  Deliberately NOT applied in `backfillDPRTasks`/`backfillFieldWorkPoints` —
  those are one-off historical-repair tools, this is a going-forward change.

## EPIC K — Site Visit / Meeting log (built 2026-06-25)
- meetlog.html (deploy at /DPR/meetlog) — unified form, replaces the two Google Forms.
- Tabs: MEETING_LOG (20 cols: id,date,time,type,project,loggedBy,team,clients,purpose,
  bodyRaw,bodyPolished,duration,driveFolder,photoIds,videoLinks,leadApproved,approvedBy,
  approvalDate,whatChanged,reportPdfId) + DECISION_LOG (id,logId,project,date,category,
  owner,task,deadline,status[Open|Done|Revised|Carried]).
- Flow: submitMeetingLog → aiPolishLog (Claude Haiku 4.5, UrlFetchApp to
  api.anthropic.com/v1/messages, key = Script Property ANTHROPIC_API_KEY; tone/clarity
  only, graceful fallback) → IDS items become tasks, log becomes a CRM_LOG connection →
  Pending → approval.html panel (edit polished, approve) → generateProjectReportPDF
  (cumulative, newest-first, cover index, photos inlined base64, saved to Drive "IDS Logs").
- uploadMeetingPhoto: base64 → per-project/date Drive folder. Routes: submitMeetingLog,
  uploadMeetingPhoto, getMeetingApprovals, approveMeetingLog, getMeetingTimeline.
- DPER site visit auto-publishes a Site Visit Log (skipTasks=true; DPER already tasks issues).
- One-time: paste ANTHROPIC_API_KEY in Script Properties; run authorizeEpicK() for
  Drive + Anthropic scopes.

## Scoring framework (current)
- Output: /50 (approvedPts / weeklyTarget × 50)
- DPR Consistency: /15 (dprDays / 6 × 15)
- Punctuality: /20 (biometric, scaled from /15 base)
- Hours: /15 (biometric, scaled from /10 base)
- Total: /100
- Delivery and Quality removed

## Key rules
- Always run: node --check filename.js or node --check filename.html after edits
- Week = Monday to Saturday (6 days)
- Punctuality threshold: 09:10 for all, 10:40 for Achal Rathore only
- Absent day deduction: −2.0 pts from Punctuality, −1.5 pts from Hours.
  DPR/Punctuality/Hours are scored on the fixed 6-day week for everyone
  (briefly changed to present-days-only 2026-07-20, reverted same day) —
  EXCEPT a day marked approved leave (BIO att status 'L', via `leaveCount()`
  in weeklyreport.html) is excluded from both the absence penalty and the
  DPR denominator; unexplained absence still costs the full penalty. First
  applied to Aman's 4-day pre-approved leave, 13-18 Jul week.
- Done pts source of truth: SelfStatusDate (col O) only
- LeadApproved = Yes required for task to count as done
- Tasks Assigned column in weekly report shows 0 (assigned pts not tracked)
- DPR tab: DAILY_SUMMARY — Date(A) Time(B) Member(C) Email(D)
- DPER tab: SITE_EXECUTION — col E = Execution Lead name

## Aman Raghuwanshi — CRM scoring (separate page in weekly report)
- getAmanWeeklyStats(weekStart) — routes: getAmanWeeklyStats (GET+POST)
- Output /50 (components are N/A when there's nothing to measure → excluded &
  output rescaled across active ones: output = activeScore/activeMax × 50):
  - Client Connection Coverage /20 = ongoing projects connected ÷ total ongoing × 20
    (connection = client contact OR meeting/visit task OR feedback this week; N/A if no ongoing)
  - Lead Management /15 = open leads worked this week ÷ leads needing attention × 15
    (worked = LEADS Last-Contacted col I in week; base = worked + still-open-untouched;
     open = status Not Contacted/Contacted over call; N/A if base = 0 — no free marks)
  - Revenue Collection /15 = min(totalCollected ÷ totalBilled, 0.70) ÷ 0.70 × 15
    (CUMULATIVE across all BILLING rows — outstanding old bills lower it; N/A if nothing billed)
  - visitsDone/meetingsDone derived from Client Contacts type (display only)
- DPR Consistency /15 = AMAN_DAILY submission days ÷ 6 × 15
- Punctuality /20 + Hours /15 = biometric from DAILY_SUMMARY (same formulas as Deepak/team, Thr 09:10)
- weeklyreport.html: buildAmanPage(d), fetched via getAmanWeeklyStats, pushed into team[] with _isAman

## Deepak Soni scoring (separate page in weekly report) — rewritten 2026-08
Current formula (this doc had drifted badly stale before this rewrite —
trust this section and getDeepakWeeklyStats/Code.js over anything else):
- **Output: /35** — `deepakOwnCompletedPts ÷ weeklyTarget × 35`, capped at 35.
  Same points-over-target mechanism as every team member's Output score.
  Replaced the old Site Visits/17.5 + Task Completion/17.5 coverage-ratio
  split (2026-08) — those measured "did he show up" / "did he finish what
  was assigned", not how much he actually produced. `deepakOwnCompletedPts`
  is ONLY points from tasks assigned TO Deepak himself (visits, meetings,
  any other work) — work he delegates to a teammate is that teammate's
  output, not double-counted here.
- **Client Communication: /10** (sites with clientUpdated=Yes / total active sites — unchanged)
- **Issue Reporting: /5** (site problems he flags for design/CRM ÷ target of 3/week — added 2026-08, was previously zero credit anywhere)
- **DPER Consistency: /15** (days with ≥1 submission / 6 — mirrors team DPR formula)
- **Punctuality: /10**, **Hours: /7** (rescaled from /15,/10 to fund Client Satisfaction below)
- **Client Satisfaction: /8** (FEEDBACK sheet Overall Satisfaction, his active sites this week — added 2026-08, fills the Planning slot team members get, since his work is reactive site execution, not pre-assigned deadline tasks)
- **Reliability: /10** (−1/late task, −2/Work Not Done, −5/unexplained absence)
- Total: /100 (35+10+5+15+10+7+8+10=100)
- **Weekly target: 72 pts/week** (TEAM tab col C, read live — NOT the old
  hardcoded 30 this doc used to say). 72/6=12/day, which is what makes him
  qualify for the senior visit/meeting point multiplier (`seniorityMultiplier`,
  dailyTarget>8 → ×1.5) — his visit/meeting task points already carry that
  before they ever reach the Output sum above.
- Active projects list: CONFIG tab, under "DEEPAK ACTIVE PROJECTS" header (lives in
  col T — getDeepakWeeklyStats scans ALL columns for the header, not just col A)
- weeklyreport.html: `buildDeepakPage(d)` (his dedicated report page — pCard
  calls use max 35 for Output, not the generic 50 team members' page uses)
  and the Overview-page team[] row builder both read `scores.s_output`
  directly from the backend now — don't recombine from s_visit/s_tasks,
  those fields no longer exist in the backend response.

## Weekly targets by role
- Junior Architect / Junior ID: 48 pts/week
- Technical Lead (Achal, Himanshu, Bhavesh): 72 pts/week
- Deepak Soni (Execution): 72 pts/week (corrected 2026-08 — was documented
  as 30 here, stale; live TEAM tab value is 72, same senior tier as Achal/
  Himanshu, which is what qualifies him for the visit/meeting point
  multiplier — see the Deepak scoring section below)
- Kritika Kaushik (Jr. ID + Social Media Manager): 48 pts/week combined — design
  task points AND Social Media Log points both feed the same Output pool, target
  not raised (2026-08 decision; revisit after a few weeks of real data if content
  work crowds out design output or vice versa).

## Social Media Manager scoring (Kritika Kaushik, built 2026-08)
- social.html — standalone form filled alongside DPR, only on days she has
  content activity (not a mandatory daily form, unlike DPR). Backend routes
  submitSocialMediaLog/getSocialMediaLog write SOCIAL_MEDIA_LOG (14 cols:
  LogID,Date,Member,Email,Type[Content|Documentation|Idea|Blocker],Platform,
  ContentType,Project,Title/Caption,Link,Status,Hours,Notes,LinkedTaskID).
- Every project field (content + documentation) offers an "Other" option
  revealing free text, for team-culture/studio content, past-project
  documentation, or a site not yet in PROJECTS.
- **Scoring (2026-08 decision, all explicit user calls, not defaults):**
  - Sent for approval, never auto-approved — same approval.html Tasks queue
    as everything else (LeadApproved 'Pending', generic getPendingTasks already
    picks it up, no new approval-flow code needed).
  - Content: only Status='Published' items score, flat points by content type
    via SM_CONTENT_PTS (Reel 2.0, Video 2.5, Carousel 1.5, Blog 2.0, Post 1.0,
    Story 0.5, else 1.0) — deliberately NOT multiplied by the project's
    discipline multiplier (logSocialMediaContentTask always writes mult=1).
    Scheduled/Draft stay visible-only in the log, no points.
  - Documentation shoots are NOT scored on their own — they're a site visit,
    reusing createDoneTask with taskType 'Site Visit' so it scores hours×2/hr
    through the exact same pipeline as every other logged visit (calcVisitPts).
    Requires an Hours field on the form.
  - createDoneTask accepts an optional data.notes override (falls back to its
    old hardcoded 'Unplanned task — self logged via DPR' string) purely so the
    Social Media source is identifiable in the approval queue/Notes column —
    every other caller is unaffected.

## Important functions in Apps Script
- syncVisitSchedule() — main visit scheduler (runs Monday 8AM)
- setupMondayTrigger() — sets up Monday trigger (run once)
- getWeeklyStats(weekStart) — returns team weekly scores
- getDeepakWeeklyStats(weekStart) — returns Deepak's weekly scores
- getAllTasks() — returns all tasks for dashboard
- getIssuesByReporter(member) — returns SITE_ISSUES filtered by col O (Reported By)
- updateIssueStatus(issueId, status, targetDate) — updates SITE_ISSUES row
- submitAmanCRM(data) — writes AMAN_DAILY, LEADS, SITE_ISSUES from CRM form
- getRecentLeads(date) — returns LEADS rows for given date with 24hr contact pending

## CRM data routing — where each form section is stored
- AMAN_DAILY  → daily INDEX only: subId, date, time, member + 4 activity Yes/No flags
- CRM_LOG → one row per client connection AND per activity (combined). Cols:
  Log ID, Submission ID, Date, Member, Category (Client Connection | Other Activity),
  Type/Activity, Project, Notes. getAmanWeeklyStats filters Category='Client Connection'.
  (Migrations: migrateAmanDaily splits old JSON → AMAN_DAILY+CRM_LOG; mergeCrmLog
   folds old CLIENT_CONNECTIONS+ACTIVITIES tabs into CRM_LOG. Both non-destructive/idempotent.)
- LEADS       → new leads (LMS format) + 24hr follow-up updates
- BILLING     → bills / payments / follow-ups (one row per bill)
- SITE_ISSUES → CRM issues & design deliverables (col O = Reported By = Aman)
- FEEDBACK    → monthly client feedback (one row per project)
- TASK_ASSIGNMENTS → auto-tasks: design issues/deliverables (1 pt) +
  site-visit/meeting tasks per team attendee (created at 0 pts; scored off hours
  the attendee logs in their DPR — Site Visit ×2/hr, Meeting ×1/hr).
  isVisitTask/calcVisitPts (Code.js) + isVisitType/isVisitRate2 (index.html) now
  recognise the plain 'Site Visit' / 'Meeting' types, not just the canonical ones.

## Sheet structure — AMAN_DAILY (8 cols A–H; daily index only)
- A: Submission ID (CRM-001…), B: Date, C: Time, D: Member
- E: Vendor Coordination, F: Site Issues Addressed, G: TnCP Coordination, H: BNI Activity (Yes/No)
- Detail rows now live in CLIENT_CONNECTIONS + ACTIVITIES (row per entry, filterable).
- NOTE: schema changed (now 8 cols, JSON removed) — delete any old AMAN_DAILY tab so headers rewrite.

## Sheet structure — CRM_LOG (8 cols; connections + activities combined)
- A: Log ID (LOG-001…), B: Submission ID, C: Date, D: Member
- E: Category (Client Connection | Other Activity)
- F: Type / Activity (Call/WhatsApp/Meeting/Site Visit/Email  OR  Vendor/Site Issues/TnCP/BNI)
- G: Project, H: Notes / Discussion

## Lead → Project promotion
- When a lead reaches "Briefing Meeting Done" or further (Design Proposal Shared /
  Fee Proposal Shared / Lead Converted), submitAmanCRM adds a PROJECTS row:
  [NL-id, ClientName, 'New Lead', '', '', ''] — de-duplicated by name.
- Fires for new leads added at that stage AND for 24hr follow-ups that set that stage.
- Aman's weekly "active projects" (Client Connection Coverage denominator) =
  all PROJECTS rows with Status = Ongoing (NOT the 'New Lead' pipeline rows).

## Sheet structure — BILLING (10 cols A–J)
- A: Bill ID (BILL-001…), B: Invoice No., C: Project, D: Bill Date, E: Bill Amount
- F: Amount Received, G: Received Date, H: Last Follow-up Date, I: Status, J: Submission ID
- CRM form asks Invoice No. on each bill AND each payment.
- writeBilling(): bills append new rows; payments match by Invoice No. FIRST, else
  OLDEST unpaid bill of that project (Partial/Paid); follow-ups stamp Last Follow-up Date.
- NOTE: schema changed (invoice col added) — delete any old BILLING tab so new headers write.

## Sheet structure — FEEDBACK (19 cols, mirrors Monthly Feedback Google Form)
- A: Feedback ID (FB-001…), B: Submission ID (CRM-…), C: Date Recorded, D: Recorded By
- E: Project/Client, F: Date of Visit/Meeting
- G: Overall Satisfaction (1-10), H: Agenda Communicated (Yes/No)
- I: Design & Functionality (1-5), J: Communication (1-5), K: Problem Resolution (1-5)
- L: Responsiveness (1-5), M: Quality of Work (1-5), N: Professionalism (1-5)
- O: Meeting On Time (Yes/No), P: Recommend/NPS (1-10)
- Q: Additional Comments, R: Appraisal of Person, S: Referrals
- CRM Section 8 "Monthly Client Feedback": Yes/No toggle → repeatable per-project blocks.
  Written to FEEDBACK sheet only (one row per project) — NOT in AMAN_DAILY.
- Agenda block attendees = Team Attendees (chips of TEAM_MEMBERS → each gets a
  meeting/visit task) + Other Attendees (free text, message only).
- Other Activities = Vendor/SiteIssues/TnCP each Yes → repeatable {project, notes};
  BNI = plain Yes/No (no detail).
- Finance = 3 collapsible Yes/No: Bills Raised (repeatable {project, amount}),
  Payment Received (repeatable {project, amount}), Follow-ups Done (project multi-select only).
- Section order: Open Issues, Report Date, Client Connections, Issues/Deliverables,
  Tomorrow's Meetings, (WhatsApp agendas), Finance, Other Activities, Lead Management, Monthly Feedback.

## Sheet structure — LEADS (matches LMS Google Sheet format; 16 cols A–P)
- A: Lead ID (LEAD-001…), B: Client Name, C: Contact No., D: Referred By
- E: Validation Check (Valid/Invalid/Not checked)
- F: Lead Status, G: Contacted By, H: Lead Creation Date, I: Last Contacted
- J: Lead Manager, K: Remarks, L: Lost Reasons, M: 24hr Contact Done
- N: Lead Source (BNI/Client/Vendor/Lead Manager/Business Associate/Friends & Family/Walk-in/Others)
- O: Briefing Date, P: Proposal Date — stamped on stage transition (SLA clocks)
- migrateLeadsColumns(sheet) adds N/O/P to existing tabs (non-destructive).
- Lost leads MUST carry a Lost Reason (8 standard: High Fee, Project on Hold,
  Portfolio Shortage, Design Proposal Not at Par, Looking for Different Plot,
  Low Scale Project, No Response/Lost Contact, Other) — enforced in CRM form.
- SLAs: contact within 24h (col M), proposal within 48h of Briefing Date (O→P).

## Leads analytics & dashboard
- getLeadsAnalytics(month='YYYY-MM') → funnel (contacted/briefing/proposal/converted/
  lost/invalid), conversionRate (converted/valid), bySource, byManager, lostByReason,
  sla24 {met/missed/pending}, sla48 {met/late/breaching/pending}, needsAttention list
  (open leads breaching 24h/48h, all months), 6-month trend. Routes: GET+POST.
- leads.html → deployed at /DPR/leads (month selector, KPIs, SLA radar, funnel,
  source/reason/manager breakdowns, trend). Uses the Forms deployment. (Pending/Yes/No)

## CRM form — recent additions (Aman feedback)
- Add New Project: Section 4 captures name + type(discipline from getConfig) + stage + client;
  createCrmProjects() writes PROJECTS rows [CP-id, name, stage, type, '', '', client] (col G = Client).
- Issue actions now include "Blocked" (status 'Blocked'); getIssuesByReporter excludes
  Blocked/Void/Invalid/Cancelled from the open list (erases invalid issues).
- Design Deliverables (kind='Deliverable') go to TASK_ASSIGNMENTS only — NOT SITE_ISSUES,
  so they don't appear in the CRM open-issues panel. Only kind='Issue' writes to SITE_ISSUES.
- PROJECTS col G = Client (added) — foundation for client-level monthly feedback KPI.

## SITE_ISSUES — col O added: Reported By (Deepak Soni / Aman Raghuwanshi)
- Used by getIssuesByReporter(member, useFallback) to filter issues per person's form
- getDeepakIssues → useFallback=true (matches blank col O via Assigned To, for legacy rows)
- getAmanIssues → useFallback=false (STRICT col O match — CRM never shows DPER's issues)
- CRM new issues carry their own per-issue project; writeSiteIssues uses iss.project || project
- CRM "Design Deliverable" items set iss.kind='Deliverable' → createIssueTask makes a design task

## Do not change
- Apps Script URL (hardcoded in all HTML forms)
- Column indices in getAllTasks — any change breaks dashboard
- WEEK_DAYS = 6 (Mon–Sat)
- cellDate() function — handles Date objects from Google Sheets
