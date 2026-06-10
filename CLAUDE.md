# Ideaform Design Studio — Productivity System

## Who I am
- Studio: Ideaform Design Studio, Bhopal
- Owner: Siddharth Inani (sidinani14@gmail.com)
- Team: ~11 members

## Key URLs
- Apps Script: https://script.google.com/macros/s/AKfycbziJSWTVf1sRqi6rKjmgjrg0DomsVKIm2DDBZWq7oHp2eCHDJ0bz9svlm8Do3G1xkEgBw/exec
- Productivity Sheet ID: 1PH1nJoPmQWS9wixuhw9B7oo0jIkhJw13htzbzlBWffk
- Scorecard Sheet ID: 198sgwhnp2GY5KYyVITMZAPLEoiEuMrHIFUat5deG3vk
- GitHub Pages: sidinani14.github.io/DPR

## Files in this repo
- IDS_Script_final.js — Google Apps Script backend (deploy via Apps Script editor)
- IDS_Task_Dashboard_final.html → deployed at /DPR/dashboard
- IDS_Projects_Dashboard.html → deployed at /DPR/projects
- IDS_Weekly_Report_Final.html → deployed at /DPR/weeklyreport
- dper.html → deployed at /DPR/dper (Deepak's DPER form)
- IDS_DPR_Form_final.html → deployed at /DPR/dpr
- IDS_Task_Assignment_final.html → deployed at /DPR/assign
- IDS_Lead_Approval_final.html → deployed at /DPR/approve

## Sheet structure — TASK_ASSIGNMENTS (23 cols A–W)
- A: TaskID, B: ProjectID, C: ProjectName, D: AssignedTo
- E: Stage/TaskType, F: Discipline, G: BasePts, H: Multiplier, I: WeightedPts
- J: AssignedDate, K: Deadline, L: Area, M: DrawingName
- N: SelfStatus, O: SelfStatusDate, P: ActualCompletionDate
- Q: LeadApproved, R: ApprovedBy, S: ApprovalDate
- T: RevisionTag, U: Notes, V: AssignedBy, W: Priority

## Sheet structure — SITE_EXECUTION (22 cols A–V)
- A: SubmissionID, B: Date, C: Time, D: ProjectName, E: Execution Lead
- F: SiteVisitDone, G: CurrentStage, H: WorksToday, I: %PlanCompletion
- J: %OverallCompletion, K: WorkTomorrow, L: OnTrackStatus, M: DelayReason
- N: IdleTime (always No), O: IdleReason (always blank)
- P: MaterialsRequired, Q: MaterialDelays, R: ClientUpdated
- S: ClientConcerns, T: BlockingTomorrow, U: DecisionsPendingSiddharth, V: AdditionalRemarks

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
- Absent day deduction: −2.0 pts from Punctuality, −1.5 pts from Hours
- Done pts source of truth: SelfStatusDate (col O) only
- LeadApproved = Yes required for task to count as done
- Tasks Assigned column in weekly report shows 0 (assigned pts not tracked)
- DPR tab: DAILY_SUMMARY — Date(A) Time(B) Member(C) Email(D)
- DPER tab: SITE_EXECUTION — col E = Execution Lead name

## Deepak Soni scoring (separate page in weekly report)
- Site Visits: /20 (active sites visited ≥1× / total active sites)
- Client Communication: /10 (sites with clientUpdated=Yes / total active sites)
- Task Completion: /20 (tasks assigned TO Deepak + assigned BY Deepak to team, with AssignedDate this week, completed within same week)
- DPER Consistency: /15 (days with ≥1 submission / 6 — mirrors team DPR formula)
- Punctuality: /20 (same base as team, biometric)
- Hours: /15 (same base as team, biometric)
- Total: /100
- Active projects list: CONFIG tab col A, under "DEEPAK ACTIVE PROJECTS" header

## Weekly targets by role
- Junior Architect / Junior ID: 48 pts/week
- Technical Lead (Achal, Himanshu, Bhavesh): 72 pts/week
- Deepak Soni (Execution): 30 pts/week (visit-based)

## Important functions in Apps Script
- syncVisitSchedule() — main visit scheduler (runs Monday 8AM)
- setupMondayTrigger() — sets up Monday trigger (run once)
- fixShiftedTaskColumns() — fixes column shift data bug (run once if needed)
- getWeeklyStats(weekStart) — returns team weekly scores
- getDeepakWeeklyStats(weekStart) — returns Deepak's weekly scores
- getAllTasks() — returns all tasks for dashboard
- getIssuesByReporter(member) — returns SITE_ISSUES filtered by col O (Reported By)
- updateIssueStatus(issueId, status, targetDate) — updates SITE_ISSUES row
- submitAmanCRM(data) — writes AMAN_DAILY, LEADS, SITE_ISSUES from CRM form
- getRecentLeads(date) — returns LEADS rows for given date with 24hr contact pending

## CRM data routing — where each form section is stored
- AMAN_DAILY  → client communication + other activities ONLY
- LEADS       → new leads (LMS format) + 24hr follow-up updates
- BILLING     → bills / payments / follow-ups (one row per bill)
- SITE_ISSUES → CRM issues & design deliverables (col O = Reported By = Aman)
- FEEDBACK    → monthly client feedback (one row per project)
- TASK_ASSIGNMENTS → auto-tasks: design issues/deliverables (1 pt) +
  site-visit/meeting tasks per team attendee (default 1 pt, overwritten by hours in DPR)

## Sheet structure — AMAN_DAILY (12 cols A–L; client comm + other activities only)
- A: Submission ID (CRM-001…), B: Date, C: Time, D: Member
- E: Client Contacts (JSON)
- F: Vendor Coordination (Yes/No), G: Vendor Entries (JSON [{project,notes}])
- H: Site Issues Addressed (Yes/No), I: Site Issue Entries (JSON)
- J: TnCP Coordination (Yes/No), K: TnCP Entries (JSON)
- L: BNI Activity (Yes/No)
- NOTE: schema changed (now 12 cols) — delete any old AMAN_DAILY tab so
  getOrCreate rewrites the new header row.

## Lead → Project promotion
- When a lead reaches "Briefing Meeting Done" or further (Design Proposal Shared /
  Fee Proposal Shared / Lead Converted), submitAmanCRM adds a PROJECTS row:
  [NL-id, ClientName, 'New Lead', '', '', ''] — de-duplicated by name.
- Fires for new leads added at that stage AND for 24hr follow-ups that set that stage.
- Aman's weekly "active projects" (Client Connection Coverage denominator) =
  all PROJECTS rows with Status = Ongoing (NOT the 'New Lead' pipeline rows).

## Sheet structure — BILLING (9 cols)
- A: Bill ID (BILL-001…), B: Project, C: Bill Date, D: Bill Amount
- E: Amount Received, F: Received Date, G: Last Follow-up Date, H: Status, I: Submission ID
- writeBilling(): bills append new rows; payments attach to OLDEST unpaid bill of that
  project (Partial/Paid); follow-ups stamp Last Follow-up Date on unpaid bills of that project.

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

## Sheet structure — LEADS (matches LMS Google Sheet format)
- A: Lead ID (LEAD-001…), B: Client Name, C: Contact No., D: Referred By
- E: Validation Check (Valid/Invalid/Not checked)
- F: Lead Status, G: Contacted By, H: Lead Creation Date, I: Last Contacted
- J: Lead Manager, K: Remarks, L: Lost Reasons, M: 24hr Contact Done (Pending/Yes/No)

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
