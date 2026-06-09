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
- getDeepakOpenIssues(lead) — returns open issues filtered by lead name

## Do not change
- Apps Script URL (hardcoded in all HTML forms)
- Column indices in getAllTasks — any change breaks dashboard
- WEEK_DAYS = 6 (Mon–Sat)
- cellDate() function — handles Date objects from Google Sheets
