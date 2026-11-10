# CF Tasks V3 — PreInspection Standard 2026-11-10-r1

**Status:** Approved business-format contract; **not** a deployed Apps Script release.
**Effective business standard:** 2026-11-10 (America/Toronto).
**Scope:** PreInspection Calendar display, Striven Task managed fields, existing notification template, safety and verification.
**Precedence:** This dated contract governs the subjects below over older presentations and assistant drafts; all unrelated protections from `CURRENT_DECISION_REGISTER.md`, `FLOW.md`, and `REQUESTED_BY_POLICY.md` remain in force. Only an explicit subsequent user decision may change the presentation/template.

## 1. Frozen exact display contracts

- Canonical PreInspection Calendar title: `<Customer Number> - <Customer Name> - (<AAA>) BBB-CCCC`. Do not add `Preinspect -`, SO number, extra punctuation or trailing IDs.
- Canonical Striven Task name: `<Customer Name> - <Verified Street, City> - (<AAA>) BBB-CCCC`. Do not insert Customer Number, task ID or vertical prefix.
- The Customer Number is a business number; do not use it as the internal Customer ID unless a real relationship lookup proves the IDs happen to be equal.
- Use the verified Customer name, contactable 10-digit phone, and Customer-owned inspection Location. Do not guess missing values or normalize a wrong Location.
- Preserve the exact original Calendar title as the first description line when the title changes; retain the authored visit description and meaningful source line breaks. Do not insert unapproved labels, separators or page breaks.
- Preserve a single existing working `-----Striven Links-----` block. Repair only absent, wrong or non-clickable entries. Include the verified Sales Orders page link (using the actual internal Customer ID) and verified `Task #<ID> – <actual Striven Task Name>` link as actual clickable Google Calendar hyperlinks. Never invent URLs or render literal Markdown-link text.
- The authoritative CF Preinspects event and every *required* Calendar copy must have the same business-visible title, description, location, appointment start/end and valid links. Read every required copy after any Calendar mutation.
- Preserve original attendees, attachments and human-authored notes. Do not add an unrelated inspector to another inspector's personal Calendar. Use verified event ownership to determine Stephen/Jay/Karen assignment and required personal-calendar participation; CF Preinspects is the shared business index. Do not conflate organizer, creator, assignee and Requested By.

### Regression fixture (fictional; no customer information)

- Calendar: `12345 - Sample Customer - (416) 555-0100`
- Task: `Sample Customer - 10 Example St, Toronto - (416) 555-0100`
- Task hyperlink label: `Task #<real task ID> – <actual task title>`

The established previously verified appointment can be used as a *read-only private regression* but must not be copied into this public repository.

## 2. Protected and managed Striven Task fields

| Field | Required treatment |
| --- | --- |
| Type | `105` — Pre Inspection |
| Name | Exact canonical Task-name format above |
| Customer | Verified internal Customer ID |
| Location | Verified Customer-owned Location ID |
| Start / Due | Canonical event start/end in Toronto local time |
| Assigned pool | Pool `8` — Preinspection Pool |
| Assigned employee | Proven inspector from the appointment's owning Calendar; preserve unrelated verified manual assignments |
| Requested By | Existing controlling PreInspection policy: exact creator/organizer email matched uniquely to Striven Employee; `Type: employee`. Do not silently change to Contact. Where source identity is ambiguous, stop this mutation and record REVIEW. |
| Task status | New appointment task OPEN; respect completed/canceled history and the lifecycle gate |
| Task Description | Exactly blank at CREATE; do not mirror Calendar notes; preserve inspector edits afterward |
| Sales Order | Not required or automatically attached on PreInspection CREATE/RECREATE |
| `InfoCustomFields` at CREATE | **None** |
| Field `854` — Install Notes | **DO_NOT_MANAGE**: automation must not prefill, copy, write, patch, clear, synchronize, gate completion on or normalize it |
| Inspector fields `852`, `853`, `860` and related assessments | Never prefill, clear or overwrite; remain inspector-owned |
| Customer Contact | May support identity evidence but is not automatically Requested By under the existing production contract |

A write planner and outbound payload validation MUST enforce the protected fields. Config constants alone are insufficient verification. Existing correct fields are left unchanged.

## 3. Task identity and lifecycle guards

- One canonical OPEN Type-105 Task per distinct actual appointment; reuse a proven existing Task.
- Zero cached matches is **not** proof of no existing task. Check persisted Event ID/Task link and relevant history.
- No history → CREATE only after fresh relationships and duplicate checks.
- Same-visit completed history → do not recreate. Older completed history for a distinct new visit → new Task may be authorized.
- Cancelled, future-conflicting, undated or ambiguous task history → REVIEW, no destructive guess.
- For a possible duplicate OPEN pair, retire only proven duplicates after verifying the canonical Task; never cancel completed history.
- All writes are Stage-7-authorized, idempotent and read-back verified. Unknown CREATE outcome requires retrieval/reconciliation before retry.

## 4. Notification template — frozen, no redesign

**Reuse only the existing PreInspection notification implementation and formatting.** Do not introduce a new generic email template, paraphrased body, alternative branded layout, decorative icons or different subject conventions.

### Existing creator missing-details email
- Producer: `tmv3_preInspectionMissingDetailEmailPreview_`; sender: `tmv3_preInspectionSendMissingDetails_`.
- Retain the current HTML structure, brand header `Pre-Inspection Update Required`, explicit `automated email notification` line, appointment summary, red missing-items section, `What it is now` / `What it should be` comparison table, original Calendar event hyperlink, and follow-up numbering.
- Preserve subject convention `MISSING DETAILS: <appointment> | <date/time> | <missing fields>` and existing `FOLLOW-UP #N` prefix, as generated by that function.
- Send **To** the uniquely verified original creator; **CC** `jay@classicfireplace.ca`, `thang@classicfireplace.ca` according to the existing missing-details template.
- Notification decision must be separate from task eligibility. Missing optional SO by itself does not justify an email or block PreInspection creation; automation-owned defects (title, task link, guest, matching or API failure) are repaired/escalated internally, not emailed as missing customer details. Never request the inspector to fill 854/852/853/860 before the visit.
- Scheduler windows: 08:00, 13:00, 17:00 America/Toronto. Same-week appointment becomes eligible at the next applicable slot; otherwise first email after two business days from event creation. Follow-up only on a subsequent business day after the existing interval and while the required human-supplied details remain missing, prior to appointment start.
- Respect the previously approved creator-exclusion list in the notification eligibility gate: Trevor Burke, Stephen Foley, Pramodh Kumar, Jay Scott and Thang Tran. Exclusion suppresses creator email, **not** record discovery or internal review.
- Track event ID, missing-field fingerprint, creator, send time, next eligible time and outcome. Never blindly resend an uncertain email delivery.
- Keep existing internal failure-alert template and its operator recipient; do not repurpose the creator missing-details template for system errors.
- Format changes to the HTML, subject, recipients or branded phrasing require a separate explicit user approval and before/after regression snapshot.

## 5. Read-only acceptance gates

A record may be called `VERIFIED_COMPLETE` only after all mandatory live checks pass:

1. Correct classification, event ID and required Calendar copies.
2. Exact Calendar title, original content retained, correct Location, Toronto times and clickable link block.
3. Internal Customer and Customer-owned Location verified; Task ID matches one canonical Task.
4. Type 105, OPEN where appropriate, exact Task name, dates, pool and employee assignment.
5. Requested By resolves per the active Employee policy; no guessed Contact substitution.
6. No prohibited inspector-owned fields in any outgoing CREATE/PATCH payload and no inspector-entered field changed.
7. Existing missing-details email template unchanged; notification correctly suppressed or eligible, without accidental sending in dry run.
8. Striven and Calendar writes independently read back; every discovered event included in the reconciliation tally.
9. State ledger and audit timestamps reflect real verification rather than `CACHE_ONLY` success.

Read-only fixture rule: a previously successful appointment must yield `MATCH_EXISTING` / `NO_CHANGE`, **zero task CREATE**, **zero email**, and no Calendar mutation when its records already comply.

## 6. Deployment boundary

This document is a standards revision only. It does **not** represent Apps Script code being deployed, tested, or a numbered Apps Script version being created. Before any implementation release: pull exact live Apps Script source; preserve a restore point; diff existing `30_PreInspection.gs` against the historical GitHub V3 files; enforce these rules in the executable source; run syntax, regression and read-only event tests; confirm zero template changes; then follow the mandatory staged production push, read-back, runtime verification, GitHub sync and five-version retention policy. Never treat the previous historical GitHub snapshot as proof of current live source.
