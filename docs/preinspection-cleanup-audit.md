# Pre-Inspection cleanup audit

Baseline: 2026-11-10 Apps Script live-source capture; branch preinspection. Static inspection only. The 30_PreInspection file contains 133 top-level function declarations; no duplicate top-level function declarations were found across the captured Apps Script files. Static counts do not reveal external calls, dynamically constructed names, Apps Script triggers, menu actions or GitHub Actions dependencies.

## Preserve until verified

- 08:00, 13:00 and 17:00 entry points, daily-slot wrapper, shared automation cycle and continuation handler.
- Notification renderer and sending boundary; preserve the approved template and protected inspector-owned fields.
- Identity, duplicate, write-authorization, readback and operator queue helpers.
- Shared APIs used by other verticals, including tmv3_taskDifferences_. The reconciliation draft changes shared behavior and requires all-vertical tests.

## Cleanup candidates, not yet approved for removal

- Dated one-off handlers tmv3_preInspectionOct21Preview, tmv3_preInspectionOct21RunNext, tmv3_preInspectionOct21RefreshCache, tmv3_preInspectionRepairNov9Stephen and tmv3_preInspectionRepairThreeAddedByOthersNov9. Check installed triggers, menus, runbooks, old and new GitHub workflow callers and external callers before retirement.
- tmv3_preInspectionScheduledCycle is a retired five-minute handler that returns without writes. Check whether a trigger or external caller still invokes it; retain a compatible stub if required.
- Terry intake and other targeted repair helpers: retain pending owner and workflow review, as single static mentions do not prove disuse.

## Operational investigation

The daily-slot wrapper may call the automation cycle repeatedly; each cycle invokes tmv3_refreshOperationalSourcesIfNeeded_(120,false). The helper returns CACHE_FRESH without refreshing when both master and operational caches are within their age thresholds. A stale or absent timestamp can trigger a master and/or transaction refresh. Repeated helper invocation is not proof of repeated report refresh; inspect live timestamps and execution logs before changing cadence. Check current authorized trigger ownership and exact schedules; current-user trigger inventory is incomplete across accounts/projects.

## Exit criteria

Capture current source, trigger inventory, call sites and workflows; establish local and Apps Script tests for matching, duplicate prevention, owner assignment, protected fields, notification template, date/time and all four verticals; remove or retire only proven obsolete paths with compatibility protection; exercise controlled read-only and end-to-end verification with correct evidence labels. Never infer live readiness from a mock or cache-only test. No Apps Script production changes are made by this audit.
