# Changelog

## Unreleased

Imported 2026-11-10 live source capture as migration baseline. Not tested as a release; no Apps Script change or deployment.

## Unreleased — ordered slot run (2026-11-20, not deployed)

- Problem: the twelve per-workflow 08:00/13:00/17:00 triggers fired in random order and overlapped (2026-11-20 08:00: Install 07:24, Service 08:02, PreInspection 08:04, Delivery 08:39). PreInspection ran while Service held the script lock: Calendar review was skipped (`PI_REVIEW_LOCK_BUSY`) and lock waits used up most of its 3-minute batch window, so only 3 events were checked and 4 were left for the next slot.
- Change: one trigger per slot (`tmv3_standardSlot0800/1300/1700`) runs PreInspection, then Install, Delivery and Service, each in its own execution, each starting one minute after the previous one finishes. A one-off safety trigger moves the run on if a step is killed by the time limit; an error in one workflow does not stop the others.
- `tmv3_installStandardDailyTriggers` removes the old per-workflow slot triggers and installs the three slot triggers (business triggers 12 → 3). The per-workflow handlers stay callable by hand.
- New read-only `tmv3_slotChainStatus()`. Operator trigger checks updated.
- Base: 2026-11-18 live capture. Requires live-source parity check, backup, and running `tmv3_installStandardDailyTriggers` once after deployment.
