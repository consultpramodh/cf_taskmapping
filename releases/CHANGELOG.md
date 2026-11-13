# Changelog

## Unreleased

- 2026-11-12 Structure: repo synced to live — `70_Contractor_Weekly.gs` added, `90_Regression_Tests.gs` removed from apps-script (kept as a test fixture), `appsscript.json` gains the Gmail scope and web app settings.

- 2026-11-12 PreInspection: missing-details emails on; Stephen-calendar scan with title/description visit type; YES/NO question inside the approved email for bookings missing CF Preinspects (switch off); Steve + Colleen Trevor paired recipients; first email never waits past the appointment; organizer-first Calendar title writes (10_Shared_Core).
- 2026-11-12 Script Properties: caches and status logs moved to CacheService, verified write records no longer kept, one-time Oct 20 order backfill retired; operator cleanup tools reduce 784 properties to 33.
- 2026-11-12 Operator tools (95_PI_Operator_Tests.gs): slot simulation, test emails to Pramodh only, property inventory/cleanup, held-record check, read-only end-to-end check.

- PreInspection r7 clean rebuild of 30_PreInspection.gs from the 2026-11-11 live capture, with branch payload and order-status guards and a read-only dry run. Deployed to live 2026-11-11 15:23 Toronto with the release gate off (self-test 36/36). See docs/preinspection-rebuild-status.md.

Imported 2026-11-10 live source capture as migration baseline. Not tested as a release; no Apps Script change or deployment.
