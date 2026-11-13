# Tests

The V3 regression suite (`90_Regression_Tests.gs`) was removed from the live project on 2026-11-12; a copy is kept as a test fixture in `tests/preinspection/baseline/`. Node tests in `tests/preinspection/` run in CI.

Standalone Node checks (no network, no Google services) run in CI via `.github/workflows/pi-rebuild-decisions.yml`:

```bash
node tests/preinspection/rebuild-equivalence.test.cjs
node tests/preinspection/rebuild-decisions.test.cjs
node tests/preinspection/order-guard.test.cjs
```

`tests/preinspection/harness.cjs` loads the `.gs` files into one VM context with in-memory stand-ins for PropertiesService, LockService, Utilities, MailApp, CalendarApp, ScriptApp and SpreadsheetApp. Mocked results are not live verification.
