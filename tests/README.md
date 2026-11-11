# Tests

Apps Script regression functions live in `apps-script/90_Regression_Tests.gs` and run inside Apps Script.

Standalone Node checks (no network, no Google services) run in CI via `.github/workflows/pi-rebuild-decisions.yml`:

```bash
node tests/preinspection/rebuild-equivalence.test.cjs
node tests/preinspection/rebuild-decisions.test.cjs
node tests/preinspection/order-guard.test.cjs
```

`tests/preinspection/harness.cjs` loads the `.gs` files into one VM context with in-memory stand-ins for PropertiesService, LockService, Utilities, MailApp, CalendarApp, ScriptApp and SpreadsheetApp. Mocked results are not live verification.
