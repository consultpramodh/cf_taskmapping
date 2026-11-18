# Changelog

## Unreleased

Imported 2026-11-10 live source capture as migration baseline. Not tested as a release; no Apps Script change or deployment.

### Delivery automation fixes (staged on the 2026-11-17 live capture, not deployed)

- Lock starvation / dropped slots: Delivery now uses the same follow-on mechanism as Install (`TMV3_VERTICAL_FOLLOW_ON`). A busy lock is retried for up to 2 minutes, and an unfinished sweep continues one minute later (max 12 follow-ons a day). `tmv3_deliverySweepContinue` was a retired stub and now mirrors `tmv3_installSweepContinue`. Service is unchanged.
- Stale intake: a new Delivery sweep refreshes the Striven source reports and rebuilds only the Delivery tab from the Delivery calendar (`tmv3_deliveryRefreshView_`, the live Install rebuild minus the Install-only creator filter). If either step fails, the sweep falls back to the existing cache and tab, with an ATTENTION audit row. Shared change: a Delivery branch in `tmv3_verticalProductionCycle_`; Install, Service and Pre-Inspection paths are unchanged.
- SO parsing: `tmv3_extractLegacyDeliveryOrderNumber_` now:
  - accepts `SO-`, `SO:`, `SO No.` and 4-digit labelled SOs;
  - ignores Customer, Task, Work Order, PO, Invoice and Quote numbers and URL IDs;
  - uses the `#12345` and bare 6-digit fallbacks only when exactly one candidate exists, otherwise sends the event to REVIEW.
- New standalone test, not part of the Apps Script project: `node tests/delivery/delivery-automation.test.cjs` (29 cases, passing locally). Not yet verified against live Apps Script.

