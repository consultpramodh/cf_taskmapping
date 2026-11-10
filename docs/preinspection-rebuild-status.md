# Pre-Inspection rebuild status

## Staged decision components — not connected

Created pure, fail-closed decision helpers for order classification, duplicate Location selection using an independently verified default, Task lifecycle, and preliminary payload validation. The helpers have no external side effects and are not called by the existing production worker. No runtime behavior changes by their presence alone.

Local validation: 26 fixture assertions and syntax verification. Fixtures do not verify Striven response schema, Apps Script behavior, integration, notification output, event correctness, or outgoing network payloads.

## Required next stages

1. Capture fresh authoritative sources and define verified order/Location adapters, including actual default-field evidence. Never infer a default from Location ordering or numeric ID. Reconcile unresolved Requested By policy before any policy change.
2. Wire classification to every authoritative event, with the approved-order policy and optional-order behavior.
3. Connect lifecycle decisions to complete fresh Task history, same-visit evidence and uncertainty handling.
4. Wire payload protection to the actual outgoing create and patch boundaries; inspect actual serialized requests; preserve inspector-owned fields 854, 852, 853 and 860.
5. Preserve and regression-test the approved notification template, owner-specific assignment, copy parity and readback.
6. Test all verticals and then run read-only, isolated and controlled live acceptance using the existing protected release workflow. Validate current live source immediately before any deployment.

The preliminary payload helper must not be treated as a complete payload guard until integrated and tested against real payload schemas, including nested shapes. No deployment, live writes, or end-to-end success is claimed.
