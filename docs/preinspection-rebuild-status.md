# Pre-Inspection rebuild status

## Staged decision components — not connected

Created pure, fail-closed decision helpers for order classification, duplicate Location selection using an independently verified default, Task lifecycle, and preliminary payload validation. The helpers have no external side effects and are not called by the existing production worker. Separately, the candidate removes the shared-calendar early return in the existing pre-create order guard: a parsed order number now receives the existing live status check, including Approved and In Progress holds. This change is in GitHub only and would change runtime behavior if deployed. It does not prove that all order references are parsed.

Local validation: 28 staged decision assertions, four order guard cases, and syntax verification. Fixtures do not verify Striven response schema, Apps Script behavior, integration, notification output, event correctness, or outgoing network payloads.

## Required next stages

1. Capture fresh authoritative sources and define verified order/Location adapters, including actual default-field evidence. Never infer a default from Location ordering or numeric ID. Reconcile unresolved Requested By policy before any policy change.
2. Wire classification to every authoritative event, with the approved-order policy and optional-order behavior.
3. Connect lifecycle decisions to complete fresh Task history, same-visit evidence and uncertainty handling.
4. Wire payload protection to the actual outgoing create and patch boundaries; inspect actual serialized requests; preserve inspector-owned fields 854, 852, 853 and 860.
5. Preserve and regression-test the approved notification template, owner-specific assignment, copy parity and readback.
6. Test all verticals and then run read-only, isolated and controlled live acceptance using the existing protected release workflow. Validate current live source immediately before any deployment.

The preliminary payload helper must not be treated as a complete payload guard until integrated and tested against real payload schemas, including nested shapes. No deployment, live writes, or end-to-end success is claimed.

## Calendar address and Primary Location — user decision

When the Calendar address exists, use a verified customer-owned address match; if strong duplicates remain, use a verified default among the matching records. When the Calendar address is absent, queue a notice to the sales rep under the approved notification rules before using a verified customer-owned Primary Location as fallback. Uncertain send outcomes do not cause blind retries. The staged connected dry-run is pure and does not send the notice or perform the fallback write.

The existing frozen notification standard names the event creator as recipient. The user now specified the sales rep for missing address; the identity source and cases where creator and sales rep differ must be resolved before live dispatch. Do not assume they are the same or alter the existing approved template.
