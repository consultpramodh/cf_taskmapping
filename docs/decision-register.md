# Decision register

- 2026-11-10: Four permanent vertical branches plus main. Baseline is the November 10 live capture; no release approval implied.
- 2026-11-10: Pre-Inspection standard 2026-11-10-r1 is the documented reference. Reconcile Requested By and all other conflicting rules with authoritative evidence before changes.
- 2026-11-10: Inspector-owned fields 854, 852, 853 and 860 are unmanaged. Preserve the approved notification template.
- 2026-11-23: Each standard slot resolves every event that needs it at that slot; deferring to the next slot is not allowed. The slot chain reruns a workflow (one pass per execution, sweep resumes from its cursor) until its sweep is done, then moves on. Stops only for API budget, no progress or the 20-pass cap, each reported in TM Audit as SLOT_CHAIN.
- 2026-11-23: Present-day events are resolved in the slot even after their appointment time. The PreInspection sweep is done only when every current event (today 00:00 onward) was handled at its current calendar fingerprint, so events added or changed behind the sweep are picked up; today's events go first. A slot-chain pass not finished after 12 minutes is treated as dead and rerun.
