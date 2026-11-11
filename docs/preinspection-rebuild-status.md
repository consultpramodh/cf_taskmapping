# Pre-Inspection rebuild status

## r7 clean rebuild (2026-11-11) — not deployed

`apps-script/30_PreInspection.gs` is a clean rewrite of the live module captured on 2026-11-11 (staging r5, `TMV3_PI_CLEANED_RELEASE_APPROVED = false`). Live behaviour is the baseline; preinspection-branch guards are added on top. The release gate stays **false**: no Striven, Calendar, guest or email mutation can run from this file until it is deliberately flipped after verification.

### Structure

One file, thirteen numbered sections that follow one appointment through the pipeline: §1 release gate, locked standard and policy · §2 source/intake · §3 classification · §4 Step 3 anchor · §5 Step 4 identity · §6 Step 5 Task match · §7 Calendar presentation and parity · §8 write guards · §9 notification (frozen renderer) · §10 batch cycle (stages A–G) · §11 daily slots · §12 operator diagnostics and dry run · §13 staged decision helpers. The 16 `tmv3_piOriginal_*` wrapper pairs are replaced by one gate check per entry point (`tmv3_piStagingDisabled_`). The 230-line cycle is split into named stage functions with identical control flow.

### Preserved exactly

- Every function called by 10_Shared_Core, 40_Installation and 90_Regression_Tests, all trigger handlers (08:00/13:00/17:00 slots and the retired continuation/five-minute stubs) and both TMV3 menu items.
- The missing-details renderer `tmv3_preInspectionMissingDetailEmailPreview_` is copied byte-for-byte from live; its fingerprint is still `b77a1a07`, matching `TMV3_HARD_RULES.PreInspection.notificationRendererFingerprint`. Subject, recipients, CC, follow-up numbering, scheduler and dedupe rules are unchanged.
- Protected fields 852/853/854/860, Type 105, Pool 8, blank Description at CREATE, no Sales Order attachment, Requested By = verified owning-Calendar Employee, no unattended guest changes or title rewrites.

### Intentional behaviour changes (stricter, fail closed)

1. **Deep payload scan** (branch `tmv3_piRebuildValidatePayload_`) now runs at the real CREATE and PATCH boundaries. A protected custom field or order attachment at any nesting depth blocks the write (`PI_PAYLOAD_REJECT_*_NO_WRITE`). The canonical create payload from `tmv3_buildReplacementCreatePayload_` passes.
2. **Sales Order status ID must agree with its name** before CREATE (`PI_SO_STATUS_ID_NAME_MISMATCH_NO_CREATE`), table `18 incomplete, 19 quoted, 20 pending approval, 22 approved, 25 in progress`. IDs 19/22/25 are already used by live classification; **18 and 20 come from branch work and must be confirmed against Striven before release**, otherwise Incomplete/Pending Approval orders will be held for review.
3. New read-only `tmv3_preInspectionDryRun(eventId)` evaluates one CF Preinspects event through source, classification, intent, hold, owner, missing-details and notification decision (with subject/recipient preview). It makes no Striven calls and no writes; it does not evaluate the Task plan.

### Removed (no callers, triggers or menu items)

Dated one-off and retired handlers: Oct 21 preview/run/cache, Nov 9 repairs, Terry quiet-intake trio, Nov 11 Terry/Jakob read-only and guarded runner, the Oct 12–19 retrospective audit, retired "Today" run/batch/unlinked/cache stubs, the schema probe, the bulk legacy-title stub, and the unused `tmv3_preInspectionStoreDiscovery_` / `tmv3_preInspectionTodayEligibleQueue_` helpers. The test suite pins this exact list.

### Not carried from the preinspection branch

The branch's 10_Shared_Core (`tmv3_taskDifferences_` pool-token fix, AUD3 runner) and 90_Regression_Tests changes are untouched by this rebuild and still need their own cross-vertical review. The branch's order guard variant (inferred-evidence lookup) was not used: nothing writes `TMV3_PI_INFERRED_*`, and live semantics are authoritative.

## Evidence

`tests/preinspection/rebuild-equivalence.test.cjs` loads the live baseline (`tests/preinspection/baseline/30_PreInspection.live-2026-11-11.gs`) and the rebuild into identical Node VM contexts with in-memory Apps Script services and a fixed clock, and compares outputs, thrown messages and every stubbed side effect (216 checks):

- interface and approved-removal list; every called `tmv3_*` resolves; renderer fingerprint; gate defaults closed;
- ~40 pure functions over generated fixtures (classification, anchors, Task matching, link block, notification decision, email HTML/subject, write guards);
- stubbed Striven read adapters (order guard, Locations, Task pagination, linked-Task hydration) and the PI Calendar Review sheet writer;
- every gated entry point with the gate closed (identical, zero side effects);
- with the gate forced open, 20 cycle scenarios × 8 entry points (create, patch, deferral, budget stop, hold, emails, uncertain send, pending recovery, history), each run twice, comparing results, traces and Script Properties;
- identical outcomes for all 19 V3 regression entry points in 90_Regression_Tests.

Passing against both the repository's shared files and the exact live capture. These are mocks: they do not prove Striven schemas, Apps Script runtime behaviour, trigger ownership or end-to-end correctness. 7 of the 19 legacy regression functions already fail against live (stale expectations); the rebuild reproduces those outcomes rather than hiding them.

## Before any release

1. Fresh live capture; confirm no drift from the baseline file used here.
2. Confirm Striven status IDs 18 and 20; resolve Karen's employee mapping; confirm Step 7 emits the owner-employee requester.
3. Run `tmv3_runPreInspectionCleanedSelfTest` (36 checks) and `tmv3_preInspectionDryRun` on real events in Apps Script.
4. Follow docs/release-process.md: backup, single writer, staged push, read-back. Flip the gate only as a separate, explicit decision.
