# Decision register

- 2026-11-10: Four permanent vertical branches plus main. Baseline is the November 10 live capture; no release approval implied.
- 2026-11-10: Pre-Inspection standard 2026-11-10-r1 is the documented reference. Reconcile Requested By and all other conflicting rules with authoritative evidence before changes.
- 2026-11-10: Inspector-owned fields 854, 852, 853 and 860 are unmanaged. Preserve the approved notification template.
- 2026-11-17: Install accepts DELIVERY_APPROVED Sales Orders as a Step 3 anchor (same as Delivery and PreInspection). Trigger: SO585729 was blocked because it was cached only as DELIVERY_APPROVED.
- 2026-11-17: Customer-number labels with common typos ("cusr#", "cus#", "cst#") parse as a Customer #. They count only at a word start and only when followed by "#".
- 2026-11-17: Install notice (TMV3_INSTALL_NOTICE_ENABLED) stays OFF until the operator approves piTest_SendInstallNoticeTest() output.
