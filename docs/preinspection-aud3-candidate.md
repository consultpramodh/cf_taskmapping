# Operator reconciliation candidate

Based on live-source inspection run 42865727615 (2026-11-10). Changes task difference comparison and one-click audit runner and adds regression fixtures. AUD3 original incorrectly accepted Pool 18 as Pool 8; candidate tests exact numeric token matching.

Local verification: JavaScript syntax for the two full files; 11 fixture cases with mocked Utilities. Not validated in Apps Script. Shared task-difference behavior affects other verticals and must pass their regressions; date parsing of Sheets Date objects and time zone handling need real runtime tests. No production deployment is authorized by this document. Existing notification renderer and inspector custom fields are not changed by this patch.
