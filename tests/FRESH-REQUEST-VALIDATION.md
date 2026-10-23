# Fresh request verification - 23 October 2026

The isolated source checkout was based on company GitHub main `f115c4fced1cd005ff447a6132359796a3e4b980`. The original local project and its dirty changes were not modified.

Final command: `XF1_TEST_EXCEL=1 node --test service/test/request.test.js service/test/distribution.test.js` (environment variable set using PowerShell).

All 34 Node test entries passed: 31 leaf checks and three parent entries. This includes partial and complete requests, same-thread second-request isolation, numbered clarification mapping, explicit reuse, existing-workbook edits, MCP schema enforcement, bridge completeness/provenance enforcement, release/cache integrity and a native Excel build with workbook/sidecar verification and collision rejection. The bridge's validator ran against a nonexistent endpoint, establishing that validation-only requests require neither network nor Excel. `git diff --check` passed.

Distribution evidence: `.build/distribution/2026-10-23T11-02-54-074Z/results.json` and its native workbook/sidecars. Earlier runs recorded an Excel COM logon-session error under the sandbox, then a manifest-preservation assertion caused by agent normalization of the manifest during a running test. The final run used the Windows user session and stable sources; both checks passed.

No workbook construction or calculation logic changed, so the full calendar matrix, large benchmark and visual layout review were not repeated. This verifies deterministic contracts and client-declared provenance, not end-to-end interpretation of arbitrary LLM conversations. A client can falsely label or replay old values; the service is stateless and does not receive values, paths or conversation text. Direct development runner use retains its legacy API.

Release 1.2.3 is prepared locally with bridge 1.1.0 at `.build/release-candidates/1.2.3-review1/`. Its exact packaged runner separately built a six-sheet 48-month workbook with zero formula errors (`verified.build.json`). Publication/deployment and live verification are separate steps; this record does not claim they have completed.
