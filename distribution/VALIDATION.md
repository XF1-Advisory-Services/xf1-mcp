# Local distribution verification - 23 September 2026

Result: **21 checks passed**, including an actual desktop Excel build. Node reports 22 passing test entries because the parent test is also counted. No release was published and no service was deployed.

Final run: `.build/distribution/2026-09-23T15-58-53-098Z/results.json`. The same tested ZIP and bridge were copied byte-for-byte to `.build/release-candidates/1.2.0-review1/`; their hashes and the native build's manifest hash were compared successfully. Builder source was `1.2.0-dev.3` after the template owner's documentation refresh (source manifest SHA-256 `55784C45EC9E77E15B08729FC7F95B6BE7DB478BAF012E982C0B8AF7F5EE6F7B`).

## Coverage

- Production handler: MCP initialization, tool discovery, no configured approval, non-cacheable responses, unrelated browser-origin rejection, descriptor validation and versioned GitHub asset restrictions.
- Local bridge: first download, exact-version cache reuse with a fresh MCP check, offline/unavailable approval refusal, altered bridge hash, corrupt ZIP, tampered cache, selected-version pinning while current approval changes, separate caches for a new version, manifest mismatch, exclusive locking and ZIP traversal rejection.
- Test boundary: local HTTP requires explicit loopback mode, which cannot point at a remote MCP endpoint.
- Native handoff: a new 48-month six-sheet workbook for January 2025-December 2028, USD, actuals through August 2026. All existing runner checks passed, with zero formula errors; workbook, map and build report were created. An attempt to reuse the output path was rejected. The `.build.json` records the same manifest hash as the candidate.
- Source boundary: the tests did not change the development manifest or the production service's null approval selection.

Test execution used Node.js 24.19.0, PowerShell 7.6.5 and local Microsoft Excel, orchestrated from Codex. The pinned pnpm 11.19.0 dependency installation passed a frozen-lockfile install and its six-entry supply-chain check. Client bridge and native runner need no Node runtime on colleagues' machines.

The first test pass identified OneDrive placeholder reparse points being mistaken for filesystem links. The bridge now distinguishes real symbolic links/junctions. The first native test could not start Excel from the sandbox (COM logon-session error); the same test passed in the authorized Windows user context. Neither issue remains in the final run.

## Verification limits

The MCP transport and bridge were tested against loopback HTTP servers serving the real prepared package. Public GitHub release redirects, actual Vercel runtime/deployment settings, and registration/calling from a colleague's installed Codex client await the explicitly authorized live publication step. Codex orchestration of the local bridge and Excel was tested; a persistent production MCP connection was not installed. Claude was not tested.

The Excel logic was not rewritten. This distribution work adds to the template owner's [builder evidence](../tests/VALIDATION.md): earlier 17-case/calendar and approximately 49 MB synthetic benchmark results, plus dev.3's bounded colour verification. The large benchmark was not repeated because distribution does not process workbook contents. There was no additional full-sheet visual review of the unchanged builder layout.

## Reproduce

See [local test commands](README.md#local-development-and-tests). Each run creates a fresh evidence directory. Set `XF1_TEST_EXCEL=1` to include the native build; otherwise the test suite exercises the distribution layer only. The final candidate and its saved evidence are local review material and must not be uploaded wholesale as release assets.
