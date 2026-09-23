# XF1 1.2.0 - first release review

This records the local review completed before publication. Vavrinec subsequently approved publishing release 1.2.0 and deploying the separate MCP service. Production selection now identifies that exact approved package. See the GitHub release and distribution setup for publication details.

## What is ready

The MCP service returns the explicitly approved package identity and local execution instructions. The PowerShell bridge checks current approval before each workbook, downloads only a missing exact version, verifies hashes and calls the existing desktop Excel builder. The template owner corrected input/list fill to `#FFFF66`. All 21 local distribution checks passed, including a real Excel build; [validation details and limitations](VALIDATION.md) are recorded.

Candidate location (local only): `.build/release-candidates/1.2.0-review1/`.

| Item | Reviewed value |
|---|---|
| Proposed release | `1.2.0`, tag `v1.2.0` |
| Builder source | `1.2.0-dev.3` |
| Bridge | `1.0.0` |
| Package | `xf1-1.2.0.zip`, 32,202 bytes |
| ZIP contents | Copied manifest plus 12 documented instruction/builder assets |
| Publishable assets | ZIP and `Invoke-XF1Build.ps1` only |
| GitHub | `XF1-Advisory-Services/xf1-mcp` (created, no source pushed by this work) |
| Vercel | Team `xf-1-advisory`; separate new `xf1-mcp` project, root `service/` |
| Authentication | None |

SHA-256 identities:

```text
ZIP       7BE70380B5B1BF952ABE50D909B89203396C4256512BC46D7EFB06CB5C48693F
Manifest  6F41B7EC7351228004BDF808DA1F56FDA708BEE1CCF9BC89C2EFA7CF783CC648
Bridge    9CAD1548E5C278E4A89E598B2A8793D2EC6CB94715AF9C061632EE9D30BE7614
```

`release-candidate.json` supplies proposed public URLs and these identities. `preparation.json` records the source manifest. `implementation-sha256.json` records the reviewed service/bridge/test source files. `evidence/` contains the passing results and native build's sidecars. Those review records are not inputs to ordinary workbook creation.

## What publication would do

After Vavrinec's explicit instruction, publish the reviewed source to the public company repository, enable immutable releases and publish these exact two assets, then create/deploy the separate Vercel service with this candidate marked approved in its selection file. Validate public downloads and a live Codex/local-Excel build, record the real endpoint and make the [colleague setup instructions](README.md) concrete. The existing dashboard project needs no changes. See the [release procedure](RELEASE.md) for exact settings and sequence.

Any implementation or package change before publication requires a new review record and relevant checks. The reviewed package must not be silently rebuilt or replaced.
