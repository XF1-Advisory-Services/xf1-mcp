# XF1 local execution and MCP setup

Status: implemented locally; no published package or deployed XF1 endpoint yet. The examples below become usable after Vavrinec explicitly approves publication and deployment. The Vercel project will be separate from the test dashboard.

## What colleagues need once

Use a Windows laptop with PowerShell 7 or later, usable desktop Microsoft Excel with `Formula2`, and Codex with permission to run local PowerShell. No Python, Node.js, Git, GitHub account or Vercel account is needed on the colleague's laptop for workbook creation. Normal organization execution policies still apply; this setup does not change them or bypass script restrictions. Install missing prerequisites through the organization's approved software process.

After the actual production URL has been confirmed, connect Codex:

```powershell
codex mcp add xf1 --url 'https://ACTUAL-XF1-HOST/api/mcp'
codex mcp list
```

Replace `ACTUAL-XF1-HOST`; it is deliberately not a claimed live hostname. There are no login/token parameters. This command follows [official Codex MCP setup guidance](https://developers.openai.com/learn/docs-mcp); it was also checked against the installed `codex mcp add --help`. In an app-only setup, add the same Streamable HTTP URL through the client's MCP configuration. Client-specific UI setup has not yet been tested against the future live endpoint.

Ask Codex to call `get_current_release`, download its `release.bridge.url`, and verify the file's SHA-256 against `release.bridge.sha256` **before running it**. Save it in a versioned local folder, for example `%LOCALAPPDATA%\XF1\bridge\1.0.0\Invoke-XF1Build.ps1`. This initial script is the execution bridge: adding a remote MCP URL alone does not grant the server access to local Excel. Keep its path available to the local agent. If a later approved release requires a different bridge version or hash, download and verify that bridge before another build; do not overwrite a bridge being used by an active build.

## Each new workbook

Ask Codex to create a new XF1 workbook and specify the start date, last actual date, financial year-end month, currency and output path. The horizon defaults to 48 months. The bridge performs a fresh MCP approval lookup and then runs the tested builder locally:

```powershell
pwsh -NoProfile -File 'C:\path\to\Invoke-XF1Build.ps1' `
  -McpUrl 'https://ACTUAL-XF1-HOST/api/mcp' `
  -StartDate '2025-01-01' -LastActualDate '2026-08-31' `
  -FinancialYearEndMonth 12 -Currency 'USD' -Months 48 `
  -OutputPath 'C:\Models\My-XF1-Model.xlsx'
```

Dates must follow the builder's [calendar rules](../docs/runner.md). A new output filename is required. Only the existing runner constructs, calculates, verifies and saves Excel. The bridge returns JSON containing the pinned version, manifest hash, cache location, download status and the runner's build report. The workbook and its `.map.json`/`.build.json` remain on the laptop. The server receives no workbook configuration, path or workbook contents.

Packages are cached at `%LOCALAPPDATA%\XF1\releases\xf1-base-six-sheet\<version>\package`. Every invocation obtains current approval and verifies the cached manifest and all listed files. It downloads only a missing approved version. `-ResolveOnly` performs the same selection and integrity checks without opening Excel, useful during first setup. Once a build has selected its version, it finishes with that version even if another is subsequently approved.

An unavailable MCP endpoint, missing approval, changed local bridge, mismatched package, modified cache, or concurrent use of the same package stops the command. It never falls back to a stale release. Wait for an active build before retrying a lock failure. For corrupted cache files, investigate or move aside that version's directory when no build is running, then retry to download the same approved version. No automatic cache deletion is performed.

Suggested persistent project instruction for colleagues:

> For every NEW XF1 workbook, call the XF1 MCP get_current_release tool, verify the approved local bridge, and invoke it with the actual production MCP URL and the user's configuration. Stop if current approval cannot be established. Reuse only the exact verified package version; keep it pinned throughout the build. Execute the packaged desktop Excel runner. Do not implement Excel construction yourself, upload workbooks, or migrate existing models.

Codex is the tested execution environment. Claude compatibility is expected but not separately tested; the same package and bridge apply to all clients capable of local PowerShell execution.

## Local development and tests

Service development needs Node.js 24 and pnpm 11.19.0. From `service/`, run `pnpm install --frozen-lockfile`, then `pnpm start`. It binds only to `127.0.0.1:3000`. With no approved release configured, the tool deliberately returns an error. The service uses [Vercel's MCP adapter](https://github.com/vercel-labs/mcp-handler) and its stateless Streamable HTTP compatibility mode; no Redis or application framework is needed.

From the project root:

```powershell
node --test service/test/distribution.test.js
# Include a real local Excel build and output-collision check:
$env:XF1_TEST_EXCEL = '1'
node --test service/test/distribution.test.js
Remove-Item Env:XF1_TEST_EXCEL
```

Tests start ephemeral loopback MCP and asset servers, prepare isolated fixtures, and write evidence under `.build/distribution/`. They do not alter production approval or publish anything. `-AllowLoopbackForTest` exists only for these fixtures: it requires a loopback MCP endpoint and permits local HTTP asset URLs. It is never part of colleague setup. Desktop Excel tests must run in a Windows user session able to launch Excel; an isolated sandbox may require normal host execution approval.
