# XF1 Excel templates

Published release: [1.2.0](https://github.com/XF1-Advisory-Services/xf1-mcp/releases/tag/v1.2.0). Public MCP endpoint: `https://xf1-mcp.vercel.app/api/mcp`. See [colleague setup](distribution/README.md).

One project contains the template specifications, local builder, and MCP distribution implementation. It creates workbooks from scratch; existing Excel files are outputs, not dependencies. Nothing is automatically published when local files change.

| Location | Purpose |
|---|---|
| `docs/` | Shared conventions and sheet specifications; legacy workbook-specific records are reference only |
| `templates/` | Generated example workbooks; they can be deleted and recreated |
| `runner/` | Native Excel creation and verification scripts |
| `tests/` | Development tests and synthetic large-workbook benchmark |
| `release-manifest.json` | Version, required software and hashes of code/instructions; no workbook input |
| `distribution/` | Local PowerShell bridge, release preparation, package contract and colleague setup |
| `service/` | Unauthenticated MCP service for the separate `xf1-mcp` Vercel project |
| `.build/` | Local test results, previews and disabled historical scripts; not release assets |

Read [runner usage](docs/runner.md) to create a new workbook and [shared guidance](docs/shared-guidance.md) for modelling conventions. Development tests and release preparation are described in [tests/README.md](tests/README.md).

The development builder is tracked separately from the approved distribution. `service/approved-release.json` selects release **1.2.0**, with immutable versioned download URLs and SHA-256 hashes. Publication and deployment were explicitly authorized by Vavrinec; future changes still require a separate release approval. Codex is the tested client; Claude compatibility is expected but untested. The service has no authentication, database or remote Excel execution.

See [distribution setup](distribution/README.md), [package contract](distribution/CONTRACT.md), and [release procedure](distribution/RELEASE.md). GitHub destination: `XF1-Advisory-Services/xf1-mcp`; Vercel team: `xf-1-advisory`, with a separate project named `xf1-mcp` and root directory `service`.
