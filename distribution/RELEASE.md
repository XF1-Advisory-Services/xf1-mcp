# Explicit release and deployment procedure

Only Vavrinec's explicit instruction authorizes preparation/update/publication of a release. Release 1.2.0 publication and the separate Vercel deployment were explicitly authorized on 2026-09-23. Future releases still require a new explicit instruction.

## Local preparation

Wait for the template owner to confirm readiness and current manifest hashes. Run the relevant builder checks and distribution tests. Prepare into a new, unused directory:

```powershell
pwsh -NoProfile -File distribution/Prepare-XF1Release.ps1 `
  -Version '1.2.0' -OutputDirectory '.build/release-candidates/1.2.0-review1'
```

This only writes local files and refuses an existing destination. It produces the fixed ZIP, a copy of the bridge, the copied package directory, `release-candidate.json`, and `preparation.json`. Keep test results with the review record. A source hash mismatch aborts preparation. An incomplete preparation directory is not publishable; use a fresh directory after resolving the issue. Do not refresh source hashes just to make a mismatch pass.

For a tested candidate, retain its exact ZIP and bridge rather than recreating them before publication. If code changes, prepare and test a new candidate; never alter an already published version. Service changes and template changes can be reviewed separately, but the selected package hash and bridge compatibility must agree.

## Publication after explicit instruction

1. Review exactly which source files will enter `XF1-Advisory-Services/xf1-mcp`. Exclude `.build/`, templates, workbooks, sidecars, caches, environment files and machine-specific artifacts. Initialize/push the shared repository only as part of that authorized source publication. Never upload the workspace wholesale.
2. Enable [GitHub release immutability](https://docs.github.com/en/code-security/how-tos/secure-your-supply-chain/establish-provenance-and-integrity/prevent-release-changes) for this repository before its first release. Create the approved version's draft release/tag and upload the **tested** `xf1-<version>.zip` and `Invoke-XF1Build.ps1` before publishing: immutable releases cannot have assets added or replaced after publication. Retain candidate metadata and validation evidence for audit. Download the public assets and compare bytes/hashes with the reviewed candidate.
3. Set `service/approved-release.json` to `{ "schemaVersion": 1, "release": <approved descriptor> }`, using the candidate's fields/hashes and changing only descriptor `approval` to `approved`. This is the current-release selection; no automatic latest-release lookup is used.
4. Create/link the dedicated Vercel project and deploy the reviewed service with that selection. Verify the live tool response, hashes, colleague setup and one end-to-end local build before announcing availability. Record the actual production URL. Public users need no authentication.

These are manual release actions, not installed automation. Selecting a different current version or rolling back the service requires another explicit instruction. Rollback changes the selected descriptor/service deployment; it does not rewrite existing release assets or migrate any workbooks. Nulling the selection and deploying it stops new builds through MCP while leaving saved workbooks intact.

## Vercel target and settings

| Setting | Value |
|---|---|
| Team | `xf-1-advisory` (`team_2Nem4Au9ovP0kTW4AY33IKeA`) |
| Project | Separate new project `xf1-mcp` |
| Repository | `XF1-Advisory-Services/xf1-mcp` |
| Project root | `service/` |
| Runtime | Node.js 24 |
| Framework preset | Other (no framework) |
| Install | `npx --yes pnpm@11.19.0 install --frozen-lockfile` |
| Package-manager discovery | Root and service `package.json` pin pnpm 11.19.0; Vercel `ENABLE_EXPERIMENTAL_COREPACK=1` |
| MCP route | `/api/mcp` using the actual production hostname |
| Function duration | 30 seconds |
| Authentication | None; production MCP must be publicly reachable |
| Runtime secrets / database | None required |
| Automatic Git deployments | Disabled in `service/vercel.json` |

Only the new project's settings are involved; do not change the dashboard project's environment, domains, build configuration or deployment protection. Team billing and limits are shared. The dedicated `xf1-mcp` project was created during the authorized publication work.

The service uses Vercel's documented [Web Request function interface](https://vercel.com/docs/functions/functions-api-reference). Automatic Git deployments are disabled using [git.deploymentEnabled](https://vercel.com/docs/project-configuration/git-configuration). Use a manual deployment from the reviewed commit. The repository-level package-manager pin supports [Vercel Corepack discovery](https://vercel.com/docs/package-managers) when the service is in a subdirectory.
