# Publication record — 2026-09-23

Vavrinec explicitly authorized publishing the release to GitHub and deploying the separate Vercel service. Source and assets were uploaded through the signed-in browser; no additional Git Credential Manager permissions were granted.

## Published identities

- Release: [v1.2.0](https://github.com/XF1-Advisory-Services/xf1-mcp/releases/tag/v1.2.0), immutable tag and assets.
- Release source commit: `1f150bbea68d5b04d8511f65e654be8019d91d34`.
- Production endpoint: `https://xf1-mcp.vercel.app/api/mcp`.
- Vercel project: `xf-1-advisory/xf1-mcp`, root `service`, framework Other, Node.js 24.
- Successful deployment: [9Yad6sJAPEFc8LLjnaDLxchDA1nn](https://vercel.com/xf-1-advisory/xf1-mcp/9Yad6sJAPEFc8LLjnaDLxchDA1nn), Production, Ready, 10-second build.
- Deployed service commit: `e6c1eec5d4816a7863a57180a4817aa3416dc261`. Changes after the release tag only pin the Vercel installer and add repository-level package-manager metadata; the approved workbook package and bridge are unchanged.

| Asset | SHA-256 |
|---|---|
| `xf1-1.2.0.zip` (32,202 bytes) | `7BE70380B5B1BF952ABE50D909B89203396C4256512BC46D7EFB06CB5C48693F` |
| `Invoke-XF1Build.ps1` (bridge 1.0.0) | `9CAD1548E5C278E4A89E598B2A8793D2EC6CB94715AF9C061632EE9D30BE7614` |
| Packaged release manifest | `6F41B7EC7351228004BDF808DA1F56FDA708BEE1CCF9BC89C2EFA7CF783CC648` |

## Verification

Both assets were downloaded from public GitHub without credentials and matched the approved SHA-256 values. The downloaded bridge contacted the public production endpoint without authentication, resolved approved release 1.2.0, downloaded a fresh package and verified all files. A second fresh approval lookup reused that exact verified cache and completed a real desktop Excel build.

The workbook contains six sheets and covers January 2025–December 2028, USD, with actuals through August 2026. Excel calculated and verified the output: 48 monthly periods, 16 quarters, 4 years and zero formula errors. This is a blank modelling template; business checks are not yet configured.

Local evidence is retained under `.build/live-publication/`, including `live-build-result.json`, the workbook and sidecars. Generated workbooks and local evidence are excluded from GitHub. Prior local distribution tests and limitations are in [VALIDATION.md](VALIDATION.md). Client-specific Codex UI setup and Claude have not been separately tested.

The initial Vercel builds failed while selecting an older pnpm installer. The successful build includes pnpm 11.19.0 discovery at repository root and an explicit pinned install command. Automatic Git deployments remain disabled in `service/vercel.json`; future release selection and deployment require explicit approval. No existing dashboard project settings were changed.
