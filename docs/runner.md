# Creating a new XF1 workbook

The runner starts with an empty desktop Excel workbook and constructs Controls, Monthly, Quarterly, Annual, Input_sheet and Checks at the requested timeline length. It creates the formatting, calendars, flags, names, validations and views, then calculates and verifies the result. No existing `.xlsx` or workbook map is needed. It does not upgrade existing models, publish releases or add business calculations/checks.

## Requirements and invocation

Use Windows, PowerShell 7 or later and installed desktop Microsoft Excel with `Formula2` support. No Python or Node runtime is needed. Codex runs the command for the analyst, subject to the computer's normal execution permissions. Excel must be installed and usable under that Windows user account.

From the package directory:

```powershell
pwsh -File ./runner/New-XF1Workbook.ps1 `
  -StartDate '2025-01-01' -LastActualDate '2026-08-31' `
  -FinancialYearEndMonth 12 -Currency 'USD' -Months 60 `
  -OutputPath 'C:\Models\My_XF1_Model.xlsx'
```

Dates use `yyyy-MM-dd`. Currency, dates, financial year-end month and output path are required. `Months` defaults to 48. Optional `ActualLabel` and `ForecastLabel` default to A and F. Calendar rules are defined in [Controls](controls.md); the runner enforces them before opening Excel. Supported start years begin at 1901; dates must fit Excel's date range. The maximum horizon is 16,358 months, allowing the timeline and spacer to fit Excel's column limit; this is a structural ceiling, not a performance benchmark.

The output path must be new; `templates/` may hold generated examples. Existing files are never overwritten. Beside the new workbook the runner writes:

- `<name>.map.json`: actual configuration, calendar extents and verification scope. Specification names refer to `docs/` in the recorded release package.
- `<name>.build.json`: package version, manifest hash, generation mode, verification results, workbook size and stage timings.

Existing workbook or sidecar paths are rejected. To regenerate the same filename, remove or move its workbook and both sidecars, or use a new filename. A failed build discards only its own staging files and publishes no completed workbook. The runner owns and closes its Excel instance; it does not attach to the analyst's open workbooks.

## Implementation and performance

The documents define the intended standard. Codex maintains the tested implementation: `runner/Layout.ps1` constructs the layout and styles, and `runner/Workbook.ps1` supplies calendar formulas, verification and output maps. Routine builds execute that code; they do not ask a model to reinterpret the prose or replay historical correction scripts. Change the relevant specification and implementation together when evolving the standard.

Construction, calculation and verification happen in one Excel session. Bulk label arrays and whole-row formula assignments avoid cell-by-cell COM writes. Calendar values are read in arrays and checked against independent date calculations. Formatting is limited to template blocks; no routine cross-engine import, rendering or full dependency rebuild is performed. Visual approval applies to the tested builder version; each routine build does not need another full visual review.

Period-type and reporting-label rows wrap and fit their calculated text, with a minimum height of 18 points. Only those rows are fitted, so long custom or partial-period labels remain legible without changing aligned timeline widths.

Checks cover every generated calendar/flag value, actual/forecast transitions, partial periods, linked headers, number/unit formats, names, validation presence, outlines, panes, visible boundaries, surplus timeline contents and blank Inputs/Checks areas. These are build-time checks; the Checks worksheet remains unconfigured.

The runner supports the six-sheet layout identified by `templateId`. Additional sheets require construction code and tests. The development benchmark adds a synthetic payload sheet solely to exercise larger builds; it is not part of normal output.

Development test instructions are in `tests/README.md` in the project. Large compressed file size alone does not describe calculation complexity; synthetic benchmark results must not be presented as a guarantee for all financial models. Stage timings exclude final Excel shutdown, which can occasionally delay completion.

## Development and release boundary

The manifest binds the instructions and builder code to a package version using SHA-256 hashes. A mismatch stops a build. Generated workbooks and their maps are not release dependencies. Deleting every example workbook does not remove the ability to build a new one. `docs/workbook-map.json`, if present, describes the older v1.1 example only and is never read by the builder.

After intentional local changes, refresh a **development** manifest with `tests/Update-DevelopmentManifest.ps1`, then run the relevant tests. This does not approve or publish anything. Published versions must be immutable; prepare a new version for subsequent changes.

The separate MCP distribution layer is implemented and locally tested, but has not been published or deployed. It obtains the current approved version before each new workbook, reuses a local cache only for that exact verified version, and keeps the version fixed throughout the build. If the approved version cannot be established, it stops rather than silently using a stale cache. Downloads, caching and release management belong to that distribution layer. Publication occurs only when Vavrinec explicitly requests it.
