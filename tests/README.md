# Runner verification

Latest completed run: [development candidate verification](VALIDATION.md).

Run from the project root using PowerShell 7 with desktop Excel available:

```powershell
./tests/Update-DevelopmentManifest.ps1
./tests/Test-Runner.ps1
./tests/Test-LargeWorkbook.ps1
```

Refreshing hashes is a development action, not a release. The refresh command refuses a manifest whose status is not `development`.

`Test-Runner.ps1` first creates a package containing only manifest assets, with no workbook or input map. It builds eight configurations, including the user's four-year example, the 48-month default with no actuals, a 12-month leap year with all actuals, fiscal year changes, different horizons, partial quarters/years and custom labels. Invalid dates, horizons, fiscal alignment, duplicate labels, output collisions and asset mismatches must be rejected. Package hashes must remain unchanged. Results and generated workbooks go into a fresh timestamped `.build/tests/` directory.

`Test-LargeWorkbook.ps1` starts from an empty workbook, calls the production layout/calendar functions for a 73-month fiscal-year configuration, and adds 800,000 random text cells and 400,000 simple formulas on a test-only sheet. It verifies calculated results and requires a compressed output of at least 30 MiB. Timings separate base construction, synthetic payload construction, calculation/verification and saving. This measures large native builds, not every financial model's calculation complexity; it is not directly comparable to the previous copy-based benchmark. Files go into `.build/benchmark-from-scratch/` and must not be published.

After a new template design, visually review each sheet; after runner changes, review affected calendar boundaries and formats. Production builds reuse that approval and do not routinely render every sheet.

Before an explicitly requested release, assign a new immutable version, include only the manifest assets, run the relevant tests and retain their results. Exclude `.build/`, benchmark data and historical scripts. The separate MCP distribution layer, including release preparation, current-release selection and caching, is implemented and locally tested; nothing has been published or deployed.
