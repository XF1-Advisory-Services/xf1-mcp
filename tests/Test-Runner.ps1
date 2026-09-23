#requires -Version 7.0
[CmdletBinding()]
param([string]$OutputDirectory=(Join-Path (Split-Path $PSScriptRoot -Parent) ('.build/tests/'+(Get-Date -Format 'yyyyMMdd-HHmmss'))))
$ErrorActionPreference='Stop'
$projectRoot=Split-Path $PSScriptRoot -Parent
$null=[IO.Directory]::CreateDirectory($OutputDirectory)
$root=Join-Path $OutputDirectory 'package-without-workbooks'
$null=[IO.Directory]::CreateDirectory($root)
$manifest=Get-Content -Raw -LiteralPath (Join-Path $projectRoot 'release-manifest.json') | ConvertFrom-Json -AsHashtable
foreach ($asset in $manifest.assets) {
    if ([IO.Path]::GetExtension($asset.path) -eq '.xlsx' -or $asset.path -like '*.map.json' -or $asset.path -like '*workbook-map.json') {throw 'A workbook or output map remains a build dependency.'}
    $destination=Join-Path $root $asset.path
    $null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
    [IO.File]::Copy((Join-Path $projectRoot $asset.path),$destination,$false)
}
[IO.File]::Copy((Join-Path $projectRoot 'release-manifest.json'),(Join-Path $root 'release-manifest.json'),$false)
$runner=Join-Path $root 'runner/New-XF1Workbook.ps1'
$sourceFiles=@($manifest.assets.path)
$before=@{};foreach ($p in $sourceFiles) {$before[$p]=(Get-FileHash -LiteralPath (Join-Path $root $p)).Hash}
$cases=@(
    @{name='baseline-60';start='2025-01-01';cutoff='2026-08-31';fy=12;months=60;currency='USD'},
    @{name='requested-48';start='2025-01-01';cutoff='2026-08-31';fy=12;months=48;currency='USD'},
    @{name='default-48';start='2025-01-01';cutoff='2024-12-31';fy=12;currency='EUR'},
    @{name='minimum-all-actual';start='2024-01-01';cutoff='2024-12-31';fy=12;months=12;currency='GBP'},
    @{name='fiscal-partial-14';start='2024-04-01';cutoff='2025-04-30';fy=3;months=14;currency='INR'},
    @{name='expanded-73';start='2023-07-01';cutoff='2024-06-30';fy=6;months=73;currency='USD'},
    @{name='quarter-complete-year-partial';start='2025-10-01';cutoff='2026-03-31';fy=9;months=15;currency='AED'},
    @{name='custom-labels';start='2025-01-01';cutoff='2025-06-30';fy=12;months=13;currency='USD';actualLabel='Actual';forecastLabel='Budget'}
)
$results=@(@{case='package-without-workbook-or-map';passed=$true})
foreach ($case in $cases) {
    $args=@{StartDate=$case.start;LastActualDate=$case.cutoff;FinancialYearEndMonth=$case.fy;Currency=$case.currency;OutputPath=(Join-Path $OutputDirectory ($case.name+'.xlsx'))}
    if ($case.ContainsKey('months')) {$args.Months=$case.months}
    if ($case.ContainsKey('actualLabel')) {$args.ActualLabel=$case.actualLabel;$args.ForecastLabel=$case.forecastLabel}
    $watch=[diagnostics.stopwatch]::StartNew()
    $raw=& $runner @args
    $report=($raw -join [environment]::NewLine) | ConvertFrom-Json -AsHashtable
    if ($report.generation -ne 'from-scratch') {throw 'Unexpected generation mode.'}
    $results+=@{case=$case.name;passed=$true;elapsedSeconds=[math]::Round($watch.Elapsed.TotalSeconds,2);timings=$report.timings;bytes=$report.workbookBytes}
    Write-Output "Passed $($case.name): $([math]::Round($watch.Elapsed.TotalSeconds,2))s"
}
$valid=@{StartDate='2025-01-01';LastActualDate='2025-06-30';FinancialYearEndMonth=12;Currency='USD';OutputPath=(Join-Path $OutputDirectory 'rejected.xlsx')}
$invalid=@(
    @{name='fiscal-mismatch';change=@{FinancialYearEndMonth=3}},
    @{name='not-month-end';change=@{LastActualDate='2025-06-29'}},
    @{name='before-start';change=@{LastActualDate='2024-11-30'}},
    @{name='after-end';change=@{LastActualDate='2035-01-31'}},
    @{name='too-short';change=@{Months=11}},
    @{name='labels-identical';change=@{ActualLabel='A';ForecastLabel='A'}},
    @{name='existing-output';change=@{OutputPath=(Join-Path $OutputDirectory 'baseline-60.xlsx')}}
)
foreach ($case in $invalid) {
    $args=$valid.Clone();foreach ($key in $case.change.Keys) {$args[$key]=$case.change[$key]}
    $rejected=$false;try {$null=& $runner @args} catch {$rejected=$true}
    if (-not $rejected) {throw "Invalid request was accepted: $($case.name)"}
    $results+=@{case=$case.name;passed=$true;rejected=$true}
}
$badManifest=Join-Path $root ('.manifest-test-'+[guid]::NewGuid().ToString('N')+'.json')
try {
    $manifest=Get-Content -Raw -LiteralPath (Join-Path $root 'release-manifest.json') | ConvertFrom-Json -AsHashtable
    $manifest.assets[0].sha256='0'*64
    $manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $badManifest -Encoding utf8
    $rejected=$false
    try {$null=& $runner @valid -ManifestPath $badManifest} catch {
        if ($_.Exception.Message -notlike '*Asset hash mismatch*') {throw}
        $rejected=$true
    }
    if (-not $rejected) {throw 'Altered release asset was accepted.'}
    $results+=@{case='asset-hash-mismatch';passed=$true;rejected=$true}
} finally {if (Test-Path -LiteralPath $badManifest) {Remove-Item -LiteralPath $badManifest}}
foreach ($p in $sourceFiles) {if ((Get-FileHash -LiteralPath (Join-Path $root $p)).Hash -ne $before[$p]) {throw "Source changed: $p"}}
if (Test-Path -LiteralPath $valid.OutputPath) {throw 'Invalid request created a workbook.'}
$results | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'results.json') -Encoding utf8
Write-Output "All $($results.Count) cases passed without an input workbook or map; package assets unchanged. Results: $OutputDirectory"
