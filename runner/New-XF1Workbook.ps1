#requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$StartDate,
    [Parameter(Mandatory)][string]$LastActualDate,
    [Parameter(Mandatory)][ValidateRange(1,12)][int]$FinancialYearEndMonth,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Currency,
    [Parameter(Mandatory)][string]$OutputPath,
    [ValidateRange(12,16358)][int]$Months = 48,
    [string]$ActualLabel = 'A',
    [string]$ForecastLabel = 'F',
    [string]$ManifestPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'release-manifest.json')
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Validate before starting Excel or creating any output.
$start = [datetime]::ParseExact($StartDate, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture)
$cutoff = [datetime]::ParseExact($LastActualDate, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture)
if ($start.Year -lt 1901 -or $start.Day -ne 1 -or $start.Month -ne ($FinancialYearEndMonth % 12 + 1)) {
    throw 'StartDate must be the first day of the financial year, in 1901 or later.'
}
$end = $start.AddMonths($Months).AddDays(-1)
if ($cutoff -lt $start.AddDays(-1) -or $cutoff -gt $end -or $cutoff.Day -ne [datetime]::DaysInMonth($cutoff.Year,$cutoff.Month)) {
    throw 'LastActualDate must be month-end, from the day before StartDate through the final model date.'
}
foreach ($label in @($Currency,$ActualLabel,$ForecastLabel)) {
    if ([string]::IsNullOrWhiteSpace($label) -or $label.Length -gt 64) { throw 'Currency and period labels must contain 1–64 characters.' }
}
if ($ActualLabel -eq $ForecastLabel) { throw 'Actual and forecast labels must be distinct.' }
$manifestFile = Get-Item -LiteralPath $ManifestPath
$root = $manifestFile.DirectoryName
$manifest = Get-Content -Raw -LiteralPath $manifestFile.FullName | ConvertFrom-Json -AsHashtable
if ($manifest.schemaVersion -ne 2 -or $manifest.templateId -ne 'xf1-base-six-sheet') { throw 'Unsupported release manifest.' }
if ([IO.Path]::GetFullPath($PSCommandPath) -ine [IO.Path]::GetFullPath((Join-Path $root 'runner/New-XF1Workbook.ps1'))) {
    throw 'Run the runner belonging to this manifest package.'
}
foreach ($asset in $manifest.assets) {
    $assetPath = [IO.Path]::GetFullPath((Join-Path $root $asset.path))
    if (-not $assetPath.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Asset path leaves the package.' }
    if ((Get-FileHash -LiteralPath $assetPath -Algorithm SHA256).Hash -ne $asset.sha256) { throw "Asset hash mismatch: $($asset.path)" }
}
foreach ($required in @('runner/New-XF1Workbook.ps1','runner/Workbook.ps1','runner/Layout.ps1')) {
    if ($required -notin $manifest.assets.path) { throw "Unverified required asset: $required" }
}
. (Join-Path $PSScriptRoot 'Workbook.ps1')
. (Join-Path $PSScriptRoot 'Layout.ps1')
$output = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ([IO.Path]::GetExtension($output) -ine '.xlsx') { throw 'OutputPath must end in .xlsx.' }
$outputMap = [IO.Path]::ChangeExtension($output, '.map.json')
$outputReport = [IO.Path]::ChangeExtension($output, '.build.json')
foreach ($path in @($output,$outputMap,$outputReport)) {
    if (Test-Path -LiteralPath $path) { throw "Output already exists; nothing will be overwritten: $path" }
}
$config = @{start=$start; cutoff=$cutoff; months=$Months; fyEnd=$FinancialYearEndMonth; currency=$Currency; actualLabel=$ActualLabel; forecastLabel=$ForecastLabel}
$map = New-XF1LayoutMap $config
$timer = [diagnostics.stopwatch]::StartNew()
$timings = [ordered]@{}
$excel = $book = $null
$savedSettings = $null
$published = [collections.generic.list[string]]::new()
$stage = Join-Path ([IO.Path]::GetDirectoryName($output)) ('.xf1-' + [guid]::NewGuid().ToString('N') + '.xlsx')
try {
    $null = [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($output))

    $excel = New-Object -ComObject Excel.Application
    $savedSettings = @{Calculation=$excel.Calculation; EnableEvents=$excel.EnableEvents; ScreenUpdating=$excel.ScreenUpdating; DisplayAlerts=$excel.DisplayAlerts; AutomationSecurity=$excel.AutomationSecurity}
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Add(-4167)
    if ($book.ReadOnly) { throw 'The new workbook opened read-only.' }
    $excel.Calculation = -4135
    $timings.openSeconds = [math]::Round($timer.Elapsed.TotalSeconds,3)
    $null = New-XF1Sheets $book $excel $config $map
    $null = Set-XF1Configuration $book $config $map
    $null = Assert-XF1Base $book $map $manifest
    $timings.configureSeconds = [math]::Round($timer.Elapsed.TotalSeconds - $timings.openSeconds,3)
    [void]$excel.Calculate()
    $null = Set-XF1LabelLayout $book $config $map
    $timings.calculateSeconds = [math]::Round($timer.Elapsed.TotalSeconds - $timings.openSeconds - $timings.configureSeconds,3)
    $verification = Test-XF1Workbook $book $excel $config $map
    $timings.verifySeconds = [math]::Round($timer.Elapsed.TotalSeconds - $timings.openSeconds - $timings.configureSeconds - $timings.calculateSeconds,3)
    [void]$book.Worksheets.Item('Controls').Activate()
    [void]$book.Worksheets.Item('Controls').Range('D9').Select()
    # Save in the caller's original calculation mode; there are no other books in this owned session.
    $excel.Calculation = $savedSettings.Calculation
    [void]$book.SaveAs($stage,51)
    [void]$book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    $timings.saveSeconds = [math]::Round($timer.Elapsed.TotalSeconds - $timings.openSeconds - $timings.configureSeconds - $timings.calculateSeconds - $timings.verifySeconds,3)
    $resultMap = New-XF1Map $map $config $output $manifest.version $verification
    $resultMap | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath ($stage + '.map.json') -Encoding utf8
    $report = [ordered]@{schemaVersion=1; releaseVersion=$manifest.version; releaseStatus=$manifest.status; generation='from-scratch'; manifestSHA256=(Get-FileHash -LiteralPath $manifestFile.FullName).Hash; workbook=$output; configuration=$resultMap.configuration; verification=$verification; timings=$timings; workbookBytes=(Get-Item -LiteralPath $stage).Length; completedAt=(Get-Date).ToUniversalTime().ToString('o')}
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath ($stage + '.build.json') -Encoding utf8
    # File.Move fails on collisions: no force, including if another process created a target meanwhile.
    [IO.File]::Move(($stage + '.map.json'),$outputMap); $published.Add($outputMap)
    [IO.File]::Move(($stage + '.build.json'),$outputReport); $published.Add($outputReport)
    [IO.File]::Move($stage,$output); $published.Add($output)
    $report | ConvertTo-Json -Depth 12
} catch {
    foreach ($path in $published) { Remove-Item -LiteralPath $path -Force }
    throw
} finally {
    if ($null -ne $book) { [void]$book.Close($false) }
    if ($null -ne $excel) {
        try { if ($null -ne $savedSettings) { foreach ($key in $savedSettings.Keys) {
            if ($key -eq 'Calculation' -and $excel.Workbooks.Count -eq 0) { continue }
            $excel.$key = $savedSettings[$key]
        } } }
        finally { [void]$excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    }
    foreach ($path in @($stage,($stage + '.map.json'),($stage + '.build.json'))) { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force } }
    # Do not block the caller on global COM finalizers after Excel has exited.
    [GC]::Collect()
}
