#requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory)][uri]$McpUrl,
    [string]$RequestPath,
    [string]$CacheRoot = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'XF1/releases'),
    [switch]$ResolveOnly,
    [switch]$ValidateRequestOnly,
    # Test-only escape hatch: requires a loopback MCP endpoint. Never use for a remote service.
    [switch]$AllowLoopbackForTest
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$bridgeVersion = '1.1.0'
$templateId = 'xf1-base-six-sheet'
$required = @('runner/New-XF1Workbook.ps1','runner/Workbook.ps1','runner/Layout.ps1')

function Assert-XF1Hash([string]$Value) {
    if ($Value -cnotmatch '^[a-fA-F0-9]{64}$') { throw 'Invalid SHA-256 in approved release.' }
}
function Assert-XF1Path([string]$Path) {
    if ($Path -cnotmatch '^(docs/[a-z][a-z0-9-]*\.md|runner/(New-XF1Workbook|Workbook|Layout)\.ps1)$') {
        throw "Unapproved package path: $Path"
    }
}
function Assert-XF1File([string]$Path, [string]$Hash) {
    Assert-XF1Hash $Hash
    $item = Get-Item -LiteralPath $Path -Force
    if ($item.PSIsContainer -or $item.LinkType) { throw "Not a regular package file: $Path" }
    if ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -ine $Hash) { throw "Package hash mismatch: $Path" }
}
function Assert-XF1Package([string]$Directory, $Release) {
    if ((Get-Item -LiteralPath $Directory -Force).LinkType) { throw 'Package links are forbidden.' }
    $manifestPath = Join-Path $Directory 'release-manifest.json'
    Assert-XF1File $manifestPath $Release.manifestSha256
    $manifest = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json -AsHashtable
    if ($manifest.schemaVersion -ne 2 -or $manifest.templateId -cne $templateId -or
        $manifest.version -cne $Release.version -or $manifest.status -cne 'release' -or
        $manifest.generation -cne 'from-scratch' -or $manifest.runner -cne $required[0]) {
        throw 'Package identity does not match the approved release.'
    }
    if ($manifest.assets.Count -lt 3 -or $manifest.assets.Count -gt 64) { throw 'Invalid package asset count.' }
    $seen = [collections.generic.hashset[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $null = $seen.Add('release-manifest.json')
    foreach ($asset in $manifest.assets) {
        Assert-XF1Path $asset.path
        if (-not $seen.Add($asset.path)) { throw 'Duplicate package asset.' }
        Assert-XF1File (Join-Path $Directory $asset.path) $asset.sha256
    }
    foreach ($path in $required) { if (-not $seen.Contains($path)) { throw "Missing runner asset: $path" } }
    foreach ($item in Get-ChildItem -LiteralPath $Directory -Recurse -Force) {
        if ($item.LinkType) { throw 'Package links are forbidden.' }
        if (-not $item.PSIsContainer) {
            $relative = [IO.Path]::GetRelativePath($Directory, $item.FullName).Replace('\','/')
            if (-not $seen.Contains($relative)) { throw "Unexpected cached file: $relative" }
        }
    }
    return $manifest
}
function Read-XF1Rpc($Response, [int]$Id) {
    $body = [string]$Response.Content
    $messages = @()
    if ($body.TrimStart().StartsWith('{')) {
        $messages = @($body | ConvertFrom-Json -AsHashtable)
    } else {
        foreach ($event in ($body -split '\r?\n\r?\n')) {
            $data = @($event -split '\r?\n' | Where-Object { $_.StartsWith('data:') } | ForEach-Object { $_.Substring(5).TrimStart() })
            if ($data.Count) { $messages += (($data -join "`n") | ConvertFrom-Json -AsHashtable) }
        }
    }
    $match = @($messages | Where-Object { $_.ContainsKey('id') -and $_.id -eq $Id })
    if ($match.Count -ne 1 -or $match[0].ContainsKey('error') -or -not $match[0].ContainsKey('result')) { throw 'MCP did not return a valid result. Stop; no cache fallback.' }
    return $match[0].result
}
function Send-XF1Rpc($Body, $Headers) {
    Invoke-WebRequest -Uri $McpUrl -Method Post -ContentType 'application/json' -Headers $Headers `
        -Body ($Body | ConvertTo-Json -Depth 12 -Compress) -TimeoutSec 30 -MaximumRedirection 0
}

if ($AllowLoopbackForTest -and -not $McpUrl.IsLoopback) { throw 'Test mode requires a loopback MCP URL.' }
if ($McpUrl.Scheme -ne 'https' -and -not ($AllowLoopbackForTest -and $McpUrl.Scheme -eq 'http')) { throw 'MCP must use HTTPS.' }
if ($McpUrl.UserInfo -or $McpUrl.Query -or $McpUrl.Fragment) { throw 'MCP URL must not contain credentials, query parameters or fragments.' }
if ($ResolveOnly -and $ValidateRequestOnly) { throw 'Choose only one validation/resolution mode.' }
if (-not $ResolveOnly) {
    if ([string]::IsNullOrWhiteSpace($RequestPath)) { throw 'A fresh -RequestPath JSON file is required for each NEW workbook request. Collect all six inputs; never reuse earlier assumptions implicitly.' }
    $request = Get-Content -Raw -LiteralPath $RequestPath | ConvertFrom-Json -AsHashtable
    if ($request.schemaVersion -ne 1 -or $request.requestId -isnot [string] -or [string]::IsNullOrWhiteSpace($request.requestId) -or $request.inputs -isnot [System.Collections.IDictionary]) { throw 'Invalid request: schemaVersion 1, a fresh requestId and inputs object are required.' }
    $questions = [ordered]@{
        StartDate='What is the model start date (YYYY-MM-DD, first day of its financial year)?'
        LastActualDate='What is the last actuals date (YYYY-MM-DD, month end; for no actuals use the day before the start date)?'
        FinancialYearEndMonth='Which month ends the financial year (1-12)?'
        Currency='Which currency should this workbook use (for example USD or EUR)?'
        Months='How many months should this workbook cover (at least 12)?'
        OutputPath="What is the new workbook's full Windows output path, including its .xlsx filename?"
    }
    $missing = @(); $values = @{}
    foreach ($field in $request.inputs.Keys) {
        if ($field -cnotin @($questions.Keys)+@('ActualLabel','ForecastLabel')) { throw "Unknown request field: $field" }
        $entry = $request.inputs[$field]
        if ($entry -isnot [System.Collections.IDictionary] -or $entry.requestId -cne $request.requestId -or $entry.source -cnotin @('current_request','clarification','explicit_reuse')) { throw "Invalid or stale provenance for $field. Only current-request inputs, its clarifications or explicit reuse are allowed." }
        if ($entry.Contains('value') -and $null -ne $entry.value -and -not [string]::IsNullOrWhiteSpace([string]$entry.value)) { $values[$field]=$entry.value }
    }
    foreach ($field in $questions.Keys) { if (-not $values.ContainsKey($field)) { $missing += $field } }
    if ($missing.Count) {
        $lines = for ($index=0; $index -lt $missing.Count; $index++) { '{0}. {1}' -f ($index+1),$questions[$missing[$index]] }
        throw ("Missing inputs for request $($request.requestId). Ask all together:`n"+($lines -join "`n"))
    }
    foreach ($field in @('StartDate','LastActualDate','Currency','OutputPath')) { if ($values[$field] -isnot [string]) { throw "$field must be text." } }
    foreach ($field in @('Months','FinancialYearEndMonth')) { if ($values[$field] -isnot [long] -and $values[$field] -isnot [int]) { throw "$field must be an integer." } }
    $StartDate=$values.StartDate; $LastActualDate=$values.LastActualDate
    $FinancialYearEndMonth=$values.FinancialYearEndMonth; $Currency=$values.Currency
    $OutputPath=$values.OutputPath; $Months=$values.Months
    $ActualLabel='A'; $ForecastLabel='F'
    if ($values.ContainsKey('ActualLabel')) { $ActualLabel=$values.ActualLabel }
    if ($values.ContainsKey('ForecastLabel')) { $ForecastLabel=$values.ForecastLabel }
    foreach ($label in @($Currency,$ActualLabel,$ForecastLabel)) { if ($label -isnot [string] -or [string]::IsNullOrWhiteSpace($label) -or $label.Length -gt 64) { throw 'Currency and period labels must contain 1-64 characters.' } }
    if ($ActualLabel -eq $ForecastLabel) { throw 'Actual and forecast labels must be distinct.' }
    if ($FinancialYearEndMonth -lt 1 -or $FinancialYearEndMonth -gt 12) { throw 'FinancialYearEndMonth must be 1-12.' }
    if ($Months -lt 12 -or $Months -gt 16358) { throw 'Months must be 12-16358; no horizon default is used by the bridge.' }
    $start=[datetime]::ParseExact($StartDate,'yyyy-MM-dd',[cultureinfo]::InvariantCulture)
    $actual=[datetime]::ParseExact($LastActualDate,'yyyy-MM-dd',[cultureinfo]::InvariantCulture)
    if ($start.Year -lt 1901 -or $start.Day -ne 1 -or $start.Month -ne (($FinancialYearEndMonth % 12)+1)) { throw 'StartDate must be the first day of the financial year, from 1901.' }
    $end=$start.AddMonths($Months).AddDays(-1)
    if ($actual -lt $start.AddDays(-1) -or $actual -gt $end -or $actual.Day -ne [datetime]::DaysInMonth($actual.Year,$actual.Month)) { throw 'LastActualDate must be a month end within the horizon or the day before StartDate.' }
    if (-not [IO.Path]::IsPathFullyQualified($OutputPath) -or [IO.Path]::GetExtension($OutputPath) -ine '.xlsx') { throw 'OutputPath must be a full path with an .xlsx filename.' }
    if ($ValidateRequestOnly) { @{requestId=$request.requestId;ready=$true;requiredFields=@($questions.Keys)} | ConvertTo-Json; return }
    if (-not $IsWindows -or $null -eq [type]::GetTypeFromProgID('Excel.Application')) { throw 'This build requires Windows and installed desktop Microsoft Excel.' }
}

# Always establish current approval first, even when all assets are already cached.
$headers = @{ Accept='application/json, text/event-stream'; 'Cache-Control'='no-cache' }
$init = Read-XF1Rpc (Send-XF1Rpc @{jsonrpc='2.0';id=1;method='initialize';params=@{protocolVersion='2025-06-18';capabilities=@{};clientInfo=@{name='xf1-powershell-bridge';version=$bridgeVersion}}} $headers) 1
if ($init.protocolVersion -notin @('2025-03-26','2025-06-18','2025-11-25')) { throw 'Unsupported negotiated MCP protocol.' }
$headers['MCP-Protocol-Version'] = $init.protocolVersion
$null = Send-XF1Rpc @{jsonrpc='2.0';method='notifications/initialized'} $headers
$result = Read-XF1Rpc (Send-XF1Rpc @{jsonrpc='2.0';id=2;method='tools/call';params=@{name='get_current_release';arguments=@{}}} $headers) 2
if ($result.ContainsKey('isError') -and $result.isError) { throw "MCP cannot establish an approved release: $($result.content[0].text)" }
if (-not $result.ContainsKey('structuredContent')) { throw 'MCP release response has no structured content.' }
$release = $result.structuredContent.release
if ($release.schemaVersion -ne 1 -or $release.approval -cne 'approved' -or
    $release.templateId -cne $templateId -or $release.version -cnotmatch '^\d+\.\d+\.\d+$' -or
    $release.generation -cne 'from-scratch' -or $release.runner -cne $required[0] -or
    $release.package.bytes -lt 1 -or $release.package.bytes -gt 5000000) { throw 'Invalid approved release identity.' }
Assert-XF1Hash $release.manifestSha256
Assert-XF1Hash $release.package.sha256
Assert-XF1Hash $release.bridge.sha256
if ($release.bridge.version -cne $bridgeVersion) { throw "Update the local bridge to version $($release.bridge.version) from the approved release before building." }
Assert-XF1File $PSCommandPath $release.bridge.sha256
$base = "https://github.com/XF1-Advisory-Services/xf1-mcp/releases/download/v$($release.version)/"
foreach ($asset in @(@($release.package.url,"xf1-$($release.version).zip"),@($release.bridge.url,'Invoke-XF1Build.ps1'))) {
    $uri = [uri]$asset[0]
    $localTest = $AllowLoopbackForTest -and $uri.IsLoopback -and $uri.Scheme -eq 'http' -and -not $uri.UserInfo
    if (-not $localTest -and $asset[0] -cne ($base+$asset[1])) { throw 'Release asset is not at the expected immutable GitHub URL.' }
}
$cache = [IO.Path]::GetFullPath($CacheRoot)
$versionDirectory = Join-Path (Join-Path $cache $templateId) $release.version
$packageDirectory = Join-Path $versionDirectory 'package'
$null = [IO.Directory]::CreateDirectory($versionDirectory)
# Reject links along the cache path before reading, writing, or executing cached code.
$ancestor = Get-Item -LiteralPath $versionDirectory -Force
while ($null -ne $ancestor) {
    # OneDrive placeholders are reparse points too; reject actual symbolic links/junctions.
    if ($ancestor.LinkType) { throw 'The cache path must not contain filesystem links.' }
    $ancestor = $ancestor.Parent
}
$lock = $null; $stage = $null
try {
    $lockPath = Join-Path $versionDirectory '.lock'
    if ((Test-Path -LiteralPath $lockPath) -and (Get-Item -LiteralPath $lockPath -Force).LinkType) { throw 'Cache lock must not be a link.' }
    try { $lock = [IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None) }
    catch { throw 'This package version is in use by another local build. Retry after it finishes.' }
    $downloaded = $false
    if (-not (Test-Path -LiteralPath $packageDirectory)) {
        $stage = Join-Path $versionDirectory ('.stage-'+[guid]::NewGuid().ToString('N'))
        $null = [IO.Directory]::CreateDirectory($stage)
        $archive = Join-Path $stage 'package.zip'
        $null = Invoke-WebRequest -Uri $release.package.url -OutFile $archive -TimeoutSec 60 -MaximumRedirection 5
        if ((Get-Item -LiteralPath $archive).Length -ne $release.package.bytes) { throw 'Downloaded package size differs from approved metadata.' }
        Assert-XF1File $archive $release.package.sha256
        $unpacked = Join-Path $stage 'unpacked'
        $null = [IO.Directory]::CreateDirectory($unpacked)
        $zip = [IO.Compression.ZipFile]::OpenRead($archive)
        try {
            $entries = [collections.generic.hashset[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $total = 0L
            foreach ($entry in $zip.Entries) {
                $path = $entry.FullName
                if ($path -cne 'release-manifest.json') { Assert-XF1Path $path }
                if (-not $entries.Add($path) -or $entries.Count -gt 65) { throw 'Duplicate or excessive ZIP entries.' }
                $total += $entry.Length
                if ($total -gt 10000000 -or (($entry.ExternalAttributes -shr 16) -band 0xF000) -eq 0xA000) { throw 'Oversized or linked ZIP entry.' }
                $target = Join-Path $unpacked $path
                $null = [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
                [IO.Compression.ZipFileExtensions]::ExtractToFile($entry,$target,$false)
            }
        } finally { $zip.Dispose() }
        $null = Assert-XF1Package $unpacked $release
        [IO.Directory]::Move($unpacked,$packageDirectory)
        $downloaded = $true
    }
    $null = Assert-XF1Package $packageDirectory $release
    $summary = [ordered]@{releaseVersion=$release.version;manifestSha256=$release.manifestSha256;packageDirectory=$packageDirectory;downloaded=$downloaded}
    if (-not $ResolveOnly) {
        $arguments = @{StartDate=$StartDate;LastActualDate=$LastActualDate;FinancialYearEndMonth=$FinancialYearEndMonth;Currency=$Currency;OutputPath=$OutputPath;Months=$Months;ActualLabel=$ActualLabel;ForecastLabel=$ForecastLabel;ManifestPath=(Join-Path $packageDirectory 'release-manifest.json')}
        $raw = & (Join-Path $packageDirectory $release.runner) @arguments
        $summary.build = ($raw -join [Environment]::NewLine) | ConvertFrom-Json -AsHashtable
        $summary.requestId = $request.requestId
    }
    $summary | ConvertTo-Json -Depth 20
} finally {
    if ($stage -and (Test-Path -LiteralPath $stage)) {
        $resolved = [IO.Path]::GetFullPath($stage)
        if (-not $resolved.StartsWith($versionDirectory+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Staging cleanup escaped its version directory.' }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
    if ($lock) { $lock.Dispose() }
}
