#requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^\d+\.\d+\.\d+$')][string]$Version,
    [string]$SourceRoot = (Split-Path $PSScriptRoot -Parent),
    [Parameter(Mandatory)][string]$OutputDirectory
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$root=[IO.Path]::GetFullPath($SourceRoot)
$destination=[IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $destination) { throw 'Release preparation never overwrites an existing directory. Choose a new candidate location.' }
$sourceManifest=Join-Path $root 'release-manifest.json'
$manifest=Get-Content -Raw -LiteralPath $sourceManifest | ConvertFrom-Json -AsHashtable
if ($manifest.schemaVersion -ne 2 -or $manifest.status -cne 'development' -or $manifest.templateId -cne 'xf1-base-six-sheet' -or $manifest.generation -cne 'from-scratch') { throw 'Expected a development six-sheet manifest.' }
$seen=[collections.generic.hashset[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($asset in $manifest.assets) {
    if ($asset.path -cnotmatch '^(docs/[a-z][a-z0-9-]*\.md|runner/(New-XF1Workbook|Workbook|Layout)\.ps1)$' -or -not $seen.Add($asset.path)) { throw 'Invalid or duplicate release asset path.' }
    if ((Get-FileHash -LiteralPath (Join-Path $root $asset.path) -Algorithm SHA256).Hash -ine $asset.sha256) { throw "Source asset hash mismatch: $($asset.path)" }
}
foreach ($required in @('runner/New-XF1Workbook.ps1','runner/Workbook.ps1','runner/Layout.ps1')) { if (-not $seen.Contains($required)) { throw "Missing $required" } }
$null=[IO.Directory]::CreateDirectory($destination)
$package=Join-Path $destination 'package'
$null=[IO.Directory]::CreateDirectory($package)
foreach ($asset in $manifest.assets) {
    $target=Join-Path $package $asset.path
    $null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
    [IO.File]::Copy((Join-Path $root $asset.path),$target,$false)
    if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ine $asset.sha256) { throw 'Source changed during packaging; candidate is incomplete.' }
}
$sourceVersion=$manifest.version
$manifest.version=$Version
# This describes a fixed release artifact, not approval/publication. Only MCP's selection can approve it.
$manifest.status='release'
$manifestPath=Join-Path $package 'release-manifest.json'
$manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $manifestPath -Encoding utf8
$archive=Join-Path $destination "xf1-$Version.zip"
$zip=[IO.Compression.ZipFile]::Open($archive,[IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($path in @('release-manifest.json')+@($manifest.assets.path)) {
        $null=[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip,(Join-Path $package $path),$path,[IO.Compression.CompressionLevel]::Optimal)
    }
} finally { $zip.Dispose() }
$bridge=Join-Path $destination 'Invoke-XF1Build.ps1'
[IO.File]::Copy((Join-Path $PSScriptRoot 'Invoke-XF1Build.ps1'),$bridge,$false)
$base="https://github.com/XF1-Advisory-Services/xf1-mcp/releases/download/v$Version/"
$descriptor=[ordered]@{
    schemaVersion=1;approval='candidate';templateId=$manifest.templateId;version=$Version;generation='from-scratch';runner=$manifest.runner
    manifestSha256=(Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
    package=@{url=($base+"xf1-$Version.zip");sha256=(Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash;bytes=(Get-Item -LiteralPath $archive).Length}
    bridge=@{version='1.0.0';url=($base+'Invoke-XF1Build.ps1');sha256=(Get-FileHash -LiteralPath $bridge -Algorithm SHA256).Hash}
    requirements=$manifest.requirements;testedClient='Codex';otherClients='Claude compatibility expected, not guaranteed or tested'
}
$descriptor | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $destination 'release-candidate.json') -Encoding utf8
[ordered]@{sourceVersion=$sourceVersion;sourceManifestSha256=(Get-FileHash -LiteralPath $sourceManifest -Algorithm SHA256).Hash;preparedAt=[datetime]::UtcNow.ToString('o');published=$false;approved=$false} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $destination 'preparation.json') -Encoding utf8
Write-Output "Prepared $Version from $sourceVersion at $destination. Not approved, published or deployed."
