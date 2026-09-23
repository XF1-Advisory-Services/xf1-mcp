#requires -Version 7.0
[CmdletBinding()]
param([string]$PackageRoot=(Split-Path $PSScriptRoot -Parent))
$ErrorActionPreference='Stop'
$path=Join-Path $PackageRoot 'release-manifest.json'
$manifest=Get-Content -Raw -LiteralPath $path | ConvertFrom-Json -AsHashtable
if ($manifest.status -ne 'development') {throw 'Released manifests are immutable. Create a development version before changing assets.'}
$files=@('runner/New-XF1Workbook.ps1','runner/Workbook.ps1','runner/Layout.ps1')
$files+=@(Get-ChildItem -LiteralPath (Join-Path $PackageRoot 'docs') -Filter '*.md' -File | ForEach-Object {'docs/'+$_.Name})
$manifest.assets=@($files | Sort-Object -Unique | ForEach-Object {@{path=$_;sha256=(Get-FileHash -LiteralPath (Join-Path $PackageRoot $_) -Algorithm SHA256).Hash}})
$manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $path -Encoding utf8
Write-Output "Updated development hashes: $($manifest.version). Nothing published."
