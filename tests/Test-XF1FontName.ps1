#requires -Version 7.0
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'runner/Workbook.ps1')
foreach ($returnedName in @('Verdana', 'verdana', 'VERDANA')) {
    Assert-XF1FontName $returnedName 'Font name'
}
$rejected = $false
try { Assert-XF1FontName 'Arial' 'Font name' } catch { $rejected = $true }
if (-not $rejected) { throw 'A different font must fail verification.' }
Write-Output 'Font-name verification accepts capitalization differences and rejects another font.'
