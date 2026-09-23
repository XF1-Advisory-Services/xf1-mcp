# Test fixture only: the bridge must reject this archive without extracting it.
#requires -Version 7.0
param([Parameter(Mandatory)][string]$OutputPath)
$ErrorActionPreference='Stop'
$zip=[IO.Compression.ZipFile]::Open($OutputPath,[IO.Compression.ZipArchiveMode]::Create)
try {
    $entry=$zip.CreateEntry('../escape.ps1')
    $writer=[IO.StreamWriter]::new($entry.Open())
    try { $writer.Write('# traversal fixture, never executed') } finally { $writer.Dispose() }
} finally { $zip.Dispose() }
