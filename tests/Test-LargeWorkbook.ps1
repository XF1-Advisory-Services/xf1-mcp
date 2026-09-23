#requires -Version 7.0
[CmdletBinding()]
param([string]$OutputDirectory=(Join-Path (Split-Path $PSScriptRoot -Parent) ('.build/benchmark-from-scratch/'+(Get-Date -Format 'yyyyMMdd-HHmmss'))))
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'runner/Workbook.ps1')
. (Join-Path $root 'runner/Layout.ps1')
$null=[IO.Directory]::CreateDirectory($OutputDirectory)
$output=Join-Path $OutputDirectory 'large-from-scratch.xlsx'
if (Test-Path -LiteralPath $output) {throw 'Benchmark output already exists.'}
$config=@{start=[datetime]'2024-04-01';cutoff=[datetime]'2025-04-30';fyEnd=3;months=73;currency='INR';actualLabel='A';forecastLabel='F'}
$map=New-XF1LayoutMap $config
$watch=[diagnostics.stopwatch]::StartNew()
$timings=[ordered]@{}
Add-Type -TypeDefinition @'
public static class XF1Fixture {
  public static object[,] TextBlock(int rows, int cols, int seed) {
    var result = new object[rows, cols]; var rng = new System.Random(seed); var bytes = new byte[48];
    for (int r=0;r<rows;r++) for (int c=0;c<cols;c++) {rng.NextBytes(bytes); result[r,c]=System.Convert.ToBase64String(bytes);}
    return result;
  }
}
'@

$excel=$book=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.ScreenUpdating=$false
    $book=$excel.Workbooks.Add(-4167);$excel.Calculation=-4135
    $null=New-XF1Sheets $book $excel $config $map
    $null=Set-XF1Configuration $book $config $map
    $timings.baseConstructionSeconds=[math]::Round($watch.Elapsed.TotalSeconds,3)
    $payload=$book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count))
    $payload.Name='Benchmark_Data'
    $first=$last=$null
    for ($r=1;$r -le 80000;$r+=2000) {
        $block=[XF1Fixture]::TextBlock(2000,10,$r)
        $payload.Range("A${r}:J$($r+1999)").NumberFormat='@'
        $payload.Range("A${r}:J$($r+1999)").Value2=$block
        if ($r -eq 1) {$first=$block[0,0]}
        $last=$block[1999,9]
    }
    $payload.Range('K1:O80000').Formula2R1C1='=ROW()*COLUMN()'
    $timings.payloadConstructionSeconds=[math]::Round($watch.Elapsed.TotalSeconds-$timings.baseConstructionSeconds,3)
    [void]$excel.Calculate()
    $null=Set-XF1LabelLayout $book $config $map
    $null=Test-XF1Workbook $book $excel $config $map
    if ($payload.Range('A1').Value2 -cne $first -or $payload.Range('J80000').Value2 -cne $last) {throw 'Payload text mismatch.'}
    if ($excel.WorksheetFunction.CountA($payload.Range('A1:J80000')) -ne 800000) {throw 'Payload cells missing.'}
    if ($payload.Range('K1:O80000').HasFormula -ne $true) {throw 'Payload formulas missing.'}
    if ($excel.WorksheetFunction.Sum($payload.Range('K1:O80000')) -ne (80000.0*80001/2*65)) {throw 'Payload calculation mismatch.'}
    $timings.calculateVerifySeconds=[math]::Round($watch.Elapsed.TotalSeconds-$timings.baseConstructionSeconds-$timings.payloadConstructionSeconds,3)
    [void]$book.Worksheets.Item('Controls').Activate()
    $excel.Calculation=-4105
    [void]$book.SaveAs($output,51);[void]$book.Close($false);$book=$null
    $timings.saveCloseSeconds=[math]::Round($watch.Elapsed.TotalSeconds-$timings.baseConstructionSeconds-$timings.payloadConstructionSeconds-$timings.calculateVerifySeconds,3)
} finally {
    if ($null -ne $book) {[void]$book.Close($false)}
    if ($null -ne $excel) {[void]$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect()
}
$bytes=(Get-Item -LiteralPath $output).Length
if ($bytes -lt 30MB) {throw "Benchmark output below 30 MiB: $bytes bytes"}
$result=@{passed=$true;generation='from-scratch';resultBytes=$bytes;totalSeconds=[math]::Round($watch.Elapsed.TotalSeconds,2);timings=$timings;payloadTextCells=800000;payloadFormulaCells=400000;scope='Native construction from an empty workbook using the production layout/calendar functions, with test-only synthetic payload. Includes payload creation; not comparable to the previous copy-and-configure benchmark or a guarantee for complex financial models.'}
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'results.json') -Encoding utf8
$result | ConvertTo-Json -Depth 8
