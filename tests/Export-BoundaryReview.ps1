#requires -Version 7.0
[CmdletBinding()]
param([Parameter(Mandatory)][string]$WorkbookPath,[Parameter(Mandatory)][string]$PdfPath)
$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'runner/Workbook.ps1')
$excel=$book=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
    $book=$excel.Workbooks.Open([IO.Path]::GetFullPath($WorkbookPath),0,$true)
    $n=[int]$book.Worksheets.Item('Controls').Range('D12').Value2
    if ($n -gt 24) {throw 'Use a short test horizon (12–24 months) for readable boundary previews.'}
    $m=Get-XF1Column (26+$n);$q=Get-XF1Column (26+[int][math]::Ceiling($n/3));$y=Get-XF1Column (26+[int][math]::Ceiling($n/12))
    $areas=@{Controls="A7:${m}38,A41:${m}65";Monthly="A1:${m}24";Quarterly="A1:${q}12";Annual="A1:${y}12";Input_sheet="A1:${m}8";Checks="A1:${m}40"}
    foreach ($s in $book.Worksheets) {
        $s.PageSetup.PrintArea=$areas[$s.Name]
        $s.PageSetup.Orientation=2;$s.PageSetup.PaperSize=9
        $s.PageSetup.Zoom=$false;$s.PageSetup.FitToPagesWide=1;$s.PageSetup.FitToPagesTall=1
        $s.PageSetup.LeftMargin=10;$s.PageSetup.RightMargin=10
    }
    [void]$book.ExportAsFixedFormat(0,[IO.Path]::GetFullPath($PdfPath))
    [void]$book.Close($false);$book=$null
} finally {
    if ($null -ne $book) {[void]$book.Close($false)}
    if ($null -ne $excel) {[void]$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
