# Native construction of the documented six-sheet layout. No workbook file is read.
function Get-XF1Colour([string]$Hex) {
    $h=$Hex.TrimStart('#')
    return [Convert]::ToInt32($h.Substring(0,2),16)+256*[Convert]::ToInt32($h.Substring(2,2),16)+65536*[Convert]::ToInt32($h.Substring(4,2),16)
}
function Set-XF1Labels($Sheet,[string]$Address,[object[]]$Labels) {
    $values=[object[,]]::new($Labels.Count,1)
    for ($i=0;$i -lt $Labels.Count;$i++) {$values[$i,0]=[string]$Labels[$i]}
    $Sheet.Range($Address).NumberFormat='@'
    $Sheet.Range($Address).Value2=$values
}
function Set-XF1Input($Range) {
    $Range.Font.Color=Get-XF1Colour '#1950FF';$Range.Interior.Color=Get-XF1Colour '#FFFF66'
    $Range.Borders.LineStyle=1;$Range.Borders.Weight=1;$Range.Borders.ColorIndex=-4105
}
function Set-XF1Band($Sheet,[int]$Row,[string]$Label,[string]$Last,[int]$Colour) {
    $Sheet.Range("A${Row}:${Last}${Row}").Interior.Color=$Colour
    $Sheet.Range("A${Row}:${Last}${Row}").Font.Bold=$true
    $Sheet.Rows.Item($Row).RowHeight=21
    $Sheet.Range("A${Row}").Value2=$Label
}
function New-XF1LayoutMap($Config) {
    # This is generated metadata for the layout below, not an input map of a saved file.
    return @{
        schemaVersion=1;configuration=@{months=$Config.months};definedNames=@{'model.currency'='Controls!$D$73';'model.check.message'='Checks!$C$40'}
        formatting='shared-guidance.md';sheets=@{
            Controls=@{spec='controls.md';controls='D9:D17';currency='D73';calendarMonths='C76:D87';sectionRows=@(7,41,71);groups=@('8:39','42:69','72:87');exceptions=@{rowHeights=@{'59'=30;'65'=30;'75'=30}};quarterBoundaryFlags=@{startAnchor='=IF(MOD(Z$23-1,3)=0,1,0)';endAnchor='=IF(MOD(Z$23,3)=0,1,0)'}}
            Monthly=@{spec='monthly.md';groups=@('8:22');newSectionsFrom=24}
            Quarterly=@{spec='quarterly.md';groups=@('8:10');newSectionsFrom=12;exceptions=@{rowHeights=@{'10'=30}}}
            Annual=@{spec='annual.md';groups=@('8:10');newSectionsFrom=12;freezePane='Z6';exceptions=@{rowHeights=@{'10'=30}}}
            Input_sheet=@{spec='inputs.md';body='blank';groups=@();newSectionsFrom=7}
            Checks=@{spec='checks.md';inherits='Monthly header and first eleven flags (rows 10:20)';sectionRows=@(7,22);groups=@('8:20','23:40');messageInput='C38';messageOutput='C40';messageFormula='=C38';configuredChecks=0;failedCheckCount='E35 (blank)'}
        }
        validations=@{'Controls!D9'='First day of financial year';'Controls!D10'='Month end, start minus one day through final model date';'Controls!D12'='Whole number 12–16358; changing horizon requires rebuilding';'Controls!D14'='Whole number 1–12';'Controls!D16:D17'='Nonblank, distinct labels'}
    }
}
function New-XF1Sheets($Book,$Excel,$Config,$Map) {
    $n=$Config.months;$q=[int][math]::Ceiling($n/3);$y=[int][math]::Ceiling($n/12)
    $integer='#,##0;[Red]-#,##0;"-"';$date='[$-en-US]mmm/d/yy;[$-en-US]mmm/d/yy;"-";@'
    $layouts=@(@('Controls',$n,87,'#FFD966'),@('Monthly',$n,22,'#00B0F0'),@('Quarterly',$q,10,'#92D050'),@('Annual',$y,10,'#92D050'),@('Input_sheet',$n,5,'#FFD966'),@('Checks',$n,40,'#F1A983'))
    $Book.Date1904=$false
    $Book.Styles.Item('Normal').Font.Name='Verdana';$Book.Styles.Item('Normal').Font.Size=10
    foreach ($layout in $layouts) {
        $name=$layout[0];$count=$layout[1];$rows=$layout[2];$colour=Get-XF1Colour $layout[3]
        if ($name -eq 'Controls') {$s=$Book.Worksheets.Item(1)} else {$s=$Book.Worksheets.Add([Type]::Missing,$Book.Worksheets.Item($Book.Worksheets.Count))}
        $s.Name=$name;$s.Tab.Color=$colour
        $last=Get-XF1Column (25+$count);$spacer=Get-XF1Column (26+$count);$hidden=Get-XF1Column (27+$count)
        $body=$s.Range("A1:${last}${rows}")
        $body.Font.Name='Verdana';$body.Font.Size=10;$body.Font.Color=0;$body.RowHeight=18;$body.VerticalAlignment=-4108
        foreach ($width in @(@('A:B',0.94),@('C:C',29.11),@('D:E',11.56),@('F:W',0.94),@('X:X',9.67),@('Y:Y',0.94))) {$s.Columns.Item($width[0]).ColumnWidth=[double]$width[1]}
        $s.Columns.Item("Z:${last}").ColumnWidth=11.56
        $s.Columns.Item("${spacer}:${spacer}").ColumnWidth=1
        $s.Columns.Item("${hidden}:XFD").Hidden=$true
        $headerEnd=$last
        if ($name -eq 'Input_sheet') {$s.Columns.Item("F:${spacer}").ColumnWidth=11.56;$headerEnd=$spacer}
        $s.Range("A1:${headerEnd}5").Interior.Color=0;$s.Range("A1:${headerEnd}5").Font.Color=16777215
        $s.Range('C2').Formula2='=model.check.message';$s.Range('C2').Font.Bold=$true
        $s.Range("C1:C${rows}").HorizontalAlignment=-4131;$s.Range("X1:X${rows}").HorizontalAlignment=-4131
        $s.Range("Z1:${last}${rows}").HorizontalAlignment=-4152
        $s.Range("Z1:${last}${rows}").NumberFormat=$integer
        $s.Outline.SummaryRow=0;$s.Outline.SummaryColumn=-4131
        if ($name -ne 'Input_sheet') {[void]$s.Columns.Item('F:W').Group()}
        foreach ($group in $Map.sheets[$name].groups) {[void]$s.Rows.Item($group).Group()}
        if ($name -eq 'Controls') {
            Set-XF1Band $s 7 'TIME' $last $colour;Set-XF1Band $s 41 'MODEL FLAGS' $last $colour;Set-XF1Band $s 71 'LISTS' $last $colour
        } elseif ($name -ne 'Input_sheet') {Set-XF1Band $s 7 'MODEL FLAGS' $last $colour}
        if ($name -eq 'Checks') {Set-XF1Band $s 22 'CHECKS' $last $colour}
        [void]$s.Activate();$Excel.ActiveWindow.DisplayGridlines=$false;$Excel.ActiveWindow.Zoom=90
        [void]$s.Range('A1').Select()
    }
    $c=$Book.Worksheets.Item('Controls')
    Set-XF1Labels $c 'C9:C17' @('Timeline start date','Last actual date','Forecast start date','Timeline length','Final model date','Financial year-end month','Last actual month position','Actual period label','Forecast period label')
    Set-XF1Labels $c 'X9:X17' @('Date','Date','Date','#months','Date','Month','#','Label','Label')
    foreach ($item in @(@('C19','Monthly calendar'),@('C27','Quarterly calendar'),@('C33','Annual calendar'),@('C43','Monthly flags'),@('C58','Quarterly flags'),@('C64','Annual flags'))) {
        $c.Range($item[0]).Value2=[string]$item[1];$c.Range($item[0]).Font.Bold=$true
    }
    Set-XF1Labels $c 'C20:C25' @('Start date','End date','Period type','Period counter','Financial year','Financial quarter')
    Set-XF1Labels $c 'X20:X25' @('Date','Date','Label','#','Year','Label')
    Set-XF1Labels $c 'C28:C31' @('Start date','End date','Period counter','Financial quarter')
    Set-XF1Labels $c 'X28:X31' @('Date','Date','#','Label')
    Set-XF1Labels $c 'C34:C38' @('Start date','End date','Forecast months','Period type','Period counter')
    Set-XF1Labels $c 'X34:X38' @('Date','Date','#','Label','#')
    $flags=@('Month','Days','Actuals flag','Forecast flag','First forecast flag','Forecast counter','Last actual flag','Year counter','Financial year end','Quarter counter','Forecast year counter','Quarter start flag','Quarter end flag')
    $units=@('#','#','1/0','1/0','1/0','#','1/0','#','1/0','#','#','1/0','1/0')
    Set-XF1Labels $c 'C44:C56' $flags;Set-XF1Labels $c 'X44:X56' $units
    $c.Range('C59').Value2='Quarter-end monthly position';$c.Range('C65').Value2='Year-end monthly position'
    $c.Range('X59').Value2='#';$c.Range('X65').Value2='#'
    foreach ($row in @(59,65)) {$c.Range("C${row}").WrapText=$true;$c.Rows.Item($row).RowHeight=30}
    $mLast=Get-XF1Column (25+$n);$qLast=Get-XF1Column (25+$q);$yLast=Get-XF1Column (25+$y)
    foreach ($address in @('D9:D11','D13',"Z20:${mLast}21","Z28:${qLast}29","Z34:${yLast}35")) {$c.Range($address).NumberFormat=$date}
    foreach ($address in @('D12','D14:D15','D76:D87')) {$c.Range($address).NumberFormat=$integer}
    $c.Range("Z24:${mLast}24").NumberFormat='0'
    $c.Range('E12').Value2='Rebuild if changed';$c.Range('E12').Font.Color=Get-XF1Colour '#C00000'
    $c.Range('C73').Value2='Model currency';Set-XF1Labels $c 'C75:C75' @('Calendar month');$c.Range('D75').Value2='Month number'
    $c.Range('C75:D75').Font.Bold=$true;$c.Range('D75').WrapText=$true;$c.Rows.Item(75).RowHeight=30
    $months=[object[,]]::new(12,2)
    for ($i=0;$i -lt 12;$i++) {$months[$i,0]=[cultureinfo]::InvariantCulture.DateTimeFormat.GetMonthName($i+1);$months[$i,1]=[double]($i+1)}
    $c.Range('C76:D87').Value2=$months
    foreach ($address in @('D9','D10','D12','D14','D16:D17','D73','C76:D87')) {Set-XF1Input $c.Range($address)}
    $c.Range('D9').Validation.Add(7,1,1,'=AND(ISNUMBER(D9),DAY(D9)=1,MONTH(D9)=MOD(D14,12)+1)')
    $c.Range('D10').Validation.Add(7,1,1,'=AND(ISNUMBER(D10),D10=EOMONTH(D10,0),D10>=D9-1,D10<=D13)')
    $c.Range('D12').Validation.Add(1,1,1,12,16358);$c.Range('D14').Validation.Add(1,1,1,1,12)
    $c.Range('D16:D17').Validation.Add(7,1,1,'=AND(LEN($D$16)>0,LEN($D$17)>0,$D$16<>$D$17)')
    foreach ($address in @('D9','D10','D12','D14','D16:D17')) {$c.Range($address).Validation.IgnoreBlank=$false;$c.Range($address).Validation.ShowError=$true}
    $c.Range('D12').Validation.InputTitle='Timeline length';$c.Range('D12').Validation.InputMessage='Generate a new workbook to change the timeline length.';$c.Range('D12').Validation.ShowInput=$true
    foreach ($name in @('Monthly','Checks')) {
        $s=$Book.Worksheets.Item($name);$endFlag=if ($name -eq 'Monthly') {22} else {20}
        Set-XF1Labels $s 'X1:X5' @('Date','Date','Label','Counter','Year')
        $s.Range("Z1:${mLast}2").NumberFormat=$date;$s.Range("Z5:${mLast}5").NumberFormat='0'
        $s.Range('C9').Value2='Monthly flags';$s.Range('C9').Font.Bold=$true
        Set-XF1Labels $s "C10:C${endFlag}" $flags[0..($endFlag-10)];Set-XF1Labels $s "X10:X${endFlag}" $units[0..($endFlag-10)]
    }
    foreach ($pair in @(@('Quarterly',$qLast,'Quarterly flags','Quarter-end monthly position'),@('Annual',$yLast,'Annual flags','Year-end monthly position'))) {
        $s=$Book.Worksheets.Item($pair[0]);$s.Range("Z1:$($pair[1])2").NumberFormat=$date
        if ($pair[0] -eq 'Quarterly') {Set-XF1Labels $s 'X1:X4' @('Date','Date','Counter','Label')} else {Set-XF1Labels $s 'X1:X5' @('Date','Date','#months','Label','Counter')}
        $s.Range('C9').Value2=[string]$pair[2];$s.Range('C9').Font.Bold=$true;$s.Range('C10').Value2=[string]$pair[3]
        $s.Range('C10').WrapText=$true;$s.Rows.Item(10).RowHeight=30;$s.Range('X10').Value2='#'
    }
    $ch=$Book.Worksheets.Item('Checks')
    foreach ($item in @(@('C27','No checks configured'),@('C35','Failed check types'),@('C37','Check messages'),@('C38','Checks not configured'))) {$ch.Range($item[0]).Value2=[string]$item[1]}
    $ch.Range('C37').Font.Bold=$true;Set-XF1Input $ch.Range('C38');$ch.Range('C40').Formula2='=C38'
    $null=$Book.Names.Add('model.currency','=Controls!$D$73');$null=$Book.Names.Add('model.check.message','=Checks!$C$40')
}
