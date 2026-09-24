# Shared implementation for the six-sheet XF1 template. No Excel session is created here.
function Get-XF1Column([int]$Number) {
    $text = ''
    while ($Number -gt 0) { $Number--; $text = [char](65 + ($Number % 26)) + $text; $Number = [int][math]::Floor($Number / 26) }
    return $text
}
function Assert-XF1Equal($Actual,$Expected,[string]$Label) {
    if ($Actual -cne $Expected) { throw "$Label expected [$Expected], got [$Actual]" }
}
function Assert-XF1FontName($Actual,[string]$Label) {
    if ($Actual -ine 'Verdana') { throw "$Label expected [Verdana], got [$Actual]" }
}
function Assert-XF1Base($Book,$Map,$Manifest) {
    $names = @($Book.Worksheets | ForEach-Object { $_.Name })
    Assert-XF1Equal ($names -join '|') ($Manifest.sheetOrder -join '|') 'Sheet order'
    Assert-XF1Equal ($names[0..5] -join '|') 'Controls|Monthly|Quarterly|Annual|Input_sheet|Checks' 'Required sheets'
    Assert-XF1Equal $Book.Date1904 $false 'Date system'
    $c = $Book.Worksheets.Item('Controls')
    Assert-XF1Equal $c.Range('D12').Value2 $Map.configuration.months 'Base horizon/map'
    Assert-XF1Equal $c.Range('C55').Value2 'Quarter start flag' 'Base quarter start'
    Assert-XF1Equal $c.Range('C56').Value2 'Quarter end flag' 'Base quarter end'
    Assert-XF1Equal $Book.Names.Item('model.currency').RefersTo '=Controls!$D$73' 'Currency name'
    Assert-XF1Equal $Book.Names.Item('model.check.message').RefersTo '=Checks!$C$40' 'Status name'
}
function Set-XF1Row($Sheet,[int]$Row,[int]$Count,[string]$Formula) {
    $last = Get-XF1Column (25 + $Count)
    $Sheet.Range("Z${Row}:${last}${Row}").Formula2R1C1 = $Formula
}
function Set-XF1Configuration($Book,$Config,$Map) {
    $n = $Config.months; $q = [int][math]::Ceiling($n/3); $y = [int][math]::Ceiling($n/12)
    $c = $Book.Worksheets.Item('Controls')
    $c.Range('D9').Value2 = $Config.start.ToOADate(); $c.Range('D10').Value2 = $Config.cutoff.ToOADate()
    $c.Range('D12').Value2 = [double]$n; $c.Range('D14').Value2 = [double]$Config.fyEnd
    # Text-first prevents dates or formulas being inferred from unit/period labels.
    $c.Range('D16:D17').NumberFormat = '@'; $c.Range('D73').NumberFormat = '@'
    $c.Range('D16').Value2 = [string]$Config.actualLabel; $c.Range('D17').Value2 = [string]$Config.forecastLabel; $c.Range('D73').Value2 = [string]$Config.currency
    $c.Range('D11').Formula2 = '=D10+1'; $c.Range('D13').Formula2 = '=EOMONTH(D9,D12-1)'
    $last = Get-XF1Column (25+$n)
    $c.Range('D15').Formula2 = ('=COUNTIF($Z$21:$' + $last + '$21,"<="&D10)')
    $monthRows = @{
        20='=R[1]C[-1]+1'; 21='=EOMONTH(R[-1]C,0)'; 22='=IF(R[1]C<=R15C4,R16C4,R17C4)'; 23='=COLUMN()-25'
        24='=YEAR(EOMONTH(R9C4,11))+INT((R[-1]C-1)/12)'; 25='="Q"&(1+INT(MOD(R[-2]C-1,12)/3))&" "&R[-1]C'
        44='=MONTH(R20C)';45='=DAY(R21C)';46='=IF(R23C<=R15C4,1,0)';47='=1-R[-1]C'
        48='=IF(AND(R23C=R15C4+1,R15C4<R12C4),1,0)';49='=MAX(0,R23C-R15C4)';50='=IF(AND(R23C=R15C4,R15C4>0),1,0)'
        51='=1+INT((R23C-1)/12)';52='=IF(R44C=R14C4,1,0)';53='=1+INT((R23C-1)/3)'
        54='=IF(R46C=1,0,R51C-INT(R15C4/12))';55='=IF(MOD(R23C-1,3)=0,1,0)';56='=IF(MOD(R23C,3)=0,1,0)'
    }
    foreach ($row in $monthRows.Keys) { Set-XF1Row $c $row $n $monthRows[$row] }
    $c.Range('Z20').Formula2 = '=$D$9'
    $endCol = 25+$n
    $quarterRows = @{
        28="=INDEX(R20C26:R20C${endCol},1,3*R30C-2)"
        29="=INDEX(R21C26:R21C${endCol},1,MIN(3*R30C,R12C4))"
        30='=COLUMN()-25'
        31="=INDEX(R25C26:R25C${endCol},1,MIN(3*R30C,R12C4))&IF(MIN(3*R30C,R12C4)-(3*R30C-2)+1<3,`" (partial)`",`"`")"
        59='=MIN(3*R30C,R12C4)'
    }
    $annualRows = @{
        34="=INDEX(R20C26:R20C${endCol},1,12*R38C-11)"
        35="=INDEX(R21C26:R21C${endCol},1,MIN(12*R38C,R12C4))"
        36='=MAX(0,MIN(12*R38C,R12C4)-MAX(12*(R38C-1),R15C4))'
        37='=IF(R36C=0,R16C4,IF(R36C=MIN(12,R12C4-12*(R38C-1)),R17C4,R16C4&(MIN(12,R12C4-12*(R38C-1))-R36C)&"+"&R17C4&R36C))&IF(MIN(12,R12C4-12*(R38C-1))<12," (partial)","")'
        38='=COLUMN()-25';65='=MIN(12*R38C,R12C4)'
    }
    foreach ($pair in @(@($quarterRows,$q),@($annualRows,$y))) {
        foreach ($row in $pair[0].Keys) {
            Set-XF1Row $c $row $pair[1] $pair[0][$row]

        }
    }
    foreach ($name in @('Monthly','Checks')) {
        $s = $Book.Worksheets.Item($name)
        for ($r=1;$r -le 5;$r++) { Set-XF1Row $s $r $n "='Controls'!R$($r+19)C" }
        $lastFlag = if ($name -eq 'Monthly') {22} else {20}
        for ($r=10;$r -le $lastFlag;$r++) { Set-XF1Row $s $r $n "='Controls'!R$($r+34)C" }
    }
    foreach ($pair in @(@('Quarterly',$q,27,4,59),@('Annual',$y,33,5,65))) {
        $s = $Book.Worksheets.Item($pair[0])
        for ($r=1;$r -le $pair[3];$r++) { Set-XF1Row $s $r $pair[1] "='Controls'!R$($r+$pair[2])C" }
        Set-XF1Row $s 10 $pair[1] "='Controls'!R$($pair[4])C"
    }
}

function Set-XF1LabelLayout($Book,$Config,$Map) {
    $n=$Config.months;$q=[int][math]::Ceiling($n/3);$y=[int][math]::Ceiling($n/12)
    # Fit only text-label rows after calculation so partial/custom labels remain legible.
    foreach ($item in @(@('Controls',22,$n),@('Controls',31,$q),@('Controls',37,$y),@('Monthly',3,$n),@('Checks',3,$n),@('Quarterly',4,$q),@('Annual',4,$y))) {
        $s=$Book.Worksheets.Item($item[0]);$row=$item[1];$last=Get-XF1Column (25+$item[2])
        $range=$s.Range("Z${row}:${last}${row}")
        $range.WrapText=$true
        [void]$range.Rows.AutoFit()
        if ($s.Rows.Item($row).RowHeight -lt 18) {$s.Rows.Item($row).RowHeight=18}
        $height=$s.Rows.Item($row).RowHeight
        if ($height -gt 18) {
            $entry=$Map.sheets[$s.Name]
            if (-not $entry.ContainsKey('exceptions')) {$entry.exceptions=@{}}
            if (-not $entry.exceptions.ContainsKey('rowHeights')) {$entry.exceptions.rowHeights=@{}}
            $entry.exceptions.rowHeights[[string]$row]=$height
        }
    }
    # Excel stores pane positions geometrically; reset them after changing header heights.
    foreach ($name in @('Controls','Monthly','Quarterly','Annual','Input_sheet','Checks')) {
        $s=$Book.Worksheets.Item($name);[void]$s.Activate();$win=$Book.Windows.Item(1)
        $win.FreezePanes=$false;$win.SplitRow=5
        $win.SplitColumn=$(if ($name -eq 'Input_sheet') {5} else {25})
        $win.FreezePanes=$true;$win.ScrollRow=1;$win.ScrollColumn=1
    }
}

function Test-XF1Workbook($Book,$Excel,$Config,$Map) {
    $n = $Config.months; $last = Get-XF1Column (25+$n)
    $h = ($Config.cutoff.Year-$Config.start.Year)*12+$Config.cutoff.Month-$Config.start.Month+1
    $c = $Book.Worksheets.Item('Controls'); $calendar = $c.Range("Z20:${last}25").Value2; $flags = $c.Range("Z44:${last}56").Value2
    Assert-XF1Equal $c.Range('D15').Value2 $h 'Actual months'
    Assert-XF1Equal $c.Range('D13').Value2 $Config.start.AddMonths($n).AddDays(-1).ToOADate() 'End date'
    Assert-XF1Equal $c.Range('D11').Value2 $Config.cutoff.AddDays(1).ToOADate() 'Forecast start'
    Assert-XF1Equal $c.Range('D73').Value2 $Config.currency 'Currency'
    for ($k=1;$k -le $n;$k++) {
        $s = $Config.start.AddMonths($k-1); $e = $s.AddMonths(1).AddDays(-1)
        $fy = $Config.start.AddMonths(11).Year + [int][math]::Floor(($k-1)/12)
        $fiscalQuarter = [int][math]::Floor((($k-1)%12)/3)+1
        $actual = [int]($e -le $Config.cutoff)
        $expectedCalendar = @($s.ToOADate(),$e.ToOADate(),$(if ($actual) {$Config.actualLabel} else {$Config.forecastLabel}),$k,$fy,"Q$fiscalQuarter $fy")
        for ($r=1;$r -le 6;$r++) { Assert-XF1Equal $calendar[$r,$k] $expectedCalendar[$r-1] "Calendar $k/$r" }
        $yearCounter = [int][math]::Floor(($k-1)/12)+1
        $expectedFlags = @($s.Month,$e.Day,$actual,(1-$actual),[int]($k -eq $h+1),[math]::Max(0,$k-$h),[int]($h -gt 0 -and $k -eq $h),$yearCounter,[int]($e.Month -eq $Config.fyEnd),([int][math]::Floor(($k-1)/3)+1),$(if ($actual) {0} else {$yearCounter-[int][math]::Floor($h/12)}),[int]((($k-1)%3)-eq 0),[int](($k%3)-eq 0))
        for ($r=1;$r -le 13;$r++) { Assert-XF1Equal $flags[$r,$k] $expectedFlags[$r-1] "Flag $k/$r" }
    }
    foreach ($pair in @(@('Quarterly',3,28,4,59),@('Annual',12,34,5,65))) {
        $name=$pair[0];$step=$pair[1];$count=[int][math]::Ceiling($n/$step);$endCol=Get-XF1Column (25+$count)
        $header=$Book.Worksheets.Item($name).Range("Z1:${endCol}$($pair[3])").Value2
        # Multi-row header arrays work even for a single annual period.
        $source=$c.Range("Z$($pair[2]):${endCol}$($pair[2]+$pair[3]-1)").Value2
        for ($i=1;$i -le $count;$i++) {
            $first=($i-1)*$step+1; $endPosition=[math]::Min($i*$step,$n);$represented=$endPosition-$first+1
            Assert-XF1Equal $header[1,$i] $Config.start.AddMonths($first-1).ToOADate() "$name start $i"
            Assert-XF1Equal $header[2,$i] $Config.start.AddMonths($endPosition).AddDays(-1).ToOADate() "$name end $i"
            for ($r=1;$r -le $pair[3];$r++) { Assert-XF1Equal $header[$r,$i] $source[$r,$i] "$name source $r/$i" }
            if ($name -eq 'Annual') {
                $forecasts=0
                for ($j=$first;$j -le $endPosition;$j++) { if ($Config.start.AddMonths($j).AddDays(-1) -gt $Config.cutoff) {$forecasts++} }
                $label=if ($forecasts -eq 0) {$Config.actualLabel} elseif ($forecasts -eq $represented) {$Config.forecastLabel} else {"$($Config.actualLabel)$($represented-$forecasts)+$($Config.forecastLabel)$forecasts"}
                if ($represented -lt 12) {$label+=' (partial)'}
                Assert-XF1Equal $header[3,$i] $forecasts "Annual forecasts $i"
                Assert-XF1Equal $header[4,$i] $label "Annual type $i"
                Assert-XF1Equal $header[5,$i] $i "Annual position $i"
            } else {
                $quarterDate = $Config.start.AddMonths($first-1)
                $label="Q$([int][math]::Floor((($first-1)%12)/3)+1) $($Config.start.AddMonths(11).Year+[int][math]::Floor(($first-1)/12))"
                if ($represented -lt 3) {$label+=' (partial)'}
                Assert-XF1Equal $header[3,$i] $i "Quarter position $i"
                Assert-XF1Equal $header[4,$i] $label "Quarter label $i"
            }
        }
        $positions=$Book.Worksheets.Item($name).Range("Z10:${endCol}10").Value2
        $sourcePositions=$c.Range("Z$($pair[4]):${endCol}$($pair[4])").Value2
        for ($i=1;$i -le $count;$i++) {
            $v=if ($count -eq 1) {$positions} else {$positions[1,$i]}
            $sv=if ($count -eq 1) {$sourcePositions} else {$sourcePositions[1,$i]}
            Assert-XF1Equal $v ([math]::Min($i*$step,$n)) "$name ending position $i"
            Assert-XF1Equal $sv $v "$name source position $i"
        }
    }
    foreach ($name in @('Monthly','Checks')) {
        $s=$Book.Worksheets.Item($name);$header=$s.Range("Z1:${last}5").Value2
        $flagCount=if ($name -eq 'Monthly') {13} else {11}
        $linkedFlags=$s.Range("Z10:${last}$(9+$flagCount)").Value2
        for ($k=1;$k -le $n;$k++) {
            for ($r=1;$r -le 5;$r++) { Assert-XF1Equal $header[$r,$k] $calendar[$r,$k] "$name header $r/$k" }
            for ($r=1;$r -le $flagCount;$r++) { Assert-XF1Equal $linkedFlags[$r,$k] $flags[$r,$k] "$name flag $r/$k" }
        }
    }
    $states=@()
    $integer='#,##0;[Red]-#,##0;"-"'
    foreach ($name in @('Controls','Monthly','Quarterly','Annual','Input_sheet','Checks')) {
        $s=$Book.Worksheets.Item($name)
        $count=switch ($name) {'Quarterly' {[int][math]::Ceiling($n/3)} 'Annual' {[int][math]::Ceiling($n/12)} default {$n}}
        $lastCol=Get-XF1Column (25+$count);$spacer=Get-XF1Column (26+$count);$hidden=Get-XF1Column (27+$count)
        $bottom=switch ($name) {'Controls' {87} 'Monthly' {22} 'Checks' {40} 'Input_sheet' {5} default {10}}
        $s.Activate();$win=$Book.Windows.Item(1)
        Assert-XF1Equal $win.SplitRow 5 "$name pane row"
        Assert-XF1Equal $win.SplitColumn $(if ($name -eq 'Input_sheet') {5} else {25}) "$name pane column"
        Assert-XF1Equal $win.FreezePanes $true "$name frozen"
        Assert-XF1Equal $s.Columns.Item("Z:${spacer}").Hidden $false "$name visible"
        Assert-XF1Equal $s.Columns.Item("${hidden}:XFD").Hidden $true "$name hidden"
        Assert-XF1Equal $s.Range('C2').Value2 'Checks not configured' "$name status"
        Assert-XF1Equal $s.Range('C2').Formula2 '=model.check.message' "$name status formula"
        foreach ($address in @('C2','Z1')) {
            $cell=$s.Range($address)
            Assert-XF1FontName $cell.Font.Name "$name header font"
            Assert-XF1Equal $cell.Font.Size 10 "$name header size"
            Assert-XF1Equal $cell.Font.Color 16777215 "$name header text colour"
            Assert-XF1Equal $cell.Interior.Color 0 "$name header fill"
        }
        $colour=switch ($name) {'Monthly' {0xF0B000} 'Quarterly' {0x50D092} 'Annual' {0x50D092} 'Checks' {0x83A9F1} default {0x66D9FF}}
        # Tab/section colours are specified in RGB; Excel's numeric colour uses BGR.
        Assert-XF1Equal $s.Tab.Color $colour "$name tab colour"
        if ($name -ne 'Input_sheet') {Assert-XF1Equal $s.Range('A7').Interior.Color $colour "$name section colour"}
        $errors=$null
        try {$errors=$s.Range("A1:${lastCol}${bottom}").SpecialCells(-4123,16)} catch [Runtime.InteropServices.COMException] {if ($_.Exception.HResult -ne -2146827284) {throw}}
        if ($null -ne $errors) {throw "Formula errors: ${name}!$($errors.Address())"}
        foreach ($group in $Map.sheets[$name].groups) {
            $ends=$group.Split(':');foreach ($row in $ends) {Assert-XF1Equal $s.Rows.Item([int]$row).OutlineLevel 2 "$name row $row group"}
        }
        if ($name -ne 'Input_sheet') {
            Assert-XF1Equal ($s.Columns.Item("${spacer}:${spacer}").ColumnWidth -lt 1.1) $true "$name spacer width"
            Assert-XF1Equal $Excel.WorksheetFunction.CountA($s.Range("${spacer}1:${spacer}${bottom}")) 0 "$name empty spacer"
            Assert-XF1Equal $s.Range("${spacer}1").Interior.ColorIndex -4142 "$name spacer fill"
        }
        if ($name -in @('Monthly','Checks')) {
            Assert-XF1Equal $s.Range("Z10:${lastCol}20").NumberFormat $integer "$name numeric format"
            Assert-XF1Equal $s.Range("Z5:${lastCol}5").NumberFormat '0' "$name year format"
        }
        if ($name -in @('Quarterly','Annual')) {Assert-XF1Equal $s.Range("Z10:${lastCol}10").NumberFormat $integer "$name positions format"}
        $states+=@{name=$name;freezeRow=5;freezeColumn=$win.SplitColumn;hiddenFrom=$hidden;statusFormula='=model.check.message'}
    }
    foreach ($pair in @(@('Controls','X55:X56'),@('Monthly','X21:X22'))) {
        $s=$Book.Worksheets.Item($pair[0]);$units=$s.Range($pair[1]).Value2
        Assert-XF1Equal $units[1,1] '1/0' 'Quarter start unit';Assert-XF1Equal $units[2,1] '1/0' 'Quarter end unit'
        Assert-XF1Equal $s.Range($pair[1]).NumberFormat '@' 'Text unit format'
    }
    Assert-XF1Equal $c.Range("Z44:${last}56").NumberFormat $integer 'Controls flag format'
    foreach ($item in @(@('Controls',31,3),@('Controls',37,12),@('Quarterly',4,3),@('Annual',4,12))) {
        if (($n % $item[2]) -ne 0) {
            $s=$Book.Worksheets.Item($item[0]);$row=$item[1]
            Assert-XF1Equal $s.Range("Z${row}").WrapText $true 'Partial period label wrap'
            Assert-XF1Equal ($s.Rows.Item($row).RowHeight -ge 18) $true 'Partial period label height'
        }
    }
    foreach ($pair in @(@(28,31,3),@(34,38,12),@(59,59,3),@(65,65,12))) {
        $after=Get-XF1Column (26+[int][math]::Ceiling($n/$pair[2]))
        Assert-XF1Equal $Excel.WorksheetFunction.CountA($c.Range("${after}$($pair[0]):${last}$($pair[1])")) 0 'Controls unused reporting columns'
    }
    Assert-XF1Equal $c.Rows.Item(14).RowHeight 18 'Controls row 14';Assert-XF1Equal $c.Rows.Item(15).RowHeight 18 'Controls row 15'
    foreach ($address in @('D9','D10','D12','D14','D16:D17','D73','C76:D87')) {
        Assert-XF1Equal $c.Range($address).Font.Color 0xFF5019 "$address input font"
        Assert-XF1Equal $c.Range($address).Interior.Color 0x66FFFF "$address input fill (#FFFF66)"
    }
    Assert-XF1Equal $Book.Worksheets.Item('Checks').Range('C38').Interior.Color 0x66FFFF 'Checks message input fill (#FFFF66)'
    Assert-XF1Equal $Book.Worksheets.Item('Checks').Range('C38').Font.Color 0xFF5019 'Checks message input font'
    foreach ($a in @('D9','D10','D12','D14','D16','D17')) {Assert-XF1Equal $c.Range($a).Validation.ShowError $true "$a validation"}
    Assert-XF1Equal $Excel.WorksheetFunction.CountA($Book.Worksheets.Item('Input_sheet').Range("A6:$(Get-XF1Column (26+$n))87")) 0 'Blank input body'
    Assert-XF1Equal $Excel.WorksheetFunction.CountA($Book.Worksheets.Item('Checks').Range("Z23:${last}40")) 0 'No configured checks'
    Assert-XF1Equal $Book.Worksheets.Item('Checks').Range('E35').Value2 $null 'Blank failed check count'
    return @{date=(Get-Date -Format 'yyyy-MM-dd');sheets=$states;actualMonths=$h;forecastMonths=$n-$h;quarterlyPeriods=[int][math]::Ceiling($n/3);annualPeriods=[int][math]::Ceiling($n/12);formulaErrors=0;checksConfigured=0;calculatedBy='Microsoft Excel';verified='All calendar/flag values, linked headers, partial periods, numeric/unit formats, names, validation presence, outlines, panes, visibility and blank Inputs/Checks areas in the six template sheets. The six sheets were created from an empty workbook.'}
}

function New-XF1Map($Map,$Config,[string]$Output,[string]$Version,$Verification) {
    $n=$Config.months;$m=Get-XF1Column (25+$n);$q=Get-XF1Column (25+[int][math]::Ceiling($n/3));$y=Get-XF1Column (25+[int][math]::Ceiling($n/12))
    $ms=Get-XF1Column (26+$n);$qs=Get-XF1Column (26+[int][math]::Ceiling($n/3));$ys=Get-XF1Column (26+[int][math]::Ceiling($n/12))
    $Map.workbook=[IO.Path]::GetFileName($Output);$Map.releaseVersion=$Version
    $Map.configuration=@{currency=$Config.currency;startDate=$Config.start.ToString('yyyy-MM-dd');lastActualDate=$Config.cutoff.ToString('yyyy-MM-dd');months=$n;financialYearEndMonth=$Config.fyEnd;actualLabel=$Config.actualLabel;forecastLabel=$Config.forecastLabel}
    $Map.verification=$Verification
    $Map.timeline=@{monthly="Z:$m";monthlySpacer=$ms;quarterly="Z:$q";quarterlySpacer=$qs;annual="Z:$y";annualSpacer=$ys}
    $c=$Map.sheets.Controls
    $c.monthlyCalendar="Z20:${m}25";$c.quarterlyCalendar="Z28:${q}31";$c.annualCalendar="Z34:${y}38"
    $c.monthlyFlags="Z44:${m}56";$c.quarterlyFlags="Z59:${q}59";$c.annualFlags="Z65:${y}65"
    $c.quarterBoundaryFlags.start="Z55:${m}55";$c.quarterBoundaryFlags.end="Z56:${m}56"
    $Map.sheets.Monthly.header="Z1:${m}5";$Map.sheets.Monthly.headerSource="Controls!Z20:${m}24"
    $Map.sheets.Monthly.flags="Z10:${m}22";$Map.sheets.Monthly.flagsSource="Controls!Z44:${m}56"
    $Map.sheets.Quarterly.header="Z1:${q}4";$Map.sheets.Quarterly.headerSource="Controls!Z28:${q}31"
    $Map.sheets.Quarterly.flags="Z10:${q}10";$Map.sheets.Quarterly.flagsSource="Controls!Z59:${q}59"
    $Map.sheets.Annual.header="Z1:${y}5";$Map.sheets.Annual.headerSource="Controls!Z34:${y}38"
    $Map.sheets.Annual.flags="Z10:${y}10";$Map.sheets.Annual.flagsSource="Controls!Z65:${y}65"
    $Map.sheets.Annual.rightSpacer=$ys;$Map.sheets.Annual.hiddenFrom=Get-XF1Column (27+[int][math]::Ceiling($n/12))
    $Map.sheets.Input_sheet.header="A1:${ms}5";$Map.sheets.Input_sheet.widthException="F:${ms} = 11.56"
    # Portable references point to the pinned package, rather than nonexistent docs beside an output.
    $Map.specifications=@{releaseVersion=$Version;location='docs/ within the matching release package';formatting='shared-guidance.md'}
    return $Map
}
