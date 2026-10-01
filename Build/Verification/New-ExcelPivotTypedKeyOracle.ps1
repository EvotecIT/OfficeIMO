<#
.SYNOPSIS
Creates independent Excel evidence for ungrouped date, error and mixed pivot keys.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus'
function Release-Com($value) { if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } }
$app=$null; $books=$null; $owned=$false
try {
    $app=New-Object -ComObject Excel.Application; $books=$app.Workbooks
    if ($books.Count -ne 0) { throw 'The Excel instance contains existing workbooks.' }
    $owned=$true; $app.Visible=$false; $app.DisplayAlerts=$false; $app.AutomationSecurity=3
    foreach ($system in @(1900,1904)) {
        $book=$null; $sheets=$null; $source=$null; $lookups=$null; $caches=$null; $closed=$false
        try {
            $book=$books.Add(); $book.Date1904=($system -eq 1904); $sheets=$book.Worksheets
            $source=$sheets.Item(1); $source.Name='Source'; $lookups=$sheets.Add(); $lookups.Name='Lookups'; $caches=$book.PivotCaches()
            $dates=if ($system -eq 1900) { @(1,59,60,60.5,61,45292,45292.5) } else { @(0,1,59,60,60.5,43830,43830.5) }
            $primaryDate=$dates[-2]
            $rows=,@('Group','DateKey','MixedKey','Amount')
            for ($index=0; $index -lt $dates.Count; $index++) { $rows+=,@('North',$dates[$index],$dates[$index],(10*($index+1))) }
            $rows+=,@('South',$primaryDate,42,60)
            $rows+=,@('South',$primaryDate,$primaryDate,120)
            $rows+=,@('South',$primaryDate,'text',70)
            $rows+=,@('South',$primaryDate,$true,80)
            $rows+=,@('South',$null,$null,90)
            $rows+=,@('South',$primaryDate,'error',100)
            $rows+=,@('West',$primaryDate,'#DIV/0!',110)
            $values=New-Object 'object[,]' $rows.Count,5
            for ($row=0; $row -lt $rows.Count; $row++) { for ($column=0; $column -lt 4; $column++) { $values[$row,$column]=$rows[$row][$column] } }
            $values[0,4]='Child'
            for ($row=1; $row -lt $rows.Count; $row++) { $values[$row,4]='Child'+($row%2) }
            $lastRow=$rows.Count
            $range=$source.Range("A1:E$lastRow"); try { $range.Value2=$values } finally { Release-Com $range }
            $range=$source.Range("B2:B$lastRow"); try { $range.NumberFormat='yyyy-mm-dd hh:mm:ss' } finally { Release-Com $range }
            $range=$source.Range('C2:C'+($dates.Count+1)); try { $range.NumberFormat='yyyy-mm-dd hh:mm:ss' } finally { Release-Com $range }
            $range=$source.Range('C'+($lastRow-1)); try { $range.Formula='=1/0' } finally { Release-Com $range }
            $range=$source.Range("C$lastRow"); try { $range.Value2="'#DIV/0!" } finally { Release-Com $range }
            $profiles=@(
                @{Name='DatesRow';Rows=@('DateKey');Columns=@()},
                @{Name='DatesCol';Rows=@();Columns=@('DateKey')},
                @{Name='DatesNested';Rows=@('Group','DateKey');Columns=@()},
                @{Name='DatesOuterRows';Rows=@('DateKey','Group');Columns=@()},
                @{Name='DatesOuterCols';Rows=@();Columns=@('DateKey','Group')},
                @{Name='DatesDeepRows';Rows=@('DateKey','Group','Child');Columns=@()},
                @{Name='DatesDeepCols';Rows=@();Columns=@('DateKey','Group','Child')},
                @{Name='MixedRow';Rows=@('MixedKey');Columns=@()},
                @{Name='MixedCol';Rows=@();Columns=@('MixedKey')},
                @{Name='MixedNested';Rows=@('Group','MixedKey');Columns=@()}
            )
            $records=@(); $lookupRow=1
            foreach ($profile in $profiles) {
                $sheet=$null; $cache=$null; $pivot=$null
                try {
                    $sheet=$sheets.Add(); $sheet.Name=$profile.Name; $cache=$caches.Create(1,"'Source'!R1C1:R${lastRow}C5",6)
                    $range=$sheet.Range('A1'); try { $pivot=$cache.CreatePivotTable($range,('Pivot'+$profile.Name)) } finally { Release-Com $range }
                    foreach ($axis in @(@{Fields=$profile.Rows;Orientation=1},@{Fields=$profile.Columns;Orientation=2})) {
                        $position=1
                        foreach ($name in $axis.Fields) {
                            $field=$null
                            try {
                                $field=$pivot.PivotFields($name); $field.Orientation=$axis.Orientation; $field.Position=$position
                                if ($name -eq 'DateKey') { try { $field.Ungroup() } catch { } }
                            } finally { Release-Com $field }
                            $position++
                        }
                    }
                    $field=$null; $data=$null
                    try { $field=$pivot.PivotFields('Amount'); $data=$pivot.AddDataField($field,'Revenue',-4157) } finally {Release-Com $data;Release-Com $field}
                    $pivot.RowAxisLayout(1); [void]$pivot.RefreshTable()
                    $fieldName=if ($profile.Name.StartsWith('Dates')) {'DateKey'} else {'MixedKey'}
                    $selections=@(@{Kind='total';Pairs=@()})
                    foreach ($date in $dates) { $selections+=@{Kind='serial';Pairs=@($fieldName,$date)} }
                    $selections+=@{Kind='blank';Pairs=@($fieldName,'(blank)')}
                    if ($fieldName -eq 'MixedKey') {
                        $selections+=@{Kind='number';Pairs=@($fieldName,42)},@{Kind='text';Pairs=@($fieldName,'text')},@{Kind='boolean';Pairs=@($fieldName,$true)},@{Kind='errorLabel';Pairs=@($fieldName,'#DIV/0!')}
                    }
                    if ($profile.Rows.Count + $profile.Columns.Count -ge 2) {
                        $selections+=@{Kind='subtotal';Pairs=@('Group','South')},@{Kind='nested';Pairs=@('Group','South',$fieldName,$primaryDate)}
                        if ($fieldName -eq 'MixedKey') { $selections+=@{Kind='errorOnly';Pairs=@('Group','South',$fieldName,'#DIV/0!')},@{Kind='errorText';Pairs=@('Group','West',$fieldName,'#DIV/0!')} }
                    }
                    $first=$lookupRow
                    foreach ($selection in $selections) {
                        $formula='=GETPIVOTDATA("Revenue",'+$profile.Name+'!A1'
                        foreach ($value in $selection.Pairs) {
                            $literal=if ($value -is [string]) { '"'+$value.Replace('"','""')+'"' } elseif ($value -is [bool]) { if ($value) {'TRUE'} else {'FALSE'} } else { $value.ToString([Globalization.CultureInfo]::InvariantCulture) }
                            $formula+=','+$literal
                        }
                        $formula+=')'; $range=$lookups.Cells.Item($lookupRow,1); try {$range.Value2=$profile.Name} finally {Release-Com $range}
                        $range=$lookups.Cells.Item($lookupRow,2); try {$range.Formula=$formula} finally {Release-Com $range}; $lookupRow++
                    }
                    $range=$pivot.TableRange1; try {$address=$range.Address($false,$false)} finally {Release-Com $range}
                    $errorLookup=$null
                    if ($fieldName -eq 'MixedKey') {
                        $range=$pivot.GetPivotData('Revenue','MixedKey',[Runtime.InteropServices.ErrorWrapper]::new(-2146826281))
                        try {$errorLookup=$range.Value2} finally {Release-Com $range}
                    }
                    $records+=[ordered]@{name='Pivot'+$profile.Name;sheet=$profile.Name;rows=$profile.Rows;columns=$profile.Columns;outputRange=$address;firstLookup=$first;lookupCount=$selections.Count;selections=$selections;errorLookupValue=$errorLookup}
                } finally {Release-Com $pivot;Release-Com $cache;Release-Com $sheet}
            }
            $app.CalculateFullRebuild(); $file="typed-keys-$system.xlsx"; $path=Join-Path $directory $file
            $book.SaveAs($path,51); $book.Close($false); $closed=$true
            [ordered]@{producer='Microsoft Excel';version=$app.Version;build=$app.Build;generatedUtc=[DateTime]::UtcNow.ToString('o');dateSystem=$system;
                regeneration='Build/Verification/New-ExcelPivotTypedKeyOracle.ps1';file=$file;sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();sourceRange="Source!A1:E$lastRow";lookupFormulaCells=$lookupRow-1;profiles=$records} |
                ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory "typed-keys-$system.provenance.json") -Encoding utf8
        } finally {
            if ($null -ne $book -and !$closed) {$book.Close($false)}
            foreach ($value in @($caches,$lookups,$source,$sheets,$book)) {Release-Com $value}
        }
    }
} finally {
    if ($owned) {$app.Quit()}; Release-Com $books; Release-Com $app
}
