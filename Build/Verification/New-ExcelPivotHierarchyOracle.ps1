<#
.SYNOPSIS
Creates native Excel evidence for hierarchical pivot axes and partial-key lookups.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus'
[void](New-Item -ItemType Directory -Path $directory -Force)
$profiles = @(
    @{ Name='Row2'; Rows=@('Region','City'); Columns=@(); Measures=1 },
    @{ Name='Col2'; Rows=@(); Columns=@('Product','Channel'); Measures=1 },
    @{ Name='Both2'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=1 },
    @{ Name='Row3'; Rows=@('Region','City','Product'); Columns=@(); Measures=1 },
    @{ Name='Col3'; Rows=@(); Columns=@('Product','Channel','Region'); Measures=1 },
    @{ Name='ManyCol'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=3; ValuesOnRows=$false; ValuesPosition=3 },
    @{ Name='ManyRow'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=3; ValuesOnRows=$true; ValuesPosition=3 },
    @{ Name='MiddleCol'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=3; ValuesOnRows=$false; ValuesPosition=2 },
    @{ Name='MiddleRow'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=3; ValuesOnRows=$true; ValuesPosition=2 },
    @{ Name='OuterCol'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=3; ValuesOnRows=$false; ValuesPosition=1 },
    @{ Name='OuterRow'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=3; ValuesOnRows=$true; ValuesPosition=1 },
    @{ Name='NoSubtotals'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=1; NoSubtotals=$true },
    @{ Name='CustomSum'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=1; SubtotalFunctions=@(2) },
    @{ Name='CustomTwo'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=1; SubtotalFunctions=@(2,4) },
    @{ Name='NoGrand'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=1; NoGrand=$true },
    @{ Name='NoSubtotal3'; Rows=@('Region','City','Product'); Columns=@('Channel'); Measures=3; ValuesOnRows=$true; ValuesPosition=2; NoSubtotals=$true },
    @{ Name='CompactManyRow'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=3; ValuesOnRows=$true; ValuesPosition=2; Layout=0 },
    @{ Name='OutlineManyCol'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=3; ValuesOnRows=$false; ValuesPosition=1; Layout=2 },
    @{ Name='TopSubtotalRow3'; Rows=@('Region','City','Product'); Columns=@(); Measures=1; Layout=2; SubtotalsAtTop=$true },
    @{ Name='CollapsedRow'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=1; CollapseNorth=$true },
    @{ Name='CompactOuterRow'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=3; ValuesOnRows=$true; ValuesPosition=1; Layout=0 },
    @{ Name='CollapsedManyRow'; Rows=@('Region','City'); Columns=@('Product','Channel'); Measures=3; ValuesOnRows=$true; ValuesPosition=3; CollapseNorth=$true }
)
$measures = @(@{ Field='Amount'; Caption='Revenue'; Function=-4157 }, @{ Field='Units'; Caption='AvgUnits'; Function=-4106 }, @{ Field='Amount'; Caption='Entries'; Function=-4112 })
$rows = @(@('Region','City','Product','Channel','Amount','Units'),
    @('North','East','A','Retail',10.0,2.0), @('South','East','B','Web',20.0,3.0),
    @('North','West','B','Retail',5.0,1.0), @('South','West','A','Web',15.0,4.0),
    @('North','East','B','Web',-2.0,1.0), @('West','Unique','C','Direct',$null,5.0),
    @('South','East','A','Retail','pending',2.0), @('North','West','A','Web',$true,1.0),
    @('South','East','A','Web',8.0,2.0), @('North','East','A','Retail',3.0,6.0),
    @('West','East','A','Retail',4.0,2.0), @('South','West','B','Retail',7.0,3.0))
$selections=@()
$keys=@('Region','North','City','East','Product','A','Channel','Retail')
for ($mask=0; $mask -lt 16; $mask++) {
    $selection=@()
    for ($field=0; $field -lt 4; $field++) {if ($mask -band (1 -shl $field)) {$selection+=@($keys[$field*2],$keys[$field*2+1])}}
    $selections+=,@($selection)
}
$selections+=,@('City','Unique'); $selections+=,@('Product','C'); $selections+=,@('Channel','Direct')
$selections+=,@('Region','North','City','missing'); $selections+=,@('Region','West','City','Unique','Product','C','Channel','Direct')
function Release-Com($value) { if ($null -ne $value) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } }
$application=$null; $workbooks=$null; $workbook=$null; $sheets=$null; $source=$null; $lookups=$null; $caches=$null
$owned=$false; $closed=$false; $lookupRow=1; $records=@()
try {
    $application=New-Object -ComObject Excel.Application; $workbooks=$application.Workbooks
    if ($workbooks.Count -ne 0) { throw 'The Excel instance contains existing workbooks.' }
    $owned=$true; $application.Visible=$false; $application.DisplayAlerts=$false; $application.AutomationSecurity=3
    $workbook=$workbooks.Add(); $sheets=$workbook.Worksheets; $source=$sheets.Item(1); $source.Name='Source'
    $values=New-Object 'object[,]' $rows.Count,6
    for ($row=0; $row -lt $rows.Count; $row++) { for ($column=0; $column -lt 6; $column++) { $values[$row,$column]=$rows[$row][$column] } }
    $range=$source.Range('A1:F13'); try { $range.Value2=$values } finally { Release-Com $range }
    $lookups=$sheets.Add(); $lookups.Name='Lookups'; $caches=$workbook.PivotCaches()
    foreach ($profile in $profiles) {
        $sheet=$null; $cache=$null; $pivot=$null; $valuesField=$null
        try {
            $sheet=$sheets.Add(); $sheet.Name=$profile.Name; $cache=$caches.Create(1,"'Source'!R1C1:R13C6",6)
            $destination=$sheet.Range('A1'); try { $pivot=$cache.CreatePivotTable($destination,('Pivot'+$profile.Name)) } finally { Release-Com $destination }
            foreach ($axis in @(@{Fields=$profile.Rows; Orientation=1},@{Fields=$profile.Columns; Orientation=2})) {
                $position=1
                foreach ($name in $axis.Fields) {
                    $field=$null
                    try { $field=$pivot.PivotFields($name); $field.Orientation=$axis.Orientation; $field.Position=$position
                        if ($profile.NoSubtotals -or $profile.SubtotalFunctions) {for ($subtotal=1; $subtotal -le 12; $subtotal++) {$field.Subtotals($subtotal)=$false}}
                        foreach ($subtotal in $profile.SubtotalFunctions) {$field.Subtotals($subtotal)=$true}
                    } finally {Release-Com $field}
                    $position++
                }
            }
            for ($measureIndex=0; $measureIndex -lt $profile.Measures; $measureIndex++) {
                $measure=$measures[$measureIndex]; $sourceField=$null; $dataField=$null
                try { $sourceField=$pivot.PivotFields($measure.Field); $dataField=$pivot.AddDataField($sourceField,$measure.Caption,$measure.Function) }
                finally { Release-Com $dataField; Release-Com $sourceField }
            }
            if ($profile.Measures -gt 1) {
                $valuesField=$pivot.DataPivotField; $valuesField.Orientation=if ($profile.ValuesOnRows) {1} else {2}; $valuesField.Position=$profile.ValuesPosition
            }
            $layout=if ($profile.ContainsKey('Layout')) {$profile.Layout} else {1}
            $pivot.RowAxisLayout($layout); $pivot.RowGrand=!$profile.NoGrand; $pivot.ColumnGrand=!$profile.NoGrand
            if ($profile.SubtotalsAtTop) {
                foreach ($name in $profile.Rows) {
                    $field=$null
                    try {$field=$pivot.PivotFields($name); $field.LayoutSubtotalLocation=1} finally {Release-Com $field}
                }
            }
            if ($profile.CollapseNorth) {
                $field=$null; $item=$null
                try {$field=$pivot.PivotFields('Region'); $item=$field.PivotItems('North'); $item.ShowDetail=$false}
                finally {Release-Com $item; Release-Com $field}
            }
            [void]$pivot.RefreshTable()
            $firstLookup=$lookupRow
            for ($measureIndex=0; $measureIndex -lt $profile.Measures; $measureIndex++) {
                foreach ($selection in $selections) {
                    $formula='=GETPIVOTDATA("'+$measures[$measureIndex].Caption+'",'+$profile.Name+'!A1'
                    foreach ($key in $selection) {$formula+=',"'+$key+'"'}
                    $formula+=')'; $cell=$lookups.Cells.Item($lookupRow,1)
                    try {$cell.Value2=$profile.Name} finally {Release-Com $cell}
                    $cell=$lookups.Cells.Item($lookupRow,2); try {$cell.Formula=$formula} finally {Release-Com $cell}; $lookupRow++
                }
            }
            $range=$pivot.TableRange1; try {$outputRange=$range.Address($false,$false)} finally {Release-Com $range}
            $records+=[ordered]@{name='Pivot'+$profile.Name; sheet=$profile.Name; rows=$profile.Rows; columns=$profile.Columns; measures=$profile.Measures;
                valuesOnRows=[bool]$profile.ValuesOnRows; valuesPosition=$profile.ValuesPosition; noSubtotals=[bool]$profile.NoSubtotals;
                subtotalFunctions=$profile.SubtotalFunctions; noGrand=[bool]$profile.NoGrand; layout=$layout;
                subtotalsAtTop=[bool]$profile.SubtotalsAtTop; collapseNorth=[bool]$profile.CollapseNorth;
                outputRange=$outputRange; firstLookup=$firstLookup; lookupCount=$profile.Measures*$selections.Count}
        } finally {foreach ($value in @($valuesField,$pivot,$cache,$sheet)) {Release-Com $value}}
    }
    $application.CalculateFullRebuild(); $path=Join-Path $directory 'hierarchy-conformance.xlsx'
    $workbook.SaveAs($path,51); $workbook.Close($false); $closed=$true
    [ordered]@{producer='Microsoft Excel'; version=$application.Version; build=$application.Build; generatedUtc=[DateTime]::UtcNow.ToString('o');
        regeneration='Build/Verification/New-ExcelPivotHierarchyOracle.ps1'; file='hierarchy-conformance.xlsx'; sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();
        sourceRange='Source!A1:F13'; lookupFormulaCells=$lookupRow-1; measures=$measures; selections=$selections; profiles=$records} |
        ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $directory 'hierarchy-conformance.provenance.json') -Encoding utf8
} finally {
    if ($null -ne $workbook -and !$closed) {$workbook.Close($false)}
    if ($owned -and $null -ne $application) {$application.Quit()}
    foreach ($value in @($caches,$lookups,$source,$sheets,$workbook,$workbooks,$application)) {Release-Com $value}
}
