<#
.SYNOPSIS
Creates native Excel evidence for multiple pivot measures and Values-axis layouts.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus'
[void](New-Item -ItemType Directory -Path $directory -Force)
$profiles = @(
    @{ Name='BothCol'; Row='Region'; Column='Product'; ValuesOnRows=$false; ValuesFirst=$false },
    @{ Name='BothRow'; Row='Region'; Column='Product'; ValuesOnRows=$true; ValuesFirst=$false },
    @{ Name='RowCol'; Row='Region'; Column=$null; ValuesOnRows=$false; ValuesFirst=$false },
    @{ Name='RowRow'; Row='Region'; Column=$null; ValuesOnRows=$true; ValuesFirst=$false },
    @{ Name='ColCol'; Row=$null; Column='Product'; ValuesOnRows=$false; ValuesFirst=$false },
    @{ Name='ColRow'; Row=$null; Column='Product'; ValuesOnRows=$true; ValuesFirst=$false },
    @{ Name='ScalarCol'; Row=$null; Column=$null; ValuesOnRows=$false; ValuesFirst=$false },
    @{ Name='ScalarRow'; Row=$null; Column=$null; ValuesOnRows=$true; ValuesFirst=$false },
    @{ Name='BothColOuter'; Row='Region'; Column='Product'; ValuesOnRows=$false; ValuesFirst=$true },
    @{ Name='BothRowOuter'; Row='Region'; Column='Product'; ValuesOnRows=$true; ValuesFirst=$true },
    @{ Name='RowRowOuter'; Row='Region'; Column=$null; ValuesOnRows=$true; ValuesFirst=$true },
    @{ Name='ColColOuter'; Row=$null; Column='Product'; ValuesOnRows=$false; ValuesFirst=$true }
)
$measures = @(@{ Field='Amount'; Caption='Revenue'; Function=-4157 }, @{ Field='Units'; Caption='AvgUnits'; Function=-4106 }, @{ Field='Amount'; Caption='Entries'; Function=-4112 })
$rows = @(@('Region','Product','Amount','Units'), @('North','A',10.0,2.0), @('North','B',20.0,3.0),
    @('South','A',5.0,1.0), @('South','B',15.0,4.0), @('North','A',-2.0,1.0), @('West','B',$null,5.0),
    @('South','A','pending',2.0), @('North','B',$true,1.0))
function Release-Com($value) { if ($null -ne $value) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } }
$application=$null; $workbooks=$null; $workbook=$null; $sheets=$null; $source=$null; $lookups=$null; $caches=$null
$owned=$false; $closed=$false; $lookupRow=1; $records=@()
try {
    $application=New-Object -ComObject Excel.Application
    $workbooks=$application.Workbooks
    if ($workbooks.Count -ne 0) { throw 'The Excel instance contains existing workbooks.' }
    $owned=$true
    $application.Visible=$false; $application.DisplayAlerts=$false; $application.AutomationSecurity=3
    $workbook=$workbooks.Add(); $sheets=$workbook.Worksheets; $source=$sheets.Item(1); $source.Name='Source'
    $values=New-Object 'object[,]' $rows.Count,4
    for ($row=0; $row -lt $rows.Count; $row++) { for ($column=0; $column -lt 4; $column++) { $values[$row,$column]=$rows[$row][$column] } }
    $range=$source.Range('A1:D9')
    try { $range.Value2=$values } finally { Release-Com $range }
    $lookups=$sheets.Add(); $lookups.Name='Lookups'; $caches=$workbook.PivotCaches()
    foreach ($profile in $profiles) {
        $sheet=$null; $cache=$null; $pivot=$null; $rowField=$null; $columnField=$null; $valuesField=$null
        try {
            $sheet=$sheets.Add(); $sheet.Name=$profile.Name
            $cache=$caches.Create(1,"'Source'!R1C1:R9C4",6)
            $destination=$sheet.Range('A1')
            try { $pivot=$cache.CreatePivotTable($destination,('Pivot'+$profile.Name)) } finally { Release-Com $destination }
            if ($profile.Row) { $rowField=$pivot.PivotFields($profile.Row); $rowField.Orientation=1; $rowField.Position=1 }
            if ($profile.Column) { $columnField=$pivot.PivotFields($profile.Column); $columnField.Orientation=2; $columnField.Position=1 }
            foreach ($measure in $measures) {
                $sourceField=$null; $dataField=$null
                try { $sourceField=$pivot.PivotFields($measure.Field); $dataField=$pivot.AddDataField($sourceField,$measure.Caption,$measure.Function) }
                finally { Release-Com $dataField; Release-Com $sourceField }
            }
            $valuesField=$pivot.DataPivotField
            $valuesField.Orientation=if ($profile.ValuesOnRows) {1} else {2}
            $valuesField.Position=if ($profile.ValuesFirst -or ($profile.ValuesOnRows -and !$profile.Row) -or (!$profile.ValuesOnRows -and !$profile.Column)) {1} else {2}
            $pivot.RowAxisLayout(1); $pivot.RowGrand=$true; $pivot.ColumnGrand=$true
            [void]$pivot.RefreshTable()
            $firstLookup=$lookupRow
            foreach ($measure in $measures) {
                foreach ($keys in @(@(), @('Region','North'), @('Region','South'), @('Region','West'), @('Product','A'), @('Product','B'),
                    @('Region','North','Product','A'), @('Region','South','Product','B'), @('Region','missing'))) {
                    $formula='=GETPIVOTDATA("'+$measure.Caption+'",'+$profile.Name+'!A1'
                    foreach ($key in $keys) { $formula+=',"'+$key+'"' }
                    $formula+=')'
                    $cell=$lookups.Cells.Item($lookupRow,1)
                    try {$cell.Value2=$profile.Name} finally {Release-Com $cell}
                    $cell=$lookups.Cells.Item($lookupRow,2)
                    try {$cell.Formula=$formula} finally {Release-Com $cell}
                    $lookupRow++
                }
            }
            $range=$pivot.TableRange1
            try { $outputRange=$range.Address($false,$false) } finally { Release-Com $range }
            $records+=[ordered]@{ name='Pivot'+$profile.Name; sheet=$profile.Name; row=$profile.Row; column=$profile.Column;
                valuesOnRows=$profile.ValuesOnRows; valuesFirst=$profile.ValuesFirst; outputRange=$outputRange; firstLookup=$firstLookup; lookupCount=27 }
        } finally { foreach ($value in @($valuesField,$columnField,$rowField,$pivot,$cache,$sheet)) { Release-Com $value } }
    }
    $application.CalculateFullRebuild()
    $path=Join-Path $directory 'layout-conformance.xlsx'
    $workbook.SaveAs($path,51); $workbook.Close($false); $closed=$true
    [ordered]@{ producer='Microsoft Excel'; version=$application.Version; build=$application.Build; generatedUtc=[DateTime]::UtcNow.ToString('o');
        regeneration='Build/Verification/New-ExcelPivotLayoutOracle.ps1'; file='layout-conformance.xlsx'; sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();
        sourceRange='Source!A1:D9'; lookupFormulaCells=$lookupRow-1; measures=$measures; profiles=$records } |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'layout-conformance.provenance.json') -Encoding utf8
} finally {
    if ($null -ne $workbook -and !$closed) { $workbook.Close($false) }
    if ($owned -and $null -ne $application) { $application.Quit() }
    foreach ($value in @($caches,$lookups,$source,$sheets,$workbook,$workbooks,$application)) { Release-Com $value }
}
