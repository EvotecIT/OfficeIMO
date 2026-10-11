<#
.SYNOPSIS
Creates an editable milestone register and a dated delivery review from one CSV.
.DESCRIPTION
Status follows progress and due date: completed milestones are Done; incomplete milestones
before the snapshot date are Overdue. Workbook status formulas recalculate after edits.
Word and PDF are dated snapshots; regenerate them after changing the input.
.EXAMPLE
./ProjectDeliveryPack.ps1 -OutputDirectory ./out -ModuleVersion 3.0.8
.EXAMPLE
./ProjectDeliveryPack.ps1 -InputPath ./Milestones.csv -AsOfDate 2026-10-02 -OutputDirectory ./out
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $OutputDirectory,
    [string] $InputPath = '',
    [datetime] $AsOfDate = '2026-09-30',
    [string] $ModuleVersion = '3.0.8'
)

$ErrorActionPreference = 'Stop'
if (-not $InputPath) { $InputPath = Join-Path $PSScriptRoot 'Data/ProjectMilestones.csv' }
Import-Module PSWriteOffice -RequiredVersion $ModuleVersion -ErrorAction Stop
$culture = [Globalization.CultureInfo]::InvariantCulture
$snapshotDate = $AsOfDate.Date
$inputRows = @(Import-Csv -LiteralPath $InputPath -Encoding UTF8)
if (-not $inputRows.Count) { throw 'Provide at least one milestone in the project CSV.' }
$rows = @(foreach ($item in $inputRows) {
    foreach ($field in 'Milestone','Owner','DueDate','Progress','NextAction') {
        if ($field -notin $item.PSObject.Properties.Name -or [string]::IsNullOrWhiteSpace($item.$field)) {
            throw "Every milestone needs a value for $field."
        }
    }
    $due = [datetime]::ParseExact($item.DueDate, 'yyyy-MM-dd', $culture)
    $progress = [int]::Parse($item.Progress, $culture)
    if ($progress -lt 0 -or $progress -gt 100) { throw "Progress must be an integer from 0 to 100 for $($item.Milestone)." }
    $status = if ($progress -eq 100) { 'Done' } elseif ($due -lt $snapshotDate) { 'Overdue' } else { 'On track' }
    [pscustomobject]@{ Milestone=$item.Milestone; Owner=$item.Owner; DueDate=$due;
        Progress=$progress; Status=$status; NextAction=$item.NextAction }
})
$overdue = @($rows | Where-Object Status -eq 'Overdue' | Sort-Object DueDate)
$done = @($rows | Where-Object Status -eq 'Done')
$lastRow = $rows.Count + 1
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$xlsx = Join-Path $OutputDirectory 'ProjectDeliveryPack.xlsx'
$docx = Join-Path $OutputDirectory 'ProjectDeliveryPack.docx'
$pdf = Join-Path $OutputDirectory 'ProjectDeliveryPack.pdf'

#region Excerpt:pswriteoffice-project-delivery
ExcelNew -Path $xlsx -EvaluateFormulas {
    ExcelSheet 'Review' {
        ExcelGridlines -Hide
        ExcelColumn -ColumnName A -Width 44
        ExcelColumn -ColumnName B -Width 24
        ExcelCell -Address A1 -Value 'Service rollout | delivery review'
        ExcelRow -Row 1 -Bold $true -Height 32
        ExcelCell -Address A2 -Value 'Snapshot date'
        ExcelCell -Address B2 -Value $snapshotDate.ToOADate() -NumberFormat 'yyyy-mm-dd'
        ExcelCell -Address A3 -Value 'Completed milestones'
        ExcelCell -Address B3 -Formula "COUNTIF(Milestones!E2:E$lastRow,`"Done`")"
        ExcelCell -Address A4 -Value 'Overdue milestones'
        ExcelCell -Address B4 -Formula "COUNTIF(Milestones!E2:E$lastRow,`"Overdue`")"
        ExcelCell -Address A6 -Value 'Update the snapshot date, due dates and progress to review another scenario.'
        ExcelRow -Row 6 -WrapText $true -Height 42
        ExcelCell -Address A7 -Value 'Regenerate the Word/PDF review from the revised CSV for distribution.'
        ExcelRow -Row 7 -WrapText $true -Height 34
        ExcelPageSetup -FitToWidth 1 -FitToHeight 0
        ExcelPrintArea -Range 'A1:B7'
    }
    ExcelSheet 'Milestones' {
        ExcelTable -Data $rows -TableName 'DeliveryMilestones' -TableStyle TableStyleMedium4
        ExcelFreeze -TopRows 1
        ExcelColumn -ColumnName A -Width 36
        ExcelColumn -ColumnName B -Width 24
        ExcelColumn -ColumnName C -Width 18
        ExcelColumn -ColumnName D -Width 15
        ExcelColumn -ColumnName E -Width 18
        ExcelColumn -ColumnName F -Width 62
        for ($row = 2; $row -le $lastRow; $row++) {
            ExcelCell -Address "C$row" -Value $rows[$row-2].DueDate.ToOADate() -NumberFormat 'yyyy-mm-dd'
            ExcelCell -Address "D$row" -Value ($rows[$row-2].Progress / 100.0) -NumberFormat '0%'
            ExcelCell -Address "E$row" -Formula "IF(D$row=1,`"Done`",IF(C$row<Review!`$B`$2,`"Overdue`",`"On track`"))"
            ExcelRow -Row $row -WrapText $true
        }
        ExcelValidationDecimal -Range "D2:D$lastRow" -Operator Between -Formula1 0 -Formula2 1 -AllowBlank:$false -ErrorTitle 'Progress from 0% to 100%' -ErrorMessage 'Enter a percentage from 0% to 100%.'
        ExcelOrientation -Orientation Landscape
        ExcelPageSetup -FitToWidth 1 -FitToHeight 0
        ExcelPrintTitles -FirstRow 1 -LastRow 1
        ExcelPrintArea -Range "A1:F$lastRow"
    }
}

$reviewRows = $rows | Select-Object Milestone,Owner,
    @{Name='Due';Expression={$_.DueDate.ToString('yyyy-MM-dd')}},
    @{Name='Progress';Expression={"$($_.Progress)%"}},Status
WordNew -Path $docx {
    WordHeader { WordParagraph -Text "DELIVERY | Service rollout | $($snapshotDate.ToString('yyyy-MM-dd'))" }
    WordFooter { WordPageNumber -IncludeTotalPages }
    WordParagraph { WordParagraphStyle -SpacingAfterPoints 12 -KeepWithNext $true; WordText 'Project delivery review' -Bold -FontSize 25 -Color '17365D' }
    WordParagraph -Text "$($done.Count) of $($rows.Count) milestones are complete; $($overdue.Count) overdue milestone(s) need an owner decision at $($snapshotDate.ToString('yyyy-MM-dd'))."
    WordParagraph -Text 'Decisions and recovery actions' -Style Heading1 { WordParagraphStyle -KeepWithNext $true -SpacingBeforePoints 12 -SpacingAfterPoints 6 }
    WordList -Style Bulleted {
        foreach ($item in $overdue) {
            $decision = WordListItem -Text "$($item.Owner) | $($item.Milestone): $($item.NextAction)." -PassThru
            WordParagraphStyle -Paragraph $decision -KeepLinesTogether $true
        }
        if (-not $overdue.Count) { WordListItem -Text 'No overdue milestones. Confirm the next due dates and acceptance evidence with their owners.' }
    }
    WordParagraph -Text 'Review owner: __________________  Decision: __________________' { WordParagraphStyle -SpacingBeforePoints 12 }
    WordParagraph -Text 'Milestone position' -Style Heading1 { WordParagraphStyle -KeepWithNext $true -SpacingBeforePoints 12 -SpacingAfterPoints 6 }
    $table = WordTable -InputObject $reviewRows -Style GridTable4Accent1 -Layout AutoFitToWindow -PassThru {
        WordTableCondition -FilterScript { $_.Status -eq 'Overdue' } -BackgroundColor '#FFF4D6'
    }
    $table.RepeatHeaderRowAtTheTopOfEachPage = $true
    foreach ($tableRow in $table.Rows) { $tableRow.AllowRowToBreakAcrossPages = $false }
    $table | Get-OfficeWordTableCell | ForEach-Object {
        $_.MarginTopCentimeters=0.12; $_.MarginBottomCentimeters=0.12
        $_.MarginLeftCentimeters=0.16; $_.MarginRightCentimeters=0.16
    }
    WordParagraph -Text 'Synthetic sample. Due today is on track; overdue means an incomplete milestone due before the snapshot date. Update the CSV and regenerate this review after changes. Record the review decision in the editable Word copy.'
}
Export-OfficeDocumentPdf -InputPath $docx -Path $pdf
#endregion
Get-Item -LiteralPath $xlsx,$docx,$pdf | Select-Object Name,Length
