#Requires -Version 5.1
<#
.SYNOPSIS
    Finds and removes comments, authors and other hidden data before files are shared.
.DESCRIPTION
    Creates a Word proposal with a reviewer comment and a custom property, an Excel
    budget with a cell comment, and a PowerPoint briefing with speaker notes, all
    with an author name. It lists what would leave with each file, writes cleaned
    copies into a Share folder, inspects those copies again and checks that the
    originals did not change.
.PARAMETER OutputDirectory
    Parent directory for a new CleanBeforeSharing folder. That folder must not
    already exist so this example cannot overwrite an earlier run.
.EXAMPLE
    ./Recipe-CleanBeforeSharing.ps1 -OutputDirectory ./Output
.OUTPUTS
    One summary row per cleaned file.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $OutputDirectory
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop

$parent = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$root = Join-Path $parent 'CleanBeforeSharing'
if (Test-Path -LiteralPath $root) {
    throw "The sample folder already exists: $root. Choose another output directory."
}
$outgoing = Join-Path $root 'Outgoing'
$share = Join-Path $root 'Share'
New-Item -ItemType Directory -Path $outgoing -Force | Out-Null

$word = New-OfficeWord -Path (Join-Path $outgoing 'Proposal.docx')
$word.AddParagraph('Proposal for the client.').AddComment('Jane Reviewer', 'JR', 'Check the price before sending.') | Out-Null
Set-OfficeWordDocumentProperty -Document $word -Name Creator -Value 'Jane Reviewer'
Set-OfficeWordDocumentProperty -Document $word -Name ClientCode -Value 'CONTOSO-7731' -Custom
Save-OfficeWord -Document $word

$excel = New-OfficeExcel -Path (Join-Path $outgoing 'Budget.xlsx')
Add-OfficeExcelSheet -Document $excel -Name Budget | Out-Null
Set-OfficeExcelCell -Document $excel -Sheet Budget -Address A1 -Value 'Budget'
Add-OfficeExcelComment -Document $excel -Sheet Budget -Address A1 -Text 'Confirm with finance.' -Author 'Jane Reviewer'
Set-OfficeExcelDocumentProperty -Document $excel -Name Creator -Value 'Jane Reviewer'
Save-OfficeExcel -Document $excel

$deck = New-OfficePowerPoint -Path (Join-Path $outgoing 'Briefing.pptx') -Author 'Jane Reviewer'
$slide = Add-OfficePowerPointSlide -Document $deck -Title 'Quarterly briefing' -PassThru
Set-OfficePowerPointNotes -Slide $slide -Text 'Do not mention the merger yet.'
Save-OfficePowerPoint -Document $deck

$originals = Get-ChildItem -LiteralPath $outgoing -File | Get-FileHash -Algorithm SHA256

# Report what would leave with each file. Nothing is changed.
$found = Get-ChildItem -LiteralPath $outgoing -File | Get-OfficeHiddenData
foreach ($file in $found) {
    if (-not ($file.Findings | Where-Object Kind -In 'Comments', 'SpeakerNotes')) {
        throw "Expected a comment or speaker notes in $($file.Path)."
    }
}

# Write cleaned copies into the share folder with the default categories.
$cleaned = Get-ChildItem -LiteralPath $outgoing -File | Remove-OfficeHiddenData -OutputDirectory $share

# Check the copies again and confirm the originals are untouched.
foreach ($copy in Get-OfficeHiddenData -LiteralPath $cleaned.OutputPath) {
    if ($copy.Findings.Count -ne 0) {
        throw "Hidden data remains in $($copy.Path): $($copy.Kinds)"
    }
}
foreach ($original in $originals) {
    if ((Get-FileHash -LiteralPath $original.Path -Algorithm SHA256).Hash -ne $original.Hash) {
        throw "The original changed: $($original.Path)"
    }
}

foreach ($report in $cleaned) {
    [pscustomobject]@{
        File    = Split-Path -Leaf $report.Path
        Removed = @($report.Findings | Where-Object Outcome -EQ 'Removed').Count
        Kinds   = $report.Kinds
        Copy    = $report.OutputPath
    }
}
