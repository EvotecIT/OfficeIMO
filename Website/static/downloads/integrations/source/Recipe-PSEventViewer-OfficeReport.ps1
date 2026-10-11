<#
.SYNOPSIS
Creates a Word/PDF incident brief and Excel event register from a bounded snapshot.
.DESCRIPTION
The default is synthetic data. -CollectLive reads recent local critical events, warnings
and errors with PSEventViewer. -Event accepts an existing Get-EVXEvent snapshot.
Counts describe the retained records, not every event in the log.
.EXAMPLE
./Recipe-PSEventViewer-OfficeReport.ps1 -OutputDirectory ./Output/EventBrief
.EXAMPLE
./Recipe-PSEventViewer-OfficeReport.ps1 -CollectLive -MaxEvents 100 -OutputDirectory ./Output/LiveBrief
.EXAMPLE
$events = @(Get-EVXEvent -LogName System -Level 2,3 -TimePeriod Last24Hours -MaxEvents 100)
./Recipe-PSEventViewer-OfficeReport.ps1 -Event $events -OutputDirectory ./Output/SystemBrief
#>
[CmdletBinding()]
param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot 'Output/EventBrief'),
    [object[]] $Event,
    [switch] $CollectLive,
    [ValidateNotNullOrEmpty()][string] $LogName = 'System',
    [ValidateRange(1, 1000)][int] $MaxEvents = 100
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -MinimumVersion 3.0.8 -ErrorAction Stop
if ($CollectLive -and $PSBoundParameters.ContainsKey('Event')) { throw 'Choose -Event or -CollectLive.' }
if ($CollectLive) {
    Import-Module PSEventViewer -MinimumVersion 4.0.1 -ErrorAction Stop
    $Event = @(Get-EVXEvent -LogName $LogName -Level 1,2,3 -TimePeriod Last24Hours -MaxEvents $MaxEvents -ReadMode Message)
    $scope = "Local $LogName critical/warning/error events / last 24 hours / newest $MaxEvents at most"
} elseif ($PSBoundParameters.ContainsKey('Event')) {
    $scope = "Provided event snapshot / first $MaxEvents records at most"
} else {
    $scope = 'Synthetic incident exercise / 11 October 2026'
    $Event = @(
        [pscustomobject]@{ TimeCreated = [datetime] '2026-10-11T08:10:00'; Id = 1001; ProviderName = 'Demo.Backup'; Level = 2; Message = 'The overnight backup did not complete.' }
        [pscustomobject]@{ TimeCreated = [datetime] '2026-10-11T08:12:00'; Id = 1001; ProviderName = 'Demo.Backup'; Level = 2; Message = 'The backup retry did not complete.' }
        [pscustomobject]@{ TimeCreated = [datetime] '2026-10-11T08:18:00'; Id = 2004; ProviderName = 'Demo.Inventory'; Level = 3; Message = 'The inventory queue is above its warning threshold.' }
        [pscustomobject]@{ TimeCreated = [datetime] '2026-10-11T08:25:00'; Id = 3002; ProviderName = 'Demo.Sync'; Level = 3; Message = 'Synchronization completed after a retry.' }
    )
}
$records = @($Event | Select-Object -First $MaxEvents | ForEach-Object {
    if (-not $_.TimeCreated -or -not $_.ProviderName -or $null -eq $_.Id -or $null -eq $_.Level) {
        throw 'Events need TimeCreated, Id, ProviderName and numeric Level properties.'
    }
    [pscustomobject]@{
        TimeCreated = [datetime] $_.TimeCreated; EventId = [int] $_.Id
        Provider = [string] $_.ProviderName; Level = [int] $_.Level; Message = [string] $_.Message
    }
})
$providerSummary = @($records | Group-Object Provider | Sort-Object Count -Descending | Select-Object -First 8 | ForEach-Object {
    [pscustomobject]@{ Provider = $_.Name; Records = $_.Count; Critical = @($_.Group | Where-Object Level -EQ 1).Count; Errors = @($_.Group | Where-Object Level -EQ 2).Count; Warnings = @($_.Group | Where-Object Level -EQ 3).Count }
})
$null = New-Item -ItemType Directory -Path $OutputDirectory -Force
$docx = Join-Path $OutputDirectory 'EventIncidentBrief.docx'
$pdf = Join-Path $OutputDirectory 'EventIncidentBrief.pdf'
$xlsx = Join-Path $OutputDirectory 'EventRegister.xlsx'

New-OfficeWord -Path $docx {
    WordHeader -Type Default { WordParagraph -Text 'OPERATIONS / EVENT REVIEW' }
    WordFooter -Type Default { WordPageNumber -IncludeTotalPages }
    WordParagraph {
        WordParagraphStyle -SpacingAfterPoints 12 -KeepWithNext $true
        WordText 'Event incident brief' -Bold -FontSize 26 -Color '17365D'
    }
    WordParagraph -Text $scope
    WordParagraph {
        WordParagraphStyle -SpacingBeforePoints 12 -SpacingAfterPoints 10
        WordText "$($records.Count) retained records" -Bold -FontSize 18 -Color '9A3412'
    }
    WordParagraph -Text 'Use this snapshot to assign investigation. Event counts alone do not establish a root cause or service impact.'
    if ($providerSummary.Count) {
        WordParagraph {
            WordParagraphStyle -SpacingBeforePoints 16 -SpacingAfterPoints 8 -KeepWithNext $true
            WordText 'Most represented providers' -Bold -FontSize 15 -Color '17365D'
        }
        $table = WordTable -InputObject $providerSummary -Style GridTable4Accent1 -Layout AutoFitToWindow -PassThru
        $table | Get-OfficeWordTableCell | ForEach-Object {
            $_.MarginTopCentimeters = 0.14; $_.MarginBottomCentimeters = 0.14
            $_.MarginLeftCentimeters = 0.16; $_.MarginRightCentimeters = 0.16
        }
    } else { WordParagraph -Text 'No matching events were returned. The workbook records that empty result.' }
    WordParagraph {
        WordParagraphStyle -SpacingBeforePoints 16 -SpacingAfterPoints 8 -KeepWithNext $true
        WordText 'Investigation handoff' -Bold -FontSize 15 -Color '17365D'
    }
    WordList -Style Numbered {
        WordListItem -Text 'Review timestamps and full messages in the attached event register.'
        WordListItem -Text 'Confirm user impact and correlate the events with recent changes.'
        WordListItem -Text 'Assign an owner, record the findings and agree a follow-up time.'
    }
    WordParagraph {
        WordParagraphStyle -SpacingBeforePoints 12
        WordText 'Only the top eight providers appear here. The workbook retains every record in this bounded snapshot. Review event messages before sharing them.' -FontSize 10 -Color '475569'
    }
}
Export-OfficeDocumentPdf -InputPath $docx -Path $pdf
New-OfficeExcel -Path $xlsx {
    ExcelSheet 'Events' {
        ExcelGridlines -Hide
        ExcelCell -Address A1 -Value 'Event register'
        ExcelRow -Row 1 -Height 30 -Bold $true -FontName 'Carlito'
        ExcelCell -Address A2 -Value $scope
        if ($records.Count) {
            ExcelTable -InputObject $records -StartRow 4 -TableName EventRegister -TableStyle TableStyleMedium2
            for ($row = 5; $row -lt 5 + $records.Count; $row++) {
                ExcelCell -Row $row -Column 1 -Value $records[$row - 5].TimeCreated -NumberFormat 'yyyy-mm-dd hh:mm'
                ExcelRow -Row $row -Height 30 -WrapText $true
            }
        } else { ExcelCell -Address A4 -Value 'No matching events returned.' }
        ExcelColumn -Column 1 -Width 23
        ExcelColumn -Column 2 -Width 12
        ExcelColumn -Column 3 -Width 28
        ExcelColumn -Column 4 -Width 10
        ExcelColumn -Column 5 -Width 70
        ExcelFreeze -TopRows 4
    }
}
Get-Item -LiteralPath $docx, $pdf, $xlsx | Select-Object Name, Length
