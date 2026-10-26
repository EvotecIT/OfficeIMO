<#
.SYNOPSIS
Creates a Word/PDF incident brief and Excel event register from a bounded snapshot.
.DESCRIPTION
The default is synthetic data. -CollectLive reads recent local critical events, warnings
and errors with PSEventViewer. -Event accepts raw or typed Get-EVXEvent results.
Counts describe the retained records, not every event in the log.
.EXAMPLE
./Recipe-PSEventViewer-OfficeReport.ps1 -OutputDirectory ./Output/EventBrief
.EXAMPLE
./Recipe-PSEventViewer-OfficeReport.ps1 -CollectLive -MaxEvents 100 -OutputDirectory ./Output/LiveBrief
.EXAMPLE
$events = @(Get-EVXEvent -LogName System -Level 2,3 -TimePeriod Last24Hours -MaxEvents 100)
./Recipe-PSEventViewer-OfficeReport.ps1 -Event $events -OutputDirectory ./Output/SystemBrief
#>
#Requires -Modules @{ ModuleName = 'PSWriteOffice'; ModuleVersion = '3.6.0' }
[CmdletBinding(DefaultParameterSetName = 'Sample')]
param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot 'Output/EventBrief'),
    [Parameter(Mandatory, ParameterSetName = 'Event')][object[]] $Event,
    [Parameter(Mandatory, ParameterSetName = 'Live')][switch] $CollectLive,
    [Parameter(ParameterSetName = 'Live')][ValidateNotNullOrEmpty()][string] $LogName = 'System',
    [ValidateRange(1, 1000)][int] $MaxEvents = 100
)

$ErrorActionPreference = 'Stop'
switch ($PSCmdlet.ParameterSetName) {
    'Live' {
        Import-Module PSEventViewer -MinimumVersion 4.0.1 -ErrorAction Stop
        $Event = @(Get-EVXEvent -LogName $LogName -Level 1,2,3 -TimePeriod Last24Hours -MaxEvents $MaxEvents -ReadMode Message)
        $scope = "Local $LogName critical/warning/error events / last 24 hours / newest $MaxEvents at most"
    }
    'Event' { $scope = "Provided event snapshot / first $MaxEvents records at most" }
    'Sample' {
        $scope = 'Synthetic incident exercise / 11 October 2026'
        $Event = @(
            [pscustomobject]@{ TimeCreated = [datetime] '2026-10-11T08:10:00'; Id = 1001; ProviderName = 'Demo.Backup'; Level = 2; MachineName = 'DEMO-SRV01'; LogName = 'Application'; Message = 'The overnight backup did not complete.' }
            [pscustomobject]@{ TimeCreated = [datetime] '2026-10-11T08:12:00'; Id = 1001; ProviderName = 'Demo.Backup'; Level = 2; MachineName = 'DEMO-SRV01'; LogName = 'Application'; Message = 'The backup retry did not complete.' }
            [pscustomobject]@{ TimeCreated = [datetime] '2026-10-11T08:18:00'; Id = 2004; ProviderName = 'Demo.Inventory'; Level = 3; MachineName = 'DEMO-SRV02'; LogName = 'Application'; Message = 'The inventory queue is above its warning threshold.' }
            [pscustomobject]@{ TimeCreated = [datetime] '2026-10-11T08:25:00'; Id = 3002; ProviderName = 'Demo.Sync'; Level = 3; MachineName = 'DEMO-SRV02'; LogName = 'System'; Message = 'Synchronization completed after a retry.' }
        )
    }
}

$records = @($Event | Select-Object -First $MaxEvents | ForEach-Object {
    # Typed rule results retain the same raw snapshot under SourceEvent.
    $source = if ($_.PSObject.Properties['SourceEvent'] -and $_.SourceEvent) { $_.SourceEvent } else { $_ }
    if (-not $source.TimeCreated -or -not $source.ProviderName -or $null -eq $source.Id -or $null -eq $source.Level) {
        throw 'Supply raw events with TimeCreated, Id, ProviderName and numeric Level, or typed results with that SourceEvent snapshot.'
    }
    [pscustomobject]@{
        TimeCreated = [datetime] $source.TimeCreated; EventId = [int] $source.Id
        Provider = [string] $source.ProviderName; Level = [int] $source.Level
        Computer = [string] $source.PSObject.Properties['MachineName'].Value
        Log = [string] $source.PSObject.Properties['LogName'].Value
        Message = [string] $source.PSObject.Properties['Message'].Value
    }
})
$computerCount = @($records.Computer | Where-Object { $_ } | Sort-Object -Unique).Count
$logCount = @($records.Log | Where-Object { $_ } | Sort-Object -Unique).Count
$providerSummary = @($records | Group-Object Provider | Sort-Object Count -Descending | Select-Object -First 8 | ForEach-Object {
    [pscustomobject]@{
        Provider = $_.Name; Records = $_.Count
        Critical = @($_.Group | Where-Object Level -EQ 1).Count
        Errors = @($_.Group | Where-Object Level -EQ 2).Count
        Warnings = @($_.Group | Where-Object Level -EQ 3).Count
    }
})
$null = New-Item -ItemType Directory -Path $OutputDirectory -Force
$docx = Join-Path $OutputDirectory 'EventIncidentBrief.docx'
$pdf = Join-Path $OutputDirectory 'EventIncidentBrief.pdf'
$xlsx = Join-Path $OutputDirectory 'EventRegister.xlsx'

New-OfficeWord -Path $docx -Theme Modern -Force {
    WordHeader { WordParagraph 'Operations / event review' }
    WordFooter { WordPageNumber -IncludeTotal }
    WordHeading 'Event incident brief'
    WordParagraph $scope
    WordParagraph { WordText "$($records.Count) retained records" -Bold -Size 18 -Color Warning }
    WordParagraph 'Use this snapshot to assign investigation. Event counts alone do not establish a root cause or service impact.'
    WordParagraph "$computerCount named source computers / $logCount named source logs. The register preserves each record's source where supplied."
    if ($providerSummary.Count) {
        WordHeading 'Most represented providers' -Level 2
        WordTable $providerSummary
    } else {
        WordParagraph 'No matching events were returned. The workbook records that empty result.'
    }
    WordHeading 'Investigation handoff' -Level 2
    WordList -Style Numbered {
        WordListItem 'Review source computers, logs, timestamps and full messages in the attached event register.'
        WordListItem 'Confirm user impact and correlate the events with recent changes.'
        WordListItem 'Assign an owner, record the findings and agree a follow-up time.'
    }
    WordParagraph {
        WordText 'Only the top eight providers appear here. The workbook retains every record in this bounded snapshot. Review event messages before sharing them.' -Size 10 -Color TextMuted
    }
}
$null = Convert-OfficeDocument -Path $docx -OutputPath $pdf

New-OfficeExcel -Path $xlsx -Theme Modern -Force {
    ExcelReportSheet 'Events' -AutoFitRows {
        ExcelReportTitle 'Event register' $scope
        if ($records.Count) {
            ExcelReportTable $records -Title 'Retained records'
        } else {
            ExcelReportParagraph 'No matching events returned.'
        }
    }
}

Get-Item -LiteralPath $docx, $pdf, $xlsx | Select-Object Name, Length
