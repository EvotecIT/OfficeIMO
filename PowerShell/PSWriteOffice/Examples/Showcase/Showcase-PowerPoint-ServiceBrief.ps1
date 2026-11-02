<#
.SYNOPSIS
Builds an editable workplace-support brief from quarterly demand data.
.DESCRIPTION
The default CSV contains synthetic planning data. Counts, comparisons and the
highlighted demand increase come from that CSV; the proposal is an example,
not a service-level or quality claim. Requires PSWriteOffice.
.EXAMPLE
./Showcase-PowerPoint-ServiceBrief.ps1 -OutputDirectory ./ServiceBrief
#>
param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '../Documents'),
    [string] $InputCsv = (Join-Path $PSScriptRoot '../Samples/ServiceBrief/Quarterly-Demand.csv'),
    [datetime] $ReviewDate = '2026-09-30',
    [string] $ModulePath = 'PSWriteOffice',
    [switch] $Open
)

$ErrorActionPreference = 'Stop'
Import-Module $ModulePath -ErrorAction Stop
$rows = @(Import-Csv -LiteralPath $InputCsv)
if ($rows.Count -lt 1 -or $rows.Count -gt 6) { throw 'Supply one to six service locations for this brief.' }
$sites = @(foreach ($row in $rows) {
    if ([string]::IsNullOrWhiteSpace($row.Site) -or [string]::IsNullOrWhiteSpace($row.Owner)) {
        throw 'Every service location needs a Site and Owner.'
    }
    foreach ($field in 'Site', 'Owner') {
        $maxLength = if ($field -eq 'Site') { 20 } else { 40 }
        if ($row.$field.Trim().Length -gt $maxLength -or $row.$field -match '[\r\n|]') {
            throw "Use a single line of at most $maxLength characters without | for $field."
        }
    }
    $values = @{}
    foreach ($field in 'SupportedPeople', 'PreviousQuarterTickets', 'CurrentQuarterTickets') {
        $number = 0
        if (-not [int]::TryParse($row.$field, [Globalization.NumberStyles]::None,
            [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
            throw "Use a nonnegative whole number for $field at $($row.Site)."
        }
        $values[$field] = $number
    }
    if ($values.SupportedPeople -eq 0) { throw "SupportedPeople must be positive at $($row.Site)." }
    [pscustomobject]@{
        Site = $row.Site.Trim(); Owner = $row.Owner.Trim()
        SupportedPeople = $values.SupportedPeople
        PreviousQuarterTickets = $values.PreviousQuarterTickets
        CurrentQuarterTickets = $values.CurrentQuarterTickets
        Change = $values.CurrentQuarterTickets - $values.PreviousQuarterTickets
    }
})
if (@($sites.Site | Sort-Object -Unique).Count -ne $sites.Count) { throw 'Service location names must be unique.' }
$people = [long](($sites | Measure-Object SupportedPeople -Sum).Sum)
$previous = [long](($sites | Measure-Object PreviousQuarterTickets -Sum).Sum)
$current = [long](($sites | Measure-Object CurrentQuarterTickets -Sum).Sum)
$change = $current - $previous
$culture = [Globalization.CultureInfo]::InvariantCulture
$changeText = if ($previous -gt 0) { ($change / [double]$previous).ToString('P1', $culture) } else { 'no prior-quarter baseline' }
$prioritySite = $sites | Sort-Object Change -Descending | Select-Object -First 1
$priorityText = if ($prioritySite.Change -gt 0) {
    "$($prioritySite.Site) increased by $($prioritySite.Change) tickets; review its intake with $($prioritySite.Owner)."
} else { 'No location increased its ticket count; review request mix before changing capacity.' }
$dateText = $ReviewDate.ToString('dd MMM yyyy', $culture)
$sampleCsv = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../Samples/ServiceBrief/Quarterly-Demand.csv'))
$dataLabel = if ([IO.Path]::GetFullPath($InputCsv) -eq $sampleCsv) { 'Synthetic data' } else { 'Supplied CSV' }
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'Showcase-PowerPoint-ServiceBrief.pptx'

$process = @(
    [pscustomobject]@{ Title = 'Baseline'; Body = 'Agree request categories and reconcile the quarterly intake.' }
    [pscustomobject]@{ Title = 'Pilot'; Body = 'Test the proposed triage queue with a nominated service owner.' }
    [pscustomobject]@{ Title = 'Review'; Body = 'Check demand, repeat requests and user feedback before rollout.' }
    [pscustomobject]@{ Title = 'Roll out'; Body = 'Publish ownership, working hours and the escalation route.' }
)
$cards = @(
    [pscustomobject]@{ Title = 'User incidents'; Items = 'triage|ownership|escalation'; AccentColor = '#2F80ED' }
    [pscustomobject]@{ Title = 'Device lifecycle'; Items = 'onboarding|replacement|returns'; AccentColor = '#219653' }
    [pscustomobject]@{ Title = 'Access requests'; Items = 'approval|least privilege|evidence'; AccentColor = '#9B51E0' }
    [pscustomobject]@{ Title = 'Service changes'; Items = 'notice|handover|review'; AccentColor = '#F2994A' }
)
$locationCards = @($sites | ForEach-Object {
    [pscustomobject]@{
        Title = $_.Site
        Items = "$($_.SupportedPeople) people|$($_.CurrentQuarterTickets) current tickets|$($_.Owner)"
        AccentColor = '#008C95'
    }
})
$capabilities = @(
    [pscustomobject]@{ Title = 'A clear front door'; Items = 'published request categories|named service owner|agreed escalation route' }
    [pscustomobject]@{ Title = 'A reliable handover'; Items = 'agreed service hours|useful operating notes|acceptance by support owner' }
    [pscustomobject]@{ Title = 'An evidence-based review'; Items = 'quarterly ticket counts|repeat request analysis|feedback from service users' }
)
$decisions = @(
    [pscustomobject]@{ Title = 'Demand baseline'; Items = "$current quarterly tickets|$changeText from $previous|$people people supported" }
    [pscustomobject]@{ Title = 'Priority review'; Items = "$priorityText|Check request mix before changing capacity." }
    [pscustomobject]@{ Title = 'Decision requested'; Items = 'Nominate the pilot owner.|Agree the review date.|Assess the pilot before changing service commitments.' }
)
$chartData = @($sites | ForEach-Object {
    [pscustomobject]@{ Site=$_.Site; 'Previous quarter'=$_.PreviousQuarterTickets; 'Current quarter'=$_.CurrentQuarterTickets }
})
$actions = @(
    [pscustomobject]@{ Action = 'Agree intake categories'; Owner = 'Service desk lead'; Review = $ReviewDate.AddDays(15).ToString('dd MMM', $culture) }
    [pscustomobject]@{ Action = "Review $($prioritySite.Site) demand"; Owner = $prioritySite.Owner; Review = $ReviewDate.AddDays(22).ToString('dd MMM', $culture) }
    [pscustomobject]@{ Action = 'Review pilot evidence'; Owner = 'Service owner'; Review = $ReviewDate.AddDays(36).ToString('dd MMM', $culture) }
)
New-OfficePowerPoint -Path $path -Theme Boardroom -Accent '#008C95' -Seed 'workplace-support' `
    -Purpose 'workplace support planning brief' -LayoutStrategy ContentFirst {
    PowerPointFooter $dateText
    PowerPointFooter $dataLabel -Alignment Right
    $cover = PowerPointSlide -Title 'Workplace support plan' -Subtitle "Demand review and service pilot | $dateText" `
        -Eyebrow 'Workplace support' -Layout Title -Seed 'service-cover' -PassThru
    PowerPointSlide -Title 'A controlled service transition' -Subtitle 'Agree the baseline, test the change and review the evidence.' -Seed 'service-transition' {
        foreach ($step in $process) {
            PowerPointStep -Title $step.Title -Text $step.Body
        }
    }
    PowerPointSlide -Title 'Service scope' -Subtitle 'Four service areas with explicit ownership and a common intake.' -Seed 'service-scope' {
        foreach ($card in $cards) {
            PowerPointCard -Title $card.Title -Items ($card.Items -split '\|') -Color $card.AccentColor
        }
    }
    PowerPointSlide -Title 'Supported locations' -Subtitle "$people people across $($sites.Count) service locations." -Seed 'service-locations' {
        foreach ($card in $locationCards) {
            PowerPointCard -Title $card.Title -Items ($card.Items -split '\|') -Color $card.AccentColor
        }
    }
    PowerPointSlide -Title 'What the operating model must provide' -Subtitle 'Practical acceptance criteria for the pilot.' -Seed 'service-acceptance' {
        foreach ($card in $capabilities) {
            PowerPointCard -Title $card.Title -Items ($card.Items -split '\|')
        }
    }
    PowerPointSlide -Title 'Demand and the proposed decision' -Subtitle 'Separate the observation, proposed review and decision.' -Seed 'service-decision' {
        foreach ($card in $decisions) {
            PowerPointCard -Title $card.Title -Items ($card.Items -split '\|')
        }
    }
    $chartSlide = PowerPointSlide -Title 'Quarterly ticket demand' -PassThru {
        $chart = PowerPointChart -Type Column -InputObject $chartData -Category Site -Value 'Previous quarter', 'Current quarter' `
            -Title "Tickets: $previous previous / $current current ($changeText)" -PassThru
        # A zero baseline keeps small changes in ticket demand in proportion.
        $chart.SetValueAxisScale(0, $null, $null, $null, $null, $null, $null) | Out-Null
        PowerPointNotes 'Both series are ticket counts. A change in demand alone does not establish improved service quality or its cause.'
    }
    PowerPointSlide -Title 'Owners and review dates' {
        PowerPointTable -InputObject $actions
        PowerPointNotes 'Example actions for a proposed pilot. Agree actual owners and dates before using this brief.'
    }
    PowerPointSection -Name 'Service planning' -Slide $cover
    PowerPointSection -Name 'Demand and actions' -Slide $chartSlide
    Set-OfficePowerPointSlide -Slide $chartSlide -Transition PushLeft
    Set-OfficePowerPointSlide -Slide $cover -Transition Fade
} -Show:$Open

$presentation = Open-OfficePowerPoint -Path $path
try { Get-OfficePowerPointSlide -Document $presentation -Summary | Select-Object SlideIndex, Title, ShapeCount, TextBoxCount, ChartCount, TableCount, HasNotes }
finally { $presentation | Close-OfficePowerPoint }
