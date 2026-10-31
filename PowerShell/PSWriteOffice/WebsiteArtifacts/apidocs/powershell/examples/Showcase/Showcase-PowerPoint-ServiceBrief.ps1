<#
.SYNOPSIS
Builds an editable workplace-support brief from quarterly demand data.
.DESCRIPTION
The default CSV contains synthetic planning data. Counts, comparisons and the
highlighted demand increase come from that CSV; the proposal is an example,
not a service-level or quality claim. Requires PSWriteOffice 3.0.8 or newer.
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
Import-Module $ModulePath -MinimumVersion 3.0.8 -ErrorAction Stop
$rows = @(Import-Csv -LiteralPath $InputCsv)
if ($rows.Count -lt 1 -or $rows.Count -gt 6) { throw 'Supply one to six service locations for this eight-slide brief.' }
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
$plan = PptDeckPlan {
    PptPlanSection -Title 'Workplace support plan' -Subtitle "Demand review and service pilot | $dateText" -Seed 'service-cover'
    PptPlanProcess -Title 'A controlled service transition' -Subtitle 'Agree the baseline, test the change and review the evidence.' -Steps $process -Seed 'service-transition'
    PptPlanCardGrid -Title 'Service scope' -Subtitle 'Four service areas with explicit ownership and a common intake.' -Cards $cards -Seed 'service-scope'
    PptPlanCardGrid -Title 'Supported locations' -Subtitle "$people people across $($sites.Count) service locations." -Cards $locationCards -Seed 'service-locations'
    PptPlanCardGrid -Title 'What the operating model must provide' -Subtitle 'Practical acceptance criteria for the pilot.' -Cards $capabilities -Seed 'service-acceptance'
    PptPlanCardGrid -Title 'Demand and the proposed decision' -Subtitle 'Separate the observation, proposed review and decision.' -Cards $decisions -Seed 'service-decision'
}
New-OfficePowerPoint -Path $path {
    Set-OfficePowerPointSlideSize -Preset Screen16x9
    PptDesignerDeck -Plan $plan -AccentColor '#008C95' -Seed 'workplace-support' -Purpose 'workplace support planning brief' -Name 'Workplace support plan' -Eyebrow 'Workplace support' -FooterLeft $dateText -FooterRight $dataLabel -CreativeDirectionPack Boardroom -LayoutStrategy ContentFirst
    $chartSlide = PowerPointSlide -Title 'Quarterly ticket demand' -PassThru
    PowerPointChart -Slide $chartSlide -Type Column -InputObject $chartData -Category Site -Value 'Previous quarter', 'Current quarter' -Title "Tickets: $previous previous / $current current ($changeText)" -X 58 -Y 118 -Width 610 -Height 235
    PowerPointNotes -Slide $chartSlide -Text 'Both series are ticket counts. A change in demand alone does not establish improved service quality or its cause.'
    $tableSlide = PowerPointSlide -Title 'Owners and review dates' -PassThru
    PowerPointTable -Slide $tableSlide -InputObject $actions -X 54 -Y 120 -Width 612 -Height 200
    PowerPointNotes -Slide $tableSlide -Text 'Example actions for a proposed pilot. Agree actual owners and dates before using this brief.'
    PowerPointSection -Name 'Service planning' -Slide 1
    PowerPointSection -Name 'Demand and actions' -Slide 7
    Set-OfficePowerPointSlide -Slide $chartSlide -Transition PushLeft
    Get-OfficePowerPointSlide -Slide 1 | Set-OfficePowerPointSlide -Transition Fade
} -Show:$Open

$presentation = Open-OfficePowerPoint -Path $path
try { Get-OfficePowerPointSlide -Document $presentation -Summary | Select-Object SlideIndex, Title, ShapeCount, TextBoxCount, ChartCount, TableCount, HasNotes }
finally { $presentation | Close-OfficePowerPoint }
