<#
.SYNOPSIS
Builds an editable workplace-support brief from quarterly demand data.
.DESCRIPTION
The default CSV contains synthetic planning data. Counts, comparisons and the
highlighted demand increase come from that CSV; the proposal is an example,
not a service-level or quality claim.
.EXAMPLE
./Showcase-PowerPoint-ServiceBrief.ps1 -OutputDirectory ./ServiceBrief
#>
#Requires -Modules @{ ModuleName = 'PSWriteOffice'; ModuleVersion = '3.6.0' }
param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '../Documents'),
    [string] $InputCsv = (Join-Path $PSScriptRoot '../Samples/ServiceBrief/Quarterly-Demand.csv'),
    [datetime] $ReviewDate = '2026-09-30',
    [switch] $Show
)

$ErrorActionPreference = 'Stop'
$culture = [Globalization.CultureInfo]::InvariantCulture

# Typed columns: a value that is not a whole number stops the import.
$counts = 'SupportedPeople', 'PreviousQuarterTickets', 'CurrentQuarterTickets'
$columnTypes = @{}
foreach ($column in $counts) { $columnTypes[$column] = [int] }
$sites = @(Import-OfficeCsv -Path $InputCsv -TrimWhitespace -ColumnType $columnTypes -Culture $culture |
    Select-Object Site, Owner, SupportedPeople, PreviousQuarterTickets, CurrentQuarterTickets,
        @{ Name = 'Change'; Expression = { $_.CurrentQuarterTickets - $_.PreviousQuarterTickets } })

if ($sites.Count -lt 1 -or $sites.Count -gt 6) { throw 'Supply one to six service locations for this eight-slide brief.' }
foreach ($site in $sites) {
    if (-not $site.Site -or $site.Site.Length -gt 20) { throw 'Every location needs a Site of at most 20 characters.' }
    if (-not $site.Owner -or $site.Owner.Length -gt 40) { throw "$($site.Site) needs an Owner of at most 40 characters." }
    if ($site.SupportedPeople -le 0 -or @($counts | Where-Object { $site.$_ -lt 0 })) { throw "Counts at $($site.Site) must be nonnegative, with at least one supported person." }
}
if (@($sites.Site | Sort-Object -Unique).Count -ne $sites.Count) { throw 'Service location names must be unique.' }

$people = ($sites | Measure-Object SupportedPeople -Sum).Sum
$previous = ($sites | Measure-Object PreviousQuarterTickets -Sum).Sum
$current = ($sites | Measure-Object CurrentQuarterTickets -Sum).Sum
$changeText = if ($previous) { (($current - $previous) / $previous).ToString('P1', $culture) } else { 'no prior-quarter baseline' }
$priority = $sites | Sort-Object Change -Descending | Select-Object -First 1
$priorityText = if ($priority.Change -gt 0) {
    "$($priority.Site) increased by $($priority.Change) tickets; review its intake with $($priority.Owner)."
} else { 'No location increased its ticket count; review request mix before changing capacity.' }
$dateText = $ReviewDate.ToString('dd MMM yyyy', $culture)
$isSample = [IO.Path]::GetFullPath($InputCsv) -eq [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../Samples/ServiceBrief/Quarterly-Demand.csv'))
$dataLabel = if ($isSample) { 'Synthetic data' } else { 'Supplied CSV' }
$chartData = $sites | Select-Object Site, @{ Name = 'Previous quarter'; Expression = { $_.PreviousQuarterTickets } },
    @{ Name = 'Current quarter'; Expression = { $_.CurrentQuarterTickets } }
$actions = @(
    [pscustomobject]@{ Action = 'Agree intake categories'; Owner = 'Service desk lead'; Review = $ReviewDate.AddDays(15).ToString('dd MMM', $culture) }
    [pscustomobject]@{ Action = "Review $($priority.Site) demand"; Owner = $priority.Owner; Review = $ReviewDate.AddDays(22).ToString('dd MMM', $culture) }
    [pscustomobject]@{ Action = 'Review pilot evidence'; Owner = 'Service owner'; Review = $ReviewDate.AddDays(36).ToString('dd MMM', $culture) }
)

$null = New-Item -ItemType Directory -Path $OutputDirectory -Force
$path = Join-Path $OutputDirectory 'Showcase-PowerPoint-ServiceBrief.pptx'

New-OfficePowerPoint -Path $path -Theme Boardroom -Accent '#008C95' -Purpose 'workplace support planning brief' -Seed 'workplace-support' -Force -Show:$Show {
    PptFooter $dateText -Alignment Left
    PptFooter $dataLabel -Alignment Right

    PptSlide -Eyebrow 'Workplace support' -Title 'Workplace support plan' -Subtitle "Demand review and service pilot | $dateText" -Transition Fade {
        PptSection 'Service planning'
    }

    PptSlide -Title 'A controlled service transition' -Subtitle 'Agree the baseline, test the change and review the evidence.' {
        PptStep 'Baseline' 'Agree request categories and reconcile the quarterly intake.'
        PptStep 'Pilot' 'Test the proposed triage queue with a nominated service owner.'
        PptStep 'Review' 'Check demand, repeat requests and user feedback before rollout.'
        PptStep 'Roll out' 'Publish ownership, working hours and the escalation route.'
    }

    PptSlide -Title 'Service scope' -Subtitle 'Four service areas with explicit ownership and a common intake.' {
        PptCard 'User incidents' -Items 'triage', 'ownership', 'escalation'
        PptCard 'Device lifecycle' -Items 'onboarding', 'replacement', 'returns'
        PptCard 'Access requests' -Items 'approval', 'least privilege', 'evidence'
        PptCard 'Service changes' -Items 'notice', 'handover', 'review'
    }

    PptSlide -Title 'Supported locations' -Subtitle "$people people across $($sites.Count) service locations." {
        foreach ($site in $sites) {
            PptCard $site.Site -Items "$($site.SupportedPeople) people", "$($site.CurrentQuarterTickets) current tickets", $site.Owner
        }
    }

    PptSlide -Title 'What the operating model must provide' -Subtitle 'Practical acceptance criteria for the pilot.' {
        PptCard 'A clear front door' -Items 'published request categories', 'named service owner', 'agreed escalation route'
        PptCard 'A reliable handover' -Items 'agreed service hours', 'useful operating notes', 'acceptance by support owner'
        PptCard 'An evidence-based review' -Items 'quarterly ticket counts', 'repeat request analysis', 'feedback from service users'
    }

    PptSlide -Title 'Demand and the proposed decision' -Subtitle 'Separate the observation, proposed review and decision.' {
        PptCard 'Demand baseline' -Items "$current quarterly tickets", "$changeText from $previous", "$people people supported"
        PptCard 'Priority review' -Items $priorityText, 'Check request mix before changing capacity.'
        PptCard 'Decision requested' -Items 'Nominate the pilot owner.', 'Agree the review date.', 'Assess the pilot before changing service commitments.'
    }

    PptSlide -Title 'Quarterly ticket demand' -Transition PushLeft {
        PptSection 'Demand and actions'
        PptChart -InputObject $chartData -Category Site -Value 'Previous quarter', 'Current quarter' -Type Column -Title "Tickets: $previous previous / $current current ($changeText)"
        PptNotes 'Both series are ticket counts. A change in demand alone does not establish improved service quality or its cause.'
    }

    PptSlide -Title 'Owners and review dates' {
        PptTable $actions
        PptNotes 'Example actions for a proposed pilot. Agree actual owners and dates before using this brief.'
    }
}

$deck = Open-OfficePowerPoint -Path $path
try { Get-OfficePowerPointSlide -Document $deck -Summary } finally { Close-OfficePowerPoint -Document $deck }
