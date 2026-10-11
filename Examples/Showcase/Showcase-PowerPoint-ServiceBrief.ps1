param(
    [switch] $Open,
    [string] $OutputDirectory = (Join-Path $PSScriptRoot 'Output')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop

$documents = $OutputDirectory
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Showcase-PowerPoint-ServiceBrief.pptx'

$process = @(
    [pscustomobject]@{ Title = 'Assess'; Body = 'Review service incidents and confirm accountable owners.' }
    [pscustomobject]@{ Title = 'Prioritize'; Body = 'Compare operational impact and select the next fixes.' }
    [pscustomobject]@{ Title = 'Remediate'; Body = 'Schedule changes with owners and recovery checks.' }
    [pscustomobject]@{ Title = 'Verify'; Body = 'Review outcomes at the next monthly service meeting.' }
)

$cards = @(
    [pscustomobject]@{ Title = 'Identity'; Items = 'owner: Platform|1 incident|stable service'; AccentColor = '#2F80ED' }
    [pscustomobject]@{ Title = 'Backup'; Items = 'owner: Operations|5 incidents|retry tuning'; AccentColor = '#219653' }
    [pscustomobject]@{ Title = 'Remote Access'; Items = 'owner: Security|7 incidents|profile retirement'; AccentColor = '#9B51E0' }
    [pscustomobject]@{ Title = 'Messaging'; Items = 'owner: Collaboration|4 incidents|queue alert review'; AccentColor = '#F2994A' }
)

$coverage = @(
    [pscustomobject]@{ Name = 'Identity'; X = 0.22; Y = 0.42; Detail = 'Platform confirms synchronization health.' }
    [pscustomobject]@{ Name = 'Backup'; X = 0.48; Y = 0.34; Detail = 'Operations owns retry tuning and restore checks.' }
    [pscustomobject]@{ Name = 'Remote Access'; X = 0.68; Y = 0.58; Detail = 'Security retires legacy VPN profiles.' }
    [pscustomobject]@{ Name = 'Messaging'; X = 0.82; Y = 0.36; Detail = 'Collaboration reviews queue alerts.' }
)

$capabilities = @(
    [pscustomobject]@{ Heading = 'Protect access'; Body = 'Retire legacy remote-access profiles without disrupting supported users.'; Items = 'owner: Security|pilot group first|rollback documented' }
    [pscustomobject]@{ Heading = 'Improve recovery'; Body = 'Tune failed backup retries and verify a representative restore.'; Items = 'owner: Operations|restore evidence|completion trend' }
    [pscustomobject]@{ Heading = 'Reduce alert noise'; Body = 'Review repetitive messaging alerts while keeping actionable failures visible.'; Items = 'owner: Collaboration|threshold review|incident follow-up' }
)

$caseStudy = @(
    [pscustomobject]@{ Heading = 'Problem'; Body = 'Repeated backup failures consume operator time and reduce confidence in recovery.' }
    [pscustomobject]@{ Heading = 'Approach'; Body = 'Tune the retry window, review exceptions and run a representative restore.' }
    [pscustomobject]@{ Heading = 'Success measure'; Body = 'Fewer repeat failures, with a recorded restore result and an accountable owner.' }
)

$metrics = @(
    [pscustomobject]@{ Value = '3'; Label = 'priority actions' }
    [pscustomobject]@{ Value = '1'; Label = 'approval requested' }
    [pscustomobject]@{ Value = '30'; Label = 'days to next review' }
)

$chartRows = @(
    [pscustomobject]@{ Service = 'Backup'; Previous = 8; Current = 5 }
    [pscustomobject]@{ Service = 'Remote Access'; Previous = 9; Current = 7 }
    [pscustomobject]@{ Service = 'Messaging'; Previous = 6; Current = 4 }
)

$tableRows = @(
    [pscustomobject]@{ Action = 'Retire legacy VPN profiles'; Owner = 'Security'; Due = '24 May' }
    [pscustomobject]@{ Action = 'Tune backup retries'; Owner = 'Operations'; Due = '31 May' }
    [pscustomobject]@{ Action = 'Review messaging alerts'; Owner = 'Collaboration'; Due = '7 June' }
)

$plan = PptDeckPlan {
    PptPlanSection -Title 'Service-health briefing' -Subtitle 'Monthly review | Synthetic example data' -Seed 'showcase-cover'
    PptPlanProcess -Title 'From review to verified action' -Subtitle 'A monthly service-improvement cycle' -Steps $process -Seed 'delivery-process'
    PptPlanCardGrid -Title 'Service posture' -Subtitle 'Current incidents and named owners for this review.' -Cards $cards -Seed 'product-cards'
    PptPlanCoverage -Title 'Accountable service owners' -Subtitle 'The map illustrates ownership; position does not represent a measured score.' -Locations $coverage -Seed 'coverage-map'
    PptPlanCapability -Title 'Three priorities for this month' -Subtitle 'Approve the actions, then verify their outcomes.' -Sections $capabilities -Seed 'quality-bar'
    PptPlanCaseStudy -Title 'Backup recovery improvement' -Sections $caseStudy -Metrics $metrics -Seed 'designer-case-study'
}

New-OfficePowerPoint -Path $path {
    PptSlideSize -Preset Screen16x9
    PptDesignerDeck -Plan $plan -AccentColor '#008C95' -Seed 'pswriteoffice-showcase' -Purpose 'technical service brief' -Name 'PSWriteOffice Showcase' -FooterLeft 'PSWriteOffice' -FooterRight 'OfficeIMO designer' -CreativeDirectionPack TechnicalMap -LayoutStrategy ContentFirst

    $chartSlide = PptSlide -PassThru
    PptTitle -Slide $chartSlide -Title 'Incident counts by service'
    PptChart -Slide $chartSlide -Type ClusteredColumn -Data $chartRows -CategoryProperty Service -SeriesProperty Previous, Current -Title 'Previous month vs current month' -X 58 -Y 118 -Width 610 -Height 265
    PptNotes -Slide $chartSlide -Text 'Synthetic counts for the same three services over two months. Do not treat them as production measurements.'

    $tableSlide = PptSlide -PassThru
    PptTitle -Slide $tableSlide -Title 'Actions for approval'
    PptTable -Slide $tableSlide -Data $tableRows -X 64 -Y 132 -Width 590 -Height 210
    PptNotes -Slide $tableSlide -Text 'Confirm each owner, due date and acceptance measure before approving this illustrative plan.'

    PptSection -Name 'Service review' -StartSlideIndex 0
    PptSection -Name 'Evidence and actions' -StartSlideIndex 6
    PptTransition -Slide $chartSlide -Transition PushLeft
    Get-OfficePowerPointSlide -Index 0 | PptTransition -Transition Fade
} -Open:$Open

$presentation = Get-OfficePowerPoint -Path $path
$summary = Get-OfficePowerPointSlideSummary -Presentation $presentation
$presentation | Close-OfficePowerPoint

$summary | Select-Object SlideIndex, Title, ShapeCount, TextBoxCount, ChartCount, TableCount, HasNotes
