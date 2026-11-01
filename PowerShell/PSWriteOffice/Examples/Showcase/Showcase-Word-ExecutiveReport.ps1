[CmdletBinding()]
param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot 'Output')
)

$ErrorActionPreference = 'Stop'

Import-Module PSWriteOffice -MinimumVersion 3.0.9 -ErrorAction Stop

$documents = $OutputDirectory
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Showcase-Word-ExecutiveReport.docx'

$services = @(
    [pscustomobject]@{ Service = 'Identity Sync'; Owner = 'Platform'; Status = 'Healthy'; Availability = 99.98; Incidents = 1; Risk = 'Low'; NextAction = 'Keep monitoring password hash sync drift' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration'; Status = 'Watch'; Availability = 99.72; Incidents = 4; Risk = 'Medium'; NextAction = 'Reduce transport queue alert noise' }
    [pscustomobject]@{ Service = 'File Services'; Owner = 'Core IT'; Status = 'Healthy'; Availability = 99.91; Incidents = 2; Risk = 'Low'; NextAction = 'Complete stale share owner review' }
    [pscustomobject]@{ Service = 'Remote Access'; Owner = 'Security'; Status = 'Needs action'; Availability = 98.84; Incidents = 7; Risk = 'High'; NextAction = 'Rotate legacy VPN profiles and publish owner runbook' }
    [pscustomobject]@{ Service = 'Endpoint Backup'; Owner = 'Operations'; Status = 'Watch'; Availability = 99.21; Incidents = 5; Risk = 'Medium'; NextAction = 'Tune failed backup retry window' }
)

$trend = @(
    [pscustomobject]@{ Month = 'Jan'; Availability = 99.20; Incidents = 8 }
    [pscustomobject]@{ Month = 'Feb'; Availability = 99.34; Incidents = 7 }
    [pscustomobject]@{ Month = 'Mar'; Availability = 99.55; Incidents = 6 }
    [pscustomobject]@{ Month = 'Apr'; Availability = 99.61; Incidents = 5 }
    [pscustomobject]@{ Month = 'May'; Availability = 99.73; Incidents = 4 }
    [pscustomobject]@{ Month = 'Jun'; Availability = 99.82; Incidents = 3 }
)

$actions = @(
    [pscustomobject]@{ Priority = 'P1'; Action = 'Retire legacy VPN profiles'; Owner = 'Security'; Due = '2026-05-24'; Outcome = 'Lower high-risk remote-access incidents' }
    [pscustomobject]@{ Priority = 'P2'; Action = 'Tune backup retry window'; Owner = 'Operations'; Due = '2026-05-31'; Outcome = 'Improve endpoint backup completion rate' }
    [pscustomobject]@{ Priority = 'P2'; Action = 'Finalize share owner review'; Owner = 'Core IT'; Due = '2026-06-07'; Outcome = 'Reduce stale access paths before audit' }
)

New-OfficeWord -Path $path {
    WordSection {
        WordPageSetup -PageSize A4 -Top 1134 -Bottom 1134 -Left 1134 -Right 1134
        WordHeader {
            WordParagraph {
                WordText -Bold 'PSWriteOffice Showcase'
                WordText ' | Executive service health'
            }
        }
        WordFooter {
            WordParagraph {
                WordText 'Synthetic service data | Page ' -Size 9
                WordField -Type Page
            }
        }

        Set-OfficeWordDocumentProperty -Name Title -Value 'Executive Service Health Report'
        Set-OfficeWordDocumentProperty -Name Creator -Value 'PSWriteOffice'
        Set-OfficeWordDocumentProperty -Name Category -Value 'Showcase'
        Set-OfficeWordDocumentProperty -Name ShowcaseProduct -Value 'Word' -Custom

        WordParagraph {
            Set-OfficeWordParagraph -SpacingAfter 12 -KeepWithNext:$true
            WordText 'Executive Service Health Report' -Size 24 -Bold -Color '#17365D'
        }
        WordParagraph 'Generated from PowerShell objects with PSWriteOffice and OfficeIMO.'
        WordParagraph {
            WordText -Bold 'Audience: '
            WordText 'technology leadership, service owners, and operational reviewers.'
            WordFootnote 'This sample uses synthetic service-health data generated entirely in PowerShell.'
        }
        WordParagraph {
            WordText -Bold 'Decision needed: '
            WordText 'approve the focused remediation plan for high-friction services and keep monthly trend review cadence.'
            WordBookmark -Name 'ExecutiveSummary'
        }

        WordParagraph -Text 'Executive Summary' -Style Heading1
        WordParagraph 'The portfolio is stable overall, but Remote Access and Endpoint Backup need targeted work before the next governance review.'
        WordList -Style Bulleted {
            WordListItem -Text 'Average availability across tracked services is above 99.5%.'
            WordListItem -Text 'Remote Access is the only high-risk service and owns the most urgent remediation action.'
            WordListItem -Text 'The next review should focus on stale access paths, retry windows, and alert quality.'
        }

        WordParagraph -Text 'Service Scorecard' -Style Heading1
        WordParagraph {
            WordText 'Rows needing attention are highlighted so the table remains useful after export.'
            WordEndnote 'Risk labels combine current incident count, owner feedback, and observed availability trend.'
        }
        $scorecard = $services | Select-Object Service, Owner, Status,
            @{ Name = 'Availability'; Expression = { '{0:N2}%' -f $_.Availability } }, Incidents, Risk
        $table = WordTable -InputObject $scorecard -Style GridTable4Accent1 -Layout AutoFitToWindow -PassThru {
            WordTableCondition -FilterScript { $_.Status -eq 'Needs action' } -BackgroundColor '#fde2e2'
            WordTableCondition -FilterScript { $_.Status -eq 'Watch' } -BackgroundColor '#fff4cc'
        }
        $table | Get-OfficeWordTableCell | Set-OfficeWordTableCell -MarginTopCentimeters 0.12 -MarginBottomCentimeters 0.12 -MarginLeftCentimeters 0.16 -MarginRightCentimeters 0.16

        WordParagraph {
            Set-OfficeWordParagraph -Style Heading1 -PageBreakBefore:$true -KeepWithNext:$true -SpacingAfter 8
            WordText 'Incident trend'
        }
        WordParagraph 'Recorded incidents fell from eight in January to three in June. June availability is 99.82%; percentages and incident counts use different units and are reported separately.'
        WordParagraph {
            Set-OfficeWordParagraph -SpacingBefore 12 -SpacingAfter 12
            WordChart -Type Line -Data $trend -Category Month -Value Incidents -Title 'Incidents by month' -XAxisTitle 'Month' -YAxisTitle 'Incidents' -FitToPageWidth -WidthFraction 1 -HeightPixels 240
        }

        WordParagraph -Text 'Action Plan' -Style Heading1
        $actionTable = WordTable -InputObject ($actions | Select-Object Priority, Action, Owner, Due) -Style GridTable4Accent1 -Layout AutoFitToWindow -PassThru
        $actionTable | Get-OfficeWordTableCell | Set-OfficeWordTableCell -MarginTopCentimeters 0.12 -MarginBottomCentimeters 0.12 -MarginLeftCentimeters 0.16 -MarginRightCentimeters 0.16
        WordParagraph {
            WordText 'Jump back to '
            WordText -Text 'Executive Summary' -Anchor 'ExecutiveSummary'
            WordText ' when reviewing the action plan.'
        }

        WordParagraph -Text 'Approval Controls' -Style Heading1
        WordParagraph {
            WordText 'Approved for publication: '
            WordCheckBox -Alias 'ApprovedForPublication' -Tag 'approval-publish'
        }
        WordParagraph {
            WordText 'Next review date: '
            WordDatePicker -Date (Get-Date '2026-06-15') -Alias 'NextReviewDate' -Tag 'review-date'
        }
        WordParagraph {
            WordText 'Review status: '
            WordDropDownList -Items 'Draft','Ready for review','Approved' -Alias 'ReviewStatus' -Tag 'review-status'
        }

        Update-OfficeWordField
    }
}

$document = Open-OfficeWord -Path $path -ReadOnly
try {
    [pscustomobject]@{
        Path            = $path
        Paragraphs      = $document.Paragraphs.Count
        Tables          = $document.Tables.Count
        Charts          = $document.Charts.Count
        ContentControls = $document.StructuredDocumentTags.Count
        Footnotes       = @(Get-OfficeWordFootnote -Document $document).Count
        Endnotes        = @(Get-OfficeWordEndnote -Document $document).Count
    } | Format-List
} finally {
    $document | Close-OfficeWord
}
