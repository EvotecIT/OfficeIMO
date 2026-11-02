[CmdletBinding()]
param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '../Documents')
)

Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$presentationPath = Join-Path $OutputDirectory 'PowerPoint-HtmlReview.pptx'
$semanticHtmlPath = Join-Path $OutputDirectory 'PowerPoint-HtmlReview.semantic.html'
$visualHtmlPath = Join-Path $OutputDirectory 'PowerPoint-HtmlReview.visual.html'

New-OfficePowerPoint -Path $presentationPath {
    PowerPointSlide -Title 'Monthly Service Review' {
        PowerPointParagraph 'Identity, Messaging, and Reporting are ready for leadership review.'
        PowerPointNotes 'Introduce the operational status summary.'
    }
    PowerPointSlide -Title 'Open Items' {
        PowerPointTable -InputObject @(
            [pscustomobject]@{
                Area = 'Messaging'
                Owner = 'Collaboration'
                NextStep = 'Review retry spikes'
            }
            [pscustomobject]@{
                Area = 'Reporting'
                Owner = 'Analytics'
                NextStep = 'Publish refreshed dashboard'
            }
        ) -Property Area, Owner, NextStep
    }
}

$semantic = New-OfficePowerPointHtmlOptions -Theme Modern -Title 'Deck Review'
$visual = New-OfficePowerPointHtmlOptions -Profile VisualReview -Theme Modern -Title 'Deck Visual Review'
Convert-OfficeDocument -Path $presentationPath -To Html -OutputPath $semanticHtmlPath -Options $semantic
Convert-OfficeDocument -Path $presentationPath -To Html -OutputPath $visualHtmlPath -Options $visual
