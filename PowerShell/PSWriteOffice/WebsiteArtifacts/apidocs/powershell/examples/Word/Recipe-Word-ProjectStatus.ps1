$path = '.\Project-Status.docx'
$milestones = @(
    [pscustomobject]@{ Milestone = 'Discovery'; Owner = 'Product'; Progress = 100; Status = 'Done' }
    [pscustomobject]@{ Milestone = 'Implementation'; Owner = 'Engineering'; Progress = 72; Status = 'On track' }
    [pscustomobject]@{ Milestone = 'Pilot'; Owner = 'Operations'; Progress = 35; Status = 'At risk' }
)
$trend = @(
    [pscustomobject]@{ Week = 'W1'; Complete = 18 }
    [pscustomobject]@{ Week = 'W2'; Complete = 37 }
    [pscustomobject]@{ Week = 'W3'; Complete = 55 }
    [pscustomobject]@{ Week = 'W4'; Complete = 72 }
)

New-OfficeWord -Path $path {
    WordSection {
        WordHeader { WordParagraph -Text 'Northwind migration | Weekly status' -Style Heading2 }
        WordFooter {
            WordText 'Internal | Page '
            WordPageNumber -IncludeTotal
        }

        WordParagraph -Text 'Northwind Migration' -Style Heading1
        WordParagraph -Text 'Weekly project status' -Style Heading2
        WordParagraph -Run @{
            Text  = 'Overall status: ', 'On track', '. The pilot needs an owner decision on the final rollout window.'
            Bold  = $true, $true, $false
            Color = $null, 'SeaGreen', $null
        }

        WordParagraph -Text 'Executive summary' -Style Heading1
        WordList -Style Bulleted {
            WordListItem -Text 'Core implementation is 72% complete.'
            WordListItem -Text 'No critical defects are open.'
            WordListItem -Text 'The pilot date is the only decision needed this week.'
        }

        WordParagraph -Text 'Milestones' -Style Heading1
        WordTable -InputObject $milestones -Style GridTable4Accent1 -Layout AutoFitToWindow {
            WordTableCondition -FilterScript { $_.Status -eq 'At risk' } -BackgroundColor '#FEE2E2'
            WordTableCondition -FilterScript { $_.Status -eq 'Done' } -BackgroundColor '#DCFCE7'
        }

        WordParagraph -Text 'Delivery trend' -Style Heading1 {
            WordParagraphStyle -SpacingBeforePoints 12 -SpacingAfterPoints 6 -KeepWithNext $true
        }
        WordChart -Type Line -Data $trend -Category Week -Value Complete -Title 'Completion by week (%)' -FitToPageWidth -HeightPixels 190

        WordParagraph -Text 'Decision' -Style Heading1 {
            WordParagraphStyle -SpacingBeforePoints 12 -SpacingAfterPoints 6 -KeepWithNext $true
        }
        WordParagraph {
            WordParagraphStyle -SpacingAfterPoints 4 -KeepWithNext $true
            WordText -Bold 'Approve the pilot window: '
            WordCheckBox -Alias 'PilotApproved' -Tag 'pilot-approved'
        }
        WordParagraph {
            WordParagraphStyle -KeepLinesTogether $true
            WordText 'Target review date: '
            WordDatePicker -Date '2026-09-01' -Alias 'PilotReviewDate' -Tag 'pilot-review-date'
        }
    }
}
