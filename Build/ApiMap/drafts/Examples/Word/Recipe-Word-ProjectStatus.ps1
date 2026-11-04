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

New-OfficeWord -Path $path -Theme Modern -Force {
    WordHeader { WordParagraph 'Northwind migration | Weekly status' }
    WordFooter { WordText 'Internal | Page '; WordPageNumber -IncludeTotal }

    WordHeading 'Northwind migration'
    WordHeading 'Weekly project status' -Level 2
    WordParagraph {
        WordText 'Overall status: ' -Bold
        WordText 'On track' -Bold -Color Success
        WordText '. The pilot needs an owner decision on the final rollout window.'
    }

    WordHeading 'Executive summary'
    WordList {
        WordListItem 'Core implementation is 72% complete.'
        WordListItem 'No critical defects are open.'
        WordListItem 'The pilot date is the only decision needed this week.'
    }

    WordHeading 'Milestones'
    WordTable $milestones {
        WordTableCondition { $_.Status -eq 'At risk' } -BackgroundColor Danger
        WordTableCondition { $_.Status -eq 'Done' } -BackgroundColor Success
    }

    WordHeading 'Delivery trend'
    WordChart -Type Line -InputObject $trend -Category Week -Value Complete -Title 'Completion by week (%)' -Height 5cm

    WordHeading 'Decision'
    WordParagraph {
        Set-OfficeWordParagraph -KeepWithNext
        WordText 'Approve the pilot window: ' -Bold
        WordCheckBox -Alias PilotApproved -Tag pilot-approved
    }
    WordParagraph {
        WordText 'Target review date: '
        WordDatePicker -Date 2026-09-01 -Alias PilotReviewDate -Tag pilot-review-date
    }
}
