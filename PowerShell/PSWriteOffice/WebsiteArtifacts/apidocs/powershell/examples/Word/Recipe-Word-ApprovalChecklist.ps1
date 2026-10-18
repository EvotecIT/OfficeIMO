$path = '.\Change-Approval-Checklist.docx'
$checks = @(
    [pscustomobject]@{ Check = 'Rollback plan attached'; Owner = 'Engineering'; Required = 'Yes' }
    [pscustomobject]@{ Check = 'Monitoring updated'; Owner = 'Operations'; Required = 'Yes' }
    [pscustomobject]@{ Check = 'Customer notice reviewed'; Owner = 'Support'; Required = 'No' }
)

WordNew -Path $path {
    WordSection {
        WordHeader { WordParagraph -Text 'Production change control' -Style Heading2 }
        WordFooter { WordPageNumber -IncludeTotalPages }

        WordParagraph -Text 'Change Approval Checklist' -Style Heading1
        WordParagraph -Text 'Complete this document before the production window opens.'
        WordTableOfContents -Style Template1

        WordParagraph -Text 'Change details' -Style Heading1
        WordParagraph {
            WordText 'Risk level: '
            WordDropDownList -Items 'Low', 'Medium', 'High' -Alias 'RiskLevel' -Tag 'risk-level'
        }
        WordParagraph {
            WordText 'Planned date: '
            WordDatePicker -Date '2026-09-15' -Alias 'PlannedDate' -Tag 'planned-date'
        }

        WordParagraph -Text 'Required evidence' -Style Heading1
        WordTable -InputObject $checks -Style GridTable1LightAccent1 -Layout AutoFitToWindow {
            WordTableCondition -FilterScript { $_.Required -eq 'Yes' } -BackgroundColor '#FEF3C7'
        }

        WordParagraph -Text 'Approvals' -Style Heading1
        foreach ($role in 'Technical owner', 'Operations', 'Change manager') {
            WordParagraph {
                WordText "$role approved: "
                WordCheckBox -Alias ($role -replace ' ', '') -Tag (($role -replace ' ', '-').ToLowerInvariant())
            }
        }

        WordParagraph -Text 'Implementation notes' -Style Heading1
        WordParagraph -Text 'Record the actual start time, validation result, and rollback decision here.'
        WordWatermark -Text 'CHANGE CONTROL'
    }
}
