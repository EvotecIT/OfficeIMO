Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Example-WordAdvanced.docx'
$data = @(
    [pscustomobject]@{ Item = 'Alpha'; Total = 1200 }
    [pscustomobject]@{ Item = 'Beta'; Total = 800 }
    [pscustomobject]@{ Item = 'Gamma'; Total = 1500 }
)

WordNew -Path $path {
    WordSection {
        WordHeader { WordParagraph -Text 'Project Status' -Style Heading2 }
        WordFooter { WordPageNumber }

        WordParagraph -Text 'Executive Summary' -Style Heading1
        WordParagraph -Text 'This report was generated automatically.'

        WordTableOfContents -Style Template1

        WordParagraph -Text 'Status Overview' -Style Heading2
        WordList -Style Bulleted {
            WordListItem -Text 'Timeline on track'
            WordListItem -Text 'Budget variance under 5%'
        }

        WordParagraph -Text 'Metrics' -Style Heading2
        WordTable -InputObject $data -Style TableGrid {
            WordTableCondition -FilterScript { $_.Total -gt 1000 } -BackgroundColor '#d9eaf7'
        }

        WordParagraph -Text 'Approvals' -Style Heading2
        WordCheckBox -Alias 'Approved' -Tag 'approved'
        WordDatePicker -Alias 'Due Date' -Tag 'due-date'
        WordDropDownList -Alias 'Status' -Tag 'status' -Items 'New','In Progress','Done'

        WordWatermark -Text 'CONFIDENTIAL'
    }
} -Open

Write-Host "Document saved to $path"
