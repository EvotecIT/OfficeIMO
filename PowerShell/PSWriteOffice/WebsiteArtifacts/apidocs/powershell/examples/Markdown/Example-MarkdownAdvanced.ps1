Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Example-MarkdownAdvanced.md'
$data = @(
    [pscustomobject]@{ Metric = 'Latency'; Value = '120ms' }
    [pscustomobject]@{ Metric = 'Errors'; Value = '0.2%' }
)

New-OfficeMarkdown -Path $path {
    MarkdownFrontMatter -Data @{
        title = 'Operations Report'
        tags  = @('ops', 'weekly')
    }
    MarkdownTableOfContents -Title 'Contents' -PlaceAtTop -MinLevel 2 -MaxLevel 3
    MarkdownHeading -Level 1 -Text 'Operations Report'
    MarkdownParagraph -Text 'Generated with PSWriteOffice.'
    MarkdownCallout -Kind 'note' -Title 'Reminder' -Body 'Review the incident log before publishing.'
    MarkdownHeading -Level 2 -Text 'Key Metrics'
    MarkdownTable -InputObject $data
    MarkdownHeading -Level 2 -Text 'Checklist'
    MarkdownTaskList -Items 'Draft', 'Review', 'Publish' -Completed 1
    MarkdownHeading -Level 2 -Text 'Glossary'
    MarkdownDefinitionList -Definition @{
        SLA = 'Service level agreement'
        SLO = 'Service level objective'
    }
    MarkdownDetails -Summary 'Implementation Notes' {
        MarkdownParagraph -Text 'This section stays collapsed in supporting renderers.'
        MarkdownList -Items 'Thin PowerShell wrapper', 'Backed by OfficeIMO.Markdown'
    }
    MarkdownCode -Language 'powershell' -Content 'Get-Service | Select-Object -First 5'
    MarkdownHorizontalRule
    MarkdownQuote -Text 'Availability is a feature.'
}

Write-Host "Markdown saved to $path"
