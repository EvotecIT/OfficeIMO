$path = '.\Markdown-Inspection.md'
MarkdownNew -Path $path {
    MarkdownFrontMatter -Data @{ title = 'Service review'; owner = 'Operations' }
    MarkdownHeading -Level 1 -Text 'Service review'
    MarkdownParagraph -Text 'This document is inspected without regex parsing.'
    MarkdownHeading -Level 2 -Text 'Controls'
    MarkdownTable -InputObject @([pscustomobject]@{ Control = 'Backups'; Status = 'Ready' })
}

Get-OfficeMarkdownFrontMatter -Path $path
Get-OfficeMarkdownHeading -Path $path
Get-OfficeMarkdownTable -Path $path -AsObject
