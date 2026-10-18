$markdownPath = '.\Markdown-Publish-Source.md'
$htmlPath = '.\Markdown-Published.html'
MarkdownNew -Path $markdownPath {
    MarkdownHeading -Level 1 -Text 'Operations handbook'
    MarkdownCallout -Kind warning -Title 'Before deployment' -Body 'Confirm the maintenance window.'
    MarkdownTaskList -Items 'Back up configuration', 'Notify users', 'Run health checks'
}

ConvertTo-OfficeMarkdownHtml `
    -Path $markdownPath `
    -OutputPath $htmlPath `
    -DocumentMode `
    -Title 'Operations handbook' `
    -IncludeAnchorLinks `
    -ExternalLinksTargetBlank
