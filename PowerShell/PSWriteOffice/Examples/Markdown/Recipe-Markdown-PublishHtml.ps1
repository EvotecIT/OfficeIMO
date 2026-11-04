$markdownPath = '.\Markdown-Publish-Source.md'
$htmlPath = '.\Markdown-Published.html'
New-OfficeMarkdown -Path $markdownPath {
    MarkdownHeading -Level 1 -Text 'Operations handbook'
    MarkdownCallout -Kind warning -Title 'Before deployment' -Text 'Confirm the maintenance window.'
    MarkdownTaskList -Items 'Back up configuration', 'Notify users', 'Run health checks'
}

Convert-OfficeDocument -To Html -Path $markdownPath -OutputPath $htmlPath -Options (New-OfficeMarkdownHtmlOptions -DocumentMode -Title 'Operations handbook' -IncludeAnchorLinks -ExternalLinksTargetBlank)
