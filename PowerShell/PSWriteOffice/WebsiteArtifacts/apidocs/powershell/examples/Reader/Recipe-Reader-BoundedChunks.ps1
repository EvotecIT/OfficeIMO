$source = '.\Reader-Chunk-Source.md'

MarkdownNew -Path $source {
    MarkdownHeading -Level 1 -Text 'Operations'
    MarkdownParagraph -Text 'Identity is healthy and the weekly review is complete.'
    MarkdownHeading -Level 2 -Text 'Actions'
    MarkdownList -Items 'Archive the evidence', 'Notify the service owner'
}

Get-OfficeDocumentChunk -Path $source -MaxChars 300 -MaxInputBytes 1048576 -MaxTableRows 50 |
    Select-Object SourcePath, Kind, HeadingPath, Text
