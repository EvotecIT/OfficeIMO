$markdownPath = '.\Markdown-Word-Source.md'
$wordPath = '.\Markdown-Converted.docx'
$roundTripPath = '.\Markdown-Word-RoundTrip.md'
New-OfficeMarkdown -Path $markdownPath {
    MarkdownHeading -Level 1 -Text 'Change plan'
    MarkdownParagraph -Text 'The same source can be reviewed in Word and returned to Markdown.'
    MarkdownList -Items 'Prepare', 'Approve', 'Deploy'
}

Convert-OfficeDocument -To Word -Path $markdownPath -OutputPath $wordPath
Convert-OfficeDocument -To Markdown -Path $wordPath -OutputPath $roundTripPath
