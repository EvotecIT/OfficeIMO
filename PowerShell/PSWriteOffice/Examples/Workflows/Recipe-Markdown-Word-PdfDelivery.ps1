$markdownPath = '.\Customer-Handoff.md'
$wordPath = '.\Customer-Handoff.docx'
$pdfPath = '.\Customer-Handoff.pdf'

$document = New-OfficeMarkdown -Path $markdownPath
Add-OfficeMarkdownHeading -Document $document -Level 1 -Text 'Customer handoff'
Add-OfficeMarkdownParagraph -Document $document -Text 'The service is ready for acceptance.'
Add-OfficeMarkdownTaskList -Document $document -Items 'Review the evidence', 'Confirm the owner', 'Approve the handoff'
$document | Save-OfficeMarkdown -Path $markdownPath

Convert-OfficeDocument -Document $document -To Word -OutputPath $wordPath
$word = Open-OfficeWord -Path $wordPath
$word | Convert-OfficeDocument -To Pdf -OutputPath $pdfPath
$word | Close-OfficeWord
