$markdownPath = '.\Customer-Handoff.md'
$wordPath = '.\Customer-Handoff.docx'
$pdfPath = '.\Customer-Handoff.pdf'

$document = New-OfficeMarkdown -Path $markdownPath -NoSave
Add-OfficeMarkdownHeading -Document $document -Level 1 -Text 'Customer handoff'
Add-OfficeMarkdownParagraph -Document $document -Text 'The service is ready for acceptance.'
Add-OfficeMarkdownTaskList -Document $document -Items 'Review the evidence', 'Confirm the owner', 'Approve the handoff'
$document | Save-OfficeMarkdown -Path $markdownPath

ConvertFrom-OfficeWordMarkdown -Document $document -OutputPath $wordPath
$word = Get-OfficeWord -Path $wordPath
$word | Export-OfficeDocumentPdf -Path $pdfPath
$word | Close-OfficeWord
