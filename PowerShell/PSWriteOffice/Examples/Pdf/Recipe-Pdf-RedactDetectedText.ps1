$source = '.\Pdf-Redaction-Source.pdf'
$redacted = '.\Pdf-Redacted.pdf'
New-OfficePdf -Path $source {
    PdfHeading 'Incident record'
    PdfParagraph 'Visible owner: Operations'
    PdfParagraph 'Secret account: 123-45-6789'
    PdfParagraph 'Visible status: Closed'
}

ConvertTo-OfficePdfRedacted `
    -Path $source `
    -OutputPath $redacted `
    -Text 'Secret account' `
    -FillColor '#111111'
