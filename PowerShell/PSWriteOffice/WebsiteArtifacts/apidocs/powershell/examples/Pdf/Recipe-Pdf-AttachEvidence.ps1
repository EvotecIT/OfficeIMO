$evidence = '.\Evidence-Summary.pdf'
$report = '.\Report-With-Evidence.pdf'

New-OfficePdf -Path $evidence {
    PdfHeading 'Evidence summary'
    PdfParagraph 'The review completed successfully.'
}

New-OfficePdf -Path $report {
    PdfTheme Report
    PdfHeading 'Audit report'
    PdfParagraph 'The supporting evidence is embedded in this PDF.'
    PdfAttachment -Path $evidence -Name 'evidence-summary.pdf' -MimeType 'application/pdf' -Relationship Data -Description 'Supporting review evidence'
}
