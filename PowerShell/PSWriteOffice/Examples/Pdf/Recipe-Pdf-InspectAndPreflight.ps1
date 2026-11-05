$path = '.\Pdf-Inspection-Source.pdf'
New-OfficePdf -Path $path -CreateOutlineFromHeadings {
    PdfMetadata -Title 'Inspection source' -Author 'PSWriteOffice'
    PdfHeading 'Summary'
    PdfParagraph 'The service is ready for review.'
    PdfPageBreak
    PdfHeading 'Evidence'
    PdfParagraph 'Evidence is retained for 90 days.'
}

(Get-OfficeDocumentInfo -Detailed -Path $path).Details
Get-OfficePdfDiagnostic -Type Preflight -Path $path
Read-OfficeDocument -As Text -Path $path -ByPage
