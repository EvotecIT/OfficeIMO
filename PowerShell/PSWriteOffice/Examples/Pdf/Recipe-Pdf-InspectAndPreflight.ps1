$path = '.\Pdf-Inspection-Source.pdf'
PdfNew -Path $path -CreateOutlineFromHeadings {
    PdfMetadata -Title 'Inspection source' -Author 'PSWriteOffice'
    PdfHeading 'Summary'
    PdfParagraph 'The service is ready for review.'
    PdfPageBreak
    PdfHeading 'Evidence'
    PdfParagraph 'Evidence is retained for 90 days.'
}

Get-OfficePdfInfo -Path $path
Get-OfficePdfPreflight -Path $path
Get-OfficePdfText -Path $path -ByPage
