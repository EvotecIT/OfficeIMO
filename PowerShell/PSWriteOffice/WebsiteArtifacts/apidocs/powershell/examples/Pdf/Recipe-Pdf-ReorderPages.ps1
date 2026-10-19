$source = '.\Review-Pack-Original.pdf'
$reordered = '.\Review-Pack-Reordered.pdf'

PdfNew -Path $source {
    PdfHeading 'Executive summary'
    PdfParagraph 'The review requires one decision.'
    PdfPageBreak
    PdfHeading 'Evidence'
    PdfParagraph 'Supporting measurements and observations.'
    PdfPageBreak
    PdfHeading 'Approval'
    PdfParagraph 'Owner sign-off page.'
}

Move-OfficePdfPage -Path $source -PageRange '3' -BeforePage 1 -OutputPath $reordered
