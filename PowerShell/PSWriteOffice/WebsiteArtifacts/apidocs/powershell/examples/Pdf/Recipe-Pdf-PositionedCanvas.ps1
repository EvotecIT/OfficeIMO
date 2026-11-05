$source = '.\Pdf-Canvas-Source.pdf'
$positioned = '.\Pdf-Positioned-Canvas.pdf'
New-OfficePdf -Path $source {
    PdfHeading 'Fixed-position review copy'
    PdfParagraph 'Normal PdfText and PdfParagraph content remains in document flow.'
    PdfPageBreak
    PdfParagraph 'The canvas callback can place content differently on every page.'
}

Add-OfficePdfStamp -Path $source -OutputPath $positioned -Content {
    PdfCanvasText -Content {
        PdfText 'Owner: ' -Bold
        PdfText 'Platform' -Color '#0F766E'
        PdfText '  |  REVIEW COPY' -Italic
    } -X 36 -Y 24 -Size 10

    PdfCanvasText 'Fixed at 36 × 780 points' -X 36 -Y 780 -Size 9
}
