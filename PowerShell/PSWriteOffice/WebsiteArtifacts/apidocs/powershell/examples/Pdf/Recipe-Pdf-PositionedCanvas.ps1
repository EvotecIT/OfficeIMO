$source = '.\Pdf-Canvas-Source.pdf'
$positioned = '.\Pdf-Positioned-Canvas.pdf'
PdfNew -Path $source {
    PdfHeading 'Fixed-position review copy'
    PdfParagraph 'Normal PdfText and PdfParagraph content remains in document flow.'
    PdfPageBreak
    PdfParagraph 'The canvas callback can place content differently on every page.'
}

Add-OfficePdfCanvas -Path $source -OutputPath $positioned -Content {
    PdfCanvasText -Run @(
        TextRun 'Owner: ' -Bold
        TextRun 'Platform' -Color '#0F766E'
        TextRun '  |  REVIEW COPY' -Italic
    ) -X 36 -Y 24 -FontSize 10

    PdfCanvasText 'Fixed at 36 × 780 points' -X 36 -Y 780 -FontSize 9
}
