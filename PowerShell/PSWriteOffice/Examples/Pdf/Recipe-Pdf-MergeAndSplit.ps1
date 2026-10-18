$first = '.\Pdf-Pack-Part-A.pdf'
$second = '.\Pdf-Pack-Part-B.pdf'
$merged = '.\Pdf-Combined-Pack.pdf'
$splitDirectory = '.\Pdf-Combined-Pack-Pages'

PdfNew -Path $first {
    PdfHeading 'Part A'
    PdfParagraph 'Operations summary.'
}

PdfNew -Path $second {
    PdfHeading 'Part B'
    PdfParagraph 'Detailed evidence.'
}

Join-OfficePdf -Path $first,$second -OutputPath $merged -PageSize A4 -ResizeMode Fit -ResizeMargin 18
Split-OfficePdf -Path $merged -OutputDirectory $splitDirectory -Prefix 'page' -PagesPerDocument 1
