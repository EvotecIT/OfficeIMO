$first = '.\Pdf-Pack-Part-A.pdf'
$second = '.\Pdf-Pack-Part-B.pdf'
$merged = '.\Pdf-Combined-Pack.pdf'

New-OfficePdf -Path $first -Force {
    PdfHeading 'Part A'
    PdfParagraph 'Operations summary.'
}
New-OfficePdf -Path $second -Force {
    PdfHeading 'Part B'
    PdfParagraph 'Detailed evidence.'
}

# Merge onto A4 pages, then split back into one file per page.
Merge-OfficeDocument -Path $first, $second -OutputPath $merged -Options (New-OfficePdfMergeOptions -PageSize A4 -ResizeMode Fit -ResizeMargin 18pt)
Split-OfficeDocument -Path $merged -OutputDirectory .\Pdf-Combined-Pack-Pages -Prefix page -PageCount 1
