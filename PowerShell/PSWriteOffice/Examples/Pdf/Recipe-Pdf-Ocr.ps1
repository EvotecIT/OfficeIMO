$imagePath = '.\Scanned-Page.png'
$inputPdf = '.\Scanned-Report.pdf'
$outputPdf = '.\Scanned-Report-Searchable.pdf'

# The review retains confidence, geometry, provider, and diagnostics alongside the text.
$ocrOptions = New-OfficeOcrOptions -Language English, Polish
$review = Read-OfficeDocument -As Ocr -Path $imagePath -OcrOptions $ocrOptions
$review.Ocr.Text

# Tesseract is discovered automatically. Missing curated English or Polish data is
# downloaded into OfficeIMO's checksum-verified per-user cache when needed.
Update-OfficePdf -Searchable `
    -Path $inputPdf `
    -OutputPath $outputPdf `
    -OcrOptions $ocrOptions

Read-OfficeDocument -As Text -Path $outputPdf
