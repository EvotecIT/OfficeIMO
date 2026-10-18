$imagePath = '.\Scanned-Page.png'
$inputPdf = '.\Scanned-Report.pdf'
$outputPdf = '.\Scanned-Report-Searchable.pdf'

# Returns text directly. Add -PassThru for confidence, geometry, provider, and diagnostics.
$recognizedText = Get-OfficeImageText -Path $imagePath -Language English, Polish
$recognizedText

# Tesseract is discovered automatically. Missing curated English or Polish data is
# downloaded into OfficeIMO's checksum-verified per-user cache when needed.
ConvertTo-OfficePdfSearchable `
    -Path $inputPdf `
    -OutputPath $outputPdf `
    -Language English, Polish

Get-OfficePdfText -Path $outputPdf
