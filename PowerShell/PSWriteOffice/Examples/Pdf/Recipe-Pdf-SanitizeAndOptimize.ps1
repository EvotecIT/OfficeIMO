$source = '.\Pdf-Delivery-Source.pdf'
$sanitized = '.\Pdf-Delivery-Sanitized.pdf'
$optimized = '.\Pdf-Delivery-Optimized.pdf'
PdfNew -Path $source {
    PdfMetadata -Title 'Delivery copy' -Author 'PSWriteOffice'
    PdfHeading 'Delivery copy'
    PdfParagraph ('Repeated delivery evidence. ' * 60)
}

ConvertTo-OfficePdfSanitized -Path $source -OutputPath $sanitized
ConvertTo-OfficePdfOptimized -Path $sanitized -OutputPath $optimized -AllowLarger -PassThruReport
