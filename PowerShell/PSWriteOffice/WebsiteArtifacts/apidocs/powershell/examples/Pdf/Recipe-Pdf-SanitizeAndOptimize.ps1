$source = '.\Pdf-Delivery-Source.pdf'
$sanitized = '.\Pdf-Delivery-Sanitized.pdf'
$optimized = '.\Pdf-Delivery-Optimized.pdf'
New-OfficePdf -Path $source {
    PdfMetadata -Title 'Delivery copy' -Author 'PSWriteOffice'
    PdfHeading 'Delivery copy'
    PdfParagraph ('Repeated delivery evidence. ' * 60)
}

Optimize-OfficePdf -Sanitize -PassThru -Path $source -OutputPath $sanitized
Optimize-OfficePdf -Path $sanitized -OutputPath $optimized -AllowLarger -PassThru
