$form = '.\Change-Request-Form.pdf'
$xfdf = '.\Change-Request-Form.xfdf'
$roundTrip = '.\Change-Request-RoundTrip.pdf'

PdfNew -Path $form {
    PdfHeading 'Change request'
    PdfText 'Owner'
    PdfFormField -Name Owner -Type Text -Value 'Platform' -Width 280
    PdfText 'Decision'
    PdfFormField -Name Decision -Type Choice -Options Approve,Reject,Defer -Value Defer -Width 220
}

Export-OfficePdfXfdf -Path $form -OutputPath $xfdf
Import-OfficePdfXfdf -Path $form -XfdfPath $xfdf -OutputPath $roundTrip
