$path = '.\Searchable-Policy.pdf'

PdfNew -Path $path {
    PdfHeading 'Remote access policy'
    PdfParagraph 'Privileged access is reviewed every 30 days.'
    PdfParagraph 'Service owners record approval evidence.'
}

$pages = Get-OfficePdfText -Path $path -ByPage
$pages | Select-Object PageNumber, Text
