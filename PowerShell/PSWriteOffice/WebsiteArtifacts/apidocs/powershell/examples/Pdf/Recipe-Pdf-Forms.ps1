$path = '.\Pdf-Approval-Form.pdf'
PdfNew -Path $path {
    PdfHeading 'Change approval'
    PdfParagraph 'Complete the fields before submitting the document.'
    PdfText 'Reviewer name'
    PdfFormField -Name Reviewer -Type Text -Value 'Unassigned' -Width 320 -Height 24
    PdfText 'Decision'
    PdfFormField -Name Decision -Type Choice -Options Approve,Reject,Defer -Value Defer -Width 220 -Height 24
    PdfText 'Evidence attached'
    PdfFormField -Name EvidenceAttached -Type CheckBox -Checked -Width 18 -Height 18
}
