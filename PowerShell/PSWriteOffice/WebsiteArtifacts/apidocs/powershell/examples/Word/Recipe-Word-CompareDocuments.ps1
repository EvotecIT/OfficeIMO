$approved = '.\Policy-Approved.docx'
$proposed = '.\Policy-Proposed.docx'
$redline = '.\Policy-Changes.docx'

WordNew -Path $approved {
    WordParagraph -Text 'Remote Access Policy' -Style Heading1
    WordParagraph -Text 'Access reviews run every 90 days.'
}

WordNew -Path $proposed {
    WordParagraph -Text 'Remote Access Policy' -Style Heading1
    WordParagraph -Text 'Access reviews run every 30 days.'
    WordParagraph -Text 'Service owners must record evidence.'
}

Compare-OfficeWordDocument -ReferencePath $approved -DifferencePath $proposed -RedlinePath $redline
