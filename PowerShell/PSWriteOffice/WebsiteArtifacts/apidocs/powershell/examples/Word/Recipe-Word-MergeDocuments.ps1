$cover = '.\Word-Pack-Cover.docx'
$detail = '.\Word-Pack-Detail.docx'
$appendix = '.\Word-Pack-Appendix.docx'
$merged = '.\Word-Combined-Pack.docx'

New-OfficeWord -Path $cover {
    WordSection {
        WordParagraph -Text 'Operations Pack' -Style Heading1
    }
}

New-OfficeWord -Path $detail {
    WordSection {
        WordParagraph -Text 'Current Status' -Style Heading1
        WordParagraph -Text 'All core services are available.'
    }
}

New-OfficeWord -Path $appendix {
    WordSection {
        WordParagraph -Text 'Appendix' -Style Heading1
        WordParagraph -Text 'Evidence retained for 90 days.'
    }
}

Join-OfficeWordDocument -Path $cover -AppendPath $detail,$appendix -OutputPath $merged
