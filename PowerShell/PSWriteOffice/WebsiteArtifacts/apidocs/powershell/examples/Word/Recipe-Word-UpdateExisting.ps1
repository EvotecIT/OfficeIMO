$path = '.\Word-Updated-In-Place.docx'
WordNew -Path $path {
    WordSection {
        WordParagraph -Text 'FY24 Service Review' -Style Heading1
        WordParagraph {
            WordText 'Open the '
            WordHyperlink -Text 'FY24 portal' -Url 'https://reports.example.test/FY24' -Tooltip 'FY24 reports'
            WordText ' for supporting evidence.'
        }
        WordParagraph {
            WordText 'Summary'
            WordBookmark -Name 'FY24Summary'
        }
    }
}

Update-OfficeWordText -Path $path -OldValue 'FY24' -NewValue 'FY25' `
    -IncludeHyperlinkText -IncludeHyperlinkUri -IncludeHyperlinkTooltip -IncludeHyperlinkAnchor
