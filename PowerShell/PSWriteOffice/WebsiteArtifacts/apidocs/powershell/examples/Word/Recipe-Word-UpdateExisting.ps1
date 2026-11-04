$path = '.\Word-Updated-In-Place.docx'
New-OfficeWord -Path $path {
    WordSection {
        WordParagraph -Text 'FY24 Service Review' -Style Heading1
        WordParagraph {
            WordText 'Open the '
            WordText -Text 'FY24 portal' -Link 'https://reports.example.test/FY24' -Tooltip 'FY24 reports'
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
