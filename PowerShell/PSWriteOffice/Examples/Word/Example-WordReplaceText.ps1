Import-Module PSWriteOffice -ErrorAction Stop
$Path = Join-Path $PSScriptRoot 'Example-WordReplaceText.docx'

New-OfficeWord -Path $Path {
    WordParagraph {
        WordText 'FY24 status is ready for review.'
        WordHyperlink -Text 'Portal FY24' -Url 'https://old.example.com/FY24' -Tooltip 'FY24 portal'
    }

    WordParagraph {
        WordHyperlink -Text 'Jump to FY24 summary' -Anchor 'FY24Summary' -Tooltip 'FY24 section'
    }

    WordParagraph {
        WordText 'Summary'
        WordBookmark -Name 'FY24Summary'
    }
}

Update-OfficeWordText -Path $Path -OldValue 'FY24' -NewValue 'FY25' -IncludeHyperlinkText -IncludeHyperlinkUri -IncludeHyperlinkAnchor -IncludeHyperlinkTooltip

Get-Item $Path
