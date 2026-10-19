BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Word Markdown conversions' {
    It 'converts Word to Markdown text and file' {
        $docPath = Join-Path $TestDrive 'MarkdownSource.docx'
        $markdownPath = Join-Path $TestDrive 'MarkdownSource.md'

        New-OfficeWord -Path $docPath {
            Add-OfficeWordParagraph -Text 'Quarterly Report' -Style Heading1
            Add-OfficeWordParagraph -Text 'Hello Markdown'
        } | Out-Null

        $markdown = ConvertTo-OfficeWordMarkdown -Path $docPath
        $markdown | Should -Match '# Quarterly Report'
        $markdown | Should -Match 'Hello Markdown'

        $file = ConvertTo-OfficeWordMarkdown -Path $docPath -OutputPath $markdownPath -PassThru
        $file | Should -BeOfType System.IO.FileInfo
        (Get-Content -Path $markdownPath -Raw) | Should -Match 'Quarterly Report'
    }

    It 'does not export Word images to files when WhatIf skips file image export' {
        $docPath = Join-Path $TestDrive 'MarkdownImageSource.docx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive
        $imageDirectory = Join-Path $TestDrive 'images'

        New-OfficeWord -Path $docPath {
            WordParagraph {
                WordImage -Path $imagePath | Out-Null
            }
        } | Out-Null

        ConvertTo-OfficeWordMarkdown -Path $docPath -ImageExportMode File -ImageDirectory $imageDirectory -WhatIf | Out-Null

        Test-Path -LiteralPath $imageDirectory | Should -BeFalse
    }

    It 'converts Markdown text and Markdown documents to Word' {
        $docPath = Join-Path $TestDrive 'MarkdownRoundtrip.docx'
        $markdown = "# Title`n`nBody text"

        $file = ConvertFrom-OfficeWordMarkdown -Markdown $markdown -OutputPath $docPath -PassThru
        $file | Should -BeOfType System.IO.FileInfo
        Test-Path $docPath | Should -BeTrue

        $paragraphs = Get-OfficeWordParagraph -Path $docPath
        ($paragraphs.Text -join "`n") | Should -Match 'Title'
        ($paragraphs.Text -join "`n") | Should -Match 'Body text'

        $markdownDocument = Get-OfficeMarkdown -Text "## Pipeline`n`nGenerated"
        $document = $markdownDocument | ConvertFrom-OfficeWordMarkdown
        try {
            $document.Paragraphs.Count | Should -BeGreaterThan 0
        } finally {
            $document.Dispose()
        }
    }

    It 'downloads explicitly allowed remote Markdown images' {
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive -Name 'RemoteMarkdown.bmp'
        $docPath = Join-Path $TestDrive 'RemoteMarkdown.docx'
        $server = Start-TestHttpFileServer -FilePath $imagePath -ContentType 'image/bmp'

        try {
            ConvertFrom-OfficeWordMarkdown -Markdown "![Remote image]($($server.Url))" -AllowRemoteImages -OutputPath $docPath | Out-Null

            $document = Get-OfficeWord -Path $docPath -ReadOnly
            try {
                $document.Images.Count | Should -Be 1
            } finally {
                $document | Close-OfficeWord
            }
        } finally {
            Stop-TestHttpFileServer -Server $server
        }
    }

    It 'inserts Markdown into a Word template bookmark' {
        $templatePath = Join-Path $TestDrive 'Template.docx'
        $docPath = Join-Path $TestDrive 'TemplateOutput.docx'

        New-OfficeWord -Path $templatePath {
            Add-OfficeWordParagraph -Text 'Before template content'
            Add-OfficeWordParagraph {
                Add-OfficeWordText -Text 'PLACEHOLDER'
                Add-OfficeWordBookmark -Name 'MainContent'
            }
            Add-OfficeWordParagraph -Text 'After template content'
        } | Out-Null

        $markdown = @'
---
title: Hidden metadata
---
# Inserted heading

Inserted body.
'@

        $file = ConvertFrom-OfficeWordMarkdown -Markdown $markdown -TemplatePath $templatePath -BookmarkName 'MainContent' -OutputPath $docPath -PassThru
        $file | Should -BeOfType System.IO.FileInfo

        $paragraphs = @(Get-OfficeWordParagraph -Path $docPath)
        $texts = @($paragraphs.Text | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        $texts | Should -Contain 'Before template content'
        $texts | Should -Contain 'Inserted heading'
        $texts | Should -Contain 'Inserted body.'
        $texts | Should -Contain 'After template content'
        $texts | Should -Not -Contain 'PLACEHOLDER'
        ($texts -join "`n") | Should -Not -Match 'Hidden metadata'
        [Array]::IndexOf($texts, 'Inserted heading') | Should -BeGreaterThan ([Array]::IndexOf($texts, 'Before template content'))
        [Array]::IndexOf($texts, 'After template content') | Should -BeGreaterThan ([Array]::IndexOf($texts, 'Inserted body.'))
    }

    It 'rejects template insertion selectors without a template path' {
        { ConvertFrom-OfficeWordMarkdown -Markdown '# Missing template' -BookmarkName 'MainContent' -ErrorAction Stop } |
            Should -Throw '*Template insertion parameters require -TemplatePath*'
    }
}
