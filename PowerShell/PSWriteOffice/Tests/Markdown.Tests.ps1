BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'Markdown cmdlets' {
    It 'parses Markdown text into a document' {
        $doc = Get-OfficeMarkdown -Text "# Title`n`nHello"
        $doc.GetType().FullName | Should -Be 'OfficeIMO.Markdown.MarkdownDoc'
    }

    It 'supports reader profiles when parsing Markdown' {
        $doc = Get-OfficeMarkdown -Text "---`ntitle: Test`n---`n# Title" -Profile CommonMark
        $doc.HasDocumentHeader | Should -BeFalse
    }

    It 'converts objects to Markdown tables' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $markdown = $rows | ConvertTo-OfficeMarkdown
        $markdown | Should -Match 'Name'
        $markdown | Should -Match 'Value'
    }

    It 'converts Markdown to HTML' {
        $html = ConvertTo-OfficeMarkdownHtml -Text "# Title`n`nHello"
        $html | Should -Match '<h1'
        $html | Should -Match 'Title'
    }

    It 'converts HTML to Markdown' {
        $markdown = ConvertFrom-OfficeMarkdownHtml -Html '<h1>Title</h1><p>Hello</p>'
        $markdown | Should -Match '# Title'
        $markdown | Should -Match 'Hello'
    }

    It 'converts HTML to a Markdown document' {
        $doc = ConvertFrom-OfficeMarkdownHtml -Html '<h1>Title</h1><p>Hello</p>' -AsDocument
        $doc.GetType().FullName | Should -Be 'OfficeIMO.Markdown.MarkdownDoc'
        $doc.ToMarkdown() | Should -Match '# Title'
    }

    It 'gets heading metadata as PowerShell-friendly objects' {
        $headings = Get-OfficeMarkdownHeading -Text "# Report`n`n## Details`n`n### Deep Dive" -MinLevel 2 -MaxLevel 3

        $headings.Count | Should -Be 2
        $headings[0].Text | Should -Be 'Details'
        $headings[0].Anchor | Should -Be 'details'
        $headings[1].Level | Should -Be 3
    }

    It 'gets the Markdown object tree as PowerShell-friendly nodes' {
        $markdown = @"
# Report

> [!NOTE]
> Remember this.

| Name | Value |
| --- | ---: |
| Alpha | 1 |
"@

        $nodes = Get-OfficeMarkdownNode -Text $markdown
        $nodes[0].Type | Should -Be 'MarkdownDoc'
        $nodes.Type | Should -Contain 'HeadingBlock'
        $nodes.Type | Should -Contain 'CalloutBlock'
        $nodes.Type | Should -Contain 'TableBlock'
        $nodes.Type | Should -Contain 'TableCell'
        ($nodes | Where-Object Type -EQ 'HeadingBlock').Text | Should -Be 'Report'
    }

    It 'filters Markdown object tree nodes by type and depth' {
        $nodes = Get-OfficeMarkdownNode -Text "# Report`n`n## Details" -NodeType '*Heading*' -MaxDepth 1

        $nodes.Count | Should -Be 2
        $nodes.Type | Should -Contain 'HeadingBlock'
        ($nodes | Where-Object Type -EQ 'MarkdownDoc').Count | Should -Be 0
    }

    It 'gets front matter entries from Markdown' {
        $frontMatter = Get-OfficeMarkdownFrontMatter -Text "---`ntitle: Report`nowner: Ops`n---`n# Report"

        $frontMatter.Count | Should -Be 2
        $frontMatter[0].Key | Should -Be 'title'
        $frontMatter[0].Value | Should -Be 'Report'
    }

    It 'gets Markdown tables as PowerShell objects' {
        $markdown = @"
| Name | Value |
| --- | ---: |
| Alpha | 1 |
| Beta | 2 |
"@

        $rows = Get-OfficeMarkdownTable -Text $markdown -AsObject

        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'Alpha'
        $rows[0].Value | Should -Be '1'
        $rows[1].Name | Should -Be 'Beta'
    }

    It 'saves converted HTML as Markdown' {
        $path = Join-Path $TestDrive 'converted.md'
        $file = ConvertFrom-OfficeMarkdownHtml -Html '<h1>Title</h1><p>Hello</p>' -OutputPath $path -PassThru

        $file | Should -BeOfType System.IO.FileInfo
        (Get-Content -Path $path -Raw) | Should -Match '# Title'
    }

    It 'does not extract HTML data URI images when WhatIf skips saving' {
        $path = Join-Path $TestDrive 'converted-whatif.md'
        $imageDirectory = Join-Path $TestDrive 'images'
        $html = '<p><img alt="Tiny" src="data:image/png;base64,AQID" /></p>'

        ConvertFrom-OfficeMarkdownHtml -Html $html -OutputPath $path -Base64ImageHandling SaveToFile -Base64ImageOutputDirectory $imageDirectory -WhatIf | Out-Null

        Test-Path -LiteralPath $path | Should -BeFalse
        Test-Path -LiteralPath $imageDirectory | Should -BeFalse
    }

    It 'does not extract HTML data URI images from no-output conversions when WhatIf is used' {
        $imageDirectory = Join-Path $TestDrive 'images-no-output'
        $html = '<p><img alt="Tiny" src="data:image/png;base64,AQID" /></p>'

        ConvertFrom-OfficeMarkdownHtml -Html $html -Base64ImageHandling SaveToFile -Base64ImageOutputDirectory $imageDirectory -WhatIf | Out-Null

        Test-Path -LiteralPath $imageDirectory | Should -BeFalse
    }

    It 'builds Markdown via DSL helpers' {
        $path = Join-Path $TestDrive 'MarkdownDsl.md'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $file = New-OfficeMarkdown -Path $path {
            MarkdownHeading -Level 1 -Text 'Report'
            MarkdownParagraph -Text 'Generated by PSWriteOffice'
            MarkdownList -Items 'One','Two'
            MarkdownTable -InputObject $rows
            MarkdownCode -Language 'powershell' -Content 'Get-Process'
            MarkdownHorizontalRule
        } -PassThru

        $file | Should -BeOfType System.IO.FileInfo
        $content = Get-Content -Path $path -Raw
        $content | Should -Match '# Report'
        $content | Should -Match 'Generated by PSWriteOffice'
        $content | Should -Match 'Name'
        $content | Should -Match 'Value'
    }

    It 'creates an empty Markdown file when no DSL content is supplied' {
        $path = Join-Path $TestDrive 'EmptyMarkdown.md'

        $file = New-OfficeMarkdown -Path $path -PassThru

        $file | Should -BeOfType System.IO.FileInfo
        Test-Path -LiteralPath $path | Should -BeTrue
        [System.IO.File]::ReadAllText($path) | Should -Be ''
    }

    It 'does not run the Markdown DSL when WhatIf skips saving' {
        $path = Join-Path $TestDrive 'MarkdownWhatIf.md'
        $script:MarkdownWhatIfDslRan = $false

        New-OfficeMarkdown -Path $path -WhatIf {
            $script:MarkdownWhatIfDslRan = $true
            MarkdownParagraph -Text 'Should not run'
        } | Out-Null

        $script:MarkdownWhatIfDslRan | Should -BeFalse
        Test-Path -LiteralPath $path | Should -BeFalse
    }

    It 'supports transposed Markdown tables' {
        $path = Join-Path $TestDrive 'MarkdownTransposedTable.md'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        New-OfficeMarkdown -Path $path {
            MarkdownTable -InputObject $rows -View Transpose
        } | Out-Null

        $content = Get-Content -Path $path -Raw
        $content | Should -Match 'Property'
        $content | Should -Match 'Row1'
        $content | Should -Match 'Row2'
        $content | Should -Match 'Alpha'
        $content | Should -Match 'Beta'
    }

    It 'adds Markdown tables to piped document objects' {
        $doc = Get-OfficeMarkdown -Text '# Seed'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $updated = $doc | MarkdownTable -InputObject $rows -PassThru
        $markdown = $updated.ToMarkdown()

        $updated | Should -Be $doc
        $markdown | Should -Match 'Name'
        $markdown | Should -Match 'Alpha'
    }

    It 'adds piped Markdown table rows to a supplied document' {
        $doc = Get-OfficeMarkdown -Text '# Seed'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $updated = $rows | MarkdownTable -Document $doc -PassThru
        $markdown = $updated.ToMarkdown()

        $updated | Should -Be $doc
        $markdown | Should -Match 'Alpha'
        $markdown | Should -Match 'Beta'
    }

    It 'adds explicit Markdown table rows to each piped document' {
        $doc1 = Get-OfficeMarkdown -Text '# First'
        $doc2 = Get-OfficeMarkdown -Text '# Second'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $updated = @($doc1, $doc2) | MarkdownTable -InputObject $rows -PassThru

        $updated.Count | Should -Be 2
        $updated[0] | Should -Be $doc1
        $updated[1] | Should -Be $doc2
        ($doc1.ToMarkdown() | Select-String -Pattern 'Alpha' -AllMatches).Matches.Count | Should -Be 1
        ($doc2.ToMarkdown() | Select-String -Pattern 'Alpha' -AllMatches).Matches.Count | Should -Be 1
    }

    It 'returns an in-memory Markdown document when NoSave is used' {
        $path = Join-Path $TestDrive 'NoSave.md'

        $document = New-OfficeMarkdown -Path $path -NoSave {
            MarkdownHeading -Level 1 -Text 'Draft'
        }

        $document.GetType().FullName | Should -Be 'OfficeIMO.Markdown.MarkdownDoc'
        $document.ToMarkdown() | Should -Match '# Draft'
        Test-Path $path | Should -BeFalse
    }

    It 'builds advanced Markdown blocks via DSL helpers' {
        $path = Join-Path $TestDrive 'MarkdownAdvancedDsl.md'

        New-OfficeMarkdown -Path $path {
            MarkdownFrontMatter -Data @{
                title = 'Operations Report'
                tags  = @('ops', 'weekly')
            }
            MarkdownTableOfContents -Title 'Contents' -PlaceAtTop -MinLevel 2 -MaxLevel 3
            MarkdownHeading -Level 1 -Text 'Operations Report'
            MarkdownHeading -Level 2 -Text 'Checklist'
            MarkdownTaskList -Items 'Draft', 'Review', 'Ship' -Completed 1
            MarkdownHeading -Level 2 -Text 'Glossary'
            MarkdownDefinitionList -Definition @{
                SLA = 'Service level agreement'
                SLO = 'Service level objective'
            }
            MarkdownDetails -Summary 'Implementation Notes' {
                MarkdownParagraph -Text 'Hidden by default.'
                MarkdownList -Items 'Thin wrapper', 'Powered by OfficeIMO'
            }
        } | Out-Null

        $content = Get-Content -Path $path -Raw
        $content | Should -Match 'title: "?Operations Report"?'
        $content | Should -Match '## Contents'
        $content | Should -Match '- \[ \] Draft'
        $content | Should -Match '- \[x\] Review'
        $content | Should -Match 'SLA'
        $content | Should -Match '<details>'
    }

    It 'supports advanced Markdown cmdlets against document objects' {
        $doc = Get-OfficeMarkdown -Text '# Seed'

        $doc |
            Add-OfficeMarkdownFrontMatter -InputObject @{ title = 'Doc pipeline' } -PassThru |
            Add-OfficeMarkdownHeading -Level 1 -Text 'Doc pipeline' -PassThru |
            Add-OfficeMarkdownTaskList -Items 'Alpha', 'Beta' -Completed 0 -PassThru |
            Add-OfficeMarkdownDefinitionList -Definition @{ API = 'Application programming interface' } -PassThru |
            Add-OfficeMarkdownDetails -Summary 'More' { MarkdownParagraph -Text 'Nested block' } |
            Out-Null

        $markdown = $doc.ToMarkdown()
        $markdown | Should -Match 'title: "?Doc pipeline"?'
        $markdown | Should -Match '- \[x\] Alpha'
        $markdown | Should -Match 'API'
        $markdown | Should -Match '<details>'
    }
}
