BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')

    function Test-OfficeLoadedMethod {
        param(
            [Parameter(Mandatory)]
            [string] $TypeName,

            [Parameter(Mandatory)]
            [string] $MethodName
        )

        $type = [AppDomain]::CurrentDomain.GetAssemblies() |
            ForEach-Object { $_.GetType($TypeName, $false) } |
            Where-Object { $null -ne $_ } |
            Select-Object -First 1
        if ($null -eq $type) {
            throw "Unable to find loaded type '$TypeName'."
        }

        @($type.GetMethods() | Where-Object Name -eq $MethodName).Count -gt 0
    }

    function Add-TestWordParagraphStyle {
        param(
            [Parameter(Mandatory)]
            [string] $Path,

            [Parameter(Mandatory)]
            [string] $StyleId
        )

        $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $entry = $archive.GetEntry('word/styles.xml')
            if (-not $entry) {
                throw "Zip entry 'word/styles.xml' not found in '$Path'."
            }

            $stream = $entry.Open()
            try {
                $reader = [System.IO.StreamReader]::new($stream)
                try {
                    [xml] $stylesXml = $reader.ReadToEnd()
                } finally {
                    $reader.Dispose()
                }
            } finally {
                $stream.Dispose()
            }

            $wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            $style = $stylesXml.CreateElement('w', 'style', $wordNamespace)
            $style.SetAttribute('type', $wordNamespace, 'paragraph')
            $style.SetAttribute('customStyle', $wordNamespace, '1')
            $style.SetAttribute('styleId', $wordNamespace, $StyleId)

            $name = $stylesXml.CreateElement('w', 'name', $wordNamespace)
            $name.SetAttribute('val', $wordNamespace, 'Issue 115 Style')
            $basedOn = $stylesXml.CreateElement('w', 'basedOn', $wordNamespace)
            $basedOn.SetAttribute('val', $wordNamespace, 'Normal')
            $quickFormat = $stylesXml.CreateElement('w', 'qFormat', $wordNamespace)

            $style.AppendChild($name) | Out-Null
            $style.AppendChild($basedOn) | Out-Null
            $style.AppendChild($quickFormat) | Out-Null
            $stylesXml.DocumentElement.AppendChild($style) | Out-Null

            $entry.Delete()
            $newEntry = $archive.CreateEntry('word/styles.xml')
            $writeStream = $newEntry.Open()
            try {
                $writer = [System.IO.StreamWriter]::new($writeStream, [System.Text.UTF8Encoding]::new($false))
                try {
                    $stylesXml.Save($writer)
                } finally {
                    $writer.Dispose()
                }
            } finally {
                $writeStream.Dispose()
            }
        } finally {
            $archive.Dispose()
        }
    }

    function Get-TestWordBodyOrder {
        param(
            [Parameter(Mandatory)]
            [string] $Path
        )

        $documentXml = Get-ZipXmlDocumentLocal -Path $Path -Entry 'word/document.xml'
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

        foreach ($node in $documentXml.SelectNodes('/w:document/w:body/*', $namespaceManager)) {
            if ($node.LocalName -eq 'p') {
                $text = ($node.SelectNodes('.//w:t', $namespaceManager) | ForEach-Object { $_.'#text' }) -join ''
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    "p:$text"
                }
            } elseif ($node.LocalName -eq 'tbl') {
                $text = ($node.SelectNodes('.//w:t', $namespaceManager) | ForEach-Object { $_.'#text' }) -join '|'
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    "tbl:$text"
                }
            }
        }
    }
}

Describe 'Word DSL surface' {
    It 'creates a document with canonical cmdlets' {
        $path = Join-Path $TestDrive 'DslCanonical.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordSection {
                Add-OfficeWordParagraph -Text 'Smoke test paragraph.'
                Add-OfficeWordList -Style 'Bulleted' {
                    Add-OfficeWordListItem -Text 'First'
                    Add-OfficeWordListItem -Text 'Second'
                }
            }
        }

        Test-Path $path | Should -BeTrue

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.Sections.Count | Should -BeGreaterThan 0
            $document.Paragraphs.Count | Should -BeGreaterThan 0
        } finally {
            $document.Dispose()
        }

    }

    It 'preserves authored paragraph and table order inside a single Word section' {
        $path = Join-Path $TestDrive 'DslSectionOrderedBlocks.docx'
        $firstTable = @(
            [PSCustomObject]@{ Control = 'A'; Status = 'Open' }
        )
        $secondTable = @(
            [PSCustomObject]@{ Control = 'B'; Status = 'Closed' }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordSection {
                Add-OfficeWordParagraph -Text 'Heading check'
                Add-OfficeWordParagraph -Text 'Text before table'
                Add-OfficeWordTable -InputObject $firstTable -Style TableGrid
                Add-OfficeWordParagraph -Text 'Text after table'
                Add-OfficeWordTable -InputObject $secondTable -Style TableGrid
                Add-OfficeWordParagraph -Text 'Tail paragraph'
            }
        }

        @(Get-TestWordBodyOrder -Path $path) | Should -Be @(
            'p:Heading check'
            'p:Text before table'
            'tbl:Control|Status|A|Open'
            'p:Text after table'
            'tbl:Control|Status|B|Closed'
            'p:Tail paragraph'
        )
    }

    It 'keeps top-level paragraphs in the implicit section' {
        $path = Join-Path $TestDrive 'DslImplicitSectionParagraphs.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph -Text 'Hello World'
            Add-OfficeWordParagraph -Text 'Hello Again'
        }

        @(Get-TestWordBodyOrder -Path $path) | Should -Be @(
            'p:Hello World'
            'p:Hello Again'
        )

        $documentXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'word/document.xml'
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

        $documentXml.SelectNodes('/w:document/w:body/w:p[w:pPr/w:sectPr and not(.//w:t[normalize-space()])]', $namespaceManager).Count | Should -Be 0
    }

    It 'preserves OfficeIMO insertion order when editing paragraphs returned by Find-OfficeWord' {
        $path = Join-Path $TestDrive 'ExistingWordInlineInsertion.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph -Text 'Before marker'
            Add-OfficeWordParagraph -Text 'Insertion marker'
            Add-OfficeWordParagraph -Text 'After marker'
        }

        $document = Get-OfficeWord -Path $path
        try {
            $marker = Find-OfficeWord -Document $document -Text 'Insertion marker' | Select-Object -First 1
            $marker | Should -Not -BeNullOrEmpty

            $inserted = $marker.AddParagraphAfterSelf()
            $inserted.Text = 'Inserted paragraph'

            $tableStyleType = [AppDomain]::CurrentDomain.GetAssemblies() |
                ForEach-Object { $_.GetType('OfficeIMO.Word.WordTableStyle', $false) } |
                Where-Object { $null -ne $_ } |
                Select-Object -First 1
            $tableStyleType | Should -Not -BeNullOrEmpty

            $table = $inserted.AddTableAfter(2, 2, [Enum]::Parse($tableStyleType, 'TableGrid'))
            $table.Rows[0].Cells[0].Paragraphs[0].Text = 'Name'
            $table.Rows[0].Cells[1].Paragraphs[0].Text = 'State'
            $table.Rows[1].Cells[0].Paragraphs[0].Text = 'Inserted table'
            $table.Rows[1].Cells[1].Paragraphs[0].Text = 'Ready'

            Close-OfficeWord -Document $document -Save
            $document = $null
        } finally {
            if ($null -ne $document) {
                $document.Dispose()
            }
        }

        @(Get-TestWordBodyOrder -Path $path) | Should -Be @(
            'p:Before marker'
            'p:Insertion marker'
            'p:Inserted paragraph'
            'tbl:Name|State|Inserted table|Ready'
            'p:After marker'
        )
    }

    It 'round-trips encrypted Word documents through lifecycle cmdlets' {
        if (-not (Test-OfficeLoadedMethod -TypeName 'OfficeIMO.Word.WordDocument' -MethodName 'LoadEncrypted')) {
            (Get-Command New-OfficeWord).Parameters.Keys | Should -Contain 'Password'
            (Get-Command Save-OfficeWord).Parameters.Keys | Should -Contain 'Password'
            (Get-Command Get-OfficeWord).Parameters.Keys | Should -Contain 'Password'
            return
        }

        $path = Join-Path $TestDrive 'EncryptedWord.docx'

        New-OfficeWord -Path $path -Password 'secret' {
            WordSection {
                WordParagraph -Text 'Encrypted Word value'
            }
        }

        { Get-ZipEntriesLocal -Path $path } | Should -Throw
        $document = Get-OfficeWord -Path $path -Password 'secret' -ReadOnly
        try {
            $document.Paragraphs.Text | Should -Contain 'Encrypted Word value'
        } finally {
            $document.Dispose()
        }

        $document = Get-OfficeWord -Path $path -Password 'secret'
        try {
            { $document | Save-OfficeWord -Path $path -ErrorAction Stop } |
                Should -Throw '*Provide -Password*'
            $document.AddParagraph('Updated encrypted Word value') | Out-Null
            $document | Save-OfficeWord -Password 'secret'
        } finally {
            $document | Close-OfficeWord
        }

        $plainCopyPath = Join-Path $TestDrive 'EncryptedWordPlainCopy.docx'
        $document = Get-OfficeWord -Path $path -Password 'secret'
        try {
            $document | Save-OfficeWord -Path $plainCopyPath
            $document.AddParagraph('Updated plain Word copy') | Out-Null
            $document | Save-OfficeWord
        } finally {
            $document | Close-OfficeWord
        }

        { Get-ZipEntriesLocal -Path $plainCopyPath } | Should -Not -Throw
        $plainCopy = Get-OfficeWord -Path $plainCopyPath -ReadOnly
        try {
            $plainCopy.Paragraphs.Text | Should -Contain 'Updated plain Word copy'
        } finally {
            $plainCopy | Close-OfficeWord
        }

        $encryptedCopyPath = Join-Path $TestDrive 'EncryptedWordCopy.docx'
        $document = Get-OfficeWord -Path $path -Password 'secret'
        try {
            $document | Save-OfficeWord -Path $encryptedCopyPath -Password 'copy-secret'
            $document.AddParagraph('Updated encrypted Word copy') | Out-Null
            { $document | Save-OfficeWord -ErrorAction Stop } | Should -Throw '*Provide -Password*'
            $document | Save-OfficeWord -Password 'copy-secret'
        } finally {
            $document | Close-OfficeWord
        }

        $encryptedCopy = Get-OfficeWord -Path $encryptedCopyPath -Password 'copy-secret' -ReadOnly
        try {
            $encryptedCopy.Paragraphs.Text | Should -Contain 'Updated encrypted Word copy'
        } finally {
            $encryptedCopy | Close-OfficeWord
        }

        $document = Get-OfficeWord -Path $path -Password 'secret' -ReadOnly
        try {
            $document.Paragraphs.Text | Should -Contain 'Updated encrypted Word value'
        } finally {
            $document | Close-OfficeWord
        }

        (Get-Command Get-OfficeWord).Parameters.Keys | Should -Not -Contain 'AutoSave'
    }

    It 'runs the Word DSL against a cloned template document' {
        $templatePath = Join-Path $TestDrive 'Issue115Template.docx'
        $outputPath = Join-Path $TestDrive 'Issue115Output.docx'
        $styleId = 'Issue115Style'

        New-OfficeWord -Path $templatePath {
            WordParagraph -Text 'Template fixed section'
        } | Out-Null
        Add-TestWordParagraphStyle -Path $templatePath -StyleId $styleId

        New-OfficeWord -TemplatePath $templatePath -Path $outputPath {
            WordParagraph -Text 'Generated appendix' -StyleId $styleId
        } | Out-Null

        $document = Get-OfficeWord -Path $outputPath -ReadOnly
        try {
            $document.Paragraphs.Text | Should -Contain 'Template fixed section'
            $document.Paragraphs.Text | Should -Contain 'Generated appendix'
        } finally {
            $document.Dispose()
        }

        $stylesXml = Get-ZipXmlDocumentLocal -Path $outputPath -Entry 'word/styles.xml'
        $documentXml = Get-ZipXmlDocumentLocal -Path $outputPath -Entry 'word/document.xml'
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

        $stylesXml.SelectSingleNode("//w:style[@w:styleId='$styleId']", $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode("//w:p[w:r/w:t='Generated appendix']/w:pPr/w:pStyle[@w:val='$styleId']", $namespaceManager) | Should -Not -BeNullOrEmpty
    }

    It 'runs the Word DSL against a loaded existing document' {
        $path = Join-Path $TestDrive 'Issue115Existing.docx'

        New-OfficeWord -Path $path {
            WordParagraph -Text 'Existing report body'
        } | Out-Null

        $document = Get-OfficeWord -Path $path {
            WordParagraph -Text 'Added after load'
        }
        try {
            $document | Save-OfficeWord | Out-Null
        } finally {
            Close-OfficeWord -Document $document
        }

        $saved = Get-OfficeWord -Path $path -ReadOnly
        try {
            $saved.Paragraphs.Text | Should -Contain 'Existing report body'
            $saved.Paragraphs.Text | Should -Contain 'Added after load'
        } finally {
            $saved.Dispose()
        }
    }

    It 'supports alias-style DSL with tables' {
        $path = Join-Path $TestDrive 'DslAlias.docx'
        $rows = @(
            [PSCustomObject]@{ Item = 'Alpha'; Qty = 1 }
            [PSCustomObject]@{ Item = 'Beta'; Qty = 5 }
        )

        New-OfficeWord -Path $path {
            WordTableOfContents -Style Template1
            WordSection {
                WordParagraph { WordText 'Alias smoke' }
                WordList {
                    WordListItem 'One'
                    WordListItem 'Two'
                }
                WordTable -InputObject $rows -Style 'GridTable1LightAccent1' {
                    WordTableCondition -FilterScript { $_.Qty -gt 2 } -BackgroundColor '#ffeeee'
                }
            }
        }

        Test-Path $path | Should -BeTrue
        (Get-Command WordNew).ResolvedCommandName | Should -Be 'New-OfficeWord'
        (Get-Command WordCheckBox).ResolvedCommandName | Should -Be 'Add-OfficeWordCheckBox'
        (Get-Command WordDatePicker).ResolvedCommandName | Should -Be 'Add-OfficeWordDatePicker'
        (Get-Command WordDropDownList).ResolvedCommandName | Should -Be 'Add-OfficeWordDropDownList'
        (Get-Command WordTableCondition).ResolvedCommandName | Should -Be 'Add-OfficeWordTableCondition'
        (Get-Command WordTableOfContents).ResolvedCommandName | Should -Be 'Add-OfficeWordTableOfContents'
        (Get-Command WordSection).ResolvedCommandName | Should -Be 'Add-OfficeWordSection'
    }

    It 'creates multiple continuous sections from a PowerShell-friendly break type' {
        $path = Join-Path $TestDrive 'DslContinuousSections.docx'

        New-OfficeWord -Path $path {
            1..8 | ForEach-Object {
                Add-OfficeWordSection {
                    Add-OfficeWordParagraph -Text "Section $_"
                } -BreakType Continuous
            }
        }

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.Sections.Count | Should -Be 8
        } finally {
            $document.Dispose()
        }

        $documentXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'word/document.xml'
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $continuousBreaks = $documentXml.SelectNodes("//w:sectPr/w:type[@w:val='continuous']", $namespaceManager)

        @($continuousBreaks).Count | Should -Be 7
    }

    It 'supports paragraphs, lists, images, and nested tables inside table cells' {
        $path = Join-Path $TestDrive 'DslTableCells.docx'
        $imagePath = Join-Path $TestDrive 'CellImage.png'
        $fixturePath = Join-Path $PSScriptRoot 'Assets/CellImage.png'
        $rows = @(
            [PSCustomObject]@{
                Topic   = 'Release readiness'
                Details = 'Pending'
            }
        )

        $nestedRows = @(
            [PSCustomObject]@{ Step = 'Validate'; State = 'Ready' }
        )

        Copy-Item -LiteralPath $fixturePath -Destination $imagePath -Force

        New-OfficeWord -Path $path {
            WordTable -InputObject $rows -Style 'GridTable1LightAccent1' {
                WordTableCell -Row 1 -Column 0 {
                    WordParagraph { WordText 'Checklist' }
                    WordImage -Path $imagePath -Width 24 -Height 24
                    WordList {
                        WordListItem 'Confirm issue coverage'
                        WordListItem 'Stage release notes'
                    }
                }

                WordTableCell -Row 1 -Column 1 {
                    WordTable -InputObject $nestedRows -NoHeader -Style 'TableGrid' {
                        WordTableCell -Row 0 -Column 0 {
                            WordParagraph { WordText 'Nested detail' }
                        }
                    }
                }
            }
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $cellTexts = $table.Rows[1].Cells[0].Paragraphs |
                ForEach-Object Text |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $listTexts = $document.Lists |
                ForEach-Object ListItems |
                ForEach-Object Text |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $nestedCellTexts = $table.NestedTables[0].Rows[0].Cells[0].Paragraphs |
                ForEach-Object Text |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $outerDetailTexts = $table.Rows[1].Cells[1].Paragraphs |
                ForEach-Object Text |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

            $cellTexts | Should -Contain 'Checklist'
            $cellTexts | Should -Contain 'Confirm issue coverage'
            $listTexts | Should -Contain 'Confirm issue coverage'
            $listTexts | Should -Contain 'Stage release notes'
            ($table.Rows[1].Cells[0].Paragraphs | Where-Object IsImage).Count | Should -Be 1

            $table.HasNestedTables | Should -BeTrue
            $table.NestedTables.Count | Should -Be 1
            $nestedCellTexts | Should -Contain 'Nested detail'
            $outerDetailTexts | Should -Not -Contain 'Nested detail'
            $table.NestedTables[0].Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'Ready'
        } finally {
            $document.Dispose()
        }
    }

    It 'adds same-paragraph Word breaks' {
        $path = Join-Path $TestDrive 'DslWordBreaks.docx'

        New-OfficeWord -Path $path {
            WordParagraph {
                WordText 'Line 1'
                WordBreak
                WordText 'Line 2'
            }
        } | Out-Null

        $documentXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'word/document.xml'
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

        $documentXml.SelectNodes('//w:br', $namespaceManager).Count | Should -Be 1
        $documentXml.SelectSingleNode('//w:t[text()="Line 1"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:t[text()="Line 2"]', $namespaceManager) | Should -Not -BeNullOrEmpty
    }

    It 'supports reader helpers and save' {
        $path = Join-Path $TestDrive 'DslReaders.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'One'; Value = 1 }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordSection {
                Add-OfficeWordParagraph -Text 'Reader smoke.'
                Add-OfficeWordTable -InputObject $rows -Style 'TableGrid'
            }
        } | Out-Null

        $document = Get-OfficeWord -Path $path
        try {
            ($document | Get-OfficeWordSection).Count | Should -BeGreaterThan 0
            ($document | Get-OfficeWordParagraph).Count | Should -BeGreaterThan 0
            ($document | Get-OfficeWordTable).Count | Should -BeGreaterThan 0
            ($document | Get-OfficeWordParagraph | Select-Object -First 1 | Get-OfficeWordText).Count | Should -BeGreaterThan 0

            $document | Save-OfficeWord | Out-Null
        } finally {
            Close-OfficeWord -Document $document
        }
    }

    It 'supports explicit Word table layout modes' {
        $path = Join-Path $TestDrive 'DslTableLayouts.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'One'; Value = 1 }
            [PSCustomObject]@{ Name = 'Two'; Value = 2 }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordTable -InputObject $rows -Style TableGrid -Layout AutoFitToContents
            Add-OfficeWordTable -InputObject $rows -Style TableGrid -Layout AutoFitToWindow
            Add-OfficeWordTable -InputObject $rows -Style TableGrid -Layout Fixed
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.Tables.Count | Should -Be 3
            $document.Tables[0].LayoutMode.ToString() | Should -Be 'AutoFit'
            $document.Tables[0].WidthType.ToString() | Should -Be 'Auto'
            $document.Tables[0].Width | Should -Be 0
            $document.Tables[1].LayoutMode.ToString() | Should -Be 'Fixed'
            $document.Tables[1].WidthType.ToString() | Should -Be 'Pct'
            $document.Tables[1].Width | Should -Be 5000
            $document.Tables[2].LayoutMode.ToString() | Should -Be 'Fixed'
        } finally {
            $document.Dispose()
        }
    }

    It 'creates span-aware Word tables from Word table cell specs' {
        (Get-Command New-OfficeWordTableCell).Parameters.Keys | Should -Contain 'ColumnSpan'
        (Get-Command New-OfficeWordTableCell).Parameters.Keys | Should -Contain 'RowSpan'
        Get-Command New-OfficeTableCell -ErrorAction SilentlyContinue | Should -BeNullOrEmpty

        $path = Join-Path $TestDrive 'DslSpanAwareWordTable.docx'

        New-OfficeWord -Path $path {
            WordTable -Style TableGrid -InputObject @(
                @('Service', 'Status', 'Owner'),
                @(New-OfficeWordTableCell -Text 'Identity systems' -ColumnSpan 3),
                @('Entra', 'Watch', 'IAM'),
                @((New-OfficeWordTableCell -Text 'Shared owner' -RowSpan 2), 'Build', 'OfficeIMO'),
                @('Release', 'PSWriteOffice')
            )
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.Rows[0].CellsCount | Should -Be 3
            $table.Rows[1].CellsCount | Should -Be 1
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Identity systems'
            $table.Rows[1].Cells[0].ColumnSpan | Should -Be 3
            $table.Rows[3].Cells[0].Paragraphs[0].Text | Should -Be 'Shared owner'
            $table.Rows[3].Cells[0].RowSpan | Should -Be 2
            $table.Rows[4].Cells[1].Paragraphs[0].Text | Should -Be 'Release'
        } finally {
            $document.Dispose()
        }
    }

    It 'creates Word paragraphs and table cells from rich text runs with named colors' {
        foreach ($name in 'WordNew', 'WordTextRun', 'WordTableCellSpec') {
            Get-Command $name | Should -Not -BeNullOrEmpty
        }

        $path = Join-Path $TestDrive 'DslRichTextRunsWord.docx'

        WordNew -Path $path {
            WordParagraph -Run @{
                Text      = 'Status: ', 'Ready'
                Color     = $null, 'SeaGreen'
                Bold      = $false, $true
                Underline = 'None', 'Dotted'
                FontSize  = 12
            }
            WordParagraph {
                WordText -Text 'Same ', 'formatting' -Color Navy -FontFamily Arial -FontSize 10
            }
            WordParagraph -Run @(
                WordTextRun 'Highlight ' -BackgroundColor Yellow
                WordTextRun 'Shaded' -BackgroundColor LightPink
                WordTextRun ' Plain' -BackgroundColor None
            )
            WordParagraph -Run @(
                WordTextRun 'x'
                WordTextRun '2' -Kind Superscript
                WordTextRun ' H'
                WordTextRun '2' -Baseline Subscript
                WordTextRun 'O'
            )
            WordParagraph -Run @(
                WordTextRun 'Styled link' -LinkUri 'https://example.org/styled' -LinkContents 'Styled tooltip' -Bold -Color Crimson -BackgroundColor Yellow -FontSize 16 -FontName 'Arial'
            )
            WordTable -Style TableGrid -InputObject @(
                , @(
                    WordTableCellSpec -Run @(
                        WordTextRun 'Build '
                        WordTextRun 'Ready' -Color SeaGreen -Bold
                    ) -ColumnSpan 2 -FillColor AliceBlue -Bold -TextColor Red -FontSize 14
                )
                , @('Owner', 'Platform')
            ) {
                WordTableCell -Row 1 -Column 0 -Run @(
                    WordTextRun 'Owner: '
                    WordTextRun 'Platform' -Color Navy -Bold
                )
            }
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.Rows[0].Cells[0].ColumnSpan | Should -Be 2
        } finally {
            $document.Dispose()
        }

        $documentXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'word/document.xml'
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $text = ($documentXml.GetElementsByTagName('t', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') | ForEach-Object { $_.InnerText }) -join ''
        $text | Should -Match 'Status: Ready'
        $text | Should -Match 'Same formatting'
        $text | Should -Match 'Highlight Shaded Plain'
        $text | Should -Match 'x2 H2O'
        $text | Should -Match 'Styled link'
        $text | Should -Match 'Build Ready'
        $text | Should -Match 'Owner: Platform'
        $documentXml.SelectSingleNode('//w:color[translate(@w:val, "abcdef", "ABCDEF")="2E8B57"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:hyperlink[.//w:t="Styled link"]//w:b', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:hyperlink[.//w:t="Styled link"]//w:color[translate(@w:val, "abcdef", "ABCDEF")="DC143C"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:hyperlink[.//w:t="Styled link"]//w:highlight[@w:val="yellow"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:hyperlink[.//w:t="Styled link"]//w:sz[@w:val="32"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:r[w:t="Build "]//w:b', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:r[w:t="Build "]//w:color[translate(@w:val, "abcdef", "ABCDEF")="FF0000"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:r[w:t="Build "]//w:sz[@w:val="28"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:r[w:t="Ready"]//w:color[translate(@w:val, "abcdef", "ABCDEF")="2E8B57"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:vertAlign[@w:val="superscript"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:vertAlign[@w:val="subscript"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:highlight[@w:val="yellow"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:shd[translate(@w:fill, "abcdef", "ABCDEF")="FFB6C1"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:highlight[translate(@w:val, "abcdef", "ABCDEF")="F0F8FF"] | //w:shd[translate(@w:fill, "abcdef", "ABCDEF")="F0F8FF"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:u[@w:val="dotted"]', $namespaceManager) | Should -Not -BeNullOrEmpty
    }

    It 'preserves half-point Word font sizes through text, rich runs, links, and style updates' {
        $path = Join-Path $TestDrive 'DslHalfPointWordFonts.docx'
        WordNew -Path $path {
            WordParagraph { WordText -Text 'DirectHalf' -FontSize 10.5 }
            WordParagraph -Run @(
                WordTextRun 'RichHalf' -FontSize 10.5
                WordTextRun 'LinkHalf' -LinkUri 'https://example.org/half' -FontSize 10.5
            )
            $styled = WordParagraph -Text 'StyledHalf' -PassThru
            @($styled.GetRuns())[0] | Set-OfficeWordTextStyle -FontSize 10.5
        } | Out-Null

        $documentXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'word/document.xml'
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        foreach ($text in 'DirectHalf', 'RichHalf', 'LinkHalf', 'StyledHalf') {
            $documentXml.SelectSingleNode("//w:r[w:t='$text']/w:rPr/w:sz[@w:val='21']", $namespaceManager) |
                Should -Not -BeNullOrEmpty
        }
    }

    It 'rejects mismatched columnar Word rich text formatting arrays' {
        $path = Join-Path $TestDrive 'DslRichTextRunsWordMismatch.docx'

        {
            WordNew -Path $path {
                WordParagraph -Run @{
                    Text  = 'Red ', 'Green '
                    Color = 'Red', 'Green', 'Blue'
                }
            }
        } | Should -Throw "*property 'Color'*match the Text count (2)*received 3 values*"

        {
            WordNew -Path $path {
                WordParagraph -Run @{
                    Text  = 'Red ', 'Green '
                    Color = @()
                }
            }
        } | Should -Throw "*property 'Color'*match the Text count (2)*received 0 values*"
    }

    It 'treats rich and styled Word table cells as structured without spans' {
        $path = Join-Path $TestDrive 'DslRichRunMixedObjectWordTable.docx'

        WordNew -Path $path {
            WordTable -Style TableGrid -InputObject @(
                [pscustomobject]@{
                    Service = 'Entra'
                    Status  = 'Ready'
                }
                , @(
                    (WordTableCellSpec -Run @(
                        WordTextRun 'Build '
                        WordTextRun 'Watch' -Color DarkOrange -Bold
                    ) -FillColor AliceBlue),
                    'Platform'
                )
            )
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Service'
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Entra'
            $table.Rows[2].Cells[1].Paragraphs[0].Text | Should -Be 'Platform'
        } finally {
            $document.Dispose()
        }

        $documentXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'word/document.xml'
        $text = ($documentXml.GetElementsByTagName('t', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') | ForEach-Object { $_.InnerText }) -join ''
        $text | Should -Match 'Build Watch'
    }

    It 'keeps ordinary Run columns in normal Word tables' {
        $path = Join-Path $TestDrive 'DslOrdinaryRunColumnWordTable.docx'

        WordNew -Path $path {
            WordTable -Style TableGrid -InputObject @(
                [pscustomobject]@{
                    Run = 'Nightly'
                }
            )
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Run'
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Nightly'
        } finally {
            $document.Dispose()
        }
    }

    It 'keeps ordinary span-like property names on normal Word tables' {
        $path = Join-Path $TestDrive 'DslOrdinarySpanNamedWordTable.docx'
        $rows = @(
            [pscustomobject]@{
                Name = 'Backlog'
                Rows = 25
                Columns = 3
                Span = 2
                ColumnSpan = 2
                RowSpan = 3
            }
        )

        New-OfficeWord -Path $path {
            WordTable -Style TableGrid -InputObject $rows
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.RowsCount | Should -Be 2
            $table.Rows[0].CellsCount | Should -Be 6
            $table.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'Rows'
            $table.Rows[0].Cells[2].Paragraphs[0].Text | Should -Be 'Columns'
            $table.Rows[0].Cells[3].Paragraphs[0].Text | Should -Be 'Span'
            $table.Rows[0].Cells[4].Paragraphs[0].Text | Should -Be 'ColumnSpan'
            $table.Rows[0].Cells[5].Paragraphs[0].Text | Should -Be 'RowSpan'
            $table.Rows[1].Cells[1].Paragraphs[0].Text | Should -Be '25'
            $table.Rows[1].Cells[2].Paragraphs[0].Text | Should -Be '3'
            $table.Rows[1].Cells[3].Paragraphs[0].Text | Should -Be '2'
            $table.Rows[1].Cells[4].Paragraphs[0].Text | Should -Be '2'
            $table.Rows[1].Cells[5].Paragraphs[0].Text | Should -Be '3'
        } finally {
            $document.Dispose()
        }
    }

    It 'keeps ordinary text and span-key properties in mixed Word tables' {
        $path = Join-Path $TestDrive 'DslMixedOrdinarySpanKeyWordTable.docx'
        $rows = @(
            [pscustomobject]@{
                Text = 'Task'
                FontSize = 'Large'
                ColumnSpan = 1
                Status = 'Open'
            }
            , @(New-OfficeWordTableCell -Text 'Follow-up' -ColumnSpan 3)
        )

        New-OfficeWord -Path $path {
            WordTable -Style TableGrid -InputObject $rows
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.RowsCount | Should -Be 3
            $table.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Text'
            $table.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'FontSize'
            $table.Rows[0].Cells[2].Paragraphs[0].Text | Should -Be 'ColumnSpan'
            $table.Rows[0].Cells[3].Paragraphs[0].Text | Should -Be 'Status'
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Task'
            $table.Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'Large'
            $table.Rows[1].Cells[2].Paragraphs[0].Text | Should -Be '1'
            $table.Rows[1].Cells[3].Paragraphs[0].Text | Should -Be 'Open'
            $table.Rows[2].Cells[0].Paragraphs[0].Text | Should -Be 'Follow-up'
            $table.Rows[2].Cells[0].ColumnSpan | Should -Be 3
        } finally {
            $document.Dispose()
        }
    }

    It 'keeps default headers on mixed object and span Word tables' {
        $path = Join-Path $TestDrive 'DslMixedObjectSpanWordTable.docx'
        $rows = @(
            [pscustomobject]@{
                Name = 'Directory'
                Status = 'Healthy'
            }
            , @(New-OfficeWordTableCell -Text 'Follow-up' -ColumnSpan 2)
        )

        New-OfficeWord -Path $path {
            WordTable -Style TableGrid -InputObject $rows
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.RowsCount | Should -Be 3
            $table.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Name'
            $table.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'Status'
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Directory'
            $table.Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'Healthy'
            $table.Rows[2].Cells[0].Paragraphs[0].Text | Should -Be 'Follow-up'
            $table.Rows[2].Cells[0].ColumnSpan | Should -Be 2
        } finally {
            $document.Dispose()
        }
    }

    It 'applies conditions to leading span rows after generated mixed Word headers' {
        $path = Join-Path $TestDrive 'DslMixedLeadingSpanConditionWordTable.docx'
        $rows = @(
            , @(New-OfficeWordTableCell -Text 'Section' -ColumnSpan 2)
            [pscustomobject]@{
                Name = 'Directory'
                Status = 'Healthy'
            }
        )

        New-OfficeWord -Path $path {
            WordTable -Style TableGrid -InputObject $rows {
                WordTableCondition -FilterScript { $_[0].Text -eq 'Section' } -BackgroundColor '#ffeeee'
            }
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.RowsCount | Should -Be 3
            $table.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Name'
            $table.Rows[0].Cells[0].ShadingFillColorHex | Should -Not -Be 'ffeeee'
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Section'
            $table.Rows[1].Cells[0].ShadingFillColorHex | Should -Be 'ffeeee'
            $table.Rows[2].Cells[0].Paragraphs[0].Text | Should -Be 'Directory'
        } finally {
            $document.Dispose()
        }
    }

    It 'keeps generated mixed Word headers out of leading row spans' {
        $path = Join-Path $TestDrive 'DslMixedLeadingRowSpanWordTable.docx'
        $rows = @(
            , @((New-OfficeWordTableCell -Text 'Group' -RowSpan 2), 'A')
            [pscustomobject]@{
                Name = 'B'
                Value = 'C'
            }
        )

        New-OfficeWord -Path $path {
            WordTable -Style TableGrid -InputObject $rows
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.RowsCount | Should -Be 3
            $table.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Name'
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Group'
            $table.Rows[1].Cells[0].RowSpan | Should -Be 2
            $table.Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'A'
            $table.Rows[2].Cells[1].Paragraphs[0].Text | Should -Be 'B'
            $table.Rows[2].Cells[2].Paragraphs[0].Text | Should -Be 'C'
        } finally {
            $document.Dispose()
        }
    }

    It 'merges adjacent and occupied-range Word spans without stale indexes' {
        $path = Join-Path $TestDrive 'DslComplexSpanWordTable.docx'

        New-OfficeWord -Path $path {
            WordTable -Style TableGrid -InputObject @(
                , @(
                    (New-OfficeWordTableCell -Text 'Left block' -ColumnSpan 2),
                    (New-OfficeWordTableCell -Text 'Right block' -ColumnSpan 2)
                )
                , @(
                    'A',
                    (New-OfficeWordTableCell -Text 'Pinned' -RowSpan 2),
                    'C'
                )
                , @(
                    (New-OfficeWordTableCell -Text 'Wide task' -ColumnSpan 2),
                    'Tail'
                )
            )
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $texts = @($table.Rows | ForEach-Object {
                $_.Cells | ForEach-Object { $_.Paragraphs[0].Text }
            })
            $texts | Should -Contain 'Left block'
            $texts | Should -Contain 'Right block'
            $texts | Should -Contain 'Pinned'
            $texts | Should -Contain 'Wide task'
            $texts | Should -Contain 'Tail'
        } finally {
            $document.Dispose()
        }
    }

    It 'merges upper row spans before lower horizontal Word spans' {
        $path = Join-Path $TestDrive 'DslUpperRowSpanBeforeLowerHorizontalWordTable.docx'

        New-OfficeWord -Path $path {
            WordTable -Style TableGrid -InputObject @(
                , @(
                    'Lead',
                    'Middle',
                    (New-OfficeWordTableCell -Text 'Pinned' -RowSpan 2)
                )
                , @(
                    (New-OfficeWordTableCell -Text 'Wide task' -ColumnSpan 2),
                    'Tail'
                )
            )
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.Rows[0].Cells[2].Paragraphs[0].Text | Should -Be 'Pinned'
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Wide task'
            $table.Rows[1].Cells[0].ColumnSpan | Should -Be 2

            $texts = @($table.Rows | ForEach-Object {
                $_.Cells | ForEach-Object { $_.Paragraphs[0].Text }
            })
            $texts | Should -Contain 'Tail'
        } finally {
            $document.Dispose()
        }

        $documentXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'word/document.xml'
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

        $documentXml.SelectSingleNode('(//w:tbl)[1]/w:tr[1]/w:tc[3]/w:tcPr/w:vMerge[@w:val="restart"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('(//w:tbl)[1]/w:tr[2]/w:tc[1]/w:tcPr/w:vMerge', $namespaceManager) | Should -BeNullOrEmpty
        $documentXml.SelectSingleNode('(//w:tbl)[1]/w:tr[2]/w:tc[2]/w:tcPr/w:vMerge[@w:val="continue"]', $namespaceManager) | Should -Not -BeNullOrEmpty
    }

    It 'rejects Word row spans past the table bottom' {
        $path = Join-Path $TestDrive 'DslInvalidRowSpanWordTable.docx'

        {
            New-OfficeWord -Path $path {
                WordTable -Style TableGrid -InputObject @(
                    , @((New-OfficeWordTableCell -Text 'Too tall' -RowSpan 2), 'Tail')
                )
            } | Out-Null
        } | Should -Throw -ExpectedMessage '*Row span cannot extend past the last table row*'
    }

    It 'applies conditions to the correct explicit Word table rows' {
        $path = Join-Path $TestDrive 'DslExplicitRowConditionWordTable.docx'

        New-OfficeWord -Path $path {
            WordTable -Style TableGrid -InputObject @(
                , @((New-OfficeWordTableCell -Text 'Open'), 'Identity')
                , @((New-OfficeWordTableCell -Text 'Closed'), 'Archive')
            ) {
                WordTableCondition -FilterScript { $_[0].Text -eq 'Open' } -BackgroundColor '#ffeeee'
            }
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.Rows[0].Cells[0].ShadingFillColorHex | Should -Be 'ffeeee'
            $table.Rows[1].Cells[0].ShadingFillColorHex | Should -Not -Be 'ffeeee'
        } finally {
            $document.Dispose()
        }
    }

    It 'adds Word pie charts inside the DSL' {
        $path = Join-Path $TestDrive 'DslWordPieChart.docx'
        $rows = @(
            [PSCustomObject]@{ Region = 'North America'; Revenue = 125000 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 98000 }
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 143000 }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph -Text 'Revenue mix'
            Add-OfficeWordChart -Type Pie -Data $rows -CategoryProperty Region -SeriesProperty Revenue -Title 'Regional Revenue Mix'
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.Charts.Count | Should -Be 1
            $document.Paragraphs.Where({ $_.IsChart }).Count | Should -Be 1
        } finally {
            $document.Dispose()
        }
    }

    It 'supports transposed Word tables' {
        $path = Join-Path $TestDrive 'DslTableTranspose.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'One'; Value = 1 }
            [PSCustomObject]@{ Name = 'Two'; Value = 2 }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordTable -InputObject $rows -View Transpose -Style TableGrid
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.RowsCount | Should -Be 3
            $table.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Property'
            $table.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'Row1'
            $table.Rows[0].Cells[2].Paragraphs[0].Text | Should -Be 'Row2'
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Name'
            $table.Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'One'
            $table.Rows[1].Cells[2].Paragraphs[0].Text | Should -Be 'Two'
            $table.Rows[2].Cells[0].Paragraphs[0].Text | Should -Be 'Value'
            $table.Rows[2].Cells[1].Paragraphs[0].Text | Should -Be '1'
            $table.Rows[2].Cells[2].Paragraphs[0].Text | Should -Be '2'
        } finally {
            $document.Dispose()
        }
    }

    It 'preserves the legacy Word table transpose switch' {
        $path = Join-Path $TestDrive 'DslTableLegacyTranspose.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'One'; Value = 1 }
            [PSCustomObject]@{ Name = 'Two'; Value = 2 }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordTable -Data $rows -SkipHeader -Transpose -Style TableGrid
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.RowsCount | Should -Be 2
            $table.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Name'
            $table.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'One'
            $table.Rows[0].Cells[2].Paragraphs[0].Text | Should -Be 'Two'
        } finally {
            $document.Dispose()
        }
    }

    It 'adds Word line charts to an open document' {
        $path = Join-Path $TestDrive 'DslWordLineChart.docx'
        $rows = @(
            [PSCustomObject]@{ Month = 'Jan'; Sales = 10; Profit = 4 }
            [PSCustomObject]@{ Month = 'Feb'; Sales = 12; Profit = 5 }
            [PSCustomObject]@{ Month = 'Mar'; Sales = 15; Profit = 7 }
        )

        $document = New-OfficeWord -Path $path -NoSave
        try {
            $chart = Add-OfficeWordChart -Document $document -Type Line -InputObject $rows -CategoryProperty Month -SeriesProperty Sales, Profit -Legend -XAxisTitle 'Month' -YAxisTitle 'Value' -Title 'Monthly Trend' -PassThru
            $chart.Title | Should -Be 'Monthly Trend'
            Save-OfficeWord -Document $document | Out-Null
        } finally {
            Close-OfficeWord -Document $document
        }

        $saved = Get-OfficeWord -Path $path -ReadOnly
        try {
            $saved.Charts.Count | Should -Be 1
            $saved.Paragraphs.Where({ $_.IsChart }).Count | Should -Be 1
        } finally {
            $saved.Dispose()
        }
    }

    It 'wraps OfficeIMO Word page setup, cover pages, equations, tab stops, and statistics' {
        $path = Join-Path $TestDrive 'DslWordRoadmap.docx'
        $omml = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x+1</m:t></m:r></m:oMath>'

        New-OfficeWord -Path $path {
            Add-OfficeWordCoverPage -Template Element -Abstract 'Executive summary' -CompanyEmail 'reports@example.test'
            Add-OfficeWordSection {
                Set-OfficeWordPageSetup -PageSize A4 -Orientation Landscape -Margin Narrow -Columns 2 -ColumnSpacing 720 -ColumnSeparator $true
                Add-OfficeWordParagraph {
                    Add-OfficeWordTabStop -Position 4320 -Alignment Decimal -Leader Dot
                    Add-OfficeWordText -Text "Score`t98.5"
                }
                Add-OfficeWordEquation -Omml $omml
            }
        }

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.CoverPage | Should -Not -BeNullOrEmpty
            $document.CoverPageProperties.Abstract | Should -Be 'Executive summary'
            $document.CoverPageProperties.CompanyEmail | Should -Be 'reports@example.test'
            $document.Sections[0].PageSettings.PageSize.ToString() | Should -Be 'A4'
            $document.Sections[0].PageOrientation.ToString() | Should -Be 'Landscape'
            $document.Sections[0].Margins.Type.ToString() | Should -Be 'Narrow'
            $document.Sections[0].ColumnCount | Should -Be 2
            $document.Sections[0].ColumnsSpace | Should -Be 720
            $document.Sections[0].HasColumnSeparator | Should -BeTrue
            ($document.Paragraphs | Where-Object { $_.TabStops.Count -gt 0 }).Count | Should -BeGreaterThan 0
            $document.Equations.Count | Should -Be 1
        } finally {
            $document.Dispose()
        }

        $stats = Get-OfficeWordStatistics -Path $path
        $stats.Paragraphs | Should -BeGreaterThan 0
        $stats.Words | Should -BeGreaterThan 0
    }

    It 'joins Word documents through the thin append wrapper' {
        $basePath = Join-Path $TestDrive 'JoinBase.docx'
        $appendPath = Join-Path $TestDrive 'JoinAppendix.docx'
        $outputPath = Join-Path $TestDrive 'JoinMerged.docx'

        New-OfficeWord -Path $basePath {
            Add-OfficeWordParagraph -Text 'Base report'
        }
        New-OfficeWord -Path $appendPath {
            Add-OfficeWordParagraph -Text 'Appendix detail'
        }

        Join-OfficeWordDocument -Path $basePath -AppendPath $appendPath -OutputPath $outputPath

        $document = Get-OfficeWord -Path $outputPath -ReadOnly
        try {
            ($document.Paragraphs.Text -join "`n") | Should -Match 'Base report'
            ($document.Paragraphs.Text -join "`n") | Should -Match 'Appendix detail'
        } finally {
            $document.Dispose()
        }
    }

    It 'adds Word charts to paragraphs created inside table cells' {
        $path = Join-Path $TestDrive 'DslWordChartInCell.docx'
        $tableRows = @(
            [PSCustomObject]@{ Topic = 'Regional revenue'; Details = 'Pending' }
        )
        $chartRows = @(
            [PSCustomObject]@{ Region = 'North America'; Revenue = 125000 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 98000 }
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 143000 }
        )

        New-OfficeWord -Path $path {
            $table = Add-OfficeWordTable -InputObject $tableRows -Style 'GridTable1LightAccent1' -PassThru
            $paragraph = $table.Rows[1].Cells[1].AddParagraph()

            Add-OfficeWordChart -Paragraph $paragraph -Type Pie -InputObject $chartRows -CategoryProperty Region -SeriesProperty Revenue -Title 'Regional Revenue Mix'
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $cellCharts = $table.Rows[1].Cells[1].Paragraphs.Where({ $_.IsChart })

            $cellCharts.Count | Should -Be 1
            $cellCharts[0].Chart | Should -Not -BeNullOrEmpty
        } finally {
            $document.Dispose()
        }
    }

    It 'ignores null rows when exporting Word tables' {
        $path = Join-Path $TestDrive 'DslTableNullRows.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'One'; Value = 1 }
            $null
            [PSCustomObject]@{ Name = 'Two'; Value = 2 }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordTable -InputObject $rows -Style TableGrid
        } | Out-Null

        Test-Path $path | Should -BeTrue

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.Tables.Count | Should -Be 1
            $table = $document.Tables[0]
            $table.RowsCount | Should -Be 3
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'One'
            $table.Rows[2].Cells[0].Paragraphs[0].Text | Should -Be 'Two'
        } finally {
            $document.Dispose()
        }
    }

    It 'adds fields, watermarks, and protection' {
        $path = Join-Path $TestDrive 'DslProtected.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph {
                Add-OfficeWordText -Text 'Page '
                Add-OfficeWordField -Type Page
            }
            Add-OfficeWordWatermark -Text 'CONFIDENTIAL'
            Protect-OfficeWordDocument -Password 'secret'
        } | Out-Null

        Test-Path $path | Should -BeTrue
    }

    It 'adds content controls and table of contents' {
        $path = Join-Path $TestDrive 'DslContentControls.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph { Add-OfficeWordContentControl -Text 'Client' -Alias 'ClientName' }
            Add-OfficeWordParagraph { Add-OfficeWordCheckBox -Checked -Alias 'Approved' }
            Add-OfficeWordParagraph { Add-OfficeWordDatePicker -Date (Get-Date) -Alias 'DueDate' }
            Add-OfficeWordParagraph { Add-OfficeWordDropDownList -Items 'Low','Medium','High' -Alias 'Priority' }
            Add-OfficeWordParagraph { Add-OfficeWordComboBox -Items 'Red','Blue' -DefaultValue 'Blue' -Alias 'Color' }
            Add-OfficeWordParagraph { Add-OfficeWordRepeatingSection -SectionTitle 'Items' -Alias 'LineItems' }
            Add-OfficeWordTableOfContents -Style Template1
            Update-OfficeWordFields
        } | Out-Null

        Test-Path $path | Should -BeTrue

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.StructuredDocumentTags.Count | Should -BeGreaterThan 0
            $document.CheckBoxes.Count | Should -BeGreaterThan 0
            $document.DatePickers.Count | Should -BeGreaterThan 0
            $document.DropDownLists.Count | Should -BeGreaterThan 0
            $document.ComboBoxes.Count | Should -BeGreaterThan 0
            $document.RepeatingSections.Count | Should -BeGreaterThan 0
            $document.TableOfContent | Should -Not -BeNullOrEmpty
        } finally {
            $document.Dispose()
        }
    }

    It 'updates piped Word table of contents outside the DSL' {
        $path = Join-Path $TestDrive 'DslPipedTableOfContentsUpdate.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordTableOfContents
            Add-OfficeWordParagraph -Text 'Scope' -Style Heading1
        } | Out-Null

        $document = Get-OfficeWord -Path $path
        try {
            $toc = $document | Get-OfficeWordTableOfContents
            $toc | Update-OfficeWordTableOfContents -PassThru | Should -Not -BeNullOrEmpty
        } finally {
            $document.Dispose()
        }
    }

    It 'adds and reads footnotes and endnotes in the DSL' {
        $path = Join-Path $TestDrive 'DslNotes.docx'

        New-OfficeWord -Path $path {
            WordParagraph {
                WordText 'Availability is calculated from successful health probes'
                WordFootnote -Text 'Probe data excludes planned maintenance windows.'
            }

            WordParagraph {
                WordText 'The appendix keeps the full scoring details'
                WordEndnote -Text 'Scoring uses freshness, severity, and service ownership weights.'
            }
        } | Out-Null

        $footnotes = @(Get-OfficeWordFootnote -Path $path)
        $endnotes = @(Get-OfficeWordEndnote -Path $path)

        $footnotes.Count | Should -Be 1
        $endnotes.Count | Should -Be 1
        $footnotes[0].Text | Should -Match 'Probe data excludes planned maintenance windows'
        $endnotes[0].Text | Should -Match 'Scoring uses freshness'

        $entries = @(Get-ZipEntriesLocal -Path $path)
        $entries | Should -Contain 'word/footnotes.xml'
        $entries | Should -Contain 'word/endnotes.xml'
    }

    It 'supports hyperlinks and document properties in the DSL' {
        $path = Join-Path $TestDrive 'DslLinksAndProperties.docx'

        New-OfficeWord -Path $path {
            Set-OfficeWordDocumentProperty -Name Title -Value 'DSL document'
            Set-OfficeWordDocumentProperty -Name Creator -Value 'PSWriteOffice'
            Set-OfficeWordDocumentProperty -Name BuildNumber -Value 21 -Custom

            WordParagraph {
                WordText 'Open '
                WordHyperlink -Text 'Example' -Url 'https://example.org' -Styled -Tooltip 'External link'
                WordText ' or jump to '
                WordHyperlink -Text 'Summary' -Anchor 'Summary'
            }

            WordParagraph {
                WordText 'Summary'
                WordBookmark -Name 'Summary'
            }

            WordParagraph -Run @(
                WordTextRun 'Run link' -LinkUri 'https://example.org/run' -LinkContents 'Run tooltip'
                WordTextRun ' and '
                WordTextRun 'run anchor' -LinkDestinationName 'Summary'
            )
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.HyperLinks.Count | Should -Be 4
            $document.BuiltinDocumentProperties.Title | Should -Be 'DSL document'
            $document.BuiltinDocumentProperties.Creator | Should -Be 'PSWriteOffice'
            $document.CustomDocumentProperties['BuildNumber'].Value | Should -Be 21
        } finally {
            $document.Dispose()
        }

        $links = @(Get-OfficeWordHyperlink -Path $path)
        $relationshipXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'word/_rels/document.xml.rels'
        $hyperlinkTargets = @($relationshipXml.Relationships.Relationship |
                Where-Object Type -eq 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' |
                ForEach-Object Target)
        $hyperlinkTargets | Should -Contain 'https://example.org/run'
        ($links | Where-Object Text -eq 'Run link' | Select-Object -First 1).Tooltip | Should -Be 'Run tooltip'
        ($links | Where-Object Text -eq 'run anchor' | Select-Object -First 1).Anchor | Should -Be 'Summary'
    }

    It 'supports background colors and mail merge in the DSL' {
        $path = Join-Path $TestDrive 'DslBackgroundMailMerge.docx'

        New-OfficeWord -Path $path {
            Set-OfficeWordBackground -Color '#ddeeff'

            Add-OfficeWordParagraph {
                Add-OfficeWordText -Text 'Dear '
                Add-OfficeWordField -Type MergeField -Parameters '"FirstName"'
                Add-OfficeWordText -Text ','
            }

            Add-OfficeWordParagraph {
                Add-OfficeWordText -Text 'Order '
                Add-OfficeWordField -Type MergeField -Parameters '"OrderId"'
                Add-OfficeWordText -Text ' is ready.'
            }

            Invoke-OfficeWordMailMerge -Values @{
                FirstName = 'Jane'
                OrderId   = 77
            }
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.Background.Color | Should -Be 'ddeeff'
            $mergeFieldType = Get-TestPSWriteOfficeEnumValue -AssemblyName 'OfficeIMO.Word' -TypeName 'OfficeIMO.Word.WordFieldType' -Name 'MergeField'
            $document.Fields.Where({ $_.FieldType -eq $mergeFieldType }).Count | Should -Be 0
        } finally {
            $document.Dispose()
        }

        (Find-OfficeWord -Path $path -Text 'Jane').Count | Should -BeGreaterThan 0
        (Find-OfficeWord -Path $path -Text '77').Count | Should -BeGreaterThan 0
    }

    It 'preserves the legacy Word mail merge data alias' {
        $path = Join-Path $TestDrive 'DslMailMergeDataAlias.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph {
                Add-OfficeWordField -Type MergeField -Parameters '"FirstName"'
            }

            Invoke-OfficeWordMailMerge -Data @{
                FirstName = 'Ada'
            }
        } | Out-Null

        (Find-OfficeWord -Path $path -Text 'Ada').Count | Should -BeGreaterThan 0
    }

    It 'updates Word text and hyperlink metadata from a path' {
        $path = Join-Path $TestDrive 'DslReplaceText.docx'

        New-OfficeWord -Path $path {
            WordParagraph {
                WordText 'FY24 status'
                WordHyperlink -Text 'Portal FY24' -Url 'https://old.example.com/FY24' -Tooltip 'FY24 link'
            }

            WordParagraph {
                WordHyperlink -Text 'Jump FY24' -Anchor 'FY24Summary' -Tooltip 'FY24 anchor'
            }

            WordParagraph {
                WordText 'Summary'
                WordBookmark -Name 'FY24Summary'
            }
        } | Out-Null

        $replacements = Update-OfficeWordText -Path $path -OldValue 'FY24' -NewValue 'FY25' -IncludeHyperlinkText -IncludeHyperlinkUri -IncludeHyperlinkAnchor -IncludeHyperlinkTooltip -PassThru
        $replacements | Should -BeGreaterThan 0

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            (Find-OfficeWord -Document $document -Text 'FY25').Count | Should -BeGreaterThan 0
            $document.HyperLinks[0].Text | Should -Be 'Portal FY25'
            $document.HyperLinks[0].Uri.OriginalString | Should -Be 'https://old.example.com/FY25'
            $document.HyperLinks[0].Tooltip | Should -Be 'FY25 link'
            $document.HyperLinks[1].Text | Should -Be 'Jump FY25'
            $document.HyperLinks[1].Anchor | Should -Be 'FY25Summary'
            $document.HyperLinks[1].Tooltip | Should -Be 'FY25 anchor'
            $document.Bookmarks.Name | Should -Contain 'FY25Summary'
            $document.Bookmarks.Name | Should -Not -Contain 'FY24Summary'
        } finally {
            $document.Dispose()
        }
    }

    It 'updates relative hyperlink targets when requested' {
        $path = Join-Path $TestDrive 'DslReplaceRelativeLink.docx'

        New-OfficeWord -Path $path {
            WordParagraph {
                WordHyperlink -Text 'Relative FY24' -Url 'https://placeholder.invalid/FY24'
            }
        } | Out-Null

        $editable = Get-OfficeWord -Path $path
        try {
            $editable.HyperLinks[0].Uri = [Uri]::new('../reports/FY24.docx', [UriKind]::RelativeOrAbsolute)
            Save-OfficeWord -Document $editable | Out-Null
        } finally {
            $editable.Dispose()
        }

        $replacements = Update-OfficeWordText -Path $path -OldValue 'FY24' -NewValue 'FY25' -IncludeHyperlinkUri -PassThru
        $replacements | Should -Be 2

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.HyperLinks[0].Text | Should -Be 'Relative FY25'
            $document.HyperLinks[0].Uri.OriginalString | Should -Be '../reports/FY25.docx'
        } finally {
            $document.Dispose()
        }
    }

    It 'does not mutate live Word documents when text update uses WhatIf' {
        $path = Join-Path $TestDrive 'DslReplaceLiveWhatIf.docx'
        $document = New-OfficeWord -Path $path -NoSave

        try {
            $document.AddParagraph('FY24 live document') | Out-Null

            Update-OfficeWordText -Document $document -OldValue 'FY24' -NewValue 'FY25' -WhatIf -PassThru | Should -BeNullOrEmpty

            (Find-OfficeWord -Document $document -Text 'FY24').Count | Should -Be 1
            (Find-OfficeWord -Document $document -Text 'FY25').Count | Should -Be 0
        } finally {
            Close-OfficeWord -Document $document
        }
    }

    It 'does not mutate the tracked Word document when text update uses WhatIf' {
        $path = Join-Path $TestDrive 'DslReplaceTrackedWhatIf.docx'
        $document = New-OfficeWord -Path $path -NoSave

        try {
            $document.AddParagraph('FY24 tracked document') | Out-Null

            Update-OfficeWordText -OldValue 'FY24' -NewValue 'FY25' -WhatIf -PassThru | Should -BeNullOrEmpty

            (Find-OfficeWord -Document $document -Text 'FY24').Count | Should -Be 1
            (Find-OfficeWord -Document $document -Text 'FY25').Count | Should -Be 0
        } finally {
            Close-OfficeWord -Document $document
        }
    }

    It 'tracks current Word documents for update and close operations' {
        $pathOne = Join-Path $TestDrive 'TrackedOne.docx'
        $pathTwo = Join-Path $TestDrive 'TrackedTwo.docx'

        $docOne = New-OfficeWord -Path $pathOne -NoSave
        $docTwo = New-OfficeWord -Path $pathTwo -NoSave

        try {
            $docOne.AddParagraph('First tracked document') | Out-Null
            $docTwo.AddParagraph('Second tracked FY24 document') | Out-Null

            Update-OfficeWordText -OldValue 'FY24' -NewValue 'FY25' -PassThru | Should -Be 1

            Close-OfficeWord -Save
            Close-OfficeWord -All -Save
        } finally {
            foreach ($doc in @($docOne, $docTwo)) {
                if ($null -ne $doc) {
                    try {
                        $doc.Dispose()
                    } catch {
                    }
                }
            }
        }

        Test-Path $pathOne | Should -BeTrue
        Test-Path $pathTwo | Should -BeTrue

        $savedOne = Get-OfficeWord -Path $pathOne -ReadOnly
        $savedTwo = Get-OfficeWord -Path $pathTwo -ReadOnly
        try {
            (Find-OfficeWord -Document $savedOne -Text 'First tracked document').Count | Should -Be 1
            (Find-OfficeWord -Document $savedTwo -Text 'FY25').Count | Should -Be 1
        } finally {
            $savedOne.Dispose()
            $savedTwo.Dispose()
        }
    }

    It 'does not fall back to the tracked document when -Document is null' {
        $path = Join-Path $TestDrive 'NullDocumentGuard.docx'
        $doc = New-OfficeWord -Path $path -NoSave

        try {
            $nullDocument = $null
            { Close-OfficeWord -Document $nullDocument } | Should -Throw

            { $doc.AddParagraph('Still tracked after null guard') | Out-Null } | Should -Not -Throw
            Close-OfficeWord -Document $doc -Save
        } finally {
            if ($null -ne $doc) {
                try {
                    $doc.Dispose()
                } catch {
                }
            }
        }

        $saved = Get-OfficeWord -Path $path -ReadOnly
        try {
            (Find-OfficeWord -Document $saved -Text 'Still tracked after null guard').Count | Should -Be 1
        } finally {
            $saved.Dispose()
        }
    }

    It 'wraps OfficeIMO Word image and shape editing helpers' {
        $path = Join-Path $TestDrive 'DslWordImagesAndShapes.docx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        New-OfficeWord -Path $path {
            WordParagraph {
                $image = WordImage -Path $imagePath -Width 32 -Height 32 -PassThru
                $image |
                    Set-OfficeWordImage -Width 48 -Height 24 -Title 'Status Logo' -Description 'Logo alt text' -HorizontalFlip $true -Rotation 5 -PassThru |
                    Should -Not -BeNullOrEmpty
            }

            WordParagraph {
                $shape = WordShape -Type RoundedRectangle -Width 144 -Height 36 -FillColor '#DDEEFF' -StrokeColor '#1F4E79' -StrokeWidth 1 -Title 'Callout' -Description 'Callout shape' -PassThru
                $shape |
                    Set-OfficeWordShape -Left 24 -Top 12 -Rotation 2 -ZIndex 3 -PassThru |
                    Should -Not -BeNullOrEmpty
                $shape.FillColorHex | Should -Be 'ddeeff'
                $shape.StrokeColorHex | Should -Be '1f4e79'
            }
        } | Out-Null

        $images = @(Get-OfficeWordImage -Path $path)
        $shapes = @(Get-OfficeWordShape -Path $path)

        $images.Count | Should -Be 1
        $images[0].Title | Should -Be 'Status Logo'
        $images[0].Description | Should -Be 'Logo alt text'
        $images[0].HorizontalFlip | Should -BeTrue

        $shapes.Count | Should -Be 1
        $shapes[0].FillColorHex | Should -Be 'ddeeff'
        $shapes[0].StrokeColorHex | Should -Be '1f4e79'
    }

    It 'wraps OfficeIMO Word table cell read and style helpers' {
        $path = Join-Path $TestDrive 'DslWordTableCells.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; State = 'Ready' }
            [PSCustomObject]@{ Name = 'Beta'; State = 'Blocked' }
        )

        New-OfficeWord -Path $path {
            $table = WordTable -InputObject $rows -Style TableGrid -PassThru
            $cell = $table |
                Get-OfficeWordTableCell -Row 1 -Column 1 |
                Set-OfficeWordTableCell -ShadingFillColor '#DDEEFF' -Width 2400 -WidthType Dxa -WrapText $false -FitText $true -PassThru

            $cell.ShadingFillColorHex | Should -Be 'DDEEFF'
            $cell.Width | Should -Be 2400
            $cell.FitText | Should -BeTrue
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $cell = $table | Get-OfficeWordTableCell -Row 1 -Column 1

            $cell.ShadingFillColorHex | Should -Be 'DDEEFF'
            $cell.Width | Should -Be 2400
            $cell.WidthType.ToString() | Should -Be 'Dxa'
        } finally {
            $document.Dispose()
        }
    }

    It 'finds an existing Word table and appends rows' {
        $path = Join-Path $TestDrive 'ExistingWordTableAppend.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Risk register'; State = 'Open' }
        )

        New-OfficeWord -Path $path {
            WordParagraph -Text 'Existing table follows'
            WordTable -InputObject $rows -Style TableGrid
        } | Out-Null

        $document = Get-OfficeWord -Path $path
        try {
            $table = Find-OfficeWordTable -Document $document -Text 'Risk register' | Select-Object -First 1
            $table | Should -Not -BeNullOrEmpty

            $row = $table | Add-OfficeWordTableRow -Values @('Mitigation plan', 'Ready') -PassThru
            $row.Cells[0].Paragraphs[0].Text | Should -Be 'Mitigation plan'
            $row.Cells[1].Paragraphs[0].Text | Should -Be 'Ready'

            $table |
                Get-OfficeWordTableCell -Row 1 -Column 1 |
                Set-OfficeWordTableCell -Text 'Investigating'

            Close-OfficeWord -Document $document -Save
            $document = $null
        } finally {
            if ($null -ne $document) {
                $document.Dispose()
            }
        }

        $saved = Get-OfficeWord -Path $path -ReadOnly
        try {
            $savedTable = Find-OfficeWordTable -Document $saved -Text 'Mitigation plan' | Select-Object -First 1
            $savedTable | Should -Not -BeNullOrEmpty
            $savedTable.RowsCount | Should -Be 3
            $savedTable.Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'Investigating'
        } finally {
            $saved.Dispose()
        }
    }

    It 'preserves Word table width when appending after a merged first row' {
        $path = Join-Path $TestDrive 'ExistingWordTableMergedHeaderAppend.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Risk register'; State = 'Open' }
        )

        New-OfficeWord -Path $path {
            $table = WordTable -InputObject $rows -Style TableGrid -PassThru
            $table |
                Get-OfficeWordTableCell -Row 0 -Column 0 |
                Set-OfficeWordTableCell -Text 'Merged risk header' -MergeRight 1
        } | Out-Null

        $document = Get-OfficeWord -Path $path
        try {
            $table = Find-OfficeWordTable -Document $document -Text 'Risk register' | Select-Object -First 1
            $table | Should -Not -BeNullOrEmpty
            $table.Rows[0].CellsCount | Should -Be 1
            $table.Rows[1].CellsCount | Should -Be 2

            $row = $table | Add-OfficeWordTableRow -Values 'Mitigation plan' -PassThru
            $row.CellsCount | Should -Be 2
            $row.Cells[0].Paragraphs[0].Text | Should -Be 'Mitigation plan'
            $row.Cells[1].Paragraphs[0].Text | Should -Be ''

            Close-OfficeWord -Document $document -Save
            $document = $null
        } finally {
            if ($null -ne $document) {
                $document.Dispose()
            }
        }

        $saved = Get-OfficeWord -Path $path -ReadOnly
        try {
            $savedTable = Find-OfficeWordTable -Document $saved -Text 'Mitigation plan' | Select-Object -First 1
            $savedTable | Should -Not -BeNullOrEmpty
            $savedTable.Rows[2].CellsCount | Should -Be 2
        } finally {
            $saved.Dispose()
        }
    }

    It 'finds an existing Word list and appends items' {
        $path = Join-Path $TestDrive 'ExistingWordListAppend.docx'

        New-OfficeWord -Path $path {
            WordParagraph -Text 'Existing checklist follows'
            WordList {
                WordListItem -Text 'Initial review'
            }
        } | Out-Null

        $document = Get-OfficeWord -Path $path
        try {
            $lists = @($document | Get-OfficeWordList)
            $lists.Count | Should -Be 1

            $item = Find-OfficeWordList -Document $document -Text 'Initial review' |
                Add-OfficeWordListItem -Text 'Final approval' -PassThru
            $item.Text | Should -Be 'Final approval'

            Close-OfficeWord -Document $document -Save
            $document = $null
        } finally {
            if ($null -ne $document) {
                $document.Dispose()
            }
        }

        $saved = Get-OfficeWord -Path $path -ReadOnly
        try {
            (Find-OfficeWord -Document $saved -Text 'Final approval').Count | Should -Be 1
            $list = Find-OfficeWordList -Document $saved -Text 'Initial review' | Select-Object -First 1
            $list.ListItems.Text | Should -Contain 'Final approval'
        } finally {
            $saved.Dispose()
        }
    }

    It 'wraps OfficeIMO Word paragraph and text style helpers' {
        $path = Join-Path $TestDrive 'DslWordStyles.docx'

        New-OfficeWord -Path $path {
            $paragraph = WordParagraph -Text 'Executive Summary' -PassThru
            $paragraph |
                Set-OfficeWordParagraphStyle -Style Heading2 -Alignment Center -SpacingBeforePoints 6 -SpacingAfterPoints 12 -IndentationBeforePoints 18 -KeepWithNext $true -PassThru |
                Should -Not -BeNullOrEmpty

            $textParagraph = WordParagraph -Text 'Initial' -PassThru
            $textItem = @($textParagraph.GetRuns())[0]
            $textItem |
                Set-OfficeWordTextStyle -Text 'Styled text' -Bold $true -Italic $true -Underline Single -Color '#C00000' -FontSize 14 -FontFamily 'Aptos' -Highlight Yellow -CapsStyle SmallCaps -Strike $true -Style Heading2Char -PassThru |
                Should -Not -BeNullOrEmpty
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $styledParagraph = $document.Paragraphs | Where-Object Text -EQ 'Executive Summary' | Select-Object -First 1
            $styledParagraph.Style.ToString() | Should -Be 'Heading2'
            $styledParagraph.ParagraphAlignment.ToString() | Should -Be 'Center'
            $styledParagraph.LineSpacingBeforePoints | Should -Be 6
            $styledParagraph.LineSpacingAfterPoints | Should -Be 12
            $styledParagraph.IndentationBeforePoints | Should -Be 18
            $styledParagraph.KeepWithNext | Should -BeTrue

            $styledText = $document.Paragraphs |
                Where-Object Text -EQ 'Styled text' |
                ForEach-Object { $_.GetRuns() } |
                Select-Object -First 1

            $styledText.Text | Should -Be 'Styled text'
            $styledText.Bold | Should -BeTrue
            $styledText.Italic | Should -BeTrue
            $styledText.Underline.ToString() | Should -Be 'Single'
            $styledText.ColorHex | Should -Be 'c00000'
            $styledText.FontSize | Should -Be 14
            $styledText.FontFamily | Should -Be 'Aptos'
            $styledText.Highlight.ToString() | Should -Be 'Yellow'
            $styledText.CapsStyle.ToString() | Should -Be 'SmallCaps'
            $styledText.Strike | Should -BeTrue
            $styledText.CharacterStyle.ToString() | Should -Be 'Heading2Char'
        } finally {
            $document.Dispose()
        }
    }

    It 'adds OfficeIMO Word text boxes for report callouts' {
        $path = Join-Path $TestDrive 'DslWordTextBox.docx'

        New-OfficeWord -Path $path {
            WordParagraph -Text 'Report callouts'
            $textBox = WordTextBox -Text 'Service posture is stable' -WidthCentimeters 7 -HeightCentimeters 2 -HorizontalOffsetCentimeters 1 -VerticalOffsetCentimeters 1 -HorizontalAlignment Center -AutoFitToTextSize -PassThru
            $textBox.WidthCentimeters | Should -BeGreaterThan 6.9
            $textBox.HeightCentimeters | Should -BeGreaterThan 1.9
            $textBox.AutoFitToTextSize | Should -BeTrue
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $textBoxes = @($document.TextBoxes)
            $textBoxes.Count | Should -Be 1
            $textBoxes[0].Paragraphs.Text -join "`n" | Should -Match 'Service posture is stable'
            [Math]::Round($textBoxes[0].WidthCentimeters, 1) | Should -Be 7.0
            [Math]::Round($textBoxes[0].HeightCentimeters, 1) | Should -Be 2.0
        } finally {
            $document.Dispose()
        }
    }
}
