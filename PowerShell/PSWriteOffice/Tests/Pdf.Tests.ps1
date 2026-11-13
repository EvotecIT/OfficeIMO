BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'PDF cmdlets' {
    It 'reads normalized PDF content and inspects the native live document' {
        $path = Join-Path $TestDrive 'structured-read.pdf'
        New-OfficePdf -Path $path {
            PdfHeading 'Structured PDF report'
            PdfParagraph 'Canonical paragraph marker'
            PdfImage -Path (Join-Path $PSScriptRoot 'Assets\CellImage.png') -Width 120 -Height 80
            PdfTable -InputObject @(
                [pscustomobject]@{ Service = 'Reader'; Status = 'Ready' }
                [pscustomobject]@{ Service = 'OCR'; Status = 'Ready' }
            )
        } | Out-Null

        $logical = Read-OfficeDocument -Path $path -Options (New-OfficePdfReadOptions -Profile Structured)
        $logical.GetType().FullName | Should -Be 'OfficeIMO.Reader.OfficeDocumentReadResult'
        @($logical.Pages).Count | Should -BeGreaterThan 0
        ($logical | Read-OfficeDocument -As Text) | Should -Match 'Canonical paragraph marker'
        @($logical.Tables).Count | Should -BeGreaterThan 0
        @($logical.Tables.Rows).Count | Should -BeGreaterThan 0
        @($logical.Assets).Count | Should -BeGreaterThan 0

        $images = @(Get-OfficePdfImage -Path $path)
        $images.Count | Should -BeGreaterThan 0
        $images[0].GetType().FullName | Should -Be 'OfficeIMO.Pdf.PdfExtractedImage'

        $live = Open-OfficePdf -LiteralPath $path
        try {
            $fromLive = $live.Read((New-OfficePdfReadOptions -Page 1).LoadOptions)
            $fromLive.GetType().FullName | Should -Be 'OfficeIMO.Pdf.PdfDocumentReadResult'
            @($fromLive.Pages).Count | Should -Be 1
            $fromLive.Text | Should -Match 'Canonical paragraph marker'
        } finally { $live.Dispose() }
    }

    It 'extracts images without replacing an existing destination' {
        $path = Join-Path $TestDrive 'image-extraction-source.pdf'
        New-OfficePdf -Path $path {
            PdfImage -Path (Join-Path $PSScriptRoot 'Assets\CellImage.png') -Width 120 -Height 80
        } | Out-Null

        $outputDirectory = Join-Path $TestDrive 'extracted-images'
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
        $firstImage = @(Get-OfficePdfImage -Path $path)[0]
        $extension = $firstImage.FileExtension.TrimStart('.')
        $existingPath = Join-Path $outputDirectory "proof-0001.$extension"
        [IO.File]::WriteAllText($existingPath, 'keep this file')

        $written = @(Get-OfficePdfImage -Path $path -OutputDirectory $outputDirectory -BaseName proof)

        $written.Count | Should -Be 1
        $written[0].Name | Should -Be "proof-0001-2.$extension"
        [IO.File]::ReadAllText($existingPath) | Should -Be 'keep this file'
        $written[0].Length | Should -BeGreaterThan 0
        @(Get-ChildItem -LiteralPath $outputDirectory -Filter '.*.tmp*').Count | Should -Be 0
    }

    It 'validates PDF operations and page selections before path I/O' {
        $missingPath = Join-Path $TestDrive 'missing-source.pdf'
        $outputPath = Join-Path $TestDrive 'validation-output.pdf'
        $exportDirectory = Join-Path $TestDrive 'validation-export'

        { Split-OfficeDocument -Path $missingPath -OutputDirectory $exportDirectory -PageCount -1 -ErrorAction Stop } |
            Should -Throw '*PageCount*'
        { Split-OfficeDocument -Path $missingPath -OutputDirectory $exportDirectory -Page 'invalid' -ErrorAction Stop } |
            Should -Throw '*page*'
        { Set-OfficePdfPage -Path $missingPath -OutputPath $outputPath -ErrorAction Stop } |
            Should -Throw '*Choose one operation*'
        { Compare-OfficeDocument -Visual -ReferencePath $missingPath -DifferencePath $missingPath -Page 'invalid' -ErrorAction Stop } |
            Should -Throw '*page*'
        { Convert-OfficeDocument -To Markdown -Path $missingPath -Page 'invalid' -ErrorAction Stop -AsString } |
            Should -Throw '*page*'
        { Export-OfficeDocumentImage -Path $missingPath -OutputDirectory $exportDirectory -Page 'invalid' -ErrorAction Stop } |
            Should -Throw '*page*'
        { Convert-OfficeDocument -To Html -Path $missingPath -Page 'invalid' -ErrorAction Stop -AsString } |
            Should -Throw '*page*'

        Test-Path -LiteralPath $exportDirectory | Should -BeFalse
    }

    It 'preserves advanced PDF read options unless a friendly parameter is supplied' {
        $path = Join-Path $TestDrive 'structured-read-options.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'First page marker'
            PdfPageBreak
            PdfParagraph 'Second page marker'
        } | Out-Null

        $options = New-OfficePdfReadOptions -Profile Fast -Page 2
        $fromOptions = @(Read-OfficeDocument -LiteralPath $path -As Text -ByPage -Options $options)
        $fromOptions.Count | Should -Be 1
        $fromOptions[0].Page.Number | Should -Be 2
        $fromOptions[0].Text | Should -Match 'Second page marker'

        $overrides = New-OfficePdfReadOptions -LoadOptions $options.LoadOptions -Profile Structured -Page 1
        $overridden = @(Read-OfficeDocument -LiteralPath $path -As Text -ByPage -Options $overrides)
        $overridden.Count | Should -Be 1
        $overridden[0].Page.Number | Should -Be 1
        $overridden[0].Text | Should -Match 'First page marker'
        $options.LoadOptions.Profile.ToString() | Should -Be 'Fast'
        @(Read-OfficeDocument -LiteralPath $path -As Text -ByPage -Options $options)[0].Page.Number | Should -Be 2
    }

    It 'supports passwords for encrypted PDF read cmdlets' {
        $encryptedPath = Join-Path $TestDrive 'encrypted.pdf'
        [IO.File]::WriteAllBytes($encryptedPath, [Convert]::FromBase64String('JVBERi0xLjQKJT8/Pz8KMSAwIG9iago8PCAvVHlwZSAvQ2F0YWxvZyAvUGFnZXMgMiAwIFIgPj4KZW5kb2JqCjIgMCBvYmoKPDwgL1R5cGUgL1BhZ2VzIC9LaWRzIFszIDAgUl0gL0NvdW50IDEgPj4KZW5kb2JqCjMgMCBvYmoKPDwgL1R5cGUgL1BhZ2UgL1BhcmVudCAyIDAgUiAvTWVkaWFCb3ggWzAgMCAzMDAgMjAwXSAvUmVzb3VyY2VzIDw8IC9Gb250IDw8IC9GMSA0IDAgUiA+PiA+PiAvQ29udGVudHMgNSAwIFIgPj4KZW5kb2JqCjQgMCBvYmoKPDwgL1R5cGUgL0ZvbnQgL1N1YnR5cGUgL1R5cGUxIC9CYXNlRm9udCAvSGVsdmV0aWNhID4+CmVuZG9iago1IDAgb2JqCjw8IC9MZW5ndGggNDYgPj4Kc3RyZWFtCi8xB1rv33VvGaaV2g01B7caayy3ttoqyqa6Fkx+aapdiBLgquqJCxhp8zWpAXQKZW5kc3RyZWFtCmVuZG9iago2IDAgb2JqCjw8IC9GaWx0ZXIgL1N0YW5kYXJkIC9WIDEgL1IgMiAvTGVuZ3RoIDQwIC9PIDw4RUVCMDk1ODE5NjYyQTc3NDQ0MkZCMDcyRTNEOUYxOUU5RDEzMEVDMDlBNEQwMDYxRTc4RkU5MjBGN0FCNjJGPiAvVSA8QjFFNzY0MTI2QzQ4RDI4RDkwNTI1NTk1MjAwREQ4MTg3NEI4NkZFMUNBNTRCQTAxODZFNThCRTJDMzU5ODhEQz4gL1AgLTQgPj4KZW5kb2JqCnhyZWYKMCA3CjAwMDAwMDAwMDAgNjU1MzUgZiAKMDAwMDAwMDAxNSAwMDAwMCBuIAowMDAwMDAwMDY0IDAwMDAwIG4gCjAwMDAwMDAxMjEgMDAwMDAgbiAKMDAwMDAwMDI0NyAwMDAwMCBuIAowMDAwMDAwMzE3IDAwMDAwIG4gCjAwMDAwMDA0MTMgMDAwMDAgbiAKdHJhaWxlcgo8PCAvU2l6ZSA3IC9Sb290IDEgMCBSIC9FbmNyeXB0IDYgMCBSIC9JRCBbPDEwNDVBODdDMjIxODRFQzE5MTRBQ0Y2NjMxRDI3NDAzPiA8MTA0NUE4N0MyMjE4NEVDMTkxNEFDRjY2MzFEMjc0MDM+XSA+PgpzdGFydHhyZWYKNjE5CiUlRU9GCg=='))

        (Get-OfficePdfDiagnostic -Type Preflight -Path $encryptedPath).CanRead | Should -BeFalse
        (Get-OfficePdfDiagnostic -Type Preflight -Path $encryptedPath -Password 'open').CanRead | Should -BeTrue
        Read-OfficeDocument -As Text -Path $encryptedPath -Password 'open' | Should -Match 'Secret PDF Text'
        (Open-OfficePdf -Path $encryptedPath -Password 'open').Read().Text | Should -Match 'Secret PDF Text'
        Convert-OfficeDocument -To Markdown -Path $encryptedPath -Password 'open' -AsString | Should -Match 'Secret PDF Text'
    }

    It 'writes password encrypted PDFs from new and save cmdlets' {
        foreach ($name in 'New-OfficePdf', 'Save-OfficePdf') {
            (Get-Command $name).Parameters.Keys | Should -Contain 'Password'
            (Get-Command $name).Parameters.Keys | Should -Contain 'OwnerPassword'
            (Get-Command $name).Parameters.Keys | Should -Contain 'Permission'
        }

        $newPath = Join-Path $TestDrive 'new-encrypted.pdf'
        New-OfficePdf -Path $newPath -Password 'open' -OwnerPassword 'owner' {
            PdfParagraph 'Generated encrypted PDF text'
        } | Out-Null

        (Get-OfficePdfDiagnostic -Type Preflight -Path $newPath).CanRead | Should -BeFalse
        (Get-OfficePdfDiagnostic -Type Preflight -Path $newPath -Password 'open').CanRead | Should -BeTrue
        Read-OfficeDocument -As Text -Path $newPath -Password 'open' | Should -Match 'Generated encrypted PDF text'
        { Read-OfficeDocument -As Text -Path $newPath -Password 'wrong' -ErrorAction Stop } | Should -Throw

        $savedPath = Join-Path $TestDrive 'saved-encrypted.pdf'
        $doc = New-OfficePdf
        PdfParagraph -Document $doc 'Saved encrypted PDF text'

        $doc | Save-OfficePdf -Path $savedPath -Password 'open' -OwnerPassword 'owner' | Out-Null

        (Get-OfficePdfDiagnostic -Type Preflight -Path $savedPath).CanRead | Should -BeFalse
        Read-OfficeDocument -As Text -Path $savedPath -Password 'open' | Should -Match 'Saved encrypted PDF text'
    }

    It 'does not run the PDF DSL when WhatIf skips saving' {
        $path = Join-Path $TestDrive 'PdfWhatIf.pdf'
        $script:PdfWhatIfDslRan = $false

        New-OfficePdf -Path $path -WhatIf {
            $script:PdfWhatIfDslRan = $true
            PdfParagraph 'Should not run'
        } | Out-Null

        $script:PdfWhatIfDslRan | Should -BeFalse
        Test-Path -LiteralPath $path | Should -BeFalse
    }

    It 'rejects document page composition and deferred context pass-through clearly' {
        $document = New-OfficePdf
        PdfParagraph -Document $document 'Existing document content'

        { $document | PdfPageSetup -Margin 36 -ErrorAction Stop } |
            Should -Throw '*input object cannot be bound*'
        { New-OfficePdf -Path (Join-Path $TestDrive 'deferred-pass-through.pdf') { PdfParagraph 'First' -PassThru -ErrorAction Stop | PdfParagraph 'Second' -ErrorAction Stop } } |
            Should -Throw '*Document*'
    }

    It 'builds a composed PDF and extracts text' {
        $path = Join-Path $TestDrive 'report.pdf'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $file = New-OfficePdf -Path $path {
            PdfMetadata -Title 'Operations Report' -Author 'PSWriteOffice'
            PdfPageSetup -PageSize A4 -Margin 54
            PdfHeader 'Operations Report'
            PdfFooter 'Page {page}/{pages}'
            PdfHeading 'Operations Report'
            PdfParagraph 'Generated by PSWriteOffice'
            PdfPanel 'This is a visually separated note.'
            PdfList -Items 'Draft', 'Review', 'Ship' -Numbered
            PdfTable -InputObject $rows
            PdfPageBreak
            PdfHeading -Level 2 -Text 'Details'
            PdfParagraph 'Second page content'
        } -PassThru

        $file | Should -BeOfType System.IO.FileInfo
        Test-Path $path | Should -BeTrue

        $info = (Get-OfficeDocumentInfo -Detailed -Path $path).Details
        $info.PageCount | Should -BeGreaterOrEqual 1

        $preflight = Get-OfficePdfDiagnostic -Type Preflight -Path $path
        $preflight.CanRead | Should -BeTrue

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Operations Report'
        $text | Should -Match 'Generated by PSWriteOffice'
    }

    It 'renders hashtable rows using their dictionary values' {
        $path = Join-Path $TestDrive 'hashtable-table.pdf'
        $rows = @(
            @{ Name = 'Alpha'; Value = 1 }
            @{ Name = 'Beta'; Value = 2 }
        )

        New-OfficePdf -Path $path {
            PdfTable -InputObject $rows -Property Name, Value
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Alpha'
        $text | Should -Match '1'
        $text | Should -Match 'Beta'
        $text | Should -Match '2'
    }

    It 'renders multiple row arrays as table rows' {
        $path = Join-Path $TestDrive 'array-row-table.pdf'

        New-OfficePdf -Path $path {
            PdfTable -InputObject @(
                @('Name', 'Value'),
                @('Alpha', 1),
                @('Beta', 2)
            )
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Name'
        $text | Should -Match 'Alpha'
        $text | Should -Match 'Beta'
    }

    It 'renders explicit headers for array row tables' {
        $path = Join-Path $TestDrive 'array-row-table-header.pdf'

        New-OfficePdf -Path $path {
            PdfTable -Header Name, Value -InputObject @(
                @('Alpha', 1),
                @('Beta', 2)
            )
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Name'
        $text | Should -Match 'Value'
        $text | Should -Match 'Alpha'
        $text | Should -Match 'Beta'
    }

    It 'renders PDF table cell spans in PDF tables' {
        (Get-Command New-OfficePdfTableCell).Parameters.Keys | Should -Contain 'ColumnSpan'
        (Get-Command New-OfficePdfTableCell).Parameters.Keys | Should -Contain 'RowSpan'
        Get-Command New-OfficeTableCell -ErrorAction SilentlyContinue | Should -BeNullOrEmpty

        $path = Join-Path $TestDrive 'span-aware-table.pdf'

        New-OfficePdf -Path $path {
            PdfTable -HeaderRowCount 1 -InputObject @(
                @('Service', 'Status', 'Owner'),
                @(New-OfficePdfTableCell -Text 'Identity systems' -ColumnSpan 3),
                @('Entra', 'Watch', 'IAM'),
                @((New-OfficePdfTableCell -Text 'Shared owner' -RowSpan 2), 'Build', 'OfficeIMO'),
                @('Release', 'PSWriteOffice')
            )
        } | Out-Null

        (Get-OfficePdfDiagnostic -Type Preflight -Path $path).CanRead | Should -BeTrue
        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Identity systems'
        $text | Should -Match 'Shared owner'
        $text | Should -Match 'Release'
    }

    It 'styles merged PDF table cells' {
        $path = Join-Path $TestDrive 'styled-span-aware-table.pdf'

        New-OfficePdf -Path $path {
            PdfTable -HeaderRowCount 1 -InputObject @(
                @('Service', 'Status', 'Owner'),
                @(New-OfficePdfTableCell -Text 'Identity systems' -ColumnSpan 3 -BackgroundColor '#DBEAFE' -Color '#1E3A8A' -Size 17 -Bold -Italic -Alignment Center -VerticalAlignment Center),
                @('Entra', 'Watch', 'IAM'),
                @(@{ Text = 'Follow-up'; ColumnSpan = 3; FillColor = '#FEF3C7'; Color = '#92400E'; FontSize = 15; Bold = $true; Align = 'Center' }),
                @('Release', 'Ready', 'PSWriteOffice')
            )
        } | Out-Null

        (Get-OfficePdfDiagnostic -Type Preflight -Path $path).CanRead | Should -BeTrue
        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Identity systems'
        $text | Should -Match 'Follow-up'

        $blocks = @(& { $pdfReadDocument = Open-OfficePdf -Path $path; try { $pdfReadDocument.Read().TextBlocks } finally { $pdfReadDocument.Dispose() } })
        $identityBlock = $blocks | Where-Object { $_.Text -match 'Identity systems' } | Select-Object -First 1
        $identityBlock | Should -Not -BeNullOrEmpty
        $identityBlock.FontSize | Should -Be 17

        $followUpBlock = $blocks | Where-Object { $_.Text -match 'Follow-up' } | Select-Object -First 1
        $followUpBlock | Should -Not -BeNullOrEmpty
        $followUpBlock.FontSize | Should -Be 15
    }

    It 'renders rich text runs and named colors in PDF text and table cells' {
        $path = Join-Path $TestDrive 'rich-run-table.pdf'

        New-OfficePdf -Path $path {
            PdfParagraph {
                PdfText 'Status: ' -Bold:$false
                PdfText 'Ready' -Color SeaGreen -Bold
            }
            PdfTable -InputObject @(
                , @(
                    PdfTableCell -Content {
                        PdfText 'Build '
                        PdfText 'Ready' -Color SeaGreen -Bold
                    } -ColumnSpan 2 -BackgroundColor AliceBlue
                )
                , @(
                    PdfTableCell -Content { PdfText 'Styled' -Bold } -Color Crimson -Size 19 -ColumnSpan 2
                )
                , @(
                    @{ Run = 'Queued' },
                    'String run'
                )
                , @('Owner', 'Platform')
            )
        } | Out-Null

        (Get-OfficePdfDiagnostic -Type Preflight -Path $path).CanRead | Should -BeTrue
        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Status'
        $text | Should -Match 'Ready'
        $text | Should -Match 'Build'
        $text | Should -Match 'Styled'
        $text | Should -Match 'Queued'
        $text | Should -Match 'Owner'

        $blocks = @(& { $pdfReadDocument = Open-OfficePdf -Path $path; try { $pdfReadDocument.Read().TextBlocks } finally { $pdfReadDocument.Dispose() } })
        $styledBlock = $blocks | Where-Object { $_.Text -match 'Styled' } | Select-Object -First 1
        $styledBlock | Should -Not -BeNullOrEmpty
        $styledBlock.FontSize | Should -Be 19
    }

    It 'preserves native PDF table-cell run link annotations' {
        $path = Join-Path $TestDrive 'rich-run-table-link.pdf'
        New-OfficePdf -Path $path {
            PdfTable -InputObject @(
                , @((PdfTableCell { PdfText 'Docs' -Link 'https://example.org/table' }))
            )
        }
        ((Get-OfficeDocumentInfo -Detailed -Path $path).Details).LinkUris | Should -Contain 'https://example.org/table'
    }

    It 'treats rich and styled PDF table cells as structured without spans' {
        $path = Join-Path $TestDrive 'rich-run-mixed-object-table.pdf'

        New-OfficePdf -Path $path {
            PdfTable -InputObject @(
                [pscustomobject]@{
                    Service = 'Entra'
                    Status  = 'Ready'
                }
                , @(
                    (PdfTableCell -BackgroundColor AliceBlue {
                        PdfText 'Build '
                        PdfText 'Watch' -Color DarkOrange -Bold
                    }),
                    'Platform'
                )
            )
        } | Out-Null

        (Get-OfficePdfDiagnostic -Type Preflight -Path $path).CanRead | Should -BeTrue
        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Service'
        $text | Should -Match 'Entra'
        $text | Should -Match 'Build'
        $text | Should -Match 'Watch'
        $text | Should -Match 'Platform'
    }

    It 'keeps ordinary style-named columns in PDF object tables' {
        $path = Join-Path $TestDrive 'ordinary-style-named-table.pdf'

        New-OfficePdf -Path $path {
            PdfTable -InputObject @(
                [pscustomobject]@{
                    Text     = 'Task'
                    Color    = 'Red'
                    FontSize = 'Large'
                }
            )
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Text'
        $text | Should -Match 'Color'
        $text | Should -Match 'FontSize'
        $text | Should -Match 'Red'
        $text | Should -Match 'Task'
        $text | Should -Match 'Large'
    }

    It 'keeps ordinary span-like property names on normal PDF tables' {
        $path = Join-Path $TestDrive 'ordinary-span-named-table.pdf'
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

        New-OfficePdf -Path $path {
            PdfTable -InputObject $rows
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Rows'
        $text | Should -Match 'Columns'
        $text | Should -Match 'Span'
        $text | Should -Match 'ColumnSpa'
        $text | Should -Match 'RowSpan'
        $text | Should -Match '25'
        $text | Should -Match '3'
        $text | Should -Match '2'
    }

    It 'keeps ordinary Run property names on normal PDF tables' {
        $path = Join-Path $TestDrive 'ordinary-run-named-table.pdf'
        $rows = @(
            [pscustomobject]@{
                Run = 'Nightly'
            }
        )

        New-OfficePdf -Path $path {
            PdfTable -InputObject $rows
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Run'
        $text | Should -Match 'Nightly'
    }

    It 'keeps ordinary text and span-key properties in mixed PDF tables' {
        $path = Join-Path $TestDrive 'mixed-ordinary-span-key-table.pdf'
        $rows = @(
            [pscustomobject]@{
                Text = 'Task'
                FontSize = 'Large'
                ColumnSpan = 1
                Status = 'Open'
            }
            , @(New-OfficePdfTableCell -Text 'Follow-up' -ColumnSpan 3)
        )

        New-OfficePdf -Path $path {
            PdfTable -InputObject $rows
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Text'
        $text | Should -Match 'FontSize'
        $text | Should -Match 'ColumnSpa'
        $text | Should -Match 'Status'
        $text | Should -Match 'Task'
        $text | Should -Match 'Large'
        $text | Should -Match 'Open'
        $text | Should -Match 'Follow-up'
    }

    It 'keeps default headers on mixed object and span PDF tables' {
        $path = Join-Path $TestDrive 'mixed-object-span-table.pdf'
        $rows = @(
            [pscustomobject]@{
                Name = 'Directory'
                Status = 'Healthy'
            }
            , @(New-OfficePdfTableCell -Text 'Follow-up' -ColumnSpan 2)
        )

        New-OfficePdf -Path $path {
            PdfTable -InputObject $rows
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Name'
        $text | Should -Match 'Status'
        $text | Should -Match 'Directory'
        $text | Should -Match 'Healthy'
        $text | Should -Match 'Follow-up'
    }

    It 'keeps generated mixed PDF headers out of leading row spans' {
        $path = Join-Path $TestDrive 'mixed-leading-row-span-table.pdf'
        $rows = @(
            , @((New-OfficePdfTableCell -Text 'Group' -RowSpan 2), 'A')
            [pscustomobject]@{
                Name = 'B'
                Value = 'C'
            }
        )

        New-OfficePdf -Path $path {
            PdfTable -InputObject $rows
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text.IndexOf('Name') | Should -BeLessThan $text.IndexOf('Group')
        $text | Should -Match 'Value'
        $text | Should -Match 'Group'
        $text | Should -Match 'A'
        $text | Should -Match 'B'
        $text | Should -Match 'C'
    }

    It 'supports transposed table views' {
        $path = Join-Path $TestDrive 'transposed-table.pdf'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        New-OfficePdf -Path $path {
            PdfTable -InputObject $rows -View Transpose
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Property'
        $text | Should -Match 'Row1'
        $text | Should -Match 'Alpha'
        $text | Should -Match 'Beta'
    }

    It 'applies OfficeIMO PDF table style options' {
        $path = Join-Path $TestDrive 'styled-table.pdf'
        $rows = @(
            [pscustomobject]@{ Service = 'Directory'; Status = 'Healthy'; Incidents = 0 }
            [pscustomobject]@{ Service = 'Mail'; Status = 'Watch'; Incidents = 2 }
        )

        New-OfficePdf -Path $path {
            PdfTable -InputObject $rows -Property Service, Status, Incidents -TableStyle Report -Caption 'Service status' -CaptionAlignment Center -AutoFit -ColumnWeight 2, 1, 1
        } | Out-Null

        (Get-OfficePdfDiagnostic -Type Preflight -Path $path).CanRead | Should -BeTrue
        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Service status'
        $text | Should -Match 'Directory'
        $text | Should -Match 'Healthy'
    }

    It 'shrinks PDF table text to fit fixed-width cells' {

        $path = Join-Path $TestDrive 'shrink-table.pdf'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 'ThisIdentifierShouldShrinkToFit' }
        )

        New-OfficePdf -Path $path {
            PdfTable -InputObject $rows -Property Name, Value -ColumnWidth 54, 108 -Size 18 -ShrinkToFit -MinimumSize 7
        } | Out-Null

        $blocks = @(& { $pdfReadDocument = Open-OfficePdf -Path $path; try { $pdfReadDocument.Read().TextBlocks } finally { $pdfReadDocument.Dispose() } })
        $valueBlock = $blocks | Where-Object { $_.Text -match 'ThisIdentifierShouldShrinkToFit' } | Select-Object -First 1
        $valueBlock | Should -Not -BeNullOrEmpty
        $valueBlock.FontSize | Should -BeLessThan 18
        $valueBlock.FontSize | Should -BeGreaterOrEqual 7
    }

    It 'adds piped PDF table rows to a supplied document' {
        $path = Join-Path $TestDrive 'piped-rows-supplied-document.pdf'
        $doc = New-OfficePdf
        PdfHeading -Document $doc 'Pipeline rows'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $updated = $rows | PdfTable -Document $doc -PassThru
        $updated | Save-OfficePdf -Path $path | Out-Null

        $updated | Should -Be $doc
        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Alpha'
        $text | Should -Match 'Beta'
    }

    It 'adds explicit PDF table rows to each piped document' {
        $path1 = Join-Path $TestDrive 'piped-document-one.pdf'
        $path2 = Join-Path $TestDrive 'piped-document-two.pdf'
        $doc1 = New-OfficePdf
        PdfHeading -Document $doc1 'First document'
        $doc2 = New-OfficePdf
        PdfHeading -Document $doc2 'Second document'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $updated = @($doc1, $doc2) | PdfTable -InputObject $rows -PassThru
        $updated[0] | Save-OfficePdf -Path $path1 | Out-Null
        $updated[1] | Save-OfficePdf -Path $path2 | Out-Null

        $updated.Count | Should -Be 2
        $updated[0] | Should -Be $doc1
        $updated[1] | Should -Be $doc2
        (Read-OfficeDocument -As Text -Path $path1) | Should -Match 'Alpha'
        (Read-OfficeDocument -As Text -Path $path2) | Should -Match 'Alpha'
    }

    It 'keeps the current page size when only margins are updated' {
        $path = Join-Path $TestDrive 'page-size-margin.pdf'

        New-OfficePdf -Path $path {
            PdfPageSetup -PageSize A4
            PdfPageSetup -Margin 36
            PdfHeading 'A4 report'
        } | Out-Null

        $page = ((Get-OfficeDocumentInfo -Detailed -Path $path).Details).Pages[0]
        [math]::Round($page.Width) | Should -Be 595
        [math]::Round($page.Height) | Should -Be 842
    }

    It 'reports append-only mutation support for generated PDFs' {


        $path = Join-Path $TestDrive 'append-plan.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'Append-only metadata readiness'
        } | Out-Null

        $plan = Get-OfficePdfDiagnostic -Type Incremental -Path $path
        $plan.CanAppendMetadata | Should -BeTrue
        $plan.CanAppendFormFields | Should -BeTrue
        $plan.SupportedActions | Should -Contain 'Metadata'
        $plan.SupportedActions | Should -Contain 'FormFill'
        $plan.Summary | Should -Match 'Incremental updates are supported'
    }

    It 'supports expanded page sizes and PDF 2 file version' {
        $path = Join-Path $TestDrive 'b5-pdf20.pdf'

        New-OfficePdf -Path $path -FileVersion Pdf20 {
            PdfPageSetup -PageSize EnvelopeB5
            PdfHeading 'B5 PDF 2'
        } | Out-Null

        $info = (Get-OfficeDocumentInfo -Detailed -Path $path).Details
        $info.HeaderVersion | Should -Be '2.0'
        $info.EffectiveVersion | Should -Be '2.0'
        $info.IsPdf20OrLater | Should -BeTrue
        [math]::Abs($info.Pages[0].Width - 176 * 72 / 25.4) | Should -BeLessThan 0.001
        [math]::Abs($info.Pages[0].Height - 250 * 72 / 25.4) | Should -BeLessThan 0.001
    }

    It 'supports PDF page operations with approved verbs' {
        $path = Join-Path $TestDrive 'pages.pdf'
        New-OfficePdf -Path $path {
            PdfHeading 'Page 1'
            PdfPageBreak
            PdfHeading 'Page 2'
        } | Out-Null

        $rotated = Join-Path $TestDrive 'rotated.pdf'
        $copy = Join-Path $TestDrive 'copy.pdf'
        $removed = Join-Path $TestDrive 'removed.pdf'

        Set-OfficePdfPage -Path $path -Rotation 90 -OutputPath $rotated -PassThru | Should -BeOfType System.IO.FileInfo
        Copy-OfficePdfPage -Path $path -Page '1' -OutputPath $copy -PassThru | Should -BeOfType System.IO.FileInfo
        Remove-OfficePdfPage -Path $path -Page '2' -OutputPath $removed -PassThru | Should -BeOfType System.IO.FileInfo

        ((Get-OfficeDocumentInfo -Detailed -Path $copy).Details).PageCount | Should -Be 1
        ((Get-OfficeDocumentInfo -Detailed -Path $removed).Details).PageCount | Should -Be 1
    }

    It 'splits PDFs by page count, page range, and headings-derived bookmarks' {
        $path = Join-Path $TestDrive 'split-source.pdf'
        New-OfficePdf -Path $path -CreateOutlineFromHeadings {
            PdfHeading 'Chapter One'
            PdfParagraph 'First chapter body'
            PdfPageBreak
            PdfHeading 'Chapter Two'
            PdfParagraph 'Second chapter body'
            PdfPageBreak
            PdfHeading 'Chapter Three'
            PdfParagraph 'Third chapter body'
        } | Out-Null

        $pages = @(Read-OfficeDocument -As Text -Path $path -ByPage)
        $pages.Count | Should -Be 3
        $pages[0].Page.Number | Should -Be 1
        $pages[0].Text | Should -Match 'Chapter One'
        $pages[2].Text | Should -Match 'Chapter Three'

        $groupDirectory = Join-Path $TestDrive 'groups'
        $groups = @(Split-OfficeDocument -Path $path -OutputDirectory $groupDirectory -Prefix 'group' -PageCount 2)
        $groups.Count | Should -Be 2
        ((Get-OfficeDocumentInfo -Detailed -Path $groups[0].FullName).Details).PageCount | Should -Be 2
        ((Get-OfficeDocumentInfo -Detailed -Path $groups[1].FullName).Details).PageCount | Should -Be 1

        $rangeDirectory = Join-Path $TestDrive 'ranges'
        $ranges = @(Split-OfficeDocument -Path $path -OutputDirectory $rangeDirectory -Prefix 'range' -Page '1-2', '3')
        $ranges.Count | Should -Be 2
        Read-OfficeDocument -As Text -Path $ranges[0].FullName | Should -Match 'Chapter Two'
        Read-OfficeDocument -As Text -Path $ranges[1].FullName | Should -Match 'Chapter Three'

        $bookmarkDirectory = Join-Path $TestDrive 'bookmarks'
        $bookmarks = @(Split-OfficeDocument -Path $path -OutputDirectory $bookmarkDirectory -Prefix 'bookmark' -ByBookmark -Bookmark 'Chapter Two')
        $bookmarks.Count | Should -Be 1
        Read-OfficeDocument -As Text -Path $bookmarks[0].FullName | Should -Match 'Chapter Two'
        Read-OfficeDocument -As Text -Path $bookmarks[0].FullName | Should -Not -Match 'Chapter Three'
    }

    It 'splits encrypted PDFs and writes padded page names' {
        $path = Join-Path $TestDrive 'split-encrypted-source.pdf'
        New-OfficePdf -Path $path -Password 'open' {
            PdfParagraph 'Encrypted page one'
            PdfPageBreak
            PdfParagraph 'Encrypted page two'
            PdfPageBreak
            PdfParagraph 'Encrypted page three'
        } | Out-Null

        $outputDirectory = Join-Path $TestDrive 'encrypted-split'
        $outputs = @(Split-OfficeDocument -Path $path -Password 'open' -OutputDirectory $outputDirectory -Prefix 'page' -IndexWidth 3)

        $outputs.Count | Should -Be 3
        $outputs[0].Name | Should -Be 'page-001.pdf'
        $outputs[1].Name | Should -Be 'page-002.pdf'
        $outputs[2].Name | Should -Be 'page-003.pdf'
        foreach ($output in $outputs) {
            { Read-OfficeDocument -As Text -Path $output.FullName -ErrorAction Stop } | Should -Throw '*requires a password*'
            ((Get-OfficeDocumentInfo -Detailed -Path $output.FullName -Password 'open').Details).Security.HasEncryption | Should -BeTrue
        }
        Read-OfficeDocument -As Text -Path $outputs[1].FullName -Password 'open' | Should -Match 'Encrypted page two'
    }

    It 'does not replace an existing split destination' {
        $path = Join-Path $TestDrive 'split-collision-source.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'Collision-safe split output'
        } | Out-Null

        $outputDirectory = Join-Path $TestDrive 'split-collision-output'
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
        $existingPath = Join-Path $outputDirectory 'split-collision-source-1.pdf'
        [IO.File]::WriteAllText($existingPath, 'keep this file')

        $written = @(Split-OfficeDocument -Path $path -OutputDirectory $outputDirectory -NoClobber -ErrorAction SilentlyContinue -ErrorVariable splitErrors)

        $written.Count | Should -Be 0
        $splitErrors.Count | Should -Be 1
        [IO.File]::ReadAllText($existingPath) | Should -Be 'keep this file'

        $replaced = @(Split-OfficeDocument -Path $path -OutputDirectory $outputDirectory -ErrorAction Stop)
        $replaced.Count | Should -Be 1
        $replaced[0].FullName | Should -Be $existingPath
        Read-OfficeDocument -As Text -Path $existingPath | Should -Match 'Collision-safe split output'
    }

    It 'does not create split output directories when WhatIf skips writes' {
        $path = Join-Path $TestDrive 'split-whatif-source.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'Page one'
            PdfPageBreak
            PdfParagraph 'Page two'
        } | Out-Null

        $outputDirectory = Join-Path $TestDrive 'split-whatif-output'
        Split-OfficeDocument -Path $path -OutputDirectory $outputDirectory -WhatIf | Out-Null

        Test-Path -LiteralPath $outputDirectory | Should -BeFalse
    }

    It 'merges and resizes PDFs to fixed paper sizes' {

        $letter = Join-Path $TestDrive 'letter-source.pdf'
        $a5 = Join-Path $TestDrive 'a5-source.pdf'
        $merged = Join-Path $TestDrive 'merged-a4.pdf'

        New-OfficePdf -Path $letter {
            PdfPageSetup -PageSize Letter
            PdfParagraph 'Letter source text'
        } | Out-Null
        New-OfficePdf -Path $a5 {
            PdfPageSetup -PageSize A5
            PdfParagraph 'A5 source text'
        } | Out-Null

        Merge-OfficeDocument -Path $letter, $a5 -OutputPath $merged -Options (New-OfficePdfMergeOptions -PageSize A4 -ResizeMargin 12 -ResizeMode Fit) -PassThru |
            Should -BeOfType System.IO.FileInfo

        $info = (Get-OfficeDocumentInfo -Detailed -Path $merged).Details
        $info.PageCount | Should -Be 2
        foreach ($page in $info.Pages) {
            [math]::Round($page.Width) | Should -Be 595
            [math]::Round($page.Height) | Should -Be 842
        }

        Read-OfficeDocument -As Text -Path $merged | Should -Match 'Letter source text'
        Read-OfficeDocument -As Text -Path $merged | Should -Match 'A5 source text'
    }

    It 'resizes selected PDF pages' {
        $path = Join-Path $TestDrive 'resize-page-source.pdf'
        $outputPath = Join-Path $TestDrive 'resize-page-output.pdf'
        New-OfficePdf -Path $path {
            PdfPageSetup -PageSize A4
            PdfParagraph 'Resize first page'
            PdfPageBreak
            PdfParagraph 'Leave second page'
        } | Out-Null

        Set-OfficePdfPage -Path $path -OutputPath $outputPath -Page '1' -PageSize Letter -ResizeMargin 18 -PassThru |
            Should -BeOfType System.IO.FileInfo

        $info = (Get-OfficeDocumentInfo -Detailed -Path $outputPath).Details
        [math]::Round($info.Pages[0].Width) | Should -Be 612
        [math]::Round($info.Pages[0].Height) | Should -Be 792
        [math]::Round($info.Pages[1].Width) | Should -Be 595
        [math]::Round($info.Pages[1].Height) | Should -Be 842
        Read-OfficeDocument -As Text -Path $outputPath | Should -Match 'Resize first page'
        Read-OfficeDocument -As Text -Path $outputPath | Should -Match 'Leave second page'
    }

    It 'exposes diagnostics, optimization hints, structured text, catalog settings, and redaction planning' {
        $path = Join-Path $TestDrive 'diagnostic-source.pdf'
        New-OfficePdf -Path $path -CreateOutlineFromHeadings -IncludePageLabels -PageLabelPrefix 'P-' -DisplayDocTitle -FitWindow -OpenActionPage 1 -OpenActionMode Fit -PageMode UseOutlines {
            PdfMetadata -Title 'Diagnostic Report'
            PdfHeading 'Diagnostic Report'
            PdfParagraph 'This paragraph should appear in structured text and redaction planning.'
        } | Out-Null

        $info = (Get-OfficeDocumentInfo -Detailed -Path $path).Details
        $info.HasOutlines | Should -BeTrue
        $info.HasReadablePageLabels | Should -BeTrue
        $info.HasReadableOpenAction | Should -BeTrue
        $info.HasReadableViewerPreferences | Should -BeTrue

        $diagnostic = Get-OfficePdfDiagnostic -Path $path
        $diagnostic.CanRead | Should -BeTrue
        $diagnostic.ObjectGraphParsed | Should -BeTrue
        $diagnostic.StreamCount | Should -BeGreaterThan 0
        $diagnostic.StreamTypeCounts.Keys | Should -Contain 'Stream'
        $diagnostic.FontCount | Should -BeGreaterThan 0
        $diagnostic.Fonts[0].ObjectNumber | Should -BeGreaterThan 0

        $optimization = Get-OfficePdfDiagnostic -Type Optimization -Path $path
        $optimization.StreamCount | Should -Be $diagnostic.StreamCount
        $optimization.LargestStreams.Count | Should -BeGreaterThan 0
        $optimization.TotalStreamBytes | Should -BeGreaterThan 0
        $optimization.LargestStreamBytes | Should -BeGreaterThan 0
        $optimization.DuplicateStreamGroupCount | Should -BeGreaterOrEqual 0

        $blocks = @(& { $pdfReadDocument = Open-OfficePdf -Path $path; try { $pdfReadDocument.Read().TextBlocks } finally { $pdfReadDocument.Dispose() } })
        $blocks.Count | Should -BeGreaterThan 0
        $blocks.Text -join "`n" | Should -Match 'Diagnostic Report'

        $plan = Get-OfficePdfRedactionPlan -Path $path -Page 1 -X 0 -Y 0 -Width 1000 -Height 1000
        $plan.HasMatches | Should -BeTrue
        $plan.Matches.Text -join "`n" | Should -Match 'structured text'

        $flatPath = Join-Path $TestDrive 'diagnostic-source-flat-annotations.pdf'
        Update-OfficePdf -FlattenAnnotation -Path $path -OutputPath $flatPath | Should -BeOfType System.IO.FileInfo
        Test-Path $flatPath | Should -BeTrue
    }

    It 'updates PDF metadata as an incremental revision' {
        $path = Join-Path $TestDrive 'incremental-source.pdf'
        $incrementalPath = Join-Path $TestDrive 'incremental-output.pdf'
        New-OfficePdf -Path $path {
            PdfMetadata -Title 'Original title' -Author 'Original author'
            PdfParagraph 'Incremental body text'
        } | Out-Null

        Set-OfficePdfMetadata -Path $path -OutputPath $incrementalPath -Title 'Updated title' -Incremental -PassThru |
            Should -BeOfType System.IO.FileInfo

        (Get-Item $incrementalPath).Length | Should -BeGreaterThan (Get-Item $path).Length
        $info = (Get-OfficeDocumentInfo -Detailed -Path $incrementalPath).Details
        $info.Metadata.Title | Should -Be 'Updated title'
        $info.Metadata.Author | Should -Be 'Original author'
        $info.Security.HasIncrementalUpdates | Should -BeTrue
        $info.Security.RevisionCount | Should -BeGreaterThan 1
        Read-OfficeDocument -As Text -Path $incrementalPath | Should -Match 'Incremental body text'
    }

    It 'updates PDF form fields as an incremental revision' {
        (Get-Command Set-OfficePdfFormField).Parameters.Keys | Should -Contain 'Incremental'
        $path = Join-Path $TestDrive 'form-incremental-source.pdf'
        $outputPath = Join-Path $TestDrive 'form-incremental-output.pdf'
        @'
%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>
endobj
2 0 obj
<< /Type /Pages /Count 1 /Kids [3 0 R] >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [5 0 R] >>
endobj
4 0 obj
<< /Producer (PSWriteOffice form fixture) >>
endobj
5 0 obj
<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /V (Ada) /Rect [50 50 180 70] /F 4 >>
endobj
6 0 obj
<< /Fields [5 0 R] >>
endobj
trailer
<< /Root 1 0 R /Info 4 0 R /Size 7 >>
startxref
123
%%EOF
'@ | Set-Content -Path $path -NoNewline -Encoding Ascii

        Set-OfficePdfFormField -Path $path -OutputPath $outputPath -Value @{ Name = 'Grace' } -Incremental -PassThru |
            Should -BeOfType System.IO.FileInfo

        (Get-Item $outputPath).Length | Should -BeGreaterThan (Get-Item $path).Length
        $field = Get-OfficePdfFormField -Path $outputPath -Name Name
        $field.Value | Should -Be 'Grace'
        $info = (Get-OfficeDocumentInfo -Detailed -Path $outputPath).Details
        $info.Security.HasIncrementalUpdates | Should -BeTrue
        $info.AcroFormNeedAppearances | Should -BeFalse
        [IO.File]::ReadAllText($outputPath) | Should -Match '/AP'
        [IO.File]::ReadAllText($outputPath) | Should -Match '/Subtype /Form'
    }

    It 'can keep NeedAppearances for incremental PDF form fills' {
        $path = Join-Path $TestDrive 'form-incremental-legacy-source.pdf'
        $outputPath = Join-Path $TestDrive 'form-incremental-legacy-output.pdf'
        @'
%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>
endobj
2 0 obj
<< /Type /Pages /Count 1 /Kids [3 0 R] >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [5 0 R] >>
endobj
4 0 obj
<< /Producer (PSWriteOffice form fixture) >>
endobj
5 0 obj
<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /V (Ada) /Rect [50 50 180 70] /F 4 >>
endobj
6 0 obj
<< /Fields [5 0 R] >>
endobj
trailer
<< /Root 1 0 R /Info 4 0 R /Size 7 >>
startxref
123
%%EOF
'@ | Set-Content -Path $path -NoNewline -Encoding Ascii

        Set-OfficePdfFormField -Path $path -OutputPath $outputPath -Value @{ Name = 'Grace' } -Incremental -KeepNeedAppearances -PassThru |
            Should -BeOfType System.IO.FileInfo

        $field = Get-OfficePdfFormField -Path $outputPath -Name Name
        $field.Value | Should -Be 'Grace'
        $info = (Get-OfficeDocumentInfo -Detailed -Path $outputPath).Details
        $info.Security.HasIncrementalUpdates | Should -BeTrue
        $info.AcroFormNeedAppearances | Should -BeTrue
    }

    It 'updates DocMDP-certified PDF form fields when permissions allow form filling' {
        $path = Join-Path $TestDrive 'form-docmdp-source.pdf'
        $outputPath = Join-Path $TestDrive 'form-docmdp-output.pdf'
        @'
%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R /AcroForm 8 0 R /Perms << /DocMDP 7 0 R >> >>
endobj
2 0 obj
<< /Type /Pages /Count 1 /Kids [3 0 R] >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [5 0 R 6 0 R] /Contents 9 0 R >>
endobj
4 0 obj
<< /Producer (PSWriteOffice DocMDP form fixture) >>
endobj
5 0 obj
<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /V (Ada) /Rect [50 50 180 70] /F 4 >>
endobj
6 0 obj
<< /FT /Sig /T (Approval) /V 7 0 R /Subtype /Widget /Rect [10 10 120 40] >>
endobj
7 0 obj
<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name (Alice) /ByteRange [0 10 20 30] /Contents <001122> /Reference [<< /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /V /1.2 /P 2 >> >>] >>
endobj
8 0 obj
<< /Fields [5 0 R 6 0 R] /SigFlags 3 >>
endobj
9 0 obj
<< /Length 44 >>
stream
BT /F1 12 Tf 72 720 Td (Signed form) Tj ET
endstream
endobj
trailer
<< /Root 1 0 R /Info 4 0 R /Size 10 >>
startxref
123
%%EOF
'@ | Set-Content -Path $path -NoNewline -Encoding Ascii

        Set-OfficePdfFormField -Path $path -OutputPath $outputPath -Value @{ Name = 'Grace' } -Incremental -PassThru |
            Should -BeOfType System.IO.FileInfo

        $field = Get-OfficePdfFormField -Path $outputPath -Name Name
        $field.Value | Should -Be 'Grace'
        # This synthetic fixture needs explicit native recovery for its incorrect stream length.
        $preflight = (Get-OfficeDocumentInfo -Detailed -Path $outputPath -Options (New-OfficePdfReadOptions)).PdfPreflight
        $preflight.CanRead | Should -BeTrue
        $preflight.ReadBlockers | Should -BeNullOrEmpty
        $security = $preflight.Probe.Security
        $security.HasIncrementalUpdates | Should -BeTrue
        $security.HasSignatures | Should -BeTrue
        $security.HasDocMDPPermissions | Should -BeTrue
        [IO.File]::ReadAllText($outputPath) | Should -Match '/NeedAppearances\s+false'
    }

    It 'sets page production boundary boxes' {
        (Get-Command Set-OfficePdfPage).Parameters.Keys | Should -Contain 'BoxName'
        $path = Join-Path $TestDrive 'pagebox-source.pdf'
        $outputPath = Join-Path $TestDrive 'pagebox-output.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'Page box source'
        } | Out-Null

        Set-OfficePdfPage -Path $path -OutputPath $outputPath -BoxName TrimBox -BoxLeft 12 -BoxBottom 14 -BoxRight 222 -BoxTop 244 -PassThru |
            Should -BeOfType System.IO.FileInfo

        $geometry = ((Get-OfficeDocumentInfo -Detailed -Path $outputPath).Details).Pages[0].Geometry
        $geometry.TrimBox.Left | Should -Be 12
        $geometry.TrimBox.Bottom | Should -Be 14
        $geometry.TrimBox.Right | Should -Be 222
        $geometry.TrimBox.Top | Should -Be 244
    }

    It 'optimizes PDFs with lossless stream compression' {

        $path = Join-Path $TestDrive 'uncompressed-source.pdf'
        $optimizedPath = Join-Path $TestDrive 'uncompressed-optimized.pdf'
        $payload = 'A' * 4096
        $stream = "BT`n/F1 12 Tf`n72 720 Td`n($payload) Tj`nET`n"
        $pdf = @"
%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Count 1 /Kids [3 0 R] >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>
endobj
4 0 obj
<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>
endobj
5 0 obj
<< /Length $([Text.Encoding]::ASCII.GetByteCount($stream)) >>
stream
$($stream.TrimEnd("`n"))
endstream
endobj
trailer
<< /Root 1 0 R /Size 6 >>
startxref
123
%%EOF
"@
        [IO.File]::WriteAllBytes($path, [Text.Encoding]::ASCII.GetBytes($pdf))

        $report = Optimize-OfficePdf -Path $path -OutputPath $optimizedPath -PassThru

        $report.Applied | Should -BeTrue
        $report.ActionCount | Should -Be 1
        $report.ReportAfter | Should -Not -BeNullOrEmpty
        $report.CandidateSavedBytes | Should -BeGreaterThan 0
        $report.SkippedActionCount | Should -BeGreaterThan -1
        $report.Actions[0].Kind | Should -Be 'CompressStream'
        Test-Path $optimizedPath | Should -BeTrue
        (Get-Item $optimizedPath).Length | Should -BeLessThan (Get-Item $path).Length
        Read-OfficeDocument -As Text -Path $optimizedPath | Should -Match ('A' * 64)
    }

    It 'reports PDF signature structure and preservation markers' {
        $path = Join-Path $TestDrive 'signed-fixture.pdf'
        $pdf = @'
%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R /AcroForm 7 0 R /Perms << /DocMDP 6 0 R /UR3 6 0 R >> /DSS 9 0 R >>
endobj
2 0 obj
<< /Type /Pages /Count 1 /Kids [3 0 R] >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [5 0 R] >>
endobj
4 0 obj
<< /Length 0 >>
stream

endstream
endobj
5 0 obj
<< /FT /Sig /T (Approval) /V 6 0 R /Subtype /Widget /Rect [10 10 120 40] >>
endobj
6 0 obj
<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name (Alice) /ByteRange [0 10 20 30] /Contents <001122> /Reference [<< /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /V /1.2 /P 2 >> >>] >>
endobj
7 0 obj
<< /Fields [5 0 R] /SigFlags 3 >>
endobj
8 0 obj
<< /Producer (PSWriteOffice signed fixture) >>
endobj
9 0 obj
<< /Certs [10 0 R] /OCSPs [11 0 R] /CRLs [12 0 R] /VRI << /ABCDEF << /Cert [10 0 R] /OCSP [11 0 R] /CRL [12 0 R] /TS 13 0 R >> >> >>
endobj
10 0 obj
<< /Type /EmbeddedFile /Length 0 >>
endobj
11 0 obj
<< /Type /EmbeddedFile /Length 0 >>
endobj
12 0 obj
<< /Type /EmbeddedFile /Length 0 >>
endobj
13 0 obj
<< /Type /TimestampEvidence /Length 0 >>
endobj
trailer
<< /Root 1 0 R /Info 8 0 R /ID [(abc) (def)] /Size 14 /Prev 100 >>
startxref
100
%%EOF
startxref
200
%%EOF
'@
        [IO.File]::WriteAllBytes($path, [Text.Encoding]::ASCII.GetBytes($pdf))

        $report = Get-OfficePdfSignature -Path $path

        $report.HasSignatures | Should -BeTrue
        $report.SignatureCount | Should -Be 1
        $report.IsStructurallyValid | Should -BeFalse
        $report.RequiresAppendOnlyMutation | Should -BeTrue
        $report.HasLongTermValidationEvidence | Should -BeTrue
        $report.CryptographicTrustVerified | Should -BeFalse
        $report.DigestVerified | Should -BeFalse
        $report.CertificateChainVerified | Should -BeFalse
        $report.RevocationChecked | Should -BeFalse
        $report.TimestampValidationPerformed | Should -BeFalse
        $report.Signatures[0].Signature.FieldName | Should -Be 'Approval'
        $report.Signatures[0].Signature.ByteRangeValues -join ',' | Should -Be '0,10,20,30'
        $report.Signatures[0].Signature.HasRecognizedSubFilter | Should -BeTrue
        $report.Signatures[0].Signature.UsesDetachedCmsSubFilter | Should -BeTrue
        $report.Signatures[0].UnsignedByteCount | Should -BeGreaterThan 0
        $report.Signatures[0].ByteRangeCoverageRatio | Should -BeGreaterThan 0
        $report.Findings.Code | Should -Contain 'cryptographic-trust-not-verified'
        $report.Findings.Code | Should -Contain 'doc-mdp-detected'
        $report.Findings.Code | Should -Contain 'long-term-validation-evidence-detected'
        $report.Findings.Code | Should -Contain 'signature-detached-cms-sub-filter'
    }

    It 'prepares and injects external PDF signature bytes' {
        $path = Join-Path $TestDrive 'signature-source.pdf'
        $preparedPath = Join-Path $TestDrive 'signature-prepared.pdf'
        $signedPath = Join-Path $TestDrive 'signature-applied.pdf'
        $signaturePath = Join-Path $TestDrive 'signature.der'

        New-OfficePdf -Path $path {
            PdfParagraph 'External signing workflow'
        } | Out-Null

        $plan = New-OfficePdfSignature -Path $path -OutputPath $preparedPath -FieldName Approval -Name Alice -Reason Approval -ReservedSize 512 -PassThru
        $plan.FieldName | Should -Be 'Approval'
        $plan.ByteRangeValues.Count | Should -Be 4
        $plan.ComputeSha256Digest().Length | Should -Be 32
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($preparedPath)) |
            Should -Be ([Convert]::ToBase64String($plan.PreparedPdf))

        $preparedReport = Get-OfficePdfSignature -Path $preparedPath
        $preparedReport.HasSignatures | Should -BeTrue
        $preparedReport.Signatures[0].Signature.FieldName | Should -Be 'Approval'
        $preparedReport.Signatures[0].ByteRangeCoversEndOfFile | Should -BeTrue

        [IO.File]::WriteAllBytes($signaturePath, [byte[]](0x30, 0x82, 0x01, 0x0A, 0xAA, 0x55))
        $signedReport = Set-OfficePdfSignature -Path $preparedPath -SignaturePath $signaturePath -OutputPath $signedPath -PassThru
        $signedReport.HasSignatures | Should -BeTrue
        $signedReport.Signatures[0].Signature.ByteRangeValues -join ',' | Should -Be ($plan.ByteRangeValues -join ',')
        $signedReport.Findings.Code | Should -Contain 'signature-detached-cms-sub-filter'
        [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($signedPath)) | Should -Match '3082010AAA55'
    }

    It 'reports complex PDF text layout diagnostics' {
        $arabic = -join ([char[]](0x0645, 0x0631, 0x062D, 0x0628, 0x0627))
        $diagnostics = @(Get-OfficePdfDiagnostic -Text $arabic -AdvancedLayout)

        $diagnostics.Code | Should -Contain 'unsupported-bidirectional-text-layout'
        $diagnostics.Code | Should -Contain 'unsupported-complex-script-shaping'
    }

    It 'applies PDF redactions using planned text block coordinates' {
        $path = Join-Path $TestDrive 'redaction-source.pdf'
        $redactedPath = Join-Path $TestDrive 'redaction-output.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'Visible before'
            PdfParagraph 'Secret account 123-45'
            PdfParagraph 'Visible after'
        } | Out-Null

        $block = & { $pdfReadDocument = Open-OfficePdf -Path $path; try { $pdfReadDocument.Read().TextBlocks } finally { $pdfReadDocument.Dispose() } } |
            Where-Object { $_.Text -match 'Secret account' } |
            Select-Object -First 1
        $block | Should -Not -BeNullOrEmpty

        $x = [math]::Min($block.XStart, $block.XEnd) - 2
        $width = [math]::Abs($block.XEnd - $block.XStart) + 4
        # Keep the selected line's area inside its neighbors' glyph bounds.
        $y = $block.BaselineY - ($block.FontSize * 0.25)
        $height = $block.FontSize * 1.1

        Remove-OfficePdfContent -Path $path -OutputPath $redactedPath -Page $block.PageNumber -X $x -Y $y -Width $width -Height $height -FillColor '#111111' |
            Should -BeOfType System.IO.FileInfo

        $text = Read-OfficeDocument -As Text -Path $redactedPath
        $text | Should -Match 'Visible before'
        $text | Should -Match 'Visible after'
        $text | Should -Not -Match 'Secret account'
        $text | Should -Not -Match '123-45'
    }

    It 'plans and applies PDF redactions from literal text' {
        $path = Join-Path $TestDrive 'redaction-text-source.pdf'
        $redactedPath = Join-Path $TestDrive 'redaction-text-output.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'Visible before'
            PdfParagraph 'Secret account 123-45'
            PdfParagraph 'Visible after'
        } | Out-Null

        $plan = Get-OfficePdfRedactionPlan -Path $path -Text 'Secret account'
        $plan.HasMatches | Should -BeTrue
        $plan.Areas | Should -HaveCount 1

        Remove-OfficePdfContent -Path $path -OutputPath $redactedPath -Text 'Secret account' |
            Should -BeOfType System.IO.FileInfo

        $text = Read-OfficeDocument -As Text -Path $redactedPath
        $text | Should -Match 'Visible before'
        $text | Should -Match 'Visible after'
        $text | Should -Not -Match 'Secret account'
        $text | Should -Not -Match '123-45'
    }

    It 'does not write a PDF when literal redaction text is absent' {
        $path = Join-Path $TestDrive 'redaction-missing-source.pdf'
        $redactedPath = Join-Path $TestDrive 'redaction-missing-output.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'Visible content only'
        } | Out-Null

        {
            Remove-OfficePdfContent -Path $path -OutputPath $redactedPath `
                -Text 'Absent secret' -ErrorAction Stop
        } | Should -Throw '*No PDF text matched*output was not written*'

        Test-Path -LiteralPath $redactedPath | Should -BeFalse
    }

    It 'creates PDFs with document-level font options' {
        $path = Join-Path $TestDrive 'font-options.pdf'
        New-OfficePdf -Path $path -Font Courier -Size 13pt {
            PdfHeading 'Font Options'
            PdfParagraph 'Generated with document-level font settings.'
        } | Out-Null

        (Get-OfficePdfDiagnostic -Type Preflight -Path $path).CanRead | Should -BeTrue
        Read-OfficeDocument -As Text -Path $path | Should -Match 'Font Options'
        $body = @(& { $pdfReadDocument = Open-OfficePdf -Path $path; try { $pdfReadDocument.Read().TextBlocks } finally { $pdfReadDocument.Dispose() } } | Where-Object Text -Match 'Generated with document-level')
        $body.Count | Should -BeGreaterThan 0
        $body | ForEach-Object { $_.FontSize | Should -Be 13 }
    }

    It 'builds visually polished PDF documents' {
        $path = Join-Path $TestDrive 'visual-polish.pdf'
        New-OfficePdf -Path $path {
            PdfBackground -Color '#F8FAFC'
            PdfPageBorder -Color '#0F766E' -Width 1.5 -Inset 24 -Opacity 0.85
            PdfBookmark 'summary'
            PdfHeading 'Visual Summary'
            PdfHorizontalRule -Color '#0F766E' -Thickness 2 -SpacingBefore 8 -SpacingAfter 10
            PdfPanel 'A polished report surface should support visual structure.'
        } | Out-Null

        $info = (Get-OfficeDocumentInfo -Detailed -Path $path).Details
        $info.PageCount | Should -Be 1
        $info.HasNamedDestinations | Should -BeTrue
        $info.NamedDestinationNames | Should -Contain 'summary'

        (Get-OfficePdfDiagnostic -Type Preflight -Path $path).CanRead | Should -BeTrue
        Read-OfficeDocument -As Text -Path $path | Should -Match 'Visual Summary'
    }

    It 'builds two-column PDF layout rows' {
        $path = Join-Path $TestDrive 'layout-row.pdf'
        $rows = @(
            [pscustomobject]@{ Area = 'Coverage'; Status = 'Good' }
            [pscustomobject]@{ Area = 'Readback'; Status = 'Verified' }
        )

        New-OfficePdf -Path $path {
            PdfHeading 'Layout Report'
            PdfSpacer 10
            PdfRow -Gap 16 -SpacingBefore 4 -SpacingAfter 8 -ColumnSeparatorColor '#CBD5E1' {
                PdfColumn -Weight 38 {
                    PdfHeading 'Highlights' -Level 2 -Color '#0F766E'
                    PdfParagraph 'Left column summary.'
                    PdfList -Item 'Designed surface', 'Readable rhythm' -Numbered
                }
                PdfColumn -Weight 62 {
                    PdfBookmark 'layout-row'
                    PdfPanel 'Right column panel.'
                    PdfTable -InputObject $rows -TableStyle Compact -Caption 'Layout row table' -NoBorder
                    PdfHorizontalRule -Color '#0F766E' -Thickness 1.2
                }
            }
            PdfParagraph 'After layout row.'
        } | Out-Null

        $info = (Get-OfficeDocumentInfo -Detailed -Path $path).Details
        $info.PageCount | Should -Be 1
        $info.NamedDestinationNames | Should -Contain 'layout-row'

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Highlights'
        $text | Should -Match 'Right column panel'
        $text | Should -Match 'Layout row table'
        $text | Should -Match 'After layout row'
    }

    It 'renders typed list content directly inside PDF row columns' {
        $path = Join-Path $TestDrive 'layout-row-list.pdf'
        New-OfficePdf -Path $path {
            PdfRow { PdfColumn -Weight 100 { PdfList -Item 'Alpha', 'Beta' } }
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Alpha'
        $text | Should -Match 'Beta'
    }

    It 'renders a single paragraph content block inside PDF row columns' {
        $path = Join-Path $TestDrive 'layout-row-single-content.pdf'
        New-OfficePdf -Path $path {
            PdfRow { PdfColumn -Weight 100 { PdfParagraph 'Single content block' } }
        } | Out-Null

        Read-OfficeDocument -As Text -Path $path | Should -Match 'Single content block'
    }

    It 'renders a single table object inside PDF row columns' {
        $path = Join-Path $TestDrive 'layout-row-single-table-object.pdf'
        New-OfficePdf -Path $path {
            PdfRow {
                PdfColumn -Weight 100 { PdfTable -InputObject @{ Name = 'Alpha'; Value = 1 } -Header 'Name', 'Value' }
            }
        } | Out-Null

        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Name'
        $text | Should -Match 'Value'
        $text | Should -Match 'Alpha'
        $text | Should -Match '1'
    }

    It 'renders rich text inside native PDF row columns' {
        $path = Join-Path $TestDrive 'layout-row-rich-run.pdf'
        New-OfficePdf -Path $path {
            PdfRow { PdfColumn -Width 100 { PdfParagraph { PdfText 'Approved' -Bold } } }
        }
        Read-OfficeDocument -As Text -Path $path | Should -Match 'Approved'
    }

    It 'builds rich PDF text with emphasis and links' {
        $path = Join-Path $TestDrive 'rich-text.pdf'
        New-OfficePdf -Path $path {
            PdfBookmark 'details'
            PdfParagraph {
                PdfText 'This paragraph contains '
                PdfText 'bold' -Bold -Color '#0F766E'
                PdfText ', '
                PdfText 'italic' -Italic
                PdfText ', highlighted ' -BackgroundColor '#FEF3C7'
                PdfText 'external link' -Link 'https://evotec.xyz' -Color '#2563EB'
                PdfText ', and '
                PdfText 'bookmark link' -LinkBookmark details -Color '#7C3AED'
                PdfText '.'
            }
            PdfRow {
                PdfColumn -Width 100 {
                    PdfParagraph {
                        PdfText 'Row layout can also contain '
                        PdfText 'rich inline text' -Bold -LinkBookmark details -Color '#0F766E'
                        PdfText '.'
                    }
                }
            }
        }
        $info = (Get-OfficeDocumentInfo -Detailed -Path $path).Details
        $info.HasLinkAnnotations | Should -BeTrue
        $info.LinkUris | Should -Contain 'https://evotec.xyz'
        $info.LinkDestinationNames | Should -Contain 'details'
        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'bold'
        $text | Should -Match 'bookmark link'
        $text | Should -Match 'rich\s+inline\s+text'
    }

    It 'renders native rich-text baselines and tab leaders' {
        $path = Join-Path $TestDrive 'rich-text-tabs.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph {
                PdfText 'Total' -Baseline Normal
                PdfText "`t" -TabLeader Hyphens -TabAlignment DecimalSeparator
                PdfText '12.34'
                PdfText "`t" -TabLeader Underscores -TabAlignment Right
                PdfText 'Done'
            }
        }
        Read-OfficeDocument -As Text -Path $path | Should -Match 'Total'
        Read-OfficeDocument -As Text -Path $path | Should -Match '12.34'
        Read-OfficeDocument -As Text -Path $path | Should -Match 'Done'
    }


    It 'applies PDF themes and decorative backgrounds' {
        $path = Join-Path $TestDrive 'styled-backgrounds.pdf'
        $imagePath = Join-Path (Join-Path $PSScriptRoot 'Assets') 'CellImage.png'

        New-OfficePdf -Path $path -Theme Modern {
            PdfBackground -Color '#FFFFFF'
            PdfBackground -ImagePath $imagePath -Fit Cover -Opacity 0.03
            PdfBackgroundShape -Shape TopBand -Height 86 -FillColor '#DBEAFE' -FillOpacity 0.75
            PdfBackgroundShape -Shape Ellipse -X 420 -Y 650 -Width 96 -Height 72 -FillColor '#99F6E4' -FillOpacity 0.35
            PdfHeading 'Styled Report'
            PdfParagraph 'Theme and decorative backgrounds are OfficeIMO-owned PDF layout features.'
            PdfPanel 'A reusable theme keeps report rhythm consistent.'
        } | Out-Null

        (Get-OfficePdfDiagnostic -Type Preflight -Path $path).CanRead | Should -BeTrue
        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Styled Report'
        $text | Should -Match 'reusable theme'
    }
}

Describe 'PDF native physical and pipeline contracts' {
    It 'writes physical rule and spacer lengths with typed colors and styled text' {
        $path = Join-Path $TestDrive 'physical-authoring.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph { PdfText 'Physical styled marker' -Bold -Italic -Underline Single -Size '14pt' -Font Helvetica }
            PdfHorizontalRule -Thickness '1mm' -SpacingBefore '6pt' -SpacingAfter '0.25in' -Color ([OfficeIMO.OfficeColor]::Red)
            PdfSpacer '0.5in'
            PdfParagraph 'After physical spacer'
        } | Out-Null
        $document = Open-OfficePdf -LiteralPath $path
        try {
            $drawing = $document.Render.Drawing(1)
            $line = @($drawing.Elements | Where-Object {
                $_ -is [OfficeIMO.Drawing.OfficeDrawingShape] -and $_.Shape.Kind -eq 'Line' -and
                $_.Shape.StrokeColor -eq [OfficeIMO.OfficeColor]::Red
            }) | Select-Object -First 1
            $line | Should -Not -BeNullOrEmpty
            [math]::Abs($line.Shape.StrokeWidth - 72 / 25.4) | Should -BeLessThan 0.002
            $line.Shape.StrokeColor | Should -Be ([OfficeIMO.OfficeColor]::Red)
            # Positioned PDF text can be split into glyph runs, with nonpainting spaces omitted.
            $styled = @($drawing.Elements | Where-Object {
                $_ -is [OfficeIMO.Drawing.OfficeDrawingText] -and $_.Y -lt $line.Y
            } | Sort-Object X)
            $styled | Should -Not -BeNullOrEmpty
            (($styled.Text -join '') -replace '\s', '') | Should -BeExactly 'Physicalstyledmarker'
            foreach ($fragment in $styled) {
                [math]::Abs($fragment.Font.Size - 14) | Should -BeLessThan 0.01
                $fragment.Font.Style.ToString() | Should -Match 'Bold'
                $fragment.Font.Style.ToString() | Should -Match 'Italic'
            }
            $text = $document.Read().Text
            $text | Should -Match 'Physical styled marker'
            $text | Should -Match 'After physical spacer'
        } finally { $document.Dispose() }
    }

    It 'collects piped chart data once and accepts native document pipelines' {
        $path = Join-Path $TestDrive 'pipeline-authoring.pdf'
        $rows = @([pscustomobject]@{ Name = 'North'; Score = 12 }, [pscustomobject]@{ Name = 'South'; Score = 25 })
        $document = New-OfficePdf
        try {
            $rows | Add-OfficePdfChart -Document $document -Category Name -Value Score -Title 'One piped chart'
            $document | Add-OfficePdfList -Item 'Piped list marker' | Out-Null
            $document | Add-OfficePdfRow { PdfColumn { PdfParagraph 'Piped row marker' } } | Out-Null
            $document | Add-OfficePdfFormField -Name PipelineField -Value Before | Out-Null
            $filled = $document | Set-OfficePdfFormField -Value @{ PipelineField = 'After' } -PassThru
            try {
                $filled | Save-OfficePdf -Path $path | Out-Null
            } finally { if (-not [object]::ReferenceEquals($filled, $document)) { $filled.Dispose() } }
        } finally { $document.Dispose() }
        $text = Read-OfficeDocument -Path $path -As Text
        ([regex]::Matches($text, 'One piped chart')).Count | Should -Be 1
        $text | Should -Match 'North'
        $text | Should -Match 'South'
        $text | Should -Match 'Piped list marker'
        $text | Should -Match 'Piped row marker'
        (Get-OfficePdfFormField -Path $path -Name PipelineField).Value | Should -Be 'After'
    }

    It 'binds literal FileInfo inputs and honors WhatIf before writing for <Operation>' -TestCases @(
        @{ Operation = 'optimize' }
        @{ Operation = 'flatten' }
        @{ Operation = 'fill' }
        @{ Operation = 'redact' }
    ) {
        param($Operation)
        $source = Join-Path $TestDrive ('literal-' + $Operation + '.pdf')
        $destination = Join-Path $TestDrive ('literal-' + $Operation + '-output')
        New-OfficePdf -Path $source { PdfParagraph 'Literal removal marker'; PdfFormField -Name Name -Value Before } | Out-Null
        $before = [Convert]::ToBase64String([IO.File]::ReadAllBytes($source))
        $file = Get-Item -LiteralPath $source
        switch ($Operation) {
            'optimize' { $file | Optimize-OfficePdf -OutputDirectory $destination -WhatIf | Out-Null }
            'flatten' { $file | Update-OfficePdf -FlattenForm -OutputDirectory $destination -WhatIf | Out-Null }
            'fill' { $file | Set-OfficePdfFormField -OutputDirectory $destination -Value @{ Name = 'After' } -WhatIf | Out-Null }
            'redact' { $file | Remove-OfficePdfContent -OutputDirectory $destination -Text 'Literal removal marker' -WhatIf | Out-Null }
        }
        Test-Path -LiteralPath $destination | Should -BeFalse
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($source)) | Should -Be $before
    }
}

Describe 'PDF printer submission' -Skip:($PSVersionTable.PSVersion -lt [version]'7.4') {
    It 'rejects a password for non-PDF input before submitting any document' {
        $pdf = Join-Path $TestDrive 'first.pdf'
        New-OfficePdf -Path $pdf { PdfParagraph 'First' } | Out-Null
        $docx = Join-Path $TestDrive 'second.docx'
        Set-Content -LiteralPath $docx -Value 'not opened'
        # The PDF comes first; validation must fail before any document is considered for submission.
        { Send-OfficePrinter -Path $pdf, $docx -PrinterName 'officeimo-unused-queue' -Password (ConvertTo-SecureString 'secret' -AsPlainText -Force) -WhatIf } |
            Should -Throw '*Password applies only to PDF documents*'
    }

    It 'binds the 3.5 parameter names of Send-OfficePdfPrinter' {
        $pdf = Join-Path $TestDrive 'aliases.pdf'
        New-OfficePdf -Path $pdf { PdfParagraph 'Aliases' } | Out-Null
        { Send-OfficePdfPrinter -Path $pdf -PrinterName 'officeimo-unused-queue' -Pages '1' -OutputFilePath (Join-Path $TestDrive 'out.prn') -WhatIf } | Should -Not -Throw
    }
}
