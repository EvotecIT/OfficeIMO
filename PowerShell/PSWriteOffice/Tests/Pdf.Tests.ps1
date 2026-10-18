BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'PDF cmdlets' {
    It 'returns the canonical structured PDF model for paths and document pipelines' {
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

        $logical = Read-OfficePdf -Path $path -Profile Structured
        $logical.GetType().FullName | Should -Be 'OfficeIMO.Pdf.PdfDocumentReadResult'
        @($logical.Pages).Count | Should -BeGreaterThan 0
        $logical.Text | Should -Match 'Canonical paragraph marker'
        @($logical.Tables).Count | Should -BeGreaterThan 0
        @($logical.Tables.Rows).Count | Should -BeGreaterThan 0
        @($logical.Images).Count | Should -BeGreaterThan 0

        $images = @(Get-OfficePdfImage -Path $path)
        $images.Count | Should -BeGreaterThan 0
        $images[0].GetType().FullName | Should -Be 'OfficeIMO.Pdf.PdfExtractedImage'

        $fromPipeline = Get-OfficePdf -Path $path | Read-OfficePdf -PageRange '1'
        $fromPipeline.GetType().FullName | Should -Be 'OfficeIMO.Pdf.PdfDocumentReadResult'
        @($fromPipeline.Pages).Count | Should -Be 1
        $fromPipeline.Text | Should -Match 'Canonical paragraph marker'
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

    It 'rejects incompatible text modes before opening the PDF' {
        $missingPath = Join-Path $TestDrive 'missing-text-source.pdf'

        { Get-OfficePdfText -Path $missingPath -AsTextBlock -AsMarkdown -ErrorAction Stop } |
            Should -Throw '*-AsTextBlock cannot be combined*'
        { Get-OfficePdfText -Path $missingPath -AsTextBlock -ByPage -ErrorAction Stop } |
            Should -Throw '*-AsTextBlock cannot be combined*'
        { Get-OfficePdfText -Path $missingPath -ByPage -AsMarkdown -ErrorAction Stop } |
            Should -Throw '*-ByPage is supported for plain text extraction only*'
    }

    It 'validates PDF operations and page selections before path I/O' {
        $missingPath = Join-Path $TestDrive 'missing-source.pdf'
        $outputPath = Join-Path $TestDrive 'validation-output.pdf'
        $exportDirectory = Join-Path $TestDrive 'validation-export'
        $joinDirectory = Join-Path $TestDrive 'validation-join'
        $joinPath = Join-Path $joinDirectory 'joined.pdf'

        { Split-OfficePdf -Path $missingPath -OutputDirectory $exportDirectory -PagesPerDocument -1 -ErrorAction Stop } |
            Should -Throw '*-PagesPerDocument must be greater than zero*'
        { Split-OfficePdf -Path $missingPath -OutputDirectory $exportDirectory -PageRange 'invalid' -ErrorAction Stop } |
            Should -Throw '*page*'
        { Set-OfficePdfForm -Path $missingPath -OutputPath $outputPath -ErrorAction Stop } |
            Should -Throw '*Provide -Field values or use -Flatten*'
        { Set-OfficePdfPage -Path $missingPath -OutputPath $outputPath -ErrorAction Stop } |
            Should -Throw '*Provide -Rotation, -BoxName with coordinates, or page resize options*'
        { Compare-OfficePdfVisual -ReferencePath $missingPath -DifferencePath $missingPath -PageRange 'invalid' -ErrorAction Stop } |
            Should -Throw '*page*'
        { ConvertTo-OfficePdfMarkdown -Path $missingPath -PageRange 'invalid' -ErrorAction Stop } |
            Should -Throw '*page*'
        { Export-OfficePdfImage -Path $missingPath -OutputPath $exportDirectory -PageRange 'invalid' -ErrorAction Stop } |
            Should -Throw '*page*'
        { Join-OfficePdf -Path $missingPath, $missingPath -Password 'one', 'two', 'three' -OutputPath $joinPath -ErrorAction Stop } |
            Should -Throw '*one -Password value or one value for every -Path source*'
        { ConvertTo-OfficePdfHtml -Path $missingPath -PageRange 'invalid' -ErrorAction Stop } |
            Should -Throw '*page*'

        Test-Path -LiteralPath $exportDirectory | Should -BeFalse
        Test-Path -LiteralPath $joinDirectory | Should -BeFalse
    }

    It 'preserves advanced PDF read options unless a friendly parameter is supplied' {
        $path = Join-Path $TestDrive 'structured-read-options.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'First page marker'
            PdfPageBreak
            PdfParagraph 'Second page marker'
        } | Out-Null

        $command = Get-Command Read-OfficePdf
        $optionsType = $command.Parameters['Options'].ParameterType
        $profileType = $command.Parameters['Profile'].ParameterType
        $pageSelectionType = $optionsType.GetProperty('PageSelection').PropertyType
        $options = [Activator]::CreateInstance($optionsType)
        $options.Profile = [Enum]::Parse($profileType, 'Fast')
        $options.PageSelection = $pageSelectionType.GetMethod('Parse', [type[]]@([string])).Invoke($null, @('2'))

        $fromOptions = Read-OfficePdf -Path $path -Options $options
        $fromOptions.Profile.ToString() | Should -Be 'Fast'
        @($fromOptions.Pages).Count | Should -Be 1
        $fromOptions.Pages[0].PageNumber | Should -Be 2

        $overridden = Read-OfficePdf -Path $path -Options $options -Profile Structured -PageRange '1'
        $overridden.Profile.ToString() | Should -Be 'Structured'
        @($overridden.Pages).Count | Should -Be 1
        $overridden.Pages[0].PageNumber | Should -Be 1
    }

    It 'supports passwords for encrypted PDF read cmdlets' {
        $encryptedPath = Join-Path $TestDrive 'encrypted.pdf'
        [IO.File]::WriteAllBytes($encryptedPath, [Convert]::FromBase64String('JVBERi0xLjQKJT8/Pz8KMSAwIG9iago8PCAvVHlwZSAvQ2F0YWxvZyAvUGFnZXMgMiAwIFIgPj4KZW5kb2JqCjIgMCBvYmoKPDwgL1R5cGUgL1BhZ2VzIC9LaWRzIFszIDAgUl0gL0NvdW50IDEgPj4KZW5kb2JqCjMgMCBvYmoKPDwgL1R5cGUgL1BhZ2UgL1BhcmVudCAyIDAgUiAvTWVkaWFCb3ggWzAgMCAzMDAgMjAwXSAvUmVzb3VyY2VzIDw8IC9Gb250IDw8IC9GMSA0IDAgUiA+PiA+PiAvQ29udGVudHMgNSAwIFIgPj4KZW5kb2JqCjQgMCBvYmoKPDwgL1R5cGUgL0ZvbnQgL1N1YnR5cGUgL1R5cGUxIC9CYXNlRm9udCAvSGVsdmV0aWNhID4+CmVuZG9iago1IDAgb2JqCjw8IC9MZW5ndGggNDYgPj4Kc3RyZWFtCi8xB1rv33VvGaaV2g01B7caayy3ttoqyqa6Fkx+aapdiBLgquqJCxhp8zWpAXQKZW5kc3RyZWFtCmVuZG9iago2IDAgb2JqCjw8IC9GaWx0ZXIgL1N0YW5kYXJkIC9WIDEgL1IgMiAvTGVuZ3RoIDQwIC9PIDw4RUVCMDk1ODE5NjYyQTc3NDQ0MkZCMDcyRTNEOUYxOUU5RDEzMEVDMDlBNEQwMDYxRTc4RkU5MjBGN0FCNjJGPiAvVSA8QjFFNzY0MTI2QzQ4RDI4RDkwNTI1NTk1MjAwREQ4MTg3NEI4NkZFMUNBNTRCQTAxODZFNThCRTJDMzU5ODhEQz4gL1AgLTQgPj4KZW5kb2JqCnhyZWYKMCA3CjAwMDAwMDAwMDAgNjU1MzUgZiAKMDAwMDAwMDAxNSAwMDAwMCBuIAowMDAwMDAwMDY0IDAwMDAwIG4gCjAwMDAwMDAxMjEgMDAwMDAgbiAKMDAwMDAwMDI0NyAwMDAwMCBuIAowMDAwMDAwMzE3IDAwMDAwIG4gCjAwMDAwMDA0MTMgMDAwMDAgbiAKdHJhaWxlcgo8PCAvU2l6ZSA3IC9Sb290IDEgMCBSIC9FbmNyeXB0IDYgMCBSIC9JRCBbPDEwNDVBODdDMjIxODRFQzE5MTRBQ0Y2NjMxRDI3NDAzPiA8MTA0NUE4N0MyMjE4NEVDMTkxNEFDRjY2MzFEMjc0MDM+XSA+PgpzdGFydHhyZWYKNjE5CiUlRU9GCg=='))

        foreach ($name in 'Get-OfficePdf', 'Get-OfficePdfInfo', 'Get-OfficePdfPreflight', 'Get-OfficePdfDiagnostic', 'Get-OfficePdfOptimization', 'Get-OfficePdfSignature', 'Get-OfficePdfText', 'Get-OfficePdfAttachment', 'Get-OfficePdfFormField', 'Get-OfficePdfRedactionPlan', 'ConvertTo-OfficePdfMarkdown', 'ConvertTo-OfficePdfHtml', 'ConvertTo-OfficePdfRedacted') {
            (Get-Command $name).Parameters.Keys | Should -Contain 'Password'
        }

        (Get-OfficePdfPreflight -Path $encryptedPath).CanRead | Should -BeFalse
        (Get-OfficePdfPreflight -Path $encryptedPath -Password 'open').CanRead | Should -BeTrue
        Get-OfficePdfText -Path $encryptedPath -Password 'open' | Should -Match 'Secret PDF Text'
        (Get-OfficePdf -Path $encryptedPath -Password 'open').Read().Text | Should -Match 'Secret PDF Text'
        ConvertTo-OfficePdfMarkdown -Path $encryptedPath -Password 'open' | Should -Match 'Secret PDF Text'
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

        (Get-OfficePdfPreflight -Path $newPath).CanRead | Should -BeFalse
        (Get-OfficePdfPreflight -Path $newPath -Password 'open').CanRead | Should -BeTrue
        Get-OfficePdfText -Path $newPath -Password 'open' | Should -Match 'Generated encrypted PDF text'
        { Get-OfficePdfText -Path $newPath -Password 'wrong' } | Should -Throw

        $savedPath = Join-Path $TestDrive 'saved-encrypted.pdf'
        $doc = New-OfficePdf {
            PdfParagraph 'Saved encrypted PDF text'
        }

        $doc | Save-OfficePdf -Path $savedPath -Password 'open' -OwnerPassword 'owner' | Out-Null

        (Get-OfficePdfPreflight -Path $savedPath).CanRead | Should -BeFalse
        Get-OfficePdfText -Path $savedPath -Password 'open' | Should -Match 'Saved encrypted PDF text'
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
        $document = New-OfficePdf {
            PdfParagraph 'Existing document content'
        }

        { $document | PdfPageSetup -Margin 36 } |
            Should -Throw '*cannot update an already-built PdfDocument*'
        { New-OfficePdf { PdfParagraph 'First' -PassThru | PdfParagraph 'Second' } } |
            Should -Throw '*cannot emit a PdfDocument before New-OfficePdf finishes composition*'
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

        $info = Get-OfficePdfInfo -Path $path
        $info.PageCount | Should -BeGreaterOrEqual 1

        $preflight = Get-OfficePdfPreflight -Path $path
        $preflight.CanRead | Should -BeTrue

        $text = Get-OfficePdfText -Path $path
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

        $text = Get-OfficePdfText -Path $path
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

        $text = Get-OfficePdfText -Path $path
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

        $text = Get-OfficePdfText -Path $path
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

        (Get-OfficePdfPreflight -Path $path).CanRead | Should -BeTrue
        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Identity systems'
        $text | Should -Match 'Shared owner'
        $text | Should -Match 'Release'
    }

    It 'styles merged PDF table cells' {
        $parameters = (Get-Command New-OfficePdfTableCell).Parameters.Keys
        $parameters | Should -Contain 'TextColor'
        $parameters | Should -Contain 'FillColor'
        $parameters | Should -Contain 'FontSize'
        $parameters | Should -Contain 'Bold'
        $parameters | Should -Contain 'Italic'
        $parameters | Should -Contain 'Align'
        $parameters | Should -Contain 'VerticalAlign'

        $path = Join-Path $TestDrive 'styled-span-aware-table.pdf'

        New-OfficePdf -Path $path {
            PdfTable -HeaderRowCount 1 -InputObject @(
                @('Service', 'Status', 'Owner'),
                @(New-OfficePdfTableCell -Text 'Identity systems' -ColumnSpan 3 -FillColor '#DBEAFE' -TextColor '#1E3A8A' -FontSize 17 -Bold -Italic -Align Center -VerticalAlign Middle),
                @('Entra', 'Watch', 'IAM'),
                @(@{ Text = 'Follow-up'; ColumnSpan = 3; FillColor = '#FEF3C7'; Color = '#92400E'; FontSize = 15; Bold = $true; Align = 'Center' }),
                @('Release', 'Ready', 'PSWriteOffice')
            )
        } | Out-Null

        (Get-OfficePdfPreflight -Path $path).CanRead | Should -BeTrue
        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Identity systems'
        $text | Should -Match 'Follow-up'

        $blocks = @(Get-OfficePdfText -Path $path -AsTextBlock)
        $identityBlock = $blocks | Where-Object { $_.Text -match 'Identity systems' } | Select-Object -First 1
        $identityBlock | Should -Not -BeNullOrEmpty
        $identityBlock.FontSize | Should -Be 17

        $followUpBlock = $blocks | Where-Object { $_.Text -match 'Follow-up' } | Select-Object -First 1
        $followUpBlock | Should -Not -BeNullOrEmpty
        $followUpBlock.FontSize | Should -Be 15
    }

    It 'renders rich text runs and named colors in PDF text and table cells' {
        foreach ($name in 'PdfNew', 'TextRun', 'PdfTextRun', 'PdfTableCell') {
            Get-Command $name | Should -Not -BeNullOrEmpty
        }

        $path = Join-Path $TestDrive 'rich-run-table.pdf'

        PdfNew -Path $path {
            PdfText -Run @{
                Text  = 'Status: ', 'Ready'
                Color = $null, 'SeaGreen'
                Bold  = $false, $true
            }
            PdfTable -InputObject @(
                , @(
                    PdfTableCell -Run @(
                        PdfTextRun 'Build '
                        PdfTextRun 'Ready' -Color SeaGreen -BackgroundColor None -Bold
                    ) -ColumnSpan 2 -FillColor AliceBlue
                )
                , @(
                    PdfTableCell -Run @(
                        PdfTextRun 'Styled'
                    ) -Bold -TextColor Crimson -FontSize 19 -ColumnSpan 2
                )
                , @(
                    @{ Run = 'Queued' },
                    'String run'
                )
                , @('Owner', 'Platform')
            )
        } | Out-Null

        (Get-OfficePdfPreflight -Path $path).CanRead | Should -BeTrue
        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Status'
        $text | Should -Match 'Ready'
        $text | Should -Match 'Build'
        $text | Should -Match 'Styled'
        $text | Should -Match 'Queued'
        $text | Should -Match 'Owner'

        $blocks = @(Get-OfficePdfText -Path $path -AsTextBlock)
        $styledBlock = $blocks | Where-Object { $_.Text -match 'Styled' } | Select-Object -First 1
        $styledBlock | Should -Not -BeNullOrEmpty
        $styledBlock.FontSize | Should -Be 19
    }

    It 'rejects PDF table cell rich text links until table annotations are supported' {
        $path = Join-Path $TestDrive 'rich-run-table-link.pdf'

        {
            PdfNew -Path $path {
                PdfTable -InputObject @(
                    , @(
                        PdfTableCell -Run @(
                            PdfTextRun 'Docs' -LinkUri 'https://example.org/table'
                        )
                    )
                )
            }
        } | Should -Throw '*PDF table cell text runs do not support links yet*'
    }

    It 'treats rich and styled PDF table cells as structured without spans' {
        $path = Join-Path $TestDrive 'rich-run-mixed-object-table.pdf'

        PdfNew -Path $path {
            PdfTable -InputObject @(
                [pscustomobject]@{
                    Service = 'Entra'
                    Status  = 'Ready'
                }
                , @(
                    (PdfTableCell -Run @(
                        PdfTextRun 'Build '
                        PdfTextRun 'Watch' -Color DarkOrange -Bold
                    ) -FillColor AliceBlue),
                    'Platform'
                )
            )
        } | Out-Null

        (Get-OfficePdfPreflight -Path $path).CanRead | Should -BeTrue
        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Service'
        $text | Should -Match 'Entra'
        $text | Should -Match 'Build'
        $text | Should -Match 'Watch'
        $text | Should -Match 'Platform'
    }

    It 'keeps ordinary style-named columns in PDF object tables' {
        $path = Join-Path $TestDrive 'ordinary-style-named-table.pdf'

        PdfNew -Path $path {
            PdfTable -InputObject @(
                [pscustomobject]@{
                    Text     = 'Task'
                    Color    = 'Red'
                    FontSize = 'Large'
                }
            )
        } | Out-Null

        $text = Get-OfficePdfText -Path $path
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

        $text = Get-OfficePdfText -Path $path
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

        $text = Get-OfficePdfText -Path $path
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

        $text = Get-OfficePdfText -Path $path
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

        $text = Get-OfficePdfText -Path $path
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

        $text = Get-OfficePdfText -Path $path
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

        $text = Get-OfficePdfText -Path $path
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
            PdfTable -InputObject $rows -Property Service, Status, Incidents -TableStyle Report -Caption 'Service status' -CaptionAlign Center -AutoFitColumns -RightAlignNumeric -ColumnWidthWeights 2, 1, 1
        } | Out-Null

        (Get-OfficePdfPreflight -Path $path).CanRead | Should -BeTrue
        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Service status'
        $text | Should -Match 'Directory'
        $text | Should -Match 'Healthy'
    }

    It 'shrinks PDF table text to fit fixed-width cells' {
        (Get-Command PdfTable).Parameters.Keys | Should -Contain 'ShrinkTextToFit'
        (Get-Command PdfTable).Parameters.Keys | Should -Contain 'MinimumShrinkFontSize'

        $path = Join-Path $TestDrive 'shrink-table.pdf'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 'ThisIdentifierShouldShrinkToFit' }
        )

        New-OfficePdf -Path $path {
            PdfTable -InputObject $rows -Property Name, Value -ColumnWidthPoints 54, 108 -FontSize 18 -ShrinkTextToFit -MinimumShrinkFontSize 7
        } | Out-Null

        $blocks = @(Get-OfficePdfText -Path $path -AsTextBlock)
        $valueBlock = $blocks | Where-Object { $_.Text -match 'ThisIdentifierShouldShrinkToFit' } | Select-Object -First 1
        $valueBlock | Should -Not -BeNullOrEmpty
        $valueBlock.FontSize | Should -BeLessThan 18
        $valueBlock.FontSize | Should -BeGreaterOrEqual 7
    }

    It 'adds piped PDF table rows to a supplied document' {
        $path = Join-Path $TestDrive 'piped-rows-supplied-document.pdf'
        $doc = New-OfficePdf {
            PdfHeading 'Pipeline rows'
        }
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $updated = $rows | PdfTable -Document $doc -PassThru
        $updated | Save-OfficePdf -Path $path | Out-Null

        $updated | Should -Be $doc
        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Alpha'
        $text | Should -Match 'Beta'
    }

    It 'adds explicit PDF table rows to each piped document' {
        $path1 = Join-Path $TestDrive 'piped-document-one.pdf'
        $path2 = Join-Path $TestDrive 'piped-document-two.pdf'
        $doc1 = New-OfficePdf {
            PdfHeading 'First document'
        }
        $doc2 = New-OfficePdf {
            PdfHeading 'Second document'
        }
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
        (Get-OfficePdfText -Path $path1) | Should -Match 'Alpha'
        (Get-OfficePdfText -Path $path2) | Should -Match 'Alpha'
    }

    It 'keeps the current page size when only margins are updated' {
        $path = Join-Path $TestDrive 'page-size-margin.pdf'

        New-OfficePdf -Path $path {
            PdfPageSetup -PageSize A4
            PdfPageSetup -Margin 36
            PdfHeading 'A4 report'
        } | Out-Null

        $page = (Get-OfficePdfInfo -Path $path).Pages[0]
        [math]::Round($page.Width) | Should -Be 595
        [math]::Round($page.Height) | Should -Be 842
    }

    It 'reports append-only mutation support for generated PDFs' {
        (Get-Command Get-OfficePdfAppendOnlyMutation).OutputType.Type.Name | Should -Contain 'PdfAppendOnlyMutationReport'

        $path = Join-Path $TestDrive 'append-plan.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'Append-only metadata readiness'
        } | Out-Null

        $plan = Get-OfficePdfAppendOnlyMutation -Path $path
        $plan.CanAppendMetadata | Should -BeTrue
        $plan.CanAppendFormFields | Should -BeTrue
        $plan.SupportedActions | Should -Contain 'Metadata'
        $plan.SupportedActions | Should -Contain 'FormFill'
        $plan.Summary | Should -Match 'Incremental updates are supported'
    }

    It 'supports expanded page sizes and PDF 2 file version' {
        $path = Join-Path $TestDrive 'b5-pdf20.pdf'

        New-OfficePdf -Path $path -FileVersion Pdf20 {
            PdfPageSetup -PageSize B5
            PdfHeading 'B5 PDF 2'
        } | Out-Null

        $info = Get-OfficePdfInfo -Path $path
        $info.HeaderVersion | Should -Be '2.0'
        $info.EffectiveVersion | Should -Be '2.0'
        $info.IsPdf20OrLater | Should -BeTrue
        [math]::Round($info.Pages[0].Width) | Should -Be 498
        [math]::Round($info.Pages[0].Height) | Should -Be 708
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
        Copy-OfficePdfPage -Path $path -PageRange '1' -OutputPath $copy -PassThru | Should -BeOfType System.IO.FileInfo
        Remove-OfficePdfPage -Path $path -PageRange '2' -OutputPath $removed -PassThru | Should -BeOfType System.IO.FileInfo

        (Get-OfficePdfInfo -Path $copy).PageCount | Should -Be 1
        (Get-OfficePdfInfo -Path $removed).PageCount | Should -Be 1
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

        $pages = @(Get-OfficePdfText -Path $path -ByPage)
        $pages.Count | Should -Be 3
        $pages[0].PageNumber | Should -Be 1
        $pages[0].Text | Should -Match 'Chapter One'
        $pages[2].Text | Should -Match 'Chapter Three'

        $groupDirectory = Join-Path $TestDrive 'groups'
        $groups = @(Split-OfficePdf -Path $path -OutputDirectory $groupDirectory -Prefix 'group' -PagesPerDocument 2)
        $groups.Count | Should -Be 2
        (Get-OfficePdfInfo -Path $groups[0].FullName).PageCount | Should -Be 2
        (Get-OfficePdfInfo -Path $groups[1].FullName).PageCount | Should -Be 1

        $rangeDirectory = Join-Path $TestDrive 'ranges'
        $ranges = @(Split-OfficePdf -Path $path -OutputDirectory $rangeDirectory -Prefix 'range' -PageRange '1-2', '3')
        $ranges.Count | Should -Be 2
        Get-OfficePdfText -Path $ranges[0].FullName | Should -Match 'Chapter Two'
        Get-OfficePdfText -Path $ranges[1].FullName | Should -Match 'Chapter Three'

        $bookmarkDirectory = Join-Path $TestDrive 'bookmarks'
        $bookmarks = @(Split-OfficePdf -Path $path -OutputDirectory $bookmarkDirectory -Prefix 'bookmark' -BookmarkName 'Chapter Two')
        $bookmarks.Count | Should -Be 1
        Get-OfficePdfText -Path $bookmarks[0].FullName | Should -Match 'Chapter Two'
        Get-OfficePdfText -Path $bookmarks[0].FullName | Should -Not -Match 'Chapter Three'
    }

    It 'splits encrypted PDFs and writes padded page names' {
        foreach ($name in 'Split-OfficePdf') {
            (Get-Command $name).Parameters.Keys | Should -Contain 'Password'
            (Get-Command $name).Parameters.Keys | Should -Contain 'PadIndex'
            (Get-Command $name).Parameters.Keys | Should -Contain 'IndexWidth'
        }

        $path = Join-Path $TestDrive 'split-encrypted-source.pdf'
        New-OfficePdf -Path $path -Password 'open' {
            PdfParagraph 'Encrypted page one'
            PdfPageBreak
            PdfParagraph 'Encrypted page two'
            PdfPageBreak
            PdfParagraph 'Encrypted page three'
        } | Out-Null

        $outputDirectory = Join-Path $TestDrive 'encrypted-split'
        $outputs = @(Split-OfficePdf -Path $path -Password 'open' -OutputDirectory $outputDirectory -Prefix 'page' -IndexWidth 3)

        $outputs.Count | Should -Be 3
        $outputs[0].Name | Should -Be 'page-001.pdf'
        $outputs[1].Name | Should -Be 'page-002.pdf'
        $outputs[2].Name | Should -Be 'page-003.pdf'
        foreach ($output in $outputs) {
            { Get-OfficePdfText -Path $output.FullName } | Should -Throw '*requires a password*'
            (Get-OfficePdfInfo -Path $output.FullName -Password 'open').Security.HasEncryption | Should -BeTrue
        }
        Get-OfficePdfText -Path $outputs[1].FullName -Password 'open' | Should -Match 'Encrypted page two'
    }

    It 'does not replace an existing split destination' {
        $path = Join-Path $TestDrive 'split-collision-source.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'Collision-safe split output'
        } | Out-Null

        $outputDirectory = Join-Path $TestDrive 'split-collision-output'
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
        $existingPath = Join-Path $outputDirectory 'page-1.pdf'
        [IO.File]::WriteAllText($existingPath, 'keep this file')

        $written = @(Split-OfficePdf -Path $path -OutputDirectory $outputDirectory)

        $written.Count | Should -Be 1
        $written[0].Name | Should -Be 'page-1-2.pdf'
        [IO.File]::ReadAllText($existingPath) | Should -Be 'keep this file'
        Get-OfficePdfText -Path $written[0].FullName | Should -Match 'Collision-safe split output'
    }

    It 'does not create split output directories when WhatIf skips writes' {
        $path = Join-Path $TestDrive 'split-whatif-source.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'Page one'
            PdfPageBreak
            PdfParagraph 'Page two'
        } | Out-Null

        $outputDirectory = Join-Path $TestDrive 'split-whatif-output'
        Split-OfficePdf -Path $path -OutputDirectory $outputDirectory -WhatIf | Out-Null

        Test-Path -LiteralPath $outputDirectory | Should -BeFalse
    }

    It 'merges and resizes PDFs to fixed paper sizes' {
        foreach ($parameter in 'FlattenVisualAnnotations', 'PageSize', 'Width', 'Height', 'Landscape', 'ResizeMode', 'ResizeMargin') {
            (Get-Command Join-OfficePdf).Parameters.Keys | Should -Contain $parameter
        }

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

        Join-OfficePdf -Path $letter, $a5 -OutputPath $merged -PageSize A4 -ResizeMargin 12 -ResizeMode Fit -PassThru |
            Should -BeOfType System.IO.FileInfo

        $info = Get-OfficePdfInfo -Path $merged
        $info.PageCount | Should -Be 2
        foreach ($page in $info.Pages) {
            [math]::Round($page.Width) | Should -Be 595
            [math]::Round($page.Height) | Should -Be 842
        }

        Get-OfficePdfText -Path $merged | Should -Match 'Letter source text'
        Get-OfficePdfText -Path $merged | Should -Match 'A5 source text'
    }

    It 'resizes selected PDF pages' {
        foreach ($parameter in 'PageSize', 'Width', 'Height', 'Landscape', 'ResizeMode', 'ResizeMargin') {
            (Get-Command Set-OfficePdfPage).Parameters.Keys | Should -Contain $parameter
        }

        $path = Join-Path $TestDrive 'resize-page-source.pdf'
        $outputPath = Join-Path $TestDrive 'resize-page-output.pdf'
        New-OfficePdf -Path $path {
            PdfPageSetup -PageSize A4
            PdfParagraph 'Resize first page'
            PdfPageBreak
            PdfParagraph 'Leave second page'
        } | Out-Null

        Set-OfficePdfPage -Path $path -OutputPath $outputPath -PageRange '1' -PageSize Letter -ResizeMargin 18 -PassThru |
            Should -BeOfType System.IO.FileInfo

        $info = Get-OfficePdfInfo -Path $outputPath
        [math]::Round($info.Pages[0].Width) | Should -Be 612
        [math]::Round($info.Pages[0].Height) | Should -Be 792
        [math]::Round($info.Pages[1].Width) | Should -Be 595
        [math]::Round($info.Pages[1].Height) | Should -Be 842
        Get-OfficePdfText -Path $outputPath | Should -Match 'Resize first page'
        Get-OfficePdfText -Path $outputPath | Should -Match 'Leave second page'
    }

    It 'exposes diagnostics, optimization hints, structured text, catalog settings, and redaction planning' {
        $path = Join-Path $TestDrive 'diagnostic-source.pdf'
        New-OfficePdf -Path $path -CreateOutlineFromHeadings -IncludePageLabels -PageLabelPrefix 'P-' -DisplayDocTitle -FitWindow -OpenActionPage 1 -OpenActionMode Fit -PageMode UseOutlines {
            PdfMetadata -Title 'Diagnostic Report'
            PdfHeading 'Diagnostic Report'
            PdfParagraph 'This paragraph should appear in structured text and redaction planning.'
        } | Out-Null

        $info = Get-OfficePdfInfo -Path $path
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

        $optimization = Get-OfficePdfOptimization -Path $path
        $optimization.StreamCount | Should -Be $diagnostic.StreamCount
        $optimization.LargestStreams.Count | Should -BeGreaterThan 0
        $optimization.TotalStreamBytes | Should -BeGreaterThan 0
        $optimization.LargestStreamBytes | Should -BeGreaterThan 0
        $optimization.DuplicateStreamGroupCount | Should -BeGreaterOrEqual 0

        $blocks = @(Get-OfficePdfText -Path $path -AsTextBlock)
        $blocks.Count | Should -BeGreaterThan 0
        $blocks.Text -join "`n" | Should -Match 'Diagnostic Report'

        $plan = Get-OfficePdfRedactionPlan -Path $path -PageNumber 1 -X 0 -Y 0 -Width 1000 -Height 1000
        $plan.HasMatches | Should -BeTrue
        $plan.Matches.Text -join "`n" | Should -Match 'structured text'

        $flatPath = Join-Path $TestDrive 'diagnostic-source-flat-annotations.pdf'
        ConvertTo-OfficePdfFlatAnnotation -Path $path -OutputPath $flatPath | Should -BeOfType System.IO.FileInfo
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
        $info = Get-OfficePdfInfo -Path $incrementalPath
        $info.Metadata.Title | Should -Be 'Updated title'
        $info.Metadata.Author | Should -Be 'Original author'
        $info.Security.HasIncrementalUpdates | Should -BeTrue
        $info.Security.RevisionCount | Should -BeGreaterThan 1
        Get-OfficePdfText -Path $incrementalPath | Should -Match 'Incremental body text'
    }

    It 'updates PDF form fields as an incremental revision' {
        (Get-Command Set-OfficePdfForm).Parameters.Keys | Should -Contain 'Incremental'
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

        Set-OfficePdfForm -Path $path -OutputPath $outputPath -Field @{ Name = 'Grace' } -Incremental -PassThru |
            Should -BeOfType System.IO.FileInfo

        (Get-Item $outputPath).Length | Should -BeGreaterThan (Get-Item $path).Length
        $field = Get-OfficePdfFormField -Path $outputPath -Name Name
        $field.Value | Should -Be 'Grace'
        $info = Get-OfficePdfInfo -Path $outputPath
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

        Set-OfficePdfForm -Path $path -OutputPath $outputPath -Field @{ Name = 'Grace' } -Incremental -KeepNeedAppearances -PassThru |
            Should -BeOfType System.IO.FileInfo

        $field = Get-OfficePdfFormField -Path $outputPath -Name Name
        $field.Value | Should -Be 'Grace'
        $info = Get-OfficePdfInfo -Path $outputPath
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

        Set-OfficePdfForm -Path $path -OutputPath $outputPath -Field @{ Name = 'Grace' } -Incremental -PassThru |
            Should -BeOfType System.IO.FileInfo

        $field = Get-OfficePdfFormField -Path $outputPath -Name Name
        $field.Value | Should -Be 'Grace'
        $info = Get-OfficePdfInfo -Path $outputPath
        $info.Security.HasIncrementalUpdates | Should -BeTrue
        $info.Security.HasSignatures | Should -BeTrue
        $info.Security.HasDocMDPPermissions | Should -BeTrue
        $info.AcroFormNeedAppearances | Should -BeFalse
    }

    It 'sets page production boundary boxes' {
        (Get-Command Set-OfficePdfPage).Parameters.Keys | Should -Contain 'BoxName'
        $path = Join-Path $TestDrive 'pagebox-source.pdf'
        $outputPath = Join-Path $TestDrive 'pagebox-output.pdf'
        New-OfficePdf -Path $path {
            PdfParagraph 'Page box source'
        } | Out-Null

        Set-OfficePdfPage -Path $path -OutputPath $outputPath -BoxName TrimBox -Left 12 -Bottom 14 -Right 222 -Top 244 -PassThru |
            Should -BeOfType System.IO.FileInfo

        $geometry = (Get-OfficePdfInfo -Path $outputPath).Pages[0].Geometry
        $geometry.TrimBox.Left | Should -Be 12
        $geometry.TrimBox.Bottom | Should -Be 14
        $geometry.TrimBox.Right | Should -Be 222
        $geometry.TrimBox.Top | Should -Be 244
    }

    It 'optimizes PDFs with lossless stream compression' {
        (Get-Command ConvertTo-OfficePdfOptimized).Parameters.Keys | Should -Contain 'PassThruReport'
        (Get-Command ConvertTo-OfficePdfOptimized).Parameters.Keys | Should -Contain 'MinimumStreamCompressionBytes'
        (Get-Command ConvertTo-OfficePdfOptimized).Parameters.Keys | Should -Contain 'KeepUnreferencedObjects'
        (Get-Command ConvertTo-OfficePdfOptimized).Parameters.Keys | Should -Contain 'KeepDuplicateStreams'

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

        $report = ConvertTo-OfficePdfOptimized -Path $path -OutputPath $optimizedPath -PassThruReport

        $report.Applied | Should -BeTrue
        $report.ActionCount | Should -Be 1
        $report.ReportAfter | Should -Not -BeNullOrEmpty
        $report.CandidateSavedBytes | Should -BeGreaterThan 0
        $report.SkippedActionCount | Should -BeGreaterThan -1
        $report.Actions[0].Kind | Should -Be 'CompressStream'
        Test-Path $optimizedPath | Should -BeTrue
        (Get-Item $optimizedPath).Length | Should -BeLessThan (Get-Item $path).Length
        Get-OfficePdfText -Path $optimizedPath | Should -Match ('A' * 64)
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
        $report.Findings.Code | Should -Contain 'CryptographicTrustNotVerified'
        $report.Findings.Code | Should -Contain 'DocMDPDetected'
        $report.Findings.Code | Should -Contain 'LongTermValidationEvidenceDetected'
        $report.Findings.Code | Should -Contain 'SignatureDetachedCmsSubFilter'
    }

    It 'prepares and injects external PDF signature bytes' {
        $path = Join-Path $TestDrive 'signature-source.pdf'
        $preparedPath = Join-Path $TestDrive 'signature-prepared.pdf'
        $signedPath = Join-Path $TestDrive 'signature-applied.pdf'
        $signaturePath = Join-Path $TestDrive 'signature.der'

        New-OfficePdf -Path $path {
            PdfParagraph 'External signing workflow'
        } | Out-Null

        $plan = New-OfficePdfSignature -Path $path -OutputPath $preparedPath -FieldName Approval -Name Alice -Reason Approval -ReservedBytes 512 -PassThruReport
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
        $signedReport = Set-OfficePdfSignature -Path $preparedPath -SignaturePath $signaturePath -OutputPath $signedPath -PassThruReport -PassThru
        $signedReport.HasSignatures | Should -BeTrue
        $signedReport.Signatures[0].Signature.ByteRangeValues -join ',' | Should -Be ($plan.ByteRangeValues -join ',')
        $signedReport.Findings.Code | Should -Contain 'SignatureDetachedCmsSubFilter'
        [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($signedPath)) | Should -Match '3082010AAA55'
    }

    It 'reports complex PDF text layout diagnostics' {
        $arabic = -join ([char[]](0x0645, 0x0631, 0x062D, 0x0628, 0x0627))
        $diagnostics = @(Get-OfficePdfTextDiagnostic -Text $arabic -AdvancedLayout)

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

        $block = Get-OfficePdfText -Path $path -AsTextBlock |
            Where-Object { $_.Text -match 'Secret account' } |
            Select-Object -First 1
        $block | Should -Not -BeNullOrEmpty

        $x = [math]::Min($block.XStart, $block.XEnd) - 2
        $width = [math]::Abs($block.XEnd - $block.XStart) + 4
        $y = $block.BaselineY - 14

        ConvertTo-OfficePdfRedacted -Path $path -OutputPath $redactedPath -PageNumber $block.PageNumber -X $x -Y $y -Width $width -Height 20 -FillColor '#111111' |
            Should -BeOfType System.IO.FileInfo

        $text = Get-OfficePdfText -Path $redactedPath
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

        ConvertTo-OfficePdfRedacted -Path $path -OutputPath $redactedPath -Text 'Secret account' |
            Should -BeOfType System.IO.FileInfo

        $text = Get-OfficePdfText -Path $redactedPath
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
            ConvertTo-OfficePdfRedacted -Path $path -OutputPath $redactedPath `
                -Text 'Absent secret' -ErrorAction Stop
        } | Should -Throw '*No PDF text blocks matched*'

        Test-Path -LiteralPath $redactedPath | Should -BeFalse
    }

    It 'creates PDFs with document-level font options' {
        $path = Join-Path $TestDrive 'font-options.pdf'
        New-OfficePdf -Path $path -DefaultFont Courier -DefaultFontSize 13 {
            PdfHeading 'Font Options'
            PdfParagraph 'Generated with document-level font settings.'
        } | Out-Null

        (Get-OfficePdfPreflight -Path $path).CanRead | Should -BeTrue
        Get-OfficePdfText -Path $path | Should -Match 'Font Options'
    }

    It 'builds visually polished PDF documents' {
        $path = Join-Path $TestDrive 'visual-polish.pdf'
        New-OfficePdf -Path $path {
            PdfBackground -Color '#F8FAFC'
            PdfPageBorder -Color '#0F766E' -Width 1.5 -Inset 24 -Opacity 0.85
            PdfBookmark 'summary'
            PdfHeading 'Visual Summary'
            PdfHr -Color '#0F766E' -Thickness 2 -SpacingBefore 8 -SpacingAfter 10
            PdfPanel 'A polished report surface should support visual structure.'
        } | Out-Null

        $info = Get-OfficePdfInfo -Path $path
        $info.PageCount | Should -Be 1
        $info.HasNamedDestinations | Should -BeTrue
        $info.NamedDestinationNames | Should -Contain 'summary'

        (Get-OfficePdfPreflight -Path $path).CanRead | Should -BeTrue
        Get-OfficePdfText -Path $path | Should -Match 'Visual Summary'
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
            PdfRow -Gap 16 -SpacingBefore 4 -SpacingAfter 8 -ColumnSeparatorColor '#CBD5E1' -Column @(
                @{
                    Width = 38
                    Content = @(
                        @{ Type = 'Heading'; Level = 2; Text = 'Highlights'; HeadingColor = '#0F766E' }
                        @{ Type = 'Paragraph'; Text = 'Left column summary.' }
                        @{ Type = 'List'; Items = @('Designed surface', 'Readable rhythm'); Numbered = $true }
                    )
                }
                @{
                    Width = 62
                    Content = @(
                        @{ Type = 'Bookmark'; Name = 'layout-row' }
                        @{ Type = 'Panel'; Text = 'Right column panel.' }
                        @{ Type = 'Table'; InputObject = $rows; TableStyle = 'Compact'; Caption = 'Layout row table'; NoBorder = $true }
                        @{ Type = 'Rule'; Color = '#0F766E'; Thickness = 1.2 }
                    )
                }
            )
            PdfParagraph 'After layout row.'
        } | Out-Null

        $info = Get-OfficePdfInfo -Path $path
        $info.PageCount | Should -Be 1
        $info.NamedDestinationNames | Should -Contain 'layout-row'

        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Highlights'
        $text | Should -Match 'Right column panel'
        $text | Should -Match 'Layout row table'
        $text | Should -Match 'After layout row'
    }

    It 'renders typed list content directly inside PDF row columns' {
        $path = Join-Path $TestDrive 'layout-row-list.pdf'
        New-OfficePdf -Path $path {
            PdfRow -Column @(
                @{
                    Width = 100
                    Type = 'List'
                    Items = @('Alpha', 'Beta')
                }
            )
        } | Out-Null

        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Alpha'
        $text | Should -Match 'Beta'
    }

    It 'renders a single hashtable content block inside PDF row columns' {
        $path = Join-Path $TestDrive 'layout-row-single-content.pdf'
        New-OfficePdf -Path $path {
            PdfRow -Column @(
                @{
                    Width = 100
                    Content = @{
                        Type = 'Paragraph'
                        Text = 'Single content block'
                    }
                }
            )
        } | Out-Null

        Get-OfficePdfText -Path $path | Should -Match 'Single content block'
    }

    It 'renders a single table object inside PDF row columns' {
        $path = Join-Path $TestDrive 'layout-row-single-table-object.pdf'
        New-OfficePdf -Path $path {
            PdfRow -Column @(
                @{
                    Width = 100
                    Content = @(
                        @{
                            Type = 'Table'
                            InputObject = @{ Name = 'Alpha'; Value = 1 }
                            Header = @('Name', 'Value')
                        }
                    )
                }
            )
        } | Out-Null

        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Name'
        $text | Should -Match 'Value'
        $text | Should -Match 'Alpha'
        $text | Should -Match '1'
    }

    It 'renders a single hashtable rich-text run inside PDF row columns' {
        $path = Join-Path $TestDrive 'layout-row-rich-run.pdf'
        New-OfficePdf -Path $path {
            PdfRow -Column @(
                @{
                    Width = 100
                    Content = @(
                        @{
                            Type = 'Paragraph'
                            Run = @{
                                Text = 'Approved'
                                Bold = $true
                            }
                        }
                    )
                }
            )
        } | Out-Null

        Get-OfficePdfText -Path $path | Should -Match 'Approved'
    }

    It 'builds rich PDF text with emphasis and links' {
        $path = Join-Path $TestDrive 'rich-text.pdf'
        New-OfficePdf -Path $path {
            PdfBookmark 'details'
            PdfText -Run @(
                @{ Text = 'This paragraph contains ' }
                @{ Text = 'bold'; Bold = $true; Color = '#0F766E' }
                @{ Text = ', ' }
                @{ Text = 'italic'; Italic = $true }
                @{ Text = ', highlighted '; BackgroundColor = '#FEF3C7' }
                @{ Text = 'external link'; LinkUri = 'https://evotec.xyz'; Color = '#2563EB' }
                @{ Text = ', and ' }
                @{ Text = 'bookmark link'; LinkDestinationName = 'details'; Color = '#7C3AED' }
                @{ Text = '.' }
            )
            PdfRow -Column @(
                @{
                    Width = 100
                    Content = @(
                        @{
                            Type = 'Paragraph'
                            Run = @(
                                @{ Text = 'Row layout can also contain ' }
                                @{ Text = 'rich inline text'; Bold = $true; LinkDestinationName = 'details'; Color = '#0F766E' }
                                @{ Text = '.' }
                            )
                        }
                    )
                }
            )
        } | Out-Null

        $info = Get-OfficePdfInfo -Path $path
        $info.HasLinkAnnotations | Should -BeTrue
        $info.LinkUris | Should -Contain 'https://evotec.xyz'
        $info.LinkDestinationNames | Should -Contain 'details'

        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'bold'
        $text | Should -Match 'bookmark link'
        $text | Should -Match 'rich inline text'
    }

    It 'renders every rich-text value offered by tab completion' {
        $path = Join-Path $TestDrive 'rich-text-completion-values.pdf'
        New-OfficePdf -Path $path {
            PdfText -Run @(
                PdfTextRun 'Total' -Baseline Normal
                PdfTextRun -Kind Tab -TabLeader Hyphens -TabAlignment DecimalSeparator
                PdfTextRun '12.34'
                PdfTextRun -Kind Tab -TabLeader Underscores -TabAlignment Right
                PdfTextRun 'Done'
            )
        } | Out-Null

        Get-OfficePdfText -Path $path | Should -Match 'Total'
        Get-OfficePdfText -Path $path | Should -Match '12.34'
        Get-OfficePdfText -Path $path | Should -Match 'Done'
    }

    It 'normalizes typed text run kinds before rendering adapters consume them' {
        $assembly = [PSWriteOffice.Cmdlets.Text.NewOfficeTextRunCommand].Assembly
        $parserType = $assembly.GetType('PSWriteOffice.Services.Text.OfficeTextRunParser', $true)
        $specType = $assembly.GetType('PSWriteOffice.Services.Text.OfficeTextRunSpec', $true)
        $parseMethod = $parserType.GetMethod('Parse', [Reflection.BindingFlags]'Static,NonPublic')
        $spec = [Activator]::CreateInstance($specType)
        $spec.Text = '2'
        $spec.Kind = 'Superscript'

        $parsed = $parseMethod.Invoke($null, @($spec))

        $parsed.Baseline | Should -Be 'Superscript'
    }

    It 'applies PDF themes and decorative backgrounds' {
        $path = Join-Path $TestDrive 'styled-backgrounds.pdf'
        $imagePath = Join-Path (Join-Path $PSScriptRoot 'Assets') 'CellImage.png'

        New-OfficePdf -Path $path {
            PdfTheme Report
            PdfBackground -Color '#FFFFFF'
            PdfBackgroundImage -Path $imagePath -Fit Cover -Opacity 0.03
            PdfBackgroundShape -Shape TopBand -Height 86 -FillColor '#DBEAFE' -FillOpacity 0.75
            PdfBackgroundShape -Shape Ellipse -X 420 -Y 650 -Width 96 -Height 72 -FillColor '#99F6E4' -FillOpacity 0.35
            PdfHeading 'Styled Report'
            PdfParagraph 'Theme and decorative backgrounds are OfficeIMO-owned PDF layout features.'
            PdfPanel 'A reusable theme keeps report rhythm consistent.'
        } | Out-Null

        (Get-OfficePdfPreflight -Path $path).CanRead | Should -BeTrue
        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Styled Report'
        $text | Should -Match 'reusable theme'
    }
}
