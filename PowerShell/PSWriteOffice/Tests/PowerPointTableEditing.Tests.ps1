BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $manifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'PowerPointTableEditing workflows' {
    BeforeEach {
        $presentation = $null
        $reloaded = $null
        $afterRemoval = $null
        $source = $null
        $target = $null
    }
    AfterEach {
        foreach ($document in @($presentation, $reloaded, $afterRemoval, $source, $target)) {
            if ($null -ne $document) {
                $document.Dispose()
            }
        }
    }

    It 'supports transposed PowerPoint tables' {
        $path = Join-Path $TestDrive 'PowerPointTransposedTable.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        try {
            $slide = $presentation.AddSlide(0, 1)
            $rows = @(
                [PSCustomObject]@{ Region = 'Europe'; Revenue = 21704714 }
                [PSCustomObject]@{ Region = 'Asia'; Revenue = 8774099 }
            )

            Add-OfficePowerPointTable -Slide $slide -InputObject $rows -View Transpose -X 40 -Y 120 -Width 420 -Height 160 | Out-Null

            $shapeInfo = @(Get-OfficePowerPointShape -Slide $slide | Where-Object Kind -eq 'Table')[0]
            $shapeInfo.RowCount | Should -Be 3
            $shapeInfo.ColumnCount | Should -Be 3
        } finally {
            Close-OfficePowerPoint -Document $presentation
        }
    }

    It 'preserves selected PowerPoint table properties and their order' {
        $path = Join-Path $TestDrive 'PowerPointSelectedProperties.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        try {
            $slide = $presentation.AddSlide(0, 1)
            $rows = @(
                [PSCustomObject]@{ Item = 'Alpha'; Qty = 10 }
                [PSCustomObject]@{ Item = 'Beta'; Qty = 20 }
            )

            $table = Add-OfficePowerPointTable -Slide $slide -InputObject $rows -Property Qty, Item -X 40 -Y 120 -Width 420 -Height 160 -PassThru

            $shapeInfo = @(Get-OfficePowerPointShape -Slide $slide | Where-Object Kind -eq 'Table')[0]
            $shapeInfo.RowCount | Should -Be 3
            $shapeInfo.ColumnCount | Should -Be 2
            $table.GetCell(0, 0).Text | Should -Be 'Qty'
            $table.GetCell(0, 1).Text | Should -Be 'Item'
            $table.GetCell(1, 0).Text | Should -Be '10'
            $table.GetCell(1, 1).Text | Should -Be 'Alpha'
        } finally {
            Close-OfficePowerPoint -Document $presentation
        }
    }

    It 'finds and modifies existing PowerPoint text boxes and tables' {
        $path = Join-Path $TestDrive 'PowerPointExistingShapeModify.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        try {
            $slide = $presentation.AddSlide(0, 1)
            Add-OfficePowerPointTextBox -Slide $slide -Text 'Draft status' -X 80 -Y 80 -Width 300 -Height 50 | Out-Null
            $rows = @(
                [PSCustomObject]@{ Metric = 'Risk'; State = 'Open' }
                [PSCustomObject]@{ Metric = 'Quality'; State = 'Ready' }
            )
            Add-OfficePowerPointTable -Slide $slide -InputObject $rows -X 80 -Y 160 -Width 420 -Height 180 | Out-Null

            $textShape = Find-OfficePowerPointShape -Presentation $presentation -Text 'Draft status' -Kind TextBox |
                Set-OfficePowerPointShapeText -Text 'Ready status' -PassThru
            $textShape.Text | Should -Be 'Ready status'

            $tableShape = Find-OfficePowerPointShape -Presentation $presentation -Text 'Risk' -Kind Table | Select-Object -First 1
            $tableShape | Should -Not -BeNullOrEmpty
            $row = $tableShape | Add-OfficePowerPointTableRow -Value 'Latency', 'Investigating' -PassThru
            $row.Cells[0].Text | Should -Be 'Latency'
            $row.Cells[1].Text | Should -Be 'Investigating'

            $cell = $tableShape | Set-OfficePowerPointTableCell -Row 1 -Column 1 -Text 'Mitigating' -PassThru
            $cell.Text | Should -Be 'Mitigating'

            Save-OfficePowerPoint -Document $presentation
            $presentation.Dispose()
            $presentation = $null
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Document $presentation
            }
        }

        $reloaded = Open-OfficePowerPoint -Path $path
        try {
            $readyShape = Find-OfficePowerPointShape -Presentation $reloaded -Text 'Ready status' -Kind TextBox | Select-Object -First 1
            $readyShape | Should -Not -BeNullOrEmpty

            $updatedTable = Find-OfficePowerPointShape -Presentation $reloaded -Text 'Mitigating' -Kind Table | Select-Object -First 1
            $updatedTable | Should -Not -BeNullOrEmpty
            $updatedTable.RowCount | Should -Be 4
            $latencyTable = Find-OfficePowerPointShape -Presentation $reloaded -Text 'Latency' -Kind Table | Select-Object -First 1
            $latencyTable | Should -Not -BeNullOrEmpty
        } finally {
            Close-OfficePowerPoint -Document $reloaded
        }
    }

    It 'creates and updates PowerPoint text and table cells from rich text runs' {
        $path = Join-Path $TestDrive 'PowerPointRichTextRuns.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        try {
            $slide = $presentation.AddSlide(0, 1)
            $textBox = Add-OfficePowerPointTextBox -Slide $slide -Run @(
                New-OfficeTextRun 'Status: '
                New-OfficeTextRun 'Ready' -Color SeaGreen -Bold
            ) -X 80 -Y 80 -Width 300 -Height 50 -PassThru
            $textBox.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointTextBox'
            $textBox.Text | Should -Be 'Status: Ready'
            $textBox = $textBox | Set-OfficePowerPointShapeText -Run @(
                New-OfficeTextRun 'Linked' -Color Crimson -BackgroundColor Yellow -FontName 'Arial' -FontSize 18 -LinkUri 'https://example.org/ppt'
            ) -PassThru
            $textBox.Paragraphs[0].Runs[0].Color | Should -Be 'DC143C'
            $textBox.Paragraphs[0].Runs[0].HighlightColor | Should -Be 'FFFF00'
            $textBox.Paragraphs[0].Runs[0].Hyperlink.AbsoluteUri | Should -Be 'https://example.org/ppt'
            $textBox = $textBox | Set-OfficePowerPointShapeText -Run @(
                New-OfficeTextRun 'Plain' -BackgroundColor None
            ) -PassThru
            $textBox.Text | Should -Be 'Plain'
            $textBox.Paragraphs[0].Runs[0].Color | Should -BeNullOrEmpty
            $textBox.Paragraphs[0].Runs[0].HighlightColor | Should -BeNullOrEmpty
            $textBox.Paragraphs[0].Runs[0].FontName | Should -BeNullOrEmpty
            $textBox.Paragraphs[0].Runs[0].FontSize | Should -BeNullOrEmpty
            $textBox.Paragraphs[0].Runs[0].Hyperlink | Should -BeNullOrEmpty

            $table = Add-OfficePowerPointTable -Slide $slide -InputObject @(
                , @(
                    @{
                        Run = @(
                            New-OfficeTextRun 'Build '
                            New-OfficeTextRun 'Ready' -Color SeaGreen
                        )
                        ColumnSpan = 2
                        FillColor = 'AliceBlue'
                        Bold = $true
                        Color = 'Red'
                        FontSize = 18
                    }
                )
                , @('Owner', 'Platform')
                , @(@{ Text = 'Filled'; FillColor = 'Yellow'; Bold = $true }, 'Styled')
                , @(@{ Run = 'Queued' }, 'String run')
            ) -X 80 -Y 150 -Width 420 -Height 140 -PassThru
            $table.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointTable'
            $table.GetCell(0, 0).GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointTableCell'
            $table.GetCell(0, 0).Text | Should -Be 'Build Ready'
            $table.GetCell(0, 0).Runs[0].Bold | Should -BeTrue
            $table.GetCell(0, 0).Runs[0].Color | Should -Be 'FF0000'
            $table.GetCell(0, 0).Runs[0].FontSize.ToPoints() | Should -Be 18
            $table.GetCell(0, 0).Runs[1].Bold | Should -BeTrue
            $table.GetCell(0, 0).Runs[1].Color | Should -Be '2E8B57'
            $table.GetCell(1, 0).Text | Should -Be 'Owner'
            $table.GetCell(2, 0).Text | Should -Be 'Filled'
            $table.GetCell(2, 0).Runs[0].Bold | Should -BeTrue
            $table.GetCell(2, 0).Runs[0].HighlightColor | Should -BeNullOrEmpty
            $table.GetCell(2, 0).FillColor | Should -Be 'FFFF00'
            $table.GetCell(3, 0).Text | Should -Be 'Queued'

            $row = $table | Add-OfficePowerPointTableRow -Value @(
                @{ Run = @(
                    New-OfficeTextRun 'Latency '
                    New-OfficeTextRun 'Ready' -Color SeaGreen -Bold
                ) },
                'SRE'
            ) -PassThru
            $row.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointTableRow'
            $row.GetCell(0).Text | Should -Be 'Latency Ready'

            $spannedRow = $table | Add-OfficePowerPointTableRow -Value @(
                @{ Text = 'Total'; ColumnSpan = 2; FillColor = 'AliceBlue' }
            ) -PassThru
            $spannedRow.GetCell(0).Text | Should -Be 'Total'
            $spannedRow.GetCell(0).Merge.Item1 | Should -Be 1
            $spannedRow.GetCell(0).Merge.Item2 | Should -Be 2

            $threeColumnTable = Add-OfficePowerPointTable -Slide $slide -InputObject @(
                , @('Metric', 'Scope', 'Value')
            ) -X 80 -Y 320 -Width 420 -Height 120 -PassThru
            $valueAfterSpan = $threeColumnTable | Add-OfficePowerPointTableRow -Value @(
                @{ Text = 'Total'; ColumnSpan = 2; FillColor = 'AliceBlue' },
                '42'
            ) -PassThru
            $valueAfterSpan.GetCell(0).Merge.Item2 | Should -Be 2
            $valueAfterSpan.GetCell(2).Text | Should -Be '42'

            {
                $threeColumnTable | Add-OfficePowerPointTableRow -Value @(
                    @{ Text = 'Too wide'; ColumnSpan = 4 }
                )
            } | Should -Throw '*ColumnSpan 4*'
            $threeColumnTable.Rows | Should -Be 2

            {
                $threeColumnTable | Add-OfficePowerPointTableRow -Value @(
                    @{ Text = 'Too tall'; RowSpan = 2 }
                )
            } | Should -Throw '*RowSpan*'
            $threeColumnTable.Rows | Should -Be 2

            {
                $threeColumnTable | Add-OfficePowerPointTableRow -Value @(
                    @{ Run = @(New-OfficeTextRun 'site' -LinkUri 'https://example.org/table-cell') }
                )
            } | Should -Throw '*do not support hyperlinks yet*'
            $threeColumnTable.Rows | Should -Be 2

            $cell = $table | Set-OfficePowerPointTableCell -Row 1 -Column 1 -Run @(
                New-OfficeTextRun 'Owner '
                New-OfficeTextRun 'Ready' -Color Navy -Bold
            ) -PassThru
            $cell.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointTableCell'
            $cell.Text | Should -Be 'Owner Ready'
            $cell = $table | Set-OfficePowerPointTableCell -Row 1 -Column 1 -Run @(
                New-OfficeTextRun 'Plain'
            ) -PassThru
            $cell.Text | Should -Be 'Plain'
            $cell.Runs[0].Color | Should -BeNullOrEmpty

            { $table | Set-OfficePowerPointTableCell -Row 1 -Column 1 -Run @(
                New-OfficeTextRun 'site' -LinkUri 'https://example.org/table-cell'
            ) } | Should -Throw
            $table.GetCell(1, 1).Text | Should -Be 'Plain'

            $baselineTextBox = Add-OfficePowerPointTextBox -Slide $slide -Run @(
                New-OfficeTextRun 'x'
                New-OfficeTextRun '2' -Kind Superscript
            ) -X 360 -Y 80 -Width 120 -Height 50 -PassThru
            $baselineTextBox.Text | Should -Be 'x2'

            $baselineTable = Add-OfficePowerPointTable -Slide $slide -InputObject @(
                , @(
                    @{
                        Run = @(
                            New-OfficeTextRun 'H'
                            New-OfficeTextRun '2' -Baseline Subscript
                            New-OfficeTextRun 'O'
                        )
                    }
                )
            ) -X 80 -Y 450 -Width 240 -Height 80 -PassThru
            $baselineTable.GetCell(0, 0).Text | Should -Be 'H2O'

            Save-OfficePowerPoint -Document $presentation
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Document $presentation
            }
        }

        $slideXml = Get-ZipXmlDocumentLocal -Path $path -Entry (Get-TestPowerPointSlideEntry -Path $path -Index 0)
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($slideXml.NameTable)
        $namespaceManager.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        $slideXml.SelectSingleNode('//a:r[a:t="2"]/a:rPr[@baseline="30000"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $slideXml.SelectSingleNode('//a:r[a:t="2"]/a:rPr[@baseline="-25000"]', $namespaceManager) | Should -Not -BeNullOrEmpty
    }

    It 'keeps ordinary style-named columns in PowerPoint object tables' {
        $path = Join-Path $TestDrive 'PowerPointOrdinaryStyleColumns.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        try {
            $slide = $presentation.AddSlide(0, 1)
            $table = Add-OfficePowerPointTable -Slide $slide -InputObject @(
                [pscustomobject]@{
                    Text     = 'Apple'
                    Color    = 'Red'
                    FontSize = 'Large'
                }
            ) -X 80 -Y 150 -Width 300 -Height 100 -PassThru

            $table.GetCell(0, 0).Text | Should -Be 'Text'
            $table.GetCell(0, 1).Text | Should -Be 'Color'
            $table.GetCell(0, 2).Text | Should -Be 'FontSize'
            $table.GetCell(1, 0).Text | Should -Be 'Apple'
            $table.GetCell(1, 1).Text | Should -Be 'Red'
            $table.GetCell(1, 2).Text | Should -Be 'Large'
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Document $presentation
            }
        }
    }

    It 'validates structured PowerPoint table runs before creating the table' {
        $path = Join-Path $TestDrive 'PowerPointStructuredTableRunValidation.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        try {
            $slide = $presentation.AddSlide(0, 1)
            @($slide.Tables).Count | Should -Be 0

            {
                Add-OfficePowerPointTable -Slide $slide -InputObject @(
                    , @(
                        @{ Run = @(New-OfficeTextRun 'site' -LinkUri 'https://example.org/table-cell') }
                    )
                ) -X 80 -Y 150 -Width 300 -Height 100 -ErrorAction Stop
            } | Should -Throw '*do not support hyperlinks yet*'

            @($slide.Tables).Count | Should -Be 0
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Document $presentation
            }
        }
    }

    It 'validates PowerPoint text box runs before creating the shape' {
        $path = Join-Path $TestDrive 'PowerPointTextBoxRunValidation.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        try {
            $slide = $presentation.AddSlide(0, 1)
            @($slide.TextBoxes).Count | Should -Be 0

            {
                Add-OfficePowerPointTextBox -Slide $slide -Run @(
                    New-OfficeTextRun 'jump' -LinkDestinationName Summary
                ) -ErrorAction Stop
            } | Should -Throw '*named PDF/Word destinations*'

            @($slide.TextBoxes).Count | Should -Be 0
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Document $presentation
            }
        }
    }

    It 'preserves explicit headers on structured PowerPoint tables' {
        $path = Join-Path $TestDrive 'PowerPointStructuredHeaders.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        try {
            $slide = $presentation.AddSlide(0, 1)
            $table = Add-OfficePowerPointTable -Slide $slide -Property Qty, Item -InputObject @(
                [pscustomobject]@{
                    Item = 'Alpha'
                    Qty  = 10
                }
                , @(
                    @{ Run = @(New-OfficeTextRun 'Total' -Bold) },
                    'Two'
                )
            ) -X 80 -Y 150 -Width 420 -Height 140 -PassThru

            $table.GetCell(0, 0).Text | Should -Be 'Qty'
            $table.GetCell(0, 1).Text | Should -Be 'Item'
            $table.GetCell(1, 0).Text | Should -Be '10'
            $table.GetCell(1, 1).Text | Should -Be 'Alpha'
            $table.GetCell(2, 0).Runs[0].Text | Should -Be 'Total'
            $table.GetCell(2, 1).Text | Should -Be 'Two'
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Document $presentation
            }
        }
    }

    It 'keeps ordinary Run columns in PowerPoint object tables' {
        $path = Join-Path $TestDrive 'PowerPointOrdinaryRunColumn.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        try {
            $slide = $presentation.AddSlide(0, 1)
            $table = Add-OfficePowerPointTable -Slide $slide -InputObject @(
                [pscustomobject]@{
                    Run = @('Nightly', 'Daily')
                }
            ) -X 80 -Y 150 -Width 300 -Height 100 -PassThru

            $table.GetCell(0, 0).Text | Should -Be 'Run'
            $table.GetCell(1, 0).Text | Should -Match 'Nightly'
            $table.GetCell(1, 0).Text | Should -Match 'Daily'
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Document $presentation
            }
        }
    }
}
