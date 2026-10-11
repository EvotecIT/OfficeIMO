BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    Import-Module $manifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $examples = Join-Path $PSScriptRoot '../Examples'
    $artifacts = @{}
    foreach ($recipe in 'Excel/Recipe-Excel-BudgetDashboard', 'PowerPoint/Recipe-PowerPoint-QuarterlyReview', 'Workflows/Recipe-MultiFormat-StatusPack', 'Pdf/Recipe-Pdf-ServiceInvoice') {
        $folder = Join-Path $TestDrive ([IO.Path]::GetFileName($recipe))
        $null = New-Item -ItemType Directory -Path $folder
        Push-Location $folder
        try { & (Join-Path $examples ($recipe + '.ps1')) -ErrorAction Stop }
        finally { Pop-Location }
        foreach ($file in Get-ChildItem -LiteralPath $folder -File) { $artifacts[$file.Name] = $file.FullName }
    }

    function Get-RecipeWorksheetXml {
        param([string] $Path, [string] $Name)
        $book = Get-ZipXmlDocumentLocal -Path $Path -Entry 'xl/workbook.xml'
        $sheet = $book.SelectSingleNode("//*[local-name()='sheet' and @name='$Name']")
        $id = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $relationships = Get-ZipXmlDocumentLocal -Path $Path -Entry 'xl/_rels/workbook.xml.rels'
        $target = $relationships.SelectSingleNode("//*[local-name()='Relationship' and @Id='$id']").Target
        $entry = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { 'xl/' + $target }
        Get-ZipXmlDocumentLocal -Path $Path -Entry $entry
    }

    function Get-RecipeLuminance {
        param([string] $Hex)
        $linear = foreach ($offset in 0, 2, 4) {
            $channel = [Convert]::ToInt32($Hex.Substring($offset, 2), 16) / 255.0
            if ($channel -le 0.04045) { $channel / 12.92 } else { [math]::Pow(($channel + 0.055) / 1.055, 2.4) }
        }
        0.2126 * $linear[0] + 0.7152 * $linear[1] + 0.0722 * $linear[2]
    }
}

Describe 'Executable recipe artifacts' {
    It 'saves calculated budget values for readers that do not recalculate formulas' {
        $dashboard = Get-RecipeWorksheetXml -Path $artifacts['Budget-Dashboard.xlsx'] -Name Dashboard
        foreach ($expected in @(@{ Cell='B3'; Value=700000 }, @{ Cell='B4'; Value=651800 }, @{ Cell='B5'; Value=-2500 })) {
            $cell = $dashboard.SelectSingleNode("//*[local-name()='c' and @r='$($expected.Cell)']")
            $cell.SelectSingleNode("*[local-name()='f']") | Should -Not -BeNullOrEmpty
            $value = $cell.SelectSingleNode("*[local-name()='v']")
            $value | Should -Not -BeNullOrEmpty
            [decimal]::Parse($value.InnerText, [Globalization.CultureInfo]::InvariantCulture) | Should -Be $expected.Value
        }
        $detail = Get-RecipeWorksheetXml -Path $artifacts['Budget-Dashboard.xlsx'] -Name Detail
        foreach ($row in 2..5) {
            $detail.SelectSingleNode("//*[local-name()='c' and @r='E$row']/*[local-name()='f']").InnerText | Should -Be "D$row-B$row" -Because 'variance must update when a user edits the workbook forecast'
        }
    }

    It 'keeps the quarterly review cover text readable against its background' {
        $path = $artifacts['Quarterly-Business-Review.pptx']
        $slide = Get-ZipXmlDocumentLocal -Path $path -Entry (Get-TestPowerPointSlideEntry -Path $path -Index 0)
        $background = $slide.SelectSingleNode("//*[local-name()='bg']//*[local-name()='srgbClr']").val
        $title = $slide.SelectSingleNode("//*[local-name()='r' and *[local-name()='t']='Quarterly Business Review']")
        $foreground = $title.SelectSingleNode("*[local-name()='rPr']//*[local-name()='srgbClr']").val
        $first = Get-RecipeLuminance $foreground
        $second = Get-RecipeLuminance $background
        (($first, $second | Measure-Object -Maximum).Maximum + 0.05) / (($first, $second | Measure-Object -Minimum).Minimum + 0.05) | Should -BeGreaterOrEqual 4.5
    }

    It 'keeps all explicitly positioned slide content inside the presentation bounds' {
        foreach ($name in 'Quarterly-Business-Review.pptx', 'Service-Status.pptx') {
            $path = $artifacts[$name]
            $presentation = Get-ZipXmlDocumentLocal -Path $path -Entry 'ppt/presentation.xml'
            $size = $presentation.SelectSingleNode("//*[local-name()='sldSz']")
            foreach ($entry in Get-ZipEntriesLocal -Path $path | Where-Object { $_ -match '^ppt/slides/slide\d+\.xml$' }) {
                $slide = Get-ZipXmlDocumentLocal -Path $path -Entry $entry
                foreach ($transform in $slide.SelectNodes("//*[local-name()='xfrm']")) {
                    $offset = $transform.SelectSingleNode("*[local-name()='off']")
                    $extent = $transform.SelectSingleNode("*[local-name()='ext']")
                    if (-not $offset -or -not $extent) { continue }
                    [long]$offset.x | Should -BeGreaterOrEqual 0 -Because "$name $entry must not clip its left edge"
                    [long]$offset.y | Should -BeGreaterOrEqual 0 -Because "$name $entry must not clip its top edge"
                    ([long]$offset.x + [long]$extent.cx) | Should -BeLessOrEqual ([long]$size.cx) -Because "$name $entry must not clip its right edge"
                    ([long]$offset.y + [long]$extent.cy) | Should -BeLessOrEqual ([long]$size.cy) -Because "$name $entry must not clip its bottom edge"
                }
            }
        }
    }

    It 'uses one numeric measure per review chart instead of mixing units or owner names' {
        foreach ($name in 'Service-Status.xlsx', 'Service-Status.pptx', 'Quarterly-Business-Review.pptx') {
            $path = $artifacts[$name]
            $entries = @(Get-ZipEntriesLocal -Path $path | Where-Object { $_ -match '/charts/chart\d+\.xml$' })
            $entries.Count | Should -BeGreaterThan 0
            foreach ($entry in $entries) {
                $chart = Get-ZipXmlDocumentLocal -Path $path -Entry $entry
                @($chart.SelectNodes("//*[local-name()='ser']")).Count | Should -Be 1 -Because "$name should compare a single measure with its own unit"
                if ($name -like 'Service-Status.*') {
                    $values = @($chart.SelectNodes("//*[local-name()='ser']/*[local-name()='val']//*[local-name()='pt']/*[local-name()='v']") | ForEach-Object InnerText)
                    ($values -join ',') | Should -Be '1,4,7' -Because 'the service chart compares incident counts'
                }
            }
        }
    }

    It 'calculates and renders invoice totals from its decimal line items' {
        $text = Get-OfficePdfText -Path $artifacts['Service-Invoice.pdf']
        $text | Should -Match 'Subtotal:\s*\$4,350\.00'
        $text | Should -Match 'Illustrative tax \(20%\):\s*\$870\.00'
        $text | Should -Match 'Total due:\s*\$5,220\.00'
    }
}
