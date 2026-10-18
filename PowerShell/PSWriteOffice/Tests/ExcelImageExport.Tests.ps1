BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'Excel image export cmdlets' {
    It 'does not load a path-based workbook under WhatIf' {
        $missingPath = Join-Path $TestDrive 'missing.xlsx'

        { Export-OfficeExcelRangeImage -Path $missingPath -WorksheetName Data -Range A1:B2 -OutputPath (Join-Path $TestDrive 'range.png') -WhatIf } |
            Should -Not -Throw
        { Export-OfficeExcelChartImage -Path $missingPath -WorksheetName Data -ChartName Chart1 -OutputPath (Join-Path $TestDrive 'chart.png') -WhatIf } |
            Should -Not -Throw
    }

    It 'creates a nested output directory and reports the saved range image path' {
        $workbookPath = Join-Path $TestDrive 'range.xlsx'
        $outputPath = Join-Path $TestDrive 'range-output\nested\range.png'
        New-OfficeExcel -Path $workbookPath {
            ExcelSheet -Name Data {
                ExcelCell -Address A1 -Value 'Name'
                ExcelCell -Address B1 -Value 'Value'
                ExcelCell -Address A2 -Value 'Alpha'
                ExcelCell -Address B2 -Value 42
            }
        } | Out-Null

        $result = Export-OfficeExcelRangeImage -Path $workbookPath -WorksheetName Data -Range A1:B2 -OutputPath $outputPath -PassThru

        Test-Path -LiteralPath $outputPath | Should -BeTrue
        $result.SavedPath | Should -Be ([System.IO.Path]::GetFullPath($outputPath))
    }

    It 'creates a nested output directory and reports the saved chart image path' {
        $workbookPath = Join-Path $TestDrive 'chart.xlsx'
        $outputPath = Join-Path $TestDrive 'chart-output\nested\chart.png'
        $script:exportChartName = $null
        New-OfficeExcel -Path $workbookPath {
            ExcelSheet -Name Data {
                ExcelCell -Address A1 -Value 'Name'
                ExcelCell -Address B1 -Value 'Value'
                ExcelCell -Address A2 -Value 'Alpha'
                ExcelCell -Address B2 -Value 42
                $chart = Add-OfficeExcelChart -Range A1:B2 -Row 4 -Column 1 -Title 'Values' -PassThru
                $script:exportChartName = $chart.Name
            }
        } | Out-Null

        $result = Export-OfficeExcelChartImage -Path $workbookPath -WorksheetName Data -ChartName $script:exportChartName -OutputPath $outputPath -PassThru

        Test-Path -LiteralPath $outputPath | Should -BeTrue
        $result.SavedPath | Should -Be ([System.IO.Path]::GetFullPath($outputPath))
    }
}
