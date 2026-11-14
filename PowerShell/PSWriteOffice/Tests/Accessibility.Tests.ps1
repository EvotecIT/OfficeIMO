BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    function New-MergedWorkbook([string] $Path) {
        $workbook = [OfficeIMO.Excel.ExcelDocument]::Create($Path)
        try {
            $sheet = $workbook.AddSheet('Data')
            $sheet.Cell(1, 1).SetValue('Total') | Out-Null
            $sheet.Range('A1:B1').Merge() | Out-Null
            $workbook.Save()
        } finally { $workbook.Dispose() }
    }
}

Describe 'Test-OfficeAccessibility' {
    It 'reports shared findings for files, keeps going past unsupported files and leaves the input unchanged' {
        $workbook = Join-Path $TestDrive 'merged.xlsx'
        $notes = Join-Path $TestDrive 'notes.txt'
        New-MergedWorkbook $workbook
        Set-Content -LiteralPath $notes -Value 'Plain text'
        $hash = (Get-FileHash -LiteralPath $workbook).Hash

        $reports = @(Get-ChildItem -LiteralPath $workbook, $notes | Test-OfficeAccessibility -ErrorAction SilentlyContinue -ErrorVariable failures)

        $reports.Count | Should -Be 1
        $reports[0].Path | Should -Be $workbook
        $reports[0].Format | Should -Be 'Excel'
        $reports[0].IsSuccessful | Should -BeTrue
        $merged = @($reports[0].Findings | Where-Object Code -EQ 'accessibility-merged-cells')
        $merged.Count | Should -Be 1
        $merged[0].Location | Should -Be 'Data!A1:B1'
        $merged[0].Severity.ToString() | Should -Be 'Warning'
        @($failures).Count | Should -Be 1
        $failures[0].FullyQualifiedErrorId | Should -BeLike 'accessibility-unsupported-format*'
        (Get-FileHash -LiteralPath $workbook).Hash | Should -Be $hash
    }

    It 'applies the strict profile, Quiet thresholds and SARIF output' {
        $workbook = Join-Path $TestDrive 'strict.xlsx'
        New-MergedWorkbook $workbook

        Test-OfficeAccessibility -Path $workbook -Quiet | Should -BeTrue
        Test-OfficeAccessibility -Path $workbook -Quiet -IncludeWarnings | Should -BeFalse
        $strict = Test-OfficeAccessibility -Path $workbook -Strict
        $strict.IsSuccessful | Should -BeFalse
        @($strict.Findings | Where-Object Code -EQ 'accessibility-missing-document-title').Count | Should -Be 1

        $sarif = Test-OfficeAccessibility -Path $workbook -Sarif | ConvertFrom-Json
        $sarif.version | Should -Be '2.1.0'
        $sarif.runs[0].results.ruleId | Should -Contain 'accessibility-merged-cells'
    }

    It 'inspects an open document including unsaved edits' {
        $workbook = Join-Path $TestDrive 'open.xlsx'
        New-MergedWorkbook $workbook
        $document = Open-OfficeExcel -Path $workbook
        try {
            $document.Sheets[0].Range('A3:B3').Merge() | Out-Null
            $report = Test-OfficeAccessibility -Document $document
            $report.Path | Should -BeNullOrEmpty
            @($report.Findings | Where-Object Code -EQ 'accessibility-merged-cells').Count | Should -Be 2
        } finally { Close-OfficeExcel -Document $document }
    }
}
