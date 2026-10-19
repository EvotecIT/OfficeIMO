BeforeAll {
    . (Join-Path $PSScriptRoot '..\Benchmarks\Excel\excel-performance.core.ps1')
    $script:csvBenchmarkText = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\Benchmarks\Csv\csv-performance.benchmark.ps1') -Raw
    $script:excelBenchmarkText = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\Benchmarks\Excel\excel-performance.benchmark.ps1') -Raw
}

Describe 'Benchmark measurement policy' {
    It 'uses stable repeated measurements for short CSV release-gate lanes' {
        Get-ExcelBenchmarkWarmupCount -Suite Smoke | Should -Be 3
        Get-CsvBenchmarkIterationCount -Suite Smoke | Should -Be 25
        Get-ExcelBenchmarkWarmupCount -Suite Standard | Should -Be 3
        Get-CsvBenchmarkIterationCount -Suite Standard | Should -Be 25
        Get-CsvBenchmarkIterationCount -Suite Full | Should -Be 25
    }

    It 'keeps slower Excel workloads repeated without copying the CSV sample depth' {
        Get-ExcelBenchmarkIterationCount -Suite Smoke | Should -Be 5
        Get-ExcelBenchmarkIterationCount -Suite Standard | Should -Be 5
        Get-ExcelBenchmarkIterationCount -Suite Full | Should -Be 5
    }

    It 'keeps super-large diagnostics bounded but statistically meaningful' {
        Get-ExcelBenchmarkWarmupCount -Suite SuperLarge | Should -Be 1
        Get-CsvBenchmarkIterationCount -Suite SuperLarge | Should -Be 7
        Get-ExcelBenchmarkIterationCount -Suite SuperLarge | Should -Be 3
    }

    It 'groups equivalent engines and cleans managed memory outside timed operations' {
        $script:csvBenchmarkText | Should -Match '-Order GroupedRotated'
        $script:csvBenchmarkText | Should -Match '-MemoryCleanup BeforeIteration'
        $script:excelBenchmarkText | Should -Match '-Order GroupedRotated'
        $script:excelBenchmarkText | Should -Match '-MemoryCleanup BeforeIteration'
    }

    It 'compares ExcelFast on equivalent default workbook shapes only' {
        $cases = @(Get-ExcelBenchmarkCase -Suite Smoke)
        $defaultCase = $cases | Where-Object Name -EQ 'objects-default'
        $textCase = $cases | Where-Object Name -EQ 'text-objects-default'
        $tableCase = $cases | Where-Object Name -EQ 'objects-table'

        Test-ExcelBenchmarkEngineCaseSupport -Engine ExcelFast -Case $defaultCase | Should -BeFalse
        Test-ExcelBenchmarkEngineCaseSupport -Engine ExcelFast -Case $textCase | Should -BeTrue
        Test-ExcelBenchmarkEngineCaseSupport -Engine ExcelFast -Case $tableCase | Should -BeFalse
    }

    It 'separates Excel reader, DataTable, and PowerShell object projection lanes' {
        $cases = @(Get-ExcelBenchmarkCase -Suite Standard)
        $projectionCases = @(
            $cases | Where-Object Name -In @(
                'import-datatable-full',
                'import-datatable-range',
                'import-datareader-full',
                'import-datareader-range'
            )
        )

        $projectionCases.Count | Should -Be 4
        foreach ($case in $projectionCases) {
            Test-ExcelBenchmarkEngineCaseSupport -Engine PSWriteOffice -Case $case | Should -BeTrue
            Test-ExcelBenchmarkEngineCaseSupport -Engine ImportExcel -Case $case | Should -BeFalse
            Test-ExcelBenchmarkEngineCaseSupport -Engine ExcelFast -Case $case | Should -BeFalse
        }
    }
}
