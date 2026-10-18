. (Join-Path $PSScriptRoot '..\Excel\excel-performance.helpers.ps1')

$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$suiteName = Get-BenchmarkInput Suite Standard
$rowCounts = Assert-ExcelBenchmarkRowCount -RowCount (Get-BenchmarkInput RowCount (Get-ExcelBenchmarkDefaultRowCount -Suite $suiteName) -Int)

New-BenchmarkSuite 'csv-performance' -OutputRoot (Join-Path $repositoryRoot 'Ignore\Benchmarks\CsvPerformance') {
    Set-BenchmarkPolicy -Warmup (Get-ExcelBenchmarkWarmupCount -Suite $suiteName) -Iterations (Get-CsvBenchmarkIterationCount -Suite $suiteName) -Order GroupedRotated -MemoryCleanup BeforeIteration -OutlierMode None
    Set-BenchmarkProfile Current -Cleanup KeepOnFailure
    Add-BenchmarkCaseSource (Get-CsvBenchmarkCase -Suite $suiteName)
    Add-BenchmarkAxis RowCount $rowCounts

    Set-BenchmarkSetup {
        param($case, $run)

        $run.RepositoryRoot = $repositoryRoot
        $run.WorksheetName = 'Data'
        $extension = Get-ExcelBenchmarkExtension -Case $case
        $run.Path = $run.OutputPath + $extension
        $run.SourcePath = $run.OutputPath + '.source' + $extension
        Initialize-ExcelBenchmarkEngine -Engine $case.Engine -Run $run
    }

    Set-BenchmarkDataFactory {
        param($case, $run)

        $profile = Get-ExcelBenchmarkData -Profile $case.DataProfile -Count ([int]$case.RowCount)
        $run.Payload = $profile.Data
        $run.ExpectedRows = [int]$case.RowCount
        $run.ColumnCount = $profile.ColumnCount
        Initialize-ExcelBenchmarkInput -Case $case -Run $run
    }

    Add-BenchmarkSkipRule {
        param($case)

        -not (Test-CsvBenchmarkEngineSupport -Engine $case.Engine -Case $case)
    }

    Add-BenchmarkEngine PSWriteOffice {
        Add-BenchmarkOperation Run {
            param($case, $run)
            Invoke-ExcelBenchmarkOperation -Engine PSWriteOffice -Case $case -Run $run
        }
    }

    Add-BenchmarkEngine NativeCsv {
        Add-BenchmarkOperation Run {
            param($case, $run)
            Invoke-ExcelBenchmarkOperation -Engine NativeCsv -Case $case -Run $run
        }
    }

    Add-BenchmarkValidation {
        param($case, $run)

        Test-CsvBenchmarkOutput -Case $case -Run $run
    }

    Add-BenchmarkMetric RowsProcessed {
        param($case, $run)

        $run.RowsProcessed
    }

    Add-BenchmarkMetric RowsPerSecond {
        param($case, $run)

        if ($run.DurationMs -le 0) {
            return 0
        }

        [double] $case.RowCount / ($run.DurationMs / 1000)
    }

    Add-BenchmarkComparison Engine -Baseline PSWriteOffice -Metric MedianMs -TieTolerance 0.05 -RequireBaselineFastest
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
