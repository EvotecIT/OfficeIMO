. (Join-Path $PSScriptRoot 'excel-performance.helpers.ps1')

$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$suiteName = Get-BenchmarkInput Suite Standard
$rowCounts = Assert-ExcelBenchmarkRowCount -RowCount (Get-BenchmarkInput RowCount (Get-ExcelBenchmarkDefaultRowCount -Suite $suiteName) -Int)
$skipWorkbookValidation = Get-BenchmarkInput SkipWorkbookValidation false -Bool
$skipImportExcelInstall = Get-BenchmarkInput SkipImportExcelInstall false -Bool
$skipExcelFastInstall = Get-BenchmarkInput SkipExcelFastInstall false -Bool

New-BenchmarkSuite 'excel-performance' -OutputRoot (Join-Path $repositoryRoot 'Ignore\Benchmarks\ExcelPerformance') {
    Set-BenchmarkPolicy -Warmup (Get-ExcelBenchmarkWarmupCount -Suite $suiteName) -Iterations (Get-ExcelBenchmarkIterationCount -Suite $suiteName) -Order GroupedRotated -MemoryCleanup BeforeIteration -OutlierMode None
    Set-BenchmarkProfile Current -Cleanup KeepOnFailure
    Add-BenchmarkCaseSource (Get-ExcelBenchmarkCase -Suite $suiteName)
    Add-BenchmarkAxis RowCount $rowCounts

    Set-BenchmarkSetup {
        param($case, $run)

        $run.RepositoryRoot = $repositoryRoot
        $run.WorksheetName = 'Data'
        $run.Path = $run.OutputPath + (Get-ExcelBenchmarkExtension -Case $case)
        $run.SourcePath = $run.OutputPath + '.source.csv'
        $run.SkipWorkbookValidation = $skipWorkbookValidation
        $run.SkipImportExcelInstall = $skipImportExcelInstall
        $run.SkipExcelFastInstall = $skipExcelFastInstall
        $run.Range = Get-ExcelBenchmarkRange -ColumnCount (Get-ExcelBenchmarkColumnCount -Profile $case.DataProfile) -Rows ([int]$case.RowCount)
        $run.RangeEndCell = Get-ExcelBenchmarkRangeEndCell -ColumnCount (Get-ExcelBenchmarkColumnCount -Profile $case.DataProfile) -Rows ([int]$case.RowCount)
        Initialize-ExcelBenchmarkEngine -Engine $case.Engine -Run $run
    }

    Set-BenchmarkDataFactory {
        param($case, $run)

        $profile = Get-ExcelBenchmarkData -Profile $case.DataProfile -Count ([int]$case.RowCount)
        $run.Payload = $profile.Data
        $run.ExpectedRows = [int]$case.RowCount
        $run.ColumnCount = $profile.ColumnCount
        $run.WorksheetName = $profile.WorksheetName
        $run.Range = Get-ExcelBenchmarkRange -ColumnCount $profile.ColumnCount -Rows ([int]$case.RowCount)
        $run.RangeEndCell = Get-ExcelBenchmarkRangeEndCell -ColumnCount $profile.ColumnCount -Rows ([int]$case.RowCount)
        Initialize-ExcelBenchmarkInput -Case $case -Run $run
    }

    Add-BenchmarkSkipRule {
        param($case)

        return ([string] $case.SupportedEngines -split ',') -notcontains [string] $case.Engine
    }

    Add-BenchmarkEngine PSWriteOffice {
        Add-BenchmarkOperation Run {
            param($case, $run)
            Invoke-ExcelBenchmarkOperation -Engine PSWriteOffice -Case $case -Run $run
        }
    }

    Add-BenchmarkEngine ImportExcel {
        Add-BenchmarkOperation Run {
            param($case, $run)
            Invoke-ExcelBenchmarkOperation -Engine ImportExcel -Case $case -Run $run
        }
    }

    Add-BenchmarkEngine ExcelFast {
        Add-BenchmarkOperation Run {
            param($case, $run)
            Invoke-ExcelBenchmarkOperation -Engine ExcelFast -Case $case -Run $run
        }
    }

    Add-BenchmarkValidation {
        param($case, $run)

        Test-ExcelBenchmarkOutput -Case $case -Run $run
    }

    Add-BenchmarkComparison Engine -Baseline PSWriteOffice -Metric MedianMs -TieTolerance 0.05 -RequireBaselineFastest
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
