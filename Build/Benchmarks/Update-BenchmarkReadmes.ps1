<#
.SYNOPSIS
Refreshes the compact CSV and Excel benchmark tables embedded in project READMEs.

.DESCRIPTION
Selects representative, directly comparable benchmark rows, writes them in the
PSPublishModule comparison schema, and delegates marker-delimited Markdown
updates to Update-BenchmarkDocument. Use -Run to execute the focused CSV, Excel,
or both benchmark sets locally before refreshing the tables. With no -Run value,
the script only renders the last committed compact comparisons. Benchmarks are
never scheduled or run by CI.

.PARAMETER Run
Benchmark sets to execute locally before updating the READMEs. Accepts Csv,
Excel, or All. Omit this parameter to refresh only from committed compact data.

.PARAMETER CsvArtifactPath
Optional BenchmarkDotNet artifact directory or report for a fresh CSV run. When
omitted, the last committed compact CSV comparison is rendered again.

.PARAMETER ExcelSummaryPath
Optional Excel comparison-suite summary to select. When supplied, the compact
Excel comparison is rebuilt from that run. If the compact artifact does not yet
exist, the committed current summary is used as the initial source.

.EXAMPLE
./Build/Benchmarks/Update-BenchmarkReadmes.ps1

.EXAMPLE
./Build/Benchmarks/Update-BenchmarkReadmes.ps1 -Run All
#>
[CmdletBinding()]
param(
    [ValidateSet("Csv", "Excel", "All")]
    [string[]] $Run = @(),
    [string] $CsvArtifactPath,
    [string] $ExcelSummaryPath = "./Docs/benchmarks/comparison-current/officeimo.excel.comparison-summary.json",
    [string] $OutputDirectory = "./Docs/benchmarks/readme-current"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "../.."))

function Resolve-RepositoryPath {
    param([Parameter(Mandatory)] [string] $Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    return [System.IO.Path]::GetFullPath((Join-Path $repositoryRoot $Path))
}

Push-Location $repositoryRoot
try {
    Import-Module PSPublishModule -MinimumVersion 3.0.57 -Force -ErrorAction Stop

    $outputRoot = Resolve-RepositoryPath $OutputDirectory
    [System.IO.Directory]::CreateDirectory($outputRoot) | Out-Null

    function Invoke-DotNetBenchmark {
        param(
            [Parameter(Mandatory)] [string] $Name,
            [Parameter(Mandatory)] [string[]] $Arguments
        )

        Write-Host "Running the focused $Name README benchmarks locally..."
        & dotnet @Arguments
        if ($LASTEXITCODE -ne 0) {
            throw "$Name benchmark execution failed with exit code $LASTEXITCODE."
        }
    }

    $runCsv = $Run -contains "Csv" -or $Run -contains "All"
    $runExcel = $Run -contains "Excel" -or $Run -contains "All"
    $artifactRoot = Resolve-RepositoryPath "./.benchmark-artifacts/readme-refresh"

    if ($runCsv) {
        $csvRunPath = Join-Path $artifactRoot "csv"
        Remove-Item -LiteralPath $csvRunPath -Recurse -Force -ErrorAction SilentlyContinue
        Invoke-DotNetBenchmark -Name "CSV" -Arguments @(
            "run", "-c", "Release", "--framework", "net8.0",
            "--project", "OfficeIMO.CSV.Benchmarks/OfficeIMO.CSV.Benchmarks.csproj", "--",
            "--filter",
            "*CsvWideBenchmarks.OfficeIMO_ReadTextFieldSpanVisitorSkipHeader*",
            "*CsvWideBenchmarks.Sep_ReadFieldSpans*",
            "*CsvWideBenchmarks.Sylvan_ReadFieldSpans*",
            "*CsvWideBenchmarks.OfficeIMO_WriteProjectedRows*",
            "*CsvWideBenchmarks.OfficeIMO_WriteDataReader*",
            "*CsvWideBenchmarks.OfficeIMO_WriteValidatedTextRows*",
            "*CsvWideBenchmarks.CsvHelper_WriteProjectedRows*",
            "*CsvWideBenchmarks.CsvHelper_WriteTextRows*",
            "*CsvWideBenchmarks.Sep_WriteProjectedRows*",
            "*CsvWideBenchmarks.Sylvan_WriteProjectedRows*",
            "*CsvWideBenchmarks.Sylvan_WriteTextRows*",
            "*CsvWideBenchmarks.Dataplat_WriteProjectedRows*",
            "*CsvWideBenchmarks.Dataplat_WriteFromReader*",
            "*CsvWideBenchmarks.Dataplat_WriteTextRows*",
            "--warmupCount", "3", "--iterationCount", "9",
            "--artifacts", $csvRunPath
        )
        $CsvArtifactPath = "./.benchmark-artifacts/readme-refresh/csv"
    }

    if ($runExcel) {
        $excelRunPath = Join-Path $artifactRoot "excel"
        Remove-Item -LiteralPath $excelRunPath -Recurse -Force -ErrorAction SilentlyContinue
        Invoke-DotNetBenchmark -Name "Excel" -Arguments @(
            "run", "-c", "Release", "--framework", "net8.0",
            "--project", "OfficeIMO.Excel.Benchmarks/OfficeIMO.Excel.Benchmarks.csproj", "--",
            "comparison-suite", "--out-dir", $excelRunPath, "--row-set", "25000",
            "--scenario", "realworld-report-all-in-one,write-datareader-table,read-objects-stream,write-datareader-compact-package",
            "--skip-package-profile", "--skip-dense-helloworld", "--skip-legacy-epplus",
            "--warmup", "20", "--iterations", "9"
        )
        $ExcelSummaryPath = "./.benchmark-artifacts/readme-refresh/excel/officeimo.excel.comparison-summary.json"
    }

    function Write-ComparisonArtifact {
        param(
            [Parameter(Mandatory)] [string] $Path,
            [Parameter(Mandatory)] [System.Collections.IDictionary] $Metadata,
            [Parameter(Mandatory)] [object[]] $Rows
        )

        $payload = [ordered]@{
            metadata = $Metadata
            comparison = $Rows
        }
        $json = $payload | ConvertTo-Json -Depth 8
        [System.IO.File]::WriteAllText($Path, $json + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
    }

    function New-ComparisonRow {
        param(
            [Parameter(Mandatory)] [string] $Scenario,
            [Parameter(Mandatory)] [string] $Operation,
            [Parameter(Mandatory)] [string] $Engine,
            [Parameter(Mandatory)] [string] $BaselineEngine,
            [Parameter(Mandatory)] [double] $Actual,
            [Parameter(Mandatory)] [double] $Baseline,
            [Parameter(Mandatory)] [string] $Metric,
            [Parameter(Mandatory)] [System.Collections.IDictionary] $Variables,
            [Parameter(Mandatory)] [string] $RuntimeHost,
            [ValidateRange(0, 1)] [double] $TieTolerance = 0.05
        )

        [ordered]@{
            suite = "OfficeIMO README highlights"
            scenario = $Scenario
            operation = $Operation
            host = $RuntimeHost
            os = ""
            runMode = ""
            variables = $Variables
            engine = $Engine
            baselineEngine = $BaselineEngine
            status = "Succeeded"
            actual = $Actual
            baseline = $Baseline
            ratio = if ($Baseline -gt 0) { $Actual / $Baseline } else { $null }
            metric = $Metric
            tieTolerance = $TieTolerance
        }
    }

    $excelComparisonPath = Join-Path $outputRoot "officeimo.excel.comparison.json"
    if ($runExcel -or $PSBoundParameters.ContainsKey("ExcelSummaryPath") -or -not [System.IO.File]::Exists($excelComparisonPath)) {
        $excelSource = Resolve-RepositoryPath $ExcelSummaryPath
        if (-not [System.IO.File]::Exists($excelSource)) {
            throw "Excel benchmark summary was not found: $excelSource"
        }

        $excel = Get-Content -LiteralPath $excelSource -Raw -Encoding UTF8 | ConvertFrom-Json
        $excelSelections = [ordered]@{
            "Feature-rich report to XLSX" = @{ Scenario = "realworld-report-all-in-one"; Operation = "Create"; Libraries = @("OfficeIMO.Excel", "EPPlus") }
            "Styled DataReader table to XLSX" = @{ Scenario = "write-datareader-table"; Operation = "Write"; Libraries = @("OfficeIMO.Excel", "ClosedXML", "EPPlus") }
            "Typed objects streamed from XLSX" = @{ Scenario = "read-objects-stream"; Operation = "Read"; Libraries = @("OfficeIMO.Excel", "ClosedXML", "EPPlus", "Sylvan.Data.Excel") }
            "Compact DataReader to XLSX" = @{ Scenario = "write-datareader-compact-package"; Operation = "Write"; Libraries = @("OfficeIMO.Excel", "SpreadCheetah", "LargeXlsx", "Sylvan.Data.Excel") }
        }
        $excelRows = [System.Collections.Generic.List[object]]::new()
        $excelSnapshot = ([DateTimeOffset] $excel.GeneratedAtUtc).UtcDateTime.ToString("yyyy-MM-dd")
        foreach ($selection in $excelSelections.GetEnumerator()) {
            $sourceRows = @($excel.Rows | Where-Object {
                $_.ArtifactKind -eq "speed-comparison" -and
                $_.RowCount -eq 25000 -and
                $_.Scenario -eq $selection.Value.Scenario -and
                $_.Library -in $selection.Value.Libraries
            })
            $baselineRow = $sourceRows | Where-Object Library -eq "OfficeIMO.Excel" | Select-Object -First 1
            if ($null -eq $baselineRow) { throw "Excel baseline is missing for '$($selection.Value.Scenario)'." }
            foreach ($row in $sourceRows) {
                $variables = [ordered]@{
                    Format = ".xlsx"
                    Rows = "25,000"
                    Snapshot = $excelSnapshot
                    Runner = "rotated local"
                }
                if ($excel.WarmupIterations -gt 0) { $variables.Warmups = [string] $excel.WarmupIterations }
                if ($excel.MeasuredIterations -gt 0) { $variables.MeasuredIterations = [string] $excel.MeasuredIterations }
                $excelRows.Add((New-ComparisonRow -Scenario $selection.Key -Operation $selection.Value.Operation `
                    -Engine $row.Library -BaselineEngine "OfficeIMO.Excel" -Actual ([double] $row.MedianMilliseconds) `
                    -Baseline ([double] $baselineRow.MedianMilliseconds) -Metric "MedianMs" -RuntimeHost ".NET 8" `
                    -Variables $variables))
            }
        }

        Write-ComparisonArtifact -Path $excelComparisonPath -Metadata ([ordered]@{
            generatedAtUtc = $excel.GeneratedAtUtc
            source = $ExcelSummaryPath
            warmupIterations = $excel.WarmupIterations
            measuredIterations = $excel.MeasuredIterations
            note = $excel.Notes
        }) -Rows $excelRows.ToArray()
    }

    $csvComparisonPath = Join-Path $outputRoot "officeimo.csv.comparison.json"
    if ($CsvArtifactPath) {
        $csvSource = Resolve-RepositoryPath $CsvArtifactPath
        $csvRun = Import-BenchmarkResult -Path $csvSource -Suite "OfficeIMO.CSV README highlights"
        $csvEngines = [ordered]@{
            "OfficeIMO_ReadTextFieldSpanVisitorSkipHeader" = "OfficeIMO.CSV"
            "Sep_ReadFieldSpans" = "Sep"
            "Sylvan_ReadFieldSpans" = "Sylvan.Data.Csv"
            "OfficeIMO_WriteProjectedRows" = "OfficeIMO.CSV"
            "OfficeIMO_WriteDataReader" = "OfficeIMO.CSV"
            "OfficeIMO_WriteValidatedTextRows" = "OfficeIMO.CSV"
            "CsvHelper_WriteProjectedRows" = "CsvHelper"
            "CsvHelper_WriteTextRows" = "CsvHelper"
            "Sep_WriteProjectedRows" = "Sep"
            "Sylvan_WriteProjectedRows" = "Sylvan.Data.Csv"
            "Sylvan_WriteTextRows" = "Sylvan.Data.Csv"
            "Dataplat_WriteProjectedRows" = "Dataplat.Dbatools.Csv"
            "Dataplat_WriteFromReader" = "Dataplat.Dbatools.Csv"
            "Dataplat_WriteTextRows" = "Dataplat.Dbatools.Csv"
        }
        $csvSelections = [ordered]@{
            "Wide field-span CSV read" = @{ Operation = "Read every field"; Baseline = "OfficeIMO_ReadTextFieldSpanVisitorSkipHeader"; Methods = @("OfficeIMO_ReadTextFieldSpanVisitorSkipHeader", "Sep_ReadFieldSpans", "Sylvan_ReadFieldSpans"); Contract = "field spans" }
            "Wide projected-array CSV write" = @{ Operation = "Format and write rows"; Baseline = "OfficeIMO_WriteProjectedRows"; Methods = @("OfficeIMO_WriteProjectedRows", "CsvHelper_WriteProjectedRows", "Dataplat_WriteProjectedRows"); Contract = "projected object arrays" }
            "Wide DataReader CSV write" = @{ Operation = "Format and write rows"; Baseline = "OfficeIMO_WriteDataReader"; Methods = @("OfficeIMO_WriteDataReader", "Sylvan_WriteProjectedRows", "Dataplat_WriteFromReader"); Contract = "IDataReader" }
            "Wide validated text-row CSV write" = @{ Operation = "Validate and write rows"; Baseline = "OfficeIMO_WriteValidatedTextRows"; Methods = @("OfficeIMO_WriteValidatedTextRows", "CsvHelper_WriteTextRows", "Sep_WriteProjectedRows", "Sylvan_WriteTextRows", "Dataplat_WriteTextRows"); Contract = "preformatted text with escaping" }
        }
        $selected = @($csvRun.Summary | Where-Object {
            $_.Scenario -in $csvEngines.Keys -and
            -not $_.Variables.ContainsKey("Shape") -and
            $_.Variables["RowCount"] -eq "25000"
        })
        $csvRows = [System.Collections.Generic.List[object]]::new()
        $csvSnapshot = (Get-Item -Force -LiteralPath $csvSource).LastWriteTimeUtc.ToString("yyyy-MM-dd")
        foreach ($selection in $csvSelections.GetEnumerator()) {
            $scenarioRows = @($selected | Where-Object Scenario -in $selection.Value.Methods)
            $baselineRow = $scenarioRows | Where-Object Scenario -eq $selection.Value.Baseline | Select-Object -First 1
            if ($null -eq $baselineRow) { throw "CSV baseline is missing for '$($selection.Key)'." }
            foreach ($row in $scenarioRows) {
                $csvRows.Add((New-ComparisonRow -Scenario $selection.Key -Operation $selection.Value.Operation `
                    -Engine $csvEngines[$row.Scenario] -BaselineEngine "OfficeIMO.CSV" -Actual ([double] $row.MeanMs) `
                    -Baseline ([double] $baselineRow.MeanMs) -Metric "MeanMs" -RuntimeHost ".NET 8" `
                    -Variables ([ordered]@{ Format = "CSV"; Rows = "25,000"; Shape = "wide"; Contract = $selection.Value.Contract; Snapshot = $csvSnapshot; Runner = "BenchmarkDotNet local" })))
            }
        }
        Write-ComparisonArtifact -Path $csvComparisonPath -Metadata ([ordered]@{
            generatedAtUtc = [DateTimeOffset]::UtcNow
            source = $CsvArtifactPath
            note = "Focused BenchmarkDotNet run; read lanes traverse every field and write lanes semantically validate every output value."
        }) -Rows $csvRows.ToArray()
    } elseif (-not [System.IO.File]::Exists($csvComparisonPath)) {
        throw "No CSV comparison artifact exists. Pass -CsvArtifactPath after a focused BenchmarkDotNet run."
    }

    foreach ($target in @(
        @{ Path = "./OfficeIMO.CSV/README.md"; Block = "officeimo-csv-benchmark-table"; Data = $csvComparisonPath },
        @{ Path = "./OfficeIMO.CSV.Benchmarks/README.md"; Block = "officeimo-csv-benchmark-table"; Data = $csvComparisonPath },
        @{ Path = "./OfficeIMO.Excel/README.md"; Block = "officeimo-excel-benchmark-table"; Data = $excelComparisonPath },
        @{ Path = "./OfficeIMO.Excel.Benchmarks/README.md"; Block = "officeimo-excel-benchmark-table"; Data = $excelComparisonPath }
    )) {
        Update-BenchmarkDocument -Path $target.Path -BlockId $target.Block -ComparisonPath $target.Data -Renderer ComparisonTable | Out-Null
    }
} finally {
    Pop-Location
}
