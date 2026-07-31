param(
    [ValidateSet('quick', 'full')]
    [string] $RunMode = 'quick',
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0',
    [ValidateSet('all', 'csv', 'csvwrite', 'xlsx', 'xlsxwrite', 'xlsb')]
    [string] $Workload = 'all',
    [string] $OutputRoot = (Join-Path ([System.IO.Path]::GetTempPath()) 'OfficeIMO\Benchmarks\Runs'),
    [string] $PowerForgeRoot = $env:POWERFORGE_ROOT,
    [ValidateSet('net8.0', 'net10.0')]
    [string] $PowerForgeFramework = 'net8.0',
    [switch] $Publish
)

$ErrorActionPreference = 'Stop'

if ($Publish -and $RunMode -ne 'full') {
    throw 'Only a full BenchmarkDotNet run can be marked publishable.'
}

. (Join-Path $PSScriptRoot 'BenchmarkEvidence.ps1')
if ([string]::IsNullOrWhiteSpace($PowerForgeRoot)) {
    Import-Module PSPublishModule -MinimumVersion 3.0.84 -Force
} else {
    $powerForgeModule = Join-Path $PowerForgeRoot "PSPublishModule\bin\Release\$PowerForgeFramework\PSPublishModule.dll"
    if (-not (Test-Path -LiteralPath $powerForgeModule -PathType Leaf)) {
        throw "The local PowerForge binary was not found at '$powerForgeModule'. Build PSPublishModule for $PowerForgeFramework in Release configuration first."
    }
    Import-Module $powerForgeModule -Force
}

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$platform = if ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::Windows)) {
    'windows'
} elseif ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::Linux)) {
    'linux'
} elseif ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::OSX)) {
    'macos'
} else {
    throw 'Unsupported benchmark platform.'
}

$definitions = [ordered]@{
    csv = [pscustomobject]@{
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*MarkPflug65KCsvBenchmarks*'
        ComparisonId = "markpflug-65k-csv-decoded-$Framework"
        Suite = 'OfficeIMO.CSV.MarkPflug65K'
        ExpectedCases = @('OfficeIMO', 'Sep', 'Sylvan', 'CsvHelper', 'DataplatDbatools', 'LumenWorks')
    }
    csvwrite = [pscustomobject]@{
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvDataReaderWriteBenchmarks*'
        ComparisonId = "csv-25k-datareader-write-$Framework"
        Suite = 'OfficeIMO.CSV.DataReaderWrite25K'
        ExpectedCases = @(
            'OfficeIMO_WriteDataReader|RowCount=25000&Shape=Mixed'
            'OfficeIMO_WriteDataReader|RowCount=25000&Shape=Quoted'
            'OfficeIMO_WriteDataReader|RowCount=25000&Shape=Multiline'
            'Sylvan_WriteDataReader|RowCount=25000&Shape=Mixed'
            'Sylvan_WriteDataReader|RowCount=25000&Shape=Quoted'
            'Sylvan_WriteDataReader|RowCount=25000&Shape=Multiline'
        )
    }
    xlsx = [pscustomobject]@{
        Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
        Filter = '*MarkPflug65KXlsxBenchmarks*'
        ComparisonId = "markpflug-65k-xlsx-typed-$Framework"
        Suite = 'OfficeIMO.Excel.Xlsx.MarkPflug65K'
        ExpectedCases = @('OfficeIMO', 'Sylvan', 'ExcelDataReader', 'ClosedXML', 'EPPlus', 'MiniExcel')
    }
    xlsxwrite = [pscustomobject]@{
        Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
        Filter = '*ExcelDataReaderWriteBenchmarks*'
        ComparisonId = "xlsx-25k-datareader-write-$Framework"
        Suite = 'OfficeIMO.Excel.DataReaderWrite25K'
        ExpectedCases = @(
            'OfficeIMO|RowCount=25000'
            'SpreadCheetah|RowCount=25000'
            'Sylvan|RowCount=25000'
            'LargeXlsx|RowCount=25000'
        )
    }
    xlsb = [pscustomobject]@{
        Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
        Filter = '*MarkPflug65KXlsbBenchmarks*'
        ComparisonId = "markpflug-65k-xlsb-typed-$Framework"
        Suite = 'OfficeIMO.Excel.Xlsb.MarkPflug65K'
        ExpectedCases = @('OfficeIMO', 'Sylvan', 'ExcelDataReader')
    }
}

$selected = if ($Workload -eq 'all') {
    @($definitions.Keys)
} else {
    @($Workload)
}

$stamp = [DateTimeOffset]::UtcNow.ToString('yyyyMMdd-HHmmss')
$staticRoot = Join-Path $repositoryRoot 'Website\static\data\benchmarks\library-comparisons'
$catalogPath = Join-Path $staticRoot 'index.json'
$catalogEligible = $RunMode -eq 'quick' -or [bool] $Publish
New-Item -ItemType Directory -Force -Path $OutputRoot, $staticRoot | Out-Null

$gitSha = (& git -C $repositoryRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($gitSha)) {
    throw 'Unable to resolve the OfficeIMO source commit for benchmark provenance.'
}
$gitDirty = @(& git -C $repositoryRoot status --porcelain --untracked-files=normal).Count -gt 0
if ($catalogEligible -and $gitDirty) {
    throw 'Cataloged benchmark evidence requires a clean Git worktree so the recorded source commit identifies the measured code exactly.'
}

$measurements = [System.Collections.Generic.List[object]]::new()
foreach ($name in $selected) {
    $definition = $definitions[$name]
    $artifactsPath = Join-Path $OutputRoot "$platform-$name-$RunMode-$stamp"
    New-Item -ItemType Directory -Force -Path $artifactsPath | Out-Null
    $provenanceMetadata = [ordered]@{
        'benchmark.workload.id' = $definition.ComparisonId
        'benchmark.workload.sourceCommit' = $gitSha
        'benchmark.workload.framework' = $Framework
    }
    $provenanceCapture = Start-BenchmarkProvenanceCapture `
        -SourceRoot $repositoryRoot `
        -ArtifactRoot $artifactsPath `
        -Metadata $provenanceMetadata `
        -RunMode $RunMode

    $arguments = @(
        'run',
        '-c', 'Release',
        '-f', $Framework,
        '--project', (Join-Path $repositoryRoot $definition.Project),
        '--',
        '--filter', $definition.Filter,
        '--artifacts', $artifactsPath
    )
    if ($RunMode -eq 'quick') {
        $arguments += @('--job', 'Dry')
    }

    Push-Location -LiteralPath $repositoryRoot
    try {
        & dotnet @arguments
        $benchmarkExitCode = $LASTEXITCODE
    } finally {
        Pop-Location
    }
    if ($benchmarkExitCode -ne 0) {
        throw "$name benchmark run failed with exit code $benchmarkExitCode."
    }
    $provenanceCapture |
        Complete-BenchmarkProvenanceCapture |
        Out-Null

    $result = Import-BenchmarkResult -Path $artifactsPath -Suite $definition.Suite
    $successfulCases = @(
        $result.Summary |
            Where-Object {
                $_.Status -in @('Success', 'Succeeded') -and
                $_.SampleCount -gt 0 -and
                $null -ne $_.MedianMs
            } |
            ForEach-Object { Get-BenchmarkEvidenceCaseIdentity -Row $_ } |
            Sort-Object -Unique
    )
    $missingCases = @(
        $definition.ExpectedCases |
            Where-Object { $_ -notin $successfulCases }
    )
    $unexpectedCases = @(
        $successfulCases |
            Where-Object { $_ -notin $definition.ExpectedCases }
    )
    if ($missingCases.Count -gt 0 -or $unexpectedCases.Count -gt 0) {
        $details = @()
        if ($missingCases.Count -gt 0) {
            $details += "missing successful cases: $($missingCases -join ', ')"
        }
        if ($unexpectedCases.Count -gt 0) {
            $details += "unexpected cases: $($unexpectedCases -join ', ')"
        }
        throw "$name benchmark evidence is incomplete ($($details -join '; '))."
    }

    $evidenceLocation = Get-BenchmarkEvidenceLocation `
        -ComparisonId $definition.ComparisonId `
        -Platform $platform `
        -RunMode $RunMode `
        -StaticRoot $staticRoot
    $normalizedPath = Join-Path $artifactsPath 'normalized-result.json'
    Write-BenchmarkEvidenceResult -Path $normalizedPath -InputObject $result

    $measurements.Add([pscustomobject]@{
        Workload = $name
        Definition = $definition
        Result = $result
        EvidenceLocation = $evidenceLocation
        ArtifactsPath = $artifactsPath
        NormalizedResult = $normalizedPath
    })
}

if (($measurements.Count -ne $selected.Count) -or
    (@($measurements | Where-Object { $null -eq $_.Result }).Count -gt 0)) {
    throw 'Benchmark measurement collection did not produce exactly one normalized result per selected workload.'
}

if ($catalogEligible) {
    foreach ($measurement in $measurements) {
        if (-not $Publish) {
            Write-BenchmarkEvidenceResult `
                -Path $measurement.EvidenceLocation.Path `
                -InputObject $measurement.Result
        }
        Update-BenchmarkEvidenceCatalog `
            -InputObject $measurement.Result `
            -Path $catalogPath `
            -ComparisonId $measurement.Definition.ComparisonId `
            -ResultPath $measurement.EvidenceLocation.ResultPath `
            -ResultArtifactPath $measurement.EvidenceLocation.Path `
            -RunMode $RunMode `
            -ExpectedPlatform windows, linux, macos `
            -Publish:$Publish | Out-Null
    }
}

$outputs = foreach ($measurement in $measurements) {
    [pscustomobject]@{
        Workload = $measurement.Workload
        Platform = $platform
        RunMode = $RunMode
        Publish = [bool] $Publish
        SourceCommit = $gitSha
        ArtifactsPath = $measurement.ArtifactsPath
        NormalizedResult = if ($catalogEligible) {
            $measurement.EvidenceLocation.Path
        } else {
            $measurement.NormalizedResult
        }
        EvidenceCatalog = if ($catalogEligible) { $catalogPath } else { $null }
    }
}

$outputs
