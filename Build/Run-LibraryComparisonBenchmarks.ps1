param(
    [ValidateSet('quick', 'full')]
    [string] $RunMode = 'quick',
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0',
    [ValidateSet('all', 'csv', 'xlsx', 'xlsb')]
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
    Import-Module PSPublishModule -MinimumVersion 3.0.81 -Force
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
        ExpectedScenarios = @('OfficeIMO', 'Sep', 'Sylvan', 'CsvHelper', 'DataplatDbatools', 'LumenWorks')
    }
    xlsx = [pscustomobject]@{
        Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
        Filter = '*MarkPflug65KXlsxBenchmarks*'
        ComparisonId = "markpflug-65k-xlsx-typed-$Framework"
        Suite = 'OfficeIMO.Excel.Xlsx.MarkPflug65K'
        ExpectedScenarios = @('OfficeIMO', 'Sylvan', 'ExcelDataReader', 'ClosedXML', 'EPPlus', 'MiniExcel')
    }
    xlsb = [pscustomobject]@{
        Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
        Filter = '*MarkPflug65KXlsbBenchmarks*'
        ComparisonId = "markpflug-65k-xlsb-typed-$Framework"
        Suite = 'OfficeIMO.Excel.Xlsb.MarkPflug65K'
        ExpectedScenarios = @('OfficeIMO', 'Sylvan', 'ExcelDataReader')
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
if ($Publish -and $gitDirty) {
    throw 'Publishable benchmark evidence requires a clean Git worktree so the recorded source commit identifies the measured code exactly.'
}

$measurements = [System.Collections.Generic.List[object]]::new()
foreach ($name in $selected) {
    $definition = $definitions[$name]
    $artifactsPath = Join-Path $OutputRoot "$platform-$name-$RunMode-$stamp"
    New-Item -ItemType Directory -Force -Path $artifactsPath | Out-Null
    $provenanceCapture = Start-BenchmarkProvenanceCapture `
        -SourceRoot $repositoryRoot `
        -ArtifactRoot $artifactsPath

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

    & dotnet @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$name benchmark run failed with exit code $LASTEXITCODE."
    }
    $provenanceCapture |
        Complete-BenchmarkProvenanceCapture |
        Out-Null

    $result = Import-BenchmarkResult -Path $artifactsPath -Suite $definition.Suite
    $successfulScenarios = @(
        $result.Summary |
            Where-Object {
                $_.Status -in @('Success', 'Succeeded') -and
                $_.SampleCount -gt 0 -and
                $null -ne $_.MedianMs
            } |
            ForEach-Object Scenario |
            Sort-Object -Unique
    )
    $missingScenarios = @(
        $definition.ExpectedScenarios |
            Where-Object { $_ -notin $successfulScenarios }
    )
    $unexpectedScenarios = @(
        $successfulScenarios |
            Where-Object { $_ -notin $definition.ExpectedScenarios }
    )
    if ($missingScenarios.Count -gt 0 -or $unexpectedScenarios.Count -gt 0) {
        $details = @()
        if ($missingScenarios.Count -gt 0) {
            $details += "missing successful scenarios: $($missingScenarios -join ', ')"
        }
        if ($unexpectedScenarios.Count -gt 0) {
            $details += "unexpected scenarios: $($unexpectedScenarios -join ', ')"
        }
        throw "$name benchmark evidence is incomplete ($($details -join '; '))."
    }

    $result.Metadata['benchmark.workload.id'] = $definition.ComparisonId
    $result.Metadata['benchmark.workload.sourceCommit'] = '5e1113a1195bed985c10788a6b89caf551663bb1'
    $result.Metadata['benchmark.workload.framework'] = $Framework
    $result.Metadata['gitSha'] = $gitSha
    $result.Metadata['gitDirty'] = $gitDirty.ToString().ToLowerInvariant()
    $result.Metadata['gitWorktreeClean'] = (-not $gitDirty).ToString().ToLowerInvariant()
    foreach ($sample in $result.Samples) {
        $sample.RunMode = $RunMode
    }
    foreach ($row in $result.Summary) {
        $row.RunMode = $RunMode
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
