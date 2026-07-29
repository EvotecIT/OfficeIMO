param(
    [ValidateSet('quick', 'full')]
    [string] $RunMode = 'quick',
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0',
    [string] $OutputRoot = (Join-Path ([System.IO.Path]::GetTempPath()) 'OfficeIMO\Benchmarks\Runs'),
    [switch] $Publish
)

$ErrorActionPreference = 'Stop'

if ($Publish -and $RunMode -ne 'full') {
    throw 'Only a full BenchmarkDotNet run can be marked publishable.'
}

. (Join-Path $PSScriptRoot 'TabularBenchmarkEvidence.ps1')
Import-Module PSPublishModule -MinimumVersion 3.0.81 -Force

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path (Join-Path $repositoryRoot 'OfficeIMO.Tabular.Benchmarks') 'OfficeIMO.Tabular.Benchmarks.csproj'
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
$stamp = [DateTimeOffset]::UtcNow.ToString('yyyyMMdd-HHmmss')
$artifactsPath = Join-Path $OutputRoot "$platform-$RunMode-$stamp"
$staticRoot = Join-Path (Join-Path (Join-Path (Join-Path $repositoryRoot 'Website') 'static') 'data') 'benchmarks'
$staticRoot = Join-Path $staticRoot 'tabular'
$catalogPath = Join-Path $staticRoot 'index.json'
$catalogEligible = $RunMode -eq 'quick' -or [bool] $Publish

New-Item -ItemType Directory -Force -Path $artifactsPath, $staticRoot | Out-Null

& dotnet run -c Release -f $Framework --project $projectPath -- --validate
if ($LASTEXITCODE -ne 0) {
    throw "Tabular benchmark validation failed with exit code $LASTEXITCODE."
}

$runnerArguments = @(
    'run', '-c', 'Release', '-f', $Framework,
    '--project', $projectPath, '--',
    '--artifacts', $artifactsPath
)
if ($RunMode -eq 'quick') {
    $runnerArguments += '--quick'
}

& dotnet @runnerArguments
if ($LASTEXITCODE -ne 0) {
    throw "Tabular benchmark run failed with exit code $LASTEXITCODE."
}

$provenancePath = Join-Path $artifactsPath 'tabular-benchmark-provenance.json'
$provenance = Get-Content -LiteralPath $provenancePath -Raw -Encoding UTF8 | ConvertFrom-Json
$comparisonId = "$($provenance.workload)-$Framework"
$evidenceLocation = Get-TabularBenchmarkEvidenceLocation `
    -ComparisonId $comparisonId `
    -Platform $platform `
    -RunMode $RunMode `
    -StaticRoot $staticRoot
$normalizedPath = if ($catalogEligible) {
    $evidenceLocation.Path
} else {
    Join-Path $artifactsPath 'normalized-result.json'
}
$result = Import-BenchmarkResult -Path $artifactsPath -Suite 'OfficeIMO.Tabular'

$result.Metadata['benchmark.workload.id'] = [string] $provenance.workload
$result.Metadata['benchmark.workload.framework'] = $Framework
$result.Metadata['benchmark.workload.sourceCommit'] = [string] $provenance.sourceCommit
foreach ($fixture in $provenance.fixtures) {
    $key = [System.IO.Path]::GetFileNameWithoutExtension([string] $fixture.name).ToLowerInvariant()
    $result.Metadata["benchmark.fixture.$key"] = [string] $fixture.sha256
}
foreach ($package in $provenance.packages.PSObject.Properties) {
    $result.Metadata["benchmark.package.$($package.Name)"] = [string] $package.Value
}

$gitSha = (& git -C $repositoryRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($gitSha)) {
    throw 'Unable to resolve the OfficeIMO source commit for benchmark provenance.'
}
$result.Metadata['gitSha'] = $gitSha
foreach ($sample in $result.Samples) {
    $sample.RunMode = $RunMode
}
foreach ($row in $result.Summary) {
    $row.RunMode = $RunMode
}

Write-TabularBenchmarkResult -Path $normalizedPath -InputObject $result
if ($catalogEligible) {
    Update-BenchmarkEvidenceCatalog `
        -InputObject $result `
        -Path $catalogPath `
        -ComparisonId $comparisonId `
        -ResultPath $evidenceLocation.ResultPath `
        -RunMode $RunMode `
        -ExpectedPlatform windows, linux, macos `
        -Publish:$Publish | Out-Null
}

[pscustomobject]@{
    Platform = $platform
    RunMode = $RunMode
    Publish = [bool] $Publish
    SourceCommit = $gitSha
    ArtifactsPath = $artifactsPath
    NormalizedResult = $normalizedPath
    EvidenceCatalog = if ($catalogEligible) { $catalogPath } else { $null }
}
