[CmdletBinding()]
param(
    # Where the fetched corpora and the reports go. It must be outside the repository: the corpora are third-party
    # test files that are never committed.
    [Parameter(Mandatory)]
    [string] $OutputDirectory,

    # Source ids from Build/RealWorldCorpus/openxml-corpus-sources.json; all of them when omitted.
    [string[]] $Source,

    [ValidateRange(1, 16)]
    [int] $Parallelism = 4,

    [ValidateRange(1, 900)]
    [int] $TimeoutSeconds = 30,

    # Reuse corpora already fetched into the output directory instead of fetching and verifying them again.
    [switch] $SkipFetch,

    # Leave document text out of the difference report.
    [switch] $WithoutDifferenceText
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repositoryRoot 'Build/RealWorldCorpus/OfficeIMO.RealWorldCorpus.Tool.csproj'
$manifest = Join-Path $repositoryRoot 'Build/RealWorldCorpus/openxml-corpus-sources.json'
$output = [System.IO.Path]::GetFullPath($OutputDirectory)
$root = [System.IO.Path]::GetFullPath($repositoryRoot).TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
if (($output.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar).StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
    throw "The output directory must be outside the repository: $output"
}

$corpus = Join-Path $output 'corpus'
$reports = Join-Path $output 'reports'
& dotnet build $project --framework net10.0 --configuration Release --nologo --verbosity quiet
if ($LASTEXITCODE -ne 0) { throw "The corpus runner did not build (exit code $LASTEXITCODE)." }

if (-not $SkipFetch) {
    $fetch = @('fetch-sources', '--manifest', $manifest, '--output', $corpus)
    if ($Source) { $fetch += @('--sources', ($Source -join ',')) }
    & dotnet run --project $project --framework net10.0 --configuration Release --no-build -- @fetch
    if ($LASTEXITCODE -ne 0) { throw "The corpora could not be fetched and verified (exit code $LASTEXITCODE)." }
}

$run = @(
    'run', '--operation', 'openxml-differential',
    '--input', (Join-Path $corpus 'files'),
    '--source-provenance', (Join-Path $corpus 'sources.json'),
    '--json', (Join-Path $reports 'report.json'),
    '--markdown', (Join-Path $reports 'report.md'),
    '--corpus-id', 'openxml-test-corpora',
    '--max-per-format', '5000', '--max-total', '10000', '--max-traversal-entries', '20000',
    '--timeout-seconds', $TimeoutSeconds, '--parallelism', $Parallelism,
    '--include-source-names'
)
if (-not $WithoutDifferenceText) { $run += '--include-difference-text' }
& dotnet run --project $project --framework net10.0 --configuration Release --no-build -- @run
if ($LASTEXITCODE -ne 0) { throw "The differential run failed (exit code $LASTEXITCODE)." }
