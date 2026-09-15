param(
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.Drawing.Benchmarks/bin/Release/net10.0'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../.validation/release-quality-images'),
    [int] $WarmupCount = 1,
    [int] $IterationCount = 3,
    [switch] $Plan
)

$ErrorActionPreference = 'Stop'
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
[void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot 'OfficeIMO.Core.dll'))
Import-Module PSPublishModule -MinimumVersion 3.0.141 -ErrorAction Stop

$result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'release-quality-images.benchmark.ps1') `
    -OutputRoot $OutputRoot -Variable @{ BinaryRoot = $BinaryRoot } `
    -WarmupCount $WarmupCount -IterationCount $IterationCount -Plan:$Plan
$result

if (-not $Plan -and @($result.Summary | Where-Object Status -ne 'Succeeded').Count -gt 0) {
    throw 'Release-quality image evidence failed validation. Inspect the benchmark artifacts.'
}
