param(
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.Studio/bin/Release/net10.0'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../.validation/assistant-benchmarks'),
    [int[]] $Pages = @(1, 100, 500),
    [int] $WarmupCount = 2,
    [int] $IterationCount = 7,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
# Run in a fresh PowerShell process: benchmark the built Core before importing tools
# that may carry a different OfficeIMO version for their own document generation.
[void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot 'OfficeIMO.Core.dll'))
Import-Module PSPublishModule -ErrorAction Stop
$result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'assistant-evidence.benchmark.ps1') `
    -OutputRoot $OutputRoot -Variable @{ BinaryRoot = $BinaryRoot; Pages = $Pages } `
    -WarmupCount $WarmupCount -IterationCount $IterationCount -Plan:$Plan
$result
if (-not $Plan -and @($result.Summary | Where-Object Status -ne 'Succeeded').Count -gt 0) {
    throw 'Assistant evidence measurement failed validation. Inspect the benchmark artifacts.'
}
