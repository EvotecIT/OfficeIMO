param(
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.AI/bin/Release/net10.0'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../.validation/ai-packing'),
    [int[]] $Blocks = @(100, 1000, 4000),
    [int[]] $RequestCharacters = @(48000, 2000000),
    [int] $WarmupCount = 2,
    [int] $IterationCount = 7,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
Import-Module PSPublishModule -ErrorAction Stop
$result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'ai-packing.benchmark.ps1') `
    -OutputRoot $OutputRoot -Variable @{ BinaryRoot = $BinaryRoot; Blocks = $Blocks; RequestCharacters = $RequestCharacters } `
    -WarmupCount $WarmupCount -IterationCount $IterationCount -Plan:$Plan
$result
if (-not $Plan -and @($result.Summary | Where-Object Status -ne 'Succeeded').Count -gt 0) {
    throw 'AI packing measurement failed validation. Inspect the benchmark artifacts.'
}
