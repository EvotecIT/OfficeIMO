param(
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.OpenDocument.Benchmarks/bin/Release/net8.0'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../Ignore/Benchmarks/DrawPageProjection'),
    [string] $ModulePath = 'PSPublishModule',
    [ValidateRange(1, 1000)] [int[]] $PageCount = @(100, 500, 1000),
    [ValidateSet('Empty', 'Fields')] [string[]] $Content = @('Empty', 'Fields'),
    [ValidateRange(0, 100)] [int] $WarmupCount = 2,
    [ValidateRange(1, 100)] [int] $IterationCount = 5,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
[void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot 'OfficeIMO.Core.dll'))
[void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot 'OfficeIMO.OpenDocument.dll'))
Import-Module $ModulePath -ErrorAction Stop
$result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'draw-page-projection.benchmark.ps1') `
    -OutputRoot $OutputRoot -Variable @{ BinaryRoot = $BinaryRoot; PageCounts = $PageCount; IncludeEmpty = 'Empty' -in $Content; IncludeFields = 'Fields' -in $Content } `
    -WarmupCount $WarmupCount -IterationCount $IterationCount -Plan:$Plan
$result
if (-not $Plan -and @($result.Samples | Where-Object Status -ne 'Succeeded').Count -gt 0) {
    throw 'Draw projection verification failed. Inspect the retained benchmark artifacts.'
}
if (-not $Plan -and @($result.Samples | Where-Object { $null -eq $_.AllocatedBytes }).Count -gt 0) {
    throw 'The shared runner did not measure managed allocation.'
}
