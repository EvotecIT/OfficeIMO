param(
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.IWork.Benchmarks/bin/Release/net8.0'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../Ignore/Benchmarks/IWorkRuntime'),
    [string] $ModulePath = 'PSPublishModule',
    [ValidateSet('Small', 'Medium', 'Large')] [string[]] $Scale = @('Small', 'Medium', 'Large'),
    [ValidateSet('Pages', 'Numbers', 'Keynote')] [string[]] $Kind = @('Pages', 'Numbers', 'Keynote'),
    [ValidateSet('LoadProject', 'ConvertSave')] [string[]] $Operation = @('LoadProject', 'ConvertSave'),
    [ValidateRange(0, 100)] [int] $WarmupCount = 2,
    [ValidateRange(1, 100)] [int] $IterationCount = 5,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
# Bind the candidate's core before the runner can load another OfficeIMO version.
[void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot 'OfficeIMO.Core.dll'))
Import-Module $ModulePath -ErrorAction Stop
$caseNames = @(foreach ($size in $Scale) { foreach ($family in $Kind) { "$family-$size" } })
$factors = @(foreach ($size in $Scale) { switch ($size) { Small { 1 } Medium { 10 } Large { 100 } } })
$result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'iwork-runtime.benchmark.ps1') `
    -OutputRoot $OutputRoot -Variable @{ BinaryRoot = $BinaryRoot; Factors = $factors } `
    -Case $caseNames -Operation $Operation -WarmupCount $WarmupCount -IterationCount $IterationCount -Plan:$Plan
$result
if (-not $Plan -and @($result.Samples | Where-Object Status -ne 'Succeeded').Count -gt 0) {
    throw 'iWork runtime qualification failed. Inspect the retained benchmark artifacts.'
}
if (-not $Plan -and @($result.Samples | Where-Object { $null -eq $_.AllocatedBytes }).Count -gt 0) {
    throw 'The runner did not measure managed allocation. Use PowerForge with operation allocation measurement, including its source build through ModulePath.'
}
