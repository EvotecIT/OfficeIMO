[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $BundlePath,
    [Parameter(Mandatory)][string] $OutputPath,
    [string] $ToolPath = (Join-Path $PSScriptRoot 'bin/Release/net8.0/OfficeIMO.ConversionConsistency.Tool.dll')
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$BundlePath = (Resolve-Path -LiteralPath $BundlePath).Path
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $OutputPath) { throw 'Negative-proof output must not already exist.' }
New-Item -ItemType Directory -Path $OutputPath | Out-Null
$bundle = Get-Content -LiteralPath (Join-Path $BundlePath 'bundle.json') -Raw | ConvertFrom-Json -Depth 100
$sourceCase = $bundle.cases | Where-Object { $_.contract.id -eq 'native-pptx' } | Select-Object -First 1
if ($null -eq $sourceCase) { throw 'Expected the native-pptx fixture in the source bundle.' }

foreach ($failure in @('missing-label', 'missing-image', 'changed-image', 'duplicate-page')) {
    $target = Join-Path $OutputPath $failure
    New-Item -ItemType Directory -Path $target | Out-Null
    Copy-Item -LiteralPath (Join-Path $BundlePath $sourceCase.contract.id) -Destination $target -Recurse
    $candidate = $bundle | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    $candidate.cases = @($candidate.cases | Where-Object { $_.contract.id -eq $sourceCase.contract.id })
    $entry = $candidate.cases[0]
    switch ($failure) {
        'missing-label' { $entry.contract.pages[0].text += 'REQUIRED-BUT-ABSENT' }
        'missing-image' { $entry.images[0].path = $entry.contract.id + '/missing.png' }
        'changed-image' { [IO.File]::AppendAllText((Join-Path $target $entry.images[0].path), 'tampered') }
        'duplicate-page' { $entry.images += $entry.images[0] }
    }
    $candidate | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath (Join-Path $target 'bundle.json') -Encoding utf8
    & dotnet $ToolPath verify --output $target *> (Join-Path $target 'verification.log')
    if ($LASTEXITCODE -ne 1) { throw "$failure did not return a structured verification failure." }
    $report = Get-Content -LiteralPath (Join-Path $target 'consistency-result.json') -Raw | ConvertFrom-Json -Depth 100
    if ($report.passed -or $report.cases[0].passed) { throw "$failure was incorrectly accepted." }
}

& dotnet $ToolPath verify --output $BundlePath --unknown-option value *> (Join-Path $OutputPath 'unknown-option.log')
if ($LASTEXITCODE -ne 2) { throw 'An unknown CLI option was not rejected.' }
& dotnet $ToolPath verify --output $BundlePath --output $BundlePath *> (Join-Path $OutputPath 'duplicate-option.log')
if ($LASTEXITCODE -ne 2) { throw 'A duplicate CLI option was not rejected.' }
Write-Host 'PASS: missing content, missing/corrupt images, duplicate pages, and invalid CLI options are rejected.'
# Expected native failures above have been validated; do not leak their exit code
# into callers such as the GitHub Actions PowerShell wrapper.
$global:LASTEXITCODE = 0
