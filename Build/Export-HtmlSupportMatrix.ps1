[CmdletBinding()]
param(
    [switch] $Check
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$testProject = Join-Path $repoRoot 'OfficeIMO.Html.Tests/OfficeIMO.Html.Tests.csproj'
$outputPath = Join-Path $repoRoot 'Docs/officeimo.html-support-matrix.md'

if ($Check) {
    dotnet test $testProject --framework net8.0 --filter 'FullyQualifiedName~HtmlSupportMatrix_CheckedInArtifactMatchesExecutableContracts' --no-restore
    if ($LASTEXITCODE -ne 0) {
        throw 'The generated HTML support matrix is out of date.'
    }
    return
}

$previousUpdateValue = $env:OFFICEIMO_UPDATE_HTML_SUPPORT_MATRIX
try {
    $env:OFFICEIMO_UPDATE_HTML_SUPPORT_MATRIX = '1'
    dotnet test $testProject --framework net8.0 --filter 'FullyQualifiedName~HtmlSupportMatrix_CheckedInArtifactMatchesExecutableContracts' --no-restore
    if ($LASTEXITCODE -ne 0) {
        throw 'Failed to generate the HTML support matrix.'
    }
} finally {
    if ($null -eq $previousUpdateValue) {
        Remove-Item Env:OFFICEIMO_UPDATE_HTML_SUPPORT_MATRIX -ErrorAction SilentlyContinue
    } else {
        $env:OFFICEIMO_UPDATE_HTML_SUPPORT_MATRIX = $previousUpdateValue
    }
}

Write-Host "Generated HTML support matrix: $outputPath"
