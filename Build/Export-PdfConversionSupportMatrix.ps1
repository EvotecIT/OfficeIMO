[CmdletBinding()]
param(
    [string] $ManifestPath,
    [string] $OutputPath,
    [switch] $Check
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $ManifestPath = Join-Path $repoRoot 'Docs/pdf-conversion-scenarios.json'
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $repoRoot 'Docs/officeimo.pdf-conversion-support-matrix.md'
}

$resolvedManifestPath = [System.IO.Path]::GetFullPath($ManifestPath)
$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $resolvedManifestPath -PathType Leaf)) {
    throw "PDF conversion scenario manifest was not found: $resolvedManifestPath"
}

$scenarioManifest = Get-Content -LiteralPath $resolvedManifestPath -Raw | ConvertFrom-Json
$qualityContract = $scenarioManifest.qualityContract
$supportLines = [System.Collections.Generic.List[string]]::new()
$supportLines.Add('# OfficeIMO PDF Conversion Support Matrix')
$supportLines.Add('')
$supportLines.Add('This matrix is generated from `Docs/pdf-conversion-scenarios.json`. Fidelity status describes the current evidence, not the intended destination.')
$supportLines.Add('')
$supportLines.Add("Premium claim rule: $($qualityContract.premiumClaimRule)")
$supportLines.Add('')
$supportLines.Add('| Source | Formats | Mode | Evidence status | Reference policy |')
$supportLines.Add('| --- | --- | --- | --- | --- |')
foreach ($converter in @($scenarioManifest.converterCatalog)) {
    $source = ([string]$converter.id).Replace('|', '\|')
    $formats = (@($converter.sourceFormats) -join ', ').Replace('|', '\|')
    $mode = ([string]$converter.conversionMode).Replace('|', '\|')
    $fidelityStatus = ([string]$converter.fidelityStatus).Replace('|', '\|')
    $referencePolicy = ([string]$converter.referencePolicy).Replace('|', '\|')
    $supportLines.Add("| $source | $formats | $mode | $fidelityStatus | $referencePolicy |")
}
$supportLines.Add('')
$supportLines.Add('## Capability Claims')
$supportLines.Add('')
$supportLines.Add('| Source | Capability | Fidelity level | Evidence scenarios |')
$supportLines.Add('| --- | --- | --- | --- |')
foreach ($converter in @($scenarioManifest.converterCatalog)) {
    if ($null -eq $converter.capabilityClaims) {
        continue
    }
    foreach ($claim in @($converter.capabilityClaims)) {
        $source = ([string]$converter.id).Replace('|', '\|')
        $capability = ([string]$claim.capability).Replace('|', '\|')
        $level = ([string]$claim.level).Replace('|', '\|')
        $evidence = (@($claim.evidenceScenarioIds) -join ', ').Replace('|', '\|')
        $supportLines.Add("| $source | $capability | $level | $evidence |")
    }
}
$supportLines.Add('')
$supportLines.Add('## Direct, Composed, And Planned Routes')
$supportLines.Add('')
$supportLines.Add('| Route | Formats | Status | Implementation owner | Contract evidence | Diagnostic contract |')
$supportLines.Add('| --- | --- | --- | --- | --- | --- |')
foreach ($route in @($scenarioManifest.compositionRoutes)) {
    $routeId = ([string]$route.id).Replace('|', '\|')
    $formats = (@($route.sourceFormats) -join ', ').Replace('|', '\|')
    $routeStatus = ([string]$route.status).Replace('|', '\|')
    $implementationProject = if ([string]::IsNullOrWhiteSpace([string]$route.implementationProject)) {
        '—'
    } else {
        ('`' + ([string]$route.implementationProject).Replace('|', '\|') + '`')
    }
    $evidenceTest = if ([string]::IsNullOrWhiteSpace([string]$route.evidenceTest)) {
        '—'
    } else {
        ('`' + ([string]$route.evidenceTest).Replace('|', '\|') + '`')
    }
    $diagnosticContract = ([string]$route.diagnosticContract).Replace('|', '\|')
    $supportLines.Add("| $routeId | $formats | $routeStatus | $implementationProject | $evidenceTest | $diagnosticContract |")
}

$expected = ($supportLines -join "`n") + "`n"
if ($Check) {
    if (-not (Test-Path -LiteralPath $resolvedOutputPath -PathType Leaf)) {
        throw "Generated PDF conversion support matrix is missing: $resolvedOutputPath"
    }

    $actual = [System.IO.File]::ReadAllText($resolvedOutputPath).Replace("`r`n", "`n")
    if ($actual -ne $expected) {
        throw "Generated PDF conversion support matrix is out of date. Run Build/Export-PdfConversionSupportMatrix.ps1."
    }

    Write-Host "PDF conversion support matrix is current: $resolvedOutputPath"
    return
}

$outputDirectory = Split-Path -Parent $resolvedOutputPath
if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
    [System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
}
[System.IO.File]::WriteAllText(
    $resolvedOutputPath,
    $expected,
    [System.Text.UTF8Encoding]::new($false))
Write-Host "Generated PDF conversion support matrix: $resolvedOutputPath"
