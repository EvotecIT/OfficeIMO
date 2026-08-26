param(
    [string] $PdfComplianceProofPath = $env:OFFICEIMO_PDF_COMPLIANCE_PROOF_PATH
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$browserProjectPath = Join-Path $PSScriptRoot '../OfficeIMO.Html.Pdf.Browser/OfficeIMO.Html.Pdf.Browser.csproj'
[xml] $browserProject = Get-Content -LiteralPath $browserProjectPath -Raw
$htmlTinkerXVersionText = [string] $browserProject.Project.PropertyGroup.HtmlTinkerXVersion
if ([string]::IsNullOrWhiteSpace($htmlTinkerXVersionText) -or
    [version] $htmlTinkerXVersionText -lt [version] '3.0.1') {
    throw 'OfficeIMO release is blocked: publish HtmlTinkerX 3.0.1 or newer with device emulation and fail-closed offline policy, then repin OfficeIMO.Html.Pdf.Browser.'
}

if ([string]::IsNullOrWhiteSpace($PdfComplianceProofPath)) {
    throw 'OfficeIMO release is blocked: provide -PdfComplianceProofPath or OFFICEIMO_PDF_COMPLIANCE_PROOF_PATH from a qualified exact-artifact validation run.'
}

$resolvedProofPath = (Resolve-Path -LiteralPath $PdfComplianceProofPath -ErrorAction Stop).Path
& "$PSScriptRoot/../.github/scripts/Assert-PdfComplianceProof.ps1" -ProofPath $resolvedProofPath

$proof = Get-Content -LiteralPath (Join-Path $resolvedProofPath 'proof.json') -Raw |
    ConvertFrom-Json
foreach ($profileName in @('PdfX1A2003', 'PdfX4')) {
    $profiles = @($proof.productProofContract.profiles | Where-Object profile -eq $profileName)
    if ($profiles.Count -ne 1) {
        throw "OfficeIMO release is blocked: proof must contain exactly one $profileName product profile."
    }

    $profile = $profiles[0]
    $validators = @($profile.externalValidatorProofs |
            Where-Object validatorKind -eq 'PdfXValidator')
    if ($profile.canClaimConformance -ne $true -or
        [string] $profile.proofStatus -ne 'Claimable' -or
        $validators.Count -ne 1 -or
        [string] $validators[0].status -ne 'Passed' -or
        $validators[0].isSatisfied -ne $true -or
        [string]::IsNullOrWhiteSpace([string] $validators[0].validatorVersion) -or
        [string] $validators[0].validatorVersion -eq 'unknown') {
        throw "OfficeIMO release is blocked: $profileName requires claimable exact-artifact PDF/X proof from a versioned qualified validator."
    }
}

$pdfXDiagnostics = @($proof.validatorDiagnostics |
        Where-Object validatorKind -eq 'PdfXValidator')
if ($pdfXDiagnostics.Count -ne 2 -or
    @($pdfXDiagnostics | Where-Object status -ne 'Passed').Count -ne 0) {
    throw 'OfficeIMO release is blocked: both exact PDF/X validator diagnostics must pass.'
}

foreach ($diagnostic in $pdfXDiagnostics) {
    $diagnosticPath = Join-Path $resolvedProofPath ([string] $diagnostic.file)
    $diagnosticText = Get-Content -LiteralPath $diagnosticPath -Raw
    if ($diagnosticText -notmatch '(?i)callas\s+pdfToolbox') {
        throw "OfficeIMO release is blocked: $($diagnostic.file) is not traceable to the qualified callas pdfToolbox adapter."
    }
}

$artifactEvidencePath = Join-Path $PSScriptRoot '../Docs/benchmarks/html-pdf-artifact-evidence/html-pdf-artifact-evidence-net10.0-windows-high.json'
$artifactEvidence = Get-Content -LiteralPath $artifactEvidencePath -Raw | ConvertFrom-Json
$expectedHtmlTinkerXCommit = [string] $artifactEvidence.report.provenance.htmlTinkerX.commit
if ($expectedHtmlTinkerXCommit -notmatch '^[0-9a-f]{40}$') {
    throw 'OfficeIMO release is blocked: committed artifact evidence lacks an exact HtmlTinkerX source commit.'
}

& "$PSScriptRoot/Test-LibraryComparisonRunnerContract.ps1"
& "$PSScriptRoot/Test-HtmlPdfBenchmarkEvidence.ps1"
& "$PSScriptRoot/Test-HtmlPdfArtifactEvidenceExporter.ps1"
& "$PSScriptRoot/Test-HtmlPdfArtifactEvidence.ps1"
& "$PSScriptRoot/Test-HtmlPdfBrowserPackages.ps1" -ExpectedHtmlTinkerXCommit $expectedHtmlTinkerXCommit
& "$PSScriptRoot/Test-TypographyPackages.ps1"

Write-Host 'OfficeIMO HTML/PDF release gate passed.'
