param(
    [Parameter(Mandatory = $true)]
    [string] $ProofPath
)

$ErrorActionPreference = 'Stop'

function Assert-Condition {
    param(
        [Parameter(Mandatory = $true)]
        [bool] $Condition,

        [Parameter(Mandatory = $true)]
        [string] $Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

$resolvedProofPath = Resolve-Path -LiteralPath $ProofPath
$jsonPath = Join-Path $resolvedProofPath 'proof.json'
$indexPath = Join-Path $resolvedProofPath 'index.md'
$productProofContractPath = Join-Path $resolvedProofPath 'officeimo-profile-proof-contract.json'

Assert-Condition -Condition (Test-Path -LiteralPath $jsonPath) -Message "Missing proof.json in $resolvedProofPath."
Assert-Condition -Condition (Test-Path -LiteralPath $indexPath) -Message "Missing index.md in $resolvedProofPath."
Assert-Condition -Condition (Test-Path -LiteralPath $productProofContractPath) -Message "Missing officeimo-profile-proof-contract.json in $resolvedProofPath."

$proof = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json
$productProofContractFile = Get-Content -LiteralPath $productProofContractPath -Raw | ConvertFrom-Json

Assert-Condition -Condition ($proof.schemaVersion -eq 3) -Message 'Unexpected proof schema version.'
Assert-Condition -Condition ($proof.testExitCode -eq 0) -Message "Expected testExitCode 0, got $($proof.testExitCode)."
Assert-Condition -Condition ($null -ne $proof.strictValidatorMode) -Message 'Missing strictValidatorMode in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration) -Message 'Missing validatorConfiguration in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.veraPdfExecutableConfigured) -Message 'Missing veraPdfExecutableConfigured in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.veraPdfArgsConfigured) -Message 'Missing veraPdfArgsConfigured in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.pdfUaValidatorExecutableConfigured) -Message 'Missing pdfUaValidatorExecutableConfigured in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.pdfUaValidatorArgsConfigured) -Message 'Missing pdfUaValidatorArgsConfigured in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.mustangExecutableConfigured) -Message 'Missing mustangExecutableConfigured in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.mustangArgsConfigured) -Message 'Missing mustangArgsConfigured in proof.json.'
Assert-Condition -Condition ($proof.contract.formalProfilesFailClosed -eq $true) -Message 'Proof contract must keep formal profiles fail-closed.'
Assert-Condition -Condition ($proof.contract.generatedPdfsAreGroundworkFixtures -eq $true) -Message 'Proof contract must mark generated PDFs as groundwork fixtures.'
Assert-Condition -Condition ($proof.contract.externalValidationRequiredForClaims -eq $true) -Message 'Proof contract must require external validation for claims.'
Assert-Condition -Condition ($null -ne $proof.productProofContract) -Message 'Missing productProofContract in proof.json.'
Assert-Condition -Condition ($proof.productProofContract.schemaVersion -eq 2) -Message 'Unexpected productProofContract schema version in proof.json.'
Assert-Condition -Condition ($productProofContractFile.schemaVersion -eq 2) -Message 'Unexpected officeimo-profile-proof-contract.json schema version.'
Assert-Condition -Condition ([string] $proof.productProofContract.generatedBy -eq [string] $productProofContractFile.generatedBy) -Message 'Product proof contract generatedBy mismatch.'
Assert-Condition -Condition ([string] $proof.productProofContract.externalEvidenceMode -eq 'NoExternalValidationInjected') -Message 'Unexpected product proof contract externalEvidenceMode in proof.json.'
Assert-Condition -Condition ([string] $productProofContractFile.externalEvidenceMode -eq 'NoExternalValidationInjected') -Message 'Unexpected externalEvidenceMode in officeimo-profile-proof-contract.json.'

$pdfFixtures = @($proof.pdfFixtures)
Assert-Condition -Condition ($pdfFixtures.Count -eq 3) -Message "Expected 3 PDF fixtures, got $($pdfFixtures.Count)."

$expectedPdfNames = @(
    'officeimo-pdfa3-groundwork.pdf',
    'officeimo-einvoice-groundwork.pdf',
    'officeimo-pdfua-groundwork.pdf'
)

foreach ($expectedName in $expectedPdfNames) {
    $entry = @($pdfFixtures | Where-Object { $_.file -eq $expectedName })
    Assert-Condition -Condition ($entry.Count -eq 1) -Message "Missing PDF fixture entry $expectedName."

    $filePath = Join-Path $resolvedProofPath $expectedName
    Assert-Condition -Condition (Test-Path -LiteralPath $filePath) -Message "Missing PDF fixture file $expectedName."
    $file = Get-Item -LiteralPath $filePath
    Assert-Condition -Condition ($file.Length -eq [long] $entry[0].sizeBytes) -Message "Size mismatch for $expectedName."

    $hash = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash.ToLowerInvariant()
    Assert-Condition -Condition ($hash -eq [string] $entry[0].sha256) -Message "SHA-256 mismatch for $expectedName."
}

$validatorDiagnostics = @($proof.validatorDiagnostics)
Assert-Condition -Condition ($validatorDiagnostics.Count -ge 4) -Message "Expected at least 4 validator diagnostics, got $($validatorDiagnostics.Count)."

$requiredValidatorKinds = @('VeraPdf', 'PdfUaValidator', 'Mustang')
foreach ($validatorKind in $requiredValidatorKinds) {
    $entry = @($validatorDiagnostics | Where-Object { $_.validatorKind -eq $validatorKind })
    Assert-Condition -Condition ($entry.Count -ge 1) -Message "Missing validator diagnostic entry for $validatorKind."
}

$validStatuses = @('NotRun', 'Passed', 'Failed', 'Error')

foreach ($entry in $validatorDiagnostics) {
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry.file)) -Message 'Validator diagnostic entry is missing file name.'
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry.validatorKind)) -Message "Validator diagnostic entry $($entry.file) is missing validatorKind."
    Assert-Condition -Condition ($validStatuses -contains [string] $entry.status) -Message "Validator diagnostic entry $($entry.file) has invalid status $($entry.status)."
    Assert-Condition -Condition ($validStatuses -contains [string] $entry.expectedStatus) -Message "Validator diagnostic entry $($entry.file) has invalid expectedStatus $($entry.expectedStatus)."
    Assert-Condition -Condition ($entry.matchesExpectedStatus -eq $true) -Message "Validator diagnostic entry $($entry.file) status $($entry.status) did not match expected status $($entry.expectedStatus)."
    $filePath = Join-Path $resolvedProofPath $entry.file
    Assert-Condition -Condition (Test-Path -LiteralPath $filePath) -Message "Missing validator diagnostic file $($entry.file)."
    $file = Get-Item -LiteralPath $filePath
    Assert-Condition -Condition ($file.Length -eq [long] $entry.sizeBytes) -Message "Size mismatch for $($entry.file)."

    $hash = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash.ToLowerInvariant()
    Assert-Condition -Condition ($hash -eq [string] $entry.sha256) -Message "SHA-256 mismatch for $($entry.file)."
}

$profileProofs = @($proof.profileProofs)
Assert-Condition -Condition ($profileProofs.Count -eq 4) -Message "Expected 4 profile proof rows, got $($profileProofs.Count)."

$expectedProfiles = @(
    'pdfa-3b-groundwork',
    'pdfua-1-groundwork',
    'einvoice-pdfa3-groundwork',
    'einvoice-groundwork'
)

foreach ($expectedProfile in $expectedProfiles) {
    $entry = @($profileProofs | Where-Object { $_.profileId -eq $expectedProfile })
    Assert-Condition -Condition ($entry.Count -eq 1) -Message "Missing profile proof row $expectedProfile."

    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry[0].displayName)) -Message "Profile proof row $expectedProfile is missing displayName."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry[0].fixtureFile)) -Message "Profile proof row $expectedProfile is missing fixtureFile."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry[0].fixtureSha256)) -Message "Profile proof row $expectedProfile is missing fixtureSha256."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry[0].validatorKind)) -Message "Profile proof row $expectedProfile is missing validatorKind."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry[0].validatorDiagnosticFile)) -Message "Profile proof row $expectedProfile is missing validatorDiagnosticFile."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry[0].readinessRequirementId)) -Message "Profile proof row $expectedProfile is missing readinessRequirementId."
    Assert-Condition -Condition ($validStatuses -contains [string] $entry[0].status) -Message "Profile proof row $expectedProfile has invalid status $($entry[0].status)."
    Assert-Condition -Condition ($validStatuses -contains [string] $entry[0].expectedStatus) -Message "Profile proof row $expectedProfile has invalid expectedStatus $($entry[0].expectedStatus)."
    Assert-Condition -Condition ($entry[0].matchesExpectedStatus -eq $true) -Message "Profile proof row $expectedProfile status $($entry[0].status) did not match expected status $($entry[0].expectedStatus)."
    Assert-Condition -Condition ($entry[0].canClaimConformance -eq $false) -Message "Profile proof row $expectedProfile must not claim conformance for groundwork fixtures."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry[0].formalClaimStatus)) -Message "Profile proof row $expectedProfile is missing formalClaimStatus."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry[0].nextAction)) -Message "Profile proof row $expectedProfile is missing nextAction."
}

$productProofProfiles = @($proof.productProofContract.profiles)
$productProofFileProfiles = @($productProofContractFile.profiles)
Assert-Condition -Condition ($productProofProfiles.Count -eq 3) -Message "Expected 3 product proof contract rows in proof.json, got $($productProofProfiles.Count)."
Assert-Condition -Condition ($productProofFileProfiles.Count -eq 3) -Message "Expected 3 product proof contract rows in officeimo-profile-proof-contract.json, got $($productProofFileProfiles.Count)."

$expectedProductProfiles = @('PdfA3B', 'PdfUa1', 'FacturX')
foreach ($expectedProductProfile in $expectedProductProfiles) {
    $entry = @($productProofProfiles | Where-Object { $_.profile -eq $expectedProductProfile })
    $fileEntry = @($productProofFileProfiles | Where-Object { $_.profile -eq $expectedProductProfile })
    Assert-Condition -Condition ($entry.Count -eq 1) -Message "Missing product proof contract row $expectedProductProfile in proof.json."
    Assert-Condition -Condition ($fileEntry.Count -eq 1) -Message "Missing product proof contract row $expectedProductProfile in officeimo-profile-proof-contract.json."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry[0].displayName)) -Message "Product proof contract row $expectedProductProfile is missing displayName."
    Assert-Condition -Condition ($entry[0].hasRequiredExternalValidation -eq $false) -Message "Product proof contract row $expectedProductProfile must not have required external validation in groundwork proof."
    Assert-Condition -Condition ($entry[0].canClaimConformance -eq $false) -Message "Product proof contract row $expectedProductProfile must not claim conformance."
    Assert-Condition -Condition (@($entry[0].missingExternalValidators).Count -ge 1) -Message "Product proof contract row $expectedProductProfile must list missing external validators."
}

Write-Host "PDF compliance proof is valid: $resolvedProofPath"
