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

function Get-ExpectedProductProofValidators {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Profile
    )

    switch ($Profile) {
        'PdfA3B' { return @('VeraPdf') }
        'PdfUa1' { return @('PdfUaValidator') }
        'FacturX' { return @('VeraPdf', 'Mustang') }
        'Zugferd' { return @('VeraPdf', 'Mustang') }
        default { throw "No expected external validator contract is defined for product proof profile $Profile." }
    }
}

function Assert-ProductProofContractProfile {
    param(
        [Parameter(Mandatory = $true)]
        $Entry,

        [Parameter(Mandatory = $true)]
        [string] $ExpectedProfile,

        [Parameter(Mandatory = $true)]
        [string] $SourceName
    )

    $validProductProofStatuses = @('Missing', 'NotRun', 'Passed', 'Failed', 'Error')
    $expectedValidators = @(Get-ExpectedProductProofValidators -Profile $ExpectedProfile)
    $requiredExternalValidators = @($Entry.requiredExternalValidators)
    $externalValidatorProofs = @($Entry.externalValidatorProofs)

    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($Entry.displayName)) -Message "Product proof contract row $ExpectedProfile in $SourceName is missing displayName."
    Assert-Condition -Condition ($Entry.PSObject.Properties.Name -contains 'hasRequiredExternalValidation') -Message "Product proof contract row $ExpectedProfile in $SourceName is missing hasRequiredExternalValidation."
    Assert-Condition -Condition ($Entry.PSObject.Properties.Name -contains 'canClaimConformance') -Message "Product proof contract row $ExpectedProfile in $SourceName is missing canClaimConformance."
    Assert-Condition -Condition ($Entry.PSObject.Properties.Name -contains 'failedExternalValidationCount') -Message "Product proof contract row $ExpectedProfile in $SourceName is missing failedExternalValidationCount."
    Assert-Condition -Condition ($requiredExternalValidators.Count -eq $expectedValidators.Count) -Message "Product proof contract row $ExpectedProfile in $SourceName has $($requiredExternalValidators.Count) required validators, expected $($expectedValidators.Count)."
    Assert-Condition -Condition ($externalValidatorProofs.Count -eq $expectedValidators.Count) -Message "Product proof contract row $ExpectedProfile in $SourceName has $($externalValidatorProofs.Count) external validator proof rows, expected $($expectedValidators.Count)."

    foreach ($expectedValidator in $expectedValidators) {
        Assert-Condition -Condition ($requiredExternalValidators -contains $expectedValidator) -Message "Product proof contract row $ExpectedProfile in $SourceName is missing required validator $expectedValidator."

        $validatorProof = @($externalValidatorProofs | Where-Object { $_.validatorKind -eq $expectedValidator })
        Assert-Condition -Condition ($validatorProof.Count -eq 1) -Message "Product proof contract row $ExpectedProfile in $SourceName must contain one $expectedValidator validator proof row."

        $row = $validatorProof[0]
        $status = [string] $row.status
        Assert-Condition -Condition ($validProductProofStatuses -contains $status) -Message "Product proof contract row $ExpectedProfile in $SourceName has invalid $expectedValidator status $status."
        Assert-Condition -Condition ($row.PSObject.Properties.Name -contains 'isSatisfied') -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator proof is missing isSatisfied."
        Assert-Condition -Condition ($row.PSObject.Properties.Name -contains 'blocksConformanceClaim') -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator proof is missing blocksConformanceClaim."
        Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($row.validatorName)) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator proof is missing validatorName."
        Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($row.diagnostic)) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator proof is missing diagnostic."
        Assert-Condition -Condition ($row.PSObject.Properties.Name -contains 'profile') -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator proof is missing profile."
        Assert-Condition -Condition ($row.PSObject.Properties.Name -contains 'exitCode') -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator proof is missing exitCode."

        if ($status -eq 'Passed') {
            Assert-Condition -Condition ($row.isSatisfied -eq $true) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator passed proof must be satisfied."
            Assert-Condition -Condition ($row.blocksConformanceClaim -eq $false) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator passed proof must not block a claim."
        } else {
            Assert-Condition -Condition ($row.isSatisfied -eq $false) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator non-passing proof must not be satisfied."
            Assert-Condition -Condition ($row.blocksConformanceClaim -eq $true) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator non-passing proof must block a claim."
        }
    }

    Assert-Condition -Condition ($Entry.hasRequiredExternalValidation -eq $false) -Message "Product proof contract row $ExpectedProfile in $SourceName must not have required external validation in groundwork proof."
    Assert-Condition -Condition ($Entry.canClaimConformance -eq $false) -Message "Product proof contract row $ExpectedProfile in $SourceName must not claim conformance."
    Assert-Condition -Condition (@($Entry.missingExternalValidators).Count -ge 1) -Message "Product proof contract row $ExpectedProfile in $SourceName must list missing or blocking external validators."
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
Assert-Condition -Condition ($proof.productProofContract.schemaVersion -eq 3) -Message 'Unexpected productProofContract schema version in proof.json.'
Assert-Condition -Condition ($productProofContractFile.schemaVersion -eq 3) -Message 'Unexpected officeimo-profile-proof-contract.json schema version.'
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
Assert-Condition -Condition ($productProofProfiles.Count -eq 4) -Message "Expected 4 product proof contract rows in proof.json, got $($productProofProfiles.Count)."
Assert-Condition -Condition ($productProofFileProfiles.Count -eq 4) -Message "Expected 4 product proof contract rows in officeimo-profile-proof-contract.json, got $($productProofFileProfiles.Count)."

$expectedProductProfiles = @('PdfA3B', 'PdfUa1', 'FacturX', 'Zugferd')
foreach ($expectedProductProfile in $expectedProductProfiles) {
    $entry = @($productProofProfiles | Where-Object { $_.profile -eq $expectedProductProfile })
    $fileEntry = @($productProofFileProfiles | Where-Object { $_.profile -eq $expectedProductProfile })
    Assert-Condition -Condition ($entry.Count -eq 1) -Message "Missing product proof contract row $expectedProductProfile in proof.json."
    Assert-Condition -Condition ($fileEntry.Count -eq 1) -Message "Missing product proof contract row $expectedProductProfile in officeimo-profile-proof-contract.json."
    Assert-ProductProofContractProfile -Entry $entry[0] -ExpectedProfile $expectedProductProfile -SourceName 'proof.json'
    Assert-ProductProofContractProfile -Entry $fileEntry[0] -ExpectedProfile $expectedProductProfile -SourceName 'officeimo-profile-proof-contract.json'
}

Write-Host "PDF compliance proof is valid: $resolvedProofPath"
