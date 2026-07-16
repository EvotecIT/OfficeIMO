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
        'PdfA2B' { return @('VeraPdf') }
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
    Assert-Condition -Condition ($Entry.PSObject.Properties.Name -contains 'proofStatus') -Message "Product proof contract row $ExpectedProfile in $SourceName is missing proofStatus."
    Assert-Condition -Condition ($Entry.hasArtifactEvidence -eq $true) -Message "Product proof contract row $ExpectedProfile in $SourceName must identify an exact artifact."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($Entry.artifactSha256) -and ([string] $Entry.artifactSha256).Length -eq 64) -Message "Product proof contract row $ExpectedProfile in $SourceName is missing an artifact SHA-256."
    Assert-Condition -Condition ([long] $Entry.artifactSizeBytes -gt 0) -Message "Product proof contract row $ExpectedProfile in $SourceName is missing artifact byte length."
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
        Assert-Condition -Condition ($row.PSObject.Properties.Name -contains 'validatorVersion') -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator proof is missing validatorVersion."
        Assert-Condition -Condition ($row.PSObject.Properties.Name -contains 'artifactSha256') -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator proof is missing artifactSha256."
        Assert-Condition -Condition ($row.PSObject.Properties.Name -contains 'artifactSizeBytes') -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator proof is missing artifactSizeBytes."
        Assert-Condition -Condition ($row.PSObject.Properties.Name -contains 'validatedAtUtc') -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator proof is missing validatedAtUtc."
        Assert-Condition -Condition ($row.PSObject.Properties.Name -contains 'warnings') -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator proof is missing warnings."

        if ($status -eq 'Passed') {
            Assert-Condition -Condition ($row.isSatisfied -eq $true) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator passed proof must be satisfied."
            Assert-Condition -Condition ($row.blocksConformanceClaim -eq $false) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator passed proof must not block a claim."
            Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($row.validatorVersion)) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator passed proof is missing validator version."
            Assert-Condition -Condition ([string] $row.artifactSha256 -eq [string] $Entry.artifactSha256) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator passed proof has the wrong artifact SHA-256."
            Assert-Condition -Condition ([long] $row.artifactSizeBytes -eq [long] $Entry.artifactSizeBytes) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator passed proof has the wrong artifact byte length."
            Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($row.validatedAtUtc)) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator passed proof is missing validation timestamp."
        } else {
            Assert-Condition -Condition ($row.isSatisfied -eq $false) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator non-passing proof must not be satisfied."
            Assert-Condition -Condition ($row.blocksConformanceClaim -eq $true) -Message "Product proof contract row $ExpectedProfile in $SourceName $expectedValidator non-passing proof must block a claim."
        }
    }

    $isFormalProfile = $ExpectedProfile -in @('PdfA2B', 'PdfA3B', 'PdfUa1', 'FacturX', 'Zugferd')
    if ($isFormalProfile -and @($externalValidatorProofs | Where-Object { $_.status -eq 'Passed' }).Count -eq $expectedValidators.Count) {
        Assert-Condition -Condition ($Entry.isInternallyReady -eq $true) -Message "$ExpectedProfile product proof must be internally ready."
        Assert-Condition -Condition ($Entry.hasRequiredExternalValidation -eq $true) -Message "$ExpectedProfile product proof must include passing external validation."
        Assert-Condition -Condition ($Entry.canClaimConformance -eq $true) -Message "$ExpectedProfile exact artifact should be claimable after internal and external proof pass."
        Assert-Condition -Condition ([string] $Entry.proofStatus -eq 'Claimable') -Message "$ExpectedProfile proofStatus must be Claimable after all proof passes."
        Assert-Condition -Condition (@($Entry.missingExternalValidators).Count -eq 0) -Message "$ExpectedProfile claimable proof must not list missing validators."
    } else {
        Assert-Condition -Condition ($Entry.hasRequiredExternalValidation -eq $false) -Message "Product proof contract row $ExpectedProfile in $SourceName must not have required external validation."
        Assert-Condition -Condition ($Entry.canClaimConformance -eq $false) -Message "Product proof contract row $ExpectedProfile in $SourceName must not claim conformance."
        Assert-Condition -Condition (@($Entry.missingExternalValidators).Count -ge 1) -Message "Product proof contract row $ExpectedProfile in $SourceName must list missing or blocking external validators."
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

Assert-Condition -Condition ($proof.schemaVersion -eq 4) -Message 'Unexpected proof schema version.'
Assert-Condition -Condition ($proof.testExitCode -eq 0) -Message "Expected testExitCode 0, got $($proof.testExitCode)."
Assert-Condition -Condition ($null -ne $proof.strictValidatorMode) -Message 'Missing strictValidatorMode in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration) -Message 'Missing validatorConfiguration in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.veraPdfExecutableConfigured) -Message 'Missing veraPdfExecutableConfigured in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.veraPdfArgsConfigured) -Message 'Missing veraPdfArgsConfigured in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.pdfUaValidatorExecutableConfigured) -Message 'Missing pdfUaValidatorExecutableConfigured in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.pdfUaValidatorArgsConfigured) -Message 'Missing pdfUaValidatorArgsConfigured in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.mustangExecutableConfigured) -Message 'Missing mustangExecutableConfigured in proof.json.'
Assert-Condition -Condition ($null -ne $proof.validatorConfiguration.mustangArgsConfigured) -Message 'Missing mustangArgsConfigured in proof.json.'
Assert-Condition -Condition ($proof.contract.unsupportedFormalProfilesFailClosed -eq $true) -Message 'Proof contract must keep unsupported formal profiles fail-closed.'
Assert-Condition -Condition ($proof.contract.formalPdfA2BGenerationEnabled -eq $true) -Message 'Proof contract must enable validator-backed PDF/A-2b generation.'
Assert-Condition -Condition ($proof.contract.formalPdfA3BGenerationEnabled -eq $true) -Message 'Proof contract must enable validator-backed PDF/A-3b generation.'
Assert-Condition -Condition ($proof.contract.formalPdfUa1GenerationEnabled -eq $true) -Message 'Proof contract must enable validator-backed PDF/UA-1 generation.'
Assert-Condition -Condition ($proof.contract.formalElectronicInvoiceGenerationEnabled -eq $true) -Message 'Proof contract must enable validator-backed Factur-X and ZUGFeRD generation.'
Assert-Condition -Condition ($proof.contract.allGeneratedPdfsAreGroundworkFixtures -eq $false) -Message 'Proof contract must distinguish formal PDF/A-2b from groundwork fixtures.'
Assert-Condition -Condition ($proof.contract.externalValidationRequiredForClaims -eq $true) -Message 'Proof contract must require external validation for claims.'
Assert-Condition -Condition ($proof.contract.externalValidationBoundToExactArtifact -eq $true) -Message 'Proof contract must bind external validation to the exact artifact.'
Assert-Condition -Condition ($null -ne $proof.productProofContract) -Message 'Missing productProofContract in proof.json.'
Assert-Condition -Condition ($proof.productProofContract.schemaVersion -eq 4) -Message 'Unexpected productProofContract schema version in proof.json.'
Assert-Condition -Condition ($productProofContractFile.schemaVersion -eq 4) -Message 'Unexpected officeimo-profile-proof-contract.json schema version.'
Assert-Condition -Condition ([string] $proof.productProofContract.generatedBy -eq [string] $productProofContractFile.generatedBy) -Message 'Product proof contract generatedBy mismatch.'
Assert-Condition -Condition ([string] $proof.productProofContract.externalEvidenceMode -eq 'ExactArtifactValidationInjectedByProofExporter') -Message 'Unexpected product proof contract externalEvidenceMode in proof.json.'
Assert-Condition -Condition ([string] $productProofContractFile.externalEvidenceMode -eq 'ExactArtifactValidationInjectedByProofExporter') -Message 'Unexpected externalEvidenceMode in officeimo-profile-proof-contract.json.'

$pdfFixtures = @($proof.pdfFixtures)
Assert-Condition -Condition ($pdfFixtures.Count -eq 5) -Message "Expected 5 PDF fixtures, got $($pdfFixtures.Count)."

$expectedPdfNames = @(
    'officeimo-pdfa2b.pdf',
    'officeimo-pdfa3b.pdf',
    'officeimo-facturx.pdf',
    'officeimo-zugferd.pdf',
    'officeimo-pdfua1.pdf'
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
Assert-Condition -Condition ($validatorDiagnostics.Count -ge 7) -Message "Expected at least 7 validator diagnostics, got $($validatorDiagnostics.Count)."

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
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry.validatorVersion)) -Message "Validator diagnostic entry $($entry.file) is missing validatorVersion."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry.artifactFile)) -Message "Validator diagnostic entry $($entry.file) is missing artifactFile."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry.artifactSha256) -and ([string] $entry.artifactSha256).Length -eq 64) -Message "Validator diagnostic entry $($entry.file) is missing artifactSha256."
    Assert-Condition -Condition ([long] $entry.artifactSizeBytes -gt 0) -Message "Validator diagnostic entry $($entry.file) is missing artifactSizeBytes."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry.validatedAtUtc)) -Message "Validator diagnostic entry $($entry.file) is missing validatedAtUtc."
    Assert-Condition -Condition ($entry.PSObject.Properties.Name -contains 'warnings') -Message "Validator diagnostic entry $($entry.file) is missing warnings."
    $filePath = Join-Path $resolvedProofPath $entry.file
    Assert-Condition -Condition (Test-Path -LiteralPath $filePath) -Message "Missing validator diagnostic file $($entry.file)."
    $file = Get-Item -LiteralPath $filePath
    Assert-Condition -Condition ($file.Length -eq [long] $entry.sizeBytes) -Message "Size mismatch for $($entry.file)."

    $hash = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash.ToLowerInvariant()
    Assert-Condition -Condition ($hash -eq [string] $entry.sha256) -Message "SHA-256 mismatch for $($entry.file)."
}

$profileProofs = @($proof.profileProofs)
Assert-Condition -Condition ($profileProofs.Count -eq 7) -Message "Expected 7 profile proof rows, got $($profileProofs.Count)."

$expectedProfiles = @(
    'pdfa-2b',
    'pdfa-3b',
    'pdfua-1',
    'facturx-pdfa3',
    'facturx',
    'zugferd-pdfa3',
    'zugferd'
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
    $expectedClaim = [string] $entry[0].status -eq 'Passed'
    Assert-Condition -Condition ($entry[0].canClaimConformance -eq $expectedClaim) -Message "Profile proof row $expectedProfile claim state does not match exact-artifact validation."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry[0].formalClaimStatus)) -Message "Profile proof row $expectedProfile is missing formalClaimStatus."
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($entry[0].nextAction)) -Message "Profile proof row $expectedProfile is missing nextAction."
}

$productProofProfiles = @($proof.productProofContract.profiles)
$productProofFileProfiles = @($productProofContractFile.profiles)
Assert-Condition -Condition ($productProofProfiles.Count -eq 5) -Message "Expected 5 product proof contract rows in proof.json, got $($productProofProfiles.Count)."
Assert-Condition -Condition ($productProofFileProfiles.Count -eq 5) -Message "Expected 5 product proof contract rows in officeimo-profile-proof-contract.json, got $($productProofFileProfiles.Count)."

$expectedProductProfiles = @('PdfA2B', 'PdfA3B', 'PdfUa1', 'FacturX', 'Zugferd')
foreach ($expectedProductProfile in $expectedProductProfiles) {
    $entry = @($productProofProfiles | Where-Object { $_.profile -eq $expectedProductProfile })
    $fileEntry = @($productProofFileProfiles | Where-Object { $_.profile -eq $expectedProductProfile })
    Assert-Condition -Condition ($entry.Count -eq 1) -Message "Missing product proof contract row $expectedProductProfile in proof.json."
    Assert-Condition -Condition ($fileEntry.Count -eq 1) -Message "Missing product proof contract row $expectedProductProfile in officeimo-profile-proof-contract.json."
    Assert-ProductProofContractProfile -Entry $entry[0] -ExpectedProfile $expectedProductProfile -SourceName 'proof.json'
    Assert-ProductProofContractProfile -Entry $fileEntry[0] -ExpectedProfile $expectedProductProfile -SourceName 'officeimo-profile-proof-contract.json'
}

Write-Host "PDF compliance proof is valid: $resolvedProofPath"
