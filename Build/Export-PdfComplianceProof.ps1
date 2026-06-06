param(
    [string] $OutputDirectory = "artifacts/pdf-compliance-proof",
    [string] $Configuration = "Debug",
    [string] $Framework = "net8.0",
    [string] $VeraPdfPath = "",
    [string] $VeraPdfArgs = "",
    [string] $PdfUaValidatorPath = "",
    [string] $PdfUaValidatorArgs = "",
    [string] $MustangPath = "",
    [string] $MustangArgs = "",
    [switch] $NoRestore,
    [switch] $RequireValidators
)

$ErrorActionPreference = 'Stop'

function Get-ValidatorKindFromFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string] $FileName
    )

    if ($FileName -like 'verapdf-*') {
        return 'VeraPdf'
    }

    if ($FileName -like 'mustang-*') {
        return 'Mustang'
    }

    if ($FileName -like 'pdfua-*') {
        return 'PdfUaValidator'
    }

    return 'Custom'
}

function Get-ValidatorProfileFromFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string] $FileName
    )

    if ($FileName -like '*pdfa3*') {
        return 'PDF/A-3b'
    }

    if ($FileName -like '*einvoice*') {
        return 'Factur-X/ZUGFeRD groundwork'
    }

    if ($FileName -like '*pdfua*') {
        return 'PDF/UA-1'
    }

    return ''
}

function Get-ValidatorStatusFromText {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Text
    )

    if ($Text -match 'was not configured') {
        return 'NotRun'
    }

    if ($Text -match 'exited with code\s+0\b') {
        return 'Passed'
    }

    if ($Text -match 'exited with code\s+\d+\b') {
        return 'Failed'
    }

    return 'Error'
}

function Get-ExpectedValidatorStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string] $ValidatorKind,

        [Parameter(Mandatory = $true)]
        $ValidatorConfiguration
    )

    if ($ValidatorKind -eq 'VeraPdf' -and $ValidatorConfiguration.veraPdfExecutableConfigured) {
        return 'Failed'
    }

    if ($ValidatorKind -eq 'PdfUaValidator' -and $ValidatorConfiguration.pdfUaValidatorExecutableConfigured) {
        return 'Failed'
    }

    if ($ValidatorKind -eq 'Mustang' -and $ValidatorConfiguration.mustangExecutableConfigured) {
        return 'Failed'
    }

    return 'NotRun'
}

function Test-AnyCommandAvailable {
    param(
        [Parameter(Mandatory = $true)]
        [string[]] $CommandNames
    )

    foreach ($commandName in $CommandNames) {
        if (Get-Command -Name $commandName -ErrorAction SilentlyContinue) {
            return $true
        }
    }

    return $false
}

function Get-ProofProfileRows {
    param(
        [Parameter(Mandatory = $true)]
        [array] $PdfRows,

        [Parameter(Mandatory = $true)]
        [array] $DiagnosticRows
    )

    $definitions = @(
        [ordered] @{
            profileId = 'pdfa-3b-groundwork'
            displayName = 'PDF/A-3b groundwork'
            fixtureFile = 'officeimo-pdfa3-groundwork.pdf'
            validatorKind = 'VeraPdf'
            readinessRequirementId = 'verapdf-validation'
            formalClaimStatus = 'BlockedUntilFormalPdfAProfileGeneration'
            nextAction = 'Implement profile-specific PDF/A generation and flip the veraPDF gate only after validator success is intentional.'
        },
        [ordered] @{
            profileId = 'pdfua-1-groundwork'
            displayName = 'PDF/UA-1 groundwork'
            fixtureFile = 'officeimo-pdfua-groundwork.pdf'
            validatorKind = 'PdfUaValidator'
            readinessRequirementId = 'pdfua-validation'
            formalClaimStatus = 'BlockedUntilFormalPdfUaProfileGeneration'
            nextAction = 'Implement full tagged structure, reading order, alternate text, font mapping, and flip the PDF/UA validator gate only after validator success is intentional.'
        },
        [ordered] @{
            profileId = 'einvoice-groundwork'
            displayName = 'Factur-X/ZUGFeRD groundwork'
            fixtureFile = 'officeimo-einvoice-groundwork.pdf'
            validatorKind = 'Mustang'
            readinessRequirementId = 'einvoice-mustang-validation'
            formalClaimStatus = 'BlockedUntilFormalEinvoiceProfileGeneration'
            nextAction = 'Implement profile-specific XML, XMP, PDF/A-3 output, and flip the Mustang gate only after validator success is intentional.'
        }
    )

    $rows = @()
    foreach ($definition in $definitions) {
        $fixture = @($PdfRows | Where-Object { $_.file -eq $definition.fixtureFile }) | Select-Object -First 1
        $diagnostic = @($DiagnosticRows | Where-Object { $_.validatorKind -eq $definition.validatorKind }) | Select-Object -First 1

        $rows += [ordered] @{
            profileId = $definition.profileId
            displayName = $definition.displayName
            fixtureFile = $definition.fixtureFile
            fixtureSha256 = if ($fixture) { $fixture.sha256 } else { $null }
            validatorKind = $definition.validatorKind
            validatorDiagnosticFile = if ($diagnostic) { $diagnostic.file } else { $null }
            validatorProfile = if ($diagnostic) { $diagnostic.profile } else { $null }
            status = if ($diagnostic) { $diagnostic.status } else { 'Error' }
            expectedStatus = if ($diagnostic) { $diagnostic.expectedStatus } else { 'Error' }
            matchesExpectedStatus = if ($diagnostic) { $diagnostic.matchesExpectedStatus } else { $false }
            readinessRequirementId = $definition.readinessRequirementId
            canClaimConformance = $false
            formalClaimStatus = $definition.formalClaimStatus
            nextAction = $definition.nextAction
        }
    }

    return $rows
}

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
$outputPath = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory
} else {
    Join-Path $repoRoot $OutputDirectory
}

New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
$resolvedOutputPath = (Resolve-Path -LiteralPath $outputPath).Path

Get-ChildItem -LiteralPath $resolvedOutputPath -File |
    Where-Object { $_.Extension -in '.pdf', '.txt', '.md', '.json' } |
    Remove-Item -Force

$previousProofOutput = $env:OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT
$previousRequireValidators = $env:OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS
$previousVeraPdfPath = $env:OFFICEIMO_VERAPDF_PATH
$previousVeraPdfArgs = $env:OFFICEIMO_VERAPDF_ARGS
$previousPdfUaValidatorPath = $env:OFFICEIMO_PDFUA_VALIDATOR_PATH
$previousPdfUaValidatorArgs = $env:OFFICEIMO_PDFUA_VALIDATOR_ARGS
$previousMustangPath = $env:OFFICEIMO_MUSTANG_PATH
$previousMustangArgs = $env:OFFICEIMO_MUSTANG_ARGS
$validatorConfiguration = $null

$testExitCode = 0
try {
    $env:OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT = $resolvedOutputPath
    if ($RequireValidators) {
        $env:OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS = '1'
    }

    if (-not [string]::IsNullOrWhiteSpace($VeraPdfPath)) {
        $env:OFFICEIMO_VERAPDF_PATH = $VeraPdfPath
    }

    if (-not [string]::IsNullOrWhiteSpace($VeraPdfArgs)) {
        $env:OFFICEIMO_VERAPDF_ARGS = $VeraPdfArgs
    }

    if (-not [string]::IsNullOrWhiteSpace($PdfUaValidatorPath)) {
        $env:OFFICEIMO_PDFUA_VALIDATOR_PATH = $PdfUaValidatorPath
    }

    if (-not [string]::IsNullOrWhiteSpace($PdfUaValidatorArgs)) {
        $env:OFFICEIMO_PDFUA_VALIDATOR_ARGS = $PdfUaValidatorArgs
    }

    if (-not [string]::IsNullOrWhiteSpace($MustangPath)) {
        $env:OFFICEIMO_MUSTANG_PATH = $MustangPath
    }

    if (-not [string]::IsNullOrWhiteSpace($MustangArgs)) {
        $env:OFFICEIMO_MUSTANG_ARGS = $MustangArgs
    }

    $validatorConfiguration = [ordered] @{
        veraPdfExecutableConfigured = -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_VERAPDF) -or -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_VERAPDF_PATH) -or (Test-AnyCommandAvailable -CommandNames @('verapdf', 'verapdf.bat', 'verapdf.exe'))
        veraPdfArgsConfigured = -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_VERAPDF_ARGS)
        pdfUaValidatorExecutableConfigured = -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_PDFUA_VALIDATOR) -or -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_PDFUA_VALIDATOR_PATH) -or (Test-AnyCommandAvailable -CommandNames @('pdfua-validator', 'pdfua-validator.bat', 'pdfua-validator.exe'))
        pdfUaValidatorArgsConfigured = -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_PDFUA_VALIDATOR_ARGS)
        mustangExecutableConfigured = -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_MUSTANG) -or -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_MUSTANG_PATH) -or (Test-AnyCommandAvailable -CommandNames @('mustangproject', 'mustangproject.bat', 'mustangproject.exe', 'mustang', 'mustang.bat', 'mustang.exe'))
        mustangArgsConfigured = -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_MUSTANG_ARGS)
    }

    $testArgs = @(
        'test',
        (Join-Path $repoRoot 'OfficeIMO.Tests/OfficeIMO.Tests.csproj'),
        '--configuration', $Configuration,
        '--framework', $Framework,
        '--filter', 'FullyQualifiedName~PdfComplianceGateTests',
        '--verbosity', 'minimal',
        '-p:WarningLevel=0'
    )

    if ($NoRestore) {
        $testArgs += '--no-restore'
    }

    Push-Location $repoRoot
    try {
        & dotnet @testArgs
        $testExitCode = $LASTEXITCODE
    } finally {
        Pop-Location
    }
} finally {
    $env:OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT = $previousProofOutput
    $env:OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS = $previousRequireValidators
    $env:OFFICEIMO_VERAPDF_PATH = $previousVeraPdfPath
    $env:OFFICEIMO_VERAPDF_ARGS = $previousVeraPdfArgs
    $env:OFFICEIMO_PDFUA_VALIDATOR_PATH = $previousPdfUaValidatorPath
    $env:OFFICEIMO_PDFUA_VALIDATOR_ARGS = $previousPdfUaValidatorArgs
    $env:OFFICEIMO_MUSTANG_PATH = $previousMustangPath
    $env:OFFICEIMO_MUSTANG_ARGS = $previousMustangArgs
}

$commit = (& git -C $repoRoot rev-parse --short HEAD).Trim()
$statusLines = @(& git -C $repoRoot status --short)
$status = ($statusLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine
$generatedAt = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
$pdfFiles = @(Get-ChildItem -LiteralPath $resolvedOutputPath -File -Filter '*.pdf' | Sort-Object Name)
$diagnosticFiles = @(Get-ChildItem -LiteralPath $resolvedOutputPath -File -Filter '*.txt' | Sort-Object Name)
$productProofContractPath = Join-Path $resolvedOutputPath 'officeimo-profile-proof-contract.json'
$indexPath = Join-Path $resolvedOutputPath 'index.md'
$jsonPath = Join-Path $resolvedOutputPath 'proof.json'

if ($pdfFiles.Count -eq 0) {
    throw "No compliance proof PDFs were generated in $resolvedOutputPath. Check the dotnet test filter and OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT wiring."
}

if (-not (Test-Path -LiteralPath $productProofContractPath)) {
    throw "No product proof contract was generated in $resolvedOutputPath. Check PdfComplianceAnalyzer proof contract test wiring."
}

$productProofContract = Get-Content -LiteralPath $productProofContractPath -Raw | ConvertFrom-Json

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('# OfficeIMO PDF Compliance Proof')
$lines.Add('')
$lines.Add("Generated: $generatedAt")
$lines.Add('')
$lines.Add("Commit: ``$commit``")
$lines.Add('')
$lines.Add("Output directory: ``$resolvedOutputPath``")
$lines.Add('')
$lines.Add("Test exit code: ``$testExitCode``")
$lines.Add('')
$lines.Add('Machine-readable summary: [proof.json](proof.json)')
$lines.Add('')
$lines.Add('Command:')
$lines.Add('')
$lines.Add('```powershell')
$commandLine = "Build/Export-PdfComplianceProof.ps1 -OutputDirectory `"$OutputDirectory`" -Configuration `"$Configuration`" -Framework `"$Framework`""
if ($NoRestore) {
    $commandLine += ' -NoRestore'
}
if ($RequireValidators) {
    $commandLine += ' -RequireValidators'
}
if (-not [string]::IsNullOrWhiteSpace($VeraPdfPath)) {
    $commandLine += " -VeraPdfPath `"$VeraPdfPath`""
}
if (-not [string]::IsNullOrWhiteSpace($VeraPdfArgs)) {
    $commandLine += " -VeraPdfArgs `"$VeraPdfArgs`""
}
if (-not [string]::IsNullOrWhiteSpace($PdfUaValidatorPath)) {
    $commandLine += " -PdfUaValidatorPath `"$PdfUaValidatorPath`""
}
if (-not [string]::IsNullOrWhiteSpace($PdfUaValidatorArgs)) {
    $commandLine += " -PdfUaValidatorArgs `"$PdfUaValidatorArgs`""
}
if (-not [string]::IsNullOrWhiteSpace($MustangPath)) {
    $commandLine += " -MustangPath `"$MustangPath`""
}
if (-not [string]::IsNullOrWhiteSpace($MustangArgs)) {
    $commandLine += " -MustangArgs `"$MustangArgs`""
}
$lines.Add($commandLine)
$lines.Add('```')
$lines.Add('')
$lines.Add('## Current Contract')
$lines.Add('')
$lines.Add('- Generated PDFs are groundwork fixtures, not formal conformance claims.')
$lines.Add('- `ComplianceProfile` values other than `None` must still fail closed until validator-backed profile generation is implemented.')
$lines.Add('- veraPDF, PDF/UA validator, and Mustang diagnostics are proof inputs; a matching expected status today means the current non-conformance guardrails are still honest.')
$lines.Add('- Validator `Failed` status is expected when the validator runs against these groundwork fixtures; formal conformance must not pass until profile generation is intentionally implemented.')
$lines.Add('- Product proof contract rows are emitted by `PdfComplianceAnalyzer.AssessProof(...)`; they are the engine-level claimability signal for this proof pack.')
$lines.Add('')
$lines.Add('## Validator Configuration')
$lines.Add('')
$lines.Add('| Setting | Value |')
$lines.Add('| --- | --- |')
$lines.Add("| Strict validator mode | $([bool] $RequireValidators) |")
$lines.Add("| veraPDF executable configured | $($validatorConfiguration.veraPdfExecutableConfigured) |")
$lines.Add("| veraPDF args configured | $($validatorConfiguration.veraPdfArgsConfigured) |")
$lines.Add("| PDF/UA validator executable configured | $($validatorConfiguration.pdfUaValidatorExecutableConfigured) |")
$lines.Add("| PDF/UA validator args configured | $($validatorConfiguration.pdfUaValidatorArgsConfigured) |")
$lines.Add("| Mustang executable configured | $($validatorConfiguration.mustangExecutableConfigured) |")
$lines.Add("| Mustang args configured | $($validatorConfiguration.mustangArgsConfigured) |")
$lines.Add('')
$lines.Add('## PDF Fixtures')
$lines.Add('')
$lines.Add('| File | Size | SHA-256 |')
$lines.Add('| --- | ---: | --- |')
$pdfRows = @()
foreach ($file in $pdfFiles) {
    $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    $name = $file.Name.Replace('|', '\|')
    $lines.Add("| [$name]($name) | $($file.Length) bytes | ``$hash`` |")
    $pdfRows += [ordered] @{
        file = $file.Name
        sizeBytes = $file.Length
        sha256 = $hash
    }
}

$lines.Add('')
$lines.Add('## Validator Diagnostics')
$lines.Add('')
$diagnosticRows = @()
if ($diagnosticFiles.Count -eq 0) {
    $lines.Add('No validator diagnostic files were written. Configure veraPDF, a PDF/UA validator, or Mustang to collect external validator output.')
} else {
    $lines.Add('| Validator | Profile | Status | Expected | Match | File | Size | SHA-256 |')
    $lines.Add('| --- | --- | --- | --- | --- | --- | ---: | --- |')
    foreach ($file in $diagnosticFiles) {
        $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        $diagnosticText = Get-Content -LiteralPath $file.FullName -Raw
        $validatorKind = Get-ValidatorKindFromFileName -FileName $file.Name
        $profile = Get-ValidatorProfileFromFileName -FileName $file.Name
        $validatorStatus = Get-ValidatorStatusFromText -Text $diagnosticText
        $expectedStatus = Get-ExpectedValidatorStatus -ValidatorKind $validatorKind -ValidatorConfiguration $validatorConfiguration
        $matchesExpectedStatus = $validatorStatus -eq $expectedStatus
        $name = $file.Name.Replace('|', '\|')
        $lines.Add("| $validatorKind | $profile | $validatorStatus | $expectedStatus | $matchesExpectedStatus | [$name]($name) | $($file.Length) bytes | ``$hash`` |")
        $diagnosticRows += [ordered] @{
            file = $file.Name
            validatorKind = $validatorKind
            profile = $profile
            status = $validatorStatus
            expectedStatus = $expectedStatus
            matchesExpectedStatus = $matchesExpectedStatus
            sizeBytes = $file.Length
            sha256 = $hash
        }
    }
}

$profileRows = @(Get-ProofProfileRows -PdfRows $pdfRows -DiagnosticRows $diagnosticRows)
$lines.Add('')
$lines.Add('## Profile Proof Matrix')
$lines.Add('')
$lines.Add('| Profile | Fixture | Validator | Status | Expected | Match | Requirement | Claim |')
$lines.Add('| --- | --- | --- | --- | --- | --- | --- | --- |')
foreach ($profileRow in $profileRows) {
    $lines.Add("| $($profileRow.displayName) | $($profileRow.fixtureFile) | $($profileRow.validatorKind) | $($profileRow.status) | $($profileRow.expectedStatus) | $($profileRow.matchesExpectedStatus) | $($profileRow.readinessRequirementId) | $($profileRow.canClaimConformance) |")
}

$lines.Add('')
$lines.Add('## Product Proof Contract')
$lines.Add('')
$lines.Add('Generated by: `' + $productProofContract.generatedBy + '`')
$lines.Add('')
$lines.Add('External evidence mode: `' + $productProofContract.externalEvidenceMode + '`')
$lines.Add('')
$lines.Add('| Profile | Internal Ready | External Validation | Claim | Missing Validators | Unsupported Requirements |')
$lines.Add('| --- | --- | --- | --- | --- | --- |')
foreach ($profile in @($productProofContract.profiles)) {
    $missingValidators = (@($profile.missingExternalValidators) -join ', ')
    $unsupportedRequirements = (@($profile.unsupportedRequirementIds) -join ', ')
    if ([string]::IsNullOrWhiteSpace($missingValidators)) {
        $missingValidators = 'none'
    }

    if ([string]::IsNullOrWhiteSpace($unsupportedRequirements)) {
        $unsupportedRequirements = 'none'
    }

    $lines.Add("| $($profile.displayName) | $($profile.isInternallyReady) | $($profile.hasRequiredExternalValidation) | $($profile.canClaimConformance) | $missingValidators | $unsupportedRequirements |")
}

if (-not [string]::IsNullOrWhiteSpace($status)) {
    $lines.Add('')
    $lines.Add('## Working Tree Note')
    $lines.Add('')
    $lines.Add('The repository had uncommitted changes when this proof pack was generated:')
    $lines.Add('')
    $lines.Add('```text')
    $lines.Add($status)
    $lines.Add('```')
}

[System.IO.File]::WriteAllLines($indexPath, $lines, [System.Text.Encoding]::UTF8)
$proof = [ordered] @{
    schemaVersion = 3
    generatedUtc = $generatedAt
    commit = $commit
    outputDirectory = $resolvedOutputPath
    testExitCode = $testExitCode
    command = $commandLine
    strictValidatorMode = [bool] $RequireValidators
    validatorConfiguration = $validatorConfiguration
    contract = [ordered] @{
        formalProfilesFailClosed = $true
        generatedPdfsAreGroundworkFixtures = $true
        externalValidationRequiredForClaims = $true
    }
    pdfFixtures = @($pdfRows)
    validatorDiagnostics = @($diagnosticRows)
    profileProofs = @($profileRows)
    productProofContract = $productProofContract
    workingTreeStatus = @($statusLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}
$proof | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
Write-Host "PDF compliance proof written to $resolvedOutputPath"
Write-Host "Index: $indexPath"
Write-Host "JSON: $jsonPath"

if ($testExitCode -ne 0) {
    throw "dotnet test failed with exit code $testExitCode. Proof artifacts were still written to $resolvedOutputPath."
}
