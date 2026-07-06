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
        [string] $Text,

        [Parameter(Mandatory = $true)]
        [string] $ValidatorKind
    )

    if ($Text -match 'was not configured') {
        return 'NotRun'
    }

    if ($Text -match 'exited with code\s+0\b') {
        return 'Passed'
    }

    if ($Text -match 'exited with code\s+\d+\b') {
        if (Test-ValidatorDiagnosticFailureEvidence -Text $Text -ValidatorKind $ValidatorKind) {
            return 'Failed'
        }

        return 'Error'
    }

    return 'Error'
}

function Test-ValidatorDiagnosticFailureEvidence {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Text,

        [Parameter(Mandatory = $true)]
        [string] $ValidatorKind
    )

    switch ($ValidatorKind) {
        'VeraPdf' {
            return $Text -match '(?i)isCompliant\s*=\s*["'']?false' -or
                $Text -match '(?i)\bnot\s+compliant\b' -or
                $Text -match '(?i)\bvalidation\s+failed\b'
        }
        'PdfUaValidator' {
            return $Text -match '(?i)isCompliant\s*=\s*["'']?false' -or
                $Text -match '(?i)\bnot\s+compliant\b' -or
                $Text -match '(?i)\bpdf/ua\b.*\b(fail|failed|not\s+compliant|violation)' -or
                $Text -match '(?i)\bvalidation\s+failed\b' -or
                $Text -match '(?i)\bviolations?\b'
        }
        'Mustang' {
            return $Text -match '(?i)\bvalidation\s+(failed|result\s*:\s*invalid)\b' -or
                $Text -match '(?i)\bnot\s+(a\s+)?valid\b' -or
                $Text -match '(?i)\bnon-?compliant\b'
        }
        default {
            return $false
        }
    }
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
            validatorDiagnosticFileName = 'verapdf-pdfa3-groundwork.txt'
            readinessRequirementId = 'verapdf-validation'
            formalClaimStatus = 'BlockedUntilFormalPdfAProfileGeneration'
            nextAction = 'Implement profile-specific PDF/A generation and flip the veraPDF gate only after validator success is intentional.'
        },
        [ordered] @{
            profileId = 'pdfua-1-groundwork'
            displayName = 'PDF/UA-1 groundwork'
            fixtureFile = 'officeimo-pdfua-groundwork.pdf'
            validatorKind = 'PdfUaValidator'
            validatorDiagnosticFileName = 'pdfua-groundwork.txt'
            readinessRequirementId = 'pdfua-validation'
            formalClaimStatus = 'BlockedUntilFormalPdfUaProfileGeneration'
            nextAction = 'Implement full tagged structure, reading order, alternate text, font mapping, and flip the PDF/UA validator gate only after validator success is intentional.'
        },
        [ordered] @{
            profileId = 'einvoice-pdfa3-groundwork'
            displayName = 'Factur-X/ZUGFeRD PDF/A-3 groundwork'
            fixtureFile = 'officeimo-einvoice-groundwork.pdf'
            validatorKind = 'VeraPdf'
            validatorDiagnosticFileName = 'verapdf-einvoice-groundwork.txt'
            readinessRequirementId = 'einvoice-verapdf-validation'
            formalClaimStatus = 'BlockedUntilFormalEinvoiceProfileGeneration'
            nextAction = 'Validate the actual e-invoice PDF/A-3 carrier with veraPDF before any Factur-X/ZUGFeRD claim.'
        },
        [ordered] @{
            profileId = 'einvoice-groundwork'
            displayName = 'Factur-X/ZUGFeRD groundwork'
            fixtureFile = 'officeimo-einvoice-groundwork.pdf'
            validatorKind = 'Mustang'
            validatorDiagnosticFileName = 'mustang-einvoice-groundwork.txt'
            readinessRequirementId = 'einvoice-mustang-validation'
            formalClaimStatus = 'BlockedUntilFormalEinvoiceProfileGeneration'
            nextAction = 'Implement profile-specific XML, XMP, PDF/A-3 output, and flip the Mustang gate only after validator success is intentional.'
        }
    )

    $rows = @()
    foreach ($definition in $definitions) {
        $fixture = @($PdfRows | Where-Object { $_.file -eq $definition.fixtureFile }) | Select-Object -First 1
        $diagnostic = @($DiagnosticRows | Where-Object { $_.file -eq $definition.validatorDiagnosticFileName }) | Select-Object -First 1

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

function Get-ProductProofDiagnosticFileName {
    param(
        [Parameter(Mandatory = $true)]
        $Profile,

        [Parameter(Mandatory = $true)]
        [string] $ValidatorKind
    )

    switch ([string] $Profile.profile) {
        'PdfA3B' {
            if ($ValidatorKind -eq 'VeraPdf') {
                return 'verapdf-pdfa3-groundwork.txt'
            }
        }
        'PdfUa1' {
            if ($ValidatorKind -eq 'PdfUaValidator') {
                return 'pdfua-groundwork.txt'
            }
        }
        { $_ -eq 'FacturX' -or $_ -eq 'Zugferd' } {
            if ($ValidatorKind -eq 'VeraPdf') {
                return 'verapdf-einvoice-groundwork.txt'
            }

            if ($ValidatorKind -eq 'Mustang') {
                return 'mustang-einvoice-groundwork.txt'
            }
        }
    }

    return $null
}

function Update-ProductProofContractWithDiagnostics {
    param(
        [Parameter(Mandatory = $true)]
        $ProductProofContract,

        [Parameter(Mandatory = $true)]
        [array] $DiagnosticRows
    )

    foreach ($profile in @($ProductProofContract.profiles)) {
        $hasRequiredExternalValidation = $true
        $missingExternalValidators = [System.Collections.Generic.List[string]]::new()
        $failedExternalValidationCount = 0

        foreach ($validatorProof in @($profile.externalValidatorProofs)) {
            $diagnosticFileName = Get-ProductProofDiagnosticFileName -Profile $profile -ValidatorKind $validatorProof.validatorKind
            $diagnosticRow = if ([string]::IsNullOrWhiteSpace($diagnosticFileName)) {
                $null
            } else {
                @($DiagnosticRows | Where-Object { $_.file -eq $diagnosticFileName }) | Select-Object -First 1
            }

            if ($null -eq $diagnosticRow) {
                $validatorProof.status = 'Missing'
                $validatorProof.isSatisfied = $false
                $validatorProof.blocksConformanceClaim = $true
                $validatorProof.validatorName = $validatorProof.validatorKind
                $validatorProof.diagnostic = 'Missing external validation.'
                $validatorProof.profile = $null
                $validatorProof.exitCode = $null
            } else {
                $validatorProof.status = $diagnosticRow.status
                $validatorProof.isSatisfied = $diagnosticRow.status -eq 'Passed'
                $validatorProof.blocksConformanceClaim = -not $validatorProof.isSatisfied
                $validatorProof.validatorName = $diagnosticRow.validatorKind
                $validatorProof.diagnostic = "Validator diagnostic status $($diagnosticRow.status) matched expected status $($diagnosticRow.expectedStatus): $($diagnosticRow.file)."
                $validatorProof.profile = $diagnosticRow.profile
                $validatorProof.exitCode = $diagnosticRow.exitCode
            }

            if (-not $validatorProof.isSatisfied) {
                $missingExternalValidators.Add([string] $validatorProof.validatorKind)
                $hasRequiredExternalValidation = $false
            }

            if ($validatorProof.status -eq 'Failed' -or $validatorProof.status -eq 'Error') {
                $failedExternalValidationCount++
            }
        }

        $profile.hasRequiredExternalValidation = $hasRequiredExternalValidation
        $profile.canClaimConformance = [bool] $profile.isInternallyReady -and $hasRequiredExternalValidation -and $failedExternalValidationCount -eq 0
        $profile.missingExternalValidators = @($missingExternalValidators)
        $profile.failedExternalValidationCount = $failedExternalValidationCount
    }
}

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
$outputPath = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory
} else {
    Join-Path $repoRoot $OutputDirectory
}

New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
$resolvedOutputPath = (Resolve-Path -LiteralPath $outputPath).Path

$generatedProofPdfFileNames = @(
    'officeimo-pdfa3-groundwork.pdf',
    'officeimo-einvoice-groundwork.pdf',
    'officeimo-pdfua-groundwork.pdf'
)

$generatedProofDiagnosticFileNames = @(
    'verapdf-pdfa3-groundwork.txt',
    'verapdf-einvoice-groundwork.txt',
    'mustang-einvoice-groundwork.txt',
    'pdfua-groundwork.txt'
)

$generatedProofFileNames = @(
    $generatedProofPdfFileNames
    $generatedProofDiagnosticFileNames
    'officeimo-profile-proof-contract.json',
    'index.md',
    'proof.json'
)

foreach ($fileName in $generatedProofFileNames) {
    $path = Join-Path $resolvedOutputPath $fileName
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

$previousProofOutput = $env:OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT
$previousRequireValidators = $env:OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS
$previousVeraPdfExecutable = $env:OFFICEIMO_VERAPDF
$previousVeraPdfPath = $env:OFFICEIMO_VERAPDF_PATH
$previousVeraPdfArgs = $env:OFFICEIMO_VERAPDF_ARGS
$previousPdfUaValidatorExecutable = $env:OFFICEIMO_PDFUA_VALIDATOR
$previousPdfUaValidatorPath = $env:OFFICEIMO_PDFUA_VALIDATOR_PATH
$previousPdfUaValidatorArgs = $env:OFFICEIMO_PDFUA_VALIDATOR_ARGS
$previousMustangExecutable = $env:OFFICEIMO_MUSTANG
$previousMustangPath = $env:OFFICEIMO_MUSTANG_PATH
$previousMustangArgs = $env:OFFICEIMO_MUSTANG_ARGS
$validatorConfiguration = $null

$testExitCode = 0
try {
    $env:OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT = $resolvedOutputPath
    if ($RequireValidators) {
        $env:OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS = '1'
    } else {
        $env:OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS = $null
    }

    if (-not [string]::IsNullOrWhiteSpace($VeraPdfPath)) {
        $env:OFFICEIMO_VERAPDF = $VeraPdfPath
        $env:OFFICEIMO_VERAPDF_PATH = $VeraPdfPath
    }

    if (-not [string]::IsNullOrWhiteSpace($VeraPdfArgs)) {
        $env:OFFICEIMO_VERAPDF_ARGS = $VeraPdfArgs
    }

    if (-not [string]::IsNullOrWhiteSpace($PdfUaValidatorPath)) {
        $env:OFFICEIMO_PDFUA_VALIDATOR = $PdfUaValidatorPath
        $env:OFFICEIMO_PDFUA_VALIDATOR_PATH = $PdfUaValidatorPath
    }

    if (-not [string]::IsNullOrWhiteSpace($PdfUaValidatorArgs)) {
        $env:OFFICEIMO_PDFUA_VALIDATOR_ARGS = $PdfUaValidatorArgs
    }

    if (-not [string]::IsNullOrWhiteSpace($MustangPath)) {
        $env:OFFICEIMO_MUSTANG = $MustangPath
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
        (Join-Path $repoRoot 'OfficeIMO.Pdf.Tests/OfficeIMO.Pdf.Tests.csproj'),
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
    $env:OFFICEIMO_VERAPDF = $previousVeraPdfExecutable
    $env:OFFICEIMO_VERAPDF_PATH = $previousVeraPdfPath
    $env:OFFICEIMO_VERAPDF_ARGS = $previousVeraPdfArgs
    $env:OFFICEIMO_PDFUA_VALIDATOR = $previousPdfUaValidatorExecutable
    $env:OFFICEIMO_PDFUA_VALIDATOR_PATH = $previousPdfUaValidatorPath
    $env:OFFICEIMO_PDFUA_VALIDATOR_ARGS = $previousPdfUaValidatorArgs
    $env:OFFICEIMO_MUSTANG = $previousMustangExecutable
    $env:OFFICEIMO_MUSTANG_PATH = $previousMustangPath
    $env:OFFICEIMO_MUSTANG_ARGS = $previousMustangArgs
}

$commit = (& git -C $repoRoot rev-parse --short HEAD).Trim()
$statusLines = @(& git -C $repoRoot status --short)
$status = ($statusLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine
$generatedAt = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
$pdfFiles = @(
    foreach ($fileName in $generatedProofPdfFileNames) {
        $path = Join-Path $resolvedOutputPath $fileName
        if (Test-Path -LiteralPath $path) {
            Get-Item -LiteralPath $path
        }
    }
) | Sort-Object Name
$diagnosticFiles = @(
    foreach ($fileName in $generatedProofDiagnosticFileNames) {
        $path = Join-Path $resolvedOutputPath $fileName
        if (Test-Path -LiteralPath $path) {
            Get-Item -LiteralPath $path
        }
    }
) | Sort-Object Name
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
        $validatorStatus = Get-ValidatorStatusFromText -Text $diagnosticText -ValidatorKind $validatorKind
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
            exitCode = if ($diagnosticText -match 'exited with code\s+(\d+)\b') { [int] $Matches[1] } else { $null }
            sizeBytes = $file.Length
            sha256 = $hash
        }
    }
}

$profileRows = @(Get-ProofProfileRows -PdfRows $pdfRows -DiagnosticRows $diagnosticRows)
Update-ProductProofContractWithDiagnostics -ProductProofContract $productProofContract -DiagnosticRows $diagnosticRows
$productProofContract | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $productProofContractPath -Encoding UTF8
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

$lines.Add('')
$lines.Add('### External Validator Proof Rows')
$lines.Add('')
$lines.Add('| Profile | Validator | Status | Blocks Claim | Diagnostic |')
$lines.Add('| --- | --- | --- | --- | --- |')
foreach ($profile in @($productProofContract.profiles)) {
    foreach ($validatorProof in @($profile.externalValidatorProofs)) {
        $diagnostic = [string] $validatorProof.diagnostic
        if ([string]::IsNullOrWhiteSpace($diagnostic)) {
            $diagnostic = 'none'
        }

        $diagnostic = $diagnostic.Replace('|', '\|')
        $lines.Add("| $($profile.displayName) | $($validatorProof.validatorKind) | $($validatorProof.status) | $($validatorProof.blocksConformanceClaim) | $diagnostic |")
    }
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
