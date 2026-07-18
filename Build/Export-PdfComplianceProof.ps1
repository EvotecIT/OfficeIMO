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

    if ($FileName -like '*facturx*') {
        return 'Factur-X'
    }

    if ($FileName -like '*zugferd*') {
        return 'ZUGFeRD'
    }

    if ($FileName -like '*pdfa2b*') {
        return 'PDF/A-2b'
    }

    if ($FileName -like '*pdfa3b*') {
        return 'PDF/A-3b'
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
        $ValidatorConfiguration,

        [Parameter(Mandatory = $true)]
        [string] $DiagnosticFileName
    )

    if ($ValidatorKind -eq 'VeraPdf' -and $ValidatorConfiguration.veraPdfExecutableConfigured) {
        return 'Passed'
    }

    if ($ValidatorKind -eq 'PdfUaValidator' -and $ValidatorConfiguration.pdfUaValidatorExecutableConfigured) {
        return 'Passed'
    }

    if ($ValidatorKind -eq 'Mustang' -and $ValidatorConfiguration.mustangExecutableConfigured) {
        return 'Passed'
    }

    return 'NotRun'
}

function Get-ValidatorVersionFromText {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Text,

        [Parameter(Mandatory = $true)]
        [string] $ValidatorKind
    )

    if ($Text -match '(?i)veraPDF\s+([0-9]+(?:\.[0-9]+)+)') {
        return $Matches[1]
    }

    if ($Text -match '(?i)<releaseDetails\s+id=["'']core["'']\s+version=["'']([0-9]+(?:\.[0-9]+)+)["'']') {
        return $Matches[1]
    }

    if ($ValidatorKind -eq 'Mustang' -and $Text -match '(?i)<validator\s+version=["'']([0-9]+(?:\.[0-9]+)+)["'']') {
        return $Matches[1]
    }

    if ($ValidatorKind -eq 'VeraPdf' -and -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_VERAPDF_VERSION)) {
        return $env:OFFICEIMO_VERAPDF_VERSION
    }

    if ($ValidatorKind -eq 'VeraPdf' -and $env:VERAPDF_CLI_IMAGE -match ':v([^@]+)') {
        return $Matches[1]
    }

    if ($Text -match '(?i)\bversion\s+([0-9]+(?:\.[0-9]+)+)') {
        return $Matches[1]
    }

    return 'unknown'
}

function Get-ArtifactFileFromDiagnosticFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string] $DiagnosticFileName
    )

    switch ($DiagnosticFileName) {
        'verapdf-pdfa2b.txt' { return 'officeimo-pdfa2b.pdf' }
        'verapdf-pdfa3b.txt' { return 'officeimo-pdfa3b.pdf' }
        'verapdf-facturx.txt' { return 'officeimo-facturx.pdf' }
        'mustang-facturx.txt' { return 'officeimo-facturx.pdf' }
        'verapdf-zugferd.txt' { return 'officeimo-zugferd.pdf' }
        'mustang-zugferd.txt' { return 'officeimo-zugferd.pdf' }
        'pdfua-pdfua1.txt' { return 'officeimo-pdfua1.pdf' }
        default { return $null }
    }
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
            profileId = 'pdfa-2b'
            displayName = 'PDF/A-2b'
            fixtureFile = 'officeimo-pdfa2b.pdf'
            validatorKind = 'VeraPdf'
            validatorDiagnosticFileName = 'verapdf-pdfa2b.txt'
            readinessRequirementId = 'verapdf-validation'
            canBecomeClaimable = $true
            formalClaimStatus = 'ValidatorBackedFormalGeneration'
            nextAction = 'Keep the exact-artifact veraPDF gate green for every change to PDF/A generation.'
        },
        [ordered] @{
            profileId = 'pdfa-3b'
            displayName = 'PDF/A-3b'
            fixtureFile = 'officeimo-pdfa3b.pdf'
            validatorKind = 'VeraPdf'
            validatorDiagnosticFileName = 'verapdf-pdfa3b.txt'
            readinessRequirementId = 'verapdf-validation'
            canBecomeClaimable = $true
            formalClaimStatus = 'ValidatorBackedFormalGeneration'
            nextAction = 'Keep the exact-artifact veraPDF gate green for every change to PDF/A-3 generation.'
        },
        [ordered] @{
            profileId = 'pdfua-1'
            displayName = 'PDF/UA-1'
            fixtureFile = 'officeimo-pdfua1.pdf'
            validatorKind = 'PdfUaValidator'
            validatorDiagnosticFileName = 'pdfua-pdfua1.txt'
            readinessRequirementId = 'pdfua-validation'
            canBecomeClaimable = $true
            formalClaimStatus = 'ValidatorBackedFormalGeneration'
            nextAction = 'Keep the exact-artifact PDF/UA-1 validator gate green across text, links, annotations, forms, and figures.'
        },
        [ordered] @{
            profileId = 'facturx-pdfa3'
            displayName = 'Factur-X PDF/A-3 carrier'
            fixtureFile = 'officeimo-facturx.pdf'
            validatorKind = 'VeraPdf'
            validatorDiagnosticFileName = 'verapdf-facturx.txt'
            readinessRequirementId = 'einvoice-verapdf-validation'
            canBecomeClaimable = $true
            formalClaimStatus = 'ValidatorBackedFormalGeneration'
            nextAction = 'Keep the exact Factur-X PDF/A-3 carrier validation green.'
        },
        [ordered] @{
            profileId = 'facturx'
            displayName = 'Factur-X invoice'
            fixtureFile = 'officeimo-facturx.pdf'
            validatorKind = 'Mustang'
            validatorDiagnosticFileName = 'mustang-facturx.txt'
            readinessRequirementId = 'einvoice-mustang-validation'
            canBecomeClaimable = $true
            formalClaimStatus = 'ValidatorBackedFormalGeneration'
            nextAction = 'Keep the exact Factur-X invoice validation green.'
        },
        [ordered] @{
            profileId = 'zugferd-pdfa3'
            displayName = 'ZUGFeRD PDF/A-3 carrier'
            fixtureFile = 'officeimo-zugferd.pdf'
            validatorKind = 'VeraPdf'
            validatorDiagnosticFileName = 'verapdf-zugferd.txt'
            readinessRequirementId = 'einvoice-verapdf-validation'
            canBecomeClaimable = $true
            formalClaimStatus = 'ValidatorBackedFormalGeneration'
            nextAction = 'Keep the exact ZUGFeRD PDF/A-3 carrier validation green.'
        },
        [ordered] @{
            profileId = 'zugferd'
            displayName = 'ZUGFeRD invoice'
            fixtureFile = 'officeimo-zugferd.pdf'
            validatorKind = 'Mustang'
            validatorDiagnosticFileName = 'mustang-zugferd.txt'
            readinessRequirementId = 'einvoice-mustang-validation'
            canBecomeClaimable = $true
            formalClaimStatus = 'ValidatorBackedFormalGeneration'
            nextAction = 'Keep the exact ZUGFeRD invoice validation green.'
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
            canClaimConformance = [bool] $definition.canBecomeClaimable -and $diagnostic -and $diagnostic.status -eq 'Passed' -and $diagnostic.artifactSha256 -eq $fixture.sha256
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
        'PdfA2B' {
            if ($ValidatorKind -eq 'VeraPdf') {
                return 'verapdf-pdfa2b.txt'
            }
        }
        'PdfA3B' {
            if ($ValidatorKind -eq 'VeraPdf') {
                return 'verapdf-pdfa3b.txt'
            }
        }
        'PdfUa1' {
            if ($ValidatorKind -eq 'PdfUaValidator') {
                return 'pdfua-pdfua1.txt'
            }
        }
        'FacturX' {
            if ($ValidatorKind -eq 'VeraPdf') {
                return 'verapdf-facturx.txt'
            }

            if ($ValidatorKind -eq 'Mustang') {
                return 'mustang-facturx.txt'
            }
        }
        'Zugferd' {
            if ($ValidatorKind -eq 'VeraPdf') {
                return 'verapdf-zugferd.txt'
            }

            if ($ValidatorKind -eq 'Mustang') {
                return 'mustang-zugferd.txt'
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
        [array] $DiagnosticRows,

        [Parameter(Mandatory = $true)]
        [array] $PdfRows
    )

    foreach ($profile in @($ProductProofContract.profiles)) {
        $profileFixture = switch ([string] $profile.profile) {
            'PdfA2B' { @($PdfRows | Where-Object { $_.file -eq 'officeimo-pdfa2b.pdf' }) | Select-Object -First 1 }
            'PdfA3B' { @($PdfRows | Where-Object { $_.file -eq 'officeimo-pdfa3b.pdf' }) | Select-Object -First 1 }
            'PdfUa1' { @($PdfRows | Where-Object { $_.file -eq 'officeimo-pdfua1.pdf' }) | Select-Object -First 1 }
            'FacturX' { @($PdfRows | Where-Object { $_.file -eq 'officeimo-facturx.pdf' }) | Select-Object -First 1 }
            'Zugferd' { @($PdfRows | Where-Object { $_.file -eq 'officeimo-zugferd.pdf' }) | Select-Object -First 1 }
            default { $null }
        }
        if ($profileFixture) {
            $profile.hasArtifactEvidence = $true
            $profile.artifactSha256 = $profileFixture.sha256
            $profile.artifactSizeBytes = $profileFixture.sizeBytes
        }

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
                $validatorProof.validatorVersion = $null
                $validatorProof.artifactSha256 = $null
                $validatorProof.artifactSizeBytes = $null
                $validatorProof.validatedAtUtc = $null
                $validatorProof.warnings = @()
            } else {
                $validatorProof.status = $diagnosticRow.status
                $validatorProof.isSatisfied = $diagnosticRow.status -eq 'Passed'
                $validatorProof.blocksConformanceClaim = -not $validatorProof.isSatisfied
                $validatorProof.validatorName = $diagnosticRow.validatorKind
                $validatorProof.diagnostic = "Validator diagnostic status $($diagnosticRow.status) matched expected status $($diagnosticRow.expectedStatus): $($diagnosticRow.file)."
                $validatorProof.profile = $diagnosticRow.profile
                $validatorProof.exitCode = $diagnosticRow.exitCode
                $validatorProof.validatorVersion = $diagnosticRow.validatorVersion
                $validatorProof.artifactSha256 = $diagnosticRow.artifactSha256
                $validatorProof.artifactSizeBytes = $diagnosticRow.artifactSizeBytes
                $validatorProof.validatedAtUtc = $diagnosticRow.validatedAtUtc
                $validatorProof.warnings = @($diagnosticRow.warnings)

                if ($validatorProof.isSatisfied -and
                    (-not $profile.hasArtifactEvidence -or
                     $validatorProof.artifactSha256 -ne $profile.artifactSha256 -or
                     $validatorProof.artifactSizeBytes -ne $profile.artifactSizeBytes)) {
                    $validatorProof.isSatisfied = $false
                    $validatorProof.blocksConformanceClaim = $true
                    $validatorProof.diagnostic = 'Validator passed, but its artifact identity does not match the profile proof artifact.'
                }
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
        $profile.proofStatus = if (-not $profile.isInternallyReady) {
            'InternalGaps'
        } elseif (-not $profile.hasArtifactEvidence) {
            'MissingArtifactEvidence'
        } elseif ($failedExternalValidationCount -gt 0) {
            'ExternalValidationFailed'
        } elseif (-not $hasRequiredExternalValidation) {
            'MissingExternalValidation'
        } else {
            'Claimable'
        }
        $profile.missingExternalValidators = @($missingExternalValidators)
        $profile.failedExternalValidationCount = $failedExternalValidationCount

        $satisfiedRequirementIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
        foreach ($validatorProof in @($profile.externalValidatorProofs)) {
            if (-not $validatorProof.isSatisfied) {
                continue
            }

            switch ([string] $validatorProof.validatorKind) {
                'VeraPdf' { [void] $satisfiedRequirementIds.Add('verapdf-validation') }
                'PdfUaValidator' { [void] $satisfiedRequirementIds.Add('pdfua-validation') }
                'Mustang' { [void] $satisfiedRequirementIds.Add('mustang-validation') }
            }
        }

        $profile.missingRequirementIds = @(
            @($profile.missingRequirementIds) |
                Where-Object { -not $satisfiedRequirementIds.Contains([string] $_) } |
                Select-Object -Unique
        )
        $profile.unsupportedRequirementIds = @(
            @($profile.unsupportedRequirementIds) |
                Where-Object { -not $satisfiedRequirementIds.Contains([string] $_) } |
                Select-Object -Unique
        )
        if ($profile.canClaimConformance -and
            (@($profile.missingRequirementIds).Count -ne 0 -or
             @($profile.unsupportedRequirementIds).Count -ne 0)) {
            throw "Claimable proof '$($profile.profile)' still contains blocking requirement identifiers."
        }
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
    'officeimo-pdfa2b.pdf',
    'officeimo-pdfa3b.pdf',
    'officeimo-facturx.pdf',
    'officeimo-zugferd.pdf',
    'officeimo-pdfua1.pdf'
)

$generatedProofDiagnosticFileNames = @(
    'verapdf-pdfa2b.txt',
    'verapdf-pdfa3b.txt',
    'verapdf-facturx.txt',
    'mustang-facturx.txt',
    'verapdf-zugferd.txt',
    'mustang-zugferd.txt',
    'pdfua-pdfua1.txt'
)

$legacyGeneratedProofFileNames = @(
    'officeimo-pdfa3-groundwork.pdf',
    'officeimo-einvoice-groundwork.pdf',
    'officeimo-pdfua-groundwork.pdf',
    'verapdf-pdfa3-groundwork.txt',
    'verapdf-einvoice-groundwork.txt',
    'mustang-einvoice-groundwork.txt',
    'pdfua-groundwork.txt'
)

$generatedProofFileNames = @(
    $generatedProofPdfFileNames
    $generatedProofDiagnosticFileNames
    $legacyGeneratedProofFileNames
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
$resolvedVeraPdfPath = if (-not [string]::IsNullOrWhiteSpace($VeraPdfPath) -and (Test-Path -LiteralPath $VeraPdfPath)) {
    (Resolve-Path -LiteralPath $VeraPdfPath).Path
} else {
    $VeraPdfPath
}
$resolvedPdfUaValidatorPath = if (-not [string]::IsNullOrWhiteSpace($PdfUaValidatorPath) -and (Test-Path -LiteralPath $PdfUaValidatorPath)) {
    (Resolve-Path -LiteralPath $PdfUaValidatorPath).Path
} else {
    $PdfUaValidatorPath
}
$resolvedMustangPath = if (-not [string]::IsNullOrWhiteSpace($MustangPath) -and (Test-Path -LiteralPath $MustangPath)) {
    (Resolve-Path -LiteralPath $MustangPath).Path
} else {
    $MustangPath
}

$testExitCode = 0
try {
    $env:OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT = $resolvedOutputPath
    if ($RequireValidators) {
        $env:OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS = '1'
    } else {
        $env:OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS = $null
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedVeraPdfPath)) {
        $env:OFFICEIMO_VERAPDF = $resolvedVeraPdfPath
        $env:OFFICEIMO_VERAPDF_PATH = $resolvedVeraPdfPath
    }

    if (-not [string]::IsNullOrWhiteSpace($VeraPdfArgs)) {
        $env:OFFICEIMO_VERAPDF_ARGS = $VeraPdfArgs
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedPdfUaValidatorPath)) {
        $env:OFFICEIMO_PDFUA_VALIDATOR = $resolvedPdfUaValidatorPath
        $env:OFFICEIMO_PDFUA_VALIDATOR_PATH = $resolvedPdfUaValidatorPath
    }

    if (-not [string]::IsNullOrWhiteSpace($PdfUaValidatorArgs)) {
        $env:OFFICEIMO_PDFUA_VALIDATOR_ARGS = $PdfUaValidatorArgs
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedMustangPath)) {
        $env:OFFICEIMO_MUSTANG = $resolvedMustangPath
        $env:OFFICEIMO_MUSTANG_PATH = $resolvedMustangPath
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
$lines.Add('- PDF/A-2b, PDF/A-3b, Factur-X, and ZUGFeRD are formal engine profiles only when their exact generated artifacts pass every required validator.')
$lines.Add('- PDF/UA-1 is formal only when the exact generated artifact passes the accessibility validator across tagged text, links, annotations, forms, figures, language, title, and embedded Unicode fonts.')
$lines.Add('- Unsupported profiles, including PDF/UA-2, remain fail-closed.')
$lines.Add('- External validators are CI/test proof tools and are not runtime dependencies of OfficeIMO.Pdf.')
$lines.Add('- Artifact SHA-256 and byte length bind each validator result to the exact generated PDF used for the conformance decision.')
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
        $expectedStatus = Get-ExpectedValidatorStatus -ValidatorKind $validatorKind -ValidatorConfiguration $validatorConfiguration -DiagnosticFileName $file.Name
        $matchesExpectedStatus = $validatorStatus -eq $expectedStatus
        $artifactFileName = Get-ArtifactFileFromDiagnosticFileName -DiagnosticFileName $file.Name
        $artifactRow = if ([string]::IsNullOrWhiteSpace($artifactFileName)) {
            $null
        } else {
            @($pdfRows | Where-Object { $_.file -eq $artifactFileName }) | Select-Object -First 1
        }
        $warnings = @($diagnosticText -split "`r?`n" | Where-Object { $_ -match '(?i)\bwarning\b' })
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
            validatorVersion = Get-ValidatorVersionFromText -Text $diagnosticText -ValidatorKind $validatorKind
            artifactFile = $artifactFileName
            artifactSha256 = if ($artifactRow) { $artifactRow.sha256 } else { $null }
            artifactSizeBytes = if ($artifactRow) { $artifactRow.sizeBytes } else { $null }
            validatedAtUtc = $generatedAt
            warnings = $warnings
            sizeBytes = $file.Length
            sha256 = $hash
        }
    }
}

$profileRows = @(Get-ProofProfileRows -PdfRows $pdfRows -DiagnosticRows $diagnosticRows)
Update-ProductProofContractWithDiagnostics -ProductProofContract $productProofContract -DiagnosticRows $diagnosticRows -PdfRows $pdfRows
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
    schemaVersion = 4
    generatedUtc = $generatedAt
    commit = $commit
    outputDirectory = $resolvedOutputPath
    testExitCode = $testExitCode
    command = $commandLine
    strictValidatorMode = [bool] $RequireValidators
    validatorConfiguration = $validatorConfiguration
    contract = [ordered] @{
        unsupportedFormalProfilesFailClosed = $true
        formalPdfA2BGenerationEnabled = $true
        formalPdfA3BGenerationEnabled = $true
        formalPdfUa1GenerationEnabled = $true
        formalElectronicInvoiceGenerationEnabled = $true
        allGeneratedPdfsAreGroundworkFixtures = $false
        externalValidationRequiredForClaims = $true
        externalValidationBoundToExactArtifact = $true
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
