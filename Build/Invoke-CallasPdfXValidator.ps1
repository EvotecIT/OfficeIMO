<#
.SYNOPSIS
Adapts qualified callas pdfToolbox PDF/X verification to the OfficeIMO proof gate.

.DESCRIPTION
Maps OfficeIMO PDF/X profile names to the vendor-supplied verification profiles.
Only exit code 0 is accepted as a clean result. Informational, warning, error,
fixup, licensing, configuration, and process failures remain fail-closed.

.EXAMPLE
$env:OFFICEIMO_CALLAS_PDFTOOLBOX = 'C:\Tools\pdfToolbox\pdfToolbox.exe'
$env:OFFICEIMO_CALLAS_PDFX1A_PROFILE = 'C:\Tools\pdfToolbox\Verify compliance with PDFX-1a.kfpx'
$env:OFFICEIMO_CALLAS_PDFX4_PROFILE = 'C:\Tools\pdfToolbox\Verify compliance with PDFX-4.kfpx'
./Build/Export-PdfComplianceProof.ps1 `
    -PdfXValidatorPath ./Build/Invoke-CallasPdfXValidator.ps1 `
    -PdfXValidatorArgs '-Profile {profile} -Pdf {pdf}' `
    -RequireValidators
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Profile,

    [Parameter(Mandatory = $true)]
    [string] $Pdf,

    [string] $PdfToolboxPath = $env:OFFICEIMO_CALLAS_PDFTOOLBOX,
    [string] $PdfX1AProfilePath = $env:OFFICEIMO_CALLAS_PDFX1A_PROFILE,
    [string] $PdfX4ProfilePath = $env:OFFICEIMO_CALLAS_PDFX4_PROFILE
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

function Resolve-RequiredFile {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $Description
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Description is not configured."
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Description was not found at '$Path'."
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

$resolvedTool = Resolve-RequiredFile -Path $PdfToolboxPath -Description 'callas pdfToolbox CLI executable'
$resolvedPdf = Resolve-RequiredFile -Path $Pdf -Description 'PDF/X artifact'
$profilePath = switch ($Profile) {
    { $_ -in @('PDF/X-1a:2003', 'PDF/X-1a', 'PdfX1A2003') } {
        Resolve-RequiredFile -Path $PdfX1AProfilePath -Description 'callas Verify compliance with PDF/X-1a profile'
        break
    }
    { $_ -in @('PDF/X-4', 'PdfX4') } {
        Resolve-RequiredFile -Path $PdfX4ProfilePath -Description 'callas Verify compliance with PDF/X-4 profile'
        break
    }
    default {
        throw "Unsupported PDF/X validation profile '$Profile'."
    }
}

Write-Host "Running qualified PDF/X preflight for $Profile."
$versionOutput = @(& $resolvedTool '--version' 2>&1)
$versionExitCode = $LASTEXITCODE
foreach ($line in $versionOutput) {
    Write-Host $line
}

if ($null -eq $versionExitCode -or $versionExitCode -ne 0) {
    throw "callas pdfToolbox version discovery failed with exit code $versionExitCode."
}

& $resolvedTool '--noprogress' $profilePath $resolvedPdf
$exitCode = $LASTEXITCODE
if ($null -eq $exitCode) {
    throw 'callas pdfToolbox did not report a process exit code.'
}

if ($exitCode -ne 0) {
    Write-Host "PDF/X preflight did not produce a clean verification result (callas exit code $exitCode)."
}

exit $exitCode
