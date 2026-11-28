param(
    [string] $ConfigPath = "$PSScriptRoot\project.build.json",
    [Nullable[bool]] $UpdateVersions,
    [Nullable[bool]] $Build,
    [Nullable[bool]] $PublishNuget = $false,
    [Nullable[bool]] $PublishGitHub = $false,
    [Nullable[bool]] $Plan,
    [string] $PlanPath,
    # Opt in to HTML/PDF qualification when publishing; ordinary package releases do not require PDF/X proof.
    [switch] $ValidateHtmlPdfRelease,
    [string] $PdfComplianceProofPath = $env:OFFICEIMO_PDF_COMPLIANCE_PROOF_PATH
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ($ValidateHtmlPdfRelease -and
    ($PublishNuget -eq $true -or $PublishGitHub -eq $true) -and $Plan -ne $true) {
    & "$PSScriptRoot/Test-HtmlPdfReleaseGate.ps1" -PdfComplianceProofPath $PdfComplianceProofPath
}

# Package builds need PSPublishModule 3.0.196, which packs OfficeIMO.Tool's per-platform packages; planning doesn't pack.
$psPublishModuleMinimum = if ($Plan -eq $true) { '3.0.194' } else { '3.0.196' }
Import-Module PSPublishModule -MinimumVersion $psPublishModuleMinimum -Force -ErrorAction Stop

if (($PublishNuget -eq $true -or $PublishGitHub -eq $true) -and $Plan -ne $true) {
    if ($ConfigPath -ne "$PSScriptRoot\project.build.json") {
        throw 'Use the canonical project configuration for a coordinated OfficeIMO release.'
    }
    & "$PSScriptRoot/Build-Release.ps1" -RunMode Publish -SignModule:$true
    return
}

$invokeParams = @{
    ConfigPath = $ConfigPath
}
if ($null -ne $UpdateVersions) { $invokeParams.UpdateVersions = $UpdateVersions }
if ($null -ne $Build) { $invokeParams.Build = $Build }
if ($null -ne $PublishNuget) { $invokeParams.PublishNuget = $PublishNuget }
if ($null -ne $PublishGitHub) { $invokeParams.PublishGitHub = $PublishGitHub }
if ($null -ne $Plan) { $invokeParams.Plan = $Plan }
if ($PlanPath) { $invokeParams.PlanPath = $PlanPath }

$result = Invoke-ProjectBuild @invokeParams
$result

if ($null -ne $result -and
    $result.PSObject.Properties.Name -contains 'Success' -and
    -not $result.Success) {
    $message = if ([string]::IsNullOrWhiteSpace([string] $result.ErrorMessage)) {
        'Project build failed without an error message.'
    } else {
        [string] $result.ErrorMessage
    }
    throw $message
}

if ($Plan -ne $true -and $Build -ne $false) {
    $config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
    $staging = Join-Path (Join-Path (Split-Path -Parent $ConfigPath) $config.RootPath) $config.StagingPath
    if (Get-ChildItem -LiteralPath $staging -Filter *.nupkg -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1) {
        $python = Get-Command python3, python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $python) {
            throw 'Python is required to check third-party notices in the built packages.'
        }
        & $python.Source "$PSScriptRoot/validate_third_party_notices.py" package $staging
        if ($LASTEXITCODE -ne 0) {
            throw 'Built packages do not declare or ship their third-party notices.'
        }
    }
}
