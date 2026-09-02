param(
    [string] $ConfigPath = "$PSScriptRoot\project.build.json",
    [Nullable[bool]] $UpdateVersions,
    [Nullable[bool]] $Build,
    [Nullable[bool]] $PublishNuget = $false,
    [Nullable[bool]] $PublishGitHub = $false,
    [Nullable[bool]] $Plan,
    [string] $PlanPath,
    [bool] $RequireHtmlPdfReleaseProof = $false,
    [string] $PdfComplianceProofPath = $env:OFFICEIMO_PDF_COMPLIANCE_PROOF_PATH
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ($RequireHtmlPdfReleaseProof -and $Plan -ne $true) {
    & "$PSScriptRoot/Test-HtmlPdfReleaseGate.ps1" -PdfComplianceProofPath $PdfComplianceProofPath
}

Import-Module PSPublishModule -Force -ErrorAction Stop

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
