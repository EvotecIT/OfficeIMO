<#
.SYNOPSIS
Converts a folder to PDF and records the files that need a review or retry.
.DESCRIPTION
Requires PSWriteOffice 3.0.9 or newer and PowerShell 7.4 or newer. The checked-in Markdown,
RTF and text inputs are synthetic service handover material. Conversion, resource
limits and hash-verified checkpoints belong to OfficeIMO.Workflows. Completed
conversion does not establish visual fidelity or archival compliance.
.PARAMETER ModuleVersion
Optional exact PSWriteOffice version. When omitted, loads an installed version
that satisfies the minimum 3.0.9 requirement and records its actual version.
.PARAMETER ModulePath
Module name or manifest path to import.
.EXAMPLE
./Example-MixedFormat-ConversionReview.ps1 -OutputDirectory ./ConversionReview
.EXAMPLE
./Example-MixedFormat-ConversionReview.ps1 -InputDirectory ./Handover -OutputDirectory ./ConversionReview -RetryFailed
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $OutputDirectory,
    [ValidateNotNullOrEmpty()][string] $InputDirectory = (Join-Path $PSScriptRoot '../Samples/ConversionReview'),
    [ValidateRange(1, 8)][int] $MaximumConcurrency = 2,
    [ValidateRange(1, 10000)][int] $MaximumFiles = 250,
    [ValidateRange(1, [long]::MaxValue)][long] $MaximumInputBytes = 16MB,
    [ValidateRange(1, [long]::MaxValue)][long] $MaximumOutputBytes = 32MB,
    [switch] $RetryFailed,
    [version] $ModuleVersion,
    [string] $ModulePath = 'PSWriteOffice'
)

$ErrorActionPreference = 'Stop'
$moduleImport = @{
    Name = $ModulePath
    MinimumVersion = '3.0.9'
    ErrorAction = 'Stop'
    PassThru = $true
}
if ($PSBoundParameters.ContainsKey('ModuleVersion')) {
    $moduleImport.Remove('MinimumVersion')
    $moduleImport.RequiredVersion = $ModuleVersion
}
$module = Import-Module @moduleImport
$sourceRoot = (Resolve-Path -LiteralPath $InputDirectory -ErrorAction Stop).ProviderPath
$outputRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$pdfRoot = Join-Path $outputRoot 'Pdf'
$checkpointRoot = Join-Path $outputRoot 'State'
$routes = @(Get-OfficeConversionRoute -PdfBatch)
$batchSummary = $null
$batch = @{
    InputDirectory = $sourceRoot
    OutputDirectory = $pdfRoot
    CheckpointDirectory = $checkpointRoot
    MaximumConcurrency = $MaximumConcurrency
    MaximumFiles = $MaximumFiles
    MaximumInputBytes = $MaximumInputBytes
    MaximumOutputBytes = $MaximumOutputBytes
    ConflictPolicy = 'Fail'
    RetryFailed = $RetryFailed
    ItemResults = $true
    SummaryVariable = 'batchSummary'
    MarkdownOptions = (New-OfficeMarkdownPdfOptions -ApplyWordLikeTheme -UseFirstHeadingAsTitle)
}
$items = @(Export-OfficeDocumentPdf @batch)

# Keep each run's review separate. The engine owns and verifies Pdf and State.
$runId = [datetime]::UtcNow.ToString('yyyyMMddTHHmmssfffffffZ')
$reportRoot = Join-Path $outputRoot "Reports/$runId"
$null = New-Item -ItemType Directory -Path $reportRoot
$outcomes = @($items | Sort-Object InputPath | ForEach-Object {
    $item = $_
    $diagnostics = @($item.Diagnostics | ForEach-Object {
        [ordered]@{
            Code = $_.Code; Severity = [string]$_.Severity; Stage = $_.Stage
            Message = $_.Message; Details = $_.Details
        }
    })
    [pscustomobject][ordered]@{
        InputPath = $item.InputPath
        OutputPath = $item.OutputPath
        Status = if ($item.Skipped) { 'Skipped' } else { [string]$item.Status }
        Reused = $item.Reused
        ReviewRequired = $item.Skipped -or [string]$item.Status -ne 'Completed' -or
            @($item.Diagnostics | Where-Object { [string]$_.Severity -ne 'Information' }).Count -gt 0
        Summary = $item.Summary
        Diagnostics = $diagnostics
    }
})
$outcomeColumns = 'InputPath','OutputPath','Status','Reused','ReviewRequired','Summary'
if ($outcomes.Count) {
    $outcomes | Select-Object $outcomeColumns |
        Export-Csv -LiteralPath (Join-Path $reportRoot 'outcomes.csv') -NoTypeInformation -Encoding utf8
} else {
    ($outcomeColumns | ForEach-Object { '"' + $_ + '"' }) -join ',' |
        Set-Content -LiteralPath (Join-Path $reportRoot 'outcomes.csv') -Encoding utf8
}
$diagnosticRows = @(foreach ($outcome in $outcomes) {
    foreach ($diagnostic in $outcome.Diagnostics) {
        [pscustomobject]@{
            InputPath = $outcome.InputPath; Status = $outcome.Status
            Severity = $diagnostic.Severity; Code = $diagnostic.Code
            Stage = $diagnostic.Stage; Message = $diagnostic.Message
        }
    }
})
if ($diagnosticRows.Count) {
    $diagnosticRows | Export-Csv -LiteralPath (Join-Path $reportRoot 'diagnostics.csv') -NoTypeInformation -Encoding utf8
}
$reviewCount = @($outcomes | Where-Object ReviewRequired).Count
$assemblyPath = (Get-Command Export-OfficeDocumentPdf).ImplementingType.Assembly.Location
[ordered]@{
    SchemaVersion = 1; RunId = $runId; ModuleVersion = [string]$module.Version
    ModuleProductVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($assemblyPath).ProductVersion
    ScriptSha256 = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
    PowerShellVersion = $PSVersionTable.PSVersion.ToString()
    SourceDirectory = $sourceRoot; PdfDirectory = $pdfRoot; CheckpointDirectory = $checkpointRoot
    Limits = [ordered]@{ Files = $MaximumFiles; InputBytes = $MaximumInputBytes; OutputBytes = $MaximumOutputBytes; Concurrency = $MaximumConcurrency }
    Summary = $batchSummary; ReviewRequired = $reviewCount
    AvailablePdfRoutes = @($routes | Select-Object Id,SourceExtensions,TargetExtension,PackageId,KnownLimitations)
    Outcomes = $outcomes
} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $reportRoot 'review.json') -Encoding utf8
@"
# Conversion review

Run: $runId. PSWriteOffice $($module.Version), PowerShell $($PSVersionTable.PSVersion).

Selected: $($batchSummary.Selected). Completed: $($batchSummary.Completed).
Reused after hash verification: $($batchSummary.Reused). Failed: $($batchSummary.Failed).
Skipped: $($batchSummary.Skipped). Cancelled: $($batchSummary.Cancelled).
Files with failures, skipped status or warning/error diagnostics: $reviewCount.

Open outcomes.csv for the file list and review.json for complete diagnostics and route limitations.
Inspect every PDF before distribution; successful conversion alone does not prove fidelity.
Select reviewed Completed files from outcomes.csv. Interrupted jobs can retain temporary
staging files for recovery; let the engine finish their publication on restart.

Rerun the same command to verify and reuse completed outputs. To retry corrected failures,
add -RetryFailed. Keep the same source, PDF, state and conversion settings. If a completed
source or PDF changes, use a new output directory after review; do not delete checkpoints
or replace PDFs to bypass their integrity checks. Checkpoints are not a backup.

The default inputs are synthetic. For real work, preserve original files and review the
reported limitations. These distribution PDFs do not establish PDF/A or archival compliance.
"@ | Set-Content -LiteralPath (Join-Path $reportRoot 'Review.md') -Encoding utf8

[pscustomobject]@{
    Selected = $batchSummary.Selected; Completed = $batchSummary.Completed; Reused = $batchSummary.Reused
    Failed = $batchSummary.Failed; Skipped = $batchSummary.Skipped; Cancelled = $batchSummary.Cancelled
    ReviewRequired = $reviewCount; PdfDirectory = $pdfRoot; ReviewDirectory = $reportRoot
}
if ($batchSummary.Cancelled -or $batchSummary.Failed -gt 0 -or $batchSummary.Completed -eq 0) {
    throw "Conversion needs attention. Review $reportRoot before retrying."
}
