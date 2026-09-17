param(
    [Parameter(Mandatory)]
    [string] $QuestPdfPackageVersion,
    [ValidateSet('pdfgenerate', 'pdfinvoice', 'pdfread', 'pdfreverse')]
    [string] $Workload = 'pdfgenerate',
    [ValidateSet('quick', 'full')]
    [string] $RunMode = 'quick',
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0',
    [ValidateSet('Community', 'Evaluation', 'Professional', 'Enterprise')]
    [string] $QuestPdfLicenseType = 'Community',
    [UInt64] $AffinityMask = 0,
    [switch] $ConfirmQuestPdfAuthorization,
    [switch] $KeepArtifacts
)

$ErrorActionPreference = 'Stop'

if (-not $ConfirmQuestPdfAuthorization) {
    throw @'
This internal-only runner does not grant rights to use the selected QuestPDF version.
Review that version's embedded license and rerun with -ConfirmQuestPdfAuthorization
only when you have permission or another valid basis for the comparison.
'@
}

$temporaryRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
$outputRoot = [System.IO.Path]::GetFullPath((Join-Path $temporaryRoot (
    'OfficeIMO-QuestPDF-Internal-' + [Guid]::NewGuid().ToString('N'))))
$requiredPrefix = $temporaryRoot.TrimEnd([System.IO.Path]::DirectorySeparatorChar) +
    [System.IO.Path]::DirectorySeparatorChar
if (-not $outputRoot.StartsWith($requiredPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Internal benchmark output '$outputRoot' is outside the temporary directory."
}

Write-Host "QuestPDF $QuestPdfPackageVersion internal-only comparison; results will not enter the website catalog."
try {
    & (Join-Path $PSScriptRoot 'Run-LibraryComparisonBenchmarks.ps1') `
        -Workload $Workload `
        -RunMode $RunMode `
        -Framework $Framework `
        -AffinityMask $AffinityMask `
        -OutputRoot $outputRoot `
        -QuestPdfPackageVersion $QuestPdfPackageVersion `
        -QuestPdfLicenseType $QuestPdfLicenseType `
        -InternalQuestPdf `
        -ConfirmQuestPdfAuthorization
} finally {
    if ($KeepArtifacts) {
        Write-Host "Internal benchmark artifacts retained at '$outputRoot'."
    } elseif (Test-Path -LiteralPath $outputRoot) {
        Remove-Item -LiteralPath $outputRoot -Recurse -ErrorAction Stop
        Write-Host 'Internal benchmark artifacts removed.'
    }
}
