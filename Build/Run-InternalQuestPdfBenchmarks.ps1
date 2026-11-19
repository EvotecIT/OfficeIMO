param(
    [Parameter(Mandatory)]
    [string] $QuestPdfPackageVersion,
    [ValidateSet('pdfgenerate', 'pdfinvoice', 'pdfread', 'pdfreverse')]
    [string] $Workload = 'pdfgenerate',
    [ValidateSet('diagnostic', 'standard', 'full')]
    [Alias('Profile', 'RunMode')]
    [string] $BenchmarkProfile = 'diagnostic',
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0',
    [UInt64] $AffinityMask = 0,
    # Pass questpdf-latest once you have reviewed that release's licence; your licence type comes from
    # OFFICEIMO_QUESTPDF_LICENSE_TYPE.
    [string[]] $AcceptLibraryLicense = @(),
    [switch] $KeepArtifacts
)

$ErrorActionPreference = 'Stop'

$temporaryRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
$outputRoot = [System.IO.Path]::GetFullPath((Join-Path $temporaryRoot (
    'OfficeIMO-QuestPDF-Local-' + [Guid]::NewGuid().ToString('N'))))
$requiredPrefix = $temporaryRoot.TrimEnd([System.IO.Path]::DirectorySeparatorChar) +
    [System.IO.Path]::DirectorySeparatorChar
if (-not $outputRoot.StartsWith($requiredPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Local benchmark output '$outputRoot' is outside the temporary directory."
}

Write-Host "QuestPDF $QuestPdfPackageVersion local comparison; results will not enter the website catalog."
try {
    & (Join-Path $PSScriptRoot 'Run-LibraryComparisonBenchmarks.ps1') `
        -Workload $Workload `
        -BenchmarkProfile $BenchmarkProfile `
        -Framework $Framework `
        -AffinityMask $AffinityMask `
        -OutputRoot $outputRoot `
        -EvidenceRoot (Join-Path $outputRoot 'bundle') `
        -LocalOnly `
        -QuestPdfPackageVersion $QuestPdfPackageVersion `
        -AcceptLibraryLicense $AcceptLibraryLicense
} finally {
    if ($KeepArtifacts) {
        Write-Host "Local benchmark artifacts retained at '$outputRoot'."
    } elseif (Test-Path -LiteralPath $outputRoot) {
        Remove-Item -LiteralPath $outputRoot -Recurse -ErrorAction Stop
        Write-Host 'Local benchmark artifacts removed.'
    }
}