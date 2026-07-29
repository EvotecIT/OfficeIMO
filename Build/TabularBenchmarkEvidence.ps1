function Get-TabularBenchmarkEvidenceLocation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $ComparisonId,
        [Parameter(Mandatory)]
        [ValidateSet('windows', 'linux', 'macos')]
        [string] $Platform,
        [Parameter(Mandatory)]
        [ValidateSet('quick', 'full')]
        [string] $RunMode,
        [Parameter(Mandatory)]
        [string] $StaticRoot
    )

    $comparisonSlug = ($ComparisonId.Trim() -replace '[^A-Za-z0-9._-]', '-').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($comparisonSlug)) {
        throw 'ComparisonId must contain at least one filename-safe character.'
    }

    $fileName = "$comparisonSlug-$Platform-$RunMode.json"
    [pscustomobject]@{
        FileName = $fileName
        Path = Join-Path $StaticRoot $fileName
        ResultPath = "/data/benchmarks/tabular/$fileName"
    }
}
