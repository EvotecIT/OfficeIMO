function Get-BenchmarkEvidenceLocation {
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
        ResultPath = "/data/benchmarks/library-comparisons/$fileName"
    }
}

function Write-BenchmarkEvidenceResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Path,
        [Parameter(Mandatory)]
        [object] $InputObject
    )

    $benchmarkJsonType = [Type]::GetType('PowerForge.BenchmarkJson, PowerForge', $true)
    $writeMethod = $benchmarkJsonType.GetMethods() |
        Where-Object { $_.Name -eq 'Write' -and $_.IsGenericMethodDefinition } |
        Select-Object -First 1
    if ($null -eq $writeMethod) {
        throw 'PowerForge.BenchmarkJson.Write<T> is unavailable.'
    }

    $value = $InputObject.PSObject.BaseObject
    $closedMethod = $writeMethod.MakeGenericMethod($value.GetType())
    $arguments = [object[]] @([string] $Path, $value)
    $null = $closedMethod.Invoke($null, $arguments)
}
