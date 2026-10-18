function Resolve-OfficeIMOSourceRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    $resolvedRoot = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    $requiredSourceFiles = @(
        'OfficeIMO.sln',
        'OfficeIMO.CSV/OfficeIMO.CSV.csproj',
        'OfficeIMO.Excel/OfficeIMO.Excel.csproj'
    )
    $missingSourceFiles = @(
        foreach ($relativePath in $requiredSourceFiles) {
            if (-not (Test-Path -LiteralPath (Join-Path $resolvedRoot $relativePath) -PathType Leaf)) {
                $relativePath
            }
        }
    )
    if ($missingSourceFiles.Count -gt 0) {
        throw "OfficeIMORoot '$resolvedRoot' is not a complete OfficeIMO source root. Missing: $($missingSourceFiles -join ', ')."
    }

    $resolvedRoot
}
