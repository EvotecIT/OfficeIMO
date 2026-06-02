param(
    [int] $Threshold = 1000,
    [switch] $IncludeTests,
    [switch] $IncludeAssets
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$excludedSegments = @('\bin\', '\obj\', '\Ignore\', '\OfficeIMO.Examples\', '\OfficeIMO.Excel.Benchmarks\', '\OfficeIMO.Excel.Benchmarks.LegacyEpPlus\')
$rg = Get-Command rg -ErrorAction SilentlyContinue

if (-not $IncludeTests) {
    $excludedSegments += '\OfficeIMO.Tests\'
    $excludedSegments += '\OfficeIMO.CSV.Tests\'
    $excludedSegments += '\OfficeIMO.VerifyTests\'
    $excludedSegments += '\OfficeIMO.MarkdownRenderer.Wpf.Tests\'
}

if (-not $IncludeAssets) {
    $excludedSegments += '\Assets\'
}

if ($rg) {
    $patterns = @(
        '--files',
        '-g', '*.cs',
        '-g', '*.ps1',
        '-g', '*.psm1',
        '-g', '!**/bin/**',
        '-g', '!**/obj/**',
        '-g', '!Ignore/**',
        '-g', '!OfficeIMO.Examples/**',
        '-g', '!OfficeIMO.Excel.Benchmarks/**',
        '-g', '!OfficeIMO.Excel.Benchmarks.LegacyEpPlus/**'
    )

    if (-not $IncludeTests) {
        $patterns += @(
            '-g', '!OfficeIMO.Tests/**',
            '-g', '!OfficeIMO.CSV.Tests/**',
            '-g', '!OfficeIMO.VerifyTests/**',
            '-g', '!OfficeIMO.MarkdownRenderer.Wpf.Tests/**'
        )
    }

    if (-not $IncludeAssets) {
        $patterns += @('-g', '!Assets/**')
    }

    $files = & $rg.Source @patterns |
        ForEach-Object { Join-Path $root $_ }
} else {
    $files = Get-ChildItem -LiteralPath $root -Recurse -File -Include *.cs,*.ps1,*.psm1 |
        Where-Object {
            $path = $_.FullName
            foreach ($segment in $excludedSegments) {
                if ($path.Contains($segment)) {
                    return $false
                }
            }

            return $true
        } |
        ForEach-Object { $_.FullName }
}

$results = foreach ($file in $files) {
    $lineCount = [System.IO.File]::ReadAllLines($file).Length
    if ($lineCount -ge $Threshold) {
        [pscustomobject]@{
            Lines = $lineCount
            Path = [System.IO.Path]::GetRelativePath($root, $file)
        }
    }
}

$results | Sort-Object Lines -Descending
