param(
    [int] $Threshold = 1000,
    [switch] $IncludeTests,
    [switch] $IncludeAssets
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$excludedDirectoryNames = @('bin', 'obj', 'Ignore', 'OfficeIMO.Examples', 'OfficeIMO.Excel.Benchmarks', 'OfficeIMO.Excel.Benchmarks.LegacyEpPlus')
$rg = Get-Command rg -ErrorAction SilentlyContinue

if (-not $IncludeTests) {
    $excludedDirectoryNames += 'OfficeIMO.Tests'
    $excludedDirectoryNames += 'OfficeIMO.Shared.Tests'
    $excludedDirectoryNames += 'OfficeIMO.CSV.Tests'
    $excludedDirectoryNames += 'OfficeIMO.VerifyTests'
    $excludedDirectoryNames += 'OfficeIMO.MarkdownRenderer.Wpf.Tests'
}

if (-not $IncludeAssets) {
    $excludedDirectoryNames += 'Assets'
    $excludedDirectoryNames += 'OfficeIMO.TestAssets'
}

function Get-CodeFiles {
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    foreach ($item in Get-ChildItem -LiteralPath $Path -Force) {
        if ($item.PSIsContainer) {
            if ($excludedDirectoryNames -contains $item.Name) {
                continue
            }

            Get-CodeFiles -Path $item.FullName
            continue
        }

        if ($item.Extension -eq '.cs' -or $item.Extension -eq '.ps1' -or $item.Extension -eq '.psm1') {
            $item.FullName
        }
    }
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
            '-g', '!OfficeIMO.Shared.Tests/**',
            '-g', '!OfficeIMO.CSV.Tests/**',
            '-g', '!OfficeIMO.VerifyTests/**',
            '-g', '!OfficeIMO.MarkdownRenderer.Wpf.Tests/**'
        )
    }

if (-not $IncludeAssets) {
    $patterns += @('-g', '!Assets/**', '-g', '!OfficeIMO.TestAssets/**')
}

    Push-Location -LiteralPath $root
    try {
        $files = & $rg.Source @patterns |
            ForEach-Object { Join-Path $root $_ }
    } finally {
        Pop-Location
    }
} else {
    $files = Get-CodeFiles -Path $root
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
