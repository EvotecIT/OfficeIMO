param(
    [string] $OutputDirectory = "artifacts/pdf-visual-review",
    [string] $Configuration = "Debug",
    [string] $Framework = "net8.0",
    [switch] $NoRestore,
    [switch] $RequireRasterizer
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
$outputPath = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory
} else {
    Join-Path $repoRoot $OutputDirectory
}

New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
$resolvedOutputPath = (Resolve-Path -LiteralPath $outputPath).Path

Get-ChildItem -LiteralPath $resolvedOutputPath -File -Filter '*.pdf' | Remove-Item -Force
$indexPath = Join-Path $resolvedOutputPath 'index.md'
if (Test-Path -LiteralPath $indexPath) {
    Remove-Item -LiteralPath $indexPath -Force
}

$previousReviewOutput = $env:OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT
$previousRequireRasterizer = $env:OFFICEIMO_REQUIRE_PDF_RASTERIZER

try {
    $env:OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT = $resolvedOutputPath
    if ($RequireRasterizer) {
        $env:OFFICEIMO_REQUIRE_PDF_RASTERIZER = '1'
    }

    $testArgs = @(
        'test',
        (Join-Path $repoRoot 'OfficeIMO.Tests/OfficeIMO.Tests.csproj'),
        '--configuration', $Configuration,
        '--framework', $Framework,
        '--filter', 'FullyQualifiedName~PdfDocumentRasterVisualBaselineTests',
        '--verbosity', 'minimal',
        '-p:WarningLevel=0'
    )

    if ($NoRestore) {
        $testArgs += '--no-restore'
    }

    Push-Location $repoRoot
    try {
        & dotnet @testArgs
        if ($LASTEXITCODE -ne 0) {
            throw "dotnet test failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }
} finally {
    $env:OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT = $previousReviewOutput
    $env:OFFICEIMO_REQUIRE_PDF_RASTERIZER = $previousRequireRasterizer
}

$commit = (& git -C $repoRoot rev-parse --short HEAD).Trim()
$statusLines = @(& git -C $repoRoot status --short)
$status = ($statusLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine
$generatedAt = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
$pdfFiles = Get-ChildItem -LiteralPath $resolvedOutputPath -File -Filter '*.pdf' | Sort-Object Name
if ($pdfFiles.Count -eq 0) {
    throw "No PDF files were generated in $resolvedOutputPath. Check the dotnet test filter and OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT wiring."
}

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('# OfficeIMO PDF Visual Review Gallery')
$lines.Add('')
$lines.Add("Generated: $generatedAt")
$lines.Add('')
$lines.Add("Commit: ``$commit``")
$lines.Add('')
$lines.Add("Output directory: ``$resolvedOutputPath``")
$lines.Add('')
$lines.Add('Command:')
$lines.Add('')
$lines.Add('```powershell')
$lines.Add("Build/Export-PdfVisualReviewGallery.ps1 -OutputDirectory `"$OutputDirectory`" -Configuration `"$Configuration`" -Framework `"$Framework`"")
$lines.Add('```')
$lines.Add('')
$lines.Add('## Review Checklist')
$lines.Add('')
$lines.Add('- Open the PDFs in Edge and Acrobat/Reader when possible.')
$lines.Add('- Check text spacing, missing glyphs, image aspect ratio, table bounds, link targets, and header/footer placement.')
$lines.Add('- Compare dense Office exports against the source intent: Word layout, Excel pagination, PowerPoint slide placement, Markdown theme rhythm, and shared PDF authoring primitives.')
$lines.Add('- Treat visual differences as product evidence, not just test noise.')
$lines.Add('')
$lines.Add('## Files')
$lines.Add('')

$lines.Add('| File | Size |')
$lines.Add('| --- | ---: |')
foreach ($file in $pdfFiles) {
    $relativeName = $file.Name.Replace('|', '\|')
    $lines.Add("| [$relativeName]($relativeName) | $($file.Length) bytes |")
}

if (-not [string]::IsNullOrWhiteSpace($status)) {
    $lines.Add('')
    $lines.Add('## Working Tree Note')
    $lines.Add('')
    $lines.Add('The repository had uncommitted changes when this gallery was generated:')
    $lines.Add('')
    $lines.Add('```text')
    $lines.Add($status)
    $lines.Add('```')
}

[System.IO.File]::WriteAllLines($indexPath, $lines, [System.Text.Encoding]::UTF8)
Write-Host "PDF visual review gallery written to $resolvedOutputPath"
Write-Host "Index: $indexPath"
