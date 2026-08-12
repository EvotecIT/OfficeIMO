[CmdletBinding()]
param(
    [string] $OutputDirectory,
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0',
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $repositoryRoot 'Ignore\Benchmarks\PdfComparisons\word-com'
} elseif (-not [System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory = Join-Path $repositoryRoot $OutputDirectory
}
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
$fixtureDirectory = Join-Path $OutputDirectory 'files'
$projectPath = Join-Path $repositoryRoot 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
$docxPath = Join-Path $fixtureDirectory 'officeimo-word-rich-25-page.docx'
$comPdfPath = Join-Path $fixtureDirectory 'microsoft-word-com-rich.pdf'

New-Item -ItemType Directory -Path $fixtureDirectory -Force | Out-Null

& dotnet run --project $projectPath -c $Configuration -f $Framework -- prepare-rich-word --repo-root $repositoryRoot --output $fixtureDirectory
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $docxPath -PathType Leaf)) {
    throw "OfficeIMO.Word fixture generation failed: $docxPath"
}

$word = $null
$document = $null
$anchorRange = $null
$layout = $null
$shape = $null
$smartArt = $null
$nodes = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($docxPath, $false, $false)

    $anchorRange = $document.Content
    $anchorRange.Collapse(0)
    $anchorRange.InsertBreak(7)
    $anchorRange.Collapse(0)
    $anchorRange.InsertAfter("COM SMARTART INTEROPERABILITY PAGE`r")
    $anchorRange.Collapse(0)

    $layout = $word.SmartArtLayouts.Item(1)
    $shape = $document.Shapes.AddSmartArt($layout, 50, 100, 450, 220, $anchorRange)
    $smartArt = $shape.SmartArt
    $nodes = $smartArt.AllNodes
    $labels = @('Collect', 'Validate', 'Publish')
    for ($index = 1; $index -le $labels.Count; $index++) {
        $node = $null
        $textFrame = $null
        $textRange = $null
        try {
            $node = $nodes.Item($index)
            $textFrame = $node.TextFrame2
            $textRange = $textFrame.TextRange
            $textRange.Text = $labels[$index - 1]
        } finally {
            if ($null -ne $textRange) {
                [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($textRange)
            }
            if ($null -ne $textFrame) {
                [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($textFrame)
            }
            if ($null -ne $node) {
                [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($node)
            }
        }
    }

    $document.ExportAsFixedFormat($comPdfPath, 17)
} finally {
    if ($null -ne $nodes) {
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($nodes)
    }
    if ($null -ne $smartArt) {
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($smartArt)
    }
    if ($null -ne $shape) {
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($shape)
    }
    if ($null -ne $layout) {
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($layout)
    }
    if ($null -ne $anchorRange) {
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($anchorRange)
    }
    if ($null -ne $document) {
        $document.Close(0)
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($null -ne $word) {
        $word.Quit()
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (-not (Test-Path -LiteralPath $comPdfPath -PathType Leaf)) {
    throw "Microsoft Word did not create the expected PDF: $comPdfPath"
}

$validationDirectory = Join-Path $OutputDirectory 'validation'
& dotnet run --project $projectPath -c $Configuration -f $Framework -- corpus --repo-root $repositoryRoot --only microsoft-word-com-rich --com-pdf $comPdfPath --skip-manipulation --output $validationDirectory
if ($LASTEXITCODE -ne 0) {
    throw "Microsoft Word COM PDF corpus validation failed with exit code $LASTEXITCODE."
}

Write-Host "COM PDF corpus validation passed: $validationDirectory"
