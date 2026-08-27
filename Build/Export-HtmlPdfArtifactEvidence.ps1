param(
    [Parameter(Mandatory)]
    [string] $EvidencePath,

    [Parameter(Mandatory)]
    [ValidateSet('windows', 'linux')]
    [string] $Platform,

    [Parameter(Mandatory)]
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$resolvedEvidencePath = (Resolve-Path -LiteralPath $EvidencePath).Path
$reportPath = if (Test-Path -LiteralPath $resolvedEvidencePath -PathType Container) {
    Join-Path $resolvedEvidencePath 'html-pdf-evidence.json'
} else {
    $resolvedEvidencePath
}
$evidenceRoot = Split-Path -Parent $reportPath
$pathComparison = if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
    [System.StringComparison]::OrdinalIgnoreCase
} else {
    [System.StringComparison]::Ordinal
}
$report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
$pdfAssemblyPath = Join-Path $PSScriptRoot '../OfficeIMO.Pdf/bin/Release/net10.0/OfficeIMO.Pdf.dll'
$coreAssemblyPath = Join-Path $PSScriptRoot '../OfficeIMO.Pdf/bin/Release/net10.0/OfficeIMO.Core.dll'
if (-not (Test-Path -LiteralPath $pdfAssemblyPath -PathType Leaf) -or
    -not (Test-Path -LiteralPath $coreAssemblyPath -PathType Leaf)) {
    throw 'HTML/PDF artifact evidence requires a Release net10.0 OfficeIMO.Pdf build for independent PDF and preview inspection.'
}
Add-Type -Path (Resolve-Path -LiteralPath $coreAssemblyPath).Path
Add-Type -Path (Resolve-Path -LiteralPath $pdfAssemblyPath).Path

if ($report.schemaVersion -ne 2 -or
    [string] $report.scale -ne 'High' -or
    [int] $report.iterations -lt 3) {
    throw 'HTML/PDF artifact evidence must be a schema-v2 High-scale run with at least three iterations.'
}

$expectedPageCount = [int] $report.input.expectedPageCount
$expectedReportMarkerCount = [int] $report.input.expectedReportMarkerCount
if ($expectedPageCount -lt 1 -or $expectedReportMarkerCount -lt 1) {
    throw 'HTML/PDF artifact evidence input must declare positive expected page and report-marker counts.'
}

function Assert-SourceProvenance {
    param(
        [Parameter(Mandatory)][string] $Subject,
        [Parameter(Mandatory)][object] $Provenance,
        [switch] $RequireSource
    )

    $kind = [string] $Provenance.kind
    if ($RequireSource -and $kind -ne 'source') {
        throw "HTML/PDF artifact evidence requires source provenance for $Subject."
    }
    if ($kind -eq 'source' -and
        ([string] $Provenance.commit -notmatch '^[0-9a-fA-F]{40}$' -or
         $Provenance.worktreeClean -ne $true)) {
        throw "HTML/PDF artifact evidence requires a clean 40-character source commit for $Subject."
    }
}

if ($null -eq $report.provenance -or $null -eq $report.provenance.officeIMO -or
    $null -eq $report.provenance.htmlTinkerX) {
    throw 'HTML/PDF artifact evidence is missing producer provenance.'
}
Assert-SourceProvenance -Subject 'OfficeIMO' -Provenance $report.provenance.officeIMO -RequireSource
Assert-SourceProvenance -Subject 'HtmlTinkerX' -Provenance $report.provenance.htmlTinkerX

$expectedOsFamily = $Platform -eq 'windows' ? 'Windows' : 'Linux'
$reportedOsFamily = [string] $report.environment.osFamily
$osDescription = [string] $report.environment.osDescription
$actualOsFamily = if ($reportedOsFamily -in @('Windows', 'Linux', 'macOS')) {
    $reportedOsFamily
} elseif ($osDescription -match '(?i)windows') {
    'Windows'
} elseif ($osDescription -match '(?i)darwin|mac\s*os|osx') {
    'macOS'
} elseif ($osDescription -match '(?i)\blinux\b|\bubuntu\b|\bdebian\b|\balpine\b|\bfedora\b|\brhel\b|\bcentos\b|\barch\b|\bsuse\b') {
    'Linux'
} else {
    'Unknown'
}
if ($actualOsFamily -ne $expectedOsFamily -or
    [string]::IsNullOrWhiteSpace($osDescription) -or
    [string]::IsNullOrWhiteSpace([string] $report.environment.externalRasterizer)) {
    throw "HTML/PDF artifact evidence is not a $expectedOsFamily run with an external rasterizer."
}

$expectedEngines = @('Chromium', 'ITextPdfHtml', 'OfficeIMO', 'PeachPDF')
$expectedCancellation = @{
    Chromium = @{ Supports = $true; Status = 'Passed' }
    OfficeIMO = @{ Supports = $true; Status = 'Passed' }
    ITextPdfHtml = @{ Supports = $false; Status = 'Unsupported' }
    PeachPDF = @{ Supports = $false; Status = 'Unsupported' }
}
$engines = @($report.engines)
$actualEngines = @($engines | ForEach-Object { [string] $_.engine } | Sort-Object)
if ($engines.Count -ne $expectedEngines.Count -or
    @(Compare-Object -ReferenceObject $expectedEngines -DifferenceObject $actualEngines).Count -ne 0) {
    throw 'HTML/PDF artifact evidence must contain exactly the required Chromium, ITextPdfHtml, OfficeIMO, and PeachPDF engines.'
}

foreach ($engine in $engines) {
    $engineName = [string] $engine.engine
    $outputs = @($engine.outputs)
    $iterations = @($outputs | ForEach-Object { [int] $_.iteration } | Sort-Object -Unique)
    $expectedIterations = @(1..([int] $report.iterations))
    if ($outputs.Count -ne [int] $report.iterations -or
        $iterations.Count -ne $expectedIterations.Count -or
        @(Compare-Object -ReferenceObject $expectedIterations -DifferenceObject $iterations).Count -ne 0) {
        throw "HTML/PDF artifact evidence engine '$engineName' must contain exactly one output for every declared iteration."
    }

    $cancellation = $expectedCancellation[$engineName]
    $requiresInFlightCancellation = $cancellation.Supports -eq $true
    if ($engine.cancellation.apiSupportsCancellation -ne $cancellation.Supports -or
        [string] $engine.cancellation.status -ne $cancellation.Status -or
        ($requiresInFlightCancellation -and
            [string] $engine.cancellation.detail -notmatch '(?i)\bin-flight\b.*\bcancelled in \d+(?:\.\d+)? ms\b') -or
        $engine.memoryComparable -ne $true) {
        throw "HTML/PDF artifact evidence engine '$engineName' does not satisfy the cancellation, memory, and determinism contract."
    }

    foreach ($output in $outputs) {
        $memory = $output.processTreeMemory
        if ($null -eq $memory -or
            [int] $memory.sampleCount -lt 1 -or
            [int] $memory.minimumObservedProcessCount -lt 1 -or
            ($engineName -eq 'Chromium' -and [int] $memory.maximumObservedProcessCount -le 1)) {
            throw "HTML/PDF artifact evidence engine '$engineName' contains output without comparable process-tree memory evidence."
        }
        $contract = $output.contract
        if ($null -eq $contract -or
            [int] $contract.pageCount -ne $expectedPageCount -or
            [int] $contract.textLength -lt 1 -or
            [int] $contract.reportMarkerCount -ne $expectedReportMarkerCount -or
            $contract.tagged -ne $true -or
            $contract.marked -ne $true -or
            [string]::IsNullOrWhiteSpace([string] $contract.catalogLanguage) -or
            [int] $contract.structureElementCount -lt 1 -or
            [int] $contract.markedContentReferenceCount -lt 1 -or
            [int] $contract.parentTreeEntryCount -lt 1 -or
            $contract.hasDocumentStructureElement -ne $true -or
            $contract.figuresHaveAlternateText -ne $true -or
            [string] $output.semanticSha256 -notmatch '^[0-9a-f]{64}$') {
            throw "HTML/PDF artifact evidence engine '$engineName' contains an output without the required semantic and tagged-PDF contract."
        }
    }
}

function Assert-PdfArtifactContract {
    param(
        [Parameter(Mandatory)][string] $RelativePath,
        [Parameter(Mandatory)][int] $ExpectedPageCount,
        [Parameter(Mandatory)][int] $ExpectedReportMarkerCount,
        [Parameter(Mandatory)][object] $ExpectedContract
    )

    $pdfInfo = @(Get-Command pdfinfo -CommandType Application -ErrorAction SilentlyContinue) | Select-Object -First 1
    $pdfToText = @(Get-Command pdftotext -CommandType Application -ErrorAction SilentlyContinue) | Select-Object -First 1
    if ($null -eq $pdfInfo -or $null -eq $pdfToText) {
        throw 'HTML/PDF artifact evidence requires executable pdfinfo and pdftotext validators.'
    }

    $fullPath = [System.IO.Path]::GetFullPath((Join-Path $evidenceRoot $RelativePath))
    $pdfInfoOutput = @(& $pdfInfo.Source $fullPath 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw "Executable PDF validation failed for artifact: $RelativePath"
    }
    $pageMatch = [regex]::Match(
        ($pdfInfoOutput -join "`n"),
        '(?m)^Pages:\s+(?<count>\d+)\s*$')
    if (-not $pageMatch.Success -or
        [int] $pageMatch.Groups['count'].Value -ne $ExpectedPageCount) {
        throw "Executable PDF validation found an unexpected page count for artifact: $RelativePath"
    }

    $pdfBytes = [System.IO.File]::ReadAllBytes($fullPath)
    $documentInfo = [OfficeIMO.Pdf.PdfDocument]::Open($pdfBytes).Inspect()
    $tagged = $documentInfo.TaggedContent
    $actualTypeCounts = if ($null -eq $tagged) {
        [string]::Empty
    } else {
        (@($tagged.StructureTypeCounts.GetEnumerator() |
                Sort-Object Key |
                ForEach-Object { "$($_.Key):$($_.Value)" }) -join ',')
    }
    $expectedTypeCounts = if ($null -eq $ExpectedContract.structureTypeCounts) {
        [string]::Empty
    } else {
        (@($ExpectedContract.structureTypeCounts.PSObject.Properties |
                Sort-Object Name |
                ForEach-Object { "$($_.Name):$($_.Value)" }) -join ',')
    }
    if ($documentInfo.HasTaggedContent -ne [bool] $ExpectedContract.tagged -or
        $null -eq $tagged -or
        $tagged.Marked -ne [bool] $ExpectedContract.marked -or
        -not [string]::Equals(
            [string] $documentInfo.CatalogLanguage,
            [string] $ExpectedContract.catalogLanguage,
            [System.StringComparison]::OrdinalIgnoreCase) -or
        $tagged.StructureElementCount -ne [int] $ExpectedContract.structureElementCount -or
        $tagged.MarkedContentReferenceCount -ne [int] $ExpectedContract.markedContentReferenceCount -or
        $tagged.ParentTreeEntryCount -ne [int] $ExpectedContract.parentTreeEntryCount -or
        $tagged.HasDocumentStructureElement -ne [bool] $ExpectedContract.hasDocumentStructureElement -or
        $tagged.FiguresHaveAlternateText -ne [bool] $ExpectedContract.figuresHaveAlternateText -or
        $actualTypeCounts -ne $expectedTypeCounts) {
        throw "Independent PDF inspection contradicted the tagged-PDF contract for artifact: $RelativePath"
    }

    $textPath = Join-Path ([System.IO.Path]::GetTempPath()) (
        'officeimo-html-pdf-evidence-' + [Guid]::NewGuid().ToString('N') + '.txt')
    try {
        $null = & $pdfToText.Source -enc UTF-8 $fullPath $textPath 2>&1
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $textPath -PathType Leaf)) {
            throw "Executable PDF text validation failed for artifact: $RelativePath"
        }
        $text = Get-Content -LiteralPath $textPath -Raw
        $normalizedText = [regex]::Replace(
            $text.ToUpperInvariant(),
            '[^\p{L}\p{Nd}]',
            [string]::Empty,
            [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)
        $actualMarkerCount = [regex]::Matches(
            $normalizedText,
            [regex]::Escape('BENCHMARKREPORT'),
            [System.Text.RegularExpressions.RegexOptions]::CultureInvariant).Count
        if ($actualMarkerCount -ne $ExpectedReportMarkerCount) {
            throw "Executable PDF validation found an unexpected report-marker count for artifact: $RelativePath"
        }

        $semanticText = [regex]::Replace(
            $text.Normalize([System.Text.NormalizationForm]::FormC).Trim(),
            '\s+',
            ' ',
            [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)
        return [Convert]::ToHexString(
            [System.Security.Cryptography.SHA256]::HashData(
                [System.Text.Encoding]::UTF8.GetBytes($semanticText))).ToLowerInvariant()
    } finally {
        if (Test-Path -LiteralPath $textPath -PathType Leaf) {
            Remove-Item -LiteralPath $textPath -Force
        }
    }
}

function Assert-PngPreview {
    param(
        [Parameter(Mandatory)][string] $RelativePath,
        [Parameter(Mandatory)][int] $ExpectedWidth,
        [Parameter(Mandatory)][int] $ExpectedHeight
    )
    $fullPath = [System.IO.Path]::GetFullPath((Join-Path $evidenceRoot $RelativePath))
    $bytes = [System.IO.File]::ReadAllBytes($fullPath)
    $image = $null
    if (-not [OfficeIMO.Drawing.OfficePngReader]::TryDecode($bytes, [ref] $image) -or
        $null -eq $image -or
        $image.Width -ne $ExpectedWidth -or
        $image.Height -ne $ExpectedHeight) {
        throw "Visual evidence is not a decodable PNG with the declared dimensions: $RelativePath"
    }
}

function Assert-ManagedPreviewMatchesPdf {
    param(
        [Parameter(Mandatory)][string] $PdfRelativePath,
        [Parameter(Mandatory)][object] $Preview
    )
    $pdfPath = [System.IO.Path]::GetFullPath((Join-Path $evidenceRoot $PdfRelativePath))
    $options = [OfficeIMO.Pdf.PdfPageRenderOptions]::new()
    $options.Format = [OfficeIMO.Pdf.PdfPageRenderFormat]::Png
    $options.Dpi = 120D
    $options.ContinueOnError = $false
    $options.MaxPages = 1
    $rendered = @([OfficeIMO.Pdf.PdfDocument]::Open([System.IO.File]::ReadAllBytes($pdfPath)).Read.RenderPages('1', $options))
    if ($rendered.Count -ne 1 -or $null -eq $rendered[0].Bytes) {
        throw "Independent managed rendering failed for artifact: $PdfRelativePath"
    }
    $actualHash = [Convert]::ToHexString(
        [System.Security.Cryptography.SHA256]::HashData($rendered[0].Bytes)).ToLowerInvariant()
    if ($rendered[0].Width -ne [int] $Preview.width -or
        $rendered[0].Height -ne [int] $Preview.height -or
        -not [string]::Equals($actualHash, [string] $Preview.sha256, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Managed preview does not match an independent page-one rendering of artifact: $PdfRelativePath"
    }
}

function Assert-ExternalPreviewMatchesPdf {
    param(
        [Parameter(Mandatory)][string] $PdfRelativePath,
        [Parameter(Mandatory)][object] $Preview
    )
    $pdfToPpm = @(Get-Command pdftoppm -CommandType Application -ErrorAction SilentlyContinue) | Select-Object -First 1
    if ($null -eq $pdfToPpm) {
        throw 'HTML/PDF artifact evidence requires executable pdftoppm preview validation.'
    }
    $pdfPath = [System.IO.Path]::GetFullPath((Join-Path $evidenceRoot $PdfRelativePath))
    $temporaryPrefix = Join-Path ([System.IO.Path]::GetTempPath()) (
        'officeimo-html-pdf-preview-' + [Guid]::NewGuid().ToString('N'))
    $temporaryPng = $temporaryPrefix + '.png'
    try {
        $null = & $pdfToPpm.Source -f 1 -l 1 -singlefile -png -r 120 $pdfPath $temporaryPrefix 2>&1
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $temporaryPng -PathType Leaf)) {
            throw "Independent external rendering failed for artifact: $PdfRelativePath"
        }
        $actualHash = (Get-FileHash -LiteralPath $temporaryPng -Algorithm SHA256).Hash.ToLowerInvariant()
        if (-not [string]::Equals($actualHash, [string] $Preview.sha256, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "External preview does not match an independent page-one rendering of artifact: $PdfRelativePath"
        }
    } finally {
        if (Test-Path -LiteralPath $temporaryPng -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPng -Force
        }
    }
}

$artifacts = [System.Collections.Generic.List[object]]::new()
$pathComparer = if ($pathComparison -eq [System.StringComparison]::OrdinalIgnoreCase) {
    [System.StringComparer]::OrdinalIgnoreCase
} else {
    [System.StringComparer]::Ordinal
}
$validatedArtifactPaths = [System.Collections.Generic.HashSet[string]]::new($pathComparer)
function Assert-NoArtifactPathLinks {
    param(
        [Parameter(Mandatory)][string] $RootPath,
        [Parameter(Mandatory)][string] $FullPath,
        [Parameter(Mandatory)][string] $RelativePath
    )

    $rootPathFull = [System.IO.Path]::GetFullPath($RootPath).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar)
    $relativeFromRoot = [System.IO.Path]::GetRelativePath($rootPathFull, $FullPath)
    $currentPath = $rootPathFull
    $pathsToInspect = [System.Collections.Generic.List[string]]::new()
    $pathsToInspect.Add($currentPath)
    foreach ($segment in $relativeFromRoot.Split(
            [char[]]@(
                [System.IO.Path]::DirectorySeparatorChar,
                [System.IO.Path]::AltDirectorySeparatorChar),
            [System.StringSplitOptions]::RemoveEmptyEntries)) {
        $currentPath = Join-Path $currentPath $segment
        $pathsToInspect.Add($currentPath)
    }

    foreach ($path in $pathsToInspect) {
        $item = Get-Item -LiteralPath $path -Force
        $isReparsePoint = ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
        $isLink = $item.PSObject.Properties.Name -contains 'LinkType' -and
            -not [string]::IsNullOrWhiteSpace([string] $item.LinkType)
        if ($isReparsePoint -or $isLink) {
            throw "Artifact path contains a symbolic link or reparse point: $RelativePath"
        }
    }
}

function Add-ValidatedArtifact {
    param(
        [Parameter(Mandatory)][string] $Kind,
        [Parameter(Mandatory)][string] $RelativePath,
        [Parameter(Mandatory)][long] $ExpectedSize,
        [Parameter(Mandatory)][string] $ExpectedSha256
    )

    $fullPath = [System.IO.Path]::GetFullPath((Join-Path $evidenceRoot $RelativePath))
    $rootPrefix = [System.IO.Path]::GetFullPath($evidenceRoot).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    if (-not $fullPath.StartsWith($rootPrefix, $pathComparison)) {
        throw "Artifact path escapes the evidence root: $RelativePath"
    }
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        throw "Artifact is missing: $RelativePath"
    }
    Assert-NoArtifactPathLinks `
        -RootPath $evidenceRoot `
        -FullPath $fullPath `
        -RelativePath $RelativePath

    $item = Get-Item -LiteralPath $fullPath
    $actualHash = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($item.Length -ne $ExpectedSize -or
        -not [string]::Equals($actualHash, $ExpectedSha256, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Artifact size or SHA-256 does not match the evidence report: $RelativePath"
    }
    $validatedArtifactPaths.Add($fullPath) | Out-Null

    $artifacts.Add([ordered]@{
            kind = $Kind
            relativePath = $RelativePath.Replace('\', '/')
            sizeBytes = $item.Length
            sha256 = $actualHash
        }) | Out-Null
    return $actualHash
}

$null = Add-ValidatedArtifact `
    -Kind 'input' `
    -RelativePath ([string] $report.input.relativePath) `
    -ExpectedSize ([long] $report.input.sizeBytes) `
    -ExpectedSha256 ([string] $report.input.sha256)

foreach ($engine in $engines) {
    $byteHashes = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $semanticHashes = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $managedVisualHashes = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $externalVisualHashes = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($output in @($engine.outputs)) {
        $byteHashes.Add((Add-ValidatedArtifact `
            -Kind ("pdf:$($engine.engine)") `
            -RelativePath ([string] $output.relativePath) `
            -ExpectedSize ([long] $output.sizeBytes) `
            -ExpectedSha256 ([string] $output.sha256))) | Out-Null
        $semanticHashes.Add((Assert-PdfArtifactContract `
            -RelativePath ([string] $output.relativePath) `
            -ExpectedPageCount ([int] $output.contract.pageCount) `
            -ExpectedReportMarkerCount ([int] $output.contract.reportMarkerCount) `
            -ExpectedContract $output.contract)) | Out-Null
        $managedVisualHashes.Add((Add-ValidatedArtifact `
            -Kind ("managed-preview:$($engine.engine)") `
            -RelativePath ([string] $output.managedVisual.relativePath) `
            -ExpectedSize ([long] $output.managedVisual.sizeBytes) `
            -ExpectedSha256 ([string] $output.managedVisual.sha256))) | Out-Null
        Assert-PngPreview `
            -RelativePath ([string] $output.managedVisual.relativePath) `
            -ExpectedWidth ([int] $output.managedVisual.width) `
            -ExpectedHeight ([int] $output.managedVisual.height)
        Assert-ManagedPreviewMatchesPdf `
            -PdfRelativePath ([string] $output.relativePath) `
            -Preview $output.managedVisual
        if ($null -eq $output.externalVisual) {
            throw "External visual evidence is missing for $($engine.engine) iteration $($output.iteration)."
        }
        $externalVisualHashes.Add((Add-ValidatedArtifact `
            -Kind ("external-preview:$($engine.engine)") `
            -RelativePath ([string] $output.externalVisual.relativePath) `
            -ExpectedSize ([long] $output.externalVisual.sizeBytes) `
            -ExpectedSha256 ([string] $output.externalVisual.sha256))) | Out-Null
        Assert-PngPreview `
            -RelativePath ([string] $output.externalVisual.relativePath) `
            -ExpectedWidth ([int] $output.externalVisual.width) `
            -ExpectedHeight ([int] $output.externalVisual.height)
        Assert-ExternalPreviewMatchesPdf `
            -PdfRelativePath ([string] $output.relativePath) `
            -Preview $output.externalVisual
    }

    $determinism = $engine.determinism
    $actualExact = $byteHashes.Count -eq 1
    $actualSemantic = $semanticHashes.Count -eq 1
    $actualManaged = $managedVisualHashes.Count -eq 1
    $actualExternal = $externalVisualHashes.Count -eq 1
    if ($null -eq $determinism -or
        $determinism.exactBytesIdentical -ne $actualExact -or
        $determinism.semanticOutputIdentical -ne $actualSemantic -or
        $determinism.managedVisualPreviewIdentical -ne $actualManaged -or
        $determinism.externalVisualPreviewIdentical -ne $actualExternal -or
        [int] $determinism.uniqueByteHashCount -ne $byteHashes.Count -or
        [int] $determinism.uniqueSemanticHashCount -ne $semanticHashes.Count -or
        [int] $determinism.uniqueManagedVisualHashCount -ne $managedVisualHashes.Count -or
        [int] $determinism.uniqueExternalVisualHashCount -ne $externalVisualHashes.Count -or
        ([string] $engine.engine -eq 'OfficeIMO' -and -not $actualExact) -or
        -not $actualSemantic -or -not $actualManaged -or -not $actualExternal) {
        throw "HTML/PDF artifact evidence engine '$($engine.engine)' does not satisfy the determinism contract derived from its validated hashes."
    }
}

$orderedArtifacts = @($artifacts | Sort-Object { $_.relativePath })
if (@($orderedArtifacts | Group-Object { $_.relativePath } | Where-Object { $_.Count -ne 1 }).Count -ne 0) {
    throw 'HTML/PDF artifact evidence contains duplicate artifact paths.'
}
$manifestText = ($orderedArtifacts | ForEach-Object {
        "$($_.kind)|$($_.relativePath)|$($_.sizeBytes)|$($_.sha256)"
    }) -join "`n"
$manifestBytes = [System.Text.Encoding]::UTF8.GetBytes($manifestText)
$manifestHash = [Convert]::ToHexString(
    [System.Security.Cryptography.SHA256]::HashData($manifestBytes)).ToLowerInvariant()

$summary = [ordered]@{
    schemaVersion = 1
    format = 'officeimo.html-pdf-artifact-evidence-summary'
    platform = $Platform
    artifactBundle = [ordered]@{
        artifactCount = $orderedArtifacts.Count
        totalBytes = [long] (@($orderedArtifacts | ForEach-Object { $_.sizeBytes } | Measure-Object -Sum).Sum)
        manifestSha256 = $manifestHash
        artifacts = $orderedArtifacts
    }
    report = $report
}

$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
if ([string]::Equals($resolvedOutputPath, [System.IO.Path]::GetFullPath($reportPath), $pathComparison) -or
    $validatedArtifactPaths.Contains($resolvedOutputPath)) {
    throw 'The artifact evidence summary output cannot overwrite the report or a validated input artifact.'
}
New-Item -ItemType Directory -Path (Split-Path -Parent $resolvedOutputPath) -Force | Out-Null
$json = ($summary | ConvertTo-Json -Depth 100).Replace("`r`n", "`n") + "`n"
[System.IO.File]::WriteAllText($resolvedOutputPath, $json, [System.Text.UTF8Encoding]::new($false))
Write-Host "Validated HTML/PDF artifact evidence summary written to '$resolvedOutputPath'."
