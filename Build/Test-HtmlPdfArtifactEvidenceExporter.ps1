$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$isWindowsHost = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT
$platform = $isWindowsHost ? 'windows' : 'linux'
$osDescription = $isWindowsHost ? 'Windows contract-test host' : 'Linux contract-test host'
$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) (
    'officeimo-html-pdf-artifact-exporter-' + [Guid]::NewGuid().ToString('N'))

try {
    $evidenceRoot = Join-Path $temporaryRoot 'Run'
    New-Item -ItemType Directory -Path $evidenceRoot -Force | Out-Null
    $insidePath = Join-Path $evidenceRoot 'input.html'
    [System.IO.File]::WriteAllText($insidePath, 'inside', [System.Text.UTF8Encoding]::new($false))
    $insideItem = Get-Item -LiteralPath $insidePath
    $insideHash = (Get-FileHash -LiteralPath $insidePath -Algorithm SHA256).Hash.ToLowerInvariant()
    $reportPath = Join-Path $evidenceRoot 'html-pdf-evidence.json'
    $outputPath = Join-Path $temporaryRoot 'summary.json'
    $engines = @(
        foreach ($engineName in @('OfficeIMO', 'PeachPDF', 'ITextPdfHtml', 'Chromium')) {
            $outputs = @(
                foreach ($iteration in 1..3) {
                    $slug = $engineName.ToLowerInvariant()
                    $pdfRelativePath = "$slug-$iteration.pdf"
                    $managedRelativePath = "$slug-$iteration-managed.png"
                    $externalRelativePath = "$slug-$iteration-external.png"
                    foreach ($artifact in @(
                            [pscustomobject]@{ Path = $pdfRelativePath; Content = "pdf-$engineName-$iteration" },
                            [pscustomobject]@{ Path = $managedRelativePath; Content = "managed-$engineName-$iteration" },
                            [pscustomobject]@{ Path = $externalRelativePath; Content = "external-$engineName-$iteration" })) {
                        $artifactPath = Join-Path $evidenceRoot $artifact.Path
                        [System.IO.File]::WriteAllText(
                            $artifactPath,
                            $artifact.Content,
                            [System.Text.UTF8Encoding]::new($false))
                    }

                    $pdfItem = Get-Item -LiteralPath (Join-Path $evidenceRoot $pdfRelativePath)
                    $managedItem = Get-Item -LiteralPath (Join-Path $evidenceRoot $managedRelativePath)
                    $externalItem = Get-Item -LiteralPath (Join-Path $evidenceRoot $externalRelativePath)
                    [ordered]@{
                        iteration = $iteration
                        relativePath = $pdfRelativePath
                        sizeBytes = $pdfItem.Length
                        sha256 = (Get-FileHash -LiteralPath $pdfItem.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
                        semanticSha256 = ('a' * 64)
                        contract = [ordered]@{
                            pageCount = 1
                            textLength = 10
                            reportMarkerCount = 1
                            characterChecksum = 1
                            tagged = $true
                            marked = $true
                            catalogLanguage = 'en-US'
                            structureElementCount = 2
                            markedContentReferenceCount = 1
                            parentTreeEntryCount = 1
                            hasDocumentStructureElement = $true
                            figuresHaveAlternateText = $true
                        }
                        managedVisual = [ordered]@{
                            relativePath = $managedRelativePath
                            sizeBytes = $managedItem.Length
                            sha256 = (Get-FileHash -LiteralPath $managedItem.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
                        }
                        externalVisual = [ordered]@{
                            relativePath = $externalRelativePath
                            sizeBytes = $externalItem.Length
                            sha256 = (Get-FileHash -LiteralPath $externalItem.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
                        }
                    }
                }
            )
            [ordered]@{
                engine = $engineName
                cancellation = [ordered]@{
                    status = $engineName -in @('OfficeIMO', 'Chromium') ? 'Passed' : 'Unsupported'
                }
                determinism = [ordered]@{
                    exactBytesIdentical = $true
                    semanticOutputIdentical = $true
                    managedVisualPreviewIdentical = $true
                    externalVisualPreviewIdentical = $true
                }
                memoryComparable = $true
                outputs = $outputs
            }
        }
    )
    $report = [ordered]@{
        schemaVersion = 2
        scale = 'High'
        iterations = 3
        environment = [ordered]@{
            osDescription = $osDescription
            externalRasterizer = 'contract-test'
        }
        input = [ordered]@{
            relativePath = 'input.html'
            sizeBytes = $insideItem.Length
            sha256 = $insideHash
        }
        engines = $engines
    }

    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))

    $validEngines = $report.engines
    $report.engines = @()
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $missingEnginesRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'exactly the required') { throw }
        $missingEnginesRejected = $true
    }
    if (-not $missingEnginesRejected) {
        throw 'HTML/PDF artifact exporter accepted evidence without the required engines.'
    }

    $report.engines = $validEngines
    $validOutputs = $report.engines[0].outputs
    $report.engines[0].outputs = @($validOutputs | Select-Object -First 2)
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $missingOutputRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'exactly one output') { throw }
        $missingOutputRejected = $true
    }
    if (-not $missingOutputRejected) {
        throw 'HTML/PDF artifact exporter accepted incomplete engine output.'
    }

    $report.engines[0].outputs = $validOutputs
    $validOsDescription = $report.environment.osDescription
    $report.environment.osDescription = 'macOS contract-test host'
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $unsupportedOsRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform linux `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'not a Linux run') { throw }
        $unsupportedOsRejected = $true
    }
    if (-not $unsupportedOsRejected) {
        throw 'HTML/PDF artifact exporter mislabeled non-Linux evidence as Linux.'
    }

    $report.environment.osDescription = $validOsDescription
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
        -EvidencePath $evidenceRoot `
        -Platform $platform `
        -OutputPath $outputPath
    if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
        throw 'HTML/PDF artifact exporter did not produce the valid in-root summary.'
    }

    $siblingName = $isWindowsHost ? 'Escape' : 'run'
    $siblingRoot = Join-Path $temporaryRoot $siblingName
    New-Item -ItemType Directory -Path $siblingRoot -Force | Out-Null
    $outsidePath = Join-Path $siblingRoot 'outside.html'
    [System.IO.File]::WriteAllText($outsidePath, 'outside', [System.Text.UTF8Encoding]::new($false))
    $outsideItem = Get-Item -LiteralPath $outsidePath
    $outsideHash = (Get-FileHash -LiteralPath $outsidePath -Algorithm SHA256).Hash.ToLowerInvariant()
    $report.input.relativePath = "../$siblingName/outside.html"
    $report.input.sizeBytes = $outsideItem.Length
    $report.input.sha256 = $outsideHash
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))

    $rejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'escapes the evidence root') { throw }
        $rejected = $true
    }
    if (-not $rejected) {
        throw 'HTML/PDF artifact exporter accepted a path outside the evidence root.'
    }

    $linkedArtifacts = Join-Path $evidenceRoot 'artifacts'
    $outsideArtifacts = Join-Path $temporaryRoot 'outside-artifacts'
    $linkPath = Join-Path $linkedArtifacts 'link'
    New-Item -ItemType Directory -Path $linkedArtifacts -Force | Out-Null
    New-Item -ItemType Directory -Path $outsideArtifacts -Force | Out-Null
    $linkedOutsidePath = Join-Path $outsideArtifacts 'linked.html'
    [System.IO.File]::WriteAllText($linkedOutsidePath, 'linked-outside', [System.Text.UTF8Encoding]::new($false))
    if ($isWindowsHost) {
        New-Item -ItemType Junction -Path $linkPath -Target $outsideArtifacts | Out-Null
    } else {
        New-Item -ItemType SymbolicLink -Path $linkPath -Target $outsideArtifacts | Out-Null
    }
    try {
        $linkedOutsideItem = Get-Item -LiteralPath $linkedOutsidePath
        $report.input.relativePath = 'artifacts/link/linked.html'
        $report.input.sizeBytes = $linkedOutsideItem.Length
        $report.input.sha256 = (Get-FileHash -LiteralPath $linkedOutsidePath -Algorithm SHA256).Hash.ToLowerInvariant()
        $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
        [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))

        $linkRejected = $false
        try {
            & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
                -EvidencePath $evidenceRoot `
                -Platform $platform `
                -OutputPath $outputPath
        } catch {
            if ($_.Exception.Message -notmatch 'symbolic link or reparse point') { throw }
            $linkRejected = $true
        }
        if (-not $linkRejected) {
            throw 'HTML/PDF artifact exporter accepted an artifact through a symbolic link or reparse point.'
        }
    } finally {
        if (Test-Path -LiteralPath $linkPath) {
            if ($isWindowsHost) {
                [System.IO.Directory]::Delete($linkPath)
            } else {
                Remove-Item -LiteralPath $linkPath -Force
            }
        }
    }
} finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}

Write-Host "HTML/PDF artifact exporter path boundary passed on $platform."
