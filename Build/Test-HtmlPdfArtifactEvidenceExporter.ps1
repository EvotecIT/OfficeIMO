$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$isWindowsHost = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT
$platform = $isWindowsHost ? 'windows' : 'linux'
$osDescription = $isWindowsHost ? 'Windows contract-test host' : 'Linux contract-test host'
$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) (
    'officeimo-html-pdf-artifact-exporter-' + [Guid]::NewGuid().ToString('N'))
$pdfProjectPath = Join-Path $PSScriptRoot '../OfficeIMO.Pdf/OfficeIMO.Pdf.csproj'
$pdfAssemblyPath = Join-Path $PSScriptRoot '../OfficeIMO.Pdf/bin/Release/net10.0/OfficeIMO.Pdf.dll'
if (-not (Test-Path -LiteralPath $pdfAssemblyPath -PathType Leaf)) {
    & dotnet build $pdfProjectPath -c Release -f net10.0 --nologo
    if ($LASTEXITCODE -ne 0) { throw 'Could not build OfficeIMO.Pdf for artifact-exporter contracts.' }
}
Add-Type -Path (Join-Path $PSScriptRoot '../OfficeIMO.Pdf/bin/Release/net10.0/OfficeIMO.Core.dll')
Add-Type -Path $pdfAssemblyPath

function New-TestPdfBytes {
    param(
        [Parameter(Mandatory)][string] $Text,
        [int] $Variant = 0
    )

    $content = "BT /F1 12 Tf 72 720 Td ($Text) Tj ET"
    $objects = @(
        '<< /Type /Catalog /Pages 2 0 R /Lang (en-US) /MarkInfo << /Marked true >> /StructTreeRoot 6 0 R >>',
        '<< /Type /Pages /Count 1 /Kids [3 0 R] >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /StructParents 0 /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
        "<< /Length $([System.Text.Encoding]::ASCII.GetByteCount("/P <</MCID 0>> BDC $content EMC")) >>`nstream`n/P <</MCID 0>> BDC $content EMC`nendstream",
        '<< /Type /StructTreeRoot /K [7 0 R] /ParentTree 8 0 R /ParentTreeNextKey 1 >>',
        '<< /Type /StructElem /S /Document /P 6 0 R /K [9 0 R] /Lang (en-US) >>',
        '<< /Nums [0 [9 0 R]] >>',
        '<< /Type /StructElem /S /P /P 7 0 R /Pg 3 0 R /K << /Type /MCR /Pg 3 0 R /MCID 0 >> >>'
    )
    $buffer = [System.Collections.Generic.List[byte]]::new()
    $offsets = [System.Collections.Generic.List[int]]::new()
    $appendAscii = {
        param([string] $Value)
        $buffer.AddRange([System.Text.Encoding]::ASCII.GetBytes($Value))
    }
    & $appendAscii "%PDF-1.4`n% variant $Variant`n"
    for ($index = 0; $index -lt $objects.Count; $index++) {
        $offsets.Add($buffer.Count)
        & $appendAscii "$($index + 1) 0 obj`n$($objects[$index])`nendobj`n"
    }
    $xrefOffset = $buffer.Count
    & $appendAscii "xref`n0 $($objects.Count + 1)`n0000000000 65535 f `n"
    foreach ($offset in $offsets) {
        & $appendAscii ($offset.ToString('0000000000') + " 00000 n `n")
    }
    & $appendAscii "trailer`n<< /Size $($objects.Count + 1) /Root 1 0 R >>`nstartxref`n$xrefOffset`n%%EOF`n"
    return $buffer.ToArray()
}

function Add-PngTextChunk {
    param([Parameter(Mandatory)][byte[]] $Bytes)

    if ($Bytes.Length -lt 20 -or
        [System.Text.Encoding]::ASCII.GetString($Bytes, $Bytes.Length - 8, 4) -ne 'IEND') {
        throw 'PNG fixture does not end with an IEND chunk.'
    }

    [byte[]] $type = [System.Text.Encoding]::ASCII.GetBytes('tEXt')
    [byte[]] $data = [System.Text.Encoding]::ASCII.GetBytes("Comment`0independent-preview")
    [uint64] $crc = 4294967295
    foreach ($value in [byte[]] ($type + $data)) {
        $crc = ($crc -bxor [uint64] $value) -band 4294967295
        foreach ($bit in 1..8) {
            $crc = if (($crc -band 1) -ne 0) {
                (($crc -shr 1) -bxor 3988292384) -band 4294967295
            } else {
                ($crc -shr 1) -band 4294967295
            }
        }
    }
    $crc = ($crc -bxor 4294967295) -band 4294967295

    $result = [System.Collections.Generic.List[byte]]::new($Bytes.Length + 12 + $data.Length)
    $iendOffset = $Bytes.Length - 12
    $result.AddRange([byte[]] $Bytes[0..($iendOffset - 1)])
    $result.AddRange([byte[]] @(
            [byte] (($data.Length -shr 24) -band 0xff),
            [byte] (($data.Length -shr 16) -band 0xff),
            [byte] (($data.Length -shr 8) -band 0xff),
            [byte] ($data.Length -band 0xff)))
    $result.AddRange($type)
    $result.AddRange($data)
    $result.AddRange([byte[]] @(
            [byte] (($crc -shr 24) -band 0xff),
            [byte] (($crc -shr 16) -band 0xff),
            [byte] (($crc -shr 8) -band 0xff),
            [byte] ($crc -band 0xff)))
    $result.AddRange([byte[]] $Bytes[$iendOffset..($Bytes.Length - 1)])
    return $result.ToArray()
}

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
                    $pdfText = "BENCHMARKREPORT $engineName"
                    $variant = $engineName -eq 'OfficeIMO' ? 0 : $iteration
                    [System.IO.File]::WriteAllBytes(
                        (Join-Path $evidenceRoot $pdfRelativePath),
                        (New-TestPdfBytes -Text $pdfText -Variant $variant))
                    $pdfFullPath = Join-Path $evidenceRoot $pdfRelativePath
                    $renderOptions = [OfficeIMO.Pdf.PdfPageRenderOptions]::new()
                    $renderOptions.Format = [OfficeIMO.Pdf.PdfPageRenderFormat]::Png
                    $renderOptions.Dpi = 120D
                    $renderOptions.ContinueOnError = $false
                    $renderOptions.MaxPages = 1
                    $managedRender = @([OfficeIMO.Pdf.PdfDocument]::Open(
                            [System.IO.File]::ReadAllBytes($pdfFullPath)).Read.RenderPages('1', $renderOptions))[0]
                    [System.IO.File]::WriteAllBytes(
                        (Join-Path $evidenceRoot $managedRelativePath),
                        $managedRender.Bytes)
                    $externalPrefix = Join-Path $evidenceRoot ([System.IO.Path]::GetFileNameWithoutExtension($externalRelativePath))
                    $null = & pdftoppm -f 1 -l 1 -singlefile -png -r 120 $pdfFullPath $externalPrefix 2>&1
                    if ($LASTEXITCODE -ne 0) { throw 'Could not create the artifact-exporter external preview fixture.' }

                    $pdfItem = Get-Item -LiteralPath (Join-Path $evidenceRoot $pdfRelativePath)
                    $managedItem = Get-Item -LiteralPath (Join-Path $evidenceRoot $managedRelativePath)
                    $externalItem = Get-Item -LiteralPath (Join-Path $evidenceRoot $externalRelativePath)
                    [ordered]@{
                        iteration = $iteration
                        relativePath = $pdfRelativePath
                        sizeBytes = $pdfItem.Length
                        sha256 = (Get-FileHash -LiteralPath $pdfItem.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
                        semanticSha256 = ('a' * 64)
                        processTreeMemory = [ordered]@{
                            peakWorkingSetBytes = 1024
                            sampleCount = 2
                            minimumObservedProcessCount = 1
                            maximumObservedProcessCount = $engineName -eq 'Chromium' ? 2 : 1
                            sampler = 'contract-test'
                        }
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
                            structureTypeCounts = [ordered]@{ Document = 1; P = 1 }
                        }
                        managedVisual = [ordered]@{
                            relativePath = $managedRelativePath
                            width = $managedRender.Width
                            height = $managedRender.Height
                            sizeBytes = $managedItem.Length
                            sha256 = (Get-FileHash -LiteralPath $managedItem.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
                        }
                        externalVisual = [ordered]@{
                            relativePath = $externalRelativePath
                            width = $managedRender.Width
                            height = $managedRender.Height
                            sizeBytes = $externalItem.Length
                            sha256 = (Get-FileHash -LiteralPath $externalItem.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
                        }
                    }
                }
            )
            [ordered]@{
                engine = $engineName
                cancellation = [ordered]@{
                    apiSupportsCancellation = $engineName -in @('OfficeIMO', 'Chromium')
                    status = $engineName -in @('OfficeIMO', 'Chromium') ? 'Passed' : 'Unsupported'
                    detail = $engineName -in @('OfficeIMO', 'Chromium') `
                        ? "An in-flight $engineName PDF request cancelled in 2 ms." `
                        : 'The compared public conversion entry point does not accept a CancellationToken.'
                }
                determinism = [ordered]@{
                    exactBytesIdentical = $engineName -eq 'OfficeIMO'
                    semanticOutputIdentical = $true
                    managedVisualPreviewIdentical = $true
                    externalVisualPreviewIdentical = $true
                    uniqueByteHashCount = $engineName -eq 'OfficeIMO' ? 1 : 3
                    uniqueSemanticHashCount = 1
                    uniqueManagedVisualHashCount = 1
                    uniqueExternalVisualHashCount = 1
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
            osFamily = $isWindowsHost ? 'Windows' : 'Linux'
            osDescription = $osDescription
            externalRasterizer = 'contract-test'
        }
        provenance = [ordered]@{
            officeIMO = [ordered]@{
                kind = 'source'
                version = '3.2.5+1111111111111111111111111111111111111111'
                commit = ('1' * 40)
                worktreeClean = $true
            }
            htmlTinkerX = [ordered]@{
                kind = 'source'
                version = '3.0.1+2222222222222222222222222222222222222222'
                commit = ('2' * 40)
                worktreeClean = $true
            }
        }
        input = [ordered]@{
            relativePath = 'input.html'
            sizeBytes = $insideItem.Length
            sha256 = $insideHash
            expectedPageCount = 1
            expectedReportMarkerCount = 1
        }
        engines = $engines
    }

    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))

    $report.provenance.officeIMO.worktreeClean = $false
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $dirtySourceRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'clean 40-character source commit') { throw }
        $dirtySourceRejected = $true
    }
    if (-not $dirtySourceRejected) {
        throw 'HTML/PDF artifact exporter accepted dirty OfficeIMO source provenance.'
    }
    $report.provenance.officeIMO.worktreeClean = $true

    $report.provenance.htmlTinkerX.commit = 'not-a-commit'
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $malformedSourceCommitRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'clean 40-character source commit') { throw }
        $malformedSourceCommitRejected = $true
    }
    if (-not $malformedSourceCommitRejected) {
        throw 'HTML/PDF artifact exporter accepted malformed HtmlTinkerX source provenance.'
    }
    $report.provenance.htmlTinkerX.commit = ('2' * 40)

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
    $officeEngine = $report.engines | Where-Object { $_.engine -eq 'OfficeIMO' }
    $officeEngine.cancellation.detail = 'A pre-cancelled OfficeIMO request was rejected.'
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $entryOnlyCancellationRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'cancellation') { throw }
        $entryOnlyCancellationRejected = $true
    }
    if (-not $entryOnlyCancellationRejected) {
        throw 'HTML/PDF artifact exporter accepted entry-only cancellation evidence.'
    }
    $officeEngine.cancellation.detail = 'An in-flight OfficeIMO PDF request cancelled in 2 ms.'

    $officeEngine.cancellation.apiSupportsCancellation = $false
    $officeEngine.cancellation.status = 'Unsupported'
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $missingSupportedCancellationRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'cancellation') { throw }
        $missingSupportedCancellationRejected = $true
    }
    if (-not $missingSupportedCancellationRejected) {
        throw 'HTML/PDF artifact exporter accepted missing cancellation proof from a supported engine.'
    }

    $officeEngine.cancellation.apiSupportsCancellation = $true
    $officeEngine.cancellation.status = 'Passed'
    $peachEngine = $report.engines | Where-Object { $_.engine -eq 'PeachPDF' }
    $peachEngine.cancellation.apiSupportsCancellation = $true
    $peachEngine.cancellation.status = 'Passed'
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $falseSupportedCancellationRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'cancellation') { throw }
        $falseSupportedCancellationRejected = $true
    }
    if (-not $falseSupportedCancellationRejected) {
        throw 'HTML/PDF artifact exporter accepted cancellation proof from an unsupported comparison API.'
    }

    $peachEngine.cancellation.apiSupportsCancellation = $false
    $peachEngine.cancellation.status = 'Unsupported'
    $officeEngine.determinism.exactBytesIdentical = $false
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $officeByteDriftRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'determinism contract') { throw }
        $officeByteDriftRejected = $true
    }
    if (-not $officeByteDriftRejected) {
        throw 'HTML/PDF artifact exporter accepted byte drift from OfficeIMO.'
    }

    $officeEngine.determinism.exactBytesIdentical = $true
    $officeEngine.determinism.uniqueByteHashCount = 2
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $falseUniqueCountRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'validated hashes') { throw }
        $falseUniqueCountRejected = $true
    }
    if (-not $falseUniqueCountRejected) {
        throw 'HTML/PDF artifact exporter trusted a false determinism hash count.'
    }
    $officeEngine.determinism.uniqueByteHashCount = 1

    $peachEngine = $report.engines | Where-Object { $_.engine -eq 'PeachPDF' }
    $semanticOutput = $peachEngine.outputs[1]
    $semanticPdfPath = Join-Path $evidenceRoot ([string] $semanticOutput.relativePath)
    $validSemanticPdfBytes = [System.IO.File]::ReadAllBytes($semanticPdfPath)
    [System.IO.File]::WriteAllBytes(
        $semanticPdfPath,
        (New-TestPdfBytes -Text 'BENCHMARKREPORT forged different content' -Variant 2))
    $semanticPdfItem = Get-Item -LiteralPath $semanticPdfPath
    $semanticOutput.sizeBytes = $semanticPdfItem.Length
    $semanticOutput.sha256 = (Get-FileHash -LiteralPath $semanticPdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $forgedSemanticHashRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'determinism contract|preview does not match') { throw }
        $forgedSemanticHashRejected = $true
    }
    if (-not $forgedSemanticHashRejected) {
        throw 'HTML/PDF artifact exporter trusted caller-supplied semantic hashes instead of validated PDF text.'
    }
    [System.IO.File]::WriteAllBytes($semanticPdfPath, $validSemanticPdfBytes)
    $semanticPdfItem = Get-Item -LiteralPath $semanticPdfPath
    $semanticOutput.sizeBytes = $semanticPdfItem.Length
    $semanticOutput.sha256 = (Get-FileHash -LiteralPath $semanticPdfPath -Algorithm SHA256).Hash.ToLowerInvariant()

    $chromiumEngine = $report.engines | Where-Object { $_.engine -eq 'Chromium' }
    $chromiumEngine.outputs[0].processTreeMemory.maximumObservedProcessCount = 1
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $singleProcessChromiumRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'process-tree memory evidence') { throw }
        $singleProcessChromiumRejected = $true
    }
    if (-not $singleProcessChromiumRejected) {
        throw 'HTML/PDF artifact exporter accepted Chromium evidence that never observed a child process.'
    }

    $chromiumEngine.outputs[0].processTreeMemory.maximumObservedProcessCount = 2
    $report.environment.osFamily = 'macOS'
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

    $report.environment.osFamily = 'FreeBSD'
    $report.environment.osDescription = 'FreeBSD contract-test host'
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $unrecognizedOsRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform linux `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'not a Linux run') { throw }
        $unrecognizedOsRejected = $true
    }
    if (-not $unrecognizedOsRejected) {
        throw 'HTML/PDF artifact exporter derived Linux provenance from the exporter host instead of the evidence report.'
    }

    $report.environment.osFamily = $isWindowsHost ? 'Windows' : 'Linux'
    $report.environment.osDescription = $isWindowsHost ? $osDescription : 'Ubuntu 24.04.3 LTS'
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
        -EvidencePath $evidenceRoot `
        -Platform $platform `
        -OutputPath $outputPath
    if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
        throw 'HTML/PDF artifact exporter did not produce the valid in-root summary.'
    }

    $taggedOutput = $report.engines[0].outputs[0]
    $taggedOutput.contract.structureElementCount = 3
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $forgedTaggedContractRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'tagged-PDF contract') { throw }
        $forgedTaggedContractRejected = $true
    }
    if (-not $forgedTaggedContractRejected) {
        throw 'HTML/PDF artifact exporter trusted tagged-PDF claims instead of independently inspecting the PDF.'
    }
    $taggedOutput.contract.structureElementCount = 2

    $managedPath = Join-Path $evidenceRoot ([string] $taggedOutput.managedVisual.relativePath)
    $originalManagedArtifacts = [System.Collections.Generic.Dictionary[string, byte[]]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($candidateOutput in @($report.engines[0].outputs)) {
        $candidatePath = Join-Path $evidenceRoot ([string] $candidateOutput.managedVisual.relativePath)
        $candidateBytes = [System.IO.File]::ReadAllBytes($candidatePath)
        $originalManagedArtifacts.Add($candidatePath, $candidateBytes)
        $repackedManagedBytes = Add-PngTextChunk -Bytes $candidateBytes
        $managedHash = [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($candidateBytes))
        $repackedManagedHash = [Convert]::ToHexString(
            [System.Security.Cryptography.SHA256]::HashData($repackedManagedBytes))
        if ([string]::Equals($managedHash, $repackedManagedHash, [System.StringComparison]::Ordinal)) {
            throw 'PNG contract fixture did not change the encoded artifact.'
        }
        [System.IO.File]::WriteAllBytes($candidatePath, $repackedManagedBytes)
        $candidateItem = Get-Item -LiteralPath $candidatePath
        $candidateOutput.managedVisual.sizeBytes = $candidateItem.Length
        $candidateOutput.managedVisual.sha256 = (Get-FileHash -LiteralPath $candidatePath -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
        -EvidencePath $evidenceRoot `
        -Platform $platform `
        -OutputPath $outputPath
    if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
        throw 'HTML/PDF artifact exporter rejected a byte-distinct PNG with identical decoded pixels.'
    }
    foreach ($candidateOutput in @($report.engines[0].outputs)) {
        $candidatePath = Join-Path $evidenceRoot ([string] $candidateOutput.managedVisual.relativePath)
        [System.IO.File]::WriteAllBytes($candidatePath, $originalManagedArtifacts[$candidatePath])
        $candidateItem = Get-Item -LiteralPath $candidatePath
        $candidateOutput.managedVisual.sizeBytes = $candidateItem.Length
        $candidateOutput.managedVisual.sha256 = (Get-FileHash -LiteralPath $candidatePath -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    $managedBytes = $originalManagedArtifacts[$managedPath]

    $externalPath = Join-Path $evidenceRoot ([string] $taggedOutput.externalVisual.relativePath)
    [System.IO.File]::WriteAllBytes($managedPath, [System.IO.File]::ReadAllBytes($externalPath))
    $managedItem = Get-Item -LiteralPath $managedPath
    $taggedOutput.managedVisual.sizeBytes = $managedItem.Length
    $taggedOutput.managedVisual.sha256 = (Get-FileHash -LiteralPath $managedPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $unrelatedManagedPreviewRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'Managed preview does not match') { throw }
        $unrelatedManagedPreviewRejected = $true
    }
    if (-not $unrelatedManagedPreviewRejected) {
        throw 'HTML/PDF artifact exporter accepted an unrelated managed preview.'
    }
    [System.IO.File]::WriteAllBytes($managedPath, $managedBytes)
    $managedItem = Get-Item -LiteralPath $managedPath
    $taggedOutput.managedVisual.sizeBytes = $managedItem.Length
    $taggedOutput.managedVisual.sha256 = (Get-FileHash -LiteralPath $managedPath -Algorithm SHA256).Hash.ToLowerInvariant()

    $report.input.expectedPageCount = 2
    foreach ($engine in $report.engines) {
        foreach ($output in $engine.outputs) {
            $output.contract.pageCount = 2
        }
    }
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $wrongPageCountRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'unexpected page count') { throw }
        $wrongPageCountRejected = $true
    }
    if (-not $wrongPageCountRejected) {
        throw 'HTML/PDF artifact exporter accepted a PDF whose page count contradicted the evidence report.'
    }
    $report.input.expectedPageCount = 1
    foreach ($engine in $report.engines) {
        foreach ($output in $engine.outputs) {
            $output.contract.pageCount = 1
        }
    }

    $firstOutput = $report.engines[0].outputs[0]
    $firstPdfPath = Join-Path $evidenceRoot ([string] $firstOutput.relativePath)
    $validPdfBytes = [System.IO.File]::ReadAllBytes($firstPdfPath)
    [System.IO.File]::WriteAllBytes($firstPdfPath, [System.Text.Encoding]::ASCII.GetBytes('not a PDF'))
    $firstPdfItem = Get-Item -LiteralPath $firstPdfPath
    $firstOutput.sizeBytes = $firstPdfItem.Length
    $firstOutput.sha256 = (Get-FileHash -LiteralPath $firstPdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $invalidPdfRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'Executable PDF validation failed') { throw }
        $invalidPdfRejected = $true
    }
    if (-not $invalidPdfRejected) {
        throw 'HTML/PDF artifact exporter accepted arbitrary bytes as validated PDF evidence.'
    }

    [System.IO.File]::WriteAllBytes($firstPdfPath, (New-TestPdfBytes -Text 'missing marker'))
    $firstPdfItem = Get-Item -LiteralPath $firstPdfPath
    $firstOutput.sizeBytes = $firstPdfItem.Length
    $firstOutput.sha256 = (Get-FileHash -LiteralPath $firstPdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    $wrongMarkerCountRejected = $false
    try {
        & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
            -EvidencePath $evidenceRoot `
            -Platform $platform `
            -OutputPath $outputPath
    } catch {
        if ($_.Exception.Message -notmatch 'unexpected report-marker count') { throw }
        $wrongMarkerCountRejected = $true
    }
    if (-not $wrongMarkerCountRejected) {
        throw 'HTML/PDF artifact exporter accepted a PDF whose report-marker count contradicted the evidence report.'
    }

    [System.IO.File]::WriteAllBytes($firstPdfPath, $validPdfBytes)
    $firstPdfItem = Get-Item -LiteralPath $firstPdfPath
    $firstOutput.sizeBytes = $firstPdfItem.Length
    $firstOutput.sha256 = (Get-FileHash -LiteralPath $firstPdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $json = ($report | ConvertTo-Json -Depth 20).Replace("`r`n", "`n") + "`n"
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))

    foreach ($protectedOutputPath in @($reportPath, $insidePath)) {
        $protectedOutputRejected = $false
        try {
            & "$PSScriptRoot/Export-HtmlPdfArtifactEvidence.ps1" `
                -EvidencePath $evidenceRoot `
                -Platform $platform `
                -OutputPath $protectedOutputPath
        } catch {
            if ($_.Exception.Message -notmatch 'cannot overwrite') { throw }
            $protectedOutputRejected = $true
        }
        if (-not $protectedOutputRejected) {
            throw 'HTML/PDF artifact exporter accepted an output path that overwrites validated evidence.'
        }
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
