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
        engines = @()
    }

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
} finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}

Write-Host "HTML/PDF artifact exporter path boundary passed on $platform."
