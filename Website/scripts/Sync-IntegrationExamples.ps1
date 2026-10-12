[CmdletBinding()]
param([switch] $Check)

$ErrorActionPreference = 'Stop'
$websiteRoot = Split-Path -Parent $PSScriptRoot
$downloadRoot = Join-Path $websiteRoot 'static/downloads/integrations'
$examples = [ordered]@{
    'supplier-review' = 'Recipe-HtmlSupplierReview.ps1'
    'event-brief' = 'Recipe-PSEventViewer-OfficeReport.ps1'
    'report-email' = 'Recipe-Mailozaurr-PdfDelivery.ps1'
}
$sourceData = [ordered]@{}
$manifestPath = Join-Path $downloadRoot 'manifest.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json -Depth 20
$zipArtifacts = @()
foreach ($id in $examples.Keys) {
    $name = $examples[$id]
    $sourcePath = Join-Path $downloadRoot "source/$name"
    $sourceBytes = [IO.File]::ReadAllBytes($sourcePath)
    $sourceArtifact = @($manifest.artifacts | Where-Object path -EQ "/downloads/integrations/source/$name")
    if ($sourceArtifact.Count -ne 1) { throw "Expected one source manifest entry for $name." }
    $sourceArtifact[0].bytes = $sourceBytes.Length
    $sourceArtifact[0].sha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($sourceBytes)).ToLowerInvariant()
    $sourceData[$id] = [ordered]@{
        download = "/downloads/integrations/source/$name"
        source = [IO.File]::ReadAllText($sourcePath).Replace("`r`n", "`n")
    }
    $stream = [IO.MemoryStream]::new()
    $archive = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create, $true)
    try {
        $files = @($name)
        if ($id -eq 'supplier-review') { $files += 'Input/SupplierOffers.html' }
        foreach ($file in $files) {
            $entry = $archive.CreateEntry($file, [IO.Compression.CompressionLevel]::Optimal)
            $entry.LastWriteTime = [datetimeoffset]::new(2000, 1, 1, 0, 0, 0, [timespan]::Zero)
            $output = $entry.Open()
            try {
                $bytes = [IO.File]::ReadAllBytes((Join-Path $downloadRoot "source/$file"))
                $output.Write($bytes, 0, $bytes.Length)
            } finally { $output.Dispose() }
        }
    } finally { $archive.Dispose() }
    $zipBytes = $stream.ToArray()
    $stream.Dispose()
    $zipPath = Join-Path $downloadRoot "$id.zip"
    if ($Check) {
        if (-not (Test-Path -LiteralPath $zipPath) -or
            [Convert]::ToHexString([IO.File]::ReadAllBytes($zipPath)) -cne [Convert]::ToHexString($zipBytes)) {
            throw "Refresh $id.zip with Sync-IntegrationExamples.ps1."
        }
    } else { [IO.File]::WriteAllBytes($zipPath, $zipBytes) }
    $zipArtifacts += [ordered]@{
        path = "/downloads/integrations/$id.zip"
        bytes = $zipBytes.Length
        sha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($zipBytes)).ToLowerInvariant()
    }
}
$dataPath = Join-Path $websiteRoot 'data/integration_examples.json'
$dataText = ($sourceData | ConvertTo-Json -Depth 5) + "`n"
$manifest.generation = 'Run a downloaded example from any folder with published modules and explicit OutputDirectory. ZIPs and displayed source: scripts/Sync-IntegrationExamples.ps1. PDF previews: pdftocairo -png -singlefile -scale-to 1400. Workbook previews: Export-OfficeExcelRangeImage, Offers A1:D10 and Events A1:G10.'
$manifest.artifacts = @($manifest.artifacts | Where-Object path -NotMatch '^/downloads/integrations/(supplier-review|event-brief|report-email)\.zip$') + $zipArtifacts
$manifestText = ($manifest | ConvertTo-Json -Depth 20) + "`n"
foreach ($item in @(@($dataPath, $dataText), @($manifestPath, $manifestText))) {
    if ($Check) {
        if (-not (Test-Path -LiteralPath $item[0]) -or [IO.File]::ReadAllText($item[0]).Replace("`r`n", "`n") -cne $item[1].Replace("`r`n", "`n")) {
            throw "Refresh $($item[0]) with Sync-IntegrationExamples.ps1."
        }
    } else { [IO.File]::WriteAllText($item[0], $item[1], [Text.UTF8Encoding]::new($false)) }
}
Write-Host 'Verified three standalone example ZIPs, displayed scripts and download hashes.'
