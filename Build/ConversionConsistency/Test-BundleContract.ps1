[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $BundlePath,
    [Parameter(Mandatory)][string] $OutputPath,
    [string] $PositionBundlePath,
    [string] $ToolPath = (Join-Path $PSScriptRoot 'bin/Release/net8.0/OfficeIMO.ConversionConsistency.Tool.dll')
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$BundlePath = (Resolve-Path -LiteralPath $BundlePath).Path
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $OutputPath) { throw 'Negative-proof output must not already exist.' }
New-Item -ItemType Directory -Path $OutputPath | Out-Null
$bundle = Get-Content -LiteralPath (Join-Path $BundlePath 'bundle.json') -Raw | ConvertFrom-Json -Depth 100
$sourceCase = $bundle.cases | Where-Object { $_.contract.id -eq 'native-pptx' } | Select-Object -First 1
if ($null -eq $sourceCase) { throw 'Expected the native-pptx fixture in the source bundle.' }

foreach ($failure in @('missing-label', 'missing-image', 'missing-format', 'changed-image', 'duplicate-page')) {
    $target = Join-Path $OutputPath $failure
    New-Item -ItemType Directory -Path $target | Out-Null
    Copy-Item -LiteralPath (Join-Path $BundlePath $sourceCase.contract.id) -Destination $target -Recurse
    $candidate = $bundle | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    $candidate.cases = @($candidate.cases | Where-Object { $_.contract.id -eq $sourceCase.contract.id })
    $entry = $candidate.cases[0]
    switch ($failure) {
        'missing-label' { $entry.contract.pages[0].text += 'REQUIRED-BUT-ABSENT' }
        'missing-image' { $entry.images[0].path = $entry.contract.id + '/missing.png' }
        'missing-format' { $entry.images = @($entry.images | Where-Object { $_.format -ne 'Webp' }) }
        'changed-image' { [IO.File]::AppendAllText((Join-Path $target $entry.images[0].path), 'tampered') }
        'duplicate-page' { $entry.images += $entry.images[0] }
    }
    $candidate | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath (Join-Path $target 'bundle.json') -Encoding utf8
    & dotnet $ToolPath verify --output $target *> (Join-Path $target 'verification.log')
    if ($LASTEXITCODE -ne 1) { throw "$failure did not return a structured verification failure." }
    $report = Get-Content -LiteralPath (Join-Path $target 'consistency-result.json') -Raw | ConvertFrom-Json -Depth 100
    if ($report.passed -or $report.cases[0].passed) { throw "$failure was incorrectly accepted." }
    if ($failure -eq 'missing-format' -and $report.cases[0].errors -notcontains 'Webp page sequence is missing, duplicated, or unexpected.') {
        throw 'The missing WebP format was not rejected by the required page-sequence contract.'
    }
}

if ($PositionBundlePath) {
    $PositionBundlePath = (Resolve-Path -LiteralPath $PositionBundlePath).Path
    $positions = Get-Content -LiteralPath (Join-Path $PositionBundlePath 'bundle.json') -Raw | ConvertFrom-Json -Depth 100
    $positionCase = $positions.cases | Where-Object { $_.contract.id -eq 'text-geometry' } | Select-Object -First 1
    if ($null -eq $positionCase) { throw 'Expected the authored text-geometry position fixture.' }
    foreach ($failure in @('moved-outline', 'missing-outline-ink', 'moved-selectable-text')) {
        $target = Join-Path $OutputPath $failure
        New-Item -ItemType Directory -Path $target | Out-Null
        Copy-Item -LiteralPath (Join-Path $PositionBundlePath $positionCase.contract.id) -Destination $target -Recurse
        $candidate = $positions | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
        $candidate.cases = @($candidate.cases | Where-Object { $_.contract.id -eq $positionCase.contract.id })
        $entry = $candidate.cases[0]
        $position = $entry.contract.pages[0].textPositions[0]
        $artifact = $entry.images | Where-Object { $_.format -eq 'Svg' -and $_.page -eq 1 } | Select-Object -First 1
        if ($null -eq $artifact) { throw 'The position fixture needs its actual SVG artifact.' }
        $svgPath = Join-Path $target $artifact.path
        [xml] $svg = Get-Content -LiteralPath $svgPath -Raw
        $carrier = $svg.SelectNodes('//*[@aria-label]') | Where-Object {
            $_.GetAttribute('aria-label').Contains($position.text) -and $null -ne $_.SelectSingleNode('.//*[local-name()="path"]')
        } | Select-Object -First 1
        if ($null -eq $carrier) { throw 'Expected labelled glyph outlines in the actual text-geometry export.' }
        if ($failure -eq 'moved-outline') {
            $carrier.SetAttribute('transform', $carrier.GetAttribute('transform') + ' translate(0 1)')
        } elseif ($failure -eq 'missing-outline-ink') {
            foreach ($path in @($carrier.SelectNodes('.//*[local-name()="path"]'))) { $path.ParentNode.RemoveChild($path) | Out-Null }
        } else {
            # A minimal selectable SVG exercises the independent browser/PDF
            # baseline path. Its intentional position error must be reported even
            # though its full-page pixels also differ from the source fixture.
            $svg = [xml] '<svg xmlns="http://www.w3.org/2000/svg"><text/></svg>'
            $svg.DocumentElement.SetAttribute('width', [string] $entry.contract.pages[0].width)
            $svg.DocumentElement.SetAttribute('height', [string] $entry.contract.pages[0].height)
            $text = $svg.DocumentElement.FirstChild
            $text.SetAttribute('x', $position.x.ToString([Globalization.CultureInfo]::InvariantCulture))
            $text.SetAttribute('y', ($position.baseline + 1).ToString([Globalization.CultureInfo]::InvariantCulture))
            $text.SetAttribute('font-family', 'Arial')
            $text.SetAttribute('font-size', '20')
            $text.InnerText = $position.text
            $entry.contract.pages[0].text = @($position.text)
            $entry.contract.pages[0].textPositions = @($position)
            $entry.contract.pages[0].svgTextPositions = @()
        }
        $svg.Save($svgPath)
        # Keep the mutated artifact's hash honest, so content/geometry rather
        # than the unrelated tamper guard must reject these controls.
        $artifact.sha256 = (Get-FileHash -LiteralPath $svgPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $candidate | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath (Join-Path $target 'bundle.json') -Encoding utf8
        & dotnet $ToolPath verify --output $target *> (Join-Path $target 'verification.log')
        if ($LASTEXITCODE -ne 1) { throw "$failure did not return a structured verification failure." }
        $report = Get-Content -LiteralPath (Join-Path $target 'consistency-result.json') -Raw | ConvertFrom-Json -Depth 100
        $prefix = if ($failure -eq 'moved-selectable-text') { 'SVG text ' } else { 'SVG outlined text ' }
        $message = $prefix + $position.text + ':'
        if ($report.passed -or -not @($report.cases[0].pages[0].errors | Where-Object { $_.StartsWith($message, [StringComparison]::Ordinal) }).Count) {
            throw "$failure was not rejected by its actual SVG text-position/ink contract."
        }
    }
}

& dotnet $ToolPath verify --output $BundlePath --unknown-option value *> (Join-Path $OutputPath 'unknown-option.log')
if ($LASTEXITCODE -ne 2) { throw 'An unknown CLI option was not rejected.' }
& dotnet $ToolPath verify --output $BundlePath --output $BundlePath *> (Join-Path $OutputPath 'duplicate-option.log')
if ($LASTEXITCODE -ne 2) { throw 'A duplicate CLI option was not rejected.' }
Write-Host 'PASS: missing content, missing/corrupt images, missing formats, duplicate pages, and invalid CLI options are rejected.'
if ($PositionBundlePath) { Write-Host 'PASS: moved outlines, labelled groups without ink, and moved selectable SVG text fail the position contract.' }
# Expected native failures above have been validated; do not leak their exit code
# into callers such as the GitHub Actions PowerShell wrapper.
$global:LASTEXITCODE = 0
