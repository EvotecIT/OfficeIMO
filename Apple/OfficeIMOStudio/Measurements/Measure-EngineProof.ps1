<#
.SYNOPSIS
Runs the engine journey and the feasibility measurements of the native Studio app on this Mac and in an iPad
simulator, and measures the engine's contribution to the app's size.

.DESCRIPTION
Requires OfficeIMOStudioEngine.xcframework in Apple/OfficeIMOStudio/Engine (see OfficeIMO.Studio.Native/Build-AppleEngine.ps1).
The script:
- builds the measurement PDFs (make-measurement-pdfs.swift) unless they exist;
- builds the macOS app, runs the journey sandboxed, and runs each measurement in an unsandboxed ad hoc copy whose
  folders point into the work directory;
- creates a temporary iPad simulator, runs the journey and the measurements there, and deletes the simulator;
- builds the iOS device app without signing and reads the linker map for the engine's share of the executable.
Each measurement launches the app once per document and run, so launch timing and the peak footprint belong to that
document. Results are written to <WorkDirectory>/results.

Simulator timings show the engine works on iPadOS and how it scales; they are not device performance.

.EXAMPLE
./Measure-EngineProof.ps1 -WorkDirectory /tmp/engine-proof
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$WorkDirectory,
    [int]$Runs = 3,
    [string]$SimulatorType = 'iPad Pro 13-inch (M5)',
    [string[]]$Documents = @('openpreserve-high-resolution-image.pdf', 'long.pdf', 'scanned.pdf'),
    [ValidateSet('macOS', 'Simulator', 'Size', 'Summary')][string[]]$Part = @('macOS', 'Simulator', 'Size')
)
$ErrorActionPreference = 'Stop'
$appRoot = Split-Path $PSScriptRoot
$repositoryRoot = (Resolve-Path (Join-Path $appRoot '../..')).Path
$project = Join-Path $appRoot 'OfficeIMOStudio.xcodeproj'
$bundleId = 'com.evotec.officeimo.studio'
New-Item -ItemType Directory -Force $WorkDirectory | Out-Null
$WorkDirectory = (Resolve-Path $WorkDirectory).Path
$corpus = Join-Path $WorkDirectory 'corpus'
$results = Join-Path $WorkDirectory 'results'
$derived = Join-Path $WorkDirectory 'DerivedData'
New-Item -ItemType Directory -Force $results | Out-Null

function Invoke-Tool([string]$Tool, [string[]]$Arguments) {
    & $Tool @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$Tool $($Arguments -join ' ') failed with exit code $LASTEXITCODE." }
}

# Collects the JSON results the app prints between its result markers.
function Read-Results([string[]]$Output) {
    $text = $Output -join "`n"
    foreach ($match in [regex]::Matches($text, '(?s)OFFICEIMO-RESULT (\S+)\n(.*?)\nOFFICEIMO-RESULT-END')) {
        [pscustomobject]@{ Name = $match.Groups[1].Value; Value = $match.Groups[2].Value | ConvertFrom-Json }
    }
}

function Save-Run([string]$Platform, [string]$Name, [int]$Run, [object[]]$Collected) {
    $result = @($Collected | Where-Object Name -eq $Name)
    if ($result.Count -ne 1) { throw "$Platform run $Run produced no $Name result." }
    $result[0].Value | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $results "$Platform-$Run-$Name")
    $result[0].Value
}

$documentSources = @{
    'openpreserve-high-resolution-image.pdf' = Join-Path $repositoryRoot 'OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/openpreserve-high-resolution-image.pdf'
    'long.pdf' = Join-Path $corpus 'long.pdf'
    'scanned.pdf' = Join-Path $corpus 'scanned.pdf'
}
if (-not (Test-Path (Join-Path $corpus 'long.pdf')) -or -not (Test-Path (Join-Path $corpus 'scanned.pdf'))) {
    Push-Location $repositoryRoot
    try { Invoke-Tool swift @((Join-Path $PSScriptRoot 'make-measurement-pdfs.swift'), $corpus) } finally { Pop-Location }
}

# macOS
if ('macOS' -in $Part) {
Write-Host 'Building the macOS app'
Invoke-Tool xcodebuild @('-project', $project, '-scheme', 'OfficeIMOStudio-macOS', '-configuration', 'Release', '-destination', 'platform=macOS',
    '-derivedDataPath', $derived, 'LD_GENERATE_MAP_FILE=YES', 'build', '-quiet')
$macApp = Join-Path $derived 'Build/Products/Release/OfficeIMO Studio.app'
$macExecutable = Join-Path $macApp 'Contents/MacOS/OfficeIMO Studio'
for ($run = 1; $run -le $Runs; $run++) {
    $null = Save-Run 'macos' 'journey.json' $run (Read-Results (& $macExecutable -OfficeIMOJourney -OfficeIMOExit))
}
# An ad hoc copy without entitlements runs outside the sandbox, so its folders can point into the work directory.
$measureApp = Join-Path $WorkDirectory 'OfficeIMO Studio (measurement).app'
if (Test-Path $measureApp) { Remove-Item -Recurse -Force $measureApp }
Copy-Item -Recurse $macApp $measureApp
Invoke-Tool codesign @('--force', '--deep', '--sign', '-', $measureApp)
$macFolder = Join-Path $WorkDirectory 'macos-folder'
New-Item -ItemType Directory -Force (Join-Path $macFolder 'Measurements') | Out-Null
foreach ($document in $Documents) { Copy-Item $documentSources[$document] (Join-Path $macFolder 'Measurements') -Force }
$env:OFFICEIMO_STUDIO_FOLDER = $macFolder
try {
    foreach ($document in $Documents) {
        for ($run = 1; $run -le $Runs; $run++) {
            Write-Host "macOS: $document, run $run"
            $output = & (Join-Path $measureApp 'Contents/MacOS/OfficeIMO Studio') -OfficeIMOMeasure $document -OfficeIMOExit
            $null = Save-Run 'macos' "measure-$document.json" $run (Read-Results $output)
        }
    }
} finally { Remove-Item Env:OFFICEIMO_STUDIO_FOLDER }
}

# iPad simulator
if ('Simulator' -in $Part) {
$runtime = (xcrun simctl list runtimes --json | ConvertFrom-Json).runtimes |
    Where-Object { $_.platform -eq 'iOS' -and $_.isAvailable -and $SimulatorType -in $_.supportedDeviceTypes.name } |
    Sort-Object { [version]$_.version } | Select-Object -Last 1
if (-not $runtime) { throw "No available iOS runtime supports $SimulatorType." }
$simulator = (xcrun simctl create "OfficeIMO Engine Proof" $SimulatorType $runtime.identifier)
if ($LASTEXITCODE -ne 0) { throw "Could not create a $SimulatorType simulator." }
$simulator = $simulator.Trim()
try {
    Invoke-Tool xcrun @('simctl', 'boot', $simulator)
    Invoke-Tool xcrun @('simctl', 'bootstatus', $simulator)
    Write-Host "Building the iOS app for $SimulatorType, iOS $($runtime.version)"
    Invoke-Tool xcodebuild @('-project', $project, '-scheme', 'OfficeIMOStudio-iOS', '-configuration', 'Release', '-destination', "id=$simulator",
        '-derivedDataPath', $derived, 'build', '-quiet')
    Invoke-Tool xcrun @('simctl', 'install', $simulator, (Join-Path $derived 'Build/Products/Release-iphonesimulator/OfficeIMO Studio.app'))
    for ($run = 1; $run -le $Runs; $run++) {
        $output = xcrun simctl launch --console-pty --terminate-running-process $simulator $bundleId -OfficeIMOJourney -OfficeIMOExit
        $null = Save-Run 'simulator' 'journey.json' $run (Read-Results $output)
    }
    $container = (xcrun simctl get_app_container $simulator $bundleId data).Trim()
    $measurements = Join-Path $container 'Documents/Measurements'
    New-Item -ItemType Directory -Force $measurements | Out-Null
    foreach ($document in $Documents) { Copy-Item $documentSources[$document] $measurements -Force }
    foreach ($document in $Documents) {
        for ($run = 1; $run -le $Runs; $run++) {
            Write-Host "Simulator: $document, run $run"
            $output = xcrun simctl launch --console-pty --terminate-running-process $simulator $bundleId -OfficeIMOMeasure $document -OfficeIMOExit
            $null = Save-Run 'simulator' "measure-$document.json" $run (Read-Results $output)
        }
    }
} finally {
    xcrun simctl shutdown $simulator 2>$null
    xcrun simctl delete $simulator
}
}

# Size: the engine's share of the device executable, from the linker map.
function Measure-LinkMap([string]$MapFile) {
    $objects = @{}
    $engine = 0L; $total = 0L
    $section = ''
    foreach ($line in [IO.File]::ReadLines($MapFile)) {
        if ($line.StartsWith('# Object files:')) { $section = 'objects'; continue }
        if ($line.StartsWith('# Sections:')) { $section = 'sections'; continue }
        if ($line.StartsWith('# Symbols:')) { $section = 'symbols'; continue }
        if ($line.StartsWith('# Dead Stripped Symbols:')) { break }
        if ($section -eq 'objects' -and $line -match '^\[\s*(\d+)\]\s+(.*)$') { $objects[[int]$Matches[1]] = $Matches[2] }
        elseif ($section -eq 'symbols' -and $line -match '^0x[0-9A-F]+\s+0x([0-9A-F]+)\s+\[\s*(\d+)\]') {
            $size = [Convert]::ToInt64($Matches[1], 16)
            $total += $size
            if ($objects[[int]$Matches[2]] -like '*libOfficeIMOStudioEngine.a*') { $engine += $size }
        }
    }
    [ordered]@{ engineBytes = $engine; linkedBytes = $total }
}

if ('Size' -in $Part) {
    Write-Host 'Building the iOS device app for size'
    Invoke-Tool xcodebuild @('-project', $project, '-scheme', 'OfficeIMOStudio-iOS', '-configuration', 'Release', '-destination', 'generic/platform=iOS',
        '-derivedDataPath', $derived, 'CODE_SIGNING_ALLOWED=NO', 'LD_GENERATE_MAP_FILE=YES', 'build', '-quiet')
    $iosApp = Join-Path $derived 'Build/Products/Release-iphoneos/OfficeIMO Studio.app'
    $iosMap = Get-ChildItem -Recurse (Join-Path $derived 'Build/Intermediates.noindex') -Filter '*-LinkMap-normal-arm64.txt' |
        Where-Object FullName -like '*Release-iphoneos*' | Select-Object -First 1
    $zip = Join-Path $WorkDirectory 'ios-app.zip'
    if (Test-Path $zip) { Remove-Item $zip }
    Invoke-Tool ditto @('-c', '-k', '--keepParent', $iosApp, $zip)
    $size = [ordered]@{ ios = Measure-LinkMap $iosMap.FullName }
    $size.ios.executableBytes = (Get-Item (Join-Path $iosApp 'OfficeIMO Studio')).Length
    $size.ios.appBytes = (Get-ChildItem -Recurse -File $iosApp | Measure-Object Length -Sum).Sum
    $size.ios.zippedAppBytes = (Get-Item $zip).Length
    $macMap = Get-ChildItem -Recurse (Join-Path $derived 'Build/Intermediates.noindex') -Filter '*-LinkMap-normal-arm64.txt' |
        Where-Object FullName -notlike '*-iphone*' | Select-Object -First 1
    if ($macMap) {
        $size.macOS = Measure-LinkMap $macMap.FullName
        $size.macOS.universalExecutableBytes = (Get-Item (Join-Path $derived 'Build/Products/Release/OfficeIMO Studio.app/Contents/MacOS/OfficeIMO Studio')).Length
    }
    $size | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $results 'size.json')
}

# Summary: medians over the runs of each platform and document.
function Get-Median([double[]]$Values) {
    $sorted = @($Values | Where-Object { $_ -ne $null } | Sort-Object)
    if ($sorted.Count -eq 0) { return $null }
    [math]::Round($sorted[[int][math]::Floor(($sorted.Count - 1) / 2)], 1)
}
$summary = [ordered]@{}
foreach ($platform in 'macos', 'simulator') {
    $journeys = @(Get-ChildItem $results -Filter "$platform-*-journey.json" | ForEach-Object { Get-Content $_ -Raw | ConvertFrom-Json })
    if ($journeys.Count -eq 0) { continue }
    $entry = [ordered]@{
        journeyPassed = @($journeys | Where-Object passed).Count -eq $journeys.Count
        engineFirstCallMs = Get-Median ($journeys.launch.engineFirstCall)
        processToAppMs = Get-Median ($journeys.launch.processToApp)
    }
    foreach ($document in $Documents) {
        $reports = @(Get-ChildItem $results -Filter "$platform-*-measure-$document.json" | ForEach-Object { Get-Content $_ -Raw | ConvertFrom-Json })
        if ($reports.Count -eq 0) { continue }
        $entry[$document] = [ordered]@{
            device = $reports[0].device
            pages = $reports[0].pages
            openMs = Get-Median $reports.openMilliseconds
            firstPageMs = Get-Median $reports.firstPageMilliseconds
            firstPagePixels = "$($reports[0].firstPage.pixelWidth) x $($reports[0].firstPage.pixelHeight)"
            firstPageRenderMs = Get-Median $reports.firstPage.milliseconds
            repeatFirstPageMs = Get-Median $reports.repeatFirstPage.milliseconds
            firstPagePngRenderMs = Get-Median $reports.firstPagePng.renderMilliseconds
            firstPagePngDecodeMs = Get-Median $reports.firstPagePng.decodeMilliseconds
            sampledMedianMs = Get-Median $reports.sampledSummary.median
            sampledP90Ms = Get-Median $reports.sampledSummary.p90
            zoom = @(foreach ($view in $reports[0].zoom) {
                $factor = $view.zoom
                $views = @($reports | ForEach-Object { $_.zoom } | Where-Object { $_.zoom -eq $factor })
                if ($view.error) { "x$($factor): $($view.error)" }
                else { "x$($factor): $($view.tiles) tiles, $($view.pixelWidth) x $($view.pixelHeight) px, $(Get-Median $views.milliseconds) ms" }
            })
            peakFootprintMB = Get-Median ($reports.footprintAfter.peak | ForEach-Object { $_ / 1MB })
            processToAppMs = Get-Median $reports.launch.processToApp
        }
    }
    $summary[$platform] = $entry
}
if (Test-Path (Join-Path $results 'size.json')) { $summary.size = Get-Content (Join-Path $results 'size.json') -Raw | ConvertFrom-Json }
$summary | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $results 'summary.json')
Write-Host "Results: $results"
