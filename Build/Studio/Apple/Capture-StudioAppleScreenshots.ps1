[CmdletBinding()]
param(
    [ValidateSet('all', 'macos', 'iphone', 'ipad')][string] $Role = 'all',
    [string] $WorkRoot
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$sourceRoot = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
$project = Join-Path $sourceRoot 'Apple/OfficeIMOStudio/OfficeIMOStudio.xcodeproj'
$captureRoot = Join-Path $sourceRoot 'Artifacts/Studio/Apple/screenshots'
$fixture = Join-Path $sourceRoot 'Website/static/downloads/showcase/pdf/showcase-dashboard.pdf'

function Invoke-CaptureXcode([string[]] $Arguments, [string] $LogPath) {
    & xcodebuild @Arguments *> $LogPath
    if ($LASTEXITCODE -ne 0) { throw "Native capture failed. See $LogPath" }
}

if ($Role -eq 'all') {
    if (-not $IsMacOS) { throw 'Native store captures require macOS and Xcode.' }
    if (Test-Path -LiteralPath $captureRoot) { throw 'Use a fresh capture checkout; existing reviewed captures are preserved.' }
    foreach ($name in 'POWERFORGE_DOTNET', 'POWERFORGE_SIMULATOR_CLI', 'POWERFORGE_CAPTURE_SIMULATORS') {
        if (-not [Environment]::GetEnvironmentVariable($name)) { throw "The pinned capture workflow must supply $name." }
    }
    $devices = $env:POWERFORGE_CAPTURE_SIMULATORS | ConvertFrom-Json -AsHashtable
    if ($devices.Count -ne 2 -or -not $devices.ContainsKey('iphone') -or -not $devices.ContainsKey('ipad')) {
        throw 'Select exactly the owned iphone and ipad simulator UUIDs.'
    }
    if (-not $WorkRoot) {
        if (-not $env:RUNNER_TEMP -or -not (Test-Path -LiteralPath $env:RUNNER_TEMP -PathType Container)) {
            throw 'Pass an explicit task scratch WorkRoot outside CI.'
        }
        $WorkRoot = Join-Path $env:RUNNER_TEMP ('studio-store-capture-' + [Guid]::NewGuid().ToString('N'))
    }
    if (-not [IO.Path]::IsPathFullyQualified($WorkRoot) -or
        -not (Test-Path -LiteralPath (Split-Path -Parent $WorkRoot) -PathType Container)) {
        throw 'Use an absolute capture work folder beneath an existing mounted scratch directory.'
    }
    if (Test-Path -LiteralPath $WorkRoot) { throw 'The capture work folder must be new.' }
    $null = New-Item -ItemType Directory -Path $WorkRoot, $captureRoot
    Write-Host "Capture build and diagnostic output: $WorkRoot"
    foreach ($platform in 'macOS', 'iOS') {
        $destination = if ($platform -eq 'macOS') { 'generic/platform=macOS' } else { 'generic/platform=iOS Simulator' }
        $signing = if ($platform -eq 'macOS') {
            @('CODE_SIGN_IDENTITY=-', 'CODE_SIGN_STYLE=Manual', 'DEVELOPMENT_TEAM=')
        } else { @('CODE_SIGNING_ALLOWED=NO') }
        Invoke-CaptureXcode -LogPath (Join-Path $WorkRoot "build-$platform.log") -Arguments (
            @('build-for-testing', '-project', $project, '-scheme', "OfficeIMOStudioCapture-$platform", '-configuration', 'Release',
              '-destination', $destination, '-derivedDataPath', (Join-Path $WorkRoot $platform),
              '-clonedSourcePackagesDirPath', (Join-Path $WorkRoot 'SourcePackages')) + $signing)
    }
    & $PSCommandPath -Role macos -WorkRoot $WorkRoot
    $receiptRoot = Join-Path $sourceRoot 'build/powerforge/screenshots'
    $null = New-Item -ItemType Directory -Path $receiptRoot -Force
    foreach ($deviceRole in 'iphone', 'ipad') {
        $owner = if ($env:GITHUB_REPOSITORY -and $env:GITHUB_RUN_ID) {
            "$env:GITHUB_REPOSITORY/$env:GITHUB_RUN_ID/$deviceRole"
        } else { "OfficeIMO/local-capture/$deviceRole" }
        # PowerForge alone owns boot, execution bounds, cancellation and release.
        & $env:POWERFORGE_DOTNET $env:POWERFORGE_SIMULATOR_CLI apple-simulator run `
            --owner $owner --device $devices[$deviceRole] --require-shutdown `
            --simulator-timeout-seconds 180 --cleanup-timeout-seconds 90 --timeout-seconds 600 --output json -- `
            pwsh -NoLogo -NoProfile -File $PSCommandPath -Role $deviceRole -WorkRoot $WorkRoot `
            > (Join-Path $receiptRoot "$deviceRole.json")
        if ($LASTEXITCODE -ne 0) { throw "Capture or simulator release failed for $deviceRole; inspect its PowerForge receipt." }
    }
    return
}

if (-not $WorkRoot -or -not (Test-Path -LiteralPath $WorkRoot -PathType Container)) { throw 'A prepared capture WorkRoot is required.' }
$platform = if ($Role -eq 'macos') { 'macOS' } else { 'iOS' }
$family = switch ($Role) { 'macos' { 'macOS' }; 'iphone' { 'iPhone' }; 'ipad' { 'iPad' } }
$products = Join-Path $WorkRoot "$platform/Build/Products"
$target = "OfficeIMOStudioCapture-$platform"
$original = @(Get-ChildItem -LiteralPath $products -Filter "$($target)_*.xctestrun")
if ($original.Count -ne 1) { throw 'Expected exactly one capture test manifest from the native build.' }
$appPath = Join-Path $products $(if ($Role -eq 'macos') { 'Release/OfficeIMO Studio.app' } else { 'Release-iphonesimulator/OfficeIMO Studio.app' })
$document = $fixture
$ownedDocumentDirectory = $null
$ownedMacFixtureDirectory = $null
$destination = 'platform=macOS'
try {
    if ($Role -eq 'macos') {
        # Keep this small public input off removable storage; large products and
        # diagnostics remain in WorkRoot. Do not grant the app volume-wide access.
        $nativeTemp = (& getconf DARWIN_USER_TEMP_DIR).Trim()
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $nativeTemp -PathType Container)) {
            throw 'Cannot locate the native user temporary directory.'
        }
        $ownedMacFixtureDirectory = Join-Path $nativeTemp ('OfficeIMOStudioCapture-' + [Guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $ownedMacFixtureDirectory
        $document = Join-Path $ownedMacFixtureDirectory 'Northwind Q3 Review.pdf'
        Copy-Item -LiteralPath $fixture -Destination $document
    } else {
        $devices = $env:POWERFORGE_CAPTURE_SIMULATORS | ConvertFrom-Json -AsHashtable
        $device = [string]$devices[$Role]
        $destination = "platform=iOS Simulator,id=$device"
        & xcrun simctl install $device $appPath
        if ($LASTEXITCODE -ne 0) { throw 'Cannot install the exact capture build.' }
        $container = (& xcrun simctl get_app_container $device com.evotec.officeimo.studio data).Trim()
        if ($LASTEXITCODE -ne 0) { throw 'Cannot locate the capture app document container.' }
        $ownedDocumentDirectory = Join-Path $container ('Documents/StudioStoreCapture-' + [Guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $ownedDocumentDirectory
        $document = Join-Path $ownedDocumentDirectory 'Northwind Q3 Review.pdf'
        Copy-Item -LiteralPath $fixture -Destination $document
    }
    $manifest = (& plutil -convert json -o - $original[0].FullName) | ConvertFrom-Json -AsHashtable
    if ($LASTEXITCODE -ne 0 -or -not $manifest.ContainsKey($target)) { throw 'Unexpected Xcode capture manifest.' }
    $test = $manifest[$target]
    # Keep the capture runner and normal app launch in the intended locale.
    $test.TestLanguage = 'en'
    $test.TestRegion = 'US'
    if ($Role -ne 'macos') {
        # The exact app is already installed and holds the public PDF. Installing
        # it again through Xcode replaces its container and invalidates that URL.
        $test.UseUITargetAppProvidedByTests = $true
        $test.Remove('UITargetAppPath')
        $test.DependentProductPaths = @($test.DependentProductPaths | Where-Object {
            [IO.Path]::GetFileName($_) -ne [IO.Path]::GetFileName($appPath)
        })
    }
    if (-not $test.ContainsKey('EnvironmentVariables')) { $test.EnvironmentVariables = @{} }
    $test.EnvironmentVariables.STUDIO_CAPTURE_DOCUMENT_PATH = $document
    $test.EnvironmentVariables.STUDIO_CAPTURE_ROLE = $Role
    $test.EnvironmentVariables.STUDIO_CAPTURE_APP_PATH = $appPath
    if ($Role -eq 'macos') {
        # Native app scratch belongs in its own sandbox, not the development
        # process's external TMPDIR. This does not change machine settings.
        $test.EnvironmentVariables.STUDIO_CAPTURE_NATIVE_TMPDIR = Join-Path $HOME 'Library/Containers/com.evotec.officeimo.studio/Data/tmp'
    }
    $testManifest = Join-Path $products "StudioStoreCapture-$Role.xctestrun"
    $manifest | ConvertTo-Json -Depth 40 | & plutil -convert xml1 -o $testManifest -
    if ($LASTEXITCODE -ne 0) { throw 'Cannot write the capture test environment.' }
    $result = Join-Path $WorkRoot "$Role.xcresult"
    Invoke-CaptureXcode -LogPath (Join-Path $WorkRoot "capture-$Role.log") -Arguments @(
        'test-without-building', '-xctestrun', $testManifest, '-destination', $destination, '-resultBundlePath', $result,
        '-only-testing', "$target/StudioStoreCaptureTests/testCapturePublicDashboard", '-parallel-testing-enabled', 'NO',
        '-collect-test-diagnostics', 'never', '-test-timeouts-enabled', 'YES',
        '-default-test-execution-time-allowance', '120', '-maximum-test-execution-time-allowance', '150')
    if ($Role -ne 'macos') {
        $after = (& xcrun simctl get_app_container $device com.evotec.officeimo.studio data).Trim()
        if ($LASTEXITCODE -ne 0 -or $after -ne $container -or -not (Test-Path -LiteralPath $document -PathType Leaf)) {
            throw 'The capture app document container changed during the native test.'
        }
    }
    $export = Join-Path $WorkRoot "$Role-attachments"
    & xcrun xcresulttool export attachments --path $result --output-path $export
    if ($LASTEXITCODE -ne 0) { throw 'Cannot export the native capture attachments.' }
    $tests = Get-Content -LiteralPath (Join-Path $export 'manifest.json') -Raw | ConvertFrom-Json
    $attachments = @($tests | ForEach-Object { $_.attachments } | Where-Object { -not $_.isAssociatedWithFailure })
    $names = @("studio-dashboard-$Role")
    if ($Role -ne 'iphone') { $names += "studio-search-$Role" }
    $output = Join-Path $captureRoot "$family/en-US"
    $null = New-Item -ItemType Directory -Path $output
    foreach ($name in $names) {
        $selected = @($attachments | Where-Object { $_.suggestedHumanReadableName.StartsWith($name + '_', [StringComparison]::Ordinal) })
        if ($selected.Count -ne 1) { throw "Expected one successful native attachment for $name." }
        Copy-Item -LiteralPath (Join-Path $export $selected[0].exportedFileName) -Destination (Join-Path $output "$name.png")
    }
} finally {
    if ($ownedMacFixtureDirectory) {
        Remove-Item -LiteralPath $ownedMacFixtureDirectory -Recurse -ErrorAction Stop
    }
    if ($ownedDocumentDirectory) {
        & xcrun simctl terminate $device com.evotec.officeimo.studio 2>$null
        Remove-Item -LiteralPath $ownedDocumentDirectory -Recurse -ErrorAction Stop
    }
}
