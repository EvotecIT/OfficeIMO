[CmdletBinding()]
param(
    [ValidateSet('all', 'macos', 'iphone', 'ipad', 'duo')][string] $Role = 'all',
    [string] $WorkRoot,
    [ValidateSet('closed', 'open')][string] $DuoPose,
    # Listing locales from store-art/locales.json that -Role all captures; all of them by default.
    [string[]] $Locales,
    # The one listing locale a single-role run captures.
    [string] $Locale = 'en-US'
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$sourceRoot = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
$project = Join-Path $sourceRoot 'Apple/OfficeIMOStudio/OfficeIMOStudio.xcodeproj'
$captureRoot = Join-Path $sourceRoot 'Artifacts/Studio/Apple/screenshots'
# Public samples from the website: the Northwind dashboard, a fillable form and an Excel workbook.
$inputs = [ordered]@{
    STUDIO_CAPTURE_DOCUMENT_PATH = @{
        Source = 'Website/static/downloads/showcase/pdf/showcase-dashboard.pdf'; Name = 'Northwind Q3 Review.pdf'
        Test = 'testCapturePublicDashboard'
    }
    STUDIO_CAPTURE_FORM_PATH = @{
        Source = 'Website/static/downloads/showcase/workflows/pdf-equipment-inspection/example.pdf'; Name = 'Equipment Inspection.pdf'
        Test = 'testCaptureFormDocument'
    }
    STUDIO_CAPTURE_OFFICE_PATH = @{
        Source = 'Website/static/downloads/showcase/features/excel-charts/quarterly-sales.xlsx'; Name = 'Quarterly Sales.xlsx'
        Test = 'testCaptureOfficeViewer'
    }
}
$localeCatalog = @((Get-Content -LiteralPath (Join-Path $PSScriptRoot 'store-art/locales.json') -Raw | ConvertFrom-Json).locales)
# The scenes each role keeps, in store order. Compact iPhone chrome shows Pages; iPad and Mac show search and the inspector.
$scenes = @{
    iphone = @('dashboard', 'markup', 'form', 'office', 'pages')
    duo = @('dashboard', 'markup', 'pages')
    ipad = @('dashboard', 'search', 'markup', 'form', 'office', 'inspector')
    macos = @('dashboard', 'search', 'markup', 'form', 'office', 'inspector')
}
# Pixel sizes the App Store screenshot sets accept for each raw capture.
$captureSizes = @{ iphone = @('1320x2868', '1290x2796'); ipad = @('2064x2752', '2048x2732'); macos = @('2560x1600') }

function Get-CaptureLocale([string] $Store) {
    $entry = @($localeCatalog | Where-Object { $_.store -ceq $Store })
    if ($entry.Count -ne 1) { throw "Unknown listing locale '$Store'; see store-art/locales.json." }
    return $entry[0]
}

function Get-PngSize([string] $Path) {
    $header = [IO.File]::ReadAllBytes($Path)[16..23]
    '{0}x{1}' -f [Buffers.Binary.BinaryPrimitives]::ReadUInt32BigEndian([byte[]]$header[0..3]),
        [Buffers.Binary.BinaryPrimitives]::ReadUInt32BigEndian([byte[]]$header[4..7])
}

function Set-SimulatorLocale([string] $Device, [string] $Language, [string] $Region) {
    # Sets a shut-down simulator's system language and region, so system text in the
    # captures (the iPad status bar date, for example) matches the listing locale.
    # PowerForge boots the device afterwards; this never boots or changes a running device.
    $state = @((& xcrun simctl list devices -j | ConvertFrom-Json -AsHashtable).devices.Values |
        ForEach-Object { $_ } | Where-Object { $_.udid -ceq $Device })
    if ($LASTEXITCODE -ne 0 -or $state.Count -ne 1) { throw "Unknown simulator: $Device" }
    if ($state[0].state -cne 'Shutdown') { throw "Set the locale of $Device while it is shut down (now $($state[0].state))." }
    $preferences = Join-Path $HOME "Library/Developer/CoreSimulator/Devices/$Device/data/Library/Preferences/.GlobalPreferences.plist"
    $null = New-Item -ItemType Directory -Path (Split-Path -Parent $preferences) -Force
    if (-not (Test-Path -LiteralPath $preferences)) {
        & plutil -create binary1 $preferences
        if ($LASTEXITCODE -ne 0) { throw "Cannot create the preferences of $Device." }
    }
    & plutil -replace AppleLanguages -json "[`"$Language`"]" $preferences
    if ($LASTEXITCODE -ne 0) { throw "Cannot set the language of $Device." }
    & plutil -replace AppleLocale -string $Region $preferences
    if ($LASTEXITCODE -ne 0) { throw "Cannot set the region of $Device." }
}

function Assert-ScreenshotContent([string] $Path) {
    # Rejects blank or flat images before they reach store art.
    & xcrun swift (Join-Path $PSScriptRoot 'verify-screenshot-content.swift') $Path
    if ($LASTEXITCODE -ne 0) { throw "The capture $Path looks blank." }
}

function Invoke-CaptureXcode([string[]] $Arguments, [string] $LogPath, [int] $TimeoutSeconds = 1800) {
    # The existing pinned PowerForge runner owns the process group and timeout,
    # including Xcode teardown. Return here so the role's finally can close its app.
    $validation = @{
        SchemaVersion = 1
        Commands = @(@{
            Name = 'Studio native capture'
            FileName = '/bin/sh'
            Arguments = @('-c', 'exec xcodebuild "$@" > "$STUDIO_CAPTURE_XCODE_LOG" 2>&1', 'studio-capture') + $Arguments
            WorkingDirectory = $sourceRoot
            TimeoutSeconds = $TimeoutSeconds
            Environment = @{ STUDIO_CAPTURE_XCODE_LOG = $LogPath }
        })
    }
    $config = "$LogPath.validation.json"
    $validation | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $config -Encoding utf8
    & $env:POWERFORGE_DOTNET $env:POWERFORGE_SIMULATOR_CLI validate-release --config $config --project-root $sourceRoot --output json `
        > "$LogPath.validation-result.json"
    if ($LASTEXITCODE -ne 0) { throw "Native capture failed or timed out. See $LogPath and its validation-result.json." }
}

foreach ($name in 'POWERFORGE_DOTNET', 'POWERFORGE_SIMULATOR_CLI') {
    if (-not [Environment]::GetEnvironmentVariable($name)) { throw "The pinned capture workflow must supply $name." }
}

if ($Role -eq 'duo') {
    if (-not $DuoPose) { throw 'Choose -DuoPose closed (outer display) or open (inner display).' }
    # App Store Connect's iPhone Duo sizes: the outer display, and the inner display in portrait.
    $duoSet = if ($DuoPose -eq 'closed') { 'duo-outer-1398x2034' } else { 'duo-inner-2007x2853' }
}

if ($Role -eq 'duo' -and -not $env:POWERFORGE_DUO_POSE) {
    # One pose per run. simctl can't fold the device, so PowerForge confirms the pose before the
    # capture starts: it folds with POWERFORGE_DUO_FOLD_TOOL when set, and otherwise asks on the
    # terminal to set it in Device Hub.
    if (-not $IsMacOS) { throw 'Native store captures require macOS and Xcode.' }
    if (-not $env:POWERFORGE_CAPTURE_SIMULATORS) { throw 'Select the owned iPhone Duo simulator UUID as {"duo":"<UUID>"}.' }
    $devices = $env:POWERFORGE_CAPTURE_SIMULATORS | ConvertFrom-Json -AsHashtable
    if ($devices.Count -ne 1 -or -not $devices.ContainsKey('duo')) { throw 'Select exactly the owned duo simulator UUID.' }
    if (Test-Path -LiteralPath (Join-Path $captureRoot "Duo/en-US/$duoSet")) {
        throw "Existing reviewed $duoSet captures are preserved; move them before capturing again."
    }
    if (-not $WorkRoot -or -not [IO.Path]::IsPathFullyQualified($WorkRoot) -or
        -not (Test-Path -LiteralPath (Split-Path -Parent $WorkRoot) -PathType Container)) {
        throw 'Pass an absolute capture work folder beneath an existing mounted scratch directory.'
    }
    # The closed and open runs can share one work folder, so the second run reuses the first build.
    $products = Join-Path $WorkRoot 'iOS/Build/Products'
    if (-not (Test-Path -LiteralPath $products -PathType Container)) {
        if (Test-Path -LiteralPath $WorkRoot) { throw 'The capture work folder must be new or hold an earlier Duo build.' }
        $null = New-Item -ItemType Directory -Path $WorkRoot
        Invoke-CaptureXcode -LogPath (Join-Path $WorkRoot 'build-iOS.log') -Arguments @(
            'build-for-testing', '-project', $project, '-scheme', 'OfficeIMOStudioCapture-iOS', '-configuration', 'Release',
            '-destination', 'generic/platform=iOS Simulator', '-derivedDataPath', (Join-Path $WorkRoot 'iOS'),
            '-clonedSourcePackagesDirPath', (Join-Path $WorkRoot 'SourcePackages'), 'CODE_SIGNING_ALLOWED=NO')
    }
    Write-Host "Capture build and diagnostic output: $WorkRoot"
    $receiptRoot = Join-Path $sourceRoot 'build/powerforge/screenshots'
    $null = New-Item -ItemType Directory -Path $receiptRoot -Force
    $owner = if ($env:GITHUB_REPOSITORY -and $env:GITHUB_RUN_ID) {
        "$env:GITHUB_REPOSITORY/$env:GITHUB_RUN_ID/$duoSet"
    } else { "OfficeIMO/local-capture/$duoSet" }
    & $env:POWERFORGE_DOTNET $env:POWERFORGE_SIMULATOR_CLI apple-simulator run `
        --owner $owner --device $devices.duo --require-shutdown `
        --simulator-timeout-seconds 180 --cleanup-timeout-seconds 90 --timeout-seconds 900 `
        --duo-pose $DuoPose --duo-orientation portrait --output json -- `
        pwsh -NoLogo -NoProfile -File $PSCommandPath -Role duo -DuoPose $DuoPose -WorkRoot $WorkRoot `
        > (Join-Path $receiptRoot "$duoSet.json")
    if ($LASTEXITCODE -ne 0) { throw "Capture or simulator release failed for $duoSet; inspect its PowerForge receipt." }
    return
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
    $selected = if ($Locales) { @($Locales | ForEach-Object { $_ -split ',' } | Where-Object { $_ }) } else { @($localeCatalog.store) }
    foreach ($store in $selected) { $null = Get-CaptureLocale $store }
    $null = New-Item -ItemType Directory -Path $WorkRoot, $captureRoot
    Write-Host "Capture build and diagnostic output: $WorkRoot"
    & (Join-Path $PSScriptRoot 'Build-StudioEngine.ps1') -RuntimeIdentifier osx-arm64, osx-x64, iossimulator-arm64 `
        -WorkDirectory (Join-Path $WorkRoot 'engine')
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
    foreach ($store in $selected) { & $PSCommandPath -Role macos -WorkRoot $WorkRoot -Locale $store }
    $receiptRoot = Join-Path $sourceRoot 'build/powerforge/screenshots'
    $null = New-Item -ItemType Directory -Path $receiptRoot -Force
    try {
        foreach ($deviceRole in 'iphone', 'ipad') {
            foreach ($store in $selected) {
                $entry = Get-CaptureLocale $store
                # One session per listing locale: the system language only changes while the device is shut down.
                # The workflow verifies the en-US session as <role>.json; other locales keep their own receipts.
                $session = if ($store -ceq 'en-US') { $deviceRole } else { "$deviceRole-$store" }
                $owner = if ($env:GITHUB_REPOSITORY -and $env:GITHUB_RUN_ID) {
                    "$env:GITHUB_REPOSITORY/$env:GITHUB_RUN_ID/$session"
                } else { "OfficeIMO/local-capture/$session" }
                Set-SimulatorLocale -Device $devices[$deviceRole] -Language $entry.language -Region $entry.region
                # PowerForge alone owns boot, execution bounds, cancellation and release.
                & $env:POWERFORGE_DOTNET $env:POWERFORGE_SIMULATOR_CLI apple-simulator run `
                    --owner $owner --device $devices[$deviceRole] --require-shutdown `
                    --simulator-timeout-seconds 180 --cleanup-timeout-seconds 90 --timeout-seconds 900 --output json -- `
                    pwsh -NoLogo -NoProfile -File $PSCommandPath -Role $deviceRole -WorkRoot $WorkRoot -Locale $store `
                    > (Join-Path $receiptRoot "$session.json")
                if ($LASTEXITCODE -ne 0) { throw "Capture or simulator release failed for $session; inspect its PowerForge receipt." }
            }
        }
    } finally {
        # Leave the shared capture devices in English for later runs.
        foreach ($deviceRole in 'iphone', 'ipad') { Set-SimulatorLocale -Device $devices[$deviceRole] -Language en -Region en_US }
    }

    # Stage the store story for the captured locales: captions, background and a drawn
    # device around the untouched captured pixels. The screenshot configs upload these sets.
    & (Join-Path $PSScriptRoot 'Invoke-PinnedPowerForge.ps1') -Tool Web -ArgumentList @(
        'apple-store-art', 'compose', '--config', (Join-Path $PSScriptRoot 'store-art/store-art.json'),
        '--out', (Join-Path $captureRoot 'store-art'), '--locales', ($selected -join ','))
    if ($LASTEXITCODE -ne 0) { throw 'Store art composition failed.' }
    if ($selected -contains 'en-US') {
        # Reviewable website assets from the same English pixels; they stay in the capture until promoted.
        $candidate = Join-Path $captureRoot 'website-candidate'
        & /bin/bash (Join-Path $PSScriptRoot 'website-screenshot-assets.sh') $captureRoot (Join-Path $candidate 'screenshots')
        if ($LASTEXITCODE -ne 0) { throw 'Website screenshot candidates failed.' }
        & xcrun swift (Join-Path $PSScriptRoot 'screenshot-contact-sheet.swift') (Join-Path $candidate 'screenshots') (Join-Path $candidate 'contact-sheet.png')
        if ($LASTEXITCODE -ne 0) { throw 'The website contact sheet failed.' }
    }
    return
}

if (-not $WorkRoot -or -not (Test-Path -LiteralPath $WorkRoot -PathType Container)) { throw 'A prepared capture WorkRoot is required.' }
$platform = if ($Role -eq 'macos') { 'macOS' } else { 'iOS' }
$family = switch ($Role) { 'macos' { 'macOS' }; 'iphone' { 'iPhone' }; 'ipad' { 'iPad' }; 'duo' { 'Duo' } }
if ($Role -eq 'duo' -and ($env:POWERFORGE_DUO_POSE -cne $DuoPose -or -not $duoSet.EndsWith("-$env:POWERFORGE_DUO_SIZE"))) {
    throw "PowerForge confirmed the $env:POWERFORGE_DUO_POSE pose at $env:POWERFORGE_DUO_SIZE, not $duoSet."
}
# The closed and open Duo runs share a work folder, so each keeps its own results.
$captureLocale = Get-CaptureLocale $Locale
if ($Role -eq 'duo' -and $Locale -cne 'en-US') { throw 'Duo captures are English only.' }
$runName = if ($Role -eq 'duo') { $duoSet } else { "$Role-$Locale" }
$products = Join-Path $WorkRoot "$platform/Build/Products"
$target = "OfficeIMOStudioCapture-$platform"
$original = @(Get-ChildItem -LiteralPath $products -Filter "$($target)_*.xctestrun")
if ($original.Count -ne 1) { throw 'Expected exactly one capture test manifest from the native build.' }
$appPath = Join-Path $products $(if ($Role -eq 'macos') { 'Release/OfficeIMO Studio.app' } else { 'Release-iphonesimulator/OfficeIMO Studio.app' })
$appPath = (& /bin/realpath $appPath).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Cannot resolve the exact native capture app.' }
$documents = [ordered]@{}
$ownedDocumentDirectory = $null
$ownedMacFixtureDirectory = $null
$ownedMacProcess = $null
$preserveMacInput = $false
$macExecutable = Join-Path $appPath 'Contents/MacOS/OfficeIMO Studio'
$macTemp = Join-Path $HOME 'Library/Containers/com.evotec.officeimo.studio/Data/tmp'
$destination = 'platform=macOS'
$testNames = @($inputs.Values | ForEach-Object { $_.Test })
if ($Role -eq 'duo') { $testNames = @('testCapturePublicDashboard') }

function Start-MacCaptureApp([string] $Document) {
    if (@(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Path -ceq $macExecutable }).Count) {
        throw 'The capture app is already running; preserve that session.'
    }
    & /usr/bin/open -n -a $appPath --env "TMPDIR=$macTemp" $Document `
        --args -AppleLanguages "($($captureLocale.language))" -AppleLocale $captureLocale.region -ApplePersistenceIgnoreState YES
    if ($LASTEXITCODE -ne 0) { throw 'Cannot launch the exact Mac capture app with its public input.' }
    $launchDeadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
        $launched = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Path -ceq $macExecutable })
        if ($launched.Count -eq 1) {
            $script:ownedMacProcess = @{ Id = $launched[0].Id; StartTime = $launched[0].StartTime }
            return
        }
        if ($launched.Count -gt 1 -or [DateTime]::UtcNow -ge $launchDeadline) {
            $script:preserveMacInput = $true
            throw 'Cannot uniquely identify the launched Mac capture process; preserve its input and inspect the session.'
        }
        Start-Sleep -Milliseconds 100
    } while ($true)
}

function Stop-MacCaptureApp {
    if (-not $script:ownedMacProcess) { return }
    $process = Get-Process -Id $script:ownedMacProcess.Id -ErrorAction SilentlyContinue
    if ($process -and $process.Path -ceq $macExecutable -and $process.StartTime -eq $script:ownedMacProcess.StartTime) {
        & /bin/kill -TERM $process.Id
        if ($LASTEXITCODE -ne 0 -or -not $process.WaitForExit(5000)) {
            $script:preserveMacInput = $true
            throw 'The owned Mac capture app did not terminate; preserve its input for inspection.'
        }
    }
    $script:ownedMacProcess = $null
}

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
        foreach ($key in $inputs.Keys) {
            $documents[$key] = Join-Path $ownedMacFixtureDirectory $inputs[$key].Name
            Copy-Item -LiteralPath (Join-Path $sourceRoot $inputs[$key].Source) -Destination $documents[$key]
        }
    } else {
        $devices = $env:POWERFORGE_CAPTURE_SIMULATORS | ConvertFrom-Json -AsHashtable
        $device = [string]$devices[$Role]
        $destination = "platform=iOS Simulator,id=$device"
        & xcrun simctl install $device $appPath
        if ($LASTEXITCODE -ne 0) { throw 'Cannot install the exact capture build.' }
        $container = (& xcrun simctl get_app_container $device com.evotec.officeimo.studio data).Trim()
        if ($LASTEXITCODE -ne 0) { throw 'Cannot locate the capture app document container.' }
        $ownedDocumentDirectory = Join-Path $container ('Documents/00 Studio Capture-' + [Guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $ownedDocumentDirectory
        foreach ($key in $inputs.Keys) {
            $documents[$key] = Join-Path $ownedDocumentDirectory $inputs[$key].Name
            Copy-Item -LiteralPath (Join-Path $sourceRoot $inputs[$key].Source) -Destination $documents[$key]
        }
    }
    $manifest = (& plutil -convert json -o - $original[0].FullName) | ConvertFrom-Json -AsHashtable
    if ($LASTEXITCODE -ne 0 -or -not $manifest.ContainsKey($target)) { throw 'Unexpected Xcode capture manifest.' }
    $test = $manifest[$target]
    # Keep the capture runner and normal app launch in the listing's language and region.
    $test.TestLanguage = $captureLocale.language
    $test.TestRegion = ($captureLocale.region -split '_')[-1]
    # The exact app is provided by this driver. Do not replace an iOS document
    # container or the Mac instance already opened with its language and PDF.
    $test.UseUITargetAppProvidedByTests = $true
    $test.Remove('UITargetAppPath')
    $test.DependentProductPaths = @($test.DependentProductPaths | Where-Object {
        [IO.Path]::GetFileName($_) -ne [IO.Path]::GetFileName($appPath)
    })
    if (-not $test.ContainsKey('EnvironmentVariables')) { $test.EnvironmentVariables = @{} }
    foreach ($key in $documents.Keys) { $test.EnvironmentVariables[$key] = $documents[$key] }
    $test.EnvironmentVariables.STUDIO_CAPTURE_ROLE = $Role
    $test.EnvironmentVariables.STUDIO_CAPTURE_APP_PATH = $appPath
    $test.EnvironmentVariables.STUDIO_CAPTURE_LANGUAGE = $captureLocale.language
    $test.EnvironmentVariables.STUDIO_CAPTURE_REGION = $captureLocale.region
    if ($Role -eq 'macos') {
        # Native app scratch belongs in its own sandbox, not the development
        # process's external TMPDIR. This does not change machine settings.
        $test.EnvironmentVariables.STUDIO_CAPTURE_NATIVE_TMPDIR = $macTemp
    }
    $testManifest = Join-Path $products "StudioStoreCapture-$runName.xctestrun"
    $manifest | ConvertTo-Json -Depth 40 | & plutil -convert xml1 -o $testManifest -
    if ($LASTEXITCODE -ne 0) { throw 'Cannot write the capture test environment.' }
    # Mac captures include the cold document accessibility tree and several scenes per
    # window. The per-image render deadline stays unchanged.
    $testAllowance = if ($Role -eq 'macos') { '300' } else { '180' }
    $maximumAllowance = if ($Role -eq 'macos') { '360' } else { '240' }
    $processTimeout = if ($Role -eq 'macos') { 420 } else { 900 }
    function Invoke-CaptureTests([string] $Name, [string[]] $Tests) {
        $result = Join-Path $WorkRoot "$Name.xcresult"
        $selection = @($Tests | ForEach-Object { '-only-testing', "$target/StudioStoreCaptureTests/$_" })
        Invoke-CaptureXcode -LogPath (Join-Path $WorkRoot "capture-$Name.log") -TimeoutSeconds $processTimeout -Arguments (@(
            'test-without-building', '-xctestrun', $testManifest, '-destination', $destination, '-resultBundlePath', $result) + $selection + @(
            '-parallel-testing-enabled', 'NO', '-collect-test-diagnostics', 'never', '-test-timeouts-enabled', 'YES',
            '-default-test-execution-time-allowance', $testAllowance, '-maximum-test-execution-time-allowance', $maximumAllowance))
        return $result
    }
    $results = @()
    if ($Role -eq 'macos') {
        # The host opens each public input in its own launch, in the listing's language.
        foreach ($key in $inputs.Keys) {
            Start-MacCaptureApp $documents[$key]
            try { $results += Invoke-CaptureTests "$runName-$($inputs[$key].Test)" @($inputs[$key].Test) }
            finally { Stop-MacCaptureApp }
        }
    } else {
        $results += Invoke-CaptureTests $runName $testNames
    }
    if ($Role -ne 'macos') {
        $after = (& xcrun simctl get_app_container $device com.evotec.officeimo.studio data).Trim()
        if ($LASTEXITCODE -ne 0 -or $after -ne $container -or @($documents.Values | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) }).Count) {
            throw 'The capture app document container changed during the native test.'
        }
    }
    $attachments = @()
    foreach ($result in $results) {
        $export = $result -replace '\.xcresult$', '-attachments'
        & xcrun xcresulttool export attachments --path $result --output-path $export
        if ($LASTEXITCODE -ne 0) { throw 'Cannot export the native capture attachments.' }
        $tests = Get-Content -LiteralPath (Join-Path $export 'manifest.json') -Raw | ConvertFrom-Json
        $attachments += @($tests | ForEach-Object { $_.attachments } | Where-Object { -not $_.isAssociatedWithFailure } |
            ForEach-Object { $_ | Add-Member -NotePropertyName ExportRoot -NotePropertyValue $export -PassThru })
    }
    $names = @($scenes[$Role] | ForEach-Object { "studio-$_-$Role" })
    $output = Join-Path $captureRoot "$family/$Locale"
    if ($Role -eq 'duo') { $output = Join-Path $output $duoSet }
    $null = New-Item -ItemType Directory -Path $output
    foreach ($name in $names) {
        $selected = @($attachments | Where-Object { $_.suggestedHumanReadableName.StartsWith($name + '_', [StringComparison]::Ordinal) })
        if ($selected.Count -ne 1) { throw "Expected one successful native attachment for $name." }
        $png = Join-Path $output "$name.png"
        Copy-Item -LiteralPath (Join-Path $selected[0].ExportRoot $selected[0].exportedFileName) -Destination $png
        Assert-ScreenshotContent $png
        $size = Get-PngSize $png
        if ($Role -eq 'duo') {
            # App Store Connect accepts only the display's exact pixel size.
            if ($size -cne $env:POWERFORGE_DUO_SIZE) { throw "$name is $size, not the $env:POWERFORGE_DUO_SIZE $duoSet size." }
        } elseif ($size -notin $captureSizes[$Role]) {
            throw "$name is $size; the $family screenshot set accepts $($captureSizes[$Role] -join ', ')."
        }
    }
} finally {
    Stop-MacCaptureApp
    if ($ownedMacFixtureDirectory -and -not $preserveMacInput) {
        Remove-Item -LiteralPath $ownedMacFixtureDirectory -Recurse -ErrorAction Stop
    }
    if ($ownedDocumentDirectory) {
        & xcrun simctl terminate $device com.evotec.officeimo.studio 2>$null
        Remove-Item -LiteralPath $ownedDocumentDirectory -Recurse -ErrorAction Stop
    }
}
