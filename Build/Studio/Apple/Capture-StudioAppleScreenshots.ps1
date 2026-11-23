[CmdletBinding()]
param(
    [ValidateSet('all', 'macos', 'iphone', 'ipad', 'duo')][string] $Role = 'all',
    [string] $WorkRoot,
    [ValidateSet('closed', 'open')][string] $DuoPose
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$sourceRoot = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
$project = Join-Path $sourceRoot 'Apple/OfficeIMOStudio/OfficeIMOStudio.xcodeproj'
$captureRoot = Join-Path $sourceRoot 'Artifacts/Studio/Apple/screenshots'
$fixture = Join-Path $sourceRoot 'Website/static/downloads/showcase/pdf/showcase-dashboard.pdf'

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
$family = switch ($Role) { 'macos' { 'macOS' }; 'iphone' { 'iPhone' }; 'ipad' { 'iPad' }; 'duo' { 'Duo' } }
if ($Role -eq 'duo' -and ($env:POWERFORGE_DUO_POSE -cne $DuoPose -or -not $duoSet.EndsWith("-$env:POWERFORGE_DUO_SIZE"))) {
    throw "PowerForge confirmed the $env:POWERFORGE_DUO_POSE pose at $env:POWERFORGE_DUO_SIZE, not $duoSet."
}
# The closed and open Duo runs share a work folder, so each keeps its own results.
$runName = if ($Role -eq 'duo') { $duoSet } else { $Role }
$products = Join-Path $WorkRoot "$platform/Build/Products"
$target = "OfficeIMOStudioCapture-$platform"
$original = @(Get-ChildItem -LiteralPath $products -Filter "$($target)_*.xctestrun")
if ($original.Count -ne 1) { throw 'Expected exactly one capture test manifest from the native build.' }
$appPath = Join-Path $products $(if ($Role -eq 'macos') { 'Release/OfficeIMO Studio.app' } else { 'Release-iphonesimulator/OfficeIMO Studio.app' })
$appPath = (& /bin/realpath $appPath).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Cannot resolve the exact native capture app.' }
$document = $fixture
$ownedDocumentDirectory = $null
$ownedMacFixtureDirectory = $null
$ownedMacProcess = $null
$preserveMacInput = $false
$macExecutable = Join-Path $appPath 'Contents/MacOS/OfficeIMO Studio'
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
        $ownedDocumentDirectory = Join-Path $container ('Documents/00 Studio Capture-' + [Guid]::NewGuid().ToString('N'))
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
    # The exact app is provided by this driver. Do not replace an iOS document
    # container or the Mac instance already opened with its language and PDF.
    $test.UseUITargetAppProvidedByTests = $true
    $test.Remove('UITargetAppPath')
    $test.DependentProductPaths = @($test.DependentProductPaths | Where-Object {
        [IO.Path]::GetFileName($_) -ne [IO.Path]::GetFileName($appPath)
    })
    if (-not $test.ContainsKey('EnvironmentVariables')) { $test.EnvironmentVariables = @{} }
    $test.EnvironmentVariables.STUDIO_CAPTURE_DOCUMENT_PATH = $document
    $test.EnvironmentVariables.STUDIO_CAPTURE_ROLE = $Role
    $test.EnvironmentVariables.STUDIO_CAPTURE_APP_PATH = $appPath
    if ($Role -eq 'macos') {
        # Native app scratch belongs in its own sandbox, not the development
        # process's external TMPDIR. This does not change machine settings.
        $test.EnvironmentVariables.STUDIO_CAPTURE_NATIVE_TMPDIR = Join-Path $HOME 'Library/Containers/com.evotec.officeimo.studio/Data/tmp'
    }
    $testManifest = Join-Path $products "StudioStoreCapture-$runName.xctestrun"
    $manifest | ConvertTo-Json -Depth 40 | & plutil -convert xml1 -o $testManifest -
    if ($LASTEXITCODE -ne 0) { throw 'Cannot write the capture test environment.' }
    $result = Join-Path $WorkRoot "$runName.xcresult"
    if ($Role -eq 'macos') {
        if (@(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Path -ceq $macExecutable }).Count) {
            throw 'The capture app is already running; preserve that session.'
        }
        & /usr/bin/open -n -a $appPath --env "TMPDIR=$($test.EnvironmentVariables.STUDIO_CAPTURE_NATIVE_TMPDIR)" $document `
            --args -AppleLanguages '(en)' -AppleLocale en_US -ApplePersistenceIgnoreState YES
        if ($LASTEXITCODE -ne 0) { throw 'Cannot launch the exact Mac capture app with its public input.' }
        $launchDeadline = [DateTime]::UtcNow.AddSeconds(15)
        do {
            $launched = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Path -ceq $macExecutable })
            if ($launched.Count -eq 1) {
                $ownedMacProcess = @{ Id = $launched[0].Id; StartTime = $launched[0].StartTime }
                break
            }
            if ($launched.Count -gt 1 -or [DateTime]::UtcNow -ge $launchDeadline) {
                $preserveMacInput = $true
                throw 'Cannot uniquely identify the launched Mac capture process; preserve its input and inspect the session.'
            }
            Start-Sleep -Milliseconds 100
        } while ($true)
    }
    # Mac captures include the cold document accessibility tree and both
    # dashboard/search windows. The per-image render deadline stays unchanged.
    $testAllowance = if ($Role -eq 'macos') { '300' } else { '120' }
    $maximumAllowance = if ($Role -eq 'macos') { '360' } else { '150' }
    $processTimeout = if ($Role -eq 'macos') { 420 } else { 240 }
    Invoke-CaptureXcode -LogPath (Join-Path $WorkRoot "capture-$runName.log") -TimeoutSeconds $processTimeout -Arguments @(
        'test-without-building', '-xctestrun', $testManifest, '-destination', $destination, '-resultBundlePath', $result,
        '-only-testing', "$target/StudioStoreCaptureTests/testCapturePublicDashboard", '-parallel-testing-enabled', 'NO',
        '-collect-test-diagnostics', 'never', '-test-timeouts-enabled', 'YES',
        '-default-test-execution-time-allowance', $testAllowance, '-maximum-test-execution-time-allowance', $maximumAllowance)
    if ($Role -ne 'macos') {
        $after = (& xcrun simctl get_app_container $device com.evotec.officeimo.studio data).Trim()
        if ($LASTEXITCODE -ne 0 -or $after -ne $container -or -not (Test-Path -LiteralPath $document -PathType Leaf)) {
            throw 'The capture app document container changed during the native test.'
        }
    }
    $export = Join-Path $WorkRoot "$runName-attachments"
    & xcrun xcresulttool export attachments --path $result --output-path $export
    if ($LASTEXITCODE -ne 0) { throw 'Cannot export the native capture attachments.' }
    $tests = Get-Content -LiteralPath (Join-Path $export 'manifest.json') -Raw | ConvertFrom-Json
    $attachments = @($tests | ForEach-Object { $_.attachments } | Where-Object { -not $_.isAssociatedWithFailure })
    $names = @("studio-dashboard-$Role")
    if ($Role -notin 'iphone', 'duo') { $names += "studio-search-$Role" }
    $output = Join-Path $captureRoot "$family/en-US"
    if ($Role -eq 'duo') { $output = Join-Path $output $duoSet }
    $null = New-Item -ItemType Directory -Path $output
    foreach ($name in $names) {
        $selected = @($attachments | Where-Object { $_.suggestedHumanReadableName.StartsWith($name + '_', [StringComparison]::Ordinal) })
        if ($selected.Count -ne 1) { throw "Expected one successful native attachment for $name." }
        Copy-Item -LiteralPath (Join-Path $export $selected[0].exportedFileName) -Destination (Join-Path $output "$name.png")
        if ($Role -eq 'duo') {
            # App Store Connect accepts only the display's exact pixel size.
            $header = [IO.File]::ReadAllBytes((Join-Path $output "$name.png"))[16..23]
            $size = '{0}x{1}' -f [Buffers.Binary.BinaryPrimitives]::ReadUInt32BigEndian([byte[]]$header[0..3]),
                [Buffers.Binary.BinaryPrimitives]::ReadUInt32BigEndian([byte[]]$header[4..7])
            if ($size -cne $env:POWERFORGE_DUO_SIZE) { throw "$name is $size, not the $env:POWERFORGE_DUO_SIZE $duoSet size." }
        }
    }
} finally {
    if ($ownedMacProcess) {
        $process = Get-Process -Id $ownedMacProcess.Id -ErrorAction SilentlyContinue
        if ($process -and $process.Path -ceq $macExecutable -and $process.StartTime -eq $ownedMacProcess.StartTime) {
            & /bin/kill -TERM $process.Id
            if ($LASTEXITCODE -ne 0 -or -not $process.WaitForExit(5000)) {
                throw 'The owned Mac capture app did not terminate; preserve its input for inspection.'
            }
        }
    }
    if ($ownedMacFixtureDirectory -and -not $preserveMacInput) {
        Remove-Item -LiteralPath $ownedMacFixtureDirectory -Recurse -ErrorAction Stop
    }
    if ($ownedDocumentDirectory) {
        & xcrun simctl terminate $device com.evotec.officeimo.studio 2>$null
        Remove-Item -LiteralPath $ownedDocumentDirectory -Recurse -ErrorAction Stop
    }
}
