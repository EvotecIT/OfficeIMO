<#
.SYNOPSIS
Builds OfficeIMOStudioEngine.xcframework: the Studio engine as static NativeAOT libraries for macOS (arm64 and x86_64),
iOS and the iOS simulator, with the C header and a Swift module.

.DESCRIPTION
For each runtime the engine is published with StudioEngineLibrary=Static. NativeAOT's static archive holds only the compiled
code, so the project also publishes the toolchain's link inputs (see PublishStaticLinkInputs in the project file). This
script prelinks the compiled code with the runtime bootstrapper, merges it with the runtime libraries and the native
shims into one static library per runtime, combines
the two macOS architectures with lipo, and packages the result with xcodebuild -create-xcframework.

The system libraries and frameworks the runtime needs are written into the module map as link directives, so a Swift
target that imports OfficeIMOStudioEngine links them automatically. The exact inputs of every runtime are kept in
<OutputDirectory>/link-inputs/<runtime>.txt.

Requires macOS with Xcode and the .NET SDK selected by global.json.

.EXAMPLE
./Build-AppleEngine.ps1 -OutputDirectory /tmp/engine
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$OutputDirectory,
    [ValidateSet('Debug', 'Release')][string]$Configuration = 'Release',
    [ValidateSet('osx-arm64', 'osx-x64', 'ios-arm64', 'iossimulator-arm64')]
    [string[]]$RuntimeIdentifier = @('osx-arm64', 'osx-x64', 'ios-arm64', 'iossimulator-arm64'),
    # The oldest systems the native apps support; the prelinked engine object is marked for them.
    [string]$MacOSDeploymentTarget = '26.0',
    [string]$IOSDeploymentTarget = '26.0'
)
$ErrorActionPreference = 'Stop'
if (-not $IsMacOS) { throw 'The XCFramework is built on macOS with Xcode.' }

$project = Join-Path $PSScriptRoot 'OfficeIMO.Studio.Native.csproj'
$name = 'OfficeIMOStudioEngine'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path $OutputDirectory).Path
$publishRoot = Join-Path $OutputDirectory 'publish'
$libraryRoot = Join-Path $OutputDirectory 'lib'
$inputsRoot = Join-Path $OutputDirectory 'link-inputs'
$framework = Join-Path $OutputDirectory "$name.xcframework"
$objectRoot = Join-Path $OutputDirectory 'objects'
foreach ($directory in $publishRoot, $libraryRoot, $inputsRoot, $objectRoot, $framework) {
    if (Test-Path $directory) { Remove-Item -Recurse -Force $directory }
}
New-Item -ItemType Directory -Force $publishRoot, $libraryRoot, $inputsRoot | Out-Null

function Invoke-Tool([string]$Tool, [string[]]$Arguments) {
    & $Tool @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$Tool failed with exit code $LASTEXITCODE." }
}

$systemInputs = [ordered]@{}
foreach ($rid in $RuntimeIdentifier) {
    Write-Host "Publishing $rid"
    $publish = Join-Path $publishRoot $rid
    Invoke-Tool dotnet @('publish', $project, '-c', $Configuration, '-r', $rid, '-p:StudioEngineLibrary=Static',
        '--artifacts-path', (Join-Path $OutputDirectory 'dotnet'), '-o', $publish)

    $linkDirectory = Join-Path $publish 'link'
    $inputs = Get-Content (Join-Path $linkDirectory 'link-inputs.txt') | Where-Object { $_ }
    # Apple SDKs ship the C++ standard library as libc++; NativeAOT names it stdc++ for every Unix target.
    $inputs = $inputs | ForEach-Object { if ($_ -eq '-lstdc++') { '-lc++' } else { $_ } }
    Set-Content (Join-Path $inputsRoot "$rid.txt") (@("$name.a") + $inputs)
    Copy-Item (Join-Path $publish "$name.deps.json") (Join-Path $inputsRoot "$rid.deps.json")

    # Object files such as the bootstrapper export nothing: they start the runtime from a static initializer. As
    # archive members the linker would never load them, so they are prelinked into the object that holds the exports,
    # and linking any engine function brings the runtime start with it. Private external symbols stay visible to the
    # runtime archives, which call into the compiled code.
    $objects = Join-Path $OutputDirectory 'objects' $rid
    if (Test-Path $objects) { Remove-Item -Recurse -Force $objects }
    New-Item -ItemType Directory -Force $objects | Out-Null
    Push-Location $objects
    try { Invoke-Tool xcrun @('ar', '-x', (Join-Path $publish "$name.a")) } finally { Pop-Location }
    $architecture = if ($rid.EndsWith('-x64')) { 'x86_64' } else { 'arm64' }
    # The compiler marks its object for the oldest supported system and the runtime objects for newer ones, so the
    # prelinked object carries the apps' deployment target.
    $platform, $sdk, $minimum = switch -Wildcard ($rid) {
        'osx-*' { 'macos', 'macosx', $MacOSDeploymentTarget }
        'iossimulator-*' { 'ios-simulator', 'iphonesimulator', $IOSDeploymentTarget }
        default { 'ios', 'iphoneos', $IOSDeploymentTarget }
    }
    $sdkVersion = (xcrun --sdk $sdk --show-sdk-version).Trim()
    $files = $inputs | Where-Object { -not $_.StartsWith('-') } | ForEach-Object { Join-Path $linkDirectory $_ }
    $engineObject = Join-Path $OutputDirectory 'objects' "$rid.o"
    $prelinked = @(Get-ChildItem $objects -Filter '*.o' | ForEach-Object FullName) + @($files | Where-Object { $_.EndsWith('.o') })
    Invoke-Tool xcrun (@('ld', '-r', '-keep_private_externs', '-arch', $architecture, '-platform_version', $platform, $minimum, $sdkVersion,
        '-o', $engineObject) + $prelinked)

    $library = Join-Path $libraryRoot $rid "lib$name.a"
    New-Item -ItemType Directory -Force (Split-Path $library) | Out-Null
    $archives = @($files | Where-Object { $_.EndsWith('.a') })
    Invoke-Tool xcrun (@('libtool', '-static', '-no_warning_for_no_symbols', '-o', $library, $engineObject) + $archives)
    $systemInputs[$rid] = @($inputs | Where-Object { $_.StartsWith('-') })
}

# Each XCFramework slice has its own headers and module map. The module map's link directives name the system libraries
# and frameworks, so a Swift target that imports the module links them without further settings. The headers sit in a
# folder named after the module so the module map cannot collide with another library's.
function New-ModuleHeaders([string]$Slice, [string[]]$Runtimes) {
    $distinct = @($Runtimes | ForEach-Object { $systemInputs[$_] -join ' ' } | Sort-Object -Unique)
    if ($distinct.Count -ne 1) { throw "The $Slice runtimes need different system libraries: $($Runtimes -join ', ')" }
    $headers = Join-Path $OutputDirectory 'headers' $Slice
    $moduleHeaders = Join-Path $headers $name
    New-Item -ItemType Directory -Force $moduleHeaders | Out-Null
    Copy-Item (Join-Path $PSScriptRoot 'include/officeimo_studio.h') $moduleHeaders
    $links = $systemInputs[$Runtimes[0]] | ForEach-Object {
        if ($_.StartsWith('-framework ')) { "    link framework `"$($_.Substring(11))`"" } else { "    link `"$($_.Substring(2))`"" }
    }
    Set-Content (Join-Path $moduleHeaders 'module.modulemap') (@("module $name {", '    header "officeimo_studio.h"', '    export *') + $links + '}')
    return $headers
}

$headersRoot = Join-Path $OutputDirectory 'headers'
if (Test-Path $headersRoot) { Remove-Item -Recurse -Force $headersRoot }
$arguments = @('-create-xcframework')
$macOS = @($RuntimeIdentifier | Where-Object { $_.StartsWith('osx-') })
if ($macOS.Count -gt 0) {
    $universal = Join-Path $libraryRoot 'macos' "lib$name.a"
    New-Item -ItemType Directory -Force (Split-Path $universal) | Out-Null
    Invoke-Tool xcrun (@('lipo', '-create', '-output', $universal) + @($macOS | ForEach-Object { Join-Path $libraryRoot $_ "lib$name.a" }))
    $arguments += '-library', $universal, '-headers', (New-ModuleHeaders 'macos' $macOS)
}
foreach ($rid in $RuntimeIdentifier | Where-Object { $_.StartsWith('ios') }) {
    $arguments += '-library', (Join-Path $libraryRoot $rid "lib$name.a"), '-headers', (New-ModuleHeaders $rid @($rid))
}
Invoke-Tool xcodebuild ($arguments + @('-output', $framework))

Write-Host "Built $framework"
Get-ChildItem -Recurse -Filter "lib$name.a" $framework | ForEach-Object {
    Write-Host ("  {0}: {1:N1} MB" -f $_.Directory.Name, ($_.Length / 1MB))
}
