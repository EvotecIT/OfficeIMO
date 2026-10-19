param(
    [string] $ModulePath = $PSScriptRoot,
    [string] $TestsPath = "$PSScriptRoot\Tests",
    [switch] $SkipDependencyInstall
)

$ResolvedModulePath = (Resolve-Path -Path $ModulePath).Path
$PrimaryModule = Get-ChildItem -Path $ResolvedModulePath -Filter '*.psd1' -Recurse -ErrorAction SilentlyContinue -Depth 1
if (-not $PrimaryModule) {
    throw "Path $ResolvedModulePath doesn't contain PSD1 files. Failing tests."
}
if ($PrimaryModule.Count -ne 1) {
    throw 'More than one PSD1 files detected. Failing tests.'
}
$ModuleName = $PrimaryModule.BaseName
$PSDInformation = Import-PowerShellDataFile -Path $PrimaryModule.FullName
$RequiredModules = @(
    'Pester'
    'PSWriteColor'
    if ($PSDInformation.RequiredModules) {
        $PSDInformation.RequiredModules
    }
)
if (-not $SkipDependencyInstall) {
    foreach ($Module in $RequiredModules) {
        if ($Module -eq 'PSWriteOffice') {
            continue
        }
        if ($Module -is [System.Collections.IDictionary]) {
            $Exists = Get-Module -ListAvailable -Name $Module.ModuleName
            if (-not $Exists) {
                Write-Warning "$ModuleName - Downloading $($Module.ModuleName) from PSGallery"
                Install-Module -Name $Module.ModuleName -Force -SkipPublisherCheck
            }
        } else {
            $Exists = Get-Module -ListAvailable $Module -ErrorAction SilentlyContinue
            if (-not $Exists) {
                Install-Module -Name $Module -Force -SkipPublisherCheck
            }
        }
    }
}

Write-Color 'ModuleName: ', $ModuleName, ' Version: ', $PSDInformation.ModuleVersion -Color Yellow, Green, Yellow, Green -LinesBefore 2
Write-Color 'PowerShell Version: ', $PSVersionTable.PSVersion -Color Yellow, Green
Write-Color 'PowerShell Edition: ', $PSVersionTable.PSEdition -Color Yellow, Green
Write-Color 'Required modules: ' -Color Yellow
foreach ($Module in $PSDInformation.RequiredModules) {
    if ($Module -is [System.Collections.IDictionary]) {
        Write-Color '   [>] ', $Module.ModuleName, ' Version: ', $Module.ModuleVersion -Color Yellow, Green, Yellow, Green
    } else {
        Write-Color '   [>] ', $Module -Color Yellow, Green
    }
}
Write-Color

$previousDevelopmentFlag = $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES
$developmentOutputPath = Join-Path $ResolvedModulePath 'Sources/PSWriteOffice/bin/Debug'
if (-not $previousDevelopmentFlag -and (Test-Path -LiteralPath $developmentOutputPath)) {
    $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
}

try {
    Import-Module $PrimaryModule.FullName -Force -Global -ErrorAction Stop
    Import-Module Pester -Force -ErrorAction Stop
} catch {
    throw "Failed to import module $ModuleName"
} finally {
    if ($null -eq $previousDevelopmentFlag) {
        Remove-Item env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES -ErrorAction SilentlyContinue
    } else {
        $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = $previousDevelopmentFlag
    }
}

$env:PSWRITEOFFICE_MODULE_MANIFEST = $PrimaryModule.FullName

$Configuration = [PesterConfiguration]::Default
$Configuration.Run.Path = $TestsPath
$Configuration.Run.Exit = $true
$Configuration.Should.ErrorAction = 'Continue'
$Configuration.CodeCoverage.Enabled = $false
$Configuration.Output.Verbosity = 'Detailed'
$Result = Invoke-Pester -Configuration $Configuration
#$result = Invoke-Pester -Script $PSScriptRoot\Tests -Verbose -Output Detailed #-EnableExit

if ($Result.FailedCount -gt 0) {
    throw "$($Result.FailedCount) tests failed."
}
