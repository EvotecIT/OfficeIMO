#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Adds the Apple apps' new interface strings to their String Catalog. Runs on a Mac with Xcode.

.DESCRIPTION
    Xcode adds new strings from the Swift source to Localizable.xcstrings when the project builds in the Xcode app, not
    when xcodebuild builds it. This builds the iOS app for the simulator, which extracts every shared string, and syncs
    the extracted strings into the catalog without marking strings stale. Run it after changing interface text and
    before exporting or checking translations.

.EXAMPLE
    ./Build/Studio/Sync-StudioAppleStrings.ps1 -DerivedDataPath /tmp/studio-derived
#>
[CmdletBinding()]
param([string] $DerivedDataPath = (Join-Path ([System.IO.Path]::GetTempPath()) 'officeimo-studio-strings'))

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
if (-not $IsMacOS) { throw 'Syncing the Apple String Catalog needs a Mac with Xcode.' }
$project = Join-Path $PSScriptRoot '..' '..' 'Apple' 'OfficeIMOStudio'
$catalog = Join-Path $project 'Shared' 'Localizable.xcstrings'

xcodebuild -project (Join-Path $project 'OfficeIMOStudio.xcodeproj') -scheme OfficeIMOStudio-iOS -configuration Debug `
    -destination 'generic/platform=iOS Simulator' -derivedDataPath $DerivedDataPath build -quiet
$extracted = Get-ChildItem -Path (Join-Path $DerivedDataPath 'Build/Intermediates.noindex') -Recurse -Filter '*.stringsdata' |
    Where-Object FullName -like '*Debug-iphonesimulator*'
if (-not $extracted) { throw 'The build extracted no strings.' }
$arguments = @('xcstringstool', 'sync', $catalog) + @($extracted | ForEach-Object { '--stringsdata', $_.FullName }) + '--skip-marking-strings-stale'
xcrun @arguments

Import-Module (Join-Path $PSScriptRoot 'StudioTranslations.psm1') -Force
# Rewrite in the layout the translation tools use, so a later import changes only translations.
Write-AppleCatalog Localizable (Read-AppleCatalog Localizable)
Write-Host 'Localizable.xcstrings has the current strings. Run ./Build/Studio/Test-StudioTranslations.ps1 to see what needs translating.'
