<#
.SYNOPSIS
Maps an OfficeIMO release to the existing date-based VS Code extension version line.
.DESCRIPTION
Accepts a coordinated three-part version tag or a historical timestamp tag.
For a coordinated tag, the source commit time provides a repeatable Marketplace
version above the existing year-based versions. SourceTime permits an explicit
source timestamp when preparing release artifacts outside a Git checkout.
.EXAMPLE
./Build/Get-MarkupExtensionReleaseVersion.ps1 -TagName OfficeIMO-v3.5.2
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^OfficeIMO-v(?:\d{14}|\d+\.\d+\.\d+)$')]
    [string] $TagName,
    [DateTimeOffset] $SourceTime
)

$ErrorActionPreference = 'Stop'
if ($TagName -match '^OfficeIMO-v(?<stamp>\d{14})$') {
    $sourceDate = [DateTimeOffset]::ParseExact($Matches.stamp, 'yyyyMMddHHmmss',
        [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal)
} elseif ($PSBoundParameters.ContainsKey('SourceTime')) {
    $sourceDate = $SourceTime.ToUniversalTime()
} else {
    $stamp = & git -C "$PSScriptRoot/.." show -s --format=%cI HEAD
    if ($LASTEXITCODE -ne 0 -or -not $stamp) { throw 'Cannot identify the extension release source time.' }
    $sourceDate = [DateTimeOffset]::Parse($stamp, [Globalization.CultureInfo]::InvariantCulture).ToUniversalTime()
}

'{0}.{1}.{2}' -f $sourceDate.Year, [int]$sourceDate.ToString('MMdd'), [int]$sourceDate.ToString('HHmmss')
