<#
.SYNOPSIS
Refreshes the pinned toml-test data used by the TOML conformance tests.

.DESCRIPTION
Clones one release of https://github.com/toml-lang/toml-test, checks that it resolves to the
expected commit, copies the TOML 1.1.0 case list (tests/files-toml-1.1.0) and the files it names,
and rewrites toml-test/manifest.json. Pass a new -Tag and -Commit together to move the pin.
#>
[CmdletBinding()]
param(
    [string] $Tag = 'v2.2.0',
    [string] $Commit = 'ce08da1ddb075d1c7596d663c7fcba9a2ae02c5c'
)

$ErrorActionPreference = 'Stop'
$repository = 'https://github.com/toml-lang/toml-test'
$target = Join-Path $PSScriptRoot 'toml-test'
$work = Join-Path ([IO.Path]::GetTempPath()) ("toml-test-" + [guid]::NewGuid().ToString('N'))

try {
    git -c advice.detachedHead=false clone --quiet --depth 1 --branch $Tag $repository $work
    if ($LASTEXITCODE -ne 0) { throw "git clone failed for $Tag." }
    $head = (git -C $work rev-parse HEAD).Trim()
    if ($head -ne $Commit) { throw "Tag $Tag resolves to $head, expected $Commit." }
    $commitDate = (git -C $work log -1 --format=%cI).Trim()

    if (Test-Path $target) { Remove-Item $target -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $target | Out-Null

    $listFile = Join-Path $work 'tests/files-toml-1.1.0'
    $list = @(Get-Content -LiteralPath $listFile | Where-Object { $_ -and -not $_.StartsWith('#') })
    foreach ($relative in $list) {
        $source = Join-Path $work (Join-Path 'tests' $relative)
        $destination = Join-Path $target $relative
        New-Item -ItemType Directory -Force -Path (Split-Path $destination) | Out-Null
        Copy-Item -LiteralPath $source -Destination $destination
    }
    Copy-Item -LiteralPath $listFile -Destination (Join-Path $target 'files-toml-1.1.0')
    Copy-Item -LiteralPath (Join-Path $work 'LICENSE') -Destination (Join-Path $target 'LICENSE')

    $manifest = [ordered]@{
        source        = $repository
        ref           = $Tag
        commit        = $Commit
        commitDate    = $commitDate
        license       = 'MIT'
        licenseSource = "$repository/blob/$Commit/LICENSE"
        tomlVersion   = '1.1.0'
        caseList      = 'files-toml-1.1.0'
        validCases    = @($list | Where-Object { $_ -like 'valid/*.toml' }).Count
        invalidCases  = @($list | Where-Object { $_ -like 'invalid/*.toml' }).Count
    }
    $json = ($manifest | ConvertTo-Json) -replace "`r`n", "`n"
    [IO.File]::WriteAllText((Join-Path $target 'manifest.json'), $json + "`n", [Text.UTF8Encoding]::new($false))
    Write-Host "Imported $($manifest.validCases) valid and $($manifest.invalidCases) invalid cases at $Commit."
}
finally {
    if (Test-Path $work) { Remove-Item $work -Recurse -Force }
}
