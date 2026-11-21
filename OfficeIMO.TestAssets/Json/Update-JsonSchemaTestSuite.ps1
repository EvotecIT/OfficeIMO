<#
.SYNOPSIS
Refreshes the pinned JSON-Schema-Test-Suite data used by the JSON Schema conformance tests.

.DESCRIPTION
Clones https://github.com/json-schema-org/JSON-Schema-Test-Suite, checks out the pinned commit, copies the
validation test cases (required and optional, without proposals) with the licence, and rewrites
JSON-Schema-Test-Suite/manifest.json. Pass a new -Commit to move the pin.
#>
[CmdletBinding()]
param(
    [string] $Commit = '38dea932b02986f496345bf007108277777630bf'
)

$ErrorActionPreference = 'Stop'
$repository = 'https://github.com/json-schema-org/JSON-Schema-Test-Suite'
$target = Join-Path $PSScriptRoot 'JSON-Schema-Test-Suite'
$work = Join-Path ([IO.Path]::GetTempPath()) ("json-schema-test-suite-" + [guid]::NewGuid().ToString('N'))

try {
    git clone --quiet $repository $work
    if ($LASTEXITCODE -ne 0) { throw "git clone failed for $repository." }
    git -C $work checkout --quiet $Commit
    if ($LASTEXITCODE -ne 0) { throw "Commit $Commit was not found." }
    $commitDate = (git -C $work log -1 --format=%cI).Trim()

    if (Test-Path $target) { Remove-Item $target -Recurse -Force }
    $tests = Join-Path $target 'tests'
    New-Item -ItemType Directory -Force -Path (Join-Path $tests 'optional') | Out-Null
    $source = Join-Path $work 'validation/tests'
    Get-ChildItem $source -File -Filter '*.json' | Copy-Item -Destination $tests
    Get-ChildItem (Join-Path $source 'optional') -File -Filter '*.json' | Copy-Item -Destination (Join-Path $tests 'optional')
    Copy-Item (Join-Path $work 'LICENSE') (Join-Path $target 'LICENSE')

    $manifest = [ordered]@{
        source     = $repository
        commit     = $Commit
        commitDate = $commitDate
        license    = 'MIT'
        path       = 'validation/tests'
        files      = @(Get-ChildItem $tests -Recurse -File -Filter '*.json').Count
    }
    $json = ($manifest | ConvertTo-Json) -replace "`r`n", "`n"
    [IO.File]::WriteAllText((Join-Path $target 'manifest.json'), $json + "`n", [Text.UTF8Encoding]::new($false))
    Write-Host "Imported $($manifest.files) test files at $Commit."
}
finally {
    if (Test-Path $work) { Remove-Item $work -Recurse -Force }
}
