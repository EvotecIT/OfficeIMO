<#
.SYNOPSIS
Refreshes the pinned yaml-test-suite data used by the YAML conformance harness.

.DESCRIPTION
Clones one dated data release of https://github.com/yaml/yaml-test-suite, checks that it
resolves to the expected commit, copies only the files the harness reads and rewrites
yaml-test-suite/manifest.json. Pass a new -Tag and -Commit together to move the pin.
#>
[CmdletBinding()]
param(
    [string] $Tag = 'data-2022-01-17',
    [string] $Commit = '6e6c296ae9c9d2d5c4134b4b64d01b29ac19ff6f',
    # The data branch carries no license file; it is taken from the matching source release.
    [string] $LicenseCommit = '45db50aecf9b1520f8258938c88f396e96f30831'
)

$ErrorActionPreference = 'Stop'
$repository = 'https://github.com/yaml/yaml-test-suite'
$target = Join-Path $PSScriptRoot 'yaml-test-suite'
$keep = @('===', 'in.yaml', 'test.event', 'in.json', 'out.yaml', 'emit.yaml', 'error')
$work = Join-Path ([IO.Path]::GetTempPath()) ("yaml-test-suite-" + [guid]::NewGuid().ToString('N'))

try {
    git clone --quiet --depth 1 --branch $Tag $repository $work
    if ($LASTEXITCODE -ne 0) { throw "git clone failed for $Tag." }
    $head = (git -C $work rev-parse HEAD).Trim()
    if ($head -ne $Commit) { throw "Tag $Tag resolves to $head, expected $Commit." }
    $commitDate = (git -C $work log -1 --format=%cI).Trim()

    if (Test-Path $target) {
        Get-ChildItem $target -Directory | Remove-Item -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $target | Out-Null

    # Case folders are four-character IDs; 'name' and 'tags' are symlink indexes.
    $cases = Get-ChildItem $work -Directory | Where-Object { $_.Name -cmatch '^[0-9A-Z]{4}$' } | Sort-Object Name
    foreach ($case in $cases) {
        $sources = @($case) + @(Get-ChildItem $case.FullName -Directory | Where-Object { $_.Name -cmatch '^[0-9]+$' })
        foreach ($source in $sources) {
            $relative = [IO.Path]::GetRelativePath($work, $source.FullName)
            $destination = Join-Path $target $relative
            foreach ($name in $keep) {
                $file = Join-Path $source.FullName $name
                if (Test-Path $file -PathType Leaf) {
                    New-Item -ItemType Directory -Force -Path $destination | Out-Null
                    Copy-Item -LiteralPath $file -Destination (Join-Path $destination $name)
                }
            }
        }
    }

    $licenseUrl = "https://raw.githubusercontent.com/yaml/yaml-test-suite/$LicenseCommit/License"
    Invoke-WebRequest -Uri $licenseUrl -OutFile (Join-Path $target 'LICENSE')

    $inputs = @(Get-ChildItem $target -Recurse -Filter 'in.yaml')
    $manifest = [ordered]@{
        source        = $repository
        ref           = $Tag
        commit        = $Commit
        commitDate    = $commitDate
        license       = 'MIT'
        licenseSource = $licenseUrl
        files         = $keep
        caseFolders   = $cases.Count
        caseCount     = $inputs.Count
        errorCases    = @(Get-ChildItem $target -Recurse -Filter 'error').Count
        jsonCases     = @(Get-ChildItem $target -Recurse -Filter 'in.json').Count
    }
    $json = ($manifest | ConvertTo-Json) -replace "`r`n", "`n"
    [IO.File]::WriteAllText((Join-Path $target 'manifest.json'), $json + "`n", [Text.UTF8Encoding]::new($false))
    Write-Host "Imported $($inputs.Count) cases from $($cases.Count) folders at $Commit."
}
finally {
    if (Test-Path $work) { Remove-Item $work -Recurse -Force }
}
