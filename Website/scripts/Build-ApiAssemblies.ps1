#Requires -Version 7.0
<#
.SYNOPSIS
  Builds the assemblies that the apidocs steps document.

.DESCRIPTION
  Release builds embed their PDBs (see Directory.Build.props), but apidocs reads
  source links from a .pdb beside each assembly. This build writes a separate
  portable PDB instead. The projects are built through one temporary solution,
  so shared references build once and output stays beside each project.

.PARAMETER Framework
  Target framework to build.

.PARAMETER Configuration
  Build configuration. Defaults to Release.

.PARAMETER SkipMissing
  Skip projects that are not in this checkout instead of failing.

.PARAMETER Project
  Project paths, relative to the current directory.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Mandatory)]
    [string] $Framework,

    [string] $Configuration = 'Release',

    [switch] $SkipMissing,

    [Parameter(Mandatory, ValueFromRemainingArguments)]
    [string[]] $Project
)

$ErrorActionPreference = 'Stop'

$projects = @($Project | ForEach-Object { [IO.Path]::GetFullPath($_, $PWD.Path) } | Select-Object -Unique)
$missing = @($projects | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
if ($missing.Count -gt 0) {
    if (-not $SkipMissing) {
        throw "Build project was not found: $($missing -join ', ')"
    }

    $projects = @($projects | Where-Object { $_ -notin $missing })
}

if ($projects.Count -eq 0) {
    Write-Host 'No existing projects to build.'
    return
}

$directory = Join-Path ([IO.Path]::GetTempPath()) ('officeimo-api-build-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $directory | Out-Null
try {
    $solution = Join-Path $directory 'projects.slnx'
    $entries = $projects | ForEach-Object {
        '  <Project Path="{0}" />' -f [Security.SecurityElement]::Escape($_.Replace('\', '/'))
    }
    Set-Content -LiteralPath $solution -Value (@('<Solution>') + $entries + '</Solution>')

    & dotnet build $solution --configuration $Configuration --framework $Framework '-p:DebugType=portable'
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet build failed with exit code $LASTEXITCODE."
    }
}
finally {
    Remove-Item -LiteralPath $directory -Recurse -Force
}
