<#
.SYNOPSIS
    Compiles the C# examples in the website documentation against the current source.

.DESCRIPTION
    Extracts every ```csharp block from Website/content/docs (or -Path), finds the
    OfficeIMO projects the blocks import, builds them once, and compiles each block
    on its own with Roslyn. A block that uses a type or member that does not exist
    fails with the file and line of the error.

    Blocks compile as standalone programs with the SDK's implicit usings. Put
    <!-- docs-check: skip --> on the line before a fence that is deliberately partial.

.PARAMETER Path
    Markdown folders to check. Defaults to Website/content/docs.

.PARAMETER WorkRoot
    Folder for the generated project and build output. Defaults to a folder under the
    system temp directory.

.PARAMETER Configuration
    Build configuration for the referenced projects. Defaults to Release.

.EXAMPLE
    ./Build/Test-DocsCodeExamples.ps1
#>
[CmdletBinding()]
param(
    [string[]] $Path,
    [string] $WorkRoot,
    [string] $Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $Path) { $Path = @(Join-Path $repoRoot 'Website/content/docs') }
if (-not $WorkRoot) { $WorkRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'officeimo-docs-examples' }
$Path = $Path | ForEach-Object { (Resolve-Path -LiteralPath $_).Path }
New-Item -ItemType Directory -Path $WorkRoot -Force | Out-Null

# Map every imported OfficeIMO namespace to the closest project folder that owns it.
$namespaces = [System.Collections.Generic.HashSet[string]]::new()
foreach ($folder in $Path) {
    foreach ($file in Get-ChildItem -LiteralPath $folder -Filter '*.md' -Recurse -File) {
        foreach ($match in [regex]::Matches((Get-Content -LiteralPath $file.FullName -Raw), '(?m)^\s*using\s+(?:static\s+)?(OfficeIMO(?:\.[A-Za-z0-9]+)*)\s*;')) {
            [void] $namespaces.Add($match.Groups[1].Value)
        }
    }
}
[void] $namespaces.Add('OfficeIMO')

$projects = [System.Collections.Generic.SortedSet[string]]::new([System.StringComparer]::Ordinal)
foreach ($namespace in $namespaces) {
    $candidate = $namespace
    while ($candidate) {
        $project = Join-Path $repoRoot "$candidate/$candidate.csproj"
        if (Test-Path -LiteralPath $project -PathType Leaf) {
            [void] $projects.Add($project)
            $generator = Join-Path $repoRoot "$candidate.SourceGeneration/$candidate.SourceGeneration.csproj"
            if (Test-Path -LiteralPath $generator -PathType Leaf) { [void] $projects.Add($generator) }
            break
        }
        $dot = $candidate.LastIndexOf('.')
        $candidate = if ($dot -gt 0) { $candidate.Substring(0, $dot) } else { $null }
    }
}

$references = ($projects | ForEach-Object { "    <ProjectReference Include=`"$_`" />" }) -join [Environment]::NewLine
$program = Join-Path $PSScriptRoot 'DocsExamples/Program.cs'
$projectFile = Join-Path $WorkRoot 'DocsExamples.csproj'
@"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <IsPackable>false</IsPackable>
    <NoWarn>`$(NoWarn);NU1608;NU1701</NoWarn>
  </PropertyGroup>
  <ItemGroup>
    <Compile Include="$program" />
    <Reference Include="Microsoft.CodeAnalysis" HintPath="`$(MSBuildSDKsPath)/../Roslyn/bincore/Microsoft.CodeAnalysis.dll" Private="true" />
    <Reference Include="Microsoft.CodeAnalysis.CSharp" HintPath="`$(MSBuildSDKsPath)/../Roslyn/bincore/Microsoft.CodeAnalysis.CSharp.dll" Private="true" />
  </ItemGroup>
  <ItemGroup>
$references
  </ItemGroup>
</Project>
"@ | Set-Content -LiteralPath $projectFile -Encoding utf8

# The repository pins the SDK in global.json; build from the repository so it applies.
Push-Location $repoRoot
try {
    Write-Host "Building $($projects.Count) projects for the documentation examples..."
    dotnet build $projectFile -c $Configuration -nologo -v:q -p:TreatWarningsAsErrors=false | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "Building the documentation example host failed (exit code $LASTEXITCODE)." }
    $exe = Join-Path $WorkRoot "bin/$Configuration/net10.0/DocsExamples.dll"
    dotnet $exe @Path
    if ($LASTEXITCODE -ne 0) { throw 'One or more documentation C# examples do not compile against the current source.' }
} finally {
    Pop-Location
}
