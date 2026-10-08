param([Parameter(Mandatory)][string] $EvidenceDirectory)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$output = [IO.Path]::GetFullPath($EvidenceDirectory)
New-Item -ItemType Directory -Path $output -Force | Out-Null
$project = Join-Path $repository 'OfficeIMO.Browser/OfficeIMO.Browser.csproj'
[xml] $metadata = Get-Content -LiteralPath $project -Raw
$version = $metadata.Project.PropertyGroup.VersionPrefix
& dotnet pack $project -c Release --nologo -o $output
if ($LASTEXITCODE -ne 0) { throw 'Asset package build/pack failed.' }
$consumer = Join-Path $output 'consumer'
New-Item -ItemType Directory -Path $consumer -Force | Out-Null
$consumerProject = Join-Path $consumer 'Consumer.csproj'
@"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup><OutputType>Exe</OutputType><TargetFrameworks>net8.0;net10.0</TargetFrameworks><ImplicitUsings>enable</ImplicitUsings></PropertyGroup>
  <ItemGroup><PackageReference Include="OfficeIMO.Browser" Version="$version" /></ItemGroup>
</Project>
"@ | Set-Content -LiteralPath $consumerProject -Encoding utf8
@'
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Browser;
foreach (var asset in new[] { BrowserAssets.Script, BrowserAssets.Module, BrowserAssets.XlsxScript,
    BrowserAssets.XlsxModule, BrowserAssets.CsvScript, BrowserAssets.CsvModule, BrowserAssets.PdfScript, BrowserAssets.PdfModule, BrowserAssets.DataTablesScript, BrowserAssets.DataTablesModule }) {
    var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(asset.Content))).ToLowerInvariant()[..16];
    if (asset.ContentHash != hash || !asset.HashedFileName.Contains(hash) || asset.Content.Length == 0)
        throw new InvalidDataException("Packed embedded asset content/hash differs.");
    if (!Encoding.UTF8.GetBytes(asset.Content).SequenceEqual(File.ReadAllBytes(Path.Combine(args[0], asset.FileName))))
        throw new InvalidDataException("Packed asset differs from the current generated source.");
}
Console.WriteLine("Packed .NET consumer verified all ten assets and their SHA-256 names.");
'@ | Set-Content -LiteralPath (Join-Path $consumer 'Program.cs') -Encoding utf8
# A task-local packages directory ensures this tests the archive, not a cached package.
$packageCache = Join-Path $consumer ('packages/' + [Guid]::NewGuid().ToString('N'))
& dotnet restore $consumerProject --source $output --packages $packageCache --force --nologo
if ($LASTEXITCODE -ne 0) { throw 'Isolated asset package restore failed.' }
foreach ($framework in 'net8.0','net10.0') {
    & dotnet run --project $consumerProject -c Release -f $framework --no-restore -- (Join-Path $repository 'OfficeIMO.JavaScript/bundles')
    if ($LASTEXITCODE -ne 0) { throw "Packed consumer failed: $framework" }
}
Write-Host "Packed .NET asset consumers passed: OfficeIMO.Browser.$version.nupkg"
