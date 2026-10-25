param([Parameter(Mandatory)][string] $EvidenceDirectory)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$output = [IO.Path]::GetFullPath($EvidenceDirectory)
New-Item -ItemType Directory -Path $output -Force | Out-Null
$project = Join-Path $repository 'OfficeIMO.JavaScript.Assets/OfficeIMO.JavaScript.Assets.csproj'
[xml] $metadata = Get-Content -LiteralPath $project -Raw
$version = $metadata.Project.PropertyGroup.VersionPrefix
& dotnet pack $project -c Release --nologo -p:UseSharedCompilation=false -o $output
if ($LASTEXITCODE -ne 0) { throw 'Asset package build/pack failed.' }
$consumer = Join-Path $output 'consumer'
New-Item -ItemType Directory -Path $consumer -Force | Out-Null
$consumerProject = Join-Path $consumer 'Consumer.csproj'
# Keep the external consumer on the repository's SDK even outside its directory tree.
Copy-Item -LiteralPath (Join-Path $repository 'global.json') -Destination $consumer
@"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup><OutputType>Exe</OutputType><TargetFrameworks>net8.0;net10.0</TargetFrameworks><ImplicitUsings>enable</ImplicitUsings></PropertyGroup>
  <ItemGroup><PackageReference Include="OfficeIMO.JavaScript.Assets" Version="$version" /></ItemGroup>
</Project>
"@ | Set-Content -LiteralPath $consumerProject -Encoding utf8
@'
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.JavaScript.Assets;
foreach (var asset in new[] { JavaScriptAssets.Script, JavaScriptAssets.Module, JavaScriptAssets.XlsxScript,
    JavaScriptAssets.XlsxModule, JavaScriptAssets.CsvScript, JavaScriptAssets.CsvModule, JavaScriptAssets.PdfScript, JavaScriptAssets.PdfModule, JavaScriptAssets.DataTablesScript, JavaScriptAssets.DataTablesModule,
    JavaScriptAssets.CanopyXScript, JavaScriptAssets.CanopyXModule, JavaScriptAssets.ExecutionScript, JavaScriptAssets.ExecutionModule }) {
    var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(asset.Content))).ToLowerInvariant()[..16];
    if (asset.ContentHash != hash || !asset.HashedFileName.Contains(hash) || asset.Content.Length == 0)
        throw new InvalidDataException("Packed embedded asset content/hash differs.");
    if (!Encoding.UTF8.GetBytes(asset.Content).SequenceEqual(File.ReadAllBytes(Path.Combine(args[0], asset.FileName))))
        throw new InvalidDataException("Packed asset differs from the current generated source.");
}
Console.WriteLine("Packed .NET consumer verified all fourteen assets and their SHA-256 names.");
'@ | Set-Content -LiteralPath (Join-Path $consumer 'Program.cs') -Encoding utf8
# A task-local packages directory ensures this tests the archive, not a cached package.
$packageCache = Join-Path $consumer ('packages/' + [Guid]::NewGuid().ToString('N'))
# The asset package must come from this archive. Missing SDK reference/host packs
# may come from NuGet; they are not product dependencies and are not always installed.
$nugetConfig = Join-Path $consumer 'NuGet.Config'
$escapedOutput = [Security.SecurityElement]::Escape($output)
@"
<configuration>
  <packageSources>
    <clear />
    <add key="local-archive" value="$escapedOutput" />
    <add key="sdk-packs" value="https://api.nuget.org/v3/index.json" />
  </packageSources>
  <packageSourceMapping>
    <packageSource key="local-archive"><package pattern="OfficeIMO.JavaScript.Assets" /></packageSource>
    <packageSource key="sdk-packs">
      <package pattern="Microsoft.NETCore.App.*" />
      <package pattern="Microsoft.AspNetCore.App.*" />
      <package pattern="Microsoft.WindowsDesktop.App.*" />
    </packageSource>
  </packageSourceMapping>
</configuration>
"@ | Set-Content -LiteralPath $nugetConfig -Encoding utf8
& dotnet restore $consumerProject --configfile $nugetConfig --packages $packageCache --force --nologo
if ($LASTEXITCODE -ne 0) { throw 'Isolated asset package restore failed.' }
$archive = Join-Path $output "OfficeIMO.JavaScript.Assets.$version.nupkg"
$restoredArchive = Join-Path $packageCache "officeimo.javascript.assets/$version/officeimo.javascript.assets.$version.nupkg"
if ((Get-FileHash -LiteralPath $archive).Hash -ne (Get-FileHash -LiteralPath $restoredArchive).Hash) {
    throw 'The consumer did not restore the freshly packed local archive.'
}
foreach ($framework in 'net8.0','net10.0') {
    & dotnet run --project $consumerProject -c Release -f $framework --no-restore -p:UseSharedCompilation=false -- (Join-Path $repository 'OfficeIMO.JavaScript/bundles')
    if ($LASTEXITCODE -ne 0) { throw "Packed consumer failed: $framework" }
}
Write-Host "Packed .NET asset consumers passed: OfficeIMO.JavaScript.Assets.$version.nupkg"
