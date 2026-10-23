param(
    [Parameter(Mandatory)][string] $PackagePath,
    [string] $BundleDirectory = "$PSScriptRoot/../../OfficeIMO.JavaScript/bundles"
)
$ErrorActionPreference = 'Stop'
$package = [IO.Path]::GetFullPath($PackagePath)
# This checks runtime resources in an explicitly supplied local or released package.
# Signature verification belongs to the signed-release caller; local CI packs are unsigned.
$archive = [IO.Compression.ZipFile]::OpenRead($package)
try {
    $manifests = @($archive.Entries | Where-Object { $_.FullName -match '^[^/]+\.nuspec$' })
    if ($manifests.Count -ne 1) { throw 'Expected one root NuGet manifest.' }
    $reader = [IO.StreamReader]::new($manifests[0].Open())
    try { [xml]$manifest = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $version = [string]$manifest.package.metadata.version
    if ($manifest.package.metadata.id -cne 'OfficeIMO.JavaScript.Assets' -or $version -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$') {
        throw 'Expected a stable OfficeIMO.JavaScript.Assets package.'
    }
    $dll = $archive.GetEntry('lib/net8.0/OfficeIMO.JavaScript.Assets.dll')
    if (-not $dll) { throw 'The JavaScript asset package has no .NET 8 asset assembly.' }
    $stream = $dll.Open(); $memory = [IO.MemoryStream]::new()
    try { $stream.CopyTo($memory); $assembly = [Reflection.Assembly]::Load($memory.ToArray()) }
    finally { $stream.Dispose(); $memory.Dispose() }
    $prefix = 'OfficeIMO.JavaScript.Assets.'
    $resources = @($assembly.GetManifestResourceNames() | Where-Object { $_.StartsWith($prefix, [StringComparison]::Ordinal) })
    $bundles = @(Get-ChildItem -LiteralPath $BundleDirectory -File | Where-Object { $_.Extension -cin @('.js', '.mjs') })
    if (-not $bundles.Count -or $resources.Count -ne $bundles.Count) { throw 'JavaScript Assets package and npm runtime asset sets differ.' }
    foreach ($bundle in $bundles) {
        $name = $prefix + $bundle.Name
        if ($resources -cnotcontains $name) { throw "The JavaScript asset package is missing $($bundle.Name)." }
        $resource = $assembly.GetManifestResourceStream($name)
        $local = [IO.File]::OpenRead($bundle.FullName)
        $hash = [Security.Cryptography.SHA256]::Create()
        try {
            $expected = [Convert]::ToHexString($hash.ComputeHash($resource))
            $actual = [Convert]::ToHexString($hash.ComputeHash($local))
            if ($actual -cne $expected) { throw "npm source differs from the JavaScript Assets package runtime: $($bundle.Name)." }
        } finally { $resource.Dispose(); $local.Dispose(); $hash.Dispose() }
    }
    $version
} finally { $archive.Dispose() }
