[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $SiteRoot
)

$ErrorActionPreference = 'Stop'
$maximumAssetBytes = 64L * 1024L * 1024L
$assetByteCache = [System.Collections.Generic.Dictionary[string, byte[]]]::new([System.StringComparer]::Ordinal)
$root = [System.IO.Path]::GetFullPath($SiteRoot).TrimEnd(
    [System.IO.Path]::DirectorySeparatorChar,
    [System.IO.Path]::AltDirectorySeparatorChar)
if (-not (Test-Path -LiteralPath $root -PathType Container)) {
    throw "Converter site root '$root' does not exist."
}

function ConvertTo-CanonicalAssetPath {
    param([Parameter(Mandatory)][string] $RelativePath)

    if ([string]::IsNullOrWhiteSpace($RelativePath) -or
        $RelativePath.StartsWith('/', [StringComparison]::Ordinal) -or
        $RelativePath.Contains('\') -or
        $RelativePath -match '[:%?#]' -or
        $RelativePath.Contains('//')) {
        throw "Converter asset path '$RelativePath' is not a safe relative path."
    }
    $canonical = if ($RelativePath.StartsWith('./', [StringComparison]::Ordinal)) {
        $RelativePath.Substring(2)
    } else {
        $RelativePath
    }
    if ([string]::IsNullOrWhiteSpace($canonical)) {
        throw "Converter asset path '$RelativePath' is empty after normalization."
    }
    foreach ($segment in $canonical.Split('/')) {
        if ([string]::IsNullOrWhiteSpace($segment) -or $segment -in @('.', '..')) {
            throw "Converter asset path '$RelativePath' contains an unsafe segment."
        }
    }
    return $canonical
}

function Get-AssetBytes {
    param(
        [Parameter(Mandatory)][string] $RelativePath,
        [long] $MaximumBytes = $maximumAssetBytes
    )

    $canonical = ConvertTo-CanonicalAssetPath -RelativePath $RelativePath
    [byte[]] $cachedBytes = $null
    if ($assetByteCache.TryGetValue($canonical, [ref] $cachedBytes)) {
        return $cachedBytes
    }

    $path = [System.IO.Path]::GetFullPath((Join-Path $root ($canonical -replace '/', [System.IO.Path]::DirectorySeparatorChar)))
    if (-not $path.StartsWith($root + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Converter asset path '$RelativePath' escapes '$root'."
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Converter asset graph references missing file '$path'."
    }
    $length = (Get-Item -LiteralPath $path).Length
    if ($length -gt $MaximumBytes) {
        throw "Converter asset '$path' exceeds the permitted $MaximumBytes bytes."
    }
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $assetByteCache[$canonical] = $bytes
    return $bytes
}

function Assert-Sha256Integrity {
    param(
        [Parameter(Mandatory)][byte[]] $Bytes,
        [Parameter(Mandatory)][string] $Expected,
        [Parameter(Mandatory)][string] $Asset
    )

    if ($Expected -notmatch '^sha256-(?<digest>[A-Za-z0-9+/]+={0,2})$') {
        throw "Converter asset '$Asset' declares unsupported integrity '$Expected'."
    }
    $actual = [Convert]::ToBase64String([System.Security.Cryptography.SHA256]::HashData($Bytes))
    if ($actual -cne $Matches.digest) {
        throw "Converter asset '$Asset' failed SHA-256 integrity validation."
    }
}

$index = [System.Text.Encoding]::UTF8.GetString((Get-AssetBytes -RelativePath 'index.html'))
$importMapMatch = [regex]::Match(
    $index,
    '<script\s+type=["'']importmap["'']\s*>(?<json>.*?)</script>',
    [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
if (-not $importMapMatch.Success) {
    throw 'Converter index does not contain a parseable import map.'
}
$importMap = $importMapMatch.Groups['json'].Value | ConvertFrom-Json -AsHashtable
$dotnetPath = [string] $importMap.imports['./_framework/dotnet.js']
if ([string]::IsNullOrWhiteSpace($dotnetPath)) {
    throw 'Converter import map does not resolve ./_framework/dotnet.js.'
}

$verified = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
foreach ($resolvedPath in $importMap.imports.Values) {
    $path = [string] $resolvedPath
    if (-not $verified.Add($path)) {
        continue
    }
    $bytes = Get-AssetBytes -RelativePath $path
    $integrity = [string] $importMap.integrity[$path]
    if ([string]::IsNullOrWhiteSpace($integrity)) {
        throw "Converter import-map asset '$path' has no integrity value."
    }
    Assert-Sha256Integrity -Bytes $bytes -Expected $integrity -Asset $path
}

$dotnet = [System.Text.Encoding]::UTF8.GetString((Get-AssetBytes -RelativePath $dotnetPath))
$configMatch = [regex]::Match(
    $dotnet,
    '/\*json-start\*/(?<json>\{.*?\})/\*json-end\*/',
    [System.Text.RegularExpressions.RegexOptions]::Singleline)
if (-not $configMatch.Success) {
    throw "Converter runtime '$dotnetPath' does not contain an embedded boot manifest."
}
$config = $configMatch.Groups['json'].Value | ConvertFrom-Json -AsHashtable

$resourceAssets = [System.Collections.Generic.List[object]]::new()
function Add-ResourceAssets {
    param(
        [Parameter(Mandatory)][object] $Node,
        [Parameter(Mandatory)][string] $BasePath
    )

    if ($Node -is [System.Collections.IDictionary]) {
        if ($Node.Contains('name')) {
            $resourceAssets.Add([pscustomobject]@{
                Path = $BasePath + [string] $Node['name']
                Hash = if ($Node.Contains('hash')) { [string] $Node['hash'] } else { $null }
            })
            return
        }
        foreach ($entry in $Node.GetEnumerator()) {
            Add-ResourceAssets -Node $entry.Value -BasePath $BasePath
        }
        return
    }
    if ($Node -is [System.Collections.IEnumerable] -and $Node -isnot [string]) {
        foreach ($item in $Node) {
            Add-ResourceAssets -Node $item -BasePath $BasePath
        }
    }
}

foreach ($group in $config.resources.GetEnumerator()) {
    if ($group.Key -eq 'satelliteResources') {
        foreach ($culture in $group.Value.GetEnumerator()) {
            Add-ResourceAssets -Node $culture.Value -BasePath ("./_framework/$($culture.Key)/")
        }
        continue
    }
    $basePath = if ($group.Key -eq 'libraryInitializers') { './' } else { './_framework/' }
    Add-ResourceAssets -Node $group.Value -BasePath $basePath
}

foreach ($asset in $resourceAssets | Sort-Object Path -Unique) {
    if (-not $verified.Add($asset.Path)) {
        continue
    }
    $bytes = Get-AssetBytes -RelativePath $asset.Path
    if (-not [string]::IsNullOrWhiteSpace($asset.Hash)) {
        Assert-Sha256Integrity -Bytes $bytes -Expected $asset.Hash -Asset $asset.Path
    }
}

if ($verified.Count -lt 10) {
    throw "Converter asset graph unexpectedly contained only $($verified.Count) resources."
}

Write-Output "Converter asset graph verified locally: $($verified.Count) fingerprinted runtime resources."
