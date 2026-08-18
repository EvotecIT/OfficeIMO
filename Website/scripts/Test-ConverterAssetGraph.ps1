[CmdletBinding(DefaultParameterSetName = 'Local')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Local')]
    [string] $SiteRoot,

    [Parameter(Mandatory, ParameterSetName = 'Remote')]
    [uri] $BaseUri,

    [string] $ExpectedDeploymentId,

    [string] $ExpectedDeploymentIdEnvironment
)

$ErrorActionPreference = 'Stop'
$isLocal = $PSCmdlet.ParameterSetName -eq 'Local'
$maximumDeploymentManifestBytes = 2L * 1024L * 1024L
$maximumAssetCount = 1024
$maximumAssetBytes = 64L * 1024L * 1024L
$maximumAggregateAssetBytes = 256L * 1024L * 1024L
$maximumRemoteAttempts = 3
$remoteRetryDelayMilliseconds = 250
$assetByteCache = [System.Collections.Generic.Dictionary[string, byte[]]]::new([System.StringComparer]::Ordinal)
$remoteHandler = $null
$remoteClient = $null

if (-not $isLocal) {
    if (-not $BaseUri.IsAbsoluteUri -or $BaseUri.Scheme -notin @('https', 'http')) {
        throw "Converter base URI '$BaseUri' must be an absolute HTTP(S) URI."
    }
    $converterRootBuilder = [System.UriBuilder]::new($BaseUri)
    $converterRootBuilder.Query = [string]::Empty
    $converterRootBuilder.Fragment = [string]::Empty
    if (-not $converterRootBuilder.Path.EndsWith('/', [StringComparison]::Ordinal)) {
        $converterRootBuilder.Path += '/'
    }
    $converterRootUri = $converterRootBuilder.Uri

    $remoteHandler = [System.Net.Http.HttpClientHandler]::new()
    $remoteHandler.AllowAutoRedirect = $false
    $remoteHandler.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Brotli
    $remoteClient = [System.Net.Http.HttpClient]::new($remoteHandler)
}

if ([string]::IsNullOrWhiteSpace($ExpectedDeploymentId) -and -not [string]::IsNullOrWhiteSpace($ExpectedDeploymentIdEnvironment)) {
    $ExpectedDeploymentId = [Environment]::GetEnvironmentVariable($ExpectedDeploymentIdEnvironment)
    if ([string]::IsNullOrWhiteSpace($ExpectedDeploymentId)) {
        throw "Expected deployment id environment '$ExpectedDeploymentIdEnvironment' is empty."
    }
}

function Get-RemoteBytes {
    param(
        [Parameter(Mandatory)][uri] $Uri,
        [Parameter(Mandatory)][long] $MaximumBytes
    )

    for ($attempt = 1; $attempt -le $maximumRemoteAttempts; $attempt++) {
        $response = $null
        $retryableFailure = $true
        try {
            $response = $remoteClient.GetAsync($Uri, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
            if (-not $response.IsSuccessStatusCode) {
                $statusCode = [int] $response.StatusCode
                if ($attempt -lt $maximumRemoteAttempts -and $statusCode -in @(408, 425, 429, 500, 502, 503, 504)) {
                    Start-Sleep -Milliseconds ($remoteRetryDelayMilliseconds * $attempt)
                    continue
                }
                $retryableFailure = $false
                throw "Converter asset '$Uri' returned HTTP $([int]$response.StatusCode)."
            }
            if ($response.Content.Headers.ContentLength.HasValue -and
                $response.Content.Headers.ContentLength.Value -gt $MaximumBytes) {
                throw "Converter asset '$Uri' declares more than the permitted $MaximumBytes bytes."
            }

            $input = $response.Content.ReadAsStream()
            try {
                $output = [System.IO.MemoryStream]::new()
                try {
                    $buffer = [byte[]]::new(81920)
                    while (($read = $input.Read($buffer, 0, $buffer.Length)) -gt 0) {
                        if ($output.Length -gt $MaximumBytes - $read) {
                            throw "Converter asset '$Uri' exceeded the permitted $MaximumBytes bytes while downloading."
                        }
                        $output.Write($buffer, 0, $read)
                    }
                    return $output.ToArray()
                } finally {
                    $output.Dispose()
                }
            } finally {
                $input.Dispose()
            }
        } catch {
            if (-not $retryableFailure -or $attempt -ge $maximumRemoteAttempts) {
                throw
            }
            Start-Sleep -Milliseconds ($remoteRetryDelayMilliseconds * $attempt)
        } finally {
            if ($null -ne $response) {
                $response.Dispose()
            }
        }
    }

    throw "Converter asset '$Uri' could not be downloaded after $maximumRemoteAttempts attempts."
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
        [Parameter(Mandatory)]
        [string] $RelativePath,

        [long] $MaximumBytes = $maximumAssetBytes
    )

    $canonical = ConvertTo-CanonicalAssetPath -RelativePath $RelativePath
    [byte[]] $cachedBytes = $null
    if ($assetByteCache.TryGetValue($canonical, [ref] $cachedBytes)) {
        if ($cachedBytes.LongLength -gt $MaximumBytes) {
            throw "Converter asset '$RelativePath' exceeds the permitted $MaximumBytes bytes."
        }
        return $cachedBytes
    }

    $normalized = $canonical -replace '/', [System.IO.Path]::DirectorySeparatorChar
    if ($isLocal) {
        $root = [System.IO.Path]::GetFullPath($SiteRoot)
        $path = [System.IO.Path]::GetFullPath((Join-Path $root $normalized))
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

    $assetUri = if ($canonical -eq 'index.html') {
        $BaseUri
    } else {
        [uri]::new($converterRootUri, $canonical)
    }
    if (-not [string]::Equals($assetUri.Scheme, $converterRootUri.Scheme, [StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals($assetUri.Authority, $converterRootUri.Authority, [StringComparison]::OrdinalIgnoreCase) -or
        -not $assetUri.AbsolutePath.StartsWith($converterRootUri.AbsolutePath, [StringComparison]::Ordinal)) {
        throw "Converter asset path '$RelativePath' resolves outside '$converterRootUri'."
    }
    $bytes = Get-RemoteBytes -Uri $assetUri -MaximumBytes $MaximumBytes
    $assetByteCache[$canonical] = $bytes
    return $bytes
}

function Get-Utf8Text {
    param([Parameter(Mandatory)][byte[]] $Bytes)
    return [System.Text.Encoding]::UTF8.GetString($Bytes)
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

function Get-Sha256Hex {
    param([Parameter(Mandatory)][byte[]] $Bytes)
    return [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($Bytes)).ToLowerInvariant()
}

try {
$deploymentManifestBytes = Get-AssetBytes -RelativePath 'deployment-assets.json' -MaximumBytes $maximumDeploymentManifestBytes
$deploymentManifest = Get-Utf8Text -Bytes $deploymentManifestBytes | ConvertFrom-Json
if ($deploymentManifest.schemaVersion -ne 1 -or [string] $deploymentManifest.deploymentId -notmatch '^[A-Fa-f0-9]{40}$|^[A-Fa-f0-9]{64}$') {
    throw 'Converter deployment manifest is missing its schema version or deployment id.'
}
if (-not [string]::IsNullOrWhiteSpace($ExpectedDeploymentId) -and
    -not [string]::Equals([string] $deploymentManifest.deploymentId, $ExpectedDeploymentId, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Converter deployment '$($deploymentManifest.deploymentId)' does not match expected source '$ExpectedDeploymentId'."
}

$manifestPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
$manifestIndex = $null
$manifestAssets = @($deploymentManifest.assets)
if ($manifestAssets.Count -gt $maximumAssetCount) {
    throw "Converter deployment manifest contains $($manifestAssets.Count) assets; the limit is $maximumAssetCount."
}
$validatedManifestAssets = [System.Collections.Generic.List[object]]::new()
[long] $declaredAggregateBytes = 0L
foreach ($asset in $manifestAssets) {
    $path = [string] $asset.path
    $expectedHash = [string] $asset.sha256
    $canonicalPath = ConvertTo-CanonicalAssetPath -RelativePath $path
    if ($canonicalPath -cne $path -or -not $manifestPaths.Add($canonicalPath)) {
        throw "Converter deployment manifest contains a blank or duplicate asset '$path'."
    }
    if ($expectedHash -notmatch '^[a-f0-9]{64}$') {
        throw "Converter deployment asset '$path' has an invalid SHA-256 digest."
    }
    [long] $declaredBytes = $asset.bytes
    if ($declaredBytes -lt 0L -or $declaredBytes -gt $maximumAssetBytes) {
        throw "Converter deployment asset '$path' declares an invalid byte length '$declaredBytes'."
    }
    if ($declaredAggregateBytes -gt $maximumAggregateAssetBytes - $declaredBytes) {
        throw "Converter deployment manifest exceeds the aggregate $maximumAggregateAssetBytes byte limit."
    }
    $declaredAggregateBytes += $declaredBytes
    $validatedManifestAssets.Add([pscustomobject]@{
        Path = $canonicalPath
        Bytes = $declaredBytes
        Sha256 = $expectedHash
    })
}

foreach ($asset in $validatedManifestAssets) {
    $bytes = Get-AssetBytes -RelativePath $asset.Path -MaximumBytes $asset.Bytes
    if ($bytes.LongLength -ne $asset.Bytes) {
        throw "Converter deployment asset '$($asset.Path)' has length $($bytes.LongLength), expected $($asset.Bytes)."
    }
    if ((Get-Sha256Hex -Bytes $bytes) -cne $asset.Sha256) {
        throw "Converter deployment asset '$($asset.Path)' does not match its deployment manifest digest."
    }
    if ($asset.Path -eq 'index.html') {
        $manifestIndex = $asset
    }
}
if ($null -eq $manifestIndex -or $manifestPaths.Count -lt 10) {
    throw "Converter deployment manifest unexpectedly contained only $($manifestPaths.Count) assets or no index.html."
}

if (-not $isLocal) {
    $entryPoint = [System.UriBuilder]::new($BaseUri)
    $entryPoint.Query = [string]::Empty
    $entryPoint.Fragment = [string]::Empty
    $plainIndexBytes = Get-RemoteBytes -Uri $entryPoint.Uri -MaximumBytes $manifestIndex.Bytes
    if ($plainIndexBytes.LongLength -ne $manifestIndex.Bytes -or
        (Get-Sha256Hex -Bytes $plainIndexBytes) -cne $manifestIndex.Sha256) {
        throw "Converter entry point '$($entryPoint.Uri)' does not match the current deployment manifest."
    }
}

$indexBytes = Get-AssetBytes -RelativePath 'index.html'
$index = Get-Utf8Text -Bytes $indexBytes
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

$dotnetBytes = Get-AssetBytes -RelativePath $dotnetPath
$dotnet = Get-Utf8Text -Bytes $dotnetBytes
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
        [Parameter(Mandatory)]
        [object] $Node,
        [Parameter(Mandatory)]
        [string] $BasePath
    )
    if ($Node -is [System.Collections.IDictionary]) {
        if ($Node.Contains('name')) {
            $name = [string] $Node['name']
            $resourceAssets.Add([pscustomobject]@{
                Path = $BasePath + $name
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

Write-Output "Converter deployment verified: $($manifestPaths.Count) public assets and $($verified.Count) fingerprinted runtime resources for $($deploymentManifest.deploymentId)."
} finally {
    if ($null -ne $remoteClient) {
        $remoteClient.Dispose()
        $remoteHandler.Dispose()
    }
}
