<#
.SYNOPSIS
Fetches the opt-in large real-world YAML corpus into a local cache and writes its index.

.DESCRIPTION
Reads sources.json next to this script. For each source it fetches the pinned commit without file contents
(a partial clone), lists the commit's tree, selects files by the source's include globs and sampling limits,
then checks out only those files and the license file. Each selected file is recorded in index.json with its
size, SHA-256 and kind ('yaml' or 'front-matter'), or as excluded with a reason: excluded paths, files over
the size bound, Markdown without YAML front matter and files a 'classify' rule identifies as templating.

The cache lives outside the repository and is never committed. Re-running is idempotent: fetched commits are
reused and only missing files are checked out. Move a pin by editing the commit in sources.json.

The cache root is -CacheRoot, else $env:OFFICEIMO_YAML_CORPUS_CACHE, else 'yaml-corpus' under the system
temporary folder. The large-corpus and mutation runs in OfficeIMO.Yaml.Benchmarks use the same rule.
#>
[CmdletBinding()]
param(
    [string] $CacheRoot,
    # Limit the run to these source ids.
    [string[]] $Source,
    # List the selection per source without checking files out or writing the index.
    [switch] $ListOnly
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 3

if (-not $CacheRoot) { $CacheRoot = $env:OFFICEIMO_YAML_CORPUS_CACHE }
if (-not $CacheRoot) { $CacheRoot = Join-Path ([IO.Path]::GetTempPath()) 'yaml-corpus' }
$CacheRoot = [IO.Path]::GetFullPath($CacheRoot)
$manifestPath = Join-Path $PSScriptRoot 'sources.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

function Invoke-Git {
    param([string] $Directory, [Parameter(ValueFromRemainingArguments)] [string[]] $Arguments)
    $output = & git -C $Directory @Arguments
    if ($LASTEXITCODE -ne 0) { throw "git $($Arguments -join ' ') failed in $Directory." }
    return $output
}

# Glob to anchored regular expression: '**/' spans zero or more folders, '*' and '?' stay inside one segment,
# '{a,b}' lists alternatives.
function ConvertTo-GlobRegex([string] $Glob) {
    $text = [Text.StringBuilder]::new('^')
    $i = 0
    while ($i -lt $Glob.Length) {
        $c = $Glob[$i]
        if ($c -eq '*' -and $i + 1 -lt $Glob.Length -and $Glob[$i + 1] -eq '*') {
            if ($i + 2 -lt $Glob.Length -and $Glob[$i + 2] -eq '/') { [void]$text.Append('(?:.*/)?'); $i += 3 }
            else { [void]$text.Append('.*'); $i += 2 }
            continue
        }
        switch ($c) {
            '*' { [void]$text.Append('[^/]*') }
            '?' { [void]$text.Append('[^/]') }
            '{' { [void]$text.Append('(?:') }
            '}' { [void]$text.Append(')') }
            ',' { [void]$text.Append('|') }
            default { [void]$text.Append([regex]::Escape([string]$c)) }
        }
        $i++
    }
    [void]$text.Append('$')
    return [regex]::new($text.ToString(), [Text.RegularExpressions.RegexOptions]::CultureInvariant)
}

function Get-Property($Object, [string] $Name, $Default = $null) {
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    return $property.Value
}

# Evenly spaced sample of a sorted list, so a bounded source still spans all of its folders.
function Select-Even([string[]] $Items, [int] $Max) {
    if ($Max -le 0 -or $Items.Count -le $Max) { return $Items }
    $result = [Collections.Generic.List[string]]::new()
    for ($k = 0; $k -lt $Max; $k++) { $result.Add($Items[[int][Math]::Floor($k * $Items.Count / $Max)]) }
    return $result.ToArray()
}

function Test-FrontMatter([byte[]] $Bytes) {
    $offset = 0
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { $offset = 3 }
    if ($Bytes.Length -lt $offset + 4) { return $false }
    if ($Bytes[$offset] -ne 0x2D -or $Bytes[$offset + 1] -ne 0x2D -or $Bytes[$offset + 2] -ne 0x2D) { return $false }
    $next = $Bytes[$offset + 3]
    return $next -eq 0x0A -or ($next -eq 0x0D)
}

$sha256 = [Security.Cryptography.SHA256]::Create()
$sourcesRoot = Join-Path $CacheRoot 'sources'
New-Item -ItemType Directory -Force -Path $sourcesRoot | Out-Null
$indexSources = [Collections.Generic.List[object]]::new()
$indexFiles = [Collections.Generic.List[object]]::new()
if ($Source) { $Source = @($Source | ForEach-Object { $_ -split ',' } | Where-Object { $_ }) }
$selectedSources = @($manifest.sources | Where-Object { -not $Source -or $Source -contains $_.id })
if ($Source -and $selectedSources.Count -ne @($Source).Count) { throw "Unknown source id in: $($Source -join ', ')." }

foreach ($entry in $selectedSources) {
    $directory = Join-Path $sourcesRoot $entry.id
    if (-not (Test-Path (Join-Path $directory '.git'))) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
        Invoke-Git $directory init --quiet | Out-Null
        Invoke-Git $directory remote add origin $entry.repository | Out-Null
        Invoke-Git $directory config core.longpaths true | Out-Null
    }

    & git -C $directory cat-file -e "$($entry.commit)^{commit}" 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Fetching $($entry.id) at $($entry.commit)..."
        # Large hosts occasionally refuse a pack under load; retry before giving up.
        for ($attempt = 1; ; $attempt++) {
            & git -C $directory -c protocol.version=2 fetch --quiet --depth 1 --filter=blob:none --no-tags origin $entry.commit
            if ($LASTEXITCODE -eq 0) { break }
            if ($attempt -ge 3) { throw "Fetching $($entry.commit) from $($entry.repository) failed." }
            Start-Sleep -Seconds (10 * $attempt)
        }
    }

    $commitDate = (Invoke-Git $directory log -1 --format=%cI $entry.commit).Trim()

    # Blobs only: symbolic links (120000) and submodules (commit entries) are not files to parse.
    $tree = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $raw = (Invoke-Git $directory ls-tree -r -z $entry.commit) -join "`n"
    foreach ($line in $raw.Split([char]0, [StringSplitOptions]::RemoveEmptyEntries)) {
        $tab = $line.IndexOf("`t")
        $meta = $line.Substring(0, $tab).Trim("`n").Split(' ')
        if ($meta[1] -eq 'blob' -and $meta[0] -ne '120000') { $tree[$line.Substring($tab + 1)] = $meta[2] }
    }
    $paths = [string[]]@($tree.Keys)
    [Array]::Sort($paths, [StringComparer]::Ordinal)

    $excludeRules = @(Get-Property $entry 'exclude' @() | ForEach-Object { [pscustomobject]@{ Regex = ConvertTo-GlobRegex $_.glob; Reason = $_.reason } })
    $excluded = [ordered]@{}
    $selected = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($include in @($entry.include)) {
        $glob = if ($include -is [string]) { $include } else { $include.glob }
        $max = if ($include -is [string]) { 0 } else { [int](Get-Property $include 'max' 0) }
        $regex = ConvertTo-GlobRegex $glob
        $found = [Collections.Generic.List[string]]::new()
        foreach ($path in $paths) {
            if (-not $regex.IsMatch($path)) { continue }
            $rule = $excludeRules | Where-Object { $_.Regex.IsMatch($path) } | Select-Object -First 1
            if ($rule) { $excluded[$path] = $rule.Reason; continue }
            $found.Add($path)
        }
        $sample = Select-Even $found.ToArray() $max
        foreach ($path in $sample) { [void]$selected.Add($path) }
        if ($ListOnly) { Write-Host ("  {0,-50} {1,6} matched, {2,6} kept" -f $glob, $found.Count, @($sample).Count) }
    }

    $files = [string[]]@($selected)
    [Array]::Sort($files, [StringComparer]::Ordinal)
    Write-Host ("{0}: {1} files selected, {2} excluded by path" -f $entry.id, $files.Count, $excluded.Count)
    if ($ListOnly) { continue }

    # Check out only the selected files and the license; the partial clone fetches their contents in one batch.
    $licenseFile = Get-Property $entry 'licenseFile'
    $checkout = [Collections.Generic.List[string]]::new($files)
    if ($licenseFile) {
        if (-not $tree.ContainsKey($licenseFile)) { throw "$($entry.id): license file '$licenseFile' is not in commit $($entry.commit)." }
        $checkout.Add($licenseFile)
    }
    $missing = @($checkout | Where-Object { -not (Test-Path -LiteralPath (Join-Path $directory $_) -PathType Leaf) })
    if ($missing.Count -gt 0) {
        # Fetch the missing contents in one request, as git's own lazy fetch does, instead of one request per file.
        $oids = @($missing | ForEach-Object { $tree[$_] } | Sort-Object -Unique)
        for ($attempt = 1; ; $attempt++) {
            $oids | & git -C $directory -c fetch.negotiationAlgorithm=noop fetch --quiet --no-tags --no-write-fetch-head --recurse-submodules=no --filter=blob:none --stdin origin
            if ($LASTEXITCODE -eq 0) { break }
            if ($attempt -ge 3) { throw "Fetching $($oids.Count) files of $($entry.id) failed." }
            Start-Sleep -Seconds (10 * $attempt)
        }
        $list = Join-Path $directory '.git/officeimo-pathspec'
        [IO.File]::WriteAllText($list, ($missing -join "`0"), [Text.UTF8Encoding]::new($false))
        Invoke-Git $directory --literal-pathspecs checkout --quiet $entry.commit --pathspec-from-file=$list --pathspec-file-nul | Out-Null
    }

    $maxBytes = [long](Get-Property $entry 'maxFileBytes' 0)
    $classifyRules = @(Get-Property $entry 'classify' @() | ForEach-Object {
        [pscustomobject]@{ Glob = ConvertTo-GlobRegex $_.glob; Pattern = [regex]::new($_.pattern); Reason = $_.reason }
    })
    $kept = 0
    $sourceExcluded = 0
    foreach ($path in $files) {
        $full = Join-Path $directory $path
        $bytes = [IO.File]::ReadAllBytes($full)
        $kind = if ($path -match '\.(md|markdown|html)$') { 'front-matter' } else { 'yaml' }
        $reason = $null
        if ($maxBytes -gt 0 -and $bytes.Length -gt $maxBytes) { $reason = "Larger than the source's $maxBytes-byte bound." }
        elseif ($kind -eq 'front-matter' -and -not (Test-FrontMatter $bytes)) { $reason = 'No YAML front matter (no leading --- line).' }
        else {
            foreach ($rule in $classifyRules) {
                if ($rule.Glob.IsMatch($path) -and $rule.Pattern.IsMatch([Text.Encoding]::UTF8.GetString($bytes))) { $reason = $rule.Reason; break }
            }
        }
        if ($reason) { $sourceExcluded++ } else { $kept++ }
        $indexFiles.Add([ordered]@{
            source   = $entry.id
            category = $entry.category
            path     = $path
            kind     = $kind
            bytes    = $bytes.Length
            sha256   = [Convert]::ToHexString($sha256.ComputeHash($bytes)).ToLowerInvariant()
            excluded = $reason
        })
    }
    foreach ($path in $excluded.Keys) {
        $sourceExcluded++
        $indexFiles.Add([ordered]@{ source = $entry.id; category = $entry.category; path = $path; kind = $null; bytes = $null; sha256 = $null; excluded = $excluded[$path] })
    }

    $indexSources.Add([ordered]@{
        id          = $entry.id
        category    = $entry.category
        repository  = $entry.repository
        commit      = $entry.commit
        commitDate  = $commitDate
        license     = $entry.license
        licenseFile = $licenseFile
        files       = $kept
        excluded    = $sourceExcluded
    })
}

if ($ListOnly) { return }

$index = [ordered]@{
    manifest  = 'OfficeIMO.TestAssets/Yaml/LargeCorpus/sources.json'
    manifestSha256 = [Convert]::ToHexString($sha256.ComputeHash([IO.File]::ReadAllBytes($manifestPath))).ToLowerInvariant()
    partial   = [bool]$Source
    sources   = $indexSources
    files     = $indexFiles
}
$json = ($index | ConvertTo-Json -Depth 5) -replace "`r`n", "`n"
$indexName = if ($Source) { 'index.partial.json' } else { 'index.json' }
[IO.File]::WriteAllText((Join-Path $CacheRoot $indexName), $json + "`n", [Text.UTF8Encoding]::new($false))
$total = ($indexSources | ForEach-Object { $_.files } | Measure-Object -Sum).Sum
$skipped = ($indexSources | ForEach-Object { $_.excluded } | Measure-Object -Sum).Sum
Write-Host "Indexed $total files from $($indexSources.Count) sources ($skipped excluded) in $(Join-Path $CacheRoot $indexName)."
