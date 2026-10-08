param([Parameter(Mandatory)][string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'This independent consumer probe requires installed Windows DAO.' }
$root = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $root) { throw 'Use a fresh directory; probes never replace existing files.' }
New-Item -ItemType Directory -Path $root | Out-Null
$engine = $null
$results = @()
try {
    $engine = New-Object -ComObject DAO.DBEngine.120
    $engineVersion = $engine.Version
    foreach ($profile in @(@{Name='jet4.mdb';Engine='Standard Jet DB';Version=1}, @{Name='ace12.accdb';Engine='Standard ACE DB';Version=2})) {
        $path = Join-Path $root $profile.Name
        # Negative feasibility control, not a writer. No file bytes or pages are copied from a seed.
        # A signature and page-type skeleton are deliberately insufficient to bootstrap a system catalog.
        $bytes = New-Object byte[] 16384
        $bytes[1] = 1
        [Text.Encoding]::ASCII.GetBytes($profile.Engine).CopyTo($bytes, 4)
        $bytes[20] = $profile.Version
        $bytes[4096] = 1; $bytes[8192] = 2; $bytes[12288] = 1
        [IO.File]::WriteAllBytes($path, $bytes)
        $db = $null
        $errorCode = $null
        try { $db = $engine.OpenDatabase($path, $false, $true); $outcome='opened' }
        catch { $outcome='rejected'; $errorCode=$_.Exception.HResult }
        finally { if ($db) { $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null } }
        $results += [ordered]@{path=$profile.Name;profile=$profile.Engine;probe='header-and-page-skeleton-negative-control';seedUsed=$false;outcome=$outcome;hresult=$errorCode;sha256=(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant();notImplemented=@('valid masked header', 'catalog fields and rows', 'allocation maps', 'system permissions', 'table indexes', 'relationships');requiredExperiment='native-bootstrap-01: encode system tables, allocation and catalog index roots from logical definitions; DAO must open the generated file before adding a user table'}
    }
} finally { if ($engine) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine) | Out-Null } }
[ordered]@{schemaVersion=1;producer='OfficeIMO verification script';consumer='Microsoft DAO';engineVersion=$engineVersion;proof='negative control only; does not prove a native writer';results=$results} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $root 'manifest.json') -Encoding utf8
$results | ConvertTo-Json -Depth 8
