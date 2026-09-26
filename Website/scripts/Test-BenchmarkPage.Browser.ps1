param(
    [Parameter(Mandatory)]
    [string] $SiteRoot
)

$ErrorActionPreference = 'Stop'

$resolvedSiteRoot = [System.IO.Path]::GetFullPath($SiteRoot)
$browserCandidates = [System.Collections.Generic.List[string]]::new()
foreach ($commandName in @('google-chrome', 'chromium', 'chromium-browser', 'chrome', 'msedge')) {
    $command = Get-Command $commandName -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) {
        $browserCandidates.Add($command.Source)
    }
}
foreach ($path in @(
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
)) {
    if (Test-Path -LiteralPath $path -PathType Leaf) {
        $browserCandidates.Add($path)
    }
}
$browserPath = $browserCandidates | Select-Object -Unique -First 1
if (-not $browserPath) {
    throw 'A Chromium browser is required for the benchmark-page behavior contract.'
}

$pythonCommand = Get-Command python3, python -ErrorAction SilentlyContinue | Select-Object -First 1
$nodeCommand = Get-Command node -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $pythonCommand -or -not $nodeCommand) {
    throw 'Python and Node.js are required for the benchmark-page browser contract.'
}

function Get-FreeTcpPort {
    $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    $listener.Start()
    $port = ([System.Net.IPEndPoint] $listener.LocalEndpoint).Port
    $listener.Stop()
    return $port
}

$sitePort = Get-FreeTcpPort
$debugPort = Get-FreeTcpPort
$profileRoot = Join-Path ([System.IO.Path]::GetTempPath()) (
    'OfficeIMO-benchmark-browser-' + [Guid]::NewGuid().ToString('N'))
[void] (New-Item -ItemType Directory -Path $profileRoot)
$server = $null
$browser = $null

try {
    $serverInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $serverInfo.FileName = $pythonCommand.Source
    $serverInfo.UseShellExecute = $false
    $serverInfo.CreateNoWindow = $true
    $serverInfo.RedirectStandardOutput = $true
    $serverInfo.RedirectStandardError = $true
    foreach ($argument in @('-m', 'http.server', [string] $sitePort, '--bind', '127.0.0.1', '--directory', $resolvedSiteRoot)) {
        [void] $serverInfo.ArgumentList.Add($argument)
    }
    $server = [System.Diagnostics.Process]::Start($serverInfo)

    $browserInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $browserInfo.FileName = $browserPath
    $browserInfo.UseShellExecute = $false
    $browserInfo.CreateNoWindow = $true
    foreach ($argument in @(
        '--headless=new',
        '--disable-gpu',
        '--disable-background-networking',
        '--no-sandbox',
        "--remote-debugging-port=$debugPort",
        "--user-data-dir=$profileRoot",
        'about:blank'
    )) {
        [void] $browserInfo.ArgumentList.Add($argument)
    }
    $browser = [System.Diagnostics.Process]::Start($browserInfo)

    $ready = $false
    foreach ($attempt in 1..60) {
        try {
            $siteResponse = Invoke-WebRequest -Uri "http://127.0.0.1:$sitePort/benchmarks/" -UseBasicParsing
            $targets = Invoke-RestMethod -Uri "http://127.0.0.1:$debugPort/json/list"
            if ($siteResponse.StatusCode -eq 200 -and @($targets | Where-Object type -eq 'page').Count -gt 0) {
                $ready = $true
                break
            }
        } catch {
            Start-Sleep -Milliseconds 250
        }
    }
    if (-not $ready) {
        throw 'The generated site and Chromium debugging endpoint did not become ready.'
    }

    $browserContract = Join-Path $PSScriptRoot 'Test-BenchmarkPage.Browser.mjs'
    & $nodeCommand.Source $browserContract $debugPort "http://127.0.0.1:$sitePort"
    if ($LASTEXITCODE -ne 0) {
        throw "The benchmark-page browser contract failed with exit code $LASTEXITCODE."
    }
} finally {
    if ($browser -and -not $browser.HasExited) {
        $browser.Kill($true)
        [void] $browser.WaitForExit(5000)
    }
    if ($server -and -not $server.HasExited) {
        $server.Kill($true)
        [void] $server.WaitForExit(5000)
    }
    if (Test-Path -LiteralPath $profileRoot) {
        foreach ($attempt in 1..20) {
            try {
                # -Force: on Unix, PowerShell treats dotfiles as hidden, and a killed Chromium can
                # leave dot-prefixed temporary files behind in the profile.
                Remove-Item -LiteralPath $profileRoot -Recurse -Force -ErrorAction Stop
                break
            } catch {
                if ($attempt -ge 20) {
                    throw
                }
                Start-Sleep -Milliseconds 250
            }
        }
    }
}

Write-Host 'Benchmark browser behavior verified for QuestPDF version rendering and both CPU-domain selections.'
