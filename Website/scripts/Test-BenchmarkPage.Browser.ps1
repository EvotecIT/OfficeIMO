param(
    [Parameter(Mandatory)]
    [string] $SiteRoot,
    [string] $EvidenceRoot
)

$ErrorActionPreference = 'Stop'
$resolvedSiteRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SiteRoot)
if (-not (Test-Path -LiteralPath (Join-Path $resolvedSiteRoot 'benchmarks/index.html') -PathType Leaf)) {
    throw "The generated benchmark page is missing from '$resolvedSiteRoot'."
}
$pythonCommand = Get-Command python3, python -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $pythonCommand) { throw 'Python is required to serve the generated benchmark page.' }
$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
$listener.Start()
$port = ([System.Net.IPEndPoint] $listener.LocalEndpoint).Port
$listener.Stop()
$origin = "http://127.0.0.1:$port"
$serverInfo = [System.Diagnostics.ProcessStartInfo]::new()
$serverInfo.FileName = $pythonCommand.Source
$serverInfo.UseShellExecute = $false
$serverInfo.CreateNoWindow = $true
$serverInfo.RedirectStandardOutput = $true
$serverInfo.RedirectStandardError = $true
foreach ($argument in @('-m', 'http.server', [string] $port, '--bind', '127.0.0.1', '--directory', $resolvedSiteRoot)) {
    [void] $serverInfo.ArgumentList.Add($argument)
}
$server = [System.Diagnostics.Process]::Start($serverInfo)
$stdout = $server.StandardOutput.ReadToEndAsync()
$stderr = $server.StandardError.ReadToEndAsync()
try {
    $ready = $false
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    while ([DateTime]::UtcNow -lt $deadline -and -not $server.HasExited) {
        try {
            $ready = (Invoke-WebRequest -Uri "$origin/benchmarks/" -TimeoutSec 2 -UseBasicParsing).StatusCode -eq 200
            if ($ready) { break }
        } catch { Start-Sleep -Milliseconds 100 }
    }
    if (-not $ready) {
        if (-not $server.HasExited) { $server.Kill($true); [void] $server.WaitForExit(5000) }
        throw "The generated benchmark page did not become ready at $origin. Server diagnostics: $($stderr.GetAwaiter().GetResult())"
    }
    $project = Join-Path $PSScriptRoot '../../Build/WebsiteBrowserContract/OfficeIMO.WebsiteBrowserContract.csproj'
    $arguments = @('run', '--project', $project, '--configuration', 'Release', '--', $origin)
    if ($EvidenceRoot) { $arguments += $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($EvidenceRoot) }
    & dotnet @arguments
    if ($LASTEXITCODE -ne 0) { throw "The benchmark-page browser contract failed with exit code $LASTEXITCODE." }
} finally {
    if (-not $server.HasExited) { $server.Kill($true); [void] $server.WaitForExit(5000) }
    $server.Dispose()
}

Write-Host 'Benchmark browser behavior verified for the rendered dependency version and both CPU-domain selections.'
