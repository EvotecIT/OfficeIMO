[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$subject = Join-Path $PSScriptRoot 'Test-ConverterAssetGraph.ps1'
$fixtureRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('officeimo-converter-graph-' + [guid]::NewGuid().ToString('N'))
$deploymentId = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'

function Get-Sha256Hex {
    param([Parameter(Mandatory)][byte[]] $Bytes)
    return [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($Bytes)).ToLowerInvariant()
}

function Get-Sha256Integrity {
    param([Parameter(Mandatory)][byte[]] $Bytes)
    return 'sha256-' + [Convert]::ToBase64String([System.Security.Cryptography.SHA256]::HashData($Bytes))
}

function Write-FixtureFile {
    param(
        [Parameter(Mandatory)][string] $RelativePath,
        [Parameter(Mandatory)][byte[]] $Bytes
    )

    $path = Join-Path $fixtureRoot ($RelativePath -replace '/', [System.IO.Path]::DirectorySeparatorChar)
    [System.IO.Directory]::CreateDirectory((Split-Path -Parent $path)) | Out-Null
    [System.IO.File]::WriteAllBytes($path, $Bytes)
}

function New-ConverterFixture {
    [System.IO.Directory]::CreateDirectory($fixtureRoot) | Out-Null
    $utf8 = [System.Text.UTF8Encoding]::new($false)
    $resourceBytes = [ordered]@{
        './_framework/dotnet.js' = $utf8.GetBytes('/*json-start*/{"resources":{}}/*json-end*/')
    }
    for ($index = 0; $index -lt 9; $index++) {
        $resourceBytes["./payload-$index.js"] = $utf8.GetBytes("export const payload$index = $index;")
    }

    $imports = [ordered]@{}
    $integrity = [ordered]@{}
    foreach ($entry in $resourceBytes.GetEnumerator()) {
        $imports[$entry.Key] = $entry.Key
        $integrity[$entry.Key] = Get-Sha256Integrity -Bytes $entry.Value
        Write-FixtureFile -RelativePath $entry.Key.Substring(2) -Bytes $entry.Value
    }

    $importMap = [ordered]@{
        imports = $imports
        integrity = $integrity
    } | ConvertTo-Json -Depth 5 -Compress
    $indexBytes = $utf8.GetBytes("<!doctype html><script type=`"importmap`">$importMap</script>")
    Write-FixtureFile -RelativePath 'index.html' -Bytes $indexBytes

    $manifestAssets = [System.Collections.Generic.List[object]]::new()
    foreach ($relativePath in @('index.html') + @($resourceBytes.Keys | ForEach-Object { $_.Substring(2) })) {
        $path = Join-Path $fixtureRoot ($relativePath -replace '/', [System.IO.Path]::DirectorySeparatorChar)
        $bytes = [System.IO.File]::ReadAllBytes($path)
        $manifestAssets.Add([ordered]@{
            path = $relativePath
            bytes = $bytes.LongLength
            sha256 = Get-Sha256Hex -Bytes $bytes
        })
    }
    $manifest = [ordered]@{
        schemaVersion = 1
        deploymentId = $deploymentId
        assets = @($manifestAssets)
    } | ConvertTo-Json -Depth 5
    Write-FixtureFile -RelativePath 'deployment-assets.json' -Bytes $utf8.GetBytes($manifest)
}

function Get-AvailablePort {
    $probe = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    try {
        $probe.Start()
        return ([System.Net.IPEndPoint] $probe.LocalEndpoint).Port
    } finally {
        $probe.Stop()
    }
}

function Start-FixtureServer {
    param([switch] $AlwaysStale)

    $serverScript = {
        param($Port, $Root, $ServeAlwaysStale)

        $listener = [System.Net.HttpListener]::new()
        $listener.Prefixes.Add("http://127.0.0.1:$Port/")
        $tokens = [System.Collections.Generic.List[string]]::new()
        $requests = [System.Collections.Generic.List[object]]::new()
        $utf8 = [System.Text.UTF8Encoding]::new($false)

        function Send-Response {
            param(
                [Parameter(Mandatory)] $Context,
                [Parameter(Mandatory)][byte[]] $Bytes,
                [int] $StatusCode = 200,
                [string] $ContentType = 'application/octet-stream'
            )
            $Context.Response.StatusCode = $StatusCode
            $Context.Response.ContentType = $ContentType
            $Context.Response.ContentLength64 = $Bytes.LongLength
            if ($Bytes.Length -gt 0) {
                $Context.Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
            }
            $Context.Response.Close()
        }

        try {
            $listener.Start()
            while ($listener.IsListening) {
                $context = $listener.GetContext()
                $requestPath = $context.Request.Url.AbsolutePath
                if ($requestPath -eq '/__health') {
                    Send-Response -Context $context -Bytes $utf8.GetBytes('ok') -ContentType 'text/plain'
                    continue
                }
                if ($requestPath -eq '/__state') {
                    $state = [ordered]@{
                        tokens = @($tokens)
                        requests = @($requests)
                    } | ConvertTo-Json -Depth 5 -Compress
                    Send-Response -Context $context -Bytes $utf8.GetBytes($state) -ContentType 'application/json'
                    continue
                }
                if ($requestPath -eq '/__stop') {
                    Send-Response -Context $context -Bytes $utf8.GetBytes('stopping') -ContentType 'text/plain'
                    $listener.Stop()
                    continue
                }

                $prefix = '/apps/officeimo-converter/'
                if (-not $requestPath.StartsWith($prefix, [StringComparison]::Ordinal)) {
                    Send-Response -Context $context -Bytes $utf8.GetBytes('not found') -StatusCode 404 -ContentType 'text/plain'
                    continue
                }
                $relativePath = $requestPath.Substring($prefix.Length)
                if ([string]::IsNullOrWhiteSpace($relativePath)) {
                    $relativePath = 'index.html'
                }
                $token = [string] $context.Request.QueryString['_officeimo_verify']
                if (-not [string]::IsNullOrWhiteSpace($token) -and -not $tokens.Contains($token)) {
                    $tokens.Add($token)
                }
                $requests.Add([ordered]@{ path = $relativePath; token = $token })

                if ($relativePath.Contains('..')) {
                    Send-Response -Context $context -Bytes $utf8.GetBytes('invalid path') -StatusCode 400 -ContentType 'text/plain'
                    continue
                }
                $path = Join-Path $Root ($relativePath -replace '/', [System.IO.Path]::DirectorySeparatorChar)
                if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                    Send-Response -Context $context -Bytes $utf8.GetBytes('not found') -StatusCode 404 -ContentType 'text/plain'
                    continue
                }
                $bytes = [System.IO.File]::ReadAllBytes($path)
                $firstToken = if ($tokens.Count -gt 0) { $tokens[0] } else { $null }
                $serveStale = $relativePath -eq 'payload-0.js' -and
                    ($ServeAlwaysStale -or [string]::IsNullOrWhiteSpace($token) -or $token -eq $firstToken)
                if ($serveStale) {
                    $bytes = $bytes + $utf8.GetBytes('-sticky-stale-response')
                }
                $context.Response.Headers['CF-Cache-Status'] = 'HIT'
                $context.Response.Headers['CF-Ray'] = 'fixture-ray'
                Send-Response -Context $context -Bytes $bytes
            }
        } finally {
            if ($listener.IsListening) {
                $listener.Stop()
            }
            $listener.Close()
        }
    }

    $lastStartupFailure = $null
    for ($bindAttempt = 1; $bindAttempt -le 5; $bindAttempt++) {
        $port = Get-AvailablePort
        $job = Start-Job -ScriptBlock $serverScript -ArgumentList $port, $fixtureRoot, ([bool] $AlwaysStale)
        $ready = $false
        for ($readinessAttempt = 0; $readinessAttempt -lt 50; $readinessAttempt++) {
            try {
                Invoke-RestMethod -Uri "http://127.0.0.1:$port/__health" -TimeoutSec 1 | Out-Null
                $ready = $true
                break
            } catch {
                if ($job.State -in @('Completed', 'Failed', 'Stopped')) {
                    break
                }
                Start-Sleep -Milliseconds 100
            }
        }

        if ($ready) {
            $baseUri = "http://127.0.0.1:$port/apps/officeimo-converter/?embedded=1"
            return [pscustomobject]@{ Job = $job; Port = $port; BaseUri = $baseUri }
        }

        $lastStartupFailure = (Receive-Job -Job $job -ErrorAction SilentlyContinue 2>&1 | Out-String).Trim()
        Stop-Job -Job $job -ErrorAction SilentlyContinue
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }

    throw "Converter asset-graph fixture server did not start after 5 bind attempts. $lastStartupFailure"
}

function Stop-FixtureServer {
    param([Parameter(Mandatory)] $Server)

    try {
        Invoke-RestMethod -Uri "http://127.0.0.1:$($Server.Port)/__stop" -TimeoutSec 2 | Out-Null
    } catch {
    }
    Wait-Job -Job $Server.Job -Timeout 5 | Out-Null
    Stop-Job -Job $Server.Job -ErrorAction SilentlyContinue
    Remove-Job -Job $Server.Job -Force -ErrorAction SilentlyContinue
}

function Get-FixtureServerState {
    param([Parameter(Mandatory)] $Server)
    return Invoke-RestMethod -Uri "http://127.0.0.1:$($Server.Port)/__state" -TimeoutSec 2
}

New-ConverterFixture
try {
    $successReportPath = Join-Path $fixtureRoot 'retry-success.json'
    $server = Start-FixtureServer
    try {
        & $subject -BaseUri $server.BaseUri -ExpectedDeploymentId $deploymentId `
            -RemoteVerificationAttempts 2 -RemoteVerificationRetryDelaySeconds 0 -ReportPath $successReportPath
        $state = Get-FixtureServerState -Server $server
        if (@($state.tokens).Count -ne 2) {
            throw "Expected two distinct verification cache identities, observed $(@($state.tokens).Count)."
        }
        if (@($state.requests | Where-Object { [string]::IsNullOrWhiteSpace([string] $_.token) }).Count -ne 0) {
            throw 'At least one converter asset request was sent without a verification cache identity.'
        }
        $report = Get-Content -LiteralPath $successReportPath -Raw | ConvertFrom-Json
        if ($report.outcome -ne 'success' -or @($report.attempts).Count -ne 2 -or $report.attempts[0].outcome -ne 'failure') {
            throw 'Successful retry report did not preserve the failed sticky-cache attempt.'
        }
    } finally {
        Stop-FixtureServer -Server $server
    }

    $failureReportPath = Join-Path $fixtureRoot 'retry-failure.json'
    $server = Start-FixtureServer -AlwaysStale
    try {
        $failed = $false
        try {
            & $subject -BaseUri $server.BaseUri -ExpectedDeploymentId $deploymentId `
                -RemoteVerificationAttempts 2 -RemoteVerificationRetryDelaySeconds 0 -ReportPath $failureReportPath
        } catch {
            $failed = $true
        }
        if (-not $failed) {
            throw 'Permanently stale converter assets unexpectedly passed integrity verification.'
        }
        $report = Get-Content -LiteralPath $failureReportPath -Raw | ConvertFrom-Json
        if ($report.outcome -ne 'failure' -or @($report.attempts).Count -ne 2 -or
            $report.lastError -notmatch 'payload-0\.js' -or
            $report.attempts[-1].lastResponse.CacheStatus -ne 'HIT') {
            throw 'Failed retry report did not identify the stale asset and cache response.'
        }
    } finally {
        Stop-FixtureServer -Server $server
    }

    Write-Output 'Converter asset-graph retry contract verified: cache identities rotate and persistent stale bytes still fail.'
} finally {
    if (Test-Path -LiteralPath $fixtureRoot) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force
    }
}
