[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $SiteRoot,

    [string] $BudgetPath = (Join-Path $PSScriptRoot '../config/converter-performance-budgets.json'),

    [string] $ReportPath = (Join-Path $PSScriptRoot '../_reports/browser-converter-performance.json'),

    [string] $BaseUrl
)

$ErrorActionPreference = 'Stop'
$resolvedSiteRoot = (Resolve-Path -LiteralPath $SiteRoot).Path
$resolvedBudgetPath = (Resolve-Path -LiteralPath $BudgetPath).Path
$runnerPath = Join-Path $PSScriptRoot 'converter-performance.playwright.js'
$budgets = Get-Content -LiteralPath $resolvedBudgetPath -Raw | ConvertFrom-Json
$playwrightPackage = "@playwright/cli@$($budgets.playwrightCliVersion)"
$browserEngine = [string] $budgets.browserEngine
if ([string]::IsNullOrWhiteSpace($budgets.playwrightCliVersion) -or [string]::IsNullOrWhiteSpace($browserEngine)) {
    throw 'The browser converter budget must pin playwrightCliVersion and browserEngine.'
}
$converterRoot = Join-Path $resolvedSiteRoot 'apps/officeimo-converter'
if (-not (Test-Path -LiteralPath $converterRoot -PathType Container)) {
    throw "Published browser converter was not found at '$converterRoot'."
}

$publishedBytes = (Get-ChildItem -LiteralPath $converterRoot -Recurse -File | Measure-Object -Property Length -Sum).Sum
if ($publishedBytes -gt [long] $budgets.maximumPublishedBytes) {
    throw "Browser converter publish is $publishedBytes bytes; budget is $($budgets.maximumPublishedBytes) bytes."
}

$server = $null
$session = 'officeimo-converter-performance-' + [Guid]::NewGuid().ToString('N')
$serverStandardOutputPath = Join-Path ([System.IO.Path]::GetTempPath()) "$session-server.stdout.log"
$serverStandardErrorPath = Join-Path ([System.IO.Path]::GetTempPath()) "$session-server.stderr.log"
try {
    if ([string]::IsNullOrWhiteSpace($BaseUrl)) {
        $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
        $listener.Start()
        $port = ([System.Net.IPEndPoint] $listener.LocalEndpoint).Port
        $listener.Stop()
        $python = Get-Command python -ErrorAction Stop
        $serverArguments = @{
            FilePath = $python.Source
            ArgumentList = @('-m', 'http.server', $port, '--bind', '127.0.0.1', '--directory', $resolvedSiteRoot)
            PassThru = $true
            RedirectStandardOutput = $serverStandardOutputPath
            RedirectStandardError = $serverStandardErrorPath
        }
        if ($IsWindows) { $serverArguments.WindowStyle = 'Hidden' }
        $server = Start-Process @serverArguments
        $BaseUrl = "http://127.0.0.1:$port/apps/officeimo-converter/"
        $ready = $false
        for ($attempt = 0; $attempt -lt 50; $attempt++) {
            try {
                $response = Invoke-WebRequest -UseBasicParsing -Uri $BaseUrl -TimeoutSec 2
                if ($response.StatusCode -eq 200) { $ready = $true; break }
            } catch {
                Start-Sleep -Milliseconds 200
            }
        }
        if (-not $ready) { throw "Local converter server did not become ready at $BaseUrl." }
    }

    $npx = Get-Command npx -ErrorAction Stop
    $installOutput = & $npx.Source --yes --package $playwrightPackage playwright-cli install-browser $browserEngine 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Playwright could not install browser '$browserEngine'.`n$($installOutput -join [Environment]::NewLine)" }
    $openOutput = & $npx.Source --yes --package $playwrightPackage playwright-cli "-s=$session" open $BaseUrl --browser $browserEngine 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Playwright could not open the converter.`n$($openOutput -join [Environment]::NewLine)" }
    $rawResult = & $npx.Source --yes --package $playwrightPackage playwright-cli "-s=$session" run-code --filename $runnerPath
    if ($LASTEXITCODE -ne 0) { throw 'Playwright browser performance run failed.' }
    $rawText = $rawResult -join [Environment]::NewLine
    $resultMatch = [regex]::Match($rawText, '(?ms)^### Result\r?\n(?<json>.+?)\r?\n### (?:Ran|Page|Error)')
    if (-not $resultMatch.Success) { throw "Playwright did not emit a parseable result block.`n$rawText" }
    $result = $resultMatch.Groups['json'].Value.Trim() | ConvertFrom-Json
    if ($result -is [string]) { $result = $result | ConvertFrom-Json }

    if ([double] $result.startupMilliseconds -gt [double] $budgets.maximumStartupMilliseconds) {
        throw "Converter startup took $($result.startupMilliseconds) ms; budget is $($budgets.maximumStartupMilliseconds) ms."
    }
    if ([long] $result.maximumBrowserHeapBytes -gt [long] $budgets.maximumBrowserHeapBytes) {
        throw "Converter browser heap reached $($result.maximumBrowserHeapBytes) bytes; budget is $($budgets.maximumBrowserHeapBytes) bytes."
    }
    $expectedRouteIds = @($budgets.routes.PSObject.Properties.Name | Sort-Object)
    $actualRouteIds = @($result.routes.routeId | Sort-Object)
    if (($expectedRouteIds -join "`n") -ne ($actualRouteIds -join "`n")) {
        throw "Representative route measurements differ from configured budgets. Expected '$($expectedRouteIds -join ', ')'; actual '$($actualRouteIds -join ', ')'."
    }
    foreach ($route in $result.routes) {
        $routeBudget = $budgets.routes.($route.routeId)
        if ($null -eq $routeBudget) { throw "No performance budget exists for route '$($route.routeId)'." }
        if ([long] $route.conversionMilliseconds -gt [long] $routeBudget.maximumConversionMilliseconds) {
            throw "Route '$($route.routeId)' took $($route.conversionMilliseconds) ms; budget is $($routeBudget.maximumConversionMilliseconds) ms."
        }
        if ([long] $route.peakRetainedBytes -gt [long] $routeBudget.maximumPeakRetainedBytes) {
            throw "Route '$($route.routeId)' retained $($route.peakRetainedBytes) bytes; budget is $($routeBudget.maximumPeakRetainedBytes) bytes."
        }
        if ([long] $route.repeatConversionMilliseconds -gt [long] $routeBudget.maximumConversionMilliseconds) {
            throw "Repeated route '$($route.routeId)' took $($route.repeatConversionMilliseconds) ms; budget is $($routeBudget.maximumConversionMilliseconds) ms."
        }
        if ([long] $route.repeatPeakRetainedBytes -gt [long] $routeBudget.maximumPeakRetainedBytes) {
            throw "Repeated route '$($route.routeId)' retained $($route.repeatPeakRetainedBytes) bytes; budget is $($routeBudget.maximumPeakRetainedBytes) bytes."
        }
        if ([long] $route.resultBytes -le 0) { throw "Route '$($route.routeId)' produced an empty result." }
        if ([long] $route.repeatResultBytes -le 0) { throw "Repeated route '$($route.routeId)' produced an empty result." }
        if ([string] $route.pdfMagic -ne '%PDF' -or [string] $route.repeatPdfMagic -ne '%PDF') {
            throw "Route '$($route.routeId)' did not produce PDF bytes in both cold and repeated conversions."
        }
        if ([long] $route.memorySamples -lt 2 -or [long] $route.peakBrowserHeapBytes -le 0) {
            throw "Route '$($route.routeId)' did not produce fail-closed in-flight browser heap evidence."
        }
        if ([long] $route.repeatMemorySamples -lt 2 -or [long] $route.repeatPeakBrowserHeapBytes -le 0) {
            throw "Repeated route '$($route.routeId)' did not produce fail-closed in-flight browser heap evidence."
        }
    }
    if (-not $result.webMcp) {
        throw 'The browser run did not capture the convert_selected_document Website Tool.'
    }
    if (-not [bool] $result.webMcpLifecycle.removedOutsideConverter -or -not [bool] $result.webMcpLifecycle.restoredWithConverter) {
        throw 'The converter Website Tool did not follow the visible converter workspace lifecycle.'
    }
    if (-not [bool] $result.webMcp.output.success -or [string] $result.webMcp.output.route -ne 'docx-pdf') {
        throw 'The convert_selected_document Website Tool did not complete the visible DOCX-to-PDF route.'
    }
    if ([long] $result.webMcp.outputCharacters -gt 1500) {
        throw "The converter Website Tool returned $($result.webMcp.outputCharacters) characters; the limit is 1500."
    }
    if ([bool] $result.webMcp.annotations.readOnlyHint -or [bool] $result.webMcp.annotations.destructiveHint -or -not [bool] $result.webMcp.annotations.untrustedContentHint) {
        throw 'The converter Website Tool annotations do not describe a local, non-destructive conversion of user-provided content.'
    }
    if ([bool] $result.webMcp.cancelled.success -or [string] $result.webMcp.cancelled.message -ne 'Conversion was cancelled before it started.') {
        throw 'The converter Website Tool did not honor a caller cancellation before conversion.'
    }
    if ($null -ne $result.webMcp.schema.properties.PSObject.Properties -and @($result.webMcp.schema.properties.PSObject.Properties).Count -gt 0) {
        throw 'The converter Website Tool must not accept file paths, bytes, or other input parameters.'
    }
    if (-not [bool] $result.longNameWebMcp.output.success -or
        [long] $result.longNameWebMcp.outputCharacters -gt 1500 -or
        [long] $result.longNameWebMcp.outputFileNameCharacters -gt 180 -or
        [bool] $result.longNameWebMcp.hasUnpairedSurrogate) {
        $longNameEvidence = $result.longNameWebMcp | ConvertTo-Json -Depth 6 -Compress
        throw "The converter Website Tool did not return a Unicode-safe bounded filename. Evidence: $longNameEvidence"
    }
    if ([bool] $result.malformedWebMcp.output.success -or [long] $result.malformedWebMcp.outputCharacters -gt 1500 -or [string] $result.malformedWebMcp.visibleDiagnostics -notmatch 'Conversion failed') {
        $malformedEvidence = $result.malformedWebMcp | ConvertTo-Json -Depth 6 -Compress
        throw "The converter Website Tool did not return a bounded failure synchronized with the visible workspace. Evidence: $malformedEvidence"
    }
    if (@($result.consoleErrors).Count -gt 0) {
        throw "Browser converter emitted console errors: $($result.consoleErrors -join ' | ')"
    }

    $report = [ordered] @{
        schemaVersion = 1
        measuredAtUtc = [DateTimeOffset]::UtcNow.ToString('O')
        publishedBytes = [long] $publishedBytes
        startupMilliseconds = [double] $result.startupMilliseconds
        maximumBrowserHeapBytes = [long] $result.maximumBrowserHeapBytes
        playwrightCliVersion = [string] $budgets.playwrightCliVersion
        browserEngine = $browserEngine
        routes = $result.routes
        webMcp = $result.webMcp
        webMcpLifecycle = $result.webMcpLifecycle
        longNameWebMcp = $result.longNameWebMcp
        malformedWebMcp = $result.malformedWebMcp
        budgets = $budgets
    }
    $reportDirectory = Split-Path -Parent $ReportPath
    if ($reportDirectory) { New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ReportPath -Encoding utf8
    Write-Output "Browser converter performance verified: $publishedBytes bytes, $([Math]::Round($result.startupMilliseconds)) ms startup, $($result.routes.Count) representative routes."
} finally {
    $npxCommand = Get-Command npx -ErrorAction SilentlyContinue
    if ($npxCommand) { & $npxCommand.Source --yes --package $playwrightPackage playwright-cli "-s=$session" close 2>$null | Out-Null }
    if ($server -and -not $server.HasExited) {
        Stop-Process -Id $server.Id -Force
        $server.WaitForExit(5000) | Out-Null
    }
    if ($server) { $server.Dispose() }
    [System.IO.File]::Delete($serverStandardOutputPath)
    [System.IO.File]::Delete($serverStandardErrorPath)
}
