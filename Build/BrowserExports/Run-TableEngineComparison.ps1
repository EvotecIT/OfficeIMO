param(
    [Parameter(Mandatory)][string] $OfficeImoModule,
    [Parameter(Mandatory)][string] $Dependencies,
    [Parameter(Mandatory)][string] $LicenseReceipt,
    [Parameter(Mandatory)][string] $Binary,
    [Parameter(Mandatory)][string] $OutputRoot,
    [string] $Qpdf,
    [ValidateSet('fresh','cold','warm')][string] $Mode = 'fresh',
    [string[]] $Cases,
    [ValidateSet('officeimo','papaparse','exceljs','exceljs-stream','sheetjs','pdfmake','jspdf','pdfkit')]
    [string[]] $Engine = @('officeimo','papaparse','exceljs','exceljs-stream','sheetjs','pdfmake','jspdf','pdfkit'),
    [ValidateSet('unicode','bmp')][string] $TextProfile = 'unicode',
    [switch] $Unique,
    [ValidateRange(1,10)][int] $WarmupCount = 1,
    [ValidateRange(1,20)][int] $IterationCount = 3,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
Import-Module PSPublishModule -ErrorAction Stop
$node = (Get-Command node -ErrorAction Stop).Source
$dotnet = (Get-Command dotnet -ErrorAction Stop).Source
$helper = Join-Path $PSScriptRoot 'table-engine-comparison.mjs'
$repository = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$evidence = [IO.Path]::GetFullPath($OutputRoot)
$owner = [guid]::NewGuid().ToString('N')
$manifest = Join-Path $evidence 'table-engines.json'
$binaryPath = (Resolve-Path -LiteralPath $Binary).Path
$qpdfPath = if ($Qpdf) { (Resolve-Path -LiteralPath $Qpdf).Path } else { (Get-Command qpdf -CommandType Application -ErrorAction Stop).Source }
$qpdfHash = (Get-FileHash -LiteralPath $qpdfPath -Algorithm SHA256).Hash
$uniqueLiteral = $Unique.IsPresent.ToString().ToLowerInvariant()
if (-not $Cases) {
    $Cases = if ($Mode -eq 'fresh') { @('flat-csv-100000x6-buffer','flat-csv-100000x20-buffer',
        'flat-xlsx-10000x6-buffer','flat-xlsx-100000x20-buffer','flat-pdf-10000x6-buffer','flat-pdf-2000x20-buffer') }
        else { @('bootstrap-csv-1000x6-buffer','bootstrap-xlsx-1000x6-buffer','bootstrap-pdf-1000x6-buffer') }
}
# Preparation only reads existing exact packages/receipts. It never installs dependencies.
& $node $helper prepare --officeimo (Resolve-Path -LiteralPath $OfficeImoModule).Path `
    --dependencies (Resolve-Path -LiteralPath $Dependencies).Path --license-receipt (Resolve-Path -LiteralPath $LicenseReceipt).Path `
    --evidence $evidence --engines ($Engine -join ',') --cases ($Cases -join ',') --unique $uniqueLiteral `
    --text-profile $TextProfile --font-regular (Join-Path $repository 'Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf') `
    --font-bold (Join-Path $repository 'Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Bold.ttf')
if ($LASTEXITCODE -ne 0) { throw 'Comparison preparation failed; no cohort ran.' }
if (-not ('DataTablesBenchmarkClient' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'DataTablesBenchmarkClient.cs') }
try {
    $result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'table-engines.benchmark.ps1') -OutputRoot $evidence `
        -Variable @{ Node = $node; Helper = $helper; Manifest = $manifest; Binary = $binaryPath; Dotnet = $dotnet; Evidence = $evidence
            Mode = $Mode; Owner = $owner; Iterations = $IterationCount; Qpdf = $qpdfPath; QpdfHash = $qpdfHash } `
        -Engine $Engine -WarmupCount $WarmupCount -IterationCount $IterationCount -Plan:$Plan
    if ($Plan) { $result | Select-Object Scenario,Engine,Operation,IsSkipped; return }
    # Successful shared-runner warmups bypass its validation callback. Qualify their actual bytes here.
    foreach ($artifact in Get-ChildItem -LiteralPath $evidence -Filter 'sample.*' -Recurse -File) {
        $iteration = 0
        if (-not [int]::TryParse($artifact.Directory.Name, [ref] $iteration) -or $iteration -ge 0) { continue }
        $measurementPath = Join-Path $artifact.DirectoryName 'measurement.json'
        if (-not (Test-Path -LiteralPath $measurementPath -PathType Leaf)) { continue }
        $measurement = Get-Content -LiteralPath $measurementPath -Raw | ConvertFrom-Json
        $uniqueValue = ([bool] $measurement.unique).ToString().ToLowerInvariant()
        $json = & $dotnet $binaryPath --validate-table-engine $artifact.FullName "--format=$($measurement.format)" `
            "--rows=$($measurement.rows)" "--columns=$($measurement.columns)" "--unique=$uniqueValue" "--text-profile=$($measurement.textProfile)"
        if ($LASTEXITCODE -ne 0) { throw 'Independent warmup readback failed; artifact retained.' }
        $proof = $json | ConvertFrom-Json
        if (-not $proof.passed -or $proof.sha256 -ne $measurement.outputSha256 -or $proof.rows -ne $measurement.rows -or
            $proof.columns -ne $measurement.columns -or $proof.format -ne $measurement.format) { throw 'Warmup proof does not match its complete artifact.' }
        if ($measurement.format -eq 'pdf') {
            $syntaxLog = Join-Path $artifact.DirectoryName 'qpdf-check.log'
            $syntaxOutput = & $qpdfPath --check $artifact.FullName 2>&1
            $syntaxExit = $LASTEXITCODE
            $syntaxOutput | Set-Content -LiteralPath $syntaxLog -Encoding utf8
            $syntaxProof = @{ passed = $syntaxExit -eq 0; exitCode = $syntaxExit; binarySha256 = $qpdfHash; outputSha256 = $measurement.outputSha256; log = $syntaxLog }
            $syntaxProof | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $artifact.DirectoryName 'qpdf-check.json') -Encoding utf8
            $proof | Add-Member -NotePropertyName qpdf -NotePropertyValue $syntaxProof
            if ($syntaxExit -ne 0) {
                $proof | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath (Join-Path $artifact.DirectoryName 'warmup-validation.json') -Encoding utf8
                throw 'Warmup PDF syntax check failed; artifact, reader proof and qpdf log retained.'
            }
        }
        $proof | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath (Join-Path $artifact.DirectoryName 'warmup-validation.json') -Encoding utf8
        Remove-Item -LiteralPath $artifact.FullName -ErrorAction Stop
    }
    & $node $helper verify --manifest $manifest | Set-Content -LiteralPath (Join-Path $evidence 'graph-final.json') -Encoding utf8
    if ($LASTEXITCODE -ne 0) { throw 'Comparison inputs changed during the cohort.' }
    if ((Get-FileHash -LiteralPath $qpdfPath -Algorithm SHA256).Hash -ne $qpdfHash) { throw 'Selected qpdf binary changed during the cohort.' }
    $result
    if (@($result.Summary | Where-Object { $_.Status -notin 'Succeeded','Skipped' }).Count) { throw 'A comparison failed; retain all failed artifacts and diagnostics.' }
} finally {
    $closed = @()
    foreach ($key in @([DataTablesBenchmarkClient]::Sessions.Keys | Where-Object { $_.StartsWith($owner + '|', [StringComparison]::Ordinal) })) {
        $client = [DataTablesBenchmarkClient]::Sessions[$key]
        $childPid = $client.ProcessId
        try { $client.Dispose(); $closed += [pscustomobject]@{ key = $key; pid = $childPid; disposed = $true } }
        finally { [void] [DataTablesBenchmarkClient]::Sessions.Remove($key) }
    }
    if (Test-Path -LiteralPath $evidence -PathType Container) {
        [pscustomobject]@{ owner = $owner; mode = $Mode; closedByFinally = $closed; remainingOwnedSessions =
            @([DataTablesBenchmarkClient]::Sessions.Keys | Where-Object { $_.StartsWith($owner + '|', [StringComparison]::Ordinal) }).Count } |
            ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $evidence 'session-closeout.json') -Encoding utf8
    }
}
