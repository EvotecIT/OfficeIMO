param([Parameter(Mandatory)][string] $EvidenceDirectory)
$ErrorActionPreference = 'Stop'
$nodeCommand = (Get-Command node -CommandType Application | Select-Object -First 1).Source
$scriptShell = (Get-Command pwsh -CommandType Application | Select-Object -First 1).Source
$repository = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$npmCommand = (Get-Command npm.cmd,npm -ErrorAction SilentlyContinue | Select-Object -First 1).Source
if (-not $npmCommand) { throw 'npm is required for the isolated consumer check.' }
$output = [IO.Path]::GetFullPath($EvidenceDirectory)
New-Item -ItemType Directory -Path $output -Force | Out-Null
$package = Join-Path $repository 'OfficeIMO.Browser'
Push-Location $package
try {
    $packed = & $npmCommand pack --script-shell $scriptShell --pack-destination $output --json
    if ($LASTEXITCODE -ne 0) { throw 'npm pack failed.' }
    $manifest = ($packed -join "`n") | ConvertFrom-Json
    $manifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $output 'archive.json') -Encoding utf8
} finally { Pop-Location }
$archive = Join-Path $output $manifest[0].filename
$consumer = Join-Path $output 'consumer'
New-Item -ItemType Directory -Path $consumer -Force | Out-Null
'{"private":true,"type":"module"}' | Set-Content -LiteralPath (Join-Path $consumer 'package.json') -Encoding utf8
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'typed-consumer.mts'),(Join-Path $PSScriptRoot 'npm-consumer.mjs') -Destination $consumer
Push-Location $consumer
try {
    & $npmCommand install --ignore-scripts --no-audit --no-fund --save-exact $archive typescript@5.9.3
    if ($LASTEXITCODE -ne 0) { throw 'Isolated archive install failed.' }
    & $nodeCommand node_modules/typescript/bin/tsc --strict --noEmit --target ES2022 --module NodeNext --moduleResolution NodeNext --lib ES2022,DOM typed-consumer.mts
    if ($LASTEXITCODE -ne 0) { throw 'Strict declaration consumer failed.' }
    & $nodeCommand npm-consumer.mjs
    if ($LASTEXITCODE -ne 0) { throw 'Packed runtime consumer failed.' }
} finally { Pop-Location }
Copy-Item -LiteralPath (Join-Path $consumer 'packed-consumer.xlsx') -Destination $output
@{ passed = $true; strictTypeScript = '5.9.3'; archive = $manifest[0].filename; runtimeWrites = $true } |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'consumer-report.json') -Encoding utf8
Write-Host "Packed npm archive and strict isolated consumer passed: $archive"
