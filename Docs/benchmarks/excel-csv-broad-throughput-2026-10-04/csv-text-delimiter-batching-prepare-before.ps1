param([Parameter(Mandatory)][string]$BaselineWorkspace, [Parameter(Mandatory)][string]$Workspace, [Parameter(Mandatory)][string]$BaselineEvidence)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$packet = Get-Content -LiteralPath $BaselineEvidence -Raw | ConvertFrom-Json
foreach ($file in $packet.Source) {
    foreach ($workspacePath in $BaselineWorkspace, $Workspace) {
        $bytes = [Text.Encoding]::UTF8.GetBytes([IO.File]::ReadAllText((Join-Path $workspacePath $file.Path)).Replace("`r`n", "`n"))
        if ([Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)) -ne $file.Sha256) { throw "CSV baseline source differs: $($file.Path)" }
    }
}
foreach ($tfm in 'net10.0', 'net8.0') {
    $binaryRoot = Join-Path $BaselineWorkspace "OfficeIMO.CSV.Benchmarks/bin/Release/$tfm"
    $binaries = @($packet.Reports | Where-Object Runtime -eq $tfm | Select-Object -First 1).Binaries
    if (-not $binaries.Count) { throw 'No qualified CSV baseline binaries' }
    foreach ($file in $binaries) {
        if ((Get-FileHash -LiteralPath (Join-Path $binaryRoot $file.Name)).Hash -ne $file.Sha256) { throw "Qualified CSV baseline binary differs: $($file.Name)" }
    }
    foreach ($side in 'baseline-csv', 'candidate-csv') {
        $destination = Join-Path $root "$tfm/$side"
        if (Test-Path -LiteralPath $destination) { throw 'Refuse existing CSV snapshot' }
        New-Item -ItemType Directory -Path $destination | Out-Null
        Get-ChildItem -LiteralPath $binaryRoot -File | Copy-Item -Destination $destination
        if (Test-Path -LiteralPath (Join-Path $binaryRoot 'runtimes')) { Copy-Item -LiteralPath (Join-Path $binaryRoot 'runtimes') -Destination $destination -Recurse }
    }
    [ordered]@{ Runtime=$tfm; BeforeSource=@{ Head=$packet.Head; Source=$packet.Source }; Binaries=$binaries; EvidenceSha256=(Get-FileHash -LiteralPath $BaselineEvidence).Hash; BeforeTestSha256=(Get-FileHash -LiteralPath (Join-Path $Workspace 'OfficeIMO.CSV.Tests/CsvDataReaderWriterRegressionTests.cs')).Hash } | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $root "before-$tfm.json") -Encoding utf8
    Write-Output "Qualified CSV baseline frozen: $tfm"
}
