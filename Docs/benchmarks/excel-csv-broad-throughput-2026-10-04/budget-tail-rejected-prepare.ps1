param([Parameter(Mandatory)][string]$BaselineRoot, [Parameter(Mandatory)][string]$Workspace)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$expectedChanges = @(
    'OfficeIMO.Excel/Read/ExcelSheetReader.Range.DataReader.Utf8.cs',
    'OfficeIMO.Excel/Read/ExcelSheetReader.Range.DataReader.Utf8.Budget.cs'
)
foreach ($tfm in 'net10.0', 'net8.0') {
    $old = Get-Content -LiteralPath (Join-Path $BaselineRoot "manifest-$tfm.json") -Raw | ConvertFrom-Json
    $beforeSource = $old.AfterSource
    $paths = @($beforeSource.Source.Path) + 'OfficeIMO.Excel/Read/ExcelSheetReader.Range.DataReader.Utf8.Budget.cs'
    $afterSource = @(foreach ($path in $paths) {
        $bytes = [Text.Encoding]::UTF8.GetBytes([IO.File]::ReadAllText((Join-Path $Workspace $path)).Replace("`r`n", "`n"))
        [pscustomobject]@{ Path = $path; Sha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)) }
    })
    $beforeHashes = @{}
    foreach ($file in $beforeSource.Source) { $beforeHashes[$file.Path] = $file.Sha256 }
    $changed = @($afterSource | Where-Object { $beforeHashes[$_.Path] -ne $_.Sha256 } | ForEach-Object Path)
    if ($changed.Count -ne $expectedChanges.Count -or @($changed | Where-Object { $_ -notin $expectedChanges }).Count) { throw 'Unexpected corrected candidate source difference' }
    $baseline = Join-Path $BaselineRoot "$tfm/candidate-excel"
    foreach ($file in $old.Manifest) {
        if ((Get-FileHash -LiteralPath (Join-Path $baseline $file.Name)).Hash -ne $file.After) { throw 'Qualified baseline snapshot changed' }
    }
    foreach ($side in 'baseline-excel', 'candidate-excel') {
        $destination = Join-Path $root "$tfm/$side"
        if (Test-Path -LiteralPath $destination) { throw 'Refuse existing corrected dimensionless snapshot' }
        New-Item -ItemType Directory -Path $destination | Out-Null
        Get-ChildItem -LiteralPath $baseline -File | Copy-Item -Destination $destination
        if (Test-Path -LiteralPath (Join-Path $baseline 'runtimes')) { Copy-Item -LiteralPath (Join-Path $baseline 'runtimes') -Destination $destination -Recurse }
    }
    foreach ($extension in 'dll', 'pdb') {
        Copy-Item -LiteralPath (Join-Path $Workspace "OfficeIMO.Excel/bin/Release/$tfm/OfficeIMO.Excel.$extension") -Destination (Join-Path $root "$tfm/candidate-excel/OfficeIMO.Excel.$extension")
    }
    $files = @(Get-ChildItem -LiteralPath (Join-Path $root "$tfm/baseline-excel") -File | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Before = (Get-FileHash -LiteralPath $_.FullName).Hash; After = (Get-FileHash -LiteralPath (Join-Path $root "$tfm/candidate-excel/$($_.Name)")).Hash }
    })
    if (@($files | Where-Object { $_.Before -ne $_.After -and $_.Name -notin 'OfficeIMO.Excel.dll', 'OfficeIMO.Excel.pdb' }).Count) { throw 'Harness or dependency differs' }
    [ordered]@{
        Runtime = $tfm
        BeforeSource = $beforeSource
        BaselineContract = 'Qualified undimensioned index candidate 1b0c3b1, normalized Excel/Core source equal to integrated b57455a0c'
        AfterSource = [ordered]@{ Head = (& git -C $Workspace rev-parse HEAD); Status = @(& git -C $Workspace status --porcelain); Source = $afterSource }
        ChangedSource = $changed
        RegressionTest = @{ Path = 'OfficeIMO.Excel.Tests/Excel.DataReaderApi.DimensionlessWorksheet.cs'; Sha256 = (Get-FileHash -LiteralPath (Join-Path $Workspace 'OfficeIMO.Excel.Tests/Excel.DataReaderApi.DimensionlessWorksheet.cs')).Hash }
        Manifest = $files
        Contract = 'Two private production paths differ; same fixture harness and dependencies; fragments only decline optional indexing; existing complete streaming validation remains authoritative; large eligible control retained'
    } | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $root "manifest-$tfm.json") -Encoding utf8
    Write-Output "Corrected dimensionless snapshots qualified: $tfm, two production paths plus regression test hash, only Excel DLL/PDB differ"
}
