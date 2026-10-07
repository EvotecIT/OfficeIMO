param([Parameter(Mandatory)][string]$Workspace)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
foreach ($tfm in 'net10.0','net8.0') {
    $before=Get-Content -LiteralPath (Join-Path $root "before-$tfm.json") -Raw | ConvertFrom-Json
    $source=@(foreach ($file in $before.BeforeSource.Source) {
        $bytes=[Text.Encoding]::UTF8.GetBytes([IO.File]::ReadAllText((Join-Path $Workspace $file.Path)).Replace("`r`n","`n"))
        [pscustomobject]@{Path=$file.Path;Sha256=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))}
    })
    $hashes=@{}
    foreach ($file in $before.BeforeSource.Source) { $hashes[$file.Path]=$file.Sha256 }
    $changed=@($source | Where-Object { $hashes[$_.Path] -ne $_.Sha256 } | ForEach-Object Path)
    $expected=@('OfficeIMO.CSV/CsvRowWriter.cs','OfficeIMO.CSV/CsvRowWriter.DataReader.cs')
    if ($changed.Count -ne 2 -or @($changed | Where-Object {$_ -notin $expected}).Count) { throw 'Unexpected CSV production or benchmark source differences' }
    foreach ($file in $before.Binaries) {
        if ((Get-FileHash -LiteralPath (Join-Path $root "$tfm/baseline-csv/$($file.Name)")).Hash -ne $file.Sha256) { throw 'CSV baseline changed' }
    }
    foreach ($extension in 'dll','pdb') {
        Copy-Item -LiteralPath (Join-Path $Workspace "OfficeIMO.CSV/bin/Release/$tfm/OfficeIMO.CSV.$extension") -Destination (Join-Path $root "$tfm/candidate-csv/OfficeIMO.CSV.$extension")
    }
    $manifest=@(foreach ($file in $before.Binaries) {
        [pscustomobject]@{Name=$file.Name;Before=$file.Sha256;After=(Get-FileHash -LiteralPath (Join-Path $root "$tfm/candidate-csv/$($file.Name)")).Hash}
    })
    if (@($manifest | Where-Object { $_.Before -ne $_.After -and $_.Name -notin 'OfficeIMO.CSV.dll','OfficeIMO.CSV.pdb' }).Count) { throw 'CSV comparison harness or dependency differs' }
    [ordered]@{Runtime=$tfm;BeforeSource=$before.BeforeSource;AfterSource=@{Head=(& git -C $Workspace rev-parse HEAD);Status=@(& git -C $Workspace status --porcelain);Source=$source};Manifest=$manifest;ChangedSource=$changed;RegressionTestSha256=(Get-FileHash -LiteralPath (Join-Path $Workspace 'OfficeIMO.CSV.Tests/CsvDataReaderWriterRegressionTests.cs')).Hash;Contract='Public DataReader text-delimiter batching; same harness, dependencies, buffers and complete output validation; partial failure/cancellation contracts qualified separately'} | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $root "manifest-$tfm.json") -Encoding utf8
    Write-Output "CSV candidate qualified: $tfm, two production paths, only CSV DLL/PDB differ"
}
