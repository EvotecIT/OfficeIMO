param([Parameter(Mandatory)][string]$Workspace,[Parameter(Mandatory)][string]$QualifiedBaselineRoot)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$production=@('OfficeIMO.CSV/CsvWriter.cs','OfficeIMO.CSV/CsvWriter.Escaping.cs')
$harness='OfficeIMO.CSV.Benchmarks/CsvTextWriteBenchmarks.cs'
$head=(& git -C $Workspace rev-parse HEAD)
if(@(& git -C $Workspace status --porcelain).Count){throw 'Candidate source must be committed and clean'}
foreach($tfm in 'net10.0','net8.0'){
    $old=Get-Content -LiteralPath (Join-Path $QualifiedBaselineRoot "before-$tfm.json") -Raw|ConvertFrom-Json
    $oldHashes=@{}
    foreach($file in $old.BeforeSource.Source){$oldHashes[$file.Path]=$file.Sha256}
    $source=@(foreach($file in $old.BeforeSource.Source){
        $text=[IO.File]::ReadAllText((Join-Path $Workspace $file.Path)).Replace("`r`n","`n")
        [pscustomobject]@{Path=$file.Path;Sha256=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($text)))}
    })
    $changed=@($source|Where-Object {$oldHashes[$_.Path] -ne $_.Sha256}|ForEach-Object Path)
    if($changed.Count -ne 3 -or @($changed|Where-Object {$_ -notin ($production+$harness)}).Count){throw 'Unexpected source or fixture difference'}
    $baseline=Join-Path $QualifiedBaselineRoot "$tfm/baseline-csv"
    foreach($file in $old.Binaries){
        if((Get-FileHash -LiteralPath (Join-Path $baseline $file.Name)).Hash -ne $file.Sha256){throw 'Qualified baseline artifact changed'}
    }
    foreach($side in 'baseline-csv','candidate-csv'){
        $destination=Join-Path $root "$tfm/$side"
        if(Test-Path -LiteralPath $destination){throw 'Refuse existing delimiter-search snapshots'}
        New-Item -ItemType Directory -Path $destination|Out-Null
        Get-ChildItem -LiteralPath $baseline -File|Copy-Item -Destination $destination
        if(Test-Path -LiteralPath (Join-Path $baseline 'runtimes')){Copy-Item -LiteralPath (Join-Path $baseline 'runtimes') -Destination $destination -Recurse}
        foreach($extension in 'dll','pdb'){
            Copy-Item -LiteralPath (Join-Path $Workspace "OfficeIMO.CSV.Benchmarks/bin/Release/$tfm/OfficeIMO.CSV.Benchmarks.$extension") -Destination $destination
            if($side -eq 'candidate-csv'){
                Copy-Item -LiteralPath (Join-Path $Workspace "OfficeIMO.CSV/bin/Release/$tfm/OfficeIMO.CSV.$extension") -Destination $destination
            }
        }
    }
    $manifest=@(foreach($file in $old.Binaries){
        [pscustomobject]@{Name=$file.Name;Before=(Get-FileHash -LiteralPath (Join-Path $root "$tfm/baseline-csv/$($file.Name)")).Hash;After=(Get-FileHash -LiteralPath (Join-Path $root "$tfm/candidate-csv/$($file.Name)")).Hash}
    })
    if(@($manifest|Where-Object {$_.Before -ne $_.After -and $_.Name -notin 'OfficeIMO.CSV.dll','OfficeIMO.CSV.pdb'}).Count){throw 'Harness or dependency differs between engines'}
    foreach($file in $old.Binaries){
        $current=@($manifest|Where-Object Name -eq $file.Name)[0]
        if($current.Before -ne $file.Sha256 -and $file.Name -notin 'OfficeIMO.CSV.Benchmarks.dll','OfficeIMO.CSV.Benchmarks.pdb'){throw 'Baseline product or dependency differs from qualified artifact'}
    }
    $beforeSource=@(foreach($file in $source){
        [pscustomobject]@{Path=$file.Path;Sha256=$(if($file.Path -in $production){$oldHashes[$file.Path]}else{$file.Sha256})}
    })
    [ordered]@{
        Runtime=$tfm
        BeforeSource=@{ProductSourceHead=$old.BeforeSource.Head;SharedHarnessSourceHead=$head;Source=$beforeSource;Contract='Qualified original CSV DLL and dependencies; updated fixture harness identical in both engines'}
        AfterSource=@{Head=$head;Status=@();Source=$source}
        ChangedProduction=$production
        SharedHarnessChange=$harness
        Manifest=$manifest
        RegressionTests=@(foreach($test in 'CsvDataReaderWriterRegressionTests.cs','CsvEscapingRegressionTests.cs'){@{Path="OfficeIMO.CSV.Tests/$test";Sha256=(Get-FileHash -LiteralPath (Join-Path $Workspace "OfficeIMO.CSV.Tests/$test")).Hash}})
        Contract='Only CSV DLL/PDB differ between engines; every other artifact and new plain/delimited fixture is identical; original qualified baseline binaries retained; complete text/byte and decoded-field validation required'
    }|ConvertTo-Json -Depth 16|Set-Content -LiteralPath (Join-Path $root "manifest-$tfm.json") -Encoding utf8
    Write-Output "Delimiter snapshots qualified: $tfm, two product paths, one shared fixture extension, only CSV DLL/PDB differ"
}
