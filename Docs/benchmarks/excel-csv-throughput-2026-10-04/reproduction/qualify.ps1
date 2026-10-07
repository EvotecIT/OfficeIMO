$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
Set-Location (Split-Path (Split-Path (Split-Path $root)))
$pending = @(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'dotnet.exe' -and $_.CommandLine -like '*Ignore/Benchmarks/excel-csv-throughput/compare/bin/Release/net10.0/Compare.dll --filter*' })
foreach ($process in $pending) { Wait-Process -Id $process.ProcessId -ErrorAction SilentlyContinue }
$env:OFFICEIMO_BENCHMARK_DATA = Join-Path $root 'fixtures'
$env:OFFICEIMO_BENCHMARK_OUTPUT = Join-Path $root 'file-output'
foreach ($mask in @('65535', '4294901760')) {
    & pwsh -NoProfile -File (Join-Path $root 'rotated.ps1') -Mask $mask -Iterations 24 -Warmup 12 -Exports 32 -ScenarioList 'CsvShortJsonAsNeeded,CsvShortJsonAlways,CsvShortQuotesAsNeeded,ExcelSstAscii,ExcelSstUnicodeTail,ExcelSstEntityTail' -Label 'rotated-small' *> (Join-Path $root "rotated-small-$mask.log")
    if ($LASTEXITCODE -ne 0) { throw "Rotated small failed on $mask" }
    & pwsh -NoProfile -File (Join-Path $root 'rotated.ps1') -Mask $mask -Iterations 24 -Warmup 12 -Exports 4 -ScenarioList 'CsvLongJsonAsNeeded,Csv25K' -Label 'rotated-large' *> (Join-Path $root "rotated-large-$mask.log")
    if ($LASTEXITCODE -ne 0) { throw "Rotated large failed on $mask" }
}
$common = @('--priority','Normal','--affinityMasks','0xFFFF,0xFFFF0000','--unrollFactor','1','--warmupCount','6','--iterationCount','12','--launchCount','1','--outliers','DontRemove')
& dotnet OfficeIMO.CSV.Benchmarks/bin/Release/net10.0/OfficeIMO.CSV.Benchmarks.dll --filter '*CsvTextWriteBenchmarks*64*Json*' --invocationCount 1024 @common --artifacts (Join-Path $root 'peers-csv-json') *> (Join-Path $root 'peers-csv-json.log')
if ($LASTEXITCODE -ne 0) { throw 'CSV JSON peers failed' }
& dotnet OfficeIMO.CSV.Benchmarks/bin/Release/net10.0/OfficeIMO.CSV.Benchmarks.dll --filter '*CsvDataReaderWriteBenchmarks.OfficeIMO_WriteDataReader*25000*Mixed*' '*CsvDataReaderWriteBenchmarks.Sylvan_WriteDataReader*25000*Mixed*' '*MarkPflug65KCsvBenchmarks.OfficeIMO' '*MarkPflug65KCsvBenchmarks.Sep' '*MarkPflug65KCsvBenchmarks.Sylvan' --invocationCount 8 @common --artifacts (Join-Path $root 'peers-csv-large') *> (Join-Path $root 'peers-csv-large.log')
if ($LASTEXITCODE -ne 0) { throw 'CSV large peers failed' }
& dotnet OfficeIMO.Excel.Benchmarks/bin/Release/net10.0/OfficeIMO.Excel.Benchmarks.dll --filter '*ExcelSharedStringReadBenchmarks*' --invocationCount 256 @common --artifacts (Join-Path $root 'peers-excel-sst') *> (Join-Path $root 'peers-excel-sst.log')
if ($LASTEXITCODE -ne 0) { throw 'Excel shared-string peers failed' }
& dotnet OfficeIMO.Excel.Benchmarks/bin/Release/net10.0/OfficeIMO.Excel.Benchmarks.dll --filter '*MarkPflug65KXlsxBenchmarks.OfficeIMO' '*MarkPflug65KXlsxBenchmarks.Sylvan' '*MarkPflug65KXlsxBenchmarks.ExcelReaderNet' '*ExcelTextWriteBenchmarks*4096*Plain*' --invocationCount 4 @common --artifacts (Join-Path $root 'peers-excel-large') *> (Join-Path $root 'peers-excel-large.log')
if ($LASTEXITCODE -ne 0) { throw 'Excel large peers failed' }
foreach ($owner in @('CSV','Excel')) {
    & dotnet build "OfficeIMO.$owner/OfficeIMO.$owner.csproj" -c Release -f netstandard2.0 *> (Join-Path $root "build-$owner-netstandard.log")
    if ($LASTEXITCODE -ne 0) { throw "$owner netstandard2.0 build failed" }
}
Write-Output 'Qualification sequence complete; inspect every benchmark report for successful case counts.'
