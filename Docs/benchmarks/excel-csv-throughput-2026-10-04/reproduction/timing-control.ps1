$ErrorActionPreference = 'Stop'
$pending = @(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'pwsh.exe' -and $_.CommandLine -like '*-File Ignore/Benchmarks/excel-csv-throughput/qualify.ps1' })
foreach ($process in $pending) { Wait-Process -Id $process.ProcessId -ErrorAction SilentlyContinue }
foreach ($mask in @('65535', '4294901760')) {
    foreach ($mode in @('aa', 'ab')) {
        $env:OFFICEIMO_PERFORMANCE_AA = if ($mode -eq 'aa') { '1' } else { '0' }
        & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'rotated.ps1') -Mask $mask -Iterations 32 -Warmup 16 -Exports 32 -ScenarioList 'CsvShortJsonAlways,CsvShortJsonAsNeeded,ExcelSstAscii' -Label "control-small-$mode" *> (Join-Path $PSScriptRoot "control-small-$mode-$mask.log")
        if ($LASTEXITCODE -ne 0) { throw "Small $mode control failed on $mask" }
        & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'rotated.ps1') -Mask $mask -Iterations 32 -Warmup 16 -Exports 4 -ScenarioList 'Csv25K,CsvLongJsonAsNeeded' -Label "control-large-$mode" *> (Join-Path $PSScriptRoot "control-large-$mode-$mask.log")
        if ($LASTEXITCODE -ne 0) { throw "Large $mode control failed on $mask" }
    }
}
Write-Output 'Timing controls complete.'
