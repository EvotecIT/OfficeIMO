param(
    [string] $Engine,
    [string] $ModuleManifest,
    [string] $InputPath,
    [string] $OutputPath,
    [string] $ObserveMemory
)

$ErrorActionPreference = 'Stop'
Import-Module -Name $ModuleManifest
if ($Engine -eq 'PSWriteOffice') {
    $null = PSWriteOffice\Convert-OfficeDocument -LiteralPath $InputPath -OutputPath $OutputPath -FailOnLoss
} else {
    $null = PowerShellNotebook\ConvertFrom-IPYNB -Path $InputPath -Destination $OutputPath -IncludeOutput
}

if ($ObserveMemory -eq 'True') {
    $process = [Diagnostics.Process]::GetCurrentProcess()
    try {
        # A zero value means the operating system did not supply the peak.
        $peak = $process.PeakWorkingSet64
        if ($peak -gt 0) {
            [Console]::WriteLine("WorkerPeakWorkingSetBytes=$peak")
        }
    } finally {
        $process.Dispose()
    }
}
