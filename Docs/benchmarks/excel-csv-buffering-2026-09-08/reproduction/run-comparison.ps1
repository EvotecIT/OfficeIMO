param(
    [string] $Label,
    [string] $ScenarioList,
    [int] $Iterations = 24,
    [int] $Exports = 16,
    [int] $Repeats = 1,
    [switch] $SkipFirstDomain
)
$ErrorActionPreference = 'Stop'
$powerShellPath = (Get-Command pwsh).Source
for ($repeat = 1; $repeat -le $Repeats; $repeat++) {
    foreach ($mask in @('65535','4294901760')) {
        if ($SkipFirstDomain -and $repeat -eq 1 -and $mask -eq '65535') { continue }
        $runLabel = $Label + '-r' + $repeat
        & $powerShellPath -NoProfile -File (Join-Path $PSScriptRoot 'rotated.ps1') -Mask $mask -Iterations $Iterations -Warmup 12 -Exports $Exports -Label $runLabel -ScenarioList $ScenarioList *> (Join-Path $PSScriptRoot ($runLabel+'-'+$mask+'.log'))
        if ($LASTEXITCODE -ne 0) { throw ('Comparison failed: '+$runLabel+' domain '+$mask) }
        Write-Output ('Completed '+$runLabel+' domain '+$mask)
    }
}
