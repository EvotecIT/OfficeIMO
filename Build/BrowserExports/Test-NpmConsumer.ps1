param([Parameter(Mandatory)][string] $EvidenceDirectory)
$ErrorActionPreference = 'Stop'
$scriptShell = (Get-Command pwsh -CommandType Application | Select-Object -First 1).Source
$repository = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$npmCommand = (Get-Command npm.cmd,npm -ErrorAction SilentlyContinue | Select-Object -First 1).Source
if (-not $npmCommand) { throw 'npm is required for the isolated consumer check.' }
$output = [IO.Path]::GetFullPath($EvidenceDirectory)
Push-Location (Join-Path $repository 'OfficeIMO.JavaScript')
try {
    & $npmCommand --script-shell $scriptShell run test:pack -- $output
    if ($LASTEXITCODE -ne 0) { throw 'Packed npm consumer qualification failed.' }
} finally { Pop-Location }
