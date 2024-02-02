param (
    $SolutionRoot = "$PSScriptRoot/.."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-Not $PSScriptRoot) {
    Write-Error "\`$PSScriptRoot variable isn't set."
}
else {
    Get-ChildItem -Recurse $SolutionRoot -Filter "*.received.txt" | ForEach-Object {
        $receivedTestResult = $_.FullName
        $approvedTestResult = $receivedTestResult.Replace('.received.txt', '.verified.txt')
        Move-Item -Force -LiteralPath $receivedTestResult $approvedTestResult
    }
}
