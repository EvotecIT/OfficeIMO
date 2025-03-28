param (
    $SolutionRoot = "$PSScriptRoot\.."
)

$ErrorActionPreference = 'Stop'
$DotnetVersions = @('6_0', '7_0', '8_0', '9_0')
$SolutionPath = [io.path]::Combine($SolutionRoot, 'OfficeImo.sln')
if ($SolutionRoot -and (Test-Path -Path $SolutionPath)) {
    Write-Host "Solution found at $($SolutionPath). Processing files..." -ForegroundColor Green
    Get-ChildItem -Recurse $SolutionRoot -Filter "*.received.txt" | ForEach-Object {
        Write-Host "Approving $($_.FullName)" -ForegroundColor Yellow
        $ReceivedTestResult = $_.FullName
        foreach ($DotNetVersion in $DotNetVersions) {
            if ($ReceivedTestResult -like "*DotNet$DotNetVersion*") {
                $ApprovedTestResult = $ReceivedTestResult.Replace(".DotNet$DotNetVersion.received.txt", '.verified.txt')
                Move-Item -LiteralPath $ReceivedTestResult -Destination $ApprovedTestResult -Force
            }
        }
    }
} else {
    Write-Host -Object "Solution not found at $SolutionPath" -ForegroundColor Red
}