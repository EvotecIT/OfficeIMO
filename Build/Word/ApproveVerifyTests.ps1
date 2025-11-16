param (
    $SolutionRoot = "$PSScriptRoot\..\.."
)

$ErrorActionPreference = 'Stop'
$SolutionPath = [io.path]::Combine($SolutionRoot, 'OfficeImo.sln')
Write-Host "Searching for solution at $SolutionPath" -ForegroundColor Cyan
if ($SolutionRoot -and (Test-Path -Path $SolutionPath)) {
    Write-Host "Solution found at $($SolutionPath). Processing files..." -ForegroundColor Green

    # Limit the search to the Verify snapshots for Word tests so we don't
    # recurse the entire repo (which can be very slow on Windows).
    $verifyRoot = [io.path]::Combine($SolutionRoot, 'OfficeIMO.VerifyTests', 'Word', 'verified')
    if (-not (Test-Path -Path $verifyRoot)) {
        Write-Host "Verify directory not found at $verifyRoot" -ForegroundColor Yellow
        return
    }

    $pattern = '*.received.txt'
    Write-Host "Looking for '$pattern' under $verifyRoot" -ForegroundColor Cyan
    # Use Where-Object instead of -Filter/-Include to avoid provider quirks
    $receivedFiles = Get-ChildItem -Path $verifyRoot -Recurse |
        Where-Object { -not $_.PSIsContainer -and $_.Name -like $pattern }

    if (-not $receivedFiles) {
        Write-Host "No '$pattern' files found under $verifyRoot" -ForegroundColor Yellow
        return
    }

    Write-Host "Found $($receivedFiles.Count) file(s) to approve." -ForegroundColor Green

    foreach ($file in $receivedFiles) {
        $ReceivedTestResult = $file.FullName

        # For multi-targeted tests Verify uses:
        #   Received: *.DotNetX_Y.received.txt
        #   Verified: *.verified.txt
        # so we must drop the framework suffix when promoting.
        if ($ReceivedTestResult -match '\.DotNet\d+_\d+\.received\.txt$') {
            $ApprovedTestResult = $ReceivedTestResult -replace '\.DotNet\d+_\d+\.received\.txt$', '.verified.txt'
        } else {
            $ApprovedTestResult = $ReceivedTestResult -replace '\.received\.txt$', '.verified.txt'
        }

        if ($ApprovedTestResult -eq $ReceivedTestResult) {
            Write-Host "Skipping $ReceivedTestResult (could not construct verified name)" -ForegroundColor DarkYellow
            continue
        }

        Write-Host "Approving $ReceivedTestResult -> $ApprovedTestResult" -ForegroundColor Yellow
        Move-Item -LiteralPath $ReceivedTestResult -Destination $ApprovedTestResult -Force
    }
} else {
    Write-Host -Object "Solution not found at $SolutionPath" -ForegroundColor Red
}
