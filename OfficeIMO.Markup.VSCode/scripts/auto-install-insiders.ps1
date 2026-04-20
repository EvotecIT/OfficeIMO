param(
    [string]$TaskName = 'OfficeIMO.Markup.Insiders.AutoInstall',
    [string]$DailyAt = '09:00',
    [switch]$DisableDaily,
    [switch]$DisableLogon,
    [switch]$Remove
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
$installScript = Join-Path $repoRoot 'scripts/install-insiders.ps1'
if (-not (Test-Path -LiteralPath $installScript)) {
    throw "install-insiders.ps1 not found at $installScript"
}

$userId = if ($env:USERDOMAIN) { "$($env:USERDOMAIN)\$($env:USERNAME)" } else { $env:USERNAME }

if ($Remove) {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed scheduled task '$TaskName'." -ForegroundColor Green
    } else {
        Write-Host "Scheduled task '$TaskName' not found." -ForegroundColor Yellow
    }
    return
}

$includeDaily = -not $DisableDaily
$includeLogon = -not $DisableLogon
if (-not $includeDaily -and -not $includeLogon) {
    throw 'Both triggers are disabled. Use at least one trigger.'
}

$triggers = @()
if ($includeDaily) {
    $triggers += New-ScheduledTaskTrigger -Daily -At $DailyAt
}
if ($includeLogon) {
    $triggers += New-ScheduledTaskTrigger -AtLogOn -User $userId
}

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$installScript`" -Force" `
    -WorkingDirectory $repoRoot

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable

$principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited
$task = New-ScheduledTask -Action $action -Trigger $triggers -Settings $settings -Principal $principal `
    -Description 'Rebuilds and installs OfficeIMO Markup into VS Code Insiders.'

Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null

$triggerSummary = @()
if ($includeLogon) { $triggerSummary += 'logon' }
if ($includeDaily) { $triggerSummary += "daily $DailyAt" }
Write-Host "Scheduled task '$TaskName' registered ($($triggerSummary -join ', '))." -ForegroundColor Green
Write-Host "To remove: scripts/auto-install-insiders.ps1 -Remove" -ForegroundColor Cyan
