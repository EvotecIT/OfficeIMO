Import-Module PSPublishModule -Force -ErrorAction Stop

$Path = "$PSScriptRoot\.."

Get-ProjectVersion -Path "$Path" -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
Set-ProjectVersion -Path "$Path" -NewVersion "0.0.24" -WhatIf:$false -Verbose -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
