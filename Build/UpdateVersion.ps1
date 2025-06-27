Import-Module PSPublishModule -Force -ErrorAction Stop

$Path = "$PSScriptRoot\.."

Get-ProjectVersion -Path "$Path" -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
Set-ProjectVersion -Path "$Path" -NewVersion "1.0.3" -WhatIf:$false -Verbose -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
