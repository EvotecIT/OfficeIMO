Import-Module PSPublishModule -Force -ErrorAction Stop

$Path = "$PSScriptRoot\..\..\OfficeIMO.Excel"

Get-ProjectVersion -Path "$Path" -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
Set-ProjectVersion -Path "$Path" -NewVersion "0.1.0" -WhatIf:$false -Verbose -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
