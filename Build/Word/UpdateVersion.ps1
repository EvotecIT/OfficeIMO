Import-Module PSPublishModule -Force -ErrorAction Stop

$Path = "$PSScriptRoot\..\..\OfficeIMO.Word"

Get-ProjectVersion -Path "$Path" -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
Set-ProjectVersion -Path "$Path" -NewVersion "1.0.9" -WhatIf:$false -Verbose -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
