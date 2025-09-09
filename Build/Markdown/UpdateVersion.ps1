Import-Module PSPublishModule -Force -ErrorAction Stop

$Path = "$PSScriptRoot\..\..\OfficeIMO.Markdown"

Get-ProjectVersion -Path "$Path" -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
Set-ProjectVersion -Path "$Path" -NewVersion "0.1" -WhatIf:$true -Verbose -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
