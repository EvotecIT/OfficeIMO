Import-Module PSPublishModule -Force -ErrorAction Stop

$Path = "$PSScriptRoot\..\..\OfficeIMO.Markdown"

Get-ProjectVersion -Path "$Path" -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
Set-ProjectVersion -Path "$Path" -NewVersion "0.2.0" -WhatIf:$false -Verbose -ExcludeFolders @("$Path\Module\Artefacts") | Format-Table
