Import-Module PSPublishModule -Force -ErrorAction Stop

$NugetAPI = Get-Content -Raw -LiteralPath "C:\Support\Important\NugetOrgEvotec.txt"
Publish-NugetPackage -Path @(
    "$PSScriptRoot\..\OfficeIMO.CSV\bin\Release"
    "$PSScriptRoot\..\OfficeIMO.Excel\bin\Release"
    "$PSScriptRoot\..\OfficeIMO.Markdown\bin\Release"
    "$PSScriptRoot\..\OfficeIMO.Word\bin\Release"
) -ApiKey $NugetAPI -SkipDuplicate
