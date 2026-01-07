Import-Module PSPublishModule -Force -ErrorAction Stop

$certificateThumbprint = '483292C9E317AA13B07BB7A96AE9D1A5ED9E7703'
Invoke-DotNetReleaseBuild -ProjectPath @(
    "$PSScriptRoot\..\OfficeIMO.CSV"
    "$PSScriptRoot\..\OfficeIMO.Excel"
    "$PSScriptRoot\..\OfficeIMO.Markdown"
    "$PSScriptRoot\..\OfficeIMO.Word"
) -CertificateThumbprint $certificateThumbprint
