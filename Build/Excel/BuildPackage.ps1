Import-Module PSPublishModule -Force -ErrorAction Stop

Invoke-DotNetReleaseBuild -ProjectPath "$PSScriptRoot\..\..\OfficeIMO.Excel" -CertificateThumbprint '483292C9E317AA13B07BB7A96AE9D1A5ED9E7703'