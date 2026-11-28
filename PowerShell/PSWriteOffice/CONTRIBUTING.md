# Contributing to PSWriteOffice

PSWriteOffice lives in the OfficeIMO repository and references the engines in the same checkout. Cmdlets stay thin: reusable behaviour belongs in the OfficeIMO packages, and the module supports Windows PowerShell 5.1 and PowerShell 7.

## Build and test

Run these from `PowerShell/PSWriteOffice`:

```powershell
dotnet build .\Sources\PSWriteOffice.slnx -c Debug
pwsh -NoLogo -NoProfile -File .\PSWriteOffice.Tests.ps1
pwsh -NoLogo -NoProfile -File .\Build\Build-Module.ps1 -RunMode Build
```

[PSWriteOffice.psm1](PSWriteOffice.psm1) loads the local debug build from `Sources\PSWriteOffice\bin\Debug\` during development.

Cmdlet help in `Docs/` is generated from the cmdlets' XML comments. Regenerate it from the repository root with `./Build/Build-Module.ps1 -RunMode Documentation -IncludePackages:$false` instead of editing the Markdown.
