---
title: "Install and Verify PSWriteOffice"
description: "Install the module, confirm the command surface, and choose a safe output location."
layout: docs
---

Install PSWriteOffice from the PowerShell Gallery before trying the curated document automation examples.

```powershell
Install-Module -Name PSWriteOffice -Scope CurrentUser
Import-Module PSWriteOffice
```

Confirm what was loaded:

```powershell
$module = Get-Module PSWriteOffice
$module.Version

Get-Command -Module PSWriteOffice |
    Group-Object Noun |
    Sort-Object Count -Descending |
    Select-Object -First 12 Count, Name
```

PSWriteOffice does not require Microsoft Office for its managed document workflows. Individual optional integrations can still have their own platform, native library, authentication, or external-process requirements; check the relevant guide before deploying a broad conversion or OCR workflow.

## Use predictable output paths

Examples should write to a script-local or explicitly configured artifact directory:

```powershell
$outputRoot = Join-Path $PSScriptRoot 'Output'
New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null

$wordPath = Join-Path $outputRoot 'Report.docx'
$pdfPath = Join-Path $outputRoot 'Report.pdf'
```

This makes cleanup, CI artifact upload, and permissions easier to reason about than relying on the current working directory.

## Next steps

- [Choose a workflow](/docs/pswriteoffice/choosing-a-workflow/)
- [Browse command families](/docs/pswriteoffice/command-families/)
- [Open the example gallery](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples)
