---
title: "Troubleshooting and Diagnostics"
description: "Diagnose module loading, file-format routing, rewrite safety, conversion fidelity, and platform-specific dependencies."
layout: docs
---

Start troubleshooting by identifying the layer that failed: module loading, command binding, source-file detection, native document parsing, mutation, conversion, or final save.

## Confirm the loaded module

```powershell
Get-Module PSWriteOffice | Format-List Name,Version,Path
Get-Command -Module PSWriteOffice -Name '*OfficePdf*' | Select-Object Name,CommandType
```

Multiple installed versions can make an interactive session differ from CI. Import the intended version in a fresh process and record the module path.

## Keep the original exception and diagnostics

Use `-ErrorAction Stop` around the bounded operation and preserve the full exception. Prefer format-specific diagnostic and preflight commands over retrying the same rewrite blindly.

## Reopen generated output

After a save or conversion, load the output through the matching reader, extract a stable marker, and inspect warnings. For visual formats, a rendered review artifact can supplement structural checks but should not replace them.

## Separate unsupported behavior from corruption

A converter can reject or warn about a feature it does not map; that is different from producing a malformed destination. Retain conversion diagnostics and choose an explicit fallback policy instead of silently dropping content.

## Check optional boundaries

OCR, web transport, authentication, native rendering, and external-process adapters can have platform or deployment requirements beyond the core managed module. Confirm those dependencies in the relevant OfficeIMO package guide and test the exact deployment path.

## Prepare a useful issue

Include the PSWriteOffice version, PowerShell version/edition, operating system, command and parameters, a minimal input when shareable, the complete exception, and relevant diagnostic output. Remove secrets and personal document content before attaching files.
