---
title: "Migrate Legacy Modules to PSWriteOffice"
description: "Move legacy PowerShell document scripts to PSWriteOffice with practical command mappings, replacement examples, and compatibility checks."
layout: docs
---

PSWriteOffice is the maintained replacement for the archived PSWriteWord, PSWriteExcel, and PSWritePDF modules. It exposes the current OfficeIMO engines through one PowerShell module, so a single automation can work with Word, Excel, PowerPoint, PDF, email, Visio, and open or text formats.

This is a migration, not a module rename. The current commands use consistent `OfficeWord`, `OfficeExcel`, and `OfficePdf` nouns, and the document DSL owns creation and saving. Existing scripts should be rewritten and tested against representative files instead of changing only the module name.

## Choose your migration

| Legacy module | Maintained path | Start with |
| --- | --- | --- |
| [PSWriteWord](https://github.com/EvotecIT/PSWriteWord) | PSWriteOffice Word commands | [PSWriteWord migration guide](/docs/pswriteoffice/migrate-from-pswriteword/) |
| [PSWriteExcel](https://github.com/EvotecIT/PSWriteExcel) | PSWriteOffice Excel commands | [PSWriteExcel migration guide](/docs/pswriteoffice/migrate-from-pswriteexcel/) |
| [PSWritePDF](https://github.com/EvotecIT/PSWritePDF) | PSWriteOffice PDF commands | [PSWritePDF migration guide](/docs/pswriteoffice/migrate-from-pswritepdf/) |

Install the maintained module from PowerShell Gallery:

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
Import-Module PSWriteOffice
```

## A safe migration sequence

1. Inventory the legacy commands, templates, fonts, sample inputs, and expected output used by the script.
2. Replace one document workflow at a time with the corresponding PSWriteOffice DSL or read/update commands.
3. Write to a new output path while both implementations are available for comparison.
4. Reopen the generated artifact and validate content, formulas, layout, forms, signatures, accessibility, or compliance as appropriate.
5. Remove the legacy module only after scheduled jobs and representative production inputs pass.

Do not load a legacy module and PSWriteOffice into the same migration script unless the command names are fully qualified. Keeping the old and new runs separate makes comparison and rollback clearer.

## What improves after migration

- one maintained PowerShell surface over the OfficeIMO libraries;
- consistent command families and generated command reference;
- document creation plus inspection, update, comparison, preflight, and repair workflows;
- cross-platform .NET behavior without Microsoft Office automation;
- current examples and release validation in the PSWriteOffice repository.

Use the [workflow chooser](/docs/pswriteoffice/choosing-a-workflow/) after migration if the script mixes generated documents, existing-file updates, format conversion, and content extraction.
