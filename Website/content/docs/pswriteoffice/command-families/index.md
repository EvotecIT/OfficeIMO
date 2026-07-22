---
title: "PSWriteOffice Command Families"
description: "The generated map of all exported cmdlets and the conceptual guide for each family."
layout: docs
---

The PSWriteOffice website catalog groups every exported cmdlet into exactly one workflow family. The build reads `PSWriteOffice.psd1`; it fails when a featured command is not exported or a new command has no family.

## Current surface

| Family | Exported commands | Guide |
| --- | ---: | --- |
| Excel | 155 | [Excel automation](/docs/pswriteoffice/excel/) |
| Word | 91 | [Word automation](/docs/pswriteoffice/word/) |
| PDF | 74 | [PDF automation](/docs/pswriteoffice/pdf/) |
| PowerPoint | 57 | [PowerPoint automation](/docs/pswriteoffice/powerpoint/) |
| Markdown | 25 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| Visio | 20 | [Visio diagrams](/docs/pswriteoffice/visio/) |
| Reader and extraction | 13 | [Reader and extraction](/docs/pswriteoffice/reader/) |
| RTF | 5 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| CSV | 5 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| OpenDocument | 5 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| Email | 4 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| AsciiDoc | 4 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| LaTeX | 4 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| HTML assets | 1 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| Shared authoring primitives | 1 | [Automation patterns](/docs/pswriteoffice/automation-patterns/) |

The total is 464 cmdlets. Aliases are intentionally counted separately because they provide a shorter DSL without replacing the canonical command names in help and automation.

## Reference versus guide

Use this conceptual catalog to choose a family. Use the generated [command reference](/api/powershell/) for exact parameter sets, common parameters, examples, source links, and related commands. Searching the reference for a noun such as `OfficePdf`, `OfficeExcelChart`, or `OfficeDocument` is often faster than scanning a flat alphabetical list.

## Evidence contract

The machine-readable catalog lives at `WebsiteArtifacts/documentation/command-catalog.json` in the PSWriteOffice repository. A Pester contract verifies:

- command and alias totals match the module manifest;
- family counts add up to the full exported command count;
- every featured command is currently exported;
- docs, examples, and API surfaces are enabled at the module version;
- regenerating the catalog produces no uncommitted difference.
