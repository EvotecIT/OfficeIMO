---
title: "PSWriteOffice Command Families"
description: "The generated map of all exported cmdlets and the conceptual guide for each family. Includes PowerShell examples, validation notes, and cmdlet links."
layout: docs
---

The PSWriteOffice website catalog groups every exported cmdlet into exactly one workflow family. The build reads `PSWriteOffice.psd1`; it fails when a featured command is not exported or a new command has no family.

## Current surface

| Family | Exported commands | Guide |
| --- | ---: | --- |
| Excel | 161 | [Excel automation](/docs/pswriteoffice/excel/) |
| Word | 98 | [Word automation](/docs/pswriteoffice/word/) |
| PDF | 96 | [PDF automation](/docs/pswriteoffice/pdf/) |
| Conversion and capabilities | 7 | [Conversion and capabilities](/docs/pswriteoffice/automation-patterns/) |
| Project | 19 | [Native document workflows](/docs/pswriteoffice/automation-patterns/) |
| Access | 9 | [Native document workflows](/docs/pswriteoffice/automation-patterns/) |
| EPUB | 8 | [Native document workflows](/docs/pswriteoffice/automation-patterns/) |
| PowerPoint | 67 | [PowerPoint automation](/docs/pswriteoffice/powerpoint/) |
| Markdown | 27 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| Visio | 24 | [Visio diagrams](/docs/pswriteoffice/visio/) |
| Reader, extraction, and OCR | 15 | [Reader, extraction, and OCR](/docs/pswriteoffice/reader/) |
| Archive migration and evidence | 5 | [Automation patterns](/docs/pswriteoffice/automation-patterns/) |
| Confluence Cloud | 7 | [Confluence Cloud publishing](/docs/pswriteoffice/confluence/) |
| RTF | 6 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| CSV | 5 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| OpenDocument | 11 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| Email | 9 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| AsciiDoc | 4 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| LaTeX | 4 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| HTML assets | 4 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| Cross-format visuals | 1 | [Automation patterns](/docs/pswriteoffice/automation-patterns/) |
| Shared authoring primitives | 1 | [Automation patterns](/docs/pswriteoffice/automation-patterns/) |

The total is generated from the module manifest. Aliases are intentionally counted separately because they provide a shorter DSL without replacing the canonical command names in help and automation.

## Reference versus guide

Use this conceptual catalog to choose a family. Use the generated [command reference](/api/powershell/) for exact parameter sets, common parameters, examples, source links, and related commands. Searching the reference for a noun such as `OfficePdf`, `OfficeExcelChart`, or `OfficeDocument` is often faster than scanning a flat alphabetical list.

## Evidence contract

The machine-readable catalog lives at `PowerShell/PSWriteOffice/WebsiteArtifacts/documentation/command-catalog.json` in the OfficeIMO repository. A Pester contract verifies:

- command and alias totals match the module manifest;
- family counts add up to the full exported command count;
- every featured command is currently exported;
- docs, examples, and API surfaces are enabled at the module version;
- regenerating the catalog produces no uncommitted difference.
