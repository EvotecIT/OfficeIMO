---
title: "PSWriteOffice Command Families"
description: "The generated map of all exported cmdlets and the conceptual guide for each family. Includes PowerShell examples, validation notes, and cmdlet links."
layout: docs
---

The PSWriteOffice website catalog groups every exported cmdlet into exactly one workflow family. The build reads `PSWriteOffice.psd1`; it fails when a featured command is not exported, a new command has no family, or a command uses a verb that has no group.

## Current surface

| Family | Exported commands | Guide |
| --- | ---: | --- |
| Excel | 102 | [Excel automation](/docs/pswriteoffice/excel/) |
| Word | 80 | [Word automation](/docs/pswriteoffice/word/) |
| PDF | 77 | [PDF automation](/docs/pswriteoffice/pdf/) |
| PowerPoint | 54 | [PowerPoint automation](/docs/pswriteoffice/powerpoint/) |
| Markdown | 31 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| Visio | 20 | [Visio diagrams](/docs/pswriteoffice/visio/) |
| Project | 20 | [Project, Access, and EPUB](/docs/pswriteoffice/project-access-epub/) |
| Cross-format operations | 14 | [Automation patterns](/docs/pswriteoffice/automation-patterns/) |
| YAML | 13 | [Configuration files and data conversion](/docs/pswriteoffice/config-data-formats/) |
| TOML | 13 | [Configuration files and data conversion](/docs/pswriteoffice/config-data-formats/) |
| OpenDocument | 11 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| Access | 9 | [Project, Access, and EPUB](/docs/pswriteoffice/project-access-epub/) |
| RTF | 9 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| Email | 9 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| EPUB | 8 | [Project, Access, and EPUB](/docs/pswriteoffice/project-access-epub/) |
| Reader, extraction, and OCR | 7 | [Reader, extraction, and OCR](/docs/pswriteoffice/reader/) |
| Confluence Cloud | 7 | [Confluence Cloud publishing](/docs/pswriteoffice/confluence/) |
| Before sharing | 6 | [Clean files before sharing](/docs/pswriteoffice/before-sharing/) |
| HTML | 6 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| AsciiDoc and LaTeX | 6 | [Open and text formats](/docs/pswriteoffice/open-text-formats/) |
| Printing | 5 | [PDF operations](/docs/pdf/operations/) |
| OneNote and iWork | 5 | [OneNote](/docs/onenote/) and [iWork](/docs/iwork/) |
| Google Docs and Sheets | 5 | [Export to Google Workspace](/docs/google-workspace/) |
| CSV | 4 | [Configuration files and data conversion](/docs/pswriteoffice/config-data-formats/) |

Inside each family, commands are grouped by what they do: create, open and save; add content; read and inspect; edit and organize; convert, export and publish; validate, protect and compare; and settings for other commands. The verb decides the group, and `New-*Options` commands are settings. The [PowerShell page](/powershell/) lists every command this way, with its aliases and help synopsis.

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
