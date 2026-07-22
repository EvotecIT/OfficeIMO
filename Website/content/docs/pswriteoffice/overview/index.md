---
title: "What PSWriteOffice Covers"
description: "A map of the module's 464 cmdlets, 15 workflow families, and relationship to OfficeIMO."
layout: docs
---

Use PSWriteOffice when a PowerShell job needs to create, convert, inspect, repair, or publish documents without automating desktop Office applications. The module is a thin PowerShell surface over the OfficeIMO libraries, so scripts and .NET applications use the same document engines and file-format behavior.

The current manifest exports 464 cmdlets and 354 aliases across 15 documented families. Those totals are generated from `PSWriteOffice.psd1`, not maintained as marketing copy.

## Choose the surface by outcome

| Need | Start with | Why |
| --- | --- | --- |
| Produce a report or template-driven artifact | Word, Excel, PowerPoint, or PDF DSL | Script blocks keep composition close to the data and make repeated jobs readable. |
| Update an existing file | `Get-*`, `Set-*`, `Update-*`, and `Save-*` commands | The object remains in the OfficeIMO model while the script performs targeted changes. |
| Review or diagnose files | inspection, preflight, comparison, and HTML export commands | Read-only diagnostics can run before a job decides whether to change or reject an artifact. |
| Normalize many formats | Reader commands | One result model exposes documents, chunks, hierarchy, tables, visuals, assets, warnings, and provenance. |
| Convert between formats | focused `ConvertFrom-*` and `ConvertTo-*` commands | Conversion behavior stays in the matching OfficeIMO adapter instead of shelling out to Office. |

## Major families

- **Excel — 155 commands:** authoring, reading, charts, pivots, validation, comments, templates, comparison, repair, accessibility, streaming, and HTML review.
- **Word — 91 commands:** sections, paragraphs, lists, tables, fields, content controls, review, mail merge, protection, merging, and conversion.
- **PDF — 74 commands:** composition, text and image extraction, merge/split, pages, forms, annotations, attachments, signatures, compliance, redaction, optimization, and diagnostics.
- **PowerPoint — 57 commands:** slides, sections, shapes, charts, tables, notes, themes, layouts, transitions, import, inspection, and HTML review.
- **Markdown, Visio, Reader, and open formats:** typed Markdown, VSDX diagrams and stencils, normalized extraction, RTF, CSV, ODT/ODS/ODP, email, AsciiDoc, and LaTeX workflows.

## How the documentation fits together

Conceptual guides answer which workflow to choose and how objects move through a script. The generated [command reference](/api/powershell/) owns parameter sets, accepted values, pipeline behavior, and source links. The [example gallery](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples) provides copyable end-to-end scripts.

When a guide and command page appear to disagree, use the current command reference and report the guide mismatch. The catalog validation prevents command totals from drifting, while examples and help remain the executable source for exact syntax.
