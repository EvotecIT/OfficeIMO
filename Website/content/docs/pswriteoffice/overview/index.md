---
title: "What PSWriteOffice Covers"
description: "A practical map of PSWriteOffice's manifest-derived commands, workflow families, OfficeIMO relationship, and supported automation outcomes."
layout: docs
---

Use PSWriteOffice when a PowerShell job needs to create, convert, inspect, repair, or publish documents without automating desktop Office applications. The module is a thin PowerShell surface over the OfficeIMO libraries, so scripts and .NET applications use the same document engines and file-format behavior.

The module manifest is the source of truth for exported cmdlets and aliases. The generated command catalog derives the current totals and family breakdown from `PSWriteOffice.psd1` so the documentation stays aligned with the packaged module.

## Choose the surface by outcome

| Need | Start with | Why |
| --- | --- | --- |
| Produce a report or template-driven artifact | Word, Excel, PowerPoint, or PDF DSL | Script blocks keep composition close to the data and make repeated jobs readable. |
| Update an existing file | `Get-*`, `Set-*`, `Update-*`, and `Save-*` commands | The object remains in the OfficeIMO model while the script performs targeted changes. |
| Review or diagnose files | inspection, preflight, comparison, and HTML export commands | Read-only diagnostics can run before a job decides whether to change or reject an artifact. |
| Normalize many formats | Reader commands | One result model exposes documents, chunks, hierarchy, tables, visuals, assets, warnings, and provenance. |
| Convert between formats or place reusable visuals | `Convert-OfficeDocument` and visual commands | Conversion behavior stays in the matching OfficeIMO adapter. PDF input can also be reconstructed as editable Word, Excel, or PowerPoint content. |
| Check protected-content support before automation | `Get-OfficeProtectionCapability` | Filterable, typed capability rows show supported operations, limits, and incomplete format coverage before a workflow handles a protected file. |

## Major families

- **Excel — 158 commands:** authoring, reading, charts, pivots, validation, comments, templates, comparison, repair, accessibility, streaming, visual placement, and image export.
- **Word — 98 commands:** sections, paragraphs, lists, tables, fields, content controls, review, mail merge, protection, merging, visual placement, and conversion.
- **PDF — 96 commands:** composition, structured reading, explicit Word/Excel/PowerPoint/Markdown/RTF export, editable Office reconstruction, text and image extraction, searchable OCR output, merge/split, pages, forms, annotations, attachments, signatures, compliance, redaction, optimization, visual placement, and diagnostics.
- **PowerPoint — 65 commands:** slides, sections, shapes, charts, tables, notes, themes, layouts, transitions, visual placement, import, inspection, and HTML review.
- **Conversion and capabilities — 7 commands:** inspect executable routes and protected-content support, then run single-file, selected-file or resumable directory conversions with limits and reports. See [Choose a conversion route](/docs/pswriteoffice/conversion-routes/).
- **Project — 19 commands:** create and schedule tasks, resources and assignments; exchange mapped tables; save native MPP or Project XML and export reports.
- **Access — 9 commands:** inspect native MDB/ACCDB schemas and rows, and create inert table databases from PowerShell objects.
- **EPUB — 8 commands:** build publications from chapters, resources and metadata, then inspect and save the package.
- **Archive migration and evidence — 5 commands:** inspect package security and provenance, identify iWork files, and convert iWork and offline OneNote content with explicit fidelity evidence.
- **Confluence Cloud — 7 commands:** plan, create, update, and remove pages; preserve managed sections; and list, upload, or download attachments.
- **Markdown, Visio, Reader, OCR, visuals, and open formats:** typed Markdown, VSDX diagrams and stencils, cross-format visual placement, normalized extraction, local image OCR, RTF, CSV, ODT/ODS/ODP, email, AsciiDoc, and LaTeX workflows.

## How the documentation fits together

Conceptual guides answer which workflow to choose and how objects move through a script. The generated [command reference](/api/powershell/) owns parameter sets, accepted values, pipeline behavior, and source links. The [example gallery](https://github.com/EvotecIT/OfficeIMO/tree/5fc3180191547ff650d54342a1de858f93e245b6/PowerShell/PSWriteOffice/Examples) provides copyable end-to-end scripts.

When a guide and command page appear to disagree, use the current command reference and report the guide mismatch. The catalog validation prevents command totals from drifting, while examples and help remain the executable source for exact syntax.
