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
| Update an existing file | `Open-*`, `Get-*`, `Set-*`, `Update-*`, and `Save-*` commands | The object remains in the OfficeIMO model while the script performs targeted changes. |
| Review or diagnose files | inspection, preflight, comparison, and HTML export commands | Read-only diagnostics can run before a job decides whether to change or reject an artifact. |
| Normalize many formats | Reader commands | One result model exposes documents, chunks, hierarchy, tables, visuals, assets, warnings, and provenance. |
| Convert between formats or place reusable visuals | `Convert-OfficeDocument` and visual commands | Conversion behavior stays in the matching OfficeIMO adapter. PDF input can also be reconstructed as editable Word, Excel, or PowerPoint content. |
| Edit configuration or move data between formats | `Open-OfficeYaml`, `Open-OfficeToml`, `Import-OfficeToml`, and `Convert-OfficeDocument` | YAML and TOML edits keep comments and layout, and data conversions report anything the target format cannot hold. See [Configuration files and data conversion](/docs/pswriteoffice/config-data-formats/). |
| Share a file outside your team | `Get-OfficeHiddenData` and `Remove-OfficeHiddenData` | Comments, authors, speaker notes, image metadata and AI labels are listed and removed into a cleaned copy, with an outcome for every item. See [Clean files before sharing](/docs/pswriteoffice/before-sharing/). |
| Check protected-content support before automation | `Get-OfficeDocumentCapability -Protection` | Filterable, typed capability rows show supported operations, limits, and incomplete format coverage before a workflow handles a protected file. |

## Major families

- **Excel — 102 commands:** authoring, reading, charts, pivots, validation, comments, templates, comparison, repair, accessibility, streaming, visual placement, and image export.
- **Word — 80 commands:** sections, paragraphs, lists, tables, fields, content controls, review, mail merge, protection, merging, visual placement, and conversion.
- **PDF — 77 commands:** composition, structured reading, explicit Word/Excel/PowerPoint/Markdown/RTF export, editable Office reconstruction, text and image extraction, searchable OCR output, merge/split, pages, forms, annotations, attachments, signatures, compliance, redaction, optimization, visual placement, and diagnostics.
- **Printing — 5 commands:** print any document OfficeIMO can turn into PDF on a named queue or IPP printer, read printer capabilities, and follow or cancel jobs.
- **PowerPoint — 54 commands:** slides, sections, shapes, charts, tables, notes, themes, layouts, transitions, visual placement, import, inspection, and HTML review.
- **Cross-format operations — 14 commands:** inspect, compare, validate, repair, protect, combine, split, convert, and render documents to images with limits and reports. See [Choose a conversion route](/docs/pswriteoffice/conversion-routes/).
- **Before sharing — 6 commands:** list and remove comments, authors, hidden content, image metadata and AI labels, inspect Content Credentials, and check Word, Excel and PowerPoint accessibility. See [Clean files before sharing](/docs/pswriteoffice/before-sharing/).
- **YAML — 13 commands:** edit files without losing comments or layout, merge overlays, validate, and convert to and from objects. See [Configuration files and data conversion](/docs/pswriteoffice/config-data-formats/).
- **TOML — 13 commands:** edit files without losing comments or layout, and read and write settings and lock files as PowerShell objects. See [Configuration files and data conversion](/docs/pswriteoffice/config-data-formats/).
- **CSV — 4 commands:** import and export delimited data with explicit quoting, encoding, and formula-injection rules.
- **Project — 20 commands:** create and schedule tasks, resources and assignments; exchange mapped tables; save native MPP or Project XML and export reports. See [Project, Access, and EPUB](/docs/pswriteoffice/project-access-epub/).
- **Access — 9 commands:** inspect native MDB/ACCDB schemas and rows, and create inert table databases from PowerShell objects.
- **EPUB — 8 commands:** build publications from chapters, resources and metadata, then inspect and save the package.
- **OneNote and iWork — 5 commands:** read offline OneNote sections and notebooks, and Apple Pages, Numbers and Keynote packages.
- **Google Docs and Sheets — 5 commands:** sign in with a service account and export Word documents to Google Docs and workbooks to Google Sheets.
- **Confluence Cloud — 7 commands:** plan, create, update, and remove pages; preserve managed sections; and list, upload, or download attachments.
- **Markdown, Visio, Reader, OCR, HTML, and open formats:** typed Markdown, VSDX diagrams and stencils, normalized extraction, local image OCR, HTML conversion settings, RTF, ODT/ODS/ODP, email, AsciiDoc, and LaTeX workflows.

## How the documentation fits together

Conceptual guides answer which workflow to choose and how objects move through a script. The generated [command reference](/api/powershell/) owns parameter sets, accepted values, pipeline behavior, and source links. The [example gallery](https://github.com/EvotecIT/OfficeIMO/tree/8b13cc8cfa5dab1b0f401f6dd5962c62ba8ce92b/PowerShell/PSWriteOffice/Examples) provides copyable end-to-end scripts.

When a guide and command page appear to disagree, use the current command reference and report the guide mismatch. The catalog validation prevents command totals from drifting, while examples and help remain the executable source for exact syntax.
