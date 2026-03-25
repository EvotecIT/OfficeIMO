---
title: "API Reference"
description: "Browse the generated API reference for OfficeIMO .NET libraries and the PSWriteOffice PowerShell module."
layout: page
---

## Reference Areas

The OfficeIMO website ships two API surfaces:

- .NET library reference generated from the compiled assemblies.
- PowerShell cmdlet reference for **PSWriteOffice**.

## .NET Libraries

| Area | Focus | Link |
|------|-------|------|
| OfficeIMO.Word | Word document creation, editing, formatting, bookmarks, tables, images, charts, and more. | [Open OfficeIMO.Word API](/api/word/) |
| OfficeIMO.Excel | Workbook generation, worksheets, tables, charts, validation, and extraction helpers. | [Open OfficeIMO.Excel API](/api/excel/) |
| OfficeIMO.PowerPoint | Slides, layouts, themes, transitions, tables, charts, and shape composition. | [Open OfficeIMO.PowerPoint API](/api/powerpoint/) |
| OfficeIMO.Markdown | Markdown builder, parser, AST, HTML rendering, and transforms. | [Open OfficeIMO.Markdown API](/api/markdown/) |
| OfficeIMO.CSV | CSV schema definition, typed mapping, validation, and streaming workflows. | [Open OfficeIMO.CSV API](/api/csv/) |
| OfficeIMO.Visio | Diagram pages, shapes, connectors, and fluent Visio generation helpers. | [Open OfficeIMO.Visio API](/api/visio/) |
| OfficeIMO.Reader | Unified extraction, chunking, and ingestion-oriented document processing. | [Open OfficeIMO.Reader API](/api/reader/) |

## PowerShell

| Area | Focus | Link |
|------|-------|------|
| PSWriteOffice | PowerShell cmdlets for Word, Excel, PowerPoint, Markdown, and CSV workflows on top of OfficeIMO. | [Open PSWriteOffice Cmdlets](/api/powershell/) |

## How To Use The API Docs

1. Start on the library landing page that matches the package you use.
2. Filter the sidebar by type name, namespace, or kind.
3. Jump into a type page for signatures, summaries, parameters, and source links.
4. Cross-reference the conceptual guides under [Getting Started](/docs/getting-started/) and the package docs under [Documentation](/docs/).

## How These Pages Are Generated

- .NET library reference is generated from the current `OfficeIMO` build outputs during website CI.
- PSWriteOffice reference is generated from synced PowerShell help XML and example scripts, with a checked-in fallback snapshot for local and clean-checkout builds.

## Need A Practical Entry Point?

- [Getting Started](/docs/getting-started/) for installation and first steps.
- [Documentation](/docs/) for task-based guides.
- [Downloads](/downloads/) for NuGet and PowerShell Gallery entry points.
