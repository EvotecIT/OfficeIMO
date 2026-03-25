---
title: Documentation
description: Complete documentation for the OfficeIMO suite of .NET libraries for creating and manipulating Office documents.
order: 1
slug: index
---

# OfficeIMO Documentation

OfficeIMO is an open-source, cross-platform suite of .NET libraries for creating and manipulating Office document formats without requiring Office to be installed. Built on top of the Open XML SDK and related package-specific components, OfficeIMO provides focused APIs for Word, Excel, PowerPoint, Markdown, CSV, and adjacent conversion workflows.

## Packages

| Package | Package Feed | Description |
|---------|--------------|-------------|
| **OfficeIMO.Word** | [OfficeIMO.Word on NuGet](https://www.nuget.org/packages/OfficeIMO.Word) | Create, read, and modify Word (.docx) documents |
| **OfficeIMO.Excel** | [OfficeIMO.Excel on NuGet](https://www.nuget.org/packages/OfficeIMO.Excel) | Create, read, and modify Excel (.xlsx) workbooks |
| **OfficeIMO.PowerPoint** | [OfficeIMO.PowerPoint on NuGet](https://www.nuget.org/packages/OfficeIMO.PowerPoint) | Generate PowerPoint (.pptx) presentations with slides, charts, and layouts |
| **OfficeIMO.Markdown** | [OfficeIMO.Markdown on NuGet](https://www.nuget.org/packages/OfficeIMO.Markdown) | Fluent Markdown builder, reader/AST, and HTML renderer |
| **OfficeIMO.CSV** | [OfficeIMO.CSV on NuGet](https://www.nuget.org/packages/OfficeIMO.CSV) | Strongly-typed CSV document model |
| **OfficeIMO.Visio** | [OfficeIMO.Visio on NuGet](https://www.nuget.org/packages/OfficeIMO.Visio) | Create and modify Visio (.vsdx) diagrams with shapes and connectors |
| **OfficeIMO.Reader** | [OfficeIMO.Reader on NuGet](https://www.nuget.org/packages/OfficeIMO.Reader) | Extract and chunk document content for indexing, search, and AI ingestion |
| **OfficeIMO.Word.Html** | [OfficeIMO.Word.Html on NuGet](https://www.nuget.org/packages/OfficeIMO.Word.Html) | Convert Word documents to/from HTML |
| **OfficeIMO.Word.Markdown** | [OfficeIMO.Word.Markdown on NuGet](https://www.nuget.org/packages/OfficeIMO.Word.Markdown) | Convert Word documents to/from Markdown |
| **PSWriteOffice** | [PSWriteOffice on PowerShell Gallery](https://www.powershellgallery.com/packages/PSWriteOffice) | PowerShell module wrapping OfficeIMO for Word, Excel, PowerPoint, and Markdown |

## Quick Links

- [Installation](/docs/getting-started/installation) -- Install any OfficeIMO package via NuGet or PowerShell Gallery.
- [Quick Start](/docs/getting-started/quickstart) -- Create your first document in under five minutes.
- [Platform Support](/docs/getting-started/platform-support) -- Supported frameworks, operating systems, and AOT notes.

### By Document Type

- [Word Documents](/docs/word/) -- Paragraphs, tables, images, headers/footers, charts, bookmarks, and more.
- [Excel Workbooks](/docs/excel/) -- Worksheets, cell formatting, tables, conditional formatting, and charts.
- [PowerPoint Presentations](/docs/powerpoint/) -- Slides, text boxes, shapes, and images.
- [Markdown](/docs/markdown/) -- Fluent builder, typed AST reader, and HTML rendering pipeline.
- [CSV](/docs/csv/) -- Typed read/write workflows with validation and streaming.
- [Visio Diagrams](/docs/visio/) -- Pages, shapes, connectors, and diagram generation patterns.
- [Reader & Extraction](/docs/reader/) -- Unified extraction, chunking, and ingestion workflows.

### Converters

- [Word to HTML](/docs/converters/word-html) -- Bidirectional conversion between Word and HTML using AngleSharp.
- [Word to Markdown](/docs/converters/word-markdown) -- Bidirectional conversion between Word and Markdown.

### PowerShell

- [PSWriteOffice Docs](/docs/pswriteoffice/) -- PowerShell cmdlets for Office document and Markdown automation.
- [PowerPoint Cmdlets](/docs/pswriteoffice/powerpoint/) -- Build slide decks with cmdlets and DSL aliases.
- [Markdown Cmdlets](/docs/pswriteoffice/markdown/) -- Generate Markdown reports, READMEs, and automation output.

### Advanced

- [AOT and Trimming](/docs/advanced/aot-trimming) -- Guidance for ahead-of-time compilation and IL trimming.

## License

OfficeIMO is licensed under the [MIT License](https://github.com/EvotecIT/OfficeIMO/blob/master/LICENSE). Copyright (c) Przemyslaw Klys @ Evotec.

## Source Code

The full source code is available on GitHub: [https://github.com/EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
