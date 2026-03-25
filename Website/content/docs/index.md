---
title: Documentation
description: Complete documentation for the OfficeIMO suite of .NET libraries for creating and manipulating Office documents.
order: 1
slug: index
---

# OfficeIMO Documentation

OfficeIMO is an open-source, cross-platform suite of .NET libraries for creating and manipulating Microsoft Office documents without requiring Office to be installed. Built on top of the Open XML SDK, OfficeIMO provides a developer-friendly API that dramatically reduces the complexity of working with Word, Excel, PowerPoint, Markdown, and CSV files.

## Packages

| Package | NuGet | Description |
|---------|-------|-------------|
| **OfficeIMO.Word** | [![NuGet](https://img.shields.io/nuget/v/OfficeIMO.Word)](https://www.nuget.org/packages/OfficeIMO.Word) | Create, read, and modify Word (.docx) documents |
| **OfficeIMO.Excel** | [![NuGet](https://img.shields.io/nuget/v/OfficeIMO.Excel)](https://www.nuget.org/packages/OfficeIMO.Excel) | Create, read, and modify Excel (.xlsx) workbooks |
| **OfficeIMO.Markdown** | [![NuGet](https://img.shields.io/nuget/v/OfficeIMO.Markdown)](https://www.nuget.org/packages/OfficeIMO.Markdown) | Fluent Markdown builder, reader/AST, and HTML renderer |
| **OfficeIMO.CSV** | [![NuGet](https://img.shields.io/nuget/v/OfficeIMO.CSV)](https://www.nuget.org/packages/OfficeIMO.CSV) | Strongly-typed CSV document model |
| **OfficeIMO.Word.Html** | [![NuGet](https://img.shields.io/nuget/v/OfficeIMO.Word.Html)](https://www.nuget.org/packages/OfficeIMO.Word.Html) | Convert Word documents to/from HTML |
| **OfficeIMO.Word.Markdown** | [![NuGet](https://img.shields.io/nuget/v/OfficeIMO.Word.Markdown)](https://www.nuget.org/packages/OfficeIMO.Word.Markdown) | Convert Word documents to/from Markdown |
| **PSWriteOffice** | [![PowerShell Gallery](https://img.shields.io/powershellgallery/v/PSWriteOffice)](https://www.powershellgallery.com/packages/PSWriteOffice) | PowerShell module wrapping OfficeIMO for Word and Excel |

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

### Converters

- [Word to HTML](/docs/converters/word-html) -- Bidirectional conversion between Word and HTML using AngleSharp.
- [Word to Markdown](/docs/converters/word-markdown) -- Bidirectional conversion between Word and Markdown.

### PowerShell

- [PSWriteOffice](/docs/pswriteoffice/) -- PowerShell cmdlets for Word and Excel automation.

### Advanced

- [AOT and Trimming](/docs/advanced/aot-trimming) -- Guidance for ahead-of-time compilation and IL trimming.

## License

OfficeIMO is licensed under the [MIT License](https://github.com/EvotecIT/OfficeIMO/blob/master/LICENSE). Copyright (c) Przemyslaw Klys @ Evotec.

## Source Code

The full source code is available on GitHub: [https://github.com/EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
