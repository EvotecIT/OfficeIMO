---
title: "Apple iWork"
description: "Read Pages, Numbers, and Keynote packages safely, then opt in to focused Word, Excel, or PowerPoint conversion adapters."
order: 39
meta.seo_title: "Apple Pages, Numbers, and Keynote import | OfficeIMO"
---

## Choose the package boundary

`OfficeIMO.IWork` is the bounded, read-only source package for modern Apple Pages, Numbers, and Keynote files. It does not add Word, Excel, or PowerPoint as dependencies.

```shell
dotnet add package OfficeIMO.IWork
```

Install a destination adapter only when the application needs that conversion:

```shell
dotnet add package OfficeIMO.Word.IWork
dotnet add package OfficeIMO.Excel.IWork
dotnet add package OfficeIMO.PowerPoint.IWork
```

The default `OfficeIMO.Word`, `OfficeIMO.Excel`, and `OfficeIMO.PowerPoint` packages do not depend on iWork. Each adapter brings together the bounded source reader and exactly one Office destination package.

## Read and inspect an iWork source

```csharp
using OfficeIMO.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("report.pages");
IWorkPagesProjection pages = source.ReadPages();

Console.WriteLine(source.Kind);
Console.WriteLine(pages.Paragraphs.Count);
```

The reader accepts ZIP packages, directory bundles, and packages with a nested `Index.zip`. It bounds package, archive, decoded-text, table, cell, image, and formula work and does not execute embedded content or fetch external resources.

## Convert through an opt-in adapter

```csharp
using OfficeIMO.IWork;
using OfficeIMO.Excel.IWork;

using IWorkNumbersLoadResult result = ExcelIWorkConverter.LoadNumbersWithReport(
    "budget.numbers",
    new IWorkReadOptions { ImportMode = IWorkImportMode.Auto });

Console.WriteLine(result.ImportReport.ProjectionKind);
Console.WriteLine(result.HasConversionLoss);
result.Document.Save("budget.xlsx");
```

Use `WordIWorkConverter.LoadPages*` for Pages, `ExcelIWorkConverter.LoadNumbers*` for Numbers, and `PowerPointIWorkConverter.LoadKeynote*` for Keynote. The short overload returns the destination document. The `WithReport` overload also exposes the typed projection, diagnostics, preserved source records, producer build history, and the exact editable or visual-fallback result.

## Preservation boundary

OfficeIMO reconstructs supported content as editable DOCX, XLSX, or PPTX. Unsupported and partially consumed records remain available in the bounded source model and are reported as conversion loss. `VisualOnly` can use a package preview when semantic reconstruction is not wanted, but that preview is not presented as editable content.

There is no Pages, Numbers, or Keynote writer. See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for the tested producer corpus, exact semantic coverage, limits, and known boundaries.
