# OfficeIMO.IWork

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.IWork)](https://www.nuget.org/packages/OfficeIMO.IWork)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.IWork?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.IWork)

Read, create and preserve Apple Pages, Numbers and Keynote documents in .NET, without iWork. `OfficeIMO.IWork` reads modern iWork packages and selected iWork '09 files, creates new native documents, and saves captured sources with explicit preservation boundaries. It never runs embedded content.

## Install

```sh
dotnet add package OfficeIMO.IWork
```

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/IWork/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO.IWork;

var presentation = IWorkKeynoteDocument.Create(); // 960 x 540 points
var slide = presentation.AddSlide("E6F2FF");
slide.AddText("Quarterly review", 60, 100, 840, 120,
    fontName: "Arial", fontSizePoints: 40, color: "003366");
slide.AddRichText(new[] {
    new IWorkTextRunData("Revenue ", bold: true),
    new IWorkTextRunData("up 12%"),
    new IWorkTextRunData(", two open risks", underline: true)
}, 60, 300, 840, 80);
presentation.Save("Review.key");
```
<!-- /powerforge:code -->

`Save` refuses to replace an existing file unless you pass `overwrite: true`.

## What it does

- **Create** native Keynote presentations, Pages documents and Numbers spreadsheets, with text, rich text, lists, tables, images and number formats.
- **Read and inspect** Pages, Numbers and Keynote sources: text, tables, cell values, comments and images.
- **Save captured sources** with edits applied in preservation sessions, keeping records the model doesn't change.
- **Bounded package reading** for ZIP and directory bundles, Snappy-framed IWA and protobuf records, with cancellation.
- **Convert** to editable Word, Excel and PowerPoint through opt-in adapters.

The [iWork guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-guide.md) shows each of these with examples.

## Converting to other formats

| Target | Package |
| --- | --- |
| Word from Pages | [OfficeIMO.Word.IWork](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.IWork/README.md) |
| Excel from Numbers | [OfficeIMO.Excel.IWork](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.IWork/README.md) |
| PowerPoint from Keynote | [OfficeIMO.PowerPoint.IWork](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint.IWork/README.md) |
| Conversion workflows | [OfficeIMO.Workflows.IWork](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Workflows.IWork/README.md) |

## Limits

Editing a captured source is limited to the preservation sessions the package qualifies; other records are kept as they are. The [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) lists what each format reads, creates and preserves.

## Learn more

- [OfficeIMO.IWork on officeimo.com](https://officeimo.com/products/iwork/) and the [iWork docs](https://officeimo.com/docs/iwork/).
- [iWork guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-guide.md) and the [support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core` only; the IWA, Snappy, protobuf and package readers are first-party.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 2 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 0 | 2 | 0 | 0 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.IWork` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
