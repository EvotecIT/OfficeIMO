# OfficeIMO.OpenDocument

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.OpenDocument)](https://www.nuget.org/packages/OfficeIMO.OpenDocument)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.OpenDocument?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.OpenDocument)

Create, read and edit OpenDocument text, spreadsheets, presentations and drawings in .NET. `OfficeIMO.OpenDocument` works with `.odt`, `.ods`, `.odp` and `.odg` and their flat XML forms directly, without LibreOffice, Microsoft Office or UNO.

## Install

```sh
dotnet add package OfficeIMO.OpenDocument
```

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/OpenDocument/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO.OpenDocument;

OdtDocument document = OdtDocument.Create();
document.AddHeading("Summary", 1);
document.AddParagraph("Created with OfficeIMO.OpenDocument.");

OdtTable table = document.AddTable(2, 2, "Results");
table.Cell(0, 0).Text = "Metric";
table.Cell(0, 1).Text = "Value";
table.Cell(1, 0).Text = "Revenue";
table.Cell(1, 1).Text = "42";

document.Save("Summary.odt");
```
<!-- /powerforge:code -->

## What it does

- **ODT**: headings, paragraphs, styles, lists, tables, sections, page layout, headers and footers, images, fields and tracked insertions and deletions.
- **ODS**: typed cells, OpenFormula with bounded evaluation, styles and number formats, merges, named ranges, annotations and validation.
- **ODP and ODG**: slides and drawings with shapes, connectors, text and images.
- **Edit without flattening**: change content while keeping unknown package entries.
- **Read text first**, before expanding images and other binary payloads.
- **Encrypt and decrypt** ODF packages with passwords, and sign and validate packages with an explicit security provider.
- **Before sharing**: inspect and remove hidden data, concealed content and provenance.

The [OpenDocument guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.opendocument-guide.md) shows each of these with examples.

## Converting to other formats

| Target | Package |
| --- | --- |
| Word | [OfficeIMO.Word.OpenDocument](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.OpenDocument/README.md) |
| Excel | [OfficeIMO.Excel.OpenDocument](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.OpenDocument/README.md) |
| PowerPoint | [OfficeIMO.PowerPoint.OpenDocument](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint.OpenDocument/README.md) |
| PDF from ODT | [OfficeIMO.OpenDocument.Odt.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.OpenDocument.Odt.Pdf/README.md) |
| PDF from ODS | [OfficeIMO.OpenDocument.Ods.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.OpenDocument.Ods.Pdf/README.md) |
| PDF from ODP | [OfficeIMO.OpenDocument.Odp.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.OpenDocument.Odp.Pdf/README.md) |
| PDF from ODG | [OfficeIMO.OpenDocument.Odg.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.OpenDocument.Odg.Pdf/README.md) |

## Limits

Formula evaluation covers arithmetic, references and common functions; external data, volatile functions and matrix formulas are not evaluated. Content outside the typed model is preserved and reported by conversions that can't map it. The [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md) lists what is supported, preserved or rejected.

## Learn more

- [OfficeIMO.OpenDocument on officeimo.com](https://officeimo.com/products/open-document/) and the [OpenDocument docs](https://officeimo.com/docs/open-document/).
- [OpenDocument guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.opendocument-guide.md) and [examples](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/OpenDocument).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core` only. Signing uses a provider you pass in, such as the optional `OfficeIMO.Security`.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 3 | 0 | 0 | 0 | 0 | 0 |
| Read | 2 | 0 | 0 | 0 | 0 | 1 |
| Edit | 2 | 0 | 0 | 1 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 4 | 0 | 0 | 0 | 0 | 0 |
| Validate | 3 | 0 | 0 | 0 | 0 | 1 |
| Remove | 3 | 0 | 0 | 0 | 0 | 0 |
| Export | 0 | 5 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OpenDocument` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
