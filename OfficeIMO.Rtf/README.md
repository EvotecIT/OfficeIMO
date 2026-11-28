# OfficeIMO.Rtf

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Rtf)](https://www.nuget.org/packages/OfficeIMO.Rtf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Rtf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Rtf)

Read, edit and write Rich Text Format in .NET. `OfficeIMO.Rtf` parses RTF into a source-preserving syntax tree and an editable semantic model, writes deterministic RTF, enforces resource limits, and reports content it can't represent.

## Install

```sh
dotnet add package OfficeIMO.Rtf
```

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/Rtf/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO.Rtf;

RtfDocument document = RtfDocument.Create();
document.AddParagraph().AddText("Quarterly report").SetBold();
document.AddParagraph("Prepared by OfficeIMO");
File.WriteAllText("Report.rtf", document.ToRtf());

RtfDocument parsed = RtfDocument.Parse(File.ReadAllText("Report.rtf"));
```
<!-- /powerforge:code -->

## What it does

- **Create and read** RTF with paragraphs, runs, styles, lists, tables, fields, footnotes, pictures and objects.
- **Semantic editing** of the document model, and **lossless structural editing** of the syntax tree when every control word must survive.
- **Read untrusted RTF** with resource limits, and defer picture and object bytes until you need them.
- **Require no conversion loss** with reports that name anything the model can't carry.
- **Encodings**: legacy code pages and Unicode, for interoperability with Word and other producers.
- **Inspect and remove** concealed content.

The [RTF guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.rtf-guide.md) shows each of these with examples.

## Converting to other formats

| Target | Package |
| --- | --- |
| Word, both directions | [OfficeIMO.Word.Rtf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Rtf/README.md) |
| PDF, both directions | [OfficeIMO.Rtf.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Rtf.Pdf/README.md) |
| Markdown | [OfficeIMO.Rtf.Markdown](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Rtf.Markdown/README.md) |
| HTML, both directions | [OfficeIMO.Html.Rtf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Html.Rtf/README.md) |

## Limits

The [RTF support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.rtf-support-matrix.md) lists feature-level coverage, boundaries and evidence.

## Learn more

- [OfficeIMO.Rtf on officeimo.com](https://officeimo.com/products/rtf/) and the [RTF docs](https://officeimo.com/docs/rtf/).
- [RTF guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.rtf-guide.md).
- Reading for search and AI: [OfficeIMO.Reader.Rtf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Reader.Rtf/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core` and `System.Text.Encoding.CodePages` for legacy encodings.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 2 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 1 |
| Edit | 2 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 2 | 0 | 0 | 0 | 0 | 0 |
| Validate | 0 | 0 | 0 | 0 | 0 | 1 |
| Remove | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Rtf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
