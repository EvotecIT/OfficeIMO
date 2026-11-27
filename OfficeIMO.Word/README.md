# OfficeIMO.Word

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word)](https://www.nuget.org/packages/OfficeIMO.Word)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word)

Create, edit, read and convert Word documents in .NET without Microsoft Word or COM. `OfficeIMO.Word` works with `.docx`, `.docm`, `.dotx` and `.dotm`, the supported subset of Word 97–2003 `.doc` and `.dot`, and imports Word 6 and Word 95 files. New documents start from the `Modern` theme, so headings, tables and charts look finished without styling each element.

## Install

```sh
dotnet add package OfficeIMO.Word
```

Add an adapter only for the conversions you need: [OfficeIMO.Word.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Pdf/README.md) for PDF and page images, [OfficeIMO.Word.Markdown](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Markdown/README.md) and [OfficeIMO.Word.Html](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Html/README.md) for Markdown and HTML in both directions.

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/Word/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO;
using OfficeIMO.Word;

Milestone[] milestones = [new("Discovery", "Product", "Done"), new("Pilot", "Operations", "At risk"), new("Rollout", "IT", "Planned")];

using var document = WordDocument.Create("Status.docx");
document.AddHeading("Northwind migration", level: 1);

WordParagraph status = document.AddParagraph("Overall status: ");
status.AddText("On track").Bold = true;

document.AddHeading("Milestones", level: 2);
document.AddTable(milestones);

document.Save();

record Milestone(string Name, string Owner, string Status);
```
<!-- /powerforge:code -->

`AddTable` takes any `IEnumerable<T>`: columns come from public properties, the header row repeats on every page, and `OfficeTableOptions` picks, orders and renames columns.

## What it does

- **Create** documents with headings, paragraphs, lists, tables, images, native editable charts, SmartArt, shapes, text boxes, equations, footnotes, endnotes, fields, tables of contents, headers, footers, sections and columns.
- **Edit** existing documents: `Find` and `ReplaceAll` with literal text or regular expressions, and `WordRange` spans that can be replaced, formatted, bookmarked, commented on or linked across runs and paragraphs.
- **Styles and themes**: create and change paragraph, character and table styles through `document.Styles`, import styles from a template, embed fonts, and restyle a loaded document with `ApplyTheme`.
- **Templates and mail merge**: fill placeholders, repeated blocks and conditions from a dictionary or an object with `WordTemplate.Apply`.
- **Review**: comments, tracked changes, document comparison with `WordDocumentComparer`, content controls and protection.
- **Merge and split**: `AppendDocument` brings in another document with its styles, lists, images and notes; `SplitBySection` and `SplitByHeading` cut one into several.
- **Before sharing**: `InspectHiddenData` and `RemoveHiddenData` find and remove authors, comments, tracked changes, hidden text and image locations; `InspectAccessibility` and `FixAccessibility` check and fix alternative text, headings, tables, title and language.
- **Render** native charts to PNG, SVG and other image formats, and whole pages through `OfficeIMO.Word.Pdf`.
- **Load untrusted files** with `WordLoadOptions.UntrustedDefaults`, which rejects macros, embedded objects, ActiveX and external relationships before parsing.

The [Word guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.word-guide.md) shows each of these with examples.

## Converting to other formats

Conversions are methods on the document. `SaveAs*` methods return an `OfficeConversionReport` that lists anything the target couldn't carry; `ToMarkdown()` and `ToWord()` return the converted document, with the report available through their `out OfficeConversionReport` overloads, and the `*Async` forms return the document and report together.

| Target | Package | Entry point |
| --- | --- | --- |
| PDF and page images | [OfficeIMO.Word.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Pdf/README.md) | `document.SaveAsPdf(path)` |
| Markdown, and Markdown to Word | [OfficeIMO.Word.Markdown](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Markdown/README.md) | `document.ToMarkdown()`, `markdown.ToWord()` |
| HTML, and HTML to Word | [OfficeIMO.Word.Html](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Html/README.md) | `document.SaveAsHtml(path)` |
| RTF | [OfficeIMO.Word.Rtf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Rtf/README.md) | |
| OpenDocument text | [OfficeIMO.Word.OpenDocument](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.OpenDocument/README.md) | |
| Google Docs | [OfficeIMO.Word.GoogleDocs](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.GoogleDocs/README.md) | |
| Apple Pages | [OfficeIMO.Word.IWork](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.IWork/README.md) | |
| PDF back to Word | [OfficeIMO.Pdf.Word](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Pdf.Word/README.md) | |

## Limits

`.docx` is the native format. Legacy `.doc` loads into the same model and saves natively for a documented subset; content outside it is reported, preserved or converted according to the compatibility mode you choose. The [compatibility matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word/COMPATIBILITY.md), [legacy DOC compatibility](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.word.legacy-doc-compatibility.md) and the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md) list exactly what is supported, preserved or rejected.

## Learn more

- [OfficeIMO.Word on officeimo.com](https://officeimo.com/products/word/) and the [Word docs](https://officeimo.com/docs/word/).
- [Word guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.word-guide.md), [advanced contracts](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.word-advanced-contracts.md), [templates and mail merge](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.word-template-mail-merge-scenarios.md) and [Word and HTML support](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.word-html-support-matrix.md).
- [Examples](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/Word) and the [migration guide](https://github.com/EvotecIT/OfficeIMO/blob/master/MIGRATION.md).
- PowerShell: [PSWriteOffice](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/README.MD) (`New-OfficeWord`, `Export-OfficeWord`). Command line: `officeimo convert report.docx report.pdf` with [OfficeIMO.Tool](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Tool/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) and `OfficeIMO.Core`. Add `OfficeIMO.Security` only to create or validate document and VBA signatures.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 14 | 12 | 0 | 8 | 2 | 2 |
| Read | 24 | 0 | 9 | 0 | 2 | 3 |
| Edit | 12 | 0 | 22 | 3 | 0 | 1 |
| Preserve | 11 | 0 | 22 | 0 | 0 | 1 |
| Inspect | 5 | 1 | 0 | 0 | 0 | 0 |
| Validate | 4 | 0 | 0 | 0 | 0 | 2 |
| Remove | 4 | 0 | 0 | 0 | 1 | 0 |
| Convert | 31 | 27 | 0 | 8 | 0 | 1 |

The complete rows for `OfficeIMO.Word` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
