# OfficeIMO.Word

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word)](https://www.nuget.org/packages/OfficeIMO.Word)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word)

Create, edit, read and save Word documents (`.docx`, and the supported subset of Word 97–2003 `.doc`) in .NET, without Microsoft Word or COM. New documents use the `Modern` theme for headings, body text, tables and charts.

## Install

```sh
dotnet add package OfficeIMO.Word
```

Add an adapter only for the conversions you need:

```sh
dotnet add package OfficeIMO.Word.Pdf        # SaveAsPdf, page images
dotnet add package OfficeIMO.Word.Markdown   # ToMarkdown, Markdown to Word
dotnet add package OfficeIMO.Word.Html       # ToHtml, SaveAsHtml
```

Targets `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.

## Create a report

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

`AddTable` takes any `IEnumerable<T>`: columns come from public properties, the header row repeats on every page and numbers are right-aligned. `OfficeTableOptions` picks, orders and renames columns:

```csharp
using OfficeIMO;
using OfficeIMO.Word;

Milestone[] milestones = [new("Discovery", "Product", "Done"), new("Pilot", "Operations", "At risk")];

using var document = WordDocument.Create("Milestones.docx");
var options = new OfficeTableOptions { Name = "Milestones" };
options.Order.AddRange(["Status", "Name"]);
options.Exclude.Add("Owner");
options.HeaderRenames["Name"] = "Milestone";
document.AddTable(milestones, options);
document.Save();

record Milestone(string Name, string Owner, string Status);
```

## Charts

Charts are native, editable Word charts with their data in an embedded workbook. Name the category and value properties, or pass selectors with `OfficeSeries`:

```csharp
using OfficeIMO;
using OfficeIMO.Word;

Week[] trend = [new("W1", 20, 18), new("W2", 45, 40), new("W3", 70, 66), new("W4", 92, 85)];

using var document = WordDocument.Create("Trend.docx");
document.AddHeading("Completion by week", level: 1);
document.AddChart(OfficeChartType.Line, trend, nameof(Week.Label), nameof(Week.Planned), nameof(Week.Complete));
document.AddChart(OfficeChartType.Column, trend, t => t.Label,
    new OfficeChartOptions { Title = "Complete (%)", Width = OfficeLength.Centimeters(14) },
    OfficeSeries.From(trend, t => t.Complete, "Complete"));
document.Save();

record Week(string Label, int Planned, int Complete);
```

## Headers, footers and images

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Create("Letter.docx");
document.HeaderDefaultOrCreate.AddParagraph("Northwind Traders");
document.FooterDefaultOrCreate.AddParagraph("Page ").AddPageNumber(includeTotalPages: true);

document.AddImage("logo.png", width: OfficeLength.Centimeters(4));
document.AddParagraph("Dear customer,");
document.Save();
```

## Edit an existing document

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Load("Status.docx");
int replaced = document.FindAndReplace("Northwind", "Contoso");
document.AddParagraph($"Updated on {DateTime.Today:d}; {replaced} replacements.");
document.Save("Status-updated.docx");
```

`Save()` writes to the path the document was created or loaded with; `Save(path)` writes a copy and keeps working on it. For files from outside your organization, load with `WordLoadOptions.UntrustedDefaults`, which rejects macros, embedded objects, ActiveX and external relationships before parsing.

## Convert

Conversions are called on the document and return an `OfficeConversionReport` that lists anything the target could not carry:

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("Status.docx");

OfficeConversionReport pdf = document.SaveAsPdf("Status.pdf");
if (pdf.HasLoss) {
    foreach (OfficeDiagnostic diagnostic in pdf.Diagnostics) Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}

document.SaveAsHtml("Status.html");
MarkdownDocument markdown = document.ToMarkdown();
File.WriteAllText("Status.md", markdown.ToMarkdown());

using WordDocument fromMarkdown = MarkdownDocument.Parse("# Notes\n\n- First\n- Second").ToWord();
fromMarkdown.Save("Notes.docx");
```

## More

- [Word guide](../Docs/officeimo.word-guide.md): lists, styles, sections and columns, fields and tables of contents, comments and revisions, content controls, mail merge, protection and signatures, equations, legacy `.doc`, image export and the full chart API.
- [Compatibility matrix](COMPATIBILITY.md) and the [generated operation contract](../Docs/Compatibility/generated/package-operations.md): exactly what is supported, preserved or rejected.
- [Word and HTML support](../Docs/officeimo.word-html-support-matrix.md), [Word advanced contracts](../Docs/officeimo.word-advanced-contracts.md), [templates and mail merge](../Docs/officeimo.word-template-mail-merge-scenarios.md) and [legacy DOC compatibility](../Docs/officeimo.word.legacy-doc-compatibility.md).
- [Examples](../OfficeIMO.Examples/Word) and the [migration guide](../MIGRATION.md).
- Adapters: [OfficeIMO.Word.Pdf](../OfficeIMO.Word.Pdf/README.md), [Markdown](../OfficeIMO.Word.Markdown/README.md), [Html](../OfficeIMO.Word.Html/README.md), [Rtf](../OfficeIMO.Word.Rtf/README.md), [OpenDocument](../OfficeIMO.Word.OpenDocument/README.md), [GoogleDocs](../OfficeIMO.Word.GoogleDocs/README.md) and [IWork](../OfficeIMO.Word.IWork/README.md). PDF back to Word lives in [OfficeIMO.Pdf.Word](../OfficeIMO.Pdf.Word/README.md).
- PowerShell: [PSWriteOffice](../PowerShell/PSWriteOffice/README.MD) (`New-OfficeWord`).

Dependencies: the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) and `OfficeIMO.Core`. Install `OfficeIMO.Security` only to create or validate document and VBA signatures. MIT licensed.

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
| Convert | 31 | 26 | 0 | 8 | 0 | 1 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Word` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
