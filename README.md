# OfficeIMO - Office and document libraries for .NET

[![CI](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml/badge.svg?branch=master)](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml)
[![codecov](https://codecov.io/gh/EvotecIT/OfficeIMO/branch/master/graph/badge.svg)](https://codecov.io/gh/EvotecIT/OfficeIMO)
[![license](https://img.shields.io/github/license/EvotecIT/OfficeIMO.svg)](LICENSE)

[![Blog](https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg)](https://evotec.xyz/hub)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn)](https://www.linkedin.com/in/pklys)
[![Discord](https://img.shields.io/discord/508328927853281280?style=flat-square&label=discord%20chat)](https://evo.yt/discord)

OfficeIMO is a family of COM-free .NET libraries for creating, reading, converting, and exporting Office and document formats. The packages are designed for services, desktop apps, build agents, and automation hosts where Microsoft Office automation is not available or not appropriate.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). Sponsorship helps keep the libraries maintained, tested, and MIT licensed.

PowerShell users should start with [EvotecIT/PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice), which is the PowerShell-facing project built around OfficeIMO.

## Main packages

| Package | Purpose |
| --- | --- |
| [OfficeIMO.Word](OfficeIMO.Word/README.md) | Create, edit, inspect, and convert `.docx` documents, with first-party support for a tested legacy `.doc` subset. |
| [OfficeIMO.Excel](OfficeIMO.Excel/README.md) | Create and modify `.xlsx` workbooks, open legacy `.xls` workbooks, worksheets, tables, ranges, styles, and reports. |
| [OfficeIMO.PowerPoint](OfficeIMO.PowerPoint/README.md) | Generate `.pptx` presentations programmatically. |
| [OfficeIMO.Visio](OfficeIMO.Visio/README.md) | Create, inspect, validate, and export `.vsdx` diagrams without Visio automation. |
| [OfficeIMO.Pdf](OfficeIMO.Pdf/README.md) | Dependency-free PDF creation, reading, inspection, page operations, and converter engine support. |
| [OfficeIMO.OpenDocument](OfficeIMO.OpenDocument/README.md) | Dependency-free native ODT, ODS, and ODP creation, editing, inspection, and preservation. |
| [OfficeIMO.Rtf](OfficeIMO.Rtf/README.md) | Bounded RTF parser, lossless syntax tree, editable semantic model, writer, and conversion reports. |
| [OfficeIMO.Markdown](OfficeIMO.Markdown/README.md) | Typed Markdown AST, builder API, reader, and HTML renderer. |
| [OfficeIMO.Reader](OfficeIMO.Reader/README.md) | Unified read-only extraction facade with modular adapters. |

## Converters and adapters

| Package | Purpose |
| --- | --- |
| [OfficeIMO.Word.Html](OfficeIMO.Word.Html/README.md) | Word to/from HTML conversion. |
| [OfficeIMO.Word.Markdown](OfficeIMO.Word.Markdown/README.md) | Word to/from Markdown conversion. |
| [OfficeIMO.Word.Pdf](OfficeIMO.Word.Pdf/README.md) | Word to PDF through `OfficeIMO.Pdf`. |
| [OfficeIMO.Word.OpenDocument](OfficeIMO.Word.OpenDocument/README.md) | Explicit Word/ODT conversion with feature-mapping reports. |
| [OfficeIMO.Word.Rtf](OfficeIMO.Word.Rtf/README.md) | Result-bearing Word/RTF conversion, mail merge, fields, merge, and comparison workflows. |
| [OfficeIMO.Excel.Pdf](OfficeIMO.Excel.Pdf/README.md) | Excel workbook to PDF through `OfficeIMO.Pdf`. |
| [OfficeIMO.Excel.OpenDocument](OfficeIMO.Excel.OpenDocument/README.md) | Explicit Excel/ODS conversion with bounded sparse expansion and feature-mapping reports. |
| [OfficeIMO.PowerPoint.Pdf](OfficeIMO.PowerPoint.Pdf/README.md) | PowerPoint presentation to PDF through `OfficeIMO.Pdf`. |
| [OfficeIMO.PowerPoint.OpenDocument](OfficeIMO.PowerPoint.OpenDocument/README.md) | Explicit PowerPoint/ODP conversion with feature-mapping reports. |
| [OfficeIMO.Markdown.Html](OfficeIMO.Markdown.Html/README.md) | HTML to Markdown document conversion. |
| [OfficeIMO.Markdown.Pdf](OfficeIMO.Markdown.Pdf/README.md) | Markdown to PDF through `OfficeIMO.Pdf`. |
| [OfficeIMO.Html.Pdf](OfficeIMO.Html.Pdf/README.md) | Direct HTML-to-PDF rendering and PDF-to-HTML projection. |
| [OfficeIMO.Html](OfficeIMO.Html/README.md) | Shared HTML parsing, resource policy, layout, PNG/SVG rendering, and HTML-to/from-RTF. |
| [OfficeIMO.Rtf.Pdf](OfficeIMO.Rtf.Pdf/README.md) | Visual RTF-to-PDF export and extractive PDF-to-RTF import. |

## Markdown, markup, and rendering

| Package | Purpose |
| --- | --- |
| [OfficeIMO.Markup](OfficeIMO.Markup/README.md) | Markdown-inspired semantic authoring model for OfficeIMO documents. |
| [OfficeIMO.Markup.Word](OfficeIMO.Markup.Word/README.md) | Render markup documents to Word. |
| [OfficeIMO.Markup.Excel](OfficeIMO.Markup.Excel/README.md) | Render markup documents to Excel workbooks. |
| [OfficeIMO.Markup.PowerPoint](OfficeIMO.Markup.PowerPoint/README.md) | Render markup documents to PowerPoint presentations. |
| [OfficeIMO.Markup.Cli](OfficeIMO.Markup.Cli/README.md) | CLI parser, validator, preview, and code-emission tooling. |
| [OfficeIMO.MarkdownRenderer](OfficeIMO.MarkdownRenderer/README.md) | Browser/WebView-friendly Markdown rendering shell. |
| [OfficeIMO.MarkdownRenderer.Wpf](OfficeIMO.MarkdownRenderer.Wpf/README.md) | WPF/WebView2 Markdown host control. |
| [OfficeIMO.MarkdownRenderer.IntelligenceX](OfficeIMO.MarkdownRenderer.IntelligenceX/README.md) | IntelligenceX renderer feature pack. |
| [OfficeIMO.MarkdownRenderer.SamplePlugin](OfficeIMO.MarkdownRenderer.SamplePlugin/README.md) | Sample third-party-style renderer plug-in package. |

## Reader family

| Package | Purpose |
| --- | --- |
| [OfficeIMO.Reader](OfficeIMO.Reader/README.md) | Common extraction model and folder/stream helpers. |
| [OfficeIMO.Reader.Csv](OfficeIMO.Reader.Csv/README.md) | CSV/TSV reader adapter. |
| [OfficeIMO.Reader.Epub](OfficeIMO.Reader.Epub/README.md) | EPUB reader adapter. |
| [OfficeIMO.Reader.Html](OfficeIMO.Reader.Html/README.md) | HTML reader adapter. |
| [OfficeIMO.Reader.Json](OfficeIMO.Reader.Json/README.md) | JSON reader adapter. |
| [OfficeIMO.Reader.OpenDocument](OfficeIMO.Reader.OpenDocument/README.md) | Native ODT, ODS, and ODP reader adapter. |
| [OfficeIMO.Reader.Ocr.Process](OfficeIMO.Reader.Ocr.Process/README.md) | Optional versioned external-process OCR provider. |
| [OfficeIMO.Reader.Ocr.Tesseract](OfficeIMO.Reader.Ocr.Tesseract/README.md) | Optional Tesseract CLI OCR provider. |
| [OfficeIMO.Reader.Pdf](OfficeIMO.Reader.Pdf/README.md) | PDF reader adapter. |
| [OfficeIMO.Reader.Rtf](OfficeIMO.Reader.Rtf/README.md) | Bounded RTF chunks, tables, visuals, warnings, and provenance. |
| [OfficeIMO.Reader.Visio](OfficeIMO.Reader.Visio/README.md) | Visio inspection snapshot adapter. |
| [OfficeIMO.Reader.Xml](OfficeIMO.Reader.Xml/README.md) | XML reader adapter. |
| [OfficeIMO.Reader.Yaml](OfficeIMO.Reader.Yaml/README.md) | YAML reader adapter. |
| [OfficeIMO.Reader.Zip](OfficeIMO.Reader.Zip/README.md) | ZIP traversal reader adapter. |

## Google Workspace and primitives

| Package | Purpose |
| --- | --- |
| [OfficeIMO.GoogleWorkspace](OfficeIMO.GoogleWorkspace/README.md) | Shared Google Workspace credentials, sessions, retry, Drive location, and translation reporting. |
| [OfficeIMO.Word.GoogleDocs](OfficeIMO.Word.GoogleDocs/README.md) | Word to Google Docs planning and export scaffolding. |
| [OfficeIMO.Excel.GoogleSheets](OfficeIMO.Excel.GoogleSheets/README.md) | Excel to Google Sheets planning and export scaffolding. |
| [OfficeIMO.CSV](OfficeIMO.CSV/README.md) | Fluent CSV document model. |
| [OfficeIMO.Drawing](OfficeIMO.Drawing/README.md) | Shared color, image, font, and drawing primitives. |
| [OfficeIMO.Zip](OfficeIMO.Zip/README.md) | Safe ZIP traversal primitives. |
| [OfficeIMO.Epub](OfficeIMO.Epub/README.md) | EPUB extraction primitives. |

## Install

Install only the packages you need:

```powershell
dotnet add package OfficeIMO.Word
dotnet add package OfficeIMO.Excel
dotnet add package OfficeIMO.PowerPoint
dotnet add package OfficeIMO.OpenDocument
dotnet add package OfficeIMO.Pdf
```

Converter packages are intentionally separate so applications can opt into the extra dependency surface only when needed:

```powershell
dotnet add package OfficeIMO.Word.Pdf
dotnet add package OfficeIMO.Word.OpenDocument
dotnet add package OfficeIMO.Excel.Pdf
dotnet add package OfficeIMO.Excel.OpenDocument
dotnet add package OfficeIMO.PowerPoint.OpenDocument
dotnet add package OfficeIMO.Markdown.Pdf
```

## Quick example

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Create("report.docx");
document.AddParagraph("OfficeIMO").SetBold();
document.AddParagraph("Created without Microsoft Office automation.");
document.Save();
```

## Common workflows

### Create an Excel report

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("sales.xlsx");
var sheet = workbook.AddWorkSheet("Sales");

sheet.CellValue(1, 1, "Product");
sheet.CellValue(1, 2, "Revenue");
sheet.CellValue(2, 1, "Alpha");
sheet.CellValue(2, 2, 120);
sheet.CellValue(3, 1, "Beta");
sheet.CellValue(3, 2, 92);
sheet.AddTable("A1:B3", hasHeader: true, name: "SalesTable", style: TableStyle.TableStyleMedium2);
sheet.AutoFitColumns();

workbook.Save();
```

### Parse CSV into typed objects

```csharp
using OfficeIMO.CSV;

List<Person> people = CsvDocument.Load("people.csv")
    .EnsureSchema(schema => schema
        .Column("Id").AsInt32().Required()
        .Column("Name").AsString().Required())
    .ValidateOrThrow()
    .Map<Person>(map => map
        .FromColumn<int>("Id", (person, value) => { person.Id = value; return person; })
        .FromColumn<string>("Name", (person, value) => { person.Name = value; return person; }))
    .ToList();

public sealed class Person {
    public int Id { get; set; }
    public string Name { get; set; } = "";
}
```

### Export Word to PDF

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("proposal.docx");
document.SaveAsPdf("proposal.pdf");
```

### Read, split, merge, and stamp PDFs

```csharp
using OfficeIMO.Pdf;

using var source = PdfDocument.Open("packet.pdf");

string firstPageText = source.Read.Text("1");
source.Pages.Extract("1-3").Save("packet-summary.pdf");

PdfDocument.Open("packet.pdf")
    .MergeWith("appendix.pdf")
    .Pages.Delete("2")
    .Stamp.Text("Reviewed")
    .Save("packet-final.pdf");
```

### Convert PDF tables back into editable Office files

```csharp
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Word.Pdf;

PdfExcelTableConverterExtensions.SavePdfTablesAsExcel(
    "statement.pdf",
    "statement-tables.xlsx");

PdfWordTableConverterExtensions.SavePdfTablesAsWord(
    "statement.pdf",
    "statement-tables.docx");
```

### Convert Markdown and HTML to PDF, PNG, and SVG

```csharp
using OfficeIMO.Markdown.Pdf;

"# Status\n\nGenerated by OfficeIMO."
    .SaveAsPdfFromMarkdown("status.pdf");
```

```csharp
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

string html = "<h1>Status</h1><p>Generated by OfficeIMO.</p>";
var options = new HtmlPdfSaveOptions {
    Margins = HtmlRenderMargins.All(32)
};

byte[] pdf = html.ToPdf(options);
byte[] png = html.ToPng(options);
string svg = html.ToSvg(options);

var pdfResult = html.ToPdfDocumentResult(options);
var pngResult = html.ToPngResult(options);
var svgResult = html.ToSvgResult(options);

html.SaveAsPdf("status.pdf", options);
html.SaveAsPng("status.png", options);
html.SaveAsSvg("status.svg", options);
```

### Extract content for indexing or RAG

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Zip;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPdfHandler()
    .AddZipHandler()
    .Build();

var chunks = reader.ReadFolder("KnowledgeBase",
    new ReaderFolderOptions {
        Recurse = true,
        MaxFiles = 500,
        DeterministicOrder = true
    },
    new ReaderOptions {
        MaxChars = 8_000,
        ComputeHashes = true
    }).ToList();
```

### Create a Visio diagram

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("network.vsdx")
    .NetworkTopologyDiagram("Branch topology", topology => topology
        .Title()
        .Root("internet", "Internet", VisioNetworkNodeKind.Internet)
        .Firewall("firewall", "Firewall")
        .Switch("core", "Core Switch")
        .Server("app", "App Server")
        .Ethernet("internet", "firewall", "WAN")
        .Trunk("firewall", "core")
        .Trunk("core", "app"))
    .Save();
```

## Target frameworks

Most shipping libraries target `netstandard2.0`, `net8.0`, and `net10.0`. Some packages also include `net472` or Windows-specific targets where the surface requires it. Check the package README or project file for exact targets.

## Deeper docs

- [Breaking API migration](Docs/officeimo.breaking-api-migration.md)

- [Examples](OfficeIMO.Examples/README.md)
- [PDF current state](Docs/officeimo.pdf.current-state.md)
- [Excel roadmap](Docs/officeimo.excel.roadmap.md)
- [Markdown correctness roadmap](Docs/officeimo.markdown.correctness-roadmap.md)
- [Visio assessment](Docs/officeimo.visio.assessment.md)
- [Website notes](Docs/officeimo.website.md)
- [Changelog](CHANGELOG.MD)
