# OfficeIMO

[![CI](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml/badge.svg?branch=master)](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml)
[![NuGet](https://img.shields.io/nuget/v/OfficeIMO.Word?label=nuget)](https://www.nuget.org/packages?q=OfficeIMO)
[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/PSWriteOffice.svg?label=PSWriteOffice)](https://www.powershellgallery.com/packages/PSWriteOffice)
[![license](https://img.shields.io/github/license/EvotecIT/OfficeIMO.svg)](LICENSE)
[![Discord](https://img.shields.io/discord/508328927853281280?style=flat-square&label=discord%20chat)](https://evo.yt/discord)

OfficeIMO creates, edits, reads and converts Word, Excel, PowerPoint and PDF files, plus Visio, Project, Access, OpenDocument, Apple iWork, email, e-books, Markdown, HTML and dozens of other formats. It is open source, MIT licensed and doesn't need Microsoft Office or COM. The same engines run in .NET applications, PowerShell scripts, a command line, AI agents, a desktop app and the browser.

## Ways to use it

| You work in | Use | Start with |
| --- | --- | --- |
| .NET | Focused NuGet packages, one per format or conversion | `dotnet add package OfficeIMO` |
| PowerShell 5.1 and 7 | [PSWriteOffice](PowerShell/PSWriteOffice/README.MD) | `Install-Module PSWriteOffice` |
| A terminal or CI job | The `officeimo` command line ([OfficeIMO.Tool](OfficeIMO.Tool/README.md)) | `dotnet tool install --global OfficeIMO.Tool` |
| An AI agent | The same tool as an MCP server, with the [Document Tools plugin](.agents/plugins/officeimo-document-tools/README.md) for Claude Code, Codex, VS Code and Cursor | `officeimo mcp serve --stdio` |
| A desktop app | [OfficeIMO Studio](https://officeimo.com/studio/) for reading, marking up, filling, converting, protecting and signing documents on Windows and Linux ([README](OfficeIMO.Studio/README.md)); native Mac, iPad and iPhone apps are in development ([README](Apple/OfficeIMOStudio/README.md)) | [Downloads](https://officeimo.com/downloads/) |
| A browser | [Browser tools](https://officeimo.com/convert/) that convert and organize files in the tab, without uploading them | [officeimo.com](https://officeimo.com/) |

## Why OfficeIMO

- **No Office, no COM, no native dependencies.** The engines are managed code that runs on Windows, macOS and Linux, in services, containers and build agents. Optional providers, such as Tesseract for OCR, are opt-in. The [package catalog](Docs/officeimo.package-catalog.md#dependencies) lists every third-party dependency.
- **One library for old and new files.** Word, Excel and PowerPoint cover the Open XML formats and the documented subsets of `.doc`, `.xls`, `.xlsb` and `.ppt`, with compatibility reports for what a subset doesn't carry.
- **Conversions say what they lost.** `OfficeConverter` and the `SaveAs` and `To` conversion methods return a report of what the target couldn't carry. You can inspect it, or set `FailOnLoss` to stop instead of writing a lossy file. The [route catalog](Docs/Compatibility/generated/conversion-routes.md) lists each route's result type and limits.
- **Careful with untrusted input.** Readers bound bytes, package entries, ZIP expansion and decoded text before allocating, and flag hidden or prompt-like content before it reaches an AI pipeline. See the [content safety matrix](Docs/officeimo.content-safety-support-matrix.md).
- **NativeAOT and WebAssembly.** CI publishes the core engines with NativeAOT on Linux and Windows, and the officeimo.com browser tools run them in WebAssembly.
- **Measured, not claimed.** Performance figures come from repeatable workloads in [`Docs/performance`](Docs/performance) and the [published benchmarks](https://officeimo.com/benchmarks/), with the method next to each result.
- **Finished-looking output.** New documents start from a theme, so tables, charts and slides look right without styling each element.

## Install

```sh
dotnet add package OfficeIMO
```

`OfficeIMO` brings Word, Excel, PowerPoint, PDF and Markdown, the conversions between them, `OfficeConverter` and the document reader. To keep a dependency small, reference only the [packages](#packages) you use.

## Quick start

A Word report, an Excel table with a chart and a PowerPoint deck from the same data, each also saved as PDF:

<!-- powerforge:code source="Build/DocumentationExamples/QuickStart/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

Sale[] sales = [new("North", 182_400m, 0.12), new("South", 141_900m, 0.08), new("West", 97_300m, -0.03)];

using (var report = WordDocument.Create("Sales.docx")) {
    report.AddHeading("Quarterly sales", level: 1);
    report.AddParagraph("Revenue by region for the third quarter.");
    report.AddTable(sales);
    report.AddChart(OfficeChartType.Column, sales, nameof(Sale.Region), nameof(Sale.Revenue));
    report.Save();
    report.SaveAsPdf("Sales.pdf");
}

using (var book = ExcelDocument.Create("Sales.xlsx", new ExcelCreateOptions { Theme = OfficeTheme.Modern })) {
    ExcelSheet sheet = book.AddSheet("Sales");
    sheet.AddTable(sales, new OfficeTableOptions { Name = "Sales" });
    sheet.AddChart(OfficeChartType.Column, sales, nameof(Sale.Region),
        new ExcelChartOptions { Title = "Revenue by region", Row = 2, Column = 6 }, nameof(Sale.Revenue));
    book.Save();
    book.SaveAsPdf("Sales-workbook.pdf");
}

using (var deck = PowerPointDocument.Create("Sales.pptx")) {
    deck.AddSlide().AddTitle("Quarterly sales");
    PowerPointSlide chartSlide = deck.AddSlide();
    chartSlide.AddTitle("Revenue by region");
    chartSlide.AddChart(OfficeChartType.Column, sales, nameof(Sale.Region), new PowerPointChartOptions {
        Left = OfficeLength.Centimeters(2.3), Top = OfficeLength.Centimeters(5.5),
        Width = OfficeLength.Centimeters(29), Height = OfficeLength.Centimeters(12)
    }, nameof(Sale.Revenue));
    deck.Save();
    deck.SaveAsPdf("Sales-deck.pdf");
}

record Sale(string Region, decimal Revenue, double Growth);
```
<!-- /powerforge:code -->
Every format follows the same shape: `{Format}Document.Create`, `Load` and `Save`; `AddParagraph`, `AddTable` and `AddChart` return what they created; `SaveAsPdf` and the other `SaveAs` methods write another format and return the conversion report. `OfficeConverter.Convert` converts files and whole folders, with resumable batches; the [OfficeIMO package README](OfficeIMO/README.md) shows how. Themes, colors and lengths are shared by every format and described in [OfficeIMO.Core](OfficeIMO.Core/README.md).

## Formats

<!-- officeimo-format-map:start -->
94 formats and 338 conversion routes. Every route runs from .NET; 68 run from the `officeimo` CLI, 68 in Studio, 92 from PSWriteOffice and 14 in the browser tools on officeimo.com.

| Family | Formats | Routes from |
| --- | --- | ---: |
| Documents | DOCX, DOC, ODT, RTF, Pages, Google Docs, DOCM, DOTM, DOTX, Word, Word XML | 38 |
| Spreadsheets | XLSX, ODS, CSV, DBF/xBase, Numbers, Google Sheets, Excel, TSV, XLAM, XLS, XLSB, XLSM, XLTM, XLTX, XML Spreadsheet 2003 | 48 |
| Presentations | PPTX, ODP, Keynote, Google Slides, POT, POTM, POTX, PowerPoint, PPAM, PPS, PPSM, PPSX, PPT, PPTM | 30 |
| PDF and fixed layout | PDF, XPS/OpenXPS, MHTML, DjVu | 32 |
| Diagrams | Visio, ODG/FODG, Binary Visio, Binary Visio stencil, Binary Visio template, VSDX, VSSX, VSTX | 18 |
| Project and Publisher | Project, Publisher | 11 |
| Images | PNG, SVG, JPEG, TIFF, WebP, HEIC | 1 |
| Notes and books | OneNote, EPUB, Comic book ZIP, Comic book 7z, FictionBook 2, Mobipocket and Kindle, Book project, Compiled HTML Help | 34 |
| Notebooks | Jupyter Notebook, Notebook source | 6 |
| Email | Email, EML, MSG, OFT, TNEF | 18 |
| Bibliography | BibTeX, BibLaTeX, CSL JSON, RIS, NBIB/MEDLINE, EndNote XML | 30 |
| Markup and text | Markdown, AsciiDoc, LaTeX, Plain text, OfficeIMO Markup | 18 |
| Web and wiki | HTML, Confluence, ADF | 19 |
| Data | JSON, JSON Lines, TOML, XML, YAML | 35 |

Pick two formats on the [interactive format map](https://officeimo.com/) to see the route between them, or read the [conversion route catalog](Docs/Compatibility/generated/conversion-routes.md) for each route's package, API, support level and limits.
<!-- officeimo-format-map:end -->

## Read documents for search and AI

`OfficeDocumentReader` from [OfficeIMO.Reader.All](OfficeIMO.Reader.All/README.md) reads any supported format into text, tables, metadata and chunks with source locations, ready for indexing or an AI pipeline. Scanned pages go through an optional OCR provider ([OfficeIMO.Ocr](OfficeIMO.Ocr/README.md)). The `officeimo` MCP server exposes the same reading, search and conversion to agents.

Before a document leaves your hands, the engines can find and remove author names, comments, tracked changes, hidden text and other [hidden data](Docs/officeimo.hidden-data-support-matrix.md), check accessibility, and sign Word, PDF and email with [OfficeIMO.Security](OfficeIMO.Security/README.md).

## Packages

Each format has its own package. Most conversions live in an adapter named `OfficeIMO.{Source}.{Target}` that brings only the two engines it joins, and readers for search and AI are `OfficeIMO.Reader.{Format}`. A few routes live in the format package itself; the [route catalog](Docs/Compatibility/generated/conversion-routes.md) names the package for every route. The main engines:

| Area | Package | Document type |
| --- | --- | --- |
| Word `.docx`, `.doc` | [OfficeIMO.Word](OfficeIMO.Word/README.md) | `WordDocument` |
| Excel `.xlsx`, `.xlsb`, `.xls` | [OfficeIMO.Excel](OfficeIMO.Excel/README.md) | `ExcelDocument` |
| PowerPoint `.pptx`, `.ppt` | [OfficeIMO.PowerPoint](OfficeIMO.PowerPoint/README.md) | `PowerPointDocument` |
| PDF | [OfficeIMO.Pdf](OfficeIMO.Pdf/README.md) | `PdfDocument` |
| Markdown | [OfficeIMO.Markdown](OfficeIMO.Markdown/README.md) | `MarkdownDocument` |
| HTML | [OfficeIMO.Html](OfficeIMO.Html/README.md) | `HtmlDocument` |
| Visio | [OfficeIMO.Visio](OfficeIMO.Visio/README.md) | `VisioDocument` |
| Project | [OfficeIMO.Project](OfficeIMO.Project/README.md) | `ProjectDocument` |
| Access | [OfficeIMO.Access](OfficeIMO.Access/README.md) | `AccessDocument` |
| OpenDocument | [OfficeIMO.OpenDocument](OfficeIMO.OpenDocument/README.md) | |
| Apple Pages, Numbers, Keynote | [OfficeIMO.IWork](OfficeIMO.IWork/README.md) | |
| OneNote | [OfficeIMO.OneNote](OfficeIMO.OneNote/README.md) | `OneNoteSection` |
| Email and mail stores | [OfficeIMO.Email](OfficeIMO.Email/README.md) | |
| EPUB | [OfficeIMO.Epub](OfficeIMO.Epub/README.md) | `EpubDocument` |
| RTF | [OfficeIMO.Rtf](OfficeIMO.Rtf/README.md) | `RtfDocument` |
| CSV, YAML, TOML | [OfficeIMO.CSV](OfficeIMO.CSV/README.md), [OfficeIMO.Yaml](OfficeIMO.Yaml/README.md), [OfficeIMO.Toml](OfficeIMO.Toml/README.md) | |
| Archives (ZIP, 7z, CAB, MSI, TAR, gzip) | [OfficeIMO.Archives](OfficeIMO.Archives/README.md) | |
| File and folder conversion, printing | [OfficeIMO](OfficeIMO/README.md), [OfficeIMO.Workflows](OfficeIMO.Workflows/README.md) | `OfficeConverter` |

The [package catalog](Docs/officeimo.package-catalog.md) lists every package, from Publisher, XPS and DjVu to Google Workspace, Confluence and electronic invoices, with its dependencies. The [operation contract](Docs/Compatibility/generated/package-operations.md) shows what each package can create, read, edit and convert.

## Platforms

Every library targets `netstandard2.0`, `net8.0` and `net10.0`; Windows builds also target `net472`, so .NET Framework applications get the same routes. A few packages that need a modern runtime, such as `OfficeIMO.Tool`, `OfficeIMO.AI` and `OfficeIMO.Odbc`, target modern .NET only, and their READMEs say so. A handful of results come from the runtime and can differ between .NET Framework and modern .NET; [platform differences](Docs/officeimo.platform-differences.md) lists them.

Word, Excel and PowerPoint use the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) for package mechanics, and HTML uses [AngleSharp](https://github.com/AngleSharp/AngleSharp) for parsing. The PDF, Markdown, RTF, OpenDocument, OneNote, email and other engines are OfficeIMO's own code.

Keep every OfficeIMO package in an application on the same version.

## Documentation

- [officeimo.com](https://officeimo.com/): guides, the format map, examples and API reference.
- [Documentation index](Docs/README.md): guides, support matrices and evidence in this repository.
- [Examples](OfficeIMO.Examples/README.md): runnable examples for every package.
- [Migration guide](MIGRATION.md) and [GitHub Releases](https://github.com/EvotecIT/OfficeIMO/releases).
- [Roadmap](Docs/ROADMAP.md).

Report bugs and request features in [OfficeIMO issues](https://github.com/EvotecIT/OfficeIMO/issues), for the libraries, the tools and PSWriteOffice alike. Questions and examples are welcome on [Discord](https://evo.yt/discord).

### Snapshot testing

[Verify.OfficeIMO](https://github.com/VerifyTests/Verify.OfficeIMO), from the [Verify](https://github.com/VerifyTests/Verify) project, turns Word, Excel, PowerPoint, PDF, email, Visio and other OfficeIMO documents into text and page images that Verify compares with approved snapshots.

## Sponsors

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys).

<!-- POWERFORGE:sponsors-readme:START -->
<p>
  <a href="https://github.com/KelvinTegelaar" title="KelvinTegelaar"><img src="https://avatars.githubusercontent.com/u/49186168?u=49610dcd84d6c868d9e47c1a64ac137f3da24808&amp;v=4" width="48" height="48" alt="KelvinTegelaar" /></a>
  <a href="https://github.com/apbirch67" title="Andrew Birch"><img src="https://avatars.githubusercontent.com/u/12010032?u=79082c6c1f026e3ab39dae7a4aff8e8fadeeeeea&amp;v=4" width="48" height="48" alt="Andrew Birch" /></a>
  <a href="https://github.com/complea" title="Complea"><img src="https://avatars.githubusercontent.com/u/144781871?v=4" width="48" height="48" alt="Complea" /></a>
  <a href="https://github.com/thomas-moeller" title="Thomas M&#248;ller"><img src="https://avatars.githubusercontent.com/u/37923349?u=e5b9b9d52dc33256937ace98623e3b7c6aead98a&amp;v=4" width="48" height="48" alt="Thomas M&#248;ller" /></a>
  <a href="https://github.com/jakehildreth" title="Jake Hildreth"><img src="https://avatars.githubusercontent.com/u/93942157?u=e1e8d3f460d775c44f5491711ad1c47a5005fb5c&amp;v=4" width="48" height="48" alt="Jake Hildreth" /></a>
  <a href="https://github.com/syazdian" title="syazdian"><img src="https://avatars.githubusercontent.com/u/15229510?v=4" width="48" height="48" alt="syazdian" /></a>
  <a href="https://coderabbit.link/evotecit" title="CodeRabbit"><img src="https://officeimo.com/assets/sponsors/coderabbit-mark.svg" width="48" height="48" alt="CodeRabbit" /></a>
</p>

[See all sponsors](SPONSORS.md)
<!-- POWERFORGE:sponsors-readme:END -->

## License

MIT. See [LICENSE](LICENSE).
