# OfficeIMO

[![CI](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml/badge.svg?branch=master)](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml)
[![codecov](https://codecov.io/gh/EvotecIT/OfficeIMO/branch/master/graph/badge.svg)](https://codecov.io/gh/EvotecIT/OfficeIMO)
[![license](https://img.shields.io/github/license/EvotecIT/OfficeIMO.svg)](LICENSE)
[![Discord](https://img.shields.io/discord/508328927853281280?style=flat-square&label=discord%20chat)](https://evo.yt/discord)

OfficeIMO creates, edits, reads and converts Word, Excel, PowerPoint, PDF, Markdown and two dozen other document formats in .NET, without Microsoft Office or COM. New documents start from a theme, so a table, chart or slide looks finished without styling each element by hand.

PowerShell users get the same engines through [PSWriteOffice](PowerShell/PSWriteOffice/README.MD), which lives in this repository.

## Install

```sh
dotnet add package OfficeIMO
```

`OfficeIMO` brings Word, Excel, PowerPoint, PDF and Markdown, the conversions between them, `OfficeConverter` and the document reader. To keep a dependency small, reference only the [packages](#packages) you use.

## Quick start

A Word report, an Excel table with a chart and a PowerPoint deck from the same data, each also saved as PDF:

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

Every format follows the same shape:

| Job | API |
| --- | --- |
| Create, load and save | `{Format}Document.Create(path)`, `Load(path)`, `Save()`, `Save(path)`, `SaveAsync(stream)` |
| Add content | `AddHeading`, `AddParagraph`, `AddTable(rows)`, `AddChart(type, rows, category, values)`, `AddImage`; each returns what it created |
| Write another format | `doc.SaveAsPdf(path)`, `doc.SaveAsHtml(path)`; returns an `OfficeConversionReport` |
| Convert in memory | `doc.ToMarkdown()`, `markdown.ToWord()` |
| Convert files or folders | `OfficeConverter.Convert(input, output)`, `OfficeConverter.ConvertAsync(files, outputDirectory, target)` |
| Read for search or AI | `OfficeDocumentReader` from `OfficeIMO.Reader.All` |

Conversions report what the target could not carry. `report.HasLoss` and `report.ThrowIfLoss()` check it; `FailOnLoss` in the conversion options turns loss into an error.

### Themes

`OfficeTheme` carries the palette, fonts, type scale, table and chart styles. The built-in themes are `Modern` (the default), `Classic`, `Minimal`, `Boardroom` and `Dark`. Loading an existing document never restyles it.

Colors can name theme roles instead of literal values, so output follows the theme:

```csharp
using OfficeIMO;

OfficeColor brand = OfficeColor.Accent1;
OfficeColor alert = OfficeColor.Danger;
OfficeLength margin = OfficeLength.Centimeters(2.5);
```

### Convert files and folders

```csharp
using OfficeIMO;

OfficeConversionReport report = OfficeConverter.Convert("Sales.docx", "Sales-copy.pdf",
    new OfficeConversionOptions { FileConflictPolicy = OfficePublicationPolicy.Replace });
Console.WriteLine($"{report.Target}: {report.Diagnostics.Count} diagnostics");

OfficeBatchReport batch = await OfficeConverter.ConvertAsync(
    Directory.EnumerateFiles("Reports", "*.docx"),
    outputDirectory: "Pdf",
    target: OfficeDocumentFormat.Pdf,
    new OfficeBatchOptions { MaxParallelism = 4, ConflictPolicy = OfficePublicationPolicy.Replace });
Console.WriteLine($"{batch.Summary.Completed} converted, {batch.Summary.Failed} failed");
```

`OfficeConverter` refuses to replace an existing output unless the conflict policy says so. A batch reports every file and can resume an interrupted run from a checkpoint (`OfficeBatchOptions.ResumeFrom`, which needs the default `Fail` policy).

## Packages

Install a format package for that format alone. Conversion adapters are named `OfficeIMO.{Source}.{Target}` and bring only the two engines they join.

| Format | Package | Document type | Common adapters |
| --- | --- | --- | --- |
| Word `.docx`, `.doc` | [OfficeIMO.Word](OfficeIMO.Word/README.md) | `WordDocument` | `.Word.Pdf`, `.Word.Markdown`, `.Word.Html`, `.Word.Rtf`, `.Word.OpenDocument` |
| Excel `.xlsx`, `.xls` | [OfficeIMO.Excel](OfficeIMO.Excel/README.md) | `ExcelDocument` | `.Excel.Pdf`, `.Excel.Html`, `.Excel.Csv`, `.Excel.OpenDocument` |
| PowerPoint `.pptx`, `.ppt` | [OfficeIMO.PowerPoint](OfficeIMO.PowerPoint/README.md) | `PowerPointDocument` | `.PowerPoint.Pdf`, `.PowerPoint.Html`, `.PowerPoint.OpenDocument` |
| PDF | [OfficeIMO.Pdf](OfficeIMO.Pdf/README.md) | `PdfDocument` | `.Pdf.Word`, `.Pdf.Excel`, `.Pdf.PowerPoint`, `.Pdf.Ocr` |
| Markdown | [OfficeIMO.Markdown](OfficeIMO.Markdown/README.md) | `MarkdownDocument` | `.Markdown.Pdf`, `.Markdown.Html` |
| HTML | [OfficeIMO.Html](OfficeIMO.Html/README.md) | `HtmlDocument` | `.Html.Pdf`, `.Html.Rtf` |
| Visio `.vsdx` | [OfficeIMO.Visio](OfficeIMO.Visio/README.md) | `VisioDocument` | `.Visio.Pdf` |
| Project `.mpp`, `.xml` | [OfficeIMO.Project](OfficeIMO.Project/README.md) | `ProjectDocument` | |
| Access `.accdb`, `.mdb` | [OfficeIMO.Access](OfficeIMO.Access/README.md) | `AccessDocument` | |
| OneNote `.one` | [OfficeIMO.OneNote](OfficeIMO.OneNote/README.md) | `OneNoteSection` | `.OneNote.Markdown`, `.OneNote.Html`, `.OneNote.Pdf` |
| Email `.eml`, `.msg`, `.pst` | [OfficeIMO.Email](OfficeIMO.Email/README.md) | | `.Email.Html`, `.Email.Image` |
| RTF | [OfficeIMO.Rtf](OfficeIMO.Rtf/README.md) | `RtfDocument` | `.Rtf.Pdf`, `.Rtf.Markdown` |
| OpenDocument `.odt`, `.ods`, `.odp` | [OfficeIMO.OpenDocument](OfficeIMO.OpenDocument/README.md) | | `.OpenDocument.Odt.Pdf`, `.Ods.Pdf`, `.Odp.Pdf` |
| EPUB | [OfficeIMO.Epub](OfficeIMO.Epub/README.md) | `EpubDocument` | `.Epub.Image` |
| CSV | [OfficeIMO.CSV](OfficeIMO.CSV/README.md) | | `.Excel.Csv` |
| Apple Pages, Numbers, Keynote | [OfficeIMO.IWork](OfficeIMO.IWork/README.md) | | `.Word.IWork`, `.Excel.IWork`, `.PowerPoint.IWork` |
| AsciiDoc, LaTeX | [OfficeIMO.AsciiDoc](OfficeIMO.AsciiDoc/README.md), [OfficeIMO.Latex](OfficeIMO.Latex/README.md) | | `.Markdown`, `.Pdf` |
| Publisher, XPS, DjVu, CHM, DocBook, OPML | [OfficeIMO.Publisher](OfficeIMO.Publisher/README.md), [OfficeIMO.Xps](OfficeIMO.Xps/README.md), [OfficeIMO.DjVu](OfficeIMO.DjVu/README.md), [OfficeIMO.Chm](OfficeIMO.Chm/README.md), [OfficeIMO.DocBook](OfficeIMO.DocBook/README.md), [OfficeIMO.Opml](OfficeIMO.Opml/README.md) | | `.Pdf` where listed |

Shared and cross-format packages:

| Need | Package |
| --- | --- |
| Theme, colors, lengths, charts, drawing and image export used by every format | [OfficeIMO.Core](OfficeIMO.Core/README.md) |
| File-to-file and batch conversion across every registered route | [OfficeIMO](OfficeIMO/README.md) (`OfficeConverter`) |
| Read, search and chunk any format for indexing and AI ingestion | [OfficeIMO.Reader.All](OfficeIMO.Reader.All/README.md), or one `OfficeIMO.Reader.{Format}` |
| OCR for scanned pages and images | [OfficeIMO.Ocr](OfficeIMO.Ocr/README.md), [OfficeIMO.Ocr.Tesseract](OfficeIMO.Ocr.Tesseract/README.md), [OfficeIMO.Pdf.Ocr](OfficeIMO.Pdf.Ocr/README.md) |
| Certificate signatures for Word, PDF and email | [OfficeIMO.Security](OfficeIMO.Security/README.md) |
| ZIP, CAB and MSI extraction | [OfficeIMO.Archives](OfficeIMO.Archives/README.md) |
| Google Docs, Sheets, Slides and Drive | [OfficeIMO.GoogleWorkspace](OfficeIMO.GoogleWorkspace/README.md) |
| Confluence pages | [OfficeIMO.Confluence](OfficeIMO.Confluence/README.md) |
| Electronic invoices | [OfficeIMO.Invoicing](OfficeIMO.Invoicing/README.md) |

The [package catalog](Docs/officeimo.package-catalog.md) lists every package, what each one covers and its third-party dependencies. The [generated operation contract](Docs/Compatibility/generated/package-operations.md) shows each package's supported operations.

## Tools

- [PSWriteOffice](PowerShell/PSWriteOffice/README.MD): PowerShell commands for every format, from `Install-Module PSWriteOffice`.
- [OfficeIMO.Tool](OfficeIMO.Tool/README.md): the `officeimo` command line, from `dotnet tool install --global OfficeIMO.Tool`. It also serves conversions, extraction and search to AI agents over MCP; the [Document Tools plugin](.agents/plugins/officeimo-document-tools/README.md) installs it in Claude Code, Codex, VS Code and Cursor.
- [OfficeIMO Studio](https://officeimo.com/studio/): an app for reading, marking up, filling, converting, protecting and signing documents on Windows and Linux ([README](OfficeIMO.Studio/README.md)), with native Mac, iPad and iPhone apps in development ([README](Apple/OfficeIMOStudio/README.md)).

## Platforms

Every library targets `netstandard2.0`, `net8.0` and `net10.0`; Windows builds also target `net472`, so .NET Framework applications get the same routes. A few packages that depend on modern-only runtimes, such as `OfficeIMO.Tool`, `OfficeIMO.AI` and `OfficeIMO.Odbc`, target modern .NET only; their READMEs say so.

OfficeIMO runs on Windows, macOS and Linux, in services, containers and build agents. It does not automate Office applications. Word, Excel and PowerPoint use the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) for package mechanics and HTML uses [AngleSharp](https://github.com/AngleSharp/AngleSharp) for parsing. The PDF, Markdown, RTF, OpenDocument, OneNote, email and other engines are OfficeIMO's own code; the [package catalog](Docs/officeimo.package-catalog.md#dependencies) lists every third-party dependency.

Keep every OfficeIMO package in an application on the same version.

## Documentation

- [Migration guide](MIGRATION.md): what changed in 3.6 and how to update code and scripts.
- [Documentation index](Docs/README.md): guides, support matrices and evidence.
- [Examples](OfficeIMO.Examples/README.md): runnable examples for every package.
- [API design](Docs/officeimo.api-design.md): the naming, lifecycle and conversion rules every package follows.
- [Roadmap](Docs/ROADMAP.md) and [GitHub Releases](https://github.com/EvotecIT/OfficeIMO/releases).

Report bugs and request features in [OfficeIMO issues](https://github.com/EvotecIT/OfficeIMO/issues), for the libraries and PSWriteOffice alike.

## Sponsors

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys).

<!-- POWERFORGE:sponsors-readme:START -->
<p>
  <a href="https://github.com/KelvinTegelaar" title="KelvinTegelaar"><img src="https://avatars.githubusercontent.com/u/49186168?u=49610dcd84d6c868d9e47c1a64ac137f3da24808&amp;v=4" width="48" height="48" alt="KelvinTegelaar" /></a>
  <a href="https://github.com/apbirch67" title="Andrew Birch"><img src="https://avatars.githubusercontent.com/u/12010032?u=79082c6c1f026e3ab39dae7a4aff8e8fadeeeeea&amp;v=4" width="48" height="48" alt="Andrew Birch" /></a>
  <a href="https://github.com/thomas-moeller" title="Thomas M&#248;ller"><img src="https://avatars.githubusercontent.com/u/37923349?u=e5b9b9d52dc33256937ace98623e3b7c6aead98a&amp;v=4" width="48" height="48" alt="Thomas M&#248;ller" /></a>
  <a href="https://github.com/jakehildreth" title="Jake Hildreth"><img src="https://avatars.githubusercontent.com/u/93942157?u=e1e8d3f460d775c44f5491711ad1c47a5005fb5c&amp;v=4" width="48" height="48" alt="Jake Hildreth" /></a>
  <a href="https://github.com/complea" title="Complea"><img src="https://avatars.githubusercontent.com/u/144781871?v=4" width="48" height="48" alt="Complea" /></a>
  <a href="https://github.com/DarthPanda12" title="DarthPanda12"><img src="https://avatars.githubusercontent.com/u/294241377?v=4" width="48" height="48" alt="DarthPanda12" /></a>
</p>

[See all sponsors](SPONSORS.md)
<!-- POWERFORGE:sponsors-readme:END -->

## License

MIT. See [LICENSE](LICENSE).
