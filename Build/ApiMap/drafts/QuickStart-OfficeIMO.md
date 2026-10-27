# OfficeIMO

OfficeIMO creates, edits, converts and reads Word, Excel, PowerPoint, PDF and Markdown documents in .NET, without Microsoft Office. Every new document starts from a theme, so a table, chart or slide looks finished without styling each element by hand. The same API serves a one-line conversion and a resumable batch of 100,000 files.

```sh
dotnet add package OfficeIMO
```

`OfficeIMO` brings Word, Excel, PowerPoint, PDF, Markdown, the conversions between them, `OfficeConverter` and the document reader. To keep a dependency small, reference only the packages you use.

## Pick your package

| Goal | Package | Start with |
| --- | --- | --- |
| Everything below in one reference | `OfficeIMO` | `OfficeConverter`, any document type |
| Word documents | `OfficeIMO.Word` | `WordDocument.Create` |
| Excel workbooks, including streamed large exports | `OfficeIMO.Excel` | `ExcelDocument.Create`, `ExcelDocument.CreateWriter` |
| PowerPoint decks | `OfficeIMO.PowerPoint` | `PowerPointDocument.Create` |
| PDF creation and editing | `OfficeIMO.Pdf` | `PdfDocument.Create`, `PdfDocument.Load` |
| Word, Excel or PowerPoint to PDF | `OfficeIMO.Word.Pdf`, `OfficeIMO.Excel.Pdf`, `OfficeIMO.PowerPoint.Pdf` | `doc.SaveAsPdf(path)` |
| PDF back to Word, Excel or PowerPoint | `OfficeIMO.Pdf.Word`, `OfficeIMO.Pdf.Excel`, `OfficeIMO.Pdf.PowerPoint` | `pdf.ToWord()` |
| Markdown documents | `OfficeIMO.Markdown` | `MarkdownDocument.Parse`, `MarkdownDocument.Load` |
| Markdown to Word and back | `OfficeIMO.Word.Markdown` | `markdown.ToWord()`, `doc.ToMarkdown()` |
| Read, search and chunk any format for AI ingestion | `OfficeIMO.Reader.All` | `OfficeDocumentReader` |
| Batch conversion with parallelism and resume | `OfficeIMO` | `OfficeConverter.ConvertAsync` |
| OCR for scanned PDFs and images | `OfficeIMO.Pdf.Ocr`, `OfficeIMO.Ocr.Tesseract` | OCR options on the reader and converter |
| Command line | `OfficeIMO.Tool` | `dotnet tool install -g OfficeIMO.Tool`, then `officeimo` |
| PowerShell | `PSWriteOffice` (PowerShell Gallery) | `Install-Module PSWriteOffice` |

## Create, load and save

Every document type, PDF included, has the same lifecycle:

```csharp
using OfficeIMO.Word;

using (var doc = WordDocument.Create("Report.docx"))
{
    doc.AddHeading("Quarterly report", level: 1);
    doc.AddParagraph("Prepared for the operations review.");
    doc.Save();
}

using var existing = WordDocument.Load("Report.docx");
existing.AddParagraph("Reviewed and approved.");
existing.Save("Report-Approved.docx");
```

## Tables and charts from your data

`AddTable` and `AddChart` take typed data in Word, Excel and PowerPoint and return the object they create:

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

using var book = ExcelDocument.Create("Orders.xlsx", new() { Theme = OfficeTheme.Modern });
ExcelSheet sheet = book.AddSheet("Orders");
ExcelTable table = sheet.AddTable(orders);      // styled, frozen header, fitted columns
sheet.AddChart(OfficeChartType.Column, orders, nameof(Order.Customer), nameof(Order.Total));
book.Save();
```

## Save as PDF

Conversions are called on the source document and return a report of anything the target could not carry:

```csharp
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var doc = WordDocument.Create("Status.docx");
doc.AddHeading("Northwind migration", level: 1);
doc.AddTable(milestones);
doc.AddChart(OfficeChartType.Line, trend, nameof(Trend.Week), nameof(Trend.Complete));
doc.Save();

OfficeConversionReport pdf = doc.SaveAsPdf("Status.pdf");
pdf.ThrowIfLoss();
```

## Convert files and formats

`OfficeConverter` handles file-to-file and batch work. `To{Format}()` returns the target document in memory:

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;

OfficeConverter.Convert("Status.docx", "Status.pdf");

OfficeBatchReport batch = await OfficeConverter.ConvertAsync(
    Directory.EnumerateFiles("Reports", "*.docx"),
    outputDirectory: "Pdf",
    target: OfficeDocumentFormat.Pdf,
    new OfficeBatchOptions { MaxParallelism = 8, ResumeFrom = "Pdf/.officeimo-batch.json" });

using var notes = MarkdownDocument.Parse(File.ReadAllText("Notes.md")).ToWord();
notes.Save("Notes.docx");
```
