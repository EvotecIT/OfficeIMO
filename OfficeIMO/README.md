# OfficeIMO

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO)](https://www.nuget.org/packages/OfficeIMO)

`OfficeIMO` is the one-reference package for documents in .NET: Word, Excel, PowerPoint, PDF and Markdown, the conversions between them, `OfficeConverter` for files and folders, and the document reader for search and AI ingestion. It runs without Microsoft Office.

```sh
dotnet add package OfficeIMO
```

Targets `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows, so .NET Framework applications get the same conversion routes. To keep a dependency small, reference the individual [format packages](https://github.com/EvotecIT/OfficeIMO#packages) instead.

## Create documents and save them as PDF

```csharp
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

Sale[] sales = [new("North", 182_400m), new("South", 141_900m), new("West", 97_300m)];

using (var report = WordDocument.Create("Sales.docx")) {
    report.AddHeading("Quarterly sales", level: 1);
    report.AddTable(sales);
    report.AddChart(OfficeChartType.Column, sales, nameof(Sale.Region), nameof(Sale.Revenue));
    report.Save();
    report.SaveAsPdf("Sales.pdf");
}

using (var book = ExcelDocument.Create("Sales.xlsx", new ExcelCreateOptions { Theme = OfficeTheme.Modern })) {
    book.AddSheet("Sales").AddTable(sales);
    book.Save();
    book.SaveAsPdf("Sales-workbook.pdf");
}

record Sale(string Region, decimal Revenue);
```

## Convert files

`OfficeConverter.Convert` picks the route from the file extensions and returns an `OfficeConversionReport`:

```csharp
using OfficeIMO;

OfficeConversionReport report = OfficeConverter.Convert("Sales.docx", "Sales.md",
    new OfficeConversionOptions { FileConflictPolicy = OfficePublicationPolicy.Replace, FailOnLoss = false });
foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity} {diagnostic.Code}: {diagnostic.Message}");
}
```

By default an existing output is not replaced; `FileConflictPolicy` changes that. `FailOnLoss = true` stops a conversion before it writes anything the target cannot carry. `OfficeConversionOptions.WordSave`, `ExcelSave` and `PowerPointSave` set native save options, such as passwords, on the output.

## Convert a folder

```csharp
using OfficeIMO;

OfficeBatchReport batch = await OfficeConverter.ConvertAsync(
    Directory.EnumerateFiles("Reports", "*.docx"),
    outputDirectory: "Pdf",
    target: OfficeDocumentFormat.Pdf,
    new OfficeBatchOptions { MaxParallelism = 4, ResumeFrom = "Pdf/.officeimo-batch.json" });

Console.WriteLine($"{batch.Summary.Completed} converted, {batch.Summary.Reused} already done, {batch.Summary.Failed} failed");
foreach (OfficeBatchItemReport item in batch.Items) {
    Console.WriteLine($"{item.Status} {item.InputPath} -> {item.OutputPath}");
}
```

With `ResumeFrom`, a rerun reuses verified outputs from the checkpoint and converts only what is missing; `RetryFailed` retries recorded failures. The checkpoint format and saved options are the same on every framework.

## Convert in memory

`To{Format}()` returns the target document model; the async form returns the model and its report together:

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using MarkdownDocument source = MarkdownDocument.Parse("# Quarterly review\n\n- Revenue up 12%\n- Two open risks");
OfficeConversionResult<WordDocument> converted = await source.ToWordAsync();
using WordDocument word = converted.Value;
converted.Report.ThrowIfLoss();
word.Save("Review.docx");
```

## Inspect any document

```csharp
using OfficeIMO;

OfficeDocumentInfo info = OfficeDocument.Inspect("Review.docx", new OfficeInspectOptions { IncludeContent = false });
Console.WriteLine(info.Source);
```

`OfficeIMO.Workflows` and `OfficeIMO.Workflows.IWork` add further routes, such as Apple iWork files, through `RegisterConversions()`.

See the [repository README](https://github.com/EvotecIT/OfficeIMO#readme) for the package map and the [migration guide](https://github.com/EvotecIT/OfficeIMO/blob/master/MIGRATION.md) when upgrading.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 27 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
