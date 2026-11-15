---
title: Extract Tables and Records
description: "Extract tables from Word, Excel, PDF and HTML as rows, CSV, Markdown or JSON, and pull key-value records with OfficeIMO.Reader."
order: 59
---

Tables come out of every format in the same shape: a `ReaderTable` with column names, rows of strings, its location in the source, and whether it was truncated. That means the same code exports a table from a Word document, a PDF page, or an Excel range.

## Extract tables

### C#

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .AddPdfHandler()
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("Quote.pdf");
int index = 0;
foreach (ReaderTable table in document.ExtractTables()) {
    index++;
    Console.WriteLine($"{table.Title ?? "Table"}: {table.Columns.Count} columns, {table.Rows.Count} rows");
    File.WriteAllText($"quote-table-{index}.csv", table.ToCsv());
    File.WriteAllText($"quote-table-{index}.md", table.ToMarkdownTable());
}
```

`ReaderOptions.MaxTableRows` (200 by default) bounds rows per table; `TotalRowCount` and `Truncated` tell you when a table was cut.

### PowerShell

```powershell
Read-OfficeDocument -Path .\Quote.pdf -As Tables
```

## Extract key-value records and sections

`ExtractStructured` returns key-value records (such as "Invoice number: 123" lines and two-column tables) and the document's sections.

### C#

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();
OfficeDocumentStructuredExtractionResult structured = reader.ReadDocument("Contract.docx").ExtractStructured();
Console.WriteLine($"{structured.Records.Count} records, {structured.Sections.Count} sections");
```

### PowerShell

```powershell
Read-OfficeDocument -Path .\Contract.docx -As Structure
```

## Limits and reference

**Limits:** table detection in PDFs works from the drawn layout; see [Extract text and tables from PDF](/docs/pdf/extract-text-tables/). See the [compatibility matrix](/compatibility/) for per-format coverage.

**Reference:** [OfficeIMO.Reader API](/api/reader/)
