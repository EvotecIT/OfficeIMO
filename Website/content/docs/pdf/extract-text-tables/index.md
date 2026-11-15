---
title: "Extract Text and Tables from PDF"
description: "Extract plain text, page text, Markdown, paragraphs, and tables from existing PDF files with OfficeIMO.Pdf, PSWriteOffice, and the officeimo CLI."
layout: docs
---

`ExtractText` returns the text of the whole document in reading order. `Read` returns a logical model: pages with paragraphs and detected tables, plus a Markdown rendering. Both work on PDFs produced by any application; scanned pages without a text layer need OCR first (see [PDF workflows in the CLI](/docs/cli/pdf/)).

## Extract all text

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("statement.pdf");
string text = pdf.ExtractText();
File.WriteAllText("statement.txt", text);
```

### PowerShell

```powershell
Read-OfficeDocument -Path .\statement.pdf -As Text
Read-OfficeDocument -Path .\statement.pdf -As Text -ByPage
```

### CLI

```sh
officeimo read statement.pdf --format markdown --output statement.md
```

## Read pages, paragraphs, and tables

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocumentReadResult result = PdfDocument.Load("statement.pdf").Read();

foreach (PdfLogicalPage page in result.Pages) {
    Console.WriteLine($"Page {page.PageNumber}: {page.Paragraphs.Count} paragraphs, {page.Tables.Count} tables");
    foreach (PdfLogicalTable table in page.Tables) {
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        foreach (IReadOnlyList<string> row in data.Rows) {
            Console.WriteLine(string.Join(" | ", row));
        }
    }
}

File.WriteAllText("statement.md", result.ExtractMarkdown());
```

### PowerShell

```powershell
Read-OfficeDocument -Path .\statement.pdf -As Tables
```

### CLI

```sh
officeimo read statement.pdf --format json --output statement.json
```

## Convert PDF tables to Excel

To move every detected table into a workbook, use the PDF-to-Excel conversion route:

```sh
officeimo convert statement.pdf statement.xlsx
```

For searchable chunks with page locations, read the PDF through [OfficeIMO.Reader](/docs/reader/chunking/) instead.

## Limits and reference

**Limits:** table detection works from the drawn layout, so tables without ruling lines or with merged cells may be split or merged differently from the source. See the [PDF package row](/compatibility/#package-pdf) and the [PDF to Excel row](/compatibility/#package-pdf-excel).

**Reference:** [OfficeIMO.Pdf API](/api/pdf/) · [Read-OfficeDocument](/api/powershell/read-officedocument/)
