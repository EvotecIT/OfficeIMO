---
title: Reader Quick Start
description: Read Word, PDF, Excel, Markdown, and other documents into Markdown, plain text, or JSON with OfficeIMO.Reader, PSWriteOffice, and the officeimo CLI.
order: 56
packages:
  - OfficeIMO.Reader.Core
cmdlets:
  - Read-OfficeDocument
---

OfficeIMO.Reader turns a document into one normalized result: text, Markdown, pages, headings, tables, and metadata, whatever the source format. Install the adapters for the formats you read; each one brings the shared `OfficeIMO.Reader.Core` contracts.

```powershell
dotnet add package OfficeIMO.Reader.Word
dotnet add package OfficeIMO.Reader.Pdf
```

## Read a document as Markdown

Build a reader once with the handlers you need, and reuse it.

### C#

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .AddPdfHandler()
    .Build();

string markdown = reader.ConvertToMarkdown("Policy.docx");
File.WriteAllText("Policy.md", markdown);
```

### PowerShell

```powershell
Read-OfficeDocument -Path .\Policy.docx -As Markdown | Set-Content .\Policy.md
```

### CLI

```sh
officeimo read Policy.docx --format markdown --output Policy.md
```

## Read the full result as JSON

The JSON result keeps the structure: source details, pages, blocks, tables, chunks, and diagnostics. It can be stored and loaded back.

### C#

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();

OfficeDocumentReadResult document = reader.ReadDocument("Handbook.pdf");
Console.WriteLine($"{document.Kind}: {document.Pages.Count} pages, {document.Chunks.Count} chunks");

string json = document.ToJson(indented: true);
File.WriteAllText("Handbook.json", json);

OfficeDocumentReadResult restored = OfficeDocumentReadResultJson.Deserialize(json);
```

### PowerShell

```powershell
Read-OfficeDocument -Path .\Handbook.pdf -As Json -Indented | Set-Content .\Handbook.json
```

### CLI

```sh
officeimo read Handbook.pdf --format json --output Handbook.json
```

## Read from a stream

Pass a source name with the stream so the reader can pick the handler from its extension:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

using FileStream upload = File.OpenRead("upload.bin");
string markdown = reader.ConvertToMarkdown(upload, "Contract.docx");
```

## Next steps

- [Chunk documents for search and AI](/docs/reader/chunking/)
- [Read whole folders](/docs/reader/folders/)
- [Extract tables and records](/docs/reader/tables-records/)
- [Check supported formats](/docs/reader/formats/)
