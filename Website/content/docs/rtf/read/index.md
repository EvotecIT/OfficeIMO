---
title: Read RTF Files and Extract Text
description: Open RTF files safely, extract paragraph text and document properties, check read diagnostics, inspect embedded pictures and objects, and chunk RTF for search.
order: 41
search_terms:
  - read rtf
  - rtf to text
  - extract text from rtf
  - rtf parser
  - untrusted rtf
packages:
  - OfficeIMO.Rtf
  - OfficeIMO.Reader.Rtf
cmdlets:
  - Open-OfficeRtf
  - Read-OfficeDocument
commands:
  - officeimo reader read
limits: The reader never fetches external resources or activates embedded objects; object bytes are available only for inspection.
verified: 2026-11-14
---

`RtfDocument.Load` returns the semantic model. `RtfDocument.LoadResult` returns an `RtfReadResult` that also carries the read diagnostics and the original bytes, so you can report what the parser skipped and still save an exact copy. For uploads and other untrusted input, pass `RtfLoadOptions.UntrustedDefaults`.

## Extract text and properties

### C#

```csharp
using OfficeIMO.Rtf;

RtfDocument sample = RtfDocument.Create();
sample.Properties.Title = "Visit summary";
sample.Properties.Author = "Front desk";
sample.AddParagraph("Patient arrived at 09:00.");
sample.AddParagraph("Follow-up booked for next week.");
File.WriteAllText("note.rtf", sample.ToRtf());

RtfDocument document = RtfDocument.Load("note.rtf");
Console.WriteLine($"{document.Properties.Title} by {document.Properties.Author}");

foreach (RtfParagraph paragraph in document.Paragraphs) {
    Console.WriteLine(paragraph.ToPlainText());
}
```

### PowerShell

```powershell
$rtf = Open-OfficeRtf -Path .\note.rtf
Read-OfficeDocument -Path .\note.rtf -As Text
```

### CLI

```sh
officeimo reader read note.rtf --format markdown --output note.md
```

## Read uploaded RTF safely

The untrusted profile caps input size, group depth, tokens, text, pictures and objects. It does not materialize embedded objects or file references and keeps only web and mail hyperlinks. A breached limit throws `RtfReadLimitException` with the limit that was hit.

### C#

```csharp
using OfficeIMO.Rtf;

File.WriteAllText("upload.rtf", @"{\rtf1\ansi Uploaded \b text\b0.\par}");

using FileStream input = File.OpenRead("upload.rtf");
RtfReadResult read = await RtfDocument.LoadResultAsync(input, RtfLoadOptions.UntrustedDefaults);

foreach (var diagnostic in read.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity} {diagnostic.Code} at {diagnostic.Position}: {diagnostic.Message}");
}

Console.WriteLine(read.Document.Paragraphs.Count);
```

## Keep an exact copy of the source

`SaveLossless` writes the original bytes back, including content the semantic model does not represent.

### C#

```csharp
using OfficeIMO.Rtf;

File.WriteAllText("legacy.rtf", @"{\rtf1\ansi{\*\customdest kept as is}Legacy text\par}");

RtfReadResult read = RtfDocument.LoadResult("legacy.rtf");
read.SaveLossless("legacy-copy.rtf");

Console.WriteLine(File.ReadAllText("legacy-copy.rtf") == File.ReadAllText("legacy.rtf"));
```

## Inspect embedded pictures and objects

With `DeferBinaryPayloads`, picture and object data stay as source fragments until you ask for them. The inspection lists every payload, and `OpenRead` returns a read-only copy of one. Objects are never activated.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Rtf;

File.WriteAllText("report.rtf", @"{\rtf1\ansi Report\par}");

RtfReadResult read = RtfDocument.LoadResult("report.rtf", new RtfLoadOptions {
    DeferBinaryPayloads = true
});

using OfficeEmbeddedPayloadInspection payloads = read.InspectEmbeddedPayloads(
    new OfficeEmbeddedPayloadInspectionOptions {
        MaxPayloadBytes = 8L * 1024 * 1024,
        MaxTotalPayloadBytes = 32L * 1024 * 1024
    });

foreach (OfficeEmbeddedPayloadInfo payload in payloads.Payloads) {
    Console.WriteLine($"{payload.Ordinal}: {payload.Id}");
}
```

## Split RTF into chunks for search

`OfficeIMO.Reader.Rtf` turns the semantic model into block-aware chunks with Markdown text, ready for an index or a retrieval pipeline.

### C#

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Rtf;

RtfDocument sample = RtfDocument.Create();
sample.AddParagraph("Retention policy");
sample.AddParagraph("Records are kept for seven years.");
File.WriteAllText("policy.rtf", sample.ToRtf());

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddRtfHandler()
    .Build();

foreach (ReaderChunk chunk in reader.Enumerate("policy.rtf")) {
    Console.WriteLine($"{chunk.Id}: {chunk.Markdown ?? chunk.Text}");
}
```

### PowerShell

```powershell
Read-OfficeDocument -Path .\policy.rtf
```

### CLI

```sh
officeimo reader read policy.rtf --format json --output policy.json
```
