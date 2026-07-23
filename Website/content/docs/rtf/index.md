---
title: "RTF Documents"
description: "Read, create, edit, preserve, and convert Rich Text Format documents with bounded parsing and explicit loss reports."
order: 40
meta.seo_title: "Read and convert RTF in .NET with OfficeIMO"
---

`OfficeIMO.Rtf` provides both an editable semantic document and a source-preserving syntax tree. Use the semantic model for normalized authoring and the lossless editor when untouched control words, destinations, binary data, or trailing bytes must survive exactly.

## Install and create

```shell
dotnet add package OfficeIMO.Rtf
```

```csharp
using OfficeIMO.Rtf;

RtfDocument document = RtfDocument.Create();
document.AddParagraph().AddText("Quarterly report").SetBold();
document.AddParagraph("Prepared by OfficeIMO");
File.WriteAllText("report.rtf", document.ToRtf());
```

## Edit semantically

```csharp
RtfDocument document = RtfDocument.Load("input.rtf").Document;

document.InsertParagraph(0, "Confidential");
document.ReplaceText("Contoso Ltd.", "Contoso Europe");
document.ReplaceBookmarkText("CustomerName", "Contoso Europe");
File.WriteAllText("updated.rtf", document.ToRtf());
```

Semantic saves normalize the document. For surgical edits that retain untouched syntax, use `EditLossless()`:

```csharp
RtfLosslessEditor editor = RtfDocument.Load("input.rtf").EditLossless();
editor.ReplaceText("Old text", "New text");
editor.SetInfo(RtfDocumentInfoField.Title, "Updated title");
editor.SaveLossless("updated-lossless.rtf");
```

## Process uploaded RTF safely

```csharp
RtfReadOptions options = RtfReadOptions.CreateUntrustedProfile();
using FileStream input = File.OpenRead("upload.rtf");

RtfReadResult read = await RtfDocument.LoadAsync(
    input,
    options,
    cancellationToken: cancellationToken);
```

The untrusted profile limits input, group depth/count, tokens, text, binary payloads, images, objects, and semantic blocks. It disables embedded-object and file-reference materialization and restricts semantic hyperlinks to web and mail schemes. The engine does not fetch external resources.

## Convert with an explicit fidelity policy

Focused adapters connect RTF to Word, Markdown, HTML, PDF, and Reader. Each adapter reports preserved, flattened, omitted, blocked, and failed content through `RtfConversionReport`.

```csharp
var report = new RtfConversionReport();
report.AddReadDiagnostics(read.Diagnostics, "upload.rtf");
report.Merge(adapterReport);
report.RequireNoLoss();
```

Use `RequireNoLoss()` for strict archival or compliance workflows. In permissive workflows, inspect each diagnostic and accept only the tradeoffs appropriate for the destination.

See the [RTF support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.rtf-support-matrix.md) and the [content publishing patterns](/docs/workflows/content-publishing/).
