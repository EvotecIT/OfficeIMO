# OfficeIMO.Rtf

`OfficeIMO.Rtf` is the managed Rich Text Format engine for OfficeIMO. It parses RTF into both a source-preserving syntax tree and an editable semantic model, writes deterministic RTF, enforces resource limits, and reports content that cannot be represented semantically.

## Install

```powershell
dotnet add package OfficeIMO.Rtf
```

## Create and read RTF

```csharp
using OfficeIMO.Rtf;

RtfDocument document = RtfDocument.Create();
document.AddParagraph().AddText("Quarterly report").SetBold();
document.AddParagraph("Prepared by OfficeIMO");

string rtf = document.ToRtf();
RtfReadResult read = RtfDocument.Read(rtf);
```

`RtfDocument.ToRtf()` writes a normalized semantic document. Use the lossless path when untouched source syntax, unknown destinations, binary payloads, or trailing bytes must remain exact:

```csharp
RtfReadResult read = RtfDocument.Load("input.rtf");
read.SaveLossless("unchanged-copy.rtf");
```

## Read untrusted RTF

The compatibility profile is intentionally permissive apart from nesting depth. Uploaded files should use the bounded profile and a cancellation token:

```csharp
RtfReadOptions options = RtfReadOptions.CreateUntrustedProfile();
using FileStream input = File.OpenRead("upload.rtf");

RtfReadResult read = await RtfDocument.LoadAsync(
    input,
    options,
    cancellationToken: cancellationToken);
```

The profile caps input bytes and characters, group depth/count, token count, text, binary payloads, images, objects, and semantic block count. A breached limit throws `RtfReadLimitException` with a stable `Code`, `LimitSource`, observed value, configured limit, and source position.

It also disables embedded-object and file-reference materialization and restricts semantic hyperlinks to web and mail schemes. The core never fetches external resources.

## Require no conversion loss

All adapters use `RtfConversionReport` for preserved, flattened, omitted, and blocked content:

```csharp
var report = new RtfConversionReport();
report.AddReadDiagnostics(read.Diagnostics, "upload.rtf");

// Merge an adapter's report here.
report.Merge(adapterReport);
report.RequireNoLoss();
```

`RequireNoLoss()` throws `RtfConversionLossException` whenever a conversion flattened, omitted, blocked, or failed content. For permissive workflows, inspect `report.Diagnostics` and accept only the actions appropriate for that destination.

## Semantic editing

Semantic editing produces normalized RTF and is the simplest option when the document meaning is more important than its original control-word layout:

```csharp
RtfDocument document = RtfDocument.Load("input.rtf").Document;

document.InsertParagraph(0, "Confidential");
document.MoveBlock(0, document.Blocks.Count - 1);
document.ReplaceText("Contoso Ltd.", "Contoso Europe");
document.ReplaceBookmarkText("CustomerName", "Contoso Europe");

RtfDocument independentCopy = document.Clone();
RtfDocumentMergeResult merge = document.AppendDocument(otherDocument);
merge.Report.RequireNoLoss();
```

`AppendDocument` remaps fonts, colors, revision authors, blocks, tables, and notes. It reports style/list flattening and source header/footer omission rather than hiding those tradeoffs.

## Lossless structural editing

`RtfLosslessEditor` changes selected syntax nodes while retaining every untouched node:

```csharp
RtfLosslessEditor editor = RtfDocument.Load("input.rtf").EditLossless();

editor.ReplaceText("Old text", "New text");
editor.SetInfo(RtfDocumentInfoField.Title, "Updated title");
editor.InsertRootParagraph(editor.RootNodeCount, "Appended note");
editor.ReplaceImage(0, replacementImage);
editor.ReplaceDestinationContent("header", @"\pard Updated header\par");

editor.SaveLossless("edited.rtf");
```

Root nodes can also be inserted, removed, or moved with `InsertRootRtf`, `RemoveRootNodes`, and `MoveRootNodes`. Those APIs are syntax-indexed; bookmark and rich-text operations belong to the semantic model.

## Encoding and interoperability

The reader supports Unicode escapes, single-byte Windows code pages 874 and 1250-1258, IBM 437/850, Mac Roman, and East Asian Windows code pages 932/936/949/950. Font charset changes can switch the active decoder within a document. Unsupported code pages emit diagnostic `RTF103` and use the documented Windows-1252 fallback while lossless source remains intact.

The checked producer corpus includes real Microsoft Word 16 and Outlook 16 output plus pinned LibreOffice regression files. Google Docs, macOS TextEdit/RTFD, EHR/CRM generators, and commercial producer output remain explicitly unverified.

## Related packages

- `OfficeIMO.Word.Rtf`: Word/DOCX conversion and result-bearing mail merge, find/replace, fields, merge, and compare workflows.
- `OfficeIMO.Html`: web-safe or trusted round-trip HTML conversion.
- `OfficeIMO.Rtf.Markdown`: Markdown conversion with footnotes and media callbacks.
- `OfficeIMO.Rtf.Pdf`: visual PDF export and extractive PDF import.
- `OfficeIMO.Reader.Rtf`: bounded chunk and provenance extraction.

See the [living support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.rtf-support-matrix.md) for feature-level boundaries and evidence.

## Dependency footprint

- **External:** No third-party RTF engine. `System.Text.Encoding.CodePages` supplies legacy encodings.
- **OfficeIMO:** `OfficeIMO.Drawing`. Lexing, parsing, semantic binding, editing, and writing are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
