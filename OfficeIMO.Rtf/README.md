# OfficeIMO.Rtf

`OfficeIMO.Rtf` is a dependency-free Rich Text Format engine for OfficeIMO.

The package owns RTF parsing, syntax-tree preservation, semantic document binding, fluent document construction, and deterministic RTF writing. Adapters such as `OfficeIMO.Word.Rtf` should reference this package instead of reimplementing RTF parsing or writing.

## Layers

- `OfficeIMO.Rtf.Syntax` tokenizes and parses RTF groups, control words, control symbols, text, and binary payloads while retaining raw source tokens.
- `OfficeIMO.Rtf.Model` exposes a semantic document model for paragraphs, runs, fonts, colors, styles, tables, images, and future document constructs.
- `OfficeIMO.Rtf.Writing` serializes the semantic model back to deterministic RTF.
- `RtfDestinationRegistry` centralizes destination categories used by readers, editors, and diagnostics.
- `RtfReadResult.ToRtfLossless()` and `RtfReadResult.SaveLossless(...)` write the parsed syntax tree without semantic normalization, preserving unknown destinations and raw binary payload bytes.

## Basic Usage

```csharp
using OfficeIMO.Rtf;

RtfDocument document = RtfDocument.Create();
document.AddParagraph()
    .AddText("Hello ")
    .Bold();
document.Paragraphs[0]
    .AddText("RTF");

string rtf = document.ToRtf();
RtfReadResult read = RtfDocument.Read(rtf);
```

## Lossless Round Trip

Use the lossless API when the goal is to preserve an existing RTF stream exactly, including destinations that are not yet semantically modeled.

```csharp
RtfReadResult read = RtfDocument.Load("input.rtf");
read.SaveLossless("output.rtf");
```

The normal `RtfDocument.ToRtf()` path writes from the semantic model and is intentionally normalized. The lossless path writes from the syntax tree captured during read.

## Encoding

RTF hex escapes such as `\'a3` and literal high-byte text loaded through the byte-preserving APIs are decoded according to the active ANSI code page. The reader currently supports dependency-free decoding for single-byte Windows ANSI code pages 874 and 1250 through 1258, with Windows-1252 as the default. Unsupported code pages emit diagnostic `RTF103` and fall back to Windows-1252 while the original syntax remains available through the lossless APIs.

## Lossless Editing

Use `RtfLosslessEditor` for targeted changes that should preserve untouched RTF syntax.

```csharp
RtfReadResult read = RtfDocument.Load("input.rtf");
RtfLosslessEditor editor = read.EditLossless();

editor.ReplaceText("Old text", "New text");
editor.SetInfo(RtfDocumentInfoField.Title, "Updated title");
editor.AppendParagraph("New paragraph");

string editedRtf = editor.ToRtf();
RtfReadResult editedRead = editor.ToReadResult();
```

The editor rewrites only affected syntax nodes and RTF-escapes inserted text. Text replacement intentionally skips structural destinations such as font tables, stylesheets, metadata, pictures, objects, list tables, ignorable destinations, and field instructions.
