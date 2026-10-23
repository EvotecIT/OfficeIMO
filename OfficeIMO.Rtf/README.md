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
RtfDocument parsed = RtfDocument.Parse(rtf);
```

`RtfDocument.ToRtf()` writes a normalized semantic document. Use the lossless path when untouched source syntax, unknown destinations, binary payloads, or trailing bytes must remain exact:

```csharp
RtfReadResult read = RtfDocument.LoadResult("input.rtf");
read.SaveLossless("unchanged-copy.rtf");
```

Use `ToRtfResult()` when a semantic save must account for source content that the model cannot retain. Its report includes unbound destinations, unsupported stylesheet syntax, style cycles and missing references, read-policy omissions, and stale encapsulated HTML, even when optional read warnings are disabled:

```csharp
RtfConversionResult<string> normalized = read.Document.ToRtfResult();
string acceptedRtf = normalized.RequireNoLoss();
```

The report belongs to that write operation. A lossless copy preserves the original source; it does not apply semantic edits.

Semantic writing materializes supported effective formatting in the body by default, so native readers such as Word use the intended appearance:

```csharp
RtfConversionResult<string> nativeOutput = document.ToRtfResult(new RtfWriteOptions {
    MaterializeStyleFormatting = true
});
File.WriteAllText("native-output.rtf", nativeOutput.Value);
```

This mode writes the effective formatting supported by the semantic model across body paragraphs, tables, headers, notes, and field results. It leaves the source model unchanged. When it materializes inherited or default values, the result reports `RtfNormalizationFormattingInheritanceMaterialized`: reopening the output treats those values as direct formatting, and `RequireNoLoss()` rejects that loss of editing information. Set `MaterializeStyleFormatting = false` to retain editable inheritance in semantic output; native readers may then display inherited formatting differently. Full character effects and effective table-style formatting remain outside this mode's coverage.

`AddFootnote`, `AddEndnote`, and `AddNoteReference` attach notes to their reference runs. Freshly authored documents use mixed-note mode (`\fet2`) unless `Settings.NoteMode` is set explicitly; endnotes are written as `\footnote\ftnalt`. Note paragraph breaks separate authored paragraphs without adding an empty trailing list item.

On reading, `Settings.NoteMode` retains the declared `\fet` value and stays `null` when the declaration is absent. Parsed documents without `\fet` use the legacy effective mode. The raw note kind and placement declarations remain in the model; the Word bridge projects ordinary notes with absent or section-end placement to endnotes at section end in legacy mode. In modern mode, ordinary notes remain footnotes and `\ftnalt` marks endnotes. Semantic writing retains the parsed note-mode policy.

Modern list bindings use the list tables, and Unicode list markers retain their native character counts.

Word 6/95 legacy paragraph numbering is exposed through `RtfParagraph.LegacyNumbering` and `RtfStyle.LegacyNumbering`. Semantic writing retains the modeled `\pn` controls, marker formatting and text before or after the number. The Word bridge maps supported single-level bullets and decimal lists to editable numbering, retaining start numbers, shared counters and supported marker formatting without changing the RTF model. Higher levels, continuation controls, special cell or section-restart behavior, legacy marker spacing and mismatched cached markers use cached or resolved marker text and report `RtfWordLegacyNumberingFlattened`. The lossless path preserves the original syntax.

Byte, stream, and file reads retain the exact original bytes. `HasOriginalBytes`, `ToBytesLossless()`, and `TryGetLosslessBytes(...)` make that contract observable. Character-only input can be written as lossless bytes only when every source character has an exact single-byte representation; the API reports or throws instead of silently transcoding it.

RTF field instructions are tokenized by `RtfFieldCodeSyntax`, including quoted arguments, switches, escapes, and unterminated-token state. Hyperlink projection uses this syntax rather than regular-expression extraction.

RTF editing restrictions are exposed through `RtfDocumentSettings.HasEditingProtection`. `IsContentEncrypted` remains `false`: form, revision, annotation, and read-only protection flags restrict editing behavior but do not encrypt RTF content.

## Paragraph spacing and inheritance

`RtfParagraph.LineSpacingTwips` stores the raw `\sl` value. Its units depend on `LineSpacingMultiple`, which represents `\slmult`:

| Value | Mode | Meaning |
| --- | --- | --- |
| Positive, with `LineSpacingMultiple = false` | Minimum | A minimum line height in twips; taller text can enlarge the line. |
| Negative, with `LineSpacingMultiple = false` | Exact | A fixed line advance whose magnitude is in twips; taller text does not enlarge it. |
| Nonzero, with `LineSpacingMultiple = true` | Multiple | A line multiple in 240ths: `240` is single spacing and `360` is 1.5 lines. |
| Zero | Automatic | Uses the renderer's natural line spacing. |

For example, `SetLineSpacing(480, false)` requests at least 24 points, while `SetLineSpacing(-480, false)` requests exactly 24 points. The Word and PDF adapters preserve this distinction.

Parsed paragraphs retain the RTF stream's active line-spacing and tab-stop state. An omitted line-spacing control resolves to single spacing; a stylesheet reference does not replace that state or restore a tab stop absent from the active stream. After source styles resolve, an absent before/after gap uses zero unless its automatic-spacing flag is true. The source model retains its nullable controls. For paragraphs created through the API, omitted spacing and an empty direct tab-stop list inherit from the selected paragraph style instead. `ResolveParagraphFormatting(...)` exposes the effective values without changing the source paragraph. Writing implicit parsed single spacing and zero non-auto gaps does not report a loss of style inheritance; materializing actual style-supplied values still does.

## Read untrusted RTF

The default OfficeIMO profile is bounded, does not materialize embedded objects or file-table references, and accepts only web and mail hyperlinks. Uploaded files can state that policy explicitly and provide a cancellation token:

```csharp
RtfReadOptions options = RtfReadOptions.CreateUntrustedProfile();
using FileStream input = File.OpenRead("upload.rtf");

RtfReadResult read = await RtfDocument.LoadResultAsync(
    input,
    options,
    cancellationToken: cancellationToken);
```

The profile caps input bytes and characters, group depth/count, token count, text, binary payloads, images, objects, and semantic block count. A breached limit throws `RtfReadLimitException` with a stable `Code`, `LimitSource`, observed value, configured limit, and source position.

The core never fetches external resources. `RtfReadOptions.CreateCompatibilityProfile()` restores the former unbounded, object-materializing, all-schemes behavior for trusted legacy inputs only.

## Defer picture and object bytes

Set `RtfReadOptions.DeferBinaryPayloads` to keep picture and object data as immutable source fragments during parsing. The parser still enforces source, group, token, binary declaration and semantic payload limits. Native `RtfImage.Data`, `RtfObject.Data` and semantic writing decode bytes when needed. The option defaults to `false` and does not enable `ReadEmbeddedObjects`, which remains `false` by default.

```csharp
using OfficeIMO;
using OfficeIMO.Rtf;

RtfReadResult read = RtfDocument.LoadResult("report.rtf", new RtfReadOptions {
    DeferBinaryPayloads = true
});
using OfficeEmbeddedPayloadInspection payloads = read.InspectEmbeddedPayloads(
    new OfficeEmbeddedPayloadInspectionOptions {
        MaxPayloadBytes = 8L * 1024 * 1024,
        MaxTotalPayloadBytes = 32L * 1024 * 1024,
        MaxReadBytes = 64L * 1024 * 1024
    });
foreach (OfficeEmbeddedPayloadInfo payload in payloads.Payloads) {
    Console.WriteLine($"{payload.Ordinal}: {payload.Id}");
}
using Stream selected = payloads.OpenRead(0, cancellationToken);
```

The session inventories `pict` and `objdata` groups retained in the original syntax, including objects omitted from the semantic document. Its `rtf-source:` locators identify source character positions, not physical byte offsets or OPC relationships. Hexadecimal lengths remain explicitly unknown until a selected read or requested hash completes. Native parser diagnostics still describe source truncation or unsupported syntax; this inventory does not restore groups the parser omitted.

Selected streams and `IncludeSha256` require complete data. Missing, negative, overflowing or truncated `bin` declarations, odd hexadecimal nibbles and non-hexadecimal garbage are refused instead of publishing partial payloads or full hashes. Semantic getters retain their ordinary compatibility behavior. `MaxReadBytes` charges repeated selected reads and hash scans. Session creation snapshots exact source fragments; subsequent semantic or lossless edits do not replace those original-source bytes. Returned streams are read-only, independent buffers that survive session disposal. Objects are never activated, and external content is never fetched. Source and selected-payload limits do not establish a peak-memory guarantee.

Deferred reading rejects `fromhtml` sources containing `pict` or `objdata` before binding because their existing HTML baseline serializes binary semantic content. The default eager profile preserves that behavior; a caller requesting deferral receives `NotSupportedException` for this profile.

## Inspect and remove concealed content

Use `RtfDocument.InspectContentSafety(...)` before sending an untrusted document to a model or automated reviewer. It reports native hidden runs (`\\v`), deleted revisions, tiny or near-zero-scaled text, resolved low-contrast text, non-primary destinations, and suspicious Unicode controls. These findings describe concealment risk; they do not prove that AI authored the document.

Cleanup is explicit and selection-based. Pass finding identifiers from the inspected report to `RtfDocument.RemoveSelectedContent(...)`; OfficeIMO removes only matching semantic text runs or exact reviewed Unicode ranges, including ranges inside alternate HTML encapsulation, writes a normalized RTF artifact, and reinspects the result. Use the shared [content-safety support matrix](../Docs/officeimo.content-safety-support-matrix.md) for exact coverage and format-specific limits.

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
RtfDocument document = RtfDocument.Load("input.rtf");

document.InsertParagraph(0, "Confidential");
document.MoveBlock(0, document.Blocks.Count - 1);
document.ReplaceText("Contoso Ltd.", "Contoso Europe");
document.ReplaceBookmarkText("CustomerName", "Contoso Europe");

RtfDocument independentCopy = document.Clone();
RtfDocumentMergeResult merge = document.AppendDocument(otherDocument);
merge.Report.RequireNoLoss();
```

`AppendDocument` imports fonts, colors, paragraph and character styles, list definitions and overrides, revision authors, blocks, tables, and notes. List resources are discovered through notes, fields, object fallbacks, and text boxes as well as body and table content. Imported resource IDs are remapped, and style inheritance and numbering remain attached to the imported content. File references, XML namespaces, document variables, custom properties, and document information are retained. Conflicting document metadata keeps the destination value and produces an omission diagnostic. Appending into the destination's layout reports source document/section layout and header/footer omission.

An empty destination adopts source document-wide settings where it has no explicit value. When content is combined, destination settings remain in effect and differences are reported as `RtfMergeDocumentSettingsFlattened`, including tab width, facing pages, and editing protection. `PreserveSections` cannot retain separate document-wide settings. Page, word, and character counts are cleared when they no longer describe combined content, with `RtfMergeStatisticsInvalidated`. Alternate source HTML is omitted and reported because it does not describe the merged semantic document.

To retain the source's page setup, columns, and header/footer stories, use section-preserving append:

```csharp
RtfDocumentMergeResult merge = document.AppendDocument(otherDocument,
    new RtfDocumentMergeOptions { PreserveSections = true });
merge.Report.RequireNoLoss();
```

An appended document starts with its own header/footer stories and page defaults. `GetEffectivePageSetup(section)` resolves document defaults beneath a section's explicit settings. `ToRtfResult()` reports when mixed page orientation requires materializing a document-wide landscape default as section settings.

## Lossless structural editing

`RtfLosslessEditor` changes selected syntax nodes while retaining every untouched node:

```csharp
RtfLosslessEditor editor = RtfDocument.LoadResult("input.rtf").EditLossless();

editor.ReplaceText("Old text", "New text");
editor.SetInfo(RtfDocumentInfoField.Title, "Updated title");
editor.InsertRootParagraph(editor.RootNodeCount, "Appended note");
editor.ReplaceImage(0, replacementImage);
editor.ReplaceDestinationContent("header", @"\pard Updated header\par");

editor.SaveLossless("edited.rtf");
```

Root nodes can also be inserted, removed, or moved with `InsertRootRtf`, `RemoveRootNodes`, and `MoveRootNodes`. Those APIs are syntax-indexed; bookmark and rich-text operations belong to the semantic model.

## Encoding and interoperability

The reader supports Unicode escapes, single-byte Windows code pages 874 and 1250-1258, IBM 437/850, Mac Roman, and East Asian Windows code pages 932/936/949/950. Font charset changes can switch the active decoder within a document. Unsupported code pages emit diagnostic `RTF103` and use the documented Windows-1252 fallback while lossless source remains intact. Normalized writing supports Unicode fallback widths (`\uc`) from 0 through 8; lossless source export preserves larger incoming values without expanding them.

Interoperability coverage includes RTF produced by Microsoft Word, Outlook, LibreOffice, Google Docs, macOS TextEdit/RTFD, an Epic EHI export, CRM and helpdesk workflows, and GemBox.Document. Each external sample is exercised through bounded reading, web-safe HTML, Markdown, and diagnostic-preserving Word conversion. See the [RTF support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.rtf-support-matrix.md) for exact producer classifications and known limits.

## Related packages

- `OfficeIMO.Word.Rtf`: Word/DOCX conversion and result-bearing mail merge, find/replace, fields, merge, and compare workflows.
- `OfficeIMO.Html`: web-safe or trusted round-trip HTML conversion.
- `OfficeIMO.Rtf.Markdown`: Markdown conversion with footnotes and media callbacks.
- `OfficeIMO.Rtf.Pdf`: visual PDF export and extractive PDF import.
- `OfficeIMO.Reader.Rtf`: bounded chunk and provenance extraction.

See the [living support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.rtf-support-matrix.md) for feature-level boundaries and evidence.

## Dependency footprint

- **External:** No third-party RTF engine. `System.Text.Encoding.CodePages` supplies legacy encodings.
- **OfficeIMO:** `OfficeIMO.Core`. Lexing, parsing, semantic binding, editing, and writing are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 2 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 1 |
| Edit | 2 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 2 | 0 | 0 | 0 | 0 | 0 |
| Validate | 0 | 0 | 0 | 0 | 0 | 1 |
| Remove | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Rtf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
