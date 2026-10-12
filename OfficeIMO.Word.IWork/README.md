# OfficeIMO.Word.IWork

`OfficeIMO.Word.IWork` is the opt-in adapter for importing modern Apple Pages files into editable `OfficeIMO.Word` documents. Installing `OfficeIMO.Word` alone does not add the iWork reader.

```bash
dotnet add package OfficeIMO.Word.IWork
```

```csharp
using OfficeIMO.IWork;
using OfficeIMO.Word.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("source.pages");
using PagesToWordResult result = source.ToWordDocumentResult(
    new IWorkConversionOptions { Mode = IWorkConversionMode.Auto });

Console.WriteLine(result.Report.ProjectionKind);
Console.WriteLine(result.HasLoss);
result.Value.Save("converted.docx");
```

`IWorkSourceDocument.Open` reads and bounds the source independently of destination policy. `ToWordDocument` returns a complete editable document and rejects partial reconstruction or preview output; `ToWordDocumentResult` also exposes the typed Pages projection, diagnostics, preserved source records, and exact editable-versus-visual-fallback result. `WordIWorkConverter.ConvertPagesToWord*` provides equivalent path and stream convenience entry points.

Qualified body image attachments retain their inline run positions and point-sized frames. Native-saved cached flow positions do not become floating DOCX offsets or require measurement rounding. The [Pages image reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-images-v15.4.json) covers PNG alpha, JPEG, descriptions and a native resize through strict save/reopen. Floating image and preview frames also convert source points to Word’s pixel API. A qualified table attachment occupies its own paragraph and becomes a Word table between the surrounding body paragraphs. The adapter removes an empty, neutral attachment paragraph when no floating object needs it as a page anchor. Consecutive tables retain a paragraph separator so receiving applications keep them distinct. Meaningful paragraph spacing, indentation, alignment and pagination flags retain their host paragraph. Nonzero placement modes, tables mixed with other paragraph content, repeated attachment positions for one drawable, and attachment-run hyperlinks remain unsupported. Other source or destination limits can still require the explicit partial policy.

An unmerged inline table with explicit fixed row sizing, complete row/column dimensions, zero offset/rotation and a matching frame can retain complete editable acceptance. Its fixed DOCX grid uses the source widths, and the frame must fit the recovered page content area. The [native table-flow reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-inline-table-flow-word-v16.113.3.json) qualifies the controlled Pages border table through strict DOCX and complete PDF conversion. Word 16.113.3 and OfficeIMO PDF preserve its 432-by-127-point frame within the recorded raster tolerance. Pages 15.4 reimports the same DOCX at 150 points high; that receiving-application row-sizing difference and font substitutions remain appearance limits.

Individual table row heights and column widths are carried into DOCX in twips under the adapter’s precision policy. A resolved native automatic-resize setting maps row heights to minimum constraints, allowing wrapped cell text to grow the row. Unresolved or malformed settings are diagnosed.

Fixed rows, including an absent automatic-resize setting, preserve their physical height in Word by accounting for its added bottom cell padding. A 32-point row with four-point bottom padding uses a 28-point DOCX exact-height declaration. The largest active bottom inset controls each row, including vertical merge continuations; removed horizontal continuations do not contribute. Word row-height declarations are limited to 31,680 twips, and exact declarations must remain positive after padding. An incompatible row box rejects editable reconstruction even under the partial policy. Automatic rows retain their separate minimum-height contract.

Table-region defaults and selected text styles preserve supported fonts, emphasis, colors, alignment, indents, spacing and pagination flags in DOCX. Explicit rich-text formatting takes precedence. Source font sizes and paragraph measurements must fit Word’s destination range; the partial policy permits rounding with the precision diagnostic.

Paragraph text formatting also applies to the paragraph mark, so empty paragraphs retain their declared font metrics. The [native Pages reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-simple-v15.4.json) qualifies a complete three-paragraph body, including an empty paragraph, against Apple DOCX and PDF exports. Saved OfficeIMO DOCX rendering in Pages matches this fixture's word positions; broader pagination and typography remain unqualified.

Selected native padding becomes Word cell margins, and top/middle/bottom alignment becomes Word vertical alignment. Source points must fit the Word cell-margin range; the partial policy permits twip rounding with a precision diagnostic. Supported selected and unbanded region solid fills become DOCX shading. Explicit no-fill overrides clear inherited fills, including on empty cells. Unsupported fills and banding retain source diagnostics and follow the conversion acceptance policy.

Supported numeric table formats become editable DOCX text: decimal precision, grouping, percentages, scientific notation, fractions, currency-code prefixes and negative-value parentheses. Red negative formats become run color. The report identifies display approximation because locale-specific symbols, automatic precision and source appearance can differ. Qualified date/time patterns and fixed abbreviated day-only, hour/minute, hour/second, minute/second or second-only durations also become editable table text through the shared numeric, calendar and elapsed formatters; see the [five-pattern source contract](../Docs/officeimo.iwork-support-matrix.md#conversion-acceptance-and-fidelity). Raw values and formula caches remain on the source projection; formatting does not change its raw display properties. Native Pages export qualification remains outside this bounded contract.

Qualified source-hidden table rows and columns remain visible in editable DOCX output. Conversion requires `AllowPartialEditableReconstruction = true` and reports `IWORK_PAGES_TABLE_VISIBILITY_OMITTED`; strict conversion uses visual fallback when available. Hidden positions and cell content remain on the typed source projection.

Qualified root table comments become DOCX comments anchored to the cell's direct paragraphs, including otherwise empty cells. Plain text, display author and UTC creation time are preserved through the Word comment owner. Replies and collaboration state remain unsupported. Comment text or author names that would require normalization use destination fallback; other source and layout limits still follow the chosen conversion policy. Native Pages export and rendered comment appearance are not qualified.

Choose the acceptance policy explicitly when source details cannot be represented:

```csharp
var options = new IWorkConversionOptions {
    Mode = IWorkConversionMode.Auto,
    AllowPartialEditableReconstruction = true,
    RequireCompleteVisualCoverage = true
};
```

This retains bounded recoverable editable content and reports incomplete details. If editable output cannot be produced, it rejects a first-page or composite preview. `AllowPartialEditableReconstruction` defaults to `false`; `RequireCompleteVisualCoverage` defaults to `true`. To accept an incomplete preview, use the result API with `RequireCompleteVisualCoverage = false` and inspect its coverage and fidelity report. Value-only APIs reject partial and preview output even when these options permit it. `result.RequireCompleteEditableReconstruction()` returns the destination after checking assessed content completeness and disposes rejected output; it does not establish identical appearance or field-level fidelity. `Report.RequireNoLoss()` also rejects unassessed record fidelity. These policies do not bypass source limits or destination safety checks.

Table geometry outside the qualified inline profile requires the partial policy and becomes a flowing editable Word table. `IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED` identifies that approximation. The partial policy also permits rounding finite measurements to DOCX units with `IWORK_PAGES_DOCX_PRECISION`; original geometry remains on the source projection.

Qualified decimal levels 0 and 1 retain starts, restarts, parent resets and continuation through ordinary text. The adapter seeds a new DOCX numbering instance with the recovered counter when a source restart, text gap or parent transition requires it. Instance starts apply once; a child level retains its default reset to one. Parent transitions use a fresh scope so receiving applications cannot carry a prior child restart into the next sequence. `IWorkTextParagraph.ListLayout` retains marker geometry. The bounded decimal profile maps marker indentation, text offset and hanging indentation into paragraph and numbering-level properties, with whole-point marker size independent of uniform body-run size. A marker-font change starts a new numbering instance so it cannot alter earlier markers. Other marker geometry requires `AllowPartialEditableReconstruction = true` and reports `IWORK_PAGES_LIST_LAYOUT_APPROXIMATED`.

The [Pages continuation reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-list-continuation-v15.4.json) checks root/nested left indents of 330/690 twips, a 330-twip hanging indent, 11-point Calibri marker declarations and 12-point Arial body text. Pages 15.4 reopens the DOCX without repair or font warnings and preserves all counters and text/marker starts. It renders the declared 11-point markers at 12 points; exact native marker-size rendering remains unqualified. The earlier [paragraph/list reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-paragraph-list-v15.4.json) retains its separate 393-twip unscaled-marker qualification.

The [counter component references](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/list-counter-mapping-v15.4.json) check child restart at five followed by a parent reset, continuation through ordinary text, and continuation after a marker-font change. Pages 15.4 retains every counter and paragraph in the saved DOCX cases. These synthetic partial-conversion cases qualify counter acceptance only.

The path and stream convenience APIs accept cancellation after the options:

```csharp
using var cancellation = new CancellationTokenSource();
using PagesToWordResult cancellable = WordIWorkConverter.ConvertPagesToWordResult(
    "source.pages", readOptions: null, conversionOptions: options,
    cancellationToken: cancellation.Token);
```

This token governs loading, projection, and destination construction. It also governs later projections from `cancellable.Source`; use `cancellable.Source.WithCancellation(newToken)` to reuse the loaded package with a new operation token. This replaces the previous token while sharing source bytes and parsed messages. Saving is a separate destination-owner operation.

The adapter directly depends on `OfficeIMO.Core`, `OfficeIMO.IWork`, and `OfficeIMO.Word`. It does not add iWork support to the default Word package graph.

Font-relative character tracking is mapped to Word run spacing when the effective font size is known. Signed-twip quantization is reported; unknown font size and destination range violations remain unsupported.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Word.IWork` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
