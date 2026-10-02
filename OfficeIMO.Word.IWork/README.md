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

`IWorkSourceDocument.Open` reads and bounds the source independently of destination policy. `ToWordDocument` returns the converted document directly; `ToWordDocumentResult` also exposes the typed Pages projection, diagnostics, preserved source records, and exact editable-versus-visual-fallback result. `WordIWorkConverter.ConvertPagesToWord*` provides equivalent path and stream convenience entry points.

Qualified body image attachments retain their inline run positions. A qualified table attachment occupies its own paragraph and becomes a Word table between the surrounding body paragraphs. Nonzero placement modes, tables mixed with other paragraph content, repeated attachment positions for one drawable, and attachment-run hyperlinks remain unsupported. Other source or destination limits can still require the explicit partial policy.

Individual table row heights and column widths are carried into DOCX in twips under the adapter’s precision policy. A resolved native automatic-resize setting maps row heights to minimum constraints, allowing wrapped cell text to grow the row. Explicit fixed settings retain exact heights; absent settings retain the existing exact-height mapping. Unresolved or malformed settings are diagnosed.

Table-region defaults and selected text styles preserve supported fonts, emphasis, colors, alignment, indents, spacing and pagination flags in DOCX. Explicit rich-text formatting takes precedence. Source font sizes and paragraph measurements must fit Word’s destination range; the partial policy permits rounding with the precision diagnostic.

Selected native padding becomes Word cell margins, and top/middle/bottom alignment becomes Word vertical alignment. Source points must fit the Word cell-margin range; the partial policy permits twip rounding with a precision diagnostic. Supported selected and unbanded region solid fills become DOCX shading. Explicit no-fill overrides clear inherited fills, including on empty cells. Unsupported fills and banding retain source diagnostics and follow the conversion acceptance policy.

Supported numeric table formats become editable DOCX text: decimal precision, grouping, percentages, scientific notation, fractions, currency-code prefixes and negative-value parentheses. Red negative formats become run color. The report identifies display approximation because locale-specific symbols, automatic precision and source appearance can differ. Qualified date/time patterns and fixed hour/minute durations also become editable table text through the shared calendar and elapsed formatters; see the [five-pattern source contract](../Docs/officeimo.iwork-support-matrix.md#conversion-acceptance-and-fidelity). Raw values and formula caches remain on the source projection; formatting does not change its raw display properties. Native Pages export qualification remains outside this bounded contract.

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

This retains bounded recoverable editable content and reports incomplete details. If editable output cannot be produced, it rejects a first-page or composite preview. Both settings default to `false`. Inspect `Report.IsPartialEditableReconstruction` before accepting the output; `Report.RequireCompleteEditableReconstruction()` rejects explicitly partial reconstruction. `Report.RequireNoLoss()` also rejects unassessed record fidelity. These policies do not bypass source limits or destination safety checks.

Under the partial policy, positioned tables become flowing editable Word tables and finite measurements are rounded to DOCX units. `IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED` and `IWORK_PAGES_DOCX_PRECISION` identify those approximations; original geometry remains on the source projection.

The path and stream convenience APIs accept cancellation after the options:

```csharp
using var cancellation = new CancellationTokenSource();
using PagesToWordResult cancellable = WordIWorkConverter.ConvertPagesToWordResult(
    "source.pages", readOptions: null, conversionOptions: options,
    cancellationToken: cancellation.Token);
```

This token governs loading, projection, and destination construction. It also governs later projections from `cancellable.Source`; reopen the source with a new token after cancellation. Saving is a separate destination-owner operation.

The adapter directly depends on `OfficeIMO.Core`, `OfficeIMO.IWork`, and `OfficeIMO.Word`. It does not add iWork support to the default Word package graph.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Word.IWork` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
